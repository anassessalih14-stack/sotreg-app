
# calc_engine.py
# Moteur de calcul + export Excel (mois -> fichier .xlsx)
# Output demandé (4 feuilles):
# 1) synthese : Prestataire, Type véhicule, KM facturé, Facture
# 2) entity_circuit : Entité, Circuit, Prestataire, Type véhicule, Age, KM facturé, Facturation
#    + règles: PIPELINE autocars -> MEA/K-MLIKATE ; PIPELINE STCR MINIBUS => 4500 (et équilibrage NAVETTE) ;
#             NAVETTE split /3 ; Minicar social: km présenté au prix Minicar 30P (tarifs period)
# 3) detail_prestataire : Period, Entité, Circuit, Prestataire, Type véhicule, Age, KM/Rotation, Rotation total,
#                         Variance (somme), KM facturé, Montant mise à disposition, Frais de km, Facture total
# 4) sotreg_lines : inchangé

from __future__ import annotations

import sqlite3
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Tuple

import pandas as pd

_LAST_TABLES = None  # populated by calculate_month_to_excel for PDF generation


FORFAIT_AUTOCAR_TEMP = 11000.0  # km/vehicule pour AUTOCAR dispo temporaire (logique existante)

# ---------- Helpers ----------
def _norm(x) -> str:
    if x is None:
        return ""
    return str(x).strip()

def _up(x) -> str:
    return _norm(x).upper()

def _f(x) -> float:
    try:
        if x is None or x == "":
            return 0.0
        return float(x)
    except Exception:
        return 0.0

def _i(x) -> int:
    try:
        if x is None or x == "":
            return 0
        return int(float(x))
    except Exception:
        return 0

def _is_sotreg(provider: str) -> bool:
    return _up(provider) == "SOTREG"

def _is_autocar_stcr_tour(provider: str, vtype: str) -> bool:
    P = _up(provider)
    V = _up(vtype)
    return (P in {"STCR", "S.TOURISME", "S.TOURISME.", "S TOURISME", "S TOURISME."}) and ("AUTOCAR" in V)

def _is_forfait_type(provider: str, vtype: str) -> bool:
    # minibus / minicar / social => facturation par forfait (logique existante)
    V = _up(vtype)
    return ("MINIBUS" in V) or ("MINICAR" in V)

def _rotation_supp(km_supp: float, km_rotation: float) -> float:
    if km_rotation and km_rotation != 0:
        return km_supp / km_rotation
    return 0.0

@dataclass
class Tariff:
    billing_mode: str
    price_mise_dispo: float
    price_km: float
    price_km_supp: float
    price_journalier: float
    price_day: float
    km_forfait_value: float

def _load_tariffs(conn: sqlite3.Connection, period: str) -> Dict[Tuple[str, str, str, str], Tariff]:
    # tariffs schema (from screenshot):
    # period, provider, vehicle_type, age_cat, billing_mode, price_mise_dispo, price_km, price_km_supp,
    # price_km_realise, price_km_non_realise, price_journalier, km_forfait_value, price_day
    rows = conn.execute("""
        SELECT period, provider, vehicle_type,
               COALESCE(age_cat,''),
               COALESCE(billing_mode,''),
               COALESCE(price_mise_dispo,0),
               COALESCE(price_km,0),
               COALESCE(price_km_supp,0),
               COALESCE(price_journalier,0),
               COALESCE(price_day,0),
               COALESCE(km_forfait_value,0)
        FROM tariffs
        WHERE period IN (?, 'base', 'DEFAULT')
    """, (period,)).fetchall()

    out: Dict[Tuple[str, str, str, str], Tariff] = {}
    for per, prov, vt, age, bm, pm, pk, pks, pj, pday, kmf in rows:
        out[(str(per), _up(prov), _up(vt), _up(age))] = Tariff(
            billing_mode=_norm(bm).lower(),
            price_mise_dispo=_f(pm),
            price_km=_f(pk),
            price_km_supp=_f(pks),
            price_journalier=_f(pj),
            price_day=_f(pday),
            km_forfait_value=_f(kmf),
        )
    return out

def _get_tariff(tariffs: Dict, period: str, provider: str, vtype: str, age_cat: str = "") -> Tariff:
    P = _up(provider); V = _up(vtype); A = _up(age_cat)
    # 1) period exact
    for per in (period, "base", "DEFAULT"):
        t = tariffs.get((str(per), P, V, A))
        if t:
            return t
        # fallback age empty
        t = tariffs.get((str(per), P, V, ""))
        if t:
            return t
    # default empty tariff
    return Tariff(billing_mode="", price_mise_dispo=0.0, price_km=0.0, price_km_supp=0.0, price_journalier=0.0, price_day=0.0, km_forfait_value=0.0)

def _get_price_km(tariffs: Dict, period: str, provider: str, vtype: str, age_cat: str = "") -> float:
    return _get_tariff(tariffs, period, provider, vtype, age_cat).price_km

# ---------- Business rules ----------
def _apply_rules(df_lines: pd.DataFrame, tariffs: Dict, period: str) -> pd.DataFrame:
    """
    df_lines columns (minimum):
      Period, Entité, Circuit, Prestataire, Type véhicule, Age,
      KM/Rotation, Rotation total, Variance, KM facturé, Facture,
      Part mise à dispo, Part km, Part supp
    """
    if df_lines.empty:
        return df_lines

    df = df_lines.copy()

    # Normalize a few
    for c in ["Entité", "Circuit", "Prestataire", "Type véhicule", "Age"]:
        if c in df.columns:
            df[c] = df[c].fillna("").astype(str)

    # --- (1) PIPELINE autocars STCR/S.TOURISME -> Mine intégré MEA / K-MLIKATE
    mask_pipe_autocar = (
        df["Entité"].str.upper().str.contains("PIPE", na=False)
        & df["Prestataire"].str.upper().isin(["STCR", "S.TOURISME", "S.TOURISME.", "S TOURISME", "S TOURISME."])
        & df["Type véhicule"].str.upper().str.contains("AUTOCAR", na=False)
    )
    df.loc[mask_pipe_autocar, "Entité"] = "Mine intégré Mea"
    df.loc[mask_pipe_autocar, "Circuit"] = "K-MLIKATE"

    # --- (2) PIPELINE STCR MINIBUS => KM=4500 & Facture=4500*prix_km, then balance on NAVETTE(STCR MINIBUS)
    mask_pipe_minibus = (
        df["Entité"].str.upper().str.contains("PIPE", na=False)
        & (df["Prestataire"].str.upper() == "STCR")
        & (df["Type véhicule"].str.upper().str.contains("MINIBUS", na=False))
    )
    if mask_pipe_minibus.any():
        # Do it per Age (because price can depend on age_cat)
        for age_val, gidx in df[mask_pipe_minibus].groupby(df.loc[mask_pipe_minibus, "Age"]).groups.items():
            idx = list(gidx)
            old_km = df.loc[idx, "KM facturé"].astype(float).sum()
            old_fact = df.loc[idx, "Facture"].astype(float).sum()

            target_km = 4500.0
            price_km = _get_price_km(tariffs, period, "STCR", "minibus", _norm(age_val))
            target_fact = target_km * price_km

            # distribute by existing km (or equal if 0)
            weights = df.loc[idx, "KM facturé"].astype(float).clip(lower=0)
            sw = weights.sum()
            if sw <= 0:
                weights = pd.Series([1.0] * len(idx), index=idx)
                sw = weights.sum()

            share = weights / sw

            # overwrite km/fact/parts for pipeline minibus
            df.loc[idx, "KM facturé"] = (share * target_km).values
            df.loc[idx, "Facture"] = (share * target_fact).values
            df.loc[idx, "Part mise à dispo"] = 0.0
            df.loc[idx, "Part supp"] = 0.0
            df.loc[idx, "Part km"] = df.loc[idx, "Facture"].astype(float)

            delta_km = target_km - old_km
            delta_fact = target_fact - old_fact

            # balance on NAVETTE lines for STCR MINIBUS same age where Entité == 'NAVETTE'
            mask_navette_stcr_minibus = (
                (df["Entité"].str.upper() == "NAVETTE")
                & (df["Prestataire"].str.upper() == "STCR")
                & (df["Type véhicule"].str.upper().str.contains("MINIBUS", na=False))
                & (df["Age"].fillna("").astype(str) == _norm(age_val))
            )
            nav_idx = df.index[mask_navette_stcr_minibus].tolist()
            if nav_idx and (delta_km != 0 or delta_fact != 0):
                nav_km = df.loc[nav_idx, "KM facturé"].astype(float).sum()
                nav_fact = df.loc[nav_idx, "Facture"].astype(float).sum()
                # remove delta proportionally (by km)
                w = df.loc[nav_idx, "KM facturé"].astype(float).clip(lower=0)
                sw2 = w.sum()
                if sw2 <= 0:
                    w = pd.Series([1.0]*len(nav_idx), index=nav_idx)
                    sw2 = w.sum()
                s2 = w / sw2
                df.loc[nav_idx, "KM facturé"] = (df.loc[nav_idx, "KM facturé"].astype(float) - s2 * delta_km).clip(lower=0).values
                df.loc[nav_idx, "Facture"] = (df.loc[nav_idx, "Facture"].astype(float) - s2 * delta_fact).clip(lower=0).values
                # parts: keep proportional to new facture as "Part km"
                df.loc[nav_idx, "Part mise à dispo"] = 0.0
                df.loc[nav_idx, "Part supp"] = 0.0
                df.loc[nav_idx, "Part km"] = df.loc[nav_idx, "Facture"].astype(float)

    # --- (3) NAVETTE split /3 (Entité == NAVETTE), per (Prestataire, Type véhicule, Age)
    nav_mask = df["Entité"].str.upper() == "NAVETTE"
    if nav_mask.any():
        nav = df[nav_mask].copy()
        df = df[~nav_mask].copy()

        split_targets = [
            "Mine intégré Mea",
            "Mine intégré Beni-Amir",
            "Mine intégré Sidi chennane-Daoui",
        ]

        group_cols = ["Prestataire", "Type véhicule", "Age"]
        num_cols = ["Rotation total", "Variance", "KM facturé", "Part mise à dispo", "Part km", "Part supp", "Facture"]
        # ensure cols exist
        for c in num_cols:
            if c not in nav.columns:
                nav[c] = 0.0

        rows = []
        for (prov, vtype, age), g in nav.groupby(group_cols, dropna=False):
            sums = {c: g[c].astype(float).sum() for c in num_cols}
            # keep period
            per = g["Period"].iloc[0] if "Period" in g.columns and len(g) else period
            # km/rotation: weighted avg by rotation total
            if "KM/Rotation" in g.columns:
                rt = g["Rotation total"].astype(float).sum()
                kmrot = (g["KM/Rotation"].astype(float) * g["Rotation total"].astype(float)).sum() / rt if rt > 0 else g["KM/Rotation"].astype(float).mean()
            else:
                kmrot = 0.0

            for ent in split_targets:
                r = {
                    "Period": per,
                    "Entité": ent,
                    "Circuit": "NAVETTE",
                    "Prestataire": prov,
                    "Type véhicule": vtype,
                    "Age": age,
                    "KM/Rotation": kmrot,
                }
                # divide numeric fields
                for c in num_cols:
                    r[c] = sums[c] / 3.0
                rows.append(r)

        df = pd.concat([df, pd.DataFrame(rows)], ignore_index=True)

    # --- (4) Présentation: Minicar social km doit correspondre au prix Minicar 30P du même mois
    mask_social = df["Type véhicule"].str.upper().str.contains("MINICAR", na=False) & df["Type véhicule"].str.upper().str.contains("SOCIAL", na=False)
    if mask_social.any():
        for i in df.index[mask_social]:
            prov = _norm(df.at[i, "Prestataire"])
            age = _norm(df.at[i, "Age"])
            fact = _f(df.at[i, "Facture"])
            price_30p = _get_price_km(tariffs, period, prov, "minicar 30P", age)
            if price_30p and price_30p > 0:
                df.at[i, "KM facturé"] = fact / price_30p
                # for reporting, set Part km = Facture (since now implied price is 30P)
                df.at[i, "Part km"] = fact

    return df

# ---------- Main export ----------

_LAST_TABLES = None  # populated by calculate_month_to_excel

def calculate_month_to_excel(db_path: Path | str, period: str, out_xlsx: Path | str) -> Path:
    """
    Produit un Excel avec 4 feuilles:
      1) synthese
      2) entity_circuit
      3) detail_prestataire
      4) sotreg_lines
    """
    db_path = Path(db_path)
    out_xlsx = Path(out_xlsx)

    conn = sqlite3.connect(str(db_path))
    tariffs = _load_tariffs(conn, period)

    # Non-sotreg raw lines
    fact = pd.read_sql_query("""
        SELECT period, provider, vehicle_type, chauffeur, mle_car, entity, circuit,
               COALESCE(km_per_rotation,0) AS km_rotation,
               COALESCE(rotation_total,0) AS rotation_total
        FROM fact_lines
        WHERE period = ?
    """, conn, params=(period,))

    # SOTREG raw lines
    sot = pd.read_sql_query("""
        SELECT period, provider, vehicle_type, entity, circuit,
               COALESCE(nb_vehicles,0) AS nb_vehicles,
               COALESCE(km_per_rotation,0) AS km_rotation,
               COALESCE(rotation_total,0) AS rotation_total
        FROM fact_lines_sotreg
        WHERE period = ?
    """, conn, params=(period,))

    # vehicle_period (for km compteur / forfait / supp / age_cat)
    veh = pd.read_sql_query("""
        SELECT period, provider, vehicle_type, mle_car,
               COALESCE(km_compteur,0) AS km_compteur,
               COALESCE(km_forfait,0) AS km_forfait,
               COALESCE(km_supp,0) AS km_supp,
               COALESCE(age_cat,'') AS age_cat,
               COALESCE(dispo_type,'permanent') AS dispo_type,
               COALESCE(nb_days,0) AS nb_days,
               COALESCE(zone,'') AS zone
        FROM vehicle_period
        WHERE period = ?
    """, conn, params=(period,))

    # --------- Build SOTREG lines for sheet 4 (unchanged output) ---------
    sot_rows = []
    if not sot.empty:
        for _, r in sot.iterrows():
            provider = _norm(r["provider"])
            vtype = _norm(r["vehicle_type"])
            age = ""  # SOTREG sheet doesn't track age in current model
            t = _get_tariff(tariffs, period, provider, vtype, age)

            km_total = _f(r["km_rotation"]) * _f(r["rotation_total"])
            facture = km_total * t.price_km + _i(r["nb_vehicles"]) * t.price_mise_dispo
            sot_rows.append({
                "Period": period,
                "Entité": _norm(r["entity"]),
                "Circuit": _norm(r["circuit"]),
                "Prestataire": provider,
                "Type véhicule": vtype,
                "Nbre Véhicule": _i(r["nb_vehicles"]),
                "KM/Rotation": _f(r["km_rotation"]),
                "Rotation total": _f(r["rotation_total"]),
                "KM total": km_total,
                "Mise à disposition": _i(r["nb_vehicles"]) * t.price_mise_dispo,
                "Frais KM": km_total * t.price_km,
                "Facture": facture
            })
    df_sot = pd.DataFrame(sot_rows)

    # --------- Non-SOTREG detail lines (for calculations) ---------
    detail_rows = []
    if not fact.empty:
        fact["km_total_line"] = fact["km_rotation"].apply(_f) * fact["rotation_total"].apply(_f)

        if not veh.empty:
            fact = fact.merge(veh, how="left", on=["provider", "vehicle_type", "mle_car"])

        fact["km_compteur"] = fact.get("km_compteur", 0).fillna(0).apply(_f)
        fact["km_forfait"] = fact.get("km_forfait", 0).fillna(0).apply(_f)
        fact["km_supp"] = fact.get("km_supp", 0).fillna(0).apply(_f)
        fact["age_cat"] = fact.get("age_cat", "").fillna("").astype(str)
        fact["dispo_type"] = fact.get("dispo_type", "permanent").fillna("permanent").astype(str).str.lower()
        fact["nb_days"] = fact.get("nb_days", 0).fillna(0).apply(_i)
        fact["zone"] = fact.get("zone", "").fillna("").astype(str)

        # variance by vehicle: share of circuit within a vehicle's total
        fact["veh_km_sum"] = fact.groupby(["provider", "vehicle_type", "mle_car"])["km_total_line"].transform("sum")
        fact["variance"] = fact.apply(lambda x: (x["km_total_line"] / x["veh_km_sum"]) if x["veh_km_sum"] else 0.0, axis=1)

        for _, r in fact.iterrows():
            provider = _norm(r["provider"])
            vtype = _norm(r["vehicle_type"])
            if _is_sotreg(provider):
                continue

            var = _f(r["variance"])
            km_line = _f(r["km_total_line"])
            dispo = _norm(r.get("dispo_type", "permanent")).lower()
            age = _norm(r.get("age_cat", ""))
            zone = _norm(r.get("zone", ""))
            is_mana_min = _up(provider) == "MANAVETTE" and _is_forfait_type(provider, vtype)
            tariff_age = zone if (is_mana_min and zone) else age
            km_comp = _f(r.get("km_compteur", 0))
            km_forfait = _f(r.get("km_forfait", 0))
            km_supp = _f(r.get("km_supp", 0))
            nb_days = _i(r.get("nb_days", 0))

            km_facture = 0.0
            facture = 0.0
            part_mise = 0.0
            part_km = 0.0
            part_supp = 0.0

            if _is_autocar_stcr_tour(provider, vtype):
                t = _get_tariff(tariffs, period, provider, vtype, tariff_age)
                if dispo == "temporaire":
                    prix_veh = nb_days * t.price_day
                    km_facture = var * FORFAIT_AUTOCAR_TEMP
                    facture = var * prix_veh
                    part_km = facture
                else:
                    part_mise = var * t.price_mise_dispo
                    part_km = var * km_comp * t.price_km
                    facture = part_mise + part_km
                    km_facture = var * km_comp

            elif _is_forfait_type(provider, vtype):
                t = _get_tariff(tariffs, period, provider, vtype, tariff_age)
                if dispo == "temporaire":
                    prix_veh = nb_days * t.price_journalier
                    facture = var * prix_veh
                    forfait = km_forfait if km_forfait else t.km_forfait_value
                    km_facture = var * forfait
                    part_km = facture
                else:
                    part_km = var * km_forfait * t.price_km
                    part_supp = var * km_supp * t.price_km_supp
                    facture = part_km + part_supp
                    km_facture = var * (km_forfait + km_supp)

            else:
                t = _get_tariff(tariffs, period, provider, vtype, tariff_age)
                part_km = var * km_forfait * t.price_km
                facture = part_km
                km_facture = var * km_forfait

            # MANAVETTE: adjust KM display using KH price (km_display = facture / price_KH)
            km_facture_display = km_facture
            if is_mana_min and facture > 0:
                t_kh = _get_tariff(tariffs, period, provider, vtype, "KH")
                if t_kh.price_km > 0:
                    km_facture_display = round(facture / t_kh.price_km, 3)

            detail_rows.append({
                "Period": period,
                "Prestataire": provider,
                "Type véhicule": vtype,
                "Chauffeur": _norm(r.get("chauffeur", "")),
                "MLE CAR": _norm(r.get("mle_car", "")),
                "Entité": _norm(r.get("entity", "")),
                "Circuit": _norm(r.get("circuit", "")),
                "Age": age,
                "Dispo": dispo,
                "Nb jours": nb_days,
                "KM/Rotation": _f(r.get("km_rotation", 0)),
                "Rotation total": _f(r.get("rotation_total", 0)),
                "KM total ligne": km_line,
                "KM total véhicule": _f(r.get("veh_km_sum", 0)),
                "Variance": var,
                "KM compteur": km_comp,
                "KM forfait": km_forfait,
                "KM supp": km_supp,
                "Rotation supp": _rotation_supp(km_supp, _f(r.get("km_rotation", 0))),
                "KM facturé": km_facture_display,
                "Part mise à dispo": part_mise,
                "Part km": part_km,
                "Part supp": part_supp,
                "Facture": facture
            })

    df_detail = pd.DataFrame(detail_rows)

    # --------- Build unified lines for aggregation (include SOTREG) ---------
    lines_parts = []

    if not df_detail.empty:
        lines_parts.append(df_detail[[
            "Period", "Entité", "Circuit", "Prestataire", "Type véhicule", "Age",
            "KM/Rotation", "Rotation total", "Variance", "KM facturé",
            "Part mise à dispo", "Part km", "Part supp", "Facture"
        ]].copy())

    if not df_sot.empty:
        tmp = df_sot.copy()
        tmp["Age"] = ""
        tmp["Variance"] = 0.0
        tmp["Part mise à dispo"] = tmp["Mise à disposition"].apply(_f)
        tmp["Part km"] = tmp["Frais KM"].apply(_f)
        tmp["Part supp"] = 0.0
        tmp = tmp.rename(columns={"KM total": "KM facturé"})
        tmp["Facture"] = tmp["Facture"].apply(_f)
        tmp["KM facturé"] = tmp["KM facturé"].apply(_f)
        tmp["KM/Rotation"] = tmp["KM/Rotation"].apply(_f)
        tmp["Rotation total"] = tmp["Rotation total"].apply(_f)
        lines_parts.append(tmp[[
            "Period", "Entité", "Circuit", "Prestataire", "Type véhicule", "Age",
            "KM/Rotation", "Rotation total", "Variance", "KM facturé",
            "Part mise à dispo", "Part km", "Part supp", "Facture"
        ]])

    if lines_parts:
        df_lines = pd.concat(lines_parts, ignore_index=True)
    else:
        df_lines = pd.DataFrame(columns=[
            "Period", "Entité", "Circuit", "Prestataire", "Type véhicule", "Age",
            "KM/Rotation", "Rotation total", "Variance", "KM facturé",
            "Part mise à dispo", "Part km", "Part supp", "Facture"
        ])

    # Apply requested rules
    df_lines_adj = _apply_rules(df_lines, tariffs, period)

    # --------- Sheet 1: synthese (prestataire/type) ---------
    df_synth = df_lines_adj.groupby(["Prestataire", "Type véhicule"], dropna=False).agg(
        **{
            "KM facturé": ("KM facturé", "sum"),
            "Facture": ("Facture", "sum"),
        }
    ).reset_index().sort_values(["Prestataire", "Type véhicule"])

    # Clear Age for MANAVETTE before groupby (KH/OZ must not appear in output)
    _mana_mask = (df_lines_adj["Prestataire"].str.upper() == "MANAVETTE") & \
                 (df_lines_adj["Type véhicule"].str.lower().str.contains("minibus", na=False))
    df_lines_adj.loc[_mana_mask, "Age"] = ""

    # --------- Sheet 2: entity_circuit (with Age) ---------
    df_entity_circuit = df_lines_adj.groupby(
        ["Entité", "Circuit", "Prestataire", "Type véhicule", "Age"], dropna=False
    ).agg(
        **{
            "KM facturé": ("KM facturé", "sum"),
            "Facturation": ("Facture", "sum"),
        }
    ).reset_index().sort_values(["Entité", "Circuit", "Prestataire", "Type véhicule", "Age"])

    # --------- Sheet 3: detail_prestataire (aggregated) ---------
    def _weighted_kmrot(g: pd.DataFrame) -> float:
        rt = g["Rotation total"].astype(float).sum()
        if rt > 0:
            return (g["KM/Rotation"].astype(float) * g["Rotation total"].astype(float)).sum() / rt
        return g["KM/Rotation"].astype(float).mean() if len(g) else 0.0

    grp_cols = ["Period", "Entité", "Circuit", "Prestataire", "Type véhicule", "Age"]
    if df_lines_adj.empty:
        df_detail_prest = pd.DataFrame(columns=[
            "Period", "Entité", "Circuit", "Prestataire", "Type véhicule", "Age",
            "KM/Rotation", "Rotation total", "Variance",
            "KM facturé", "Montant mise à disposition", "Frais de km", "Facture total"
        ])
    else:
        # base aggregates
        base_agg = df_lines_adj.groupby(grp_cols, dropna=False).agg(
            rotation_total=("Rotation total", "sum"),
            variance=("Variance", "sum"),
            km_facture=("KM facturé", "sum"),
            mise_dispo=("Part mise à dispo", "sum"),
            frais_km=("Part km", "sum"),
            facture_total=("Facture", "sum"),
        ).reset_index()

        # KM/Rotation weighted avg
        kmrot = df_lines_adj.groupby(grp_cols, dropna=False).apply(_weighted_kmrot).reset_index(name="KM/Rotation")

        df_detail_prest = base_agg.merge(kmrot, on=grp_cols, how="left")

        df_detail_prest = df_detail_prest.rename(columns={
            "rotation_total": "Rotation total",
            "variance": "Variance",
            "km_facture": "KM facturé",
            "mise_dispo": "Montant mise à disposition",
            "frais_km": "Frais de km",
            "facture_total": "Facture total",
        })

        # order columns
        df_detail_prest = df_detail_prest[[
            "Period", "Entité", "Circuit", "Prestataire", "Type véhicule", "Age",
            "KM/Rotation", "Rotation total", "Variance",
            "KM facturé", "Montant mise à disposition", "Frais de km", "Facture total"
        ]].sort_values(["Period", "Entité", "Circuit", "Prestataire", "Type véhicule", "Age"])

        # Prix km ref = prix KH de la DB (reference fixe, pas un ratio)
        import sqlite3 as _sq3
        _conn_ref = _sq3.connect(str(db_path))
        _kh_prices = {}
        for _period_val in df_detail_prest["Period"].unique():
            _row = _conn_ref.execute(
                "SELECT price_km FROM tariffs WHERE UPPER(provider)='MANAVETTE' AND period=? AND UPPER(age_cat)='KH' LIMIT 1",
                (str(_period_val),)
            ).fetchone()
            if _row:
                _kh_prices[str(_period_val)] = round(float(_row[0]), 4)
        _conn_ref.close()

        is_mana_rows = df_detail_prest["Prestataire"].str.upper() == "MANAVETTE"
        df_detail_prest["Prix km ref"] = None
        for _per, _price in _kh_prices.items():
            _mask = is_mana_rows & (df_detail_prest["Period"].astype(str) == _per)
            df_detail_prest.loc[_mask, "Prix km ref"] = _price

    # --------- Sheet 4: sotreg_lines (unchanged) ---------
    df_sot_lines = df_sot.copy()
    # Keep a consistent order if present
    if not df_sot_lines.empty:
        cols = list(df_sot_lines.columns)
        preferred = ["Period", "Entité", "Circuit", "Prestataire", "Type véhicule"]
        ordered = preferred + [c for c in cols if c not in preferred]
        df_sot_lines = df_sot_lines[ordered]

    # --------- Write Excel ----------
    out_xlsx.parent.mkdir(parents=True, exist_ok=True)

    # expose last tables for PDF generation
    global _LAST_TABLES
    _LAST_TABLES = (df_synth.copy(), df_entity_circuit.copy(), df_detail_prest.copy(), df_sot_lines.copy())

    with pd.ExcelWriter(out_xlsx, engine="openpyxl") as writer:
        df_synth.to_excel(writer, sheet_name="synthese", index=False)
        df_entity_circuit.to_excel(writer, sheet_name="entity_circuit", index=False)
        df_detail_prest.to_excel(writer, sheet_name="detail_prestataire", index=False)
        df_sot_lines.to_excel(writer, sheet_name="sotreg_lines", index=False)

    conn.close()
    return out_xlsx


# ---------------------------
# PDF attachement (2 tables)
# ---------------------------

def _month_label(period: str) -> str:
    p = period.strip()
    return p[:7] if len(p) >= 7 and p[4] == "-" else p

def build_pdf_tables_from_dfs(df_entity_circuit: pd.DataFrame, df_detail_prestataire: pd.DataFrame,
                             period: str, entity: str, provider: str, db_path: str = None):
    per = _month_label(period)

    # Get KH reference price for MANAVETTE from DB
    prix_km_kh = None
    if db_path and _up(provider) == "MANAVETTE":
        import sqlite3 as _sq2
        _c = _sq2.connect(str(db_path))
        _r = _c.execute(
            "SELECT price_km FROM tariffs WHERE UPPER(provider)='MANAVETTE' AND period=? AND UPPER(age_cat)='KH' LIMIT 1",
            (period,)
        ).fetchone()
        if _r:
            prix_km_kh = round(float(_r[0]), 4)
        _c.close()

    ec = df_entity_circuit.copy()
    ec = ec[(ec["Entité"].astype(str).str.lower() == str(entity).lower()) &
            (ec["Prestataire"].astype(str).str.lower() == str(provider).lower())].copy()

    dp = df_detail_prestataire.copy()
    dp = dp[(dp["Entité"].astype(str).str.lower() == str(entity).lower()) &
            (dp["Prestataire"].astype(str).str.lower() == str(provider).lower())].copy()

    # Remove rows with null/zero billing for the PDF
    if "Facturation" in ec.columns:
        ec["Facturation"] = pd.to_numeric(ec["Facturation"], errors="coerce").fillna(0.0)
        ec = ec[ec["Facturation"] > 0].copy()

    billing_col = "Facture total" if "Facture total" in dp.columns else ("Facturation" if "Facturation" in dp.columns else None)
    if billing_col:
        dp[billing_col] = pd.to_numeric(dp[billing_col], errors="coerce").fillna(0.0)
        dp = dp[dp[billing_col] > 0].copy()

    def _present_vehicle_name(v: str) -> str:
        vu = (v or '').strip()
        if 'MINICAR' in vu.upper() and 'SOCIAL' in vu.upper() and str(provider).upper() == 'STCR':
            return 'minicar 30P'
        return vu

    # FACTURE
    f_cols = ["Commande", "Itinéraires", "Montant (HT)"]
    f_rows = []
    if not ec.empty:
        tmp = ec.copy()
        tmp["Age"] = tmp.get("Age", "").fillna("").astype(str)
        tmp["Commande"] = tmp.apply(lambda r: (_present_vehicle_name(str(r["Type véhicule"])) .upper() + (f" {r['Age']}" if r["Age"] and r["Age"] != "nan" else "")), axis=1)
        tmp["Itinéraires"] = tmp["Circuit"].astype(str)
        tmp["Montant (HT)"] = tmp["Facturation"].astype(float)
        g = tmp.groupby(["Commande", "Itinéraires"], dropna=False)["Montant (HT)"].sum().reset_index()
        for _, r in g.iterrows():
            f_rows.append([r["Commande"], r["Itinéraires"], round(float(r["Montant (HT)"]), 2)])

    total_ht = sum((r[2] for r in f_rows), 0.0)
    frais_gestion = round(total_ht * 0.10, 2)
    tva = round((total_ht + frais_gestion) * 0.10, 2)
    total_ttc = round(total_ht + frais_gestion + tva, 2)

    f_rows.append(["", "TOTAL HORS TAXE", round(total_ht, 2)])
    f_rows.append(["", "10% (Frais de Gestion)", frais_gestion])
    f_rows.append(["", "TVA 10%", tva])
    f_rows.append(["", "TOTAL A PAYER T.T.C", total_ttc])

    f_title = f"ANNEXE - Facture ({provider})"

    # DETAIL — colonnes simplifiées pour vérification client
    # MANAVETTE : Circuit | KM facturé | Prix km | Facture total
    # Autres    : Circuit | Véhicule | KM facturé | Prix km | Facture total
    is_mana = _up(provider) == "MANAVETTE"

    if is_mana:
        d_cols = ["Entité", "Circuit", "KM facturé", "Prix km (ref. KH)", "Facture total"]
    else:
        d_cols = ["Entité", "Véhicule", "Circuit", "Age", "KM facturé", "Prix km", "Facture total"]

    d_rows = []
    if not dp.empty:
        tmp = dp.copy()
        for col in ["KM facturé", "Prix km ref", "Montant mise à disposition",
                    "Frais de km", "Facture total"]:
            if col not in tmp.columns:
                tmp[col] = 0.0
        tmp["Age"] = tmp.get("Age", "").fillna("").astype(str)

        for _, r in tmp.iterrows():
            km   = round(float(r.get("KM facturé", 0) or 0), 3)
            ft   = round(float(r.get("Facture total", 0) or 0), 2)
            prix = r.get("Prix km ref", None)
            # If Prix km ref not in df, use prix_km_kh from DB
            if (prix is None or (hasattr(prix,'__float__') and pd.isna(float(prix)))) and prix_km_kh:
                prix = prix_km_kh

            if is_mana:
                prix_display = round(float(prix), 4) if prix and float(prix or 0) > 0 else "-"
                d_rows.append([
                    r.get("Entité", ""),
                    r.get("Circuit", ""),
                    km,
                    prix_display,
                    ft,
                ])
            else:
                # Show: Entité | Véhicule | Circuit | Age | KM | Prix km | Facture
                prix_display = round(float(prix), 4) if prix and float(prix or 0) > 0 else "-"
                d_rows.append([
                    r.get("Entité", ""),
                    _present_vehicle_name(str(r.get("Type véhicule", ""))),
                    r.get("Circuit", ""),
                    r.get("Age", ""),
                    km,
                    prix_display,
                    ft,
                ])

    d_title = "Détail de vérification"
    return (f_title, f_cols, f_rows), (d_title, d_cols, d_rows)


def _build_tariff_table_for_pdf(db_path, period, provider):
    """Tariff table for PDF page 2. MANAVETTE: show KH price as reference."""
    import sqlite3 as _sq
    conn = _sq.connect(str(db_path))
    rows = conn.execute("""
        SELECT vehicle_type, age_cat, price_km, price_journalier, km_forfait_value, billing_mode
        FROM tariffs WHERE period=? AND UPPER(provider)=UPPER(?)
        ORDER BY vehicle_type, age_cat
    """, (period, provider)).fetchall()
    conn.close()

    is_mana = _up(provider) == "MANAVETTE"
    if is_mana:
        kh = next((r for r in rows if str(r[1]).upper() == "KH"), rows[0] if rows else None)
        cols  = ["Type vehicule", "Prix km (ref. KH)"]
        trows = [["minibus", round(float(kh[2] or 0), 4) if kh else "-"]]
    else:
        cols = ["Type vehicule", "Age", "Prix km", "Prix journalier", "KM forfait"]
        seen = set(); trows = []
        for vt, age, pk, pj, kmf, bm in rows:
            key = (vt, age)
            if key not in seen:
                seen.add(key)
                trows.append([str(vt or ""), str(age or ""),
                    round(float(pk), 4) if pk else "-",
                    round(float(pj), 2) if pj else "-",
                    round(float(kmf), 0) if kmf else "-"])
    return "Grille tarifaire", cols, trows


def _render_pdf_direct(out_path, entity, provider, period,
                       f_title, f_cols, f_rows,
                       d_title, d_cols, d_rows,
                       tarif_table=None):
    """Render PDF directly from data — no DB cache, no pdf_attachment dependency."""
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.platypus import (BaseDocTemplate, Frame, PageTemplate, Table,
                                    TableStyle, Paragraph, Spacer, PageBreak, NextPageTemplate)
    from reportlab.lib import colors
    from reportlab.lib.styles import getSampleStyleSheet
    styles = getSampleStyleSheet()

    def _ts(fs=8):
        return TableStyle([
            ("BACKGROUND",    (0,0), (-1,0),  colors.lightgrey),
            ("FONTNAME",      (0,0), (-1,0),  "Helvetica-Bold"),
            ("FONTSIZE",      (0,0), (-1,-1), fs),
            ("GRID",          (0,0), (-1,-1), 0.5, colors.grey),
            ("ALIGN",         (0,0), (-1,-1), "CENTER"),
            ("VALIGN",        (0,0), (-1,-1), "MIDDLE"),
            ("ROWBACKGROUNDS",(0,1), (-1,-1), [colors.whitesmoke, colors.lightyellow]),
            ("LEFTPADDING",   (0,0), (-1,-1), 3),
            ("RIGHTPADDING",  (0,0), (-1,-1), 3),
            ("TOPPADDING",    (0,0), (-1,-1), 2),
            ("BOTTOMPADDING", (0,0), (-1,-1), 2),
        ])

    def _tbl(cols, rows, fs=8, cw=None):
        data = [cols] + [["" if v is None else str(v) for v in r] for r in rows]
        t = Table(data, repeatRows=1, colWidths=cw)
        t.setStyle(_ts(fs))
        return t

    def _sig(ent):
        sig = "\n\n\nSignature : ____________________"
        t = Table([[f"Responsable de l'entite ({ent})", "Le Chef d'exploitation"],
                   [sig, sig]], colWidths=[260, 260])
        t.setStyle(TableStyle([
            ("FONTNAME",      (0,0), (-1,0), "Helvetica-Bold"),
            ("FONTSIZE",      (0,0), (-1,-1), 9),
            ("ALIGN",         (0,0), (-1,-1), "CENTER"),
            ("VALIGN",        (0,0), (-1,-1), "TOP"),
            ("BOX",           (0,0), (-1,-1), 0.8, colors.grey),
            ("INNERGRID",     (0,0), (-1,-1), 0.5, colors.grey),
            ("TOPPADDING",    (0,0), (-1,-1), 8),
            ("BOTTOMPADDING", (0,0), (-1,-1), 8),
        ]))
        return t

    doc = BaseDocTemplate(str(out_path), pagesize=A4,
                          leftMargin=12, rightMargin=12, topMargin=14, bottomMargin=14)
    fp = Frame(doc.leftMargin, doc.bottomMargin, doc.width, doc.height, id="F1")
    fl = Frame(12, 12, landscape(A4)[0]-24, landscape(A4)[1]-24, id="F2")
    doc.addPageTemplates([
        PageTemplate(id="PORTRAIT",  frames=[fp], pagesize=A4),
        PageTemplate(id="LANDSCAPE", frames=[fl], pagesize=landscape(A4)),
    ])

    # Page 1: Facture
    els = [
        Paragraph("Attachement de Transport du Personnel", styles["Title"]),
        Paragraph(
            f"<b>Entite:</b> {entity} &nbsp;&nbsp; "
            f"<b>Prestataire:</b> {provider} &nbsp;&nbsp; "
            f"<b>Periode:</b> {period}",
            styles["Normal"]),
        Spacer(1, 8),
        Paragraph(f_title, styles["Heading2"]),
        Spacer(1, 6),
    ]
    t1 = _tbl(f_cols, f_rows, fs=9, cw=[210, 260, 90])
    n = len(f_rows) + 1
    if n >= 5:
        t1.setStyle(TableStyle([
            ("FONTNAME",   (0, n-4), (-1, n-1), "Helvetica-Bold"),
            ("BACKGROUND", (0, n-4), (-1, n-1), colors.beige),
        ]))
    els += [t1, Spacer(1, 14), _sig(entity)]

    # Page 2: Detail + Tarif
    els += [NextPageTemplate("LANDSCAPE"), PageBreak()]
    els += [Paragraph(d_title, styles["Heading2"]), Spacer(1, 6)]

    if d_rows:
        t2 = _tbl(d_cols, d_rows, fs=7)
        t2.setStyle(TableStyle([
            ("LEFTPADDING",  (0,0), (-1,-1), 2),
            ("RIGHTPADDING", (0,0), (-1,-1), 2),
            ("TOPPADDING",   (0,0), (-1,-1), 2),
            ("BOTTOMPADDING",(0,0), (-1,-1), 2),
        ]))
        els.append(t2)
    else:
        els.append(Paragraph("(Aucune donnee)", styles["Normal"]))

    if tarif_table:
        t_title, t_cols, t_rows = tarif_table
        els += [Spacer(1, 14), Paragraph(t_title, styles["Heading3"]), Spacer(1, 4)]
        els.append(_tbl(t_cols, t_rows, fs=8))

    doc.build(els)
    return out_path


def generate_attachment_pdf_for_selection(db_path: Path | str, period: str, entity: str,
                                          provider: str, out_pdf: Path | str | None = None) -> Path:
    from tempfile import NamedTemporaryFile
    import os

    db_path = Path(db_path)
    if out_pdf is None:
        out_pdf = Path.home() / "Downloads" / f"Attachement_{period}_{entity}_{provider}.pdf"
    out_pdf = Path(out_pdf)

    # Calculate and get fresh tables
    tmp = NamedTemporaryFile(suffix=".xlsx", delete=False)
    tmp.close()
    calculate_month_to_excel(db_path, period, tmp.name)
    try:
        os.unlink(tmp.name)
    except Exception:
        pass

    global _LAST_TABLES
    if _LAST_TABLES is None:
        raise RuntimeError("Tables non disponibles apres calcul")

    df_synth, df_ec, df_dp, df_sot = _LAST_TABLES

    # Build table data
    (f_title, f_cols, f_rows), (d_title, d_cols, d_rows) = build_pdf_tables_from_dfs(
        df_ec, df_dp, period, entity, provider, db_path=str(db_path))

    # Build tarif table
    tarif_table = _build_tariff_table_for_pdf(str(db_path), period, provider)

    # Render PDF directly — no DB cache
    _render_pdf_direct(str(out_pdf), entity, provider, period,
                       f_title, f_cols, f_rows,
                       d_title, d_cols, d_rows,
                       tarif_table=tarif_table)
    return out_pdf

def generate_global_pdf_for_entity(db_path, period: str, entity: str, out_pdf=None):
    """Global PDF: one facture page per provider, no detail page."""
    from reportlab.lib.pagesizes import A4
    from reportlab.platypus import (BaseDocTemplate, Frame, PageTemplate,
                                    Table, TableStyle, Paragraph, Spacer, PageBreak)
    from reportlab.lib import colors
    from reportlab.lib.styles import getSampleStyleSheet
    from tempfile import NamedTemporaryFile
    import os

    db_path = Path(db_path)
    if out_pdf is None:
        out_pdf = Path.home() / "Downloads" / f"Attachement_GLOBAL_{period}_{entity}.pdf"
    out_pdf = Path(out_pdf)

    tmp = NamedTemporaryFile(suffix=".xlsx", delete=False)
    tmp.close()
    calculate_month_to_excel(db_path, period, tmp.name)
    try: os.unlink(tmp.name)
    except: pass

    global _LAST_TABLES
    if _LAST_TABLES is None:
        raise RuntimeError("Tables non disponibles apres calcul")
    df_synth, df_ec, df_dp, df_sot = _LAST_TABLES

    try:    ec_ent = df_ec[df_ec["Entité"].astype(str).str.lower() == str(entity).lower()]
    except: ec_ent = df_ec[df_ec["Entite"].astype(str).str.lower() == str(entity).lower()]
    providers = sorted(ec_ent["Prestataire"].dropna().unique().tolist())

    styles = getSampleStyleSheet()
    doc = BaseDocTemplate(str(out_pdf), pagesize=A4,
                          leftMargin=20, rightMargin=20, topMargin=20, bottomMargin=20)
    frm = Frame(doc.leftMargin, doc.bottomMargin, doc.width, doc.height, id="F1")
    doc.addPageTemplates([PageTemplate(id="PORT", frames=[frm], pagesize=A4)])

    def _ts9():
        return TableStyle([
            ("BACKGROUND",    (0,0),(-1,0), colors.lightgrey),
            ("FONTNAME",      (0,0),(-1,0), "Helvetica-Bold"),
            ("FONTSIZE",      (0,0),(-1,-1),9),
            ("GRID",          (0,0),(-1,-1),0.5,colors.grey),
            ("ALIGN",         (0,0),(-1,-1),"CENTER"),
            ("VALIGN",        (0,0),(-1,-1),"MIDDLE"),
            ("ROWBACKGROUNDS",(0,1),(-1,-1),[colors.whitesmoke,colors.lightyellow]),
            ("LEFTPADDING",   (0,0),(-1,-1),3),("RIGHTPADDING",(0,0),(-1,-1),3),
            ("TOPPADDING",    (0,0),(-1,-1),2),("BOTTOMPADDING",(0,0),(-1,-1),2),
        ])

    els = []; first = True
    for provider in providers:
        if not first: els.append(PageBreak())
        first = False

        (ft, fc, fr), _ = build_pdf_tables_from_dfs(
            df_ec, df_dp, period, entity, provider, db_path=str(db_path))

        els += [
            Paragraph("Attachement de Transport du Personnel", styles["Title"]),
            Paragraph(
                f"<b>Entite:</b> {entity} &nbsp;&nbsp; "
                f"<b>Prestataire:</b> {provider} &nbsp;&nbsp; "
                f"<b>Periode:</b> {period}",
                styles["Normal"]),
            Spacer(1, 10), Paragraph(ft, styles["Heading2"]), Spacer(1, 6),
        ]
        data = [fc] + [["" if c is None else str(c) for c in row] for row in fr]
        n = len(data)
        t = Table(data, repeatRows=1, colWidths=[210, 260, 90])
        ts = _ts9()
        for ri in range(max(1, n-4), n):
            ts.add("FONTNAME",   (0,ri),(-1,ri),"Helvetica-Bold")
            ts.add("BACKGROUND", (0,ri),(-1,ri),colors.beige)
        t.setStyle(ts)
        els += [t, Spacer(1, 16)]

        sig = "\n\n\nSignature : ____________________"
        s = Table([[f"Responsable de l'entite ({entity})", "Le Chef d'exploitation"],
                   [sig, sig]], colWidths=[260, 260])
        s.setStyle(TableStyle([
            ("FONTNAME",      (0,0),(-1,0),"Helvetica-Bold"),
            ("FONTSIZE",      (0,0),(-1,-1),9),
            ("ALIGN",         (0,0),(-1,-1),"CENTER"),
            ("VALIGN",        (0,0),(-1,-1),"TOP"),
            ("BOX",           (0,0),(-1,-1),0.8,colors.grey),
            ("INNERGRID",     (0,0),(-1,-1),0.5,colors.grey),
            ("TOPPADDING",    (0,0),(-1,-1),8),
            ("BOTTOMPADDING", (0,0),(-1,-1),8),
        ]))
        els.append(s)

    doc.build(els)
    return out_pdf
