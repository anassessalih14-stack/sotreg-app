"""
SOTREG - Application Web (Flask)
=================================
Workflow 3 utilisateurs:
  user1 (saisie)   → saisit les données → télécharge sa DB → l'envoie à user2
  user2 (exports)  → importe la DB de user1 → génère Excel/PDF
  admin            → importe une DB consolidée → dashboard + tout

Fichier unique autonome. Lancez avec:
    python sotreg_web.py

Accès: http://127.0.0.1:5000
"""

import os, sys, io, json, shutil, sqlite3, tempfile
from pathlib import Path
from functools import wraps
from flask import (Flask, render_template_string, request, redirect,
                   url_for, session, flash, send_file, jsonify, g)

# ── Config ─────────────────────────────────────────────────────────────────────
BASE_DIR   = Path(__file__).resolve().parent
WORK_DB    = BASE_DIR / "sotreg_work.db"   # DB de saisie (user1)
TMPL_DB    = BASE_DIR / "sotreg_work.db"   # template = même fichier (pas de reset destructif)
UPLOAD_DIR = BASE_DIR / "uploads"
UPLOAD_DIR.mkdir(exist_ok=True)

# Rôles: saisie | exports | admin
# Modifiez ici les mots de passe avant de partager l'application
USERS = {
    "user1": ("pass1", "saisie"),   # Utilisateur 1 — saisie uniquement
    "user2": ("pass2", "exports"),  # Utilisateur 2 — exports/calculs uniquement
    "admin": ("admin", "admin"),    # Administrateur — tout
}

# Page d'accueil par rôle après login
ROLE_HOME = {
    "saisie":  "saisie",
    "exports": "exports",
    "admin":   "dashboard",
}

app = Flask(__name__)
app.secret_key = "sotreg-secret-2024"

# ── DB helpers ──────────────────────────────────────────────────────────────────
def get_db(path=None):
    db_path = path or str(WORK_DB)
    if "db" not in g or (path and g.get("db_path") != db_path):
        if hasattr(g, "db") and g.db:
            try: g.db.close()
            except: pass
        g.db = sqlite3.connect(db_path)
        g.db.row_factory = sqlite3.Row
        g.db_path = db_path
        _auto_migrate(g.db)
    return g.db

@app.teardown_appcontext
def close_db(e=None):
    db = g.pop("db", None)
    if db: db.close()

def _auto_migrate(conn):
    cur = conn.cursor()
    # tariffs: price_day
    cur.execute("PRAGMA table_info(tariffs)")
    tcols = [r[1] for r in cur.fetchall()]
    if "price_day" not in tcols:
        cur.execute("ALTER TABLE tariffs ADD COLUMN price_day REAL DEFAULT 0")
    # vehicle_period: dispo_type + nb_days + age_cat
    cur.execute("PRAGMA table_info(vehicle_period)")
    vcols = [r[1] for r in cur.fetchall()]
    for col, typ in [("dispo_type","TEXT DEFAULT 'permanent'"),("nb_days","REAL DEFAULT 0"),("age_cat","TEXT DEFAULT ''"),("zone","TEXT DEFAULT ''")]:
        if col not in vcols:
            cur.execute(f"ALTER TABLE vehicle_period ADD COLUMN {col} {typ}")
    cur.executescript("""
        CREATE TABLE IF NOT EXISTS entity(id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT UNIQUE NOT NULL);
        CREATE TABLE IF NOT EXISTS circuit(id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT UNIQUE NOT NULL);
        CREATE TABLE IF NOT EXISTS entity_circuit(entity_id INTEGER NOT NULL, circuit_id INTEGER NOT NULL, UNIQUE(entity_id,circuit_id));
    """)
    conn.commit()

def query(sql, params=(), path=None):
    return get_db(path).execute(sql, params).fetchall()

def execute(sql, params=(), path=None):
    db = get_db(path)
    db.execute(sql, params)
    db.commit()

def executemany(sql, rows, path=None):
    db = get_db(path)
    db.executemany(sql, rows)
    db.commit()

# ── Auth ────────────────────────────────────────────────────────────────────────
def login_required(roles=None):
    def decorator(f):
        @wraps(f)
        def wrapped(*args, **kwargs):
            if "user" not in session:
                return redirect(url_for("login"))
            if roles:
                role = session.get("role","")
                if role != "admin" and role not in roles:
                    flash("Accès non autorisé.", "error")
                    return redirect(url_for("dashboard"))
            return f(*args, **kwargs)
        return wrapped
    return decorator

# ── Routes Auth ─────────────────────────────────────────────────────────────────
@app.route("/", methods=["GET","POST"])
def login():
    if "user" in session:
        return redirect(url_for(ROLE_HOME[session["role"]]))
    error = None
    if request.method == "POST":
        u = request.form.get("username","").strip()
        p = request.form.get("password","").strip()
        if u in USERS and USERS[u][0] == p:
            session["user"] = u
            session["role"] = USERS[u][1]
            return redirect(url_for(ROLE_HOME[USERS[u][1]]))
        error = "Identifiants invalides."
    return render_template_string(LOGIN_HTML, error=error)

@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))

# ── KPI API (admin only, uses session db) ──────────────────────────────────────
@app.route("/api/kpi")
@login_required(roles=["admin"])
def api_kpi():
    db_path = session.get("admin_db", str(WORK_DB))
    def q(sql): return query(sql, path=db_path)
    rows_km = q("""
        SELECT period, provider,
               COUNT(*) as nb_lines,
               COALESCE(SUM(rotation_total),0) as total_rot,
               COALESCE(SUM(km_per_rotation * rotation_total),0) as total_km
        FROM fact_lines WHERE period != 'base'
        GROUP BY period, provider ORDER BY period, provider
    """)
    rows_sot = q("""
        SELECT period, provider,
               COUNT(*) as nb_lines,
               COALESCE(SUM(rotation_total),0) as total_rot,
               COALESCE(SUM(km_per_rotation * rotation_total),0) as total_km
        FROM fact_lines_sotreg WHERE period != 'base'
        GROUP BY period, provider ORDER BY period, provider
    """)
    rows_veh = q("""
        SELECT period, provider, COUNT(DISTINCT mle_car) as nb_veh,
               COALESCE(SUM(km_compteur),0) as total_km_veh
        FROM vehicle_period WHERE period != 'base'
        GROUP BY period, provider ORDER BY period, provider
    """)
    rows_fact = q("""
        SELECT vp.period, vp.provider,
               COALESCE(SUM(vp.km_compteur * COALESCE(t.price_km,0)),0) as frais_km,
               COALESCE(COUNT(DISTINCT vp.mle_car) * MAX(COALESCE(t.price_mise_dispo,0)),0) as frais_dispo
        FROM vehicle_period vp
        LEFT JOIN tariffs t ON t.period=vp.period AND t.provider=vp.provider COLLATE NOCASE
                             AND t.vehicle_type=vp.vehicle_type COLLATE NOCASE
        WHERE vp.period != 'base'
        GROUP BY vp.period, vp.provider ORDER BY vp.period, vp.provider
    """)
    # SOTREG facturation séparée (basée sur fact_lines_sotreg)
    rows_fact_sotreg = q("""
        SELECT fs.period, fs.provider,
               COALESCE(SUM(fs.km_per_rotation * fs.rotation_total * COALESCE(t.price_km,0)),0) as frais_km,
               COALESCE(SUM(fs.nb_vehicles * COALESCE(t.price_mise_dispo,0)),0) as frais_dispo
        FROM fact_lines_sotreg fs
        LEFT JOIN tariffs t ON t.period=fs.period AND t.provider=fs.provider COLLATE NOCASE
                             AND t.vehicle_type=fs.vehicle_type COLLATE NOCASE
        WHERE fs.period != 'base'
        GROUP BY fs.period, fs.provider ORDER BY fs.period, fs.provider
    """)
    rows_entity = q("""
        SELECT fl.period, fl.entity,
               ROUND(SUM(fl.km_per_rotation * fl.rotation_total * COALESCE(t.price_km,0)),2) as fact_ht
        FROM fact_lines fl
        LEFT JOIN tariffs t ON t.period=fl.period AND t.provider=fl.provider COLLATE NOCASE
                             AND t.vehicle_type=fl.vehicle_type COLLATE NOCASE
        WHERE fl.period != 'base' AND fl.entity IS NOT NULL AND fl.entity != 'NAVETTE'
        GROUP BY fl.period, fl.entity ORDER BY fl.period, fl.entity
    """)
    def to_list(rows): return [dict(r) for r in rows]
    return jsonify({
        "km_rotations": to_list(rows_km) + to_list(rows_sot),
        "vehicles":     to_list(rows_veh),
        "facturation":  to_list(rows_fact) + to_list(rows_fact_sotreg),
        "entity_fact":  to_list(rows_entity),
    })

@app.route("/api/entity_circuits")
@login_required(roles=["saisie","admin"])
def api_entity_circuits():
    """Retourne tous les couples entité/circuit avec leur valeur KM la plus récente."""
    rows = query("""
        WITH all_lines AS (
          SELECT id, period, entity, circuit, km_per_rotation FROM fact_lines
          UNION ALL
          SELECT id, period, entity, circuit, km_per_rotation FROM fact_lines_sotreg
        ), ranked AS (
          SELECT entity, circuit, km_per_rotation, period AS latest_period, id AS latest_id,
                 ROW_NUMBER() OVER (
                   PARTITION BY UPPER(TRIM(entity)), UPPER(TRIM(circuit))
                   ORDER BY CASE WHEN period='base' THEN 0 ELSE 1 END DESC,
                            period DESC, id DESC
                 ) AS rn
          FROM all_lines
          WHERE entity IS NOT NULL AND TRIM(entity) != ''
            AND circuit IS NOT NULL AND TRIM(circuit) != ''
        )
        SELECT entity, circuit, COALESCE(km_per_rotation,0) AS km_per_rotation
        FROM ranked WHERE rn=1 ORDER BY latest_period DESC, latest_id DESC
    """)
    mapping = {}
    for r in rows:
        ent = r["entity"]; cir = r["circuit"]
        if ent not in mapping: mapping[ent] = {}
        mapping[ent][cir] = r["km_per_rotation"]
    return jsonify(mapping)

@app.route("/api/circuit_km")
@login_required(roles=["saisie","admin"])
def api_circuit_km():
    """Mapping circuit → km_per_rotation — dernière valeur saisie (MAX id)."""
    rows = query("""
        SELECT circuit, km_per_rotation
        FROM fact_lines
        WHERE period = (SELECT MAX(period) FROM fact_lines WHERE period != 'base')
          AND km_per_rotation > 0
          AND id IN (
              SELECT MAX(id) FROM fact_lines
              WHERE period = (SELECT MAX(period) FROM fact_lines WHERE period != 'base')
                AND km_per_rotation > 0
              GROUP BY circuit
          )
        ORDER BY circuit
    """)
    mapping = {r["circuit"]: r["km_per_rotation"] for r in rows}
    return jsonify(mapping)

@app.route("/api/provider_vtypes")
@login_required(roles=["saisie","admin"])
def api_provider_vtypes():
    """Retourne les types de véhicule disponibles par prestataire."""
    rows = query("""
        SELECT provider, vehicle_type FROM (
            SELECT provider, vehicle_type FROM fact_lines WHERE vehicle_type != ''
            UNION
            SELECT provider, vehicle_type FROM fact_lines_sotreg WHERE vehicle_type != ''
        ) GROUP BY provider, vehicle_type ORDER BY provider, vehicle_type
    """)
    mapping = {}
    for r in rows:
        p = r["provider"]
        if p not in mapping:
            mapping[p] = []
        mapping[p].append(r["vehicle_type"])
    return jsonify(mapping)

@app.route("/api/lists")
@login_required(roles=["saisie","admin"])
def api_lists():
    """Listes pour l'autocomplete de la saisie."""
    entities = [r[0] for r in query(
        "SELECT DISTINCT entity FROM fact_lines WHERE entity IS NOT NULL AND entity != '' "
        "UNION SELECT DISTINCT entity FROM fact_lines_sotreg WHERE entity IS NOT NULL AND entity != '' "
        "ORDER BY entity")]
    circuits = [r[0] for r in query(
        "SELECT DISTINCT circuit FROM fact_lines WHERE circuit IS NOT NULL AND circuit != '' "
        "UNION SELECT DISTINCT circuit FROM fact_lines_sotreg WHERE circuit IS NOT NULL AND circuit != '' "
        "ORDER BY circuit")]
    mles = [r[0] for r in query(
        "SELECT DISTINCT mle_car FROM fact_lines WHERE mle_car IS NOT NULL AND mle_car != '' ORDER BY mle_car")]
    chauffeurs = [r[0] for r in query(
        "SELECT DISTINCT chauffeur FROM fact_lines WHERE chauffeur IS NOT NULL AND chauffeur != '' ORDER BY chauffeur")]
    return jsonify({
        "entities":   entities,
        "circuits":   circuits,
        "mles":       mles,
        "chauffeurs": chauffeurs,
    })

@app.route("/dashboard")
@login_required(roles=["admin"])
def dashboard():
    db_path = session.get("admin_db", str(WORK_DB))
    try:
        periods   = [r[0] for r in query(
            "SELECT DISTINCT period FROM fact_lines WHERE period != 'base' "
            "UNION SELECT DISTINCT period FROM fact_lines_sotreg WHERE period != 'base' "
            "ORDER BY period DESC LIMIT 24", path=db_path)]
        providers = [r[0] for r in query(
            "SELECT DISTINCT provider FROM fact_lines "
            "UNION SELECT DISTINCT provider FROM fact_lines_sotreg ORDER BY provider", path=db_path)]
        nb_lines = query("SELECT COUNT(*) as n FROM fact_lines", path=db_path)[0]["n"]
        nb_veh   = query("SELECT COUNT(DISTINCT mle_car) as n FROM vehicle_period WHERE period != 'base'", path=db_path)[0]["n"]
    except Exception as e:
        flash(str(e), "error")
        periods, providers, nb_lines, nb_veh = [], [], 0, 0
    db_name = Path(db_path).name
    return render_template_string(DASHBOARD_HTML, periods=periods,
                                  providers=providers, nb_lines=nb_lines,
                                  nb_veh=nb_veh, db_name=db_name,
                                  user=session["user"])

@app.route("/admin/upload_db", methods=["POST"])
@login_required(roles=["admin"])
def admin_upload_db():
    f = request.files.get("dbfile")
    if not f or not f.filename.endswith(".db"):
        flash("Fichier .db requis.", "error")
        return redirect(url_for("dashboard"))
    dest = UPLOAD_DIR / f.filename
    f.save(str(dest))
    session["admin_db"] = str(dest)
    flash(f"DB chargée: {f.filename} — dashboard mis à jour.", "success")
    return redirect(url_for("dashboard"))

@app.route("/admin/reset_db_view")
@login_required(roles=["admin"])
def admin_reset_db_view():
    session.pop("admin_db", None)
    flash("Retour à la DB locale.", "success")
    return redirect(url_for("dashboard"))

# ── Module 1: Saisie ────────────────────────────────────────────────────────────
@app.route("/saisie")
@login_required(roles=["saisie","admin"])
def saisie():
    # Non-SOTREG providers from fact_lines
    providers = [r[0] for r in query("SELECT DISTINCT provider FROM fact_lines WHERE provider != '' ORDER BY provider")]
    # Add SOTREG if not already present
    if "SOTREG" not in [p.upper() for p in providers]:
        providers = ["SOTREG"] + providers

    # All vehicle types: union of fact_lines + fact_lines_sotreg
    vtypes_fl  = [r[0] for r in query("SELECT DISTINCT vehicle_type FROM fact_lines WHERE vehicle_type != '' ORDER BY vehicle_type")]
    vtypes_sot = [r[0] for r in query("SELECT DISTINCT vehicle_type FROM fact_lines_sotreg WHERE vehicle_type != '' ORDER BY vehicle_type")]
    vtypes = list(dict.fromkeys(vtypes_sot + vtypes_fl))  # SOTREG types first, then others, deduplicated

    # All periods available (for the selector dropdown)
    periods_db = [r[0] for r in query(
        "SELECT DISTINCT period FROM fact_lines WHERE period != 'base' "
        "UNION SELECT DISTINCT period FROM fact_lines_sotreg WHERE period != 'base' "
        "ORDER BY period DESC")]

    return render_template_string(SAISIE_HTML, providers=providers, vtypes=vtypes,
                                  periods_db=periods_db, user=session["user"])

@app.route("/api/saisie/load")
@login_required(roles=["saisie","admin"])
def saisie_load():
    provider = request.args.get("provider","").strip()
    vtype    = request.args.get("vtype","").strip()
    period   = request.args.get("period","").strip()
    if not provider or not vtype or not period:
        return jsonify({"error": "Paramètres manquants"}), 400

    # Ensure month data exists
    _ensure_month_data(provider, vtype, period)

    sotreg = provider.upper() == "SOTREG"

    if sotreg:
        lines = [dict(r) for r in query(
            "SELECT entity, circuit, nb_vehicles, km_per_rotation, rotation_total FROM fact_lines_sotreg WHERE period=? AND provider=? COLLATE NOCASE AND vehicle_type=? COLLATE NOCASE",
            (period, provider, vtype))]
        tariff_rows = [dict(r) for r in query(
            "SELECT billing_mode, price_mise_dispo, price_km FROM tariffs WHERE period=? AND provider=? COLLATE NOCASE AND vehicle_type=? COLLATE NOCASE LIMIT 1",
            (period, provider, vtype))]
        return jsonify({"sotreg": True, "lines": lines, "tariff": tariff_rows[0] if tariff_rows else {}})
    else:
        lines = [dict(r) for r in query(
            "SELECT chauffeur, mle_car, entity, circuit, km_per_rotation, rotation_total FROM fact_lines WHERE period=? AND provider=? COLLATE NOCASE AND vehicle_type=? COLLATE NOCASE",
            (period, provider, vtype))]
        veh = [dict(r) for r in query(
            "SELECT mle_car, km_compteur, km_forfait, km_supp, age_cat, dispo_type, nb_days, COALESCE(zone,'') as zone FROM vehicle_period WHERE period=? AND provider=? COLLATE NOCASE AND vehicle_type=? COLLATE NOCASE",
            (period, provider, vtype))]
        tariff_rows = [dict(r) for r in query(
            "SELECT billing_mode, price_mise_dispo, price_km, price_km_supp, price_journalier, price_day, km_forfait_value, age_cat FROM tariffs WHERE period=? AND provider=? COLLATE NOCASE AND vehicle_type=? COLLATE NOCASE",
            (period, provider, vtype))]
        return jsonify({"sotreg": False, "lines": lines, "vehicles": veh, "tariffs": tariff_rows})

@app.route("/api/saisie/save", methods=["POST"])
@login_required(roles=["saisie","admin"])
def saisie_save():
    data     = request.get_json()
    provider = data.get("provider","").strip()
    vtype    = data.get("vtype","").strip()
    period   = data.get("period","").strip()
    sotreg   = provider.upper() == "SOTREG"

    if not provider or not vtype or not period:
        return jsonify({"error": "Paramètres manquants"}), 400

    db = get_db()

    # Save tariffs
    tariffs_data = data.get("tariffs", [])
    db.execute("DELETE FROM tariffs WHERE period=? AND provider=? COLLATE NOCASE AND vehicle_type=? COLLATE NOCASE",
               (period, provider, vtype))
    for t in tariffs_data:
        db.execute("""INSERT INTO tariffs(period,provider,vehicle_type,age_cat,billing_mode,
                      price_mise_dispo,price_km,price_km_supp,price_journalier,price_day,km_forfait_value)
                      VALUES(?,?,?,?,?,?,?,?,?,?,?)""",
                   (period, provider, vtype,
                    t.get("age_cat",""), t.get("billing_mode",""),
                    float(t.get("price_mise_dispo",0) or 0),
                    float(t.get("price_km",0) or 0),
                    float(t.get("price_km_supp",0) or 0),
                    float(t.get("price_journalier",0) or 0),
                    float(t.get("price_day",0) or 0),
                    float(t.get("km_forfait_value",0) or 0)))

    if sotreg:
        db.execute("DELETE FROM fact_lines_sotreg WHERE period=? AND provider=? COLLATE NOCASE AND vehicle_type=? COLLATE NOCASE",
                   (period, provider, vtype))
        for row in data.get("lines", []):
            if not row.get("circuit","").strip():
                continue
            _upsert_entity(db, row.get("entity",""))
            _upsert_circuit(db, row.get("circuit",""))
            db.execute("INSERT INTO fact_lines_sotreg(period,provider,vehicle_type,entity,circuit,nb_vehicles,km_per_rotation,rotation_total) VALUES(?,?,?,?,?,?,?,?)",
                       (period, provider, vtype,
                        row.get("entity","") or None,
                        row.get("circuit",""),
                        int(float(row.get("nb_vehicles",0) or 0)),
                        float(row.get("km_per_rotation",0) or 0),
                        float(row.get("rotation_total",0) or 0)))
    else:
        db.execute("DELETE FROM fact_lines WHERE period=? AND provider=? COLLATE NOCASE AND vehicle_type=? COLLATE NOCASE",
                   (period, provider, vtype))
        for row in data.get("lines", []):
            if not row.get("mle_car","").strip() or not row.get("circuit","").strip():
                continue
            _upsert_entity(db, row.get("entity",""))
            _upsert_circuit(db, row.get("circuit",""))
            db.execute("INSERT INTO fact_lines(period,provider,vehicle_type,chauffeur,mle_car,entity,circuit,km_per_rotation,rotation_total) VALUES(?,?,?,?,?,?,?,?,?)",
                       (period, provider, vtype,
                        row.get("chauffeur",""),
                        row.get("mle_car",""),
                        row.get("entity","") or None,
                        row.get("circuit",""),
                        float(row.get("km_per_rotation",0) or 0),
                        float(row.get("rotation_total",0) or 0)))

        db.execute("DELETE FROM vehicle_period WHERE period=? AND provider=? COLLATE NOCASE AND vehicle_type=? COLLATE NOCASE",
                   (period, provider, vtype))
        for v in data.get("vehicles", []):
            if not v.get("mle_car","").strip():
                continue
            db.execute("INSERT INTO vehicle_period(period,provider,vehicle_type,mle_car,km_compteur,km_forfait,km_supp,age_cat,dispo_type,nb_days,zone) VALUES(?,?,?,?,?,?,?,?,?,?,?)",
                       (period, provider, vtype,
                        v.get("mle_car",""),
                        float(v.get("km_compteur",0) or 0),
                        float(v.get("km_forfait",0) or 0),
                        float(v.get("km_supp",0) or 0),
                        v.get("age_cat",""),
                        v.get("dispo_type","permanent"),
                        float(v.get("nb_days",0) or 0),
                        v.get("zone","")))

    db.commit()
    return jsonify({"ok": True, "msg": f"Sauvegardé: {period} / {provider} / {vtype}"})

def _upsert_entity(db, name):
    name = (name or "").strip()
    if name:
        db.execute("INSERT OR IGNORE INTO entity(name) VALUES(?)", (name,))

def _upsert_circuit(db, name):
    name = (name or "").strip()
    if name:
        db.execute("INSERT OR IGNORE INTO circuit(name) VALUES(?)", (name,))

def _ensure_month_data(provider, vtype, month):
    if not month or month == "base":
        return
    db = get_db()
    sotreg = provider.upper() == "SOTREG"
    if not sotreg:
        row = db.execute("SELECT 1 FROM fact_lines WHERE period=? AND provider=? COLLATE NOCASE AND vehicle_type=? COLLATE NOCASE LIMIT 1", (month, provider, vtype)).fetchone()
        if not row:
            db.execute("INSERT INTO fact_lines(period,provider,vehicle_type,chauffeur,mle_car,entity,circuit,km_per_rotation,rotation_total) SELECT ?,provider,vehicle_type,chauffeur,mle_car,entity,circuit,km_per_rotation,rotation_total FROM fact_lines WHERE period='base' AND provider=? COLLATE NOCASE AND vehicle_type=? COLLATE NOCASE", (month, provider, vtype))
        db.execute("INSERT OR IGNORE INTO vehicle_period(period,provider,vehicle_type,mle_car,km_compteur,km_forfait,km_supp,age_cat,dispo_type,nb_days) SELECT ?,provider,vehicle_type,mle_car,COALESCE(km_compteur,0),COALESCE(km_forfait,0),COALESCE(km_supp,0),COALESCE(age_cat,''),COALESCE(dispo_type,'permanent'),COALESCE(nb_days,0) FROM vehicle_period WHERE period='base' AND provider=? COLLATE NOCASE AND vehicle_type=? COLLATE NOCASE", (month, provider, vtype))
    row2 = db.execute("SELECT 1 FROM fact_lines_sotreg WHERE period=? AND provider=? COLLATE NOCASE AND vehicle_type=? COLLATE NOCASE LIMIT 1", (month, provider, vtype)).fetchone()
    if not row2:
        db.execute("INSERT INTO fact_lines_sotreg(period,provider,vehicle_type,entity,circuit,nb_vehicles,km_per_rotation,rotation_total) SELECT ?,provider,vehicle_type,entity,circuit,nb_vehicles,km_per_rotation,rotation_total FROM fact_lines_sotreg WHERE period='base' AND provider=? COLLATE NOCASE AND vehicle_type=? COLLATE NOCASE", (month, provider, vtype))
    row3 = db.execute("SELECT 1 FROM tariffs WHERE period=? AND provider=? COLLATE NOCASE AND vehicle_type=? COLLATE NOCASE LIMIT 1", (month, provider, vtype)).fetchone()
    if not row3:
        db.execute("INSERT INTO tariffs(period,provider,vehicle_type,age_cat,billing_mode,price_mise_dispo,price_km,price_km_supp,price_journalier,price_day,km_forfait_value) SELECT ?,provider,vehicle_type,age_cat,billing_mode,price_mise_dispo,price_km,price_km_supp,price_journalier,COALESCE(price_day,0),COALESCE(km_forfait_value,0) FROM tariffs WHERE period IN ('DEFAULT','base') AND provider=? COLLATE NOCASE AND vehicle_type=? COLLATE NOCASE", (month, provider, vtype))
    db.commit()

# ── Module 2: Exports ───────────────────────────────────────────────────────────
@app.route("/exports")
@login_required(roles=["exports","admin"])
def exports():
    db_path = session.get("export_db", str(WORK_DB))
    role    = session.get("role","")
    try:
        periods   = [r[0] for r in query("SELECT DISTINCT period FROM fact_lines WHERE period != 'base' ORDER BY period DESC", path=db_path)]
        entities  = [r[0] for r in query("SELECT DISTINCT entity FROM fact_lines WHERE entity IS NOT NULL AND UPPER(entity) != 'NAVETTE' ORDER BY entity", path=db_path)]
        providers = [r[0] for r in query("SELECT DISTINCT provider FROM fact_lines ORDER BY provider", path=db_path)]
        if "SOTREG" not in [p.upper() for p in providers]:
            providers.append("SOTREG")
    except Exception as e:
        periods, entities, providers = [], [], []
        flash(str(e), "error")
    db_name = Path(db_path).name
    return render_template_string(EXPORTS_HTML, periods=periods, entities=entities,
                                  providers=providers, user=session["user"],
                                  db_name=db_name, role=role)

@app.route("/exports/upload_db", methods=["POST"])
@login_required(roles=["exports","admin"])
def upload_db():
    f = request.files.get("dbfile")
    if not f or not f.filename.endswith(".db"):
        flash("Fichier .db requis.", "error")
        return redirect(url_for("exports"))
    dest = UPLOAD_DIR / f.filename
    f.save(str(dest))
    session["export_db"] = str(dest)
    flash(f"DB chargée: {f.filename}", "success")
    return redirect(url_for("exports"))

@app.route("/exports/reset_db")
@login_required(roles=["exports","admin"])
def reset_export_db():
    session.pop("export_db", None)
    return redirect(url_for("exports"))

@app.route("/exports/excel_mois")
@login_required(roles=["exports","admin"])
def export_excel_mois():
    period   = request.args.get("period","").strip()
    db_path  = session.get("export_db", str(WORK_DB))
    if not period:
        flash("Période requise.", "error")
        return redirect(url_for("exports"))
    try:
        sys.path.insert(0, str(BASE_DIR))
        from calc_engine import calculate_month_to_excel
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
            tmp = f.name
        calculate_month_to_excel(db_path, period, tmp)
        return send_file(tmp, as_attachment=True, download_name=f"facturation_{period}.xlsx",
                         mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    except Exception as e:
        flash(f"Erreur: {e}", "error")
        return redirect(url_for("exports"))

@app.route("/exports/attachment_excel")
@login_required(roles=["exports","admin"])
def export_attachment_excel():
    period   = request.args.get("period","").strip()
    entity   = request.args.get("entity","").strip()
    provider = request.args.get("provider","").strip()
    db_path  = session.get("export_db", str(WORK_DB))
    if not all([period, entity, provider]):
        flash("Période, entité et prestataire requis.", "error")
        return redirect(url_for("exports"))
    try:
        sys.path.insert(0, str(BASE_DIR))
        from excel_attachments import generate_provider_attachment_excel
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
            tmp = f.name
        generate_provider_attachment_excel(db_path, period, entity, provider, tmp)
        fname = f"Attachement_{period}_{entity}_{provider}.xlsx"
        return send_file(tmp, as_attachment=True, download_name=fname,
                         mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    except Exception as e:
        flash(f"Erreur: {e}", "error")
        return redirect(url_for("exports"))

@app.route("/exports/attachment_global")
@login_required(roles=["exports","admin"])
def export_attachment_global():
    period  = request.args.get("period","").strip()
    entity  = request.args.get("entity","").strip()
    db_path = session.get("export_db", str(WORK_DB))
    if not all([period, entity]):
        flash("Période et entité requises.", "error")
        return redirect(url_for("exports"))
    try:
        sys.path.insert(0, str(BASE_DIR))
        from excel_attachments import generate_global_attachment_excel
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
            tmp = f.name
        generate_global_attachment_excel(db_path, period, entity, tmp)
        fname = f"Attachement_GLOBAL_{period}_{entity}.xlsx"
        return send_file(tmp, as_attachment=True, download_name=fname,
                         mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    except Exception as e:
        flash(f"Erreur: {e}", "error")
        return redirect(url_for("exports"))

@app.route("/exports/pdf_global")
@login_required(roles=["exports","admin"])
def export_pdf_global():
    period  = request.args.get("period","").strip()
    entity  = request.args.get("entity","").strip()
    db_path = session.get("export_db", str(WORK_DB))
    if not all([period, entity]):
        flash("Période et entité requises.", "error")
        return redirect(url_for("exports"))
    try:
        sys.path.insert(0, str(BASE_DIR))
        from calc_engine import generate_global_pdf_for_entity
        with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as f:
            tmp = f.name
        generate_global_pdf_for_entity(db_path, period, entity, tmp)
        fname = f"Attachement_GLOBAL_{period}_{entity}.pdf"
        return send_file(tmp, as_attachment=True, download_name=fname, mimetype="application/pdf")
    except Exception as e:
        flash(f"Erreur PDF Global: {e}", "error")
        return redirect(url_for("exports"))

@app.route("/exports/pdf")
@login_required(roles=["exports","admin"])
def export_pdf():
    period   = request.args.get("period","").strip()
    entity   = request.args.get("entity","").strip()
    provider = request.args.get("provider","").strip()
    db_path  = session.get("export_db", str(WORK_DB))
    if not all([period, entity, provider]):
        flash("Période, entité et prestataire requis.", "error")
        return redirect(url_for("exports"))
    try:
        sys.path.insert(0, str(BASE_DIR))
        from calc_engine import generate_attachment_pdf_for_selection
        with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as f:
            tmp = f.name
        generate_attachment_pdf_for_selection(db_path, period, entity, provider, tmp)
        fname = f"Attachement_{period}_{entity}_{provider}.pdf"
        return send_file(tmp, as_attachment=True, download_name=fname, mimetype="application/pdf")
    except Exception as e:
        flash(f"Erreur PDF: {e}", "error")
        return redirect(url_for("exports"))

# ── Admin: reset DB template ───────────────────────────────────────────────────
@app.route("/admin/reset_db", methods=["POST"])
@login_required(roles=["admin"])
def admin_reset_db():
    if not TMPL_DB.exists():
        flash("Template DB introuvable.", "error")
        return redirect(url_for("dashboard"))
    close_db()
    shutil.copy(str(TMPL_DB), str(WORK_DB))
    flash("Base de données de saisie réinitialisée depuis le template.", "success")
    return redirect(url_for("dashboard"))

# ── User1: télécharger sa DB de saisie ────────────────────────────────────────
@app.route("/saisie/download_db")
@login_required(roles=["saisie","admin"])
def saisie_download_db():
    """User1 télécharge sa DB pour l'envoyer à user2 ou admin."""
    if not WORK_DB.exists():
        flash("DB introuvable.", "error")
        return redirect(url_for("saisie"))
    return send_file(str(WORK_DB), as_attachment=True,
                     download_name=f"sotreg_saisie.db",
                     mimetype="application/octet-stream")

@app.route("/saisie/upload_db", methods=["POST"])
@login_required(roles=["saisie","admin"])
def saisie_upload_db():
    """Charge la dernière DB sauvegardée par user1 pour reprendre la saisie."""
    uploaded = request.files.get("dbfile")
    if not uploaded or not uploaded.filename.lower().endswith(".db"):
        flash("Veuillez sélectionner un fichier .db.", "error")
        return redirect(url_for("saisie"))

    tmp_path = None
    try:
        with tempfile.NamedTemporaryFile(suffix=".db", delete=False) as tmp:
            tmp_path = tmp.name
            uploaded.save(tmp_path)

        # Contrôle d'intégrité et vérification des tables avant tout remplacement.
        check = sqlite3.connect(tmp_path)
        integrity = check.execute("PRAGMA integrity_check").fetchone()[0]
        tables = {r[0] for r in check.execute(
            "SELECT name FROM sqlite_master WHERE type='table'"
        ).fetchall()}
        check.close()
        required = {"fact_lines", "fact_lines_sotreg", "vehicle_period", "tariffs"}
        missing = required - tables
        if integrity != "ok":
            raise ValueError("La base SQLite est endommagée.")
        if missing:
            raise ValueError("Base incompatible — tables manquantes : " + ", ".join(sorted(missing)))

        close_db()
        shutil.copyfile(tmp_path, str(WORK_DB))
        # Applique les petites migrations prévues par l'application.
        db = sqlite3.connect(str(WORK_DB))
        _auto_migrate(db)
        db.close()
        flash("Base chargée avec succès. Vous pouvez continuer la saisie.", "success")
    except Exception as e:
        flash(f"Chargement refusé : {e}", "error")
    finally:
        if tmp_path:
            try: os.unlink(tmp_path)
            except OSError: pass
    return redirect(url_for("saisie"))

# ══════════════════════════════════════════════════════════════════════════════
# TEMPLATES HTML
# ══════════════════════════════════════════════════════════════════════════════

_BASE_STYLE = """
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<link href="https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@300;400;500;600&family=IBM+Plex+Mono:wght@400;500&display=swap" rel="stylesheet">
<style>
  :root {
    --bg:       #0f1117;
    --bg2:      #161b27;
    --bg3:      #1e2535;
    --border:   #2a3347;
    --text:     #e2e8f0;
    --muted:    #7c8fa6;
    --accent:   #3b82f6;
    --accent2:  #60a5fa;
    --success:  #22c55e;
    --warning:  #f59e0b;
    --danger:   #ef4444;
    --radius:   8px;
  }
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: 'IBM Plex Sans', sans-serif; background: var(--bg); color: var(--text);
         font-size: 14px; line-height: 1.6; }
  a { color: var(--accent2); text-decoration: none; }
  a:hover { text-decoration: underline; }

  /* Nav */
  .nav { background: var(--bg2); border-bottom: 1px solid var(--border);
         padding: 0 24px; display: flex; align-items: center; gap: 8px; height: 52px; }
  .nav-brand { font-family: 'IBM Plex Mono', monospace; font-weight: 500; font-size: 15px;
               color: var(--accent2); letter-spacing: 1px; margin-right: auto; }
  .nav a { color: var(--muted); font-size: 13px; padding: 6px 12px; border-radius: var(--radius);
           transition: all .15s; }
  .nav a:hover, .nav a.active { background: var(--bg3); color: var(--text); text-decoration: none; }
  .nav .logout { color: var(--danger) !important; }

  /* Layout */
  .container { max-width: 1400px; margin: 0 auto; padding: 24px; }
  .page-title { font-size: 20px; font-weight: 600; color: var(--text); margin-bottom: 20px;
                border-left: 3px solid var(--accent); padding-left: 12px; }

  /* Cards */
  .card { background: var(--bg2); border: 1px solid var(--border); border-radius: var(--radius); padding: 20px; }
  .card-title { font-size: 13px; font-weight: 500; color: var(--muted); text-transform: uppercase;
                letter-spacing: .8px; margin-bottom: 12px; }

  /* Forms */
  .form-row { display: flex; gap: 12px; flex-wrap: wrap; align-items: flex-end; margin-bottom: 16px; }
  .form-group { display: flex; flex-direction: column; gap: 5px; min-width: 160px; }
  label { font-size: 12px; color: var(--muted); font-weight: 500; text-transform: uppercase; letter-spacing: .5px; }
  select, input[type=text], input[type=password] {
    background: var(--bg3); border: 1px solid var(--border); color: var(--text);
    padding: 8px 12px; border-radius: var(--radius); font-size: 13px; font-family: inherit;
    transition: border-color .15s; outline: none; }
  select:focus, input:focus { border-color: var(--accent); }

  /* Buttons */
  .btn { display: inline-flex; align-items: center; gap: 6px; padding: 8px 16px;
         border-radius: var(--radius); font-size: 13px; font-weight: 500; cursor: pointer;
         border: 1px solid transparent; transition: all .15s; font-family: inherit; }
  .btn-primary { background: var(--accent); color: #fff; border-color: var(--accent); }
  .btn-primary:hover { background: var(--accent2); }
  .btn-secondary { background: var(--bg3); color: var(--text); border-color: var(--border); }
  .btn-secondary:hover { border-color: var(--accent); color: var(--accent2); }
  .btn-success { background: #166534; color: #86efac; border-color: #166534; }
  .btn-success:hover { background: #15803d; }
  .btn-danger { background: #7f1d1d; color: #fca5a5; border-color: #7f1d1d; }
  .btn-danger:hover { background: #991b1b; }
  .btn-sm { padding: 5px 10px; font-size: 12px; }
  .btn-icon { padding: 6px 8px; }

  /* Tables */
  .table-wrap { overflow-x: auto; border-radius: var(--radius); border: 1px solid var(--border); }
  table { width: 100%; border-collapse: collapse; font-size: 13px; }
  thead tr { background: var(--bg3); }
  th { padding: 10px 12px; text-align: left; font-size: 11px; text-transform: uppercase;
       letter-spacing: .7px; color: var(--muted); font-weight: 500; white-space: nowrap; }
  td { padding: 8px 12px; border-top: 1px solid var(--border); vertical-align: middle; }
  tbody tr:hover { background: var(--bg3); }
  td input { background: transparent; border: none; color: var(--text); width: 100%;
             font-family: 'IBM Plex Mono', monospace; font-size: 12px; padding: 2px 4px; outline: none; }
  td input:focus { background: var(--bg); border-radius: 4px; border: 1px solid var(--accent); }

  /* Alerts */
  .alert { padding: 10px 16px; border-radius: var(--radius); margin-bottom: 16px;
           font-size: 13px; border-left: 3px solid; }
  .alert-success { background: #052e16; color: #86efac; border-color: var(--success); }
  .alert-error   { background: #450a0a; color: #fca5a5; border-color: var(--danger); }
  .alert-info    { background: #172554; color: #93c5fd; border-color: var(--accent); }

  /* Stats */
  .stats-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(180px,1fr)); gap: 16px; margin-bottom: 24px; }
  .stat-card { background: var(--bg2); border: 1px solid var(--border); border-radius: var(--radius);
               padding: 16px; }
  .stat-value { font-size: 28px; font-weight: 600; color: var(--accent2); font-family: 'IBM Plex Mono', monospace; }
  .stat-label { font-size: 11px; color: var(--muted); text-transform: uppercase; letter-spacing: .7px; margin-top: 4px; }

  /* Badge */
  .badge { display: inline-block; padding: 2px 8px; border-radius: 100px; font-size: 11px; font-weight: 500; }
  .badge-blue   { background: #1e3a5f; color: #93c5fd; }
  .badge-green  { background: #052e16; color: #86efac; }
  .badge-orange { background: #431407; color: #fdba74; }

  /* Section tabs */
  .tabs { display: flex; gap: 4px; margin-bottom: 20px; border-bottom: 1px solid var(--border); }
  .tab { padding: 10px 18px; font-size: 13px; cursor: pointer; color: var(--muted);
         border-bottom: 2px solid transparent; transition: all .15s; }
  .tab:hover { color: var(--text); }
  .tab.active { color: var(--accent2); border-bottom-color: var(--accent); }

  /* Loading */
  .spinner { display: inline-block; width: 14px; height: 14px; border: 2px solid var(--border);
             border-top-color: var(--accent); border-radius: 50%; animation: spin .6s linear infinite; }
  @keyframes spin { to { transform: rotate(360deg); } }

  /* Toast */
  #toast { position: fixed; top: 16px; right: 16px; z-index: 999;
           background: var(--bg2); border: 1px solid var(--border); border-radius: var(--radius);
           padding: 12px 18px; font-size: 13px; transform: translateX(120%);
           transition: transform .2s; max-width: 320px; }
  #toast.show { transform: translateX(0); }
  #toast.toast-success { border-left: 3px solid var(--success); color: #86efac; }
  #toast.toast-error   { border-left: 3px solid var(--danger);  color: #fca5a5; }

  .grid-2 { display: grid; grid-template-columns: 1fr 1fr; gap: 16px; }
  .grid-3 { display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 16px; }
  @media(max-width:768px) { .grid-2, .grid-3 { grid-template-columns: 1fr; } }
  .mt-16 { margin-top: 16px; }
  .mb-16 { margin-bottom: 16px; }
  .gap-8 { gap: 8px; }
  .flex { display: flex; }
  .items-center { align-items: center; }
  .text-muted { color: var(--muted); }
  .text-mono { font-family: 'IBM Plex Mono', monospace; }
  .text-success { color: var(--success); }
  .text-danger { color: var(--danger); }
</style>
"""

_NAV = """
<nav class="nav">
  <span class="nav-brand">▸ SOTREG</span>
  {% if role == 'admin' %}
  <a href="/dashboard" {% if active=='dashboard' %}class="active"{% endif %}>Dashboard</a>
  {% endif %}
  {% if role in ['saisie','admin'] %}
  <a href="/saisie" {% if active=='saisie' %}class="active"{% endif %}>Saisie</a>
  {% endif %}
  {% if role in ['exports','admin'] %}
  <a href="/exports" {% if active=='exports' %}class="active"{% endif %}>Exports</a>
  {% endif %}
  <span style="margin-left:auto;font-size:12px;color:var(--muted)">{{ user }} · <span class="badge badge-{% if role=='admin' %}orange{% elif role=='saisie' %}blue{% else %}green{% endif %}">{{ role }}</span></span>
  <a href="/logout" class="logout">Déconnexion</a>
</nav>
<div id="toast"></div>
<script>
function showToast(msg, type='success'){
  const t=document.getElementById('toast');
  t.textContent=msg; t.className='show toast-'+type;
  setTimeout(()=>t.className='',3500);
}
</script>
"""

LOGIN_HTML = """<!doctype html><html><head>""" + _BASE_STYLE + """
<title>SOTREG · Connexion</title>
<style>
.login-wrap { min-height:100vh; display:flex; align-items:center; justify-content:center; }
.login-box { width:360px; }
.logo { font-family:'IBM Plex Mono',monospace; font-size:24px; font-weight:600;
        color:var(--accent2); margin-bottom:8px; letter-spacing:2px; }
.sub { color:var(--muted); font-size:13px; margin-bottom:32px; }
.login-form { background:var(--bg2); border:1px solid var(--border); border-radius:var(--radius); padding:28px; }
.login-form .form-group { margin-bottom:16px; }
.login-form input { width:100%; }
.login-form button { width:100%; margin-top:8px; justify-content:center; padding:10px; font-size:14px; }
</style></head><body>
<div class="login-wrap"><div class="login-box">
  <div class="logo">SOTREG</div>
  <div class="sub">Application de facturation</div>
  {% if error %}<div class="alert alert-error">{{ error }}</div>{% endif %}
  <div class="login-form">
    <form method="post">
      <div class="form-group"><label>Identifiant</label><input type="text" name="username" autofocus></div>
      <div class="form-group"><label>Mot de passe</label><input type="password" name="password"></div>
      <button type="submit" class="btn btn-primary">Se connecter</button>
    </form>
    <div style="margin-top:20px;font-size:11px;color:var(--muted)">
      Comptes test: user1/pass1 (saisie) · user2/pass2 (exports) · admin/admin
    </div>
  </div>
</div></div></body></html>"""

DASHBOARD_HTML = """<!doctype html><html><head>""" + _BASE_STYLE + """
<title>SOTREG · Dashboard</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js"></script>
<style>
.charts-grid { display:grid; grid-template-columns:1fr 1fr; gap:16px; margin-top:16px; }
.chart-card  { background:var(--bg2); border:1px solid var(--border); border-radius:var(--radius); padding:20px; }
.chart-card canvas { max-height:240px; }
@media(max-width:900px){ .charts-grid { grid-template-columns:1fr; } }
.kpi-row { display:grid; grid-template-columns:repeat(5,1fr); gap:12px; margin-bottom:8px; }
@media(max-width:900px){ .kpi-row { grid-template-columns:repeat(2,1fr); } }
.kpi-card { background:var(--bg2); border:1px solid var(--border); border-radius:var(--radius);
            padding:14px 16px; border-top:3px solid var(--accent); transition: opacity .2s; }
.kpi-val  { font-size:26px; font-weight:600; color:var(--accent2); font-family:'IBM Plex Mono',monospace; }
.kpi-lbl  { font-size:11px; color:var(--muted); text-transform:uppercase; letter-spacing:.7px; margin-top:2px; }
.kpi-sub  { font-size:11px; color:var(--muted); margin-top:4px; }
.month-pills { display:flex; gap:6px; align-items:center; flex-wrap:wrap; margin-bottom:16px; }
.mpill { padding:5px 14px; border-radius:100px; font-size:12px; font-weight:500; cursor:pointer;
         border:1px solid var(--border); background:var(--bg3); color:var(--muted);
         transition: all .15s; font-family:'IBM Plex Mono',monospace; }
.mpill:hover { border-color:var(--accent); color:var(--accent2); }
.mpill.active { background:var(--accent); border-color:var(--accent); color:#fff; }
.mpill-all { border-style:dashed; }
</style>
</head><body>
""" + _NAV.replace("{{ active }}", "") + """
<div class="container">
  <div class="page-title">Tableau de bord</div>
  {% with msgs = get_flashed_messages(with_categories=true) %}{% for cat,msg in msgs %}
    <div class="alert alert-{{ cat }}">{{ msg }}</div>{% endfor %}{% endwith %}

  <!-- Filtre par mois -->
  <div class="month-pills">
    <span style="font-size:11px;color:var(--muted);text-transform:uppercase;letter-spacing:.7px;margin-right:4px">Période :</span>
    <button class="mpill mpill-all active" onclick="filterMonth('all',this)">Tous les mois</button>
    {% for p in periods %}
    <button class="mpill" onclick="filterMonth('{{ p }}',this)">{{ p }}</button>
    {% endfor %}
  </div>

  <!-- KPI Cards (filtrées) -->
  <div class="kpi-row" id="kpi-row">
    <div class="kpi-card" style="border-top-color:#3b82f6">
      <div class="kpi-val" id="kpi-lines">{{ nb_lines }}</div>
      <div class="kpi-lbl">Lignes de saisie</div>
      <div class="kpi-sub" id="kpi-lines-sub">tous mois</div>
    </div>
    <div class="kpi-card" style="border-top-color:#22c55e">
      <div class="kpi-val" id="kpi-veh">{{ nb_veh }}</div>
      <div class="kpi-lbl">Véhicules distincts</div>
      <div class="kpi-sub" id="kpi-veh-sub">tous mois</div>
    </div>
    <div class="kpi-card" style="border-top-color:#f59e0b">
      <div class="kpi-val" id="kpi-km">—</div>
      <div class="kpi-lbl">KM total parcourus</div>
      <div class="kpi-sub" id="kpi-km-sub">tous mois</div>
    </div>
    <div class="kpi-card" style="border-top-color:#a78bfa">
      <div class="kpi-val" id="kpi-rot">—</div>
      <div class="kpi-lbl">Rotations totales</div>
      <div class="kpi-sub" id="kpi-rot-sub">tous mois</div>
    </div>
    <div class="kpi-card" style="border-top-color:#f472b6">
      <div class="kpi-val" id="kpi-total-fact">—</div>
      <div class="kpi-lbl">Facturation HT estimée</div>
      <div class="kpi-sub" id="kpi-fact-sub">tous mois</div>
    </div>
  </div>

  <!-- Charts -->
  <div class="charts-grid">
    <div class="chart-card">
      <div class="card-title">KM parcourus — prestataire</div>
      <canvas id="chart-km"></canvas>
    </div>
    <div class="chart-card">
      <div class="card-title">Rotations — prestataire</div>
      <canvas id="chart-rot"></canvas>
    </div>
    <div class="chart-card">
      <div class="card-title">Facturation estimée (MAD)</div>
      <canvas id="chart-fact"></canvas>
    </div>
    <div class="chart-card">
      <div class="card-title">Véhicules actifs</div>
      <canvas id="chart-veh"></canvas>
    </div>
    <div class="chart-card" style="grid-column:1/-1">
      <div class="card-title" style="display:flex;align-items:center;gap:12px">
        Facture HT par entité
        <span style="font-size:11px;color:var(--muted);font-weight:400">(KM × tarif)</span>
      </div>
      <canvas id="chart-entity" style="max-height:300px"></canvas>
    </div>
  </div>

  <!-- Admin tools -->
  <div class="grid-2 mt-16">
    <div class="card">
      <div class="card-title">Source de données</div>
      <div style="display:flex;align-items:center;gap:8px;margin-bottom:12px;flex-wrap:wrap">
        <span class="badge badge-{% if db_name == 'sotreg_work.db' %}orange{% else %}green{% endif %} text-mono">
          {{ db_name }}
        </span>
        {% if db_name != 'sotreg_work.db' %}
        <a href="/admin/reset_db_view" class="btn btn-secondary btn-sm">↺ DB locale</a>
        {% endif %}
      </div>
      <form method="post" action="/admin/upload_db" enctype="multipart/form-data"
            style="display:flex;gap:8px;align-items:center;flex-wrap:wrap">
        <input type="file" name="dbfile" accept=".db"
               style="font-size:12px;color:var(--muted);flex:1;min-width:180px">
        <button type="submit" class="btn btn-primary btn-sm">Charger DB consolidée</button>
      </form>
    </div>
    <div class="card">
      <div class="card-title">Administration</div>
      <div style="display:flex;gap:8px;flex-wrap:wrap;margin-bottom:12px">
        <a href="/exports" class="btn btn-success btn-sm">→ Exports</a>
        <a href="/saisie"  class="btn btn-primary btn-sm">→ Saisie</a>
        <form method="post" action="/admin/reset_db" style="display:inline"
              onsubmit="return confirm('Réinitialiser la DB de saisie depuis le template ?')">
          <button class="btn btn-danger btn-sm">↺ Réinitialiser DB</button>
        </form>
      </div>
      {% for p in periods %}<span class="badge badge-blue" style="margin:2px">{{ p }}</span>{% endfor %}
      <br style="margin:4px 0">
      {% for p in providers %}<span class="badge badge-green" style="margin:2px">{{ p }}</span>{% endfor %}
    </div>
  </div>
</div>

<script>
const COLORS    = ['#3b82f6','#22c55e','#f59e0b','#a78bfa','#f472b6','#34d399'];
const ENT_COLORS= ['#3b82f6','#22c55e','#f59e0b','#a78bfa','#f472b6','#34d399','#fb923c','#e879f9'];
const GRID      = { color:'rgba(255,255,255,0.06)' };
const TICK      = { color:'#7c8fa6', font:{size:11} };
const FMT = v => v>=1e6?(v/1e6).toFixed(2)+' M':v>=1e3?(v/1e3).toFixed(0)+'k':String(Math.round(v));
Chart.defaults.color = '#7c8fa6';
Chart.defaults.font.family = "'IBM Plex Sans', sans-serif";

let _kpiData = null;
let _charts  = {};
let _selMonth = 'all';

function makeChart(id, type, labels, datasets, opts={}) {
  if(_charts[id]) _charts[id].destroy();
  const ctx = document.getElementById(id);
  if(!ctx) return;
  _charts[id] = new Chart(ctx, {
    type, data:{labels, datasets},
    options:{
      responsive:true, maintainAspectRatio:true,
      plugins:{ legend:{ labels:{ color:'#a0aec0', font:{size:11}, boxWidth:12 } } },
      scales: type!=='pie' ? {
        x:{ ticks:TICK, grid:GRID },
        y:{ ticks:{...TICK, callback:FMT}, grid:GRID }
      } : undefined,
      ...opts
    }
  });
}

function filterMonth(month, btn){
  _selMonth = month;
  document.querySelectorAll('.mpill').forEach(b=>b.classList.remove('active'));
  btn.classList.add('active');
  if(_kpiData) renderAll(_kpiData, month);
}

function renderKpiCards(d, month){
  const filt = r => month==='all' || r.period===month;
  const sub  = month==='all' ? 'tous mois' : month;

  const km  = d.km_rotations.filter(filt).reduce((s,r)=>s+(r.total_km||0),0);
  const rot = d.km_rotations.filter(filt).reduce((s,r)=>s+(r.total_rot||0),0);
  const fact= d.facturation.filter(filt).reduce((s,r)=>s+(r.frais_km||0)+(r.frais_dispo||0),0);
  const veh = month==='all'
    ? [...new Set(d.vehicles.map(r=>r.provider+'|'+r.period))].length   // proxy
    : d.vehicles.filter(r=>r.period===month).reduce((s,r)=>s+(r.nb_veh||0),0);
  const lines = d.km_rotations.filter(filt).reduce((s,r)=>s+(r.nb_lines||0),0);

  document.getElementById('kpi-lines').textContent = lines || '—';
  document.getElementById('kpi-veh').textContent   = veh   || '—';
  document.getElementById('kpi-km').textContent    = FMT(km);
  document.getElementById('kpi-rot').textContent   = Math.round(rot).toLocaleString('fr');
  document.getElementById('kpi-total-fact').textContent = FMT(fact);
  ['kpi-lines-sub','kpi-veh-sub','kpi-km-sub','kpi-rot-sub','kpi-fact-sub']
    .forEach(id=>{ const el=document.getElementById(id); if(el) el.textContent=sub; });
}

function renderAll(d, month){
  renderKpiCards(d, month);
  const filt = r => month==='all' || r.period===month;

  const periods   = [...new Set(d.km_rotations.filter(filt).map(r=>r.period))].sort();
  const providers = [...new Set(d.km_rotations.map(r=>r.provider))].sort();

  // KM chart
  makeChart('chart-km','bar',periods, providers.map((p,i)=>({
    label:p,
    data: periods.map(per=>{ const r=d.km_rotations.find(x=>x.period===per&&x.provider===p); return r?Math.round(r.total_km):0; }),
    backgroundColor:COLORS[i%COLORS.length]+'cc', borderColor:COLORS[i%COLORS.length], borderWidth:1
  })));

  // Rotations chart
  makeChart('chart-rot','line',periods, providers.map((p,i)=>({
    label:p,
    data: periods.map(per=>{ const r=d.km_rotations.find(x=>x.period===per&&x.provider===p); return r?Math.round(r.total_rot):0; }),
    backgroundColor:'transparent', borderColor:COLORS[i%COLORS.length], borderWidth:2, pointRadius:4, tension:0.3
  })));

  // Facturation chart
  const fPers  = [...new Set(d.facturation.filter(filt).map(r=>r.period))].sort();
  const fProvs = [...new Set(d.facturation.map(r=>r.provider))].sort();
  makeChart('chart-fact','bar',fPers, fProvs.map((p,i)=>({
    label:p,
    data: fPers.map(per=>{ const rows=d.facturation.filter(x=>x.period===per&&x.provider===p); return rows.reduce((s,r)=>s+(r.frais_km||0)+(r.frais_dispo||0),0); }),
    backgroundColor:COLORS[i%COLORS.length]+'99', borderColor:COLORS[i%COLORS.length], borderWidth:1
  })),{scales:{x:{stacked:true,ticks:TICK,grid:GRID},y:{stacked:true,ticks:{...TICK,callback:FMT},grid:GRID}}});

  // Vehicles chart
  const vPers  = [...new Set(d.vehicles.filter(filt).map(r=>r.period))].sort();
  const vProvs = [...new Set(d.vehicles.map(r=>r.provider))].sort();
  makeChart('chart-veh','bar',vPers, vProvs.map((p,i)=>({
    label:p,
    data: vPers.map(per=>{ const r=d.vehicles.find(x=>x.period===per&&x.provider===p); return r?r.nb_veh:0; }),
    backgroundColor:COLORS[i%COLORS.length]+'cc', borderColor:COLORS[i%COLORS.length], borderWidth:1
  })));

  // Entity billing chart
  const ePers = [...new Set(d.entity_fact.filter(filt).map(r=>r.period))].sort();
  const ents  = [...new Set(d.entity_fact.map(r=>r.entity))].sort();
  makeChart('chart-entity','bar',ePers, ents.map((ent,i)=>({
    label:ent,
    data: ePers.map(per=>{ const r=d.entity_fact.find(x=>x.period===per&&x.entity===ent); return r?Math.round(r.fact_ht):0; }),
    backgroundColor:ENT_COLORS[i%ENT_COLORS.length]+'bb', borderColor:ENT_COLORS[i%ENT_COLORS.length], borderWidth:1, borderRadius:3
  })),{
    plugins:{ legend:{ position:'right', labels:{ color:'#a0aec0', font:{size:11}, boxWidth:12, padding:10 } } },
    scales:{ x:{ticks:TICK,grid:GRID}, y:{ticks:{...TICK,callback:FMT},grid:GRID} }
  });
}

async function loadKPI(){
  try {
    const res = await fetch('/api/kpi');
    _kpiData  = await res.json();
    renderAll(_kpiData, _selMonth);
  } catch(e){ console.error('KPI error:', e); }
}
loadKPI();
</script>
</body></html>"""

SAISIE_HTML = """<!doctype html><html><head>""" + _BASE_STYLE + """
<title>SOTREG · Saisie</title>
<style>
.toolbar { display:flex; gap:8px; align-items:center; flex-wrap:wrap; margin-bottom:16px; }
.section-label { font-size:11px;text-transform:uppercase;letter-spacing:.8px;color:var(--muted);
                  font-weight:600;padding:8px 0 4px; display:flex; align-items:center; gap:10px; }
#status-bar { font-size:12px; color:var(--muted); min-height:20px; }
.add-row-btn { font-size:12px; padding:4px 10px; }
.prev-hint { font-size:11px; color:var(--warning); background:#431407; border:1px solid #92400e;
             border-radius:4px; padding:4px 10px; display:none; }
</style>
</head><body>
""" + _NAV.replace("{{ active }}", "") + """
<div class="container">
  <div class="page-title">Saisie des données</div>
  {% with msgs = get_flashed_messages(with_categories=true) %}{% for cat,msg in msgs %}
    <div class="alert alert-{{ cat }}">{{ msg }}</div>{% endfor %}{% endwith %}

  <!-- Sélecteurs -->
  <div class="card mb-16">
    <div class="form-row">
      <div class="form-group">
        <label>Prestataire</label>
        <select id="sel-provider" onchange="onProviderChange()">
          {% for p in providers %}<option>{{ p }}</option>{% endfor %}
        </select>
      </div>
      <div class="form-group">
        <label>Type de véhicule</label>
        <select id="sel-vtype" onchange="onSelChange()">
          {% for v in vtypes %}<option>{{ v }}</option>{% endfor %}
        </select>
      </div>
      <div class="form-group">
        <label>Période (YYYY-MM)</label>
        <div style="display:flex;gap:6px;align-items:center">
          <select id="sel-period-pick" onchange="onPeriodPick(this.value)" style="width:140px">
            {% for p in periods_db %}<option>{{ p }}</option>{% endfor %}
            <option value="__custom__">Autre…</option>
          </select>
          <input type="text" id="sel-period" style="width:110px;display:none" placeholder="YYYY-MM" oninput="onSelChange()">
        </div>
      </div>
      <div class="form-group">
        <label>&nbsp;</label>
        <button class="btn btn-primary" onclick="loadData()">
          <span id="load-spinner" class="spinner" style="display:none"></span> Charger
        </button>
      </div>
    </div>
    <div id="prev-hint" class="prev-hint">
      ⚡ Nouveau mois — données pré-remplies depuis le mois précédent. Modifiez puis sauvegardez.
    </div>
    <div id="status-bar" style="margin-top:8px">Sélectionnez une combinaison et cliquez Charger.</div>
  </div>

  <div id="data-section" style="display:none">

    <!-- Tarifs -->
    <div class="section-label">Tarifs</div>
    <div class="card mb-16">
      <div id="tarifs-container"></div>
    </div>

    <!-- Lignes -->
    <div class="section-label">
      Lignes de facturation
      <button class="btn btn-secondary add-row-btn" onclick="addLine()">+ Ligne</button>
    </div>
    <!-- Barre de filtre lignes -->
    <div style="display:flex;gap:8px;align-items:center;margin-bottom:8px;flex-wrap:wrap">
      <input type="text" id="filter-lines" placeholder="🔍  Filtrer par circuit, entité, chauffeur, MLE…"
             oninput="applyFilter('lines-body',this.value)"
             style="flex:1;min-width:220px;background:var(--bg3);border:1px solid var(--border);
                    color:var(--text);padding:7px 12px;border-radius:var(--radius);font-size:13px;
                    font-family:inherit;outline:none">
      <button class="btn btn-secondary btn-sm" onclick="clearFilter('filter-lines','lines-body')">✕ Effacer</button>
      <span id="filter-lines-count" style="font-size:12px;color:var(--muted)"></span>
    </div>
    <div class="card mb-16 table-wrap">
      <table id="lines-table"><thead id="lines-head"></thead><tbody id="lines-body"></tbody></table>
    </div>

    <!-- Véhicules (non-SOTREG) -->
    <div id="veh-section">
      <div class="section-label">
        Véhicules
        <button class="btn btn-secondary add-row-btn" onclick="addVeh()">+ Véhicule</button>
      </div>
      <!-- Barre de filtre véhicules -->
      <div style="display:flex;gap:8px;align-items:center;margin-bottom:8px;flex-wrap:wrap">
        <input type="text" id="filter-veh" placeholder="🔍  Filtrer par MLE, age, dispo…"
               oninput="applyFilter('veh-body',this.value)"
               style="flex:1;min-width:220px;background:var(--bg3);border:1px solid var(--border);
                      color:var(--text);padding:7px 12px;border-radius:var(--radius);font-size:13px;
                      font-family:inherit;outline:none">
        <button class="btn btn-secondary btn-sm" onclick="clearFilter('filter-veh','veh-body')">✕ Effacer</button>
        <span id="filter-veh-count" style="font-size:12px;color:var(--muted)"></span>
      </div>
      <div class="card mb-16 table-wrap">
        <table id="veh-table"><thead id="veh-head"></thead><tbody id="veh-body"></tbody></table>
      </div>
    </div>

    <!-- Actions -->
    <div class="toolbar">
      <button class="btn btn-success" onclick="saveData()">
        <span id="save-spinner" class="spinner" style="display:none"></span> ✓ Sauvegarder
      </button>
      <a href="#" class="btn btn-primary" onclick="return calcExcel()">⬇ Calculer Excel</a>
      <span id="save-status" style="font-size:12px;color:var(--muted)"></span>
    </div>
  </div>

  <!-- Bandeau export DB — toujours visible -->
  <div style="margin-top:32px;border:1px solid #059669;border-radius:8px;
              background:#052e2b;padding:16px 20px">
    <div style="font-size:13px;font-weight:600;color:#6ee7b7;margin-bottom:5px">
      📂 Reprendre une saisie existante
    </div>
    <div style="font-size:12px;color:#a7f3d0;margin-bottom:12px">
      Au début de la session, chargez la dernière base téléchargée. Son contenu remplacera
      la base temporaire actuellement présente sur Render.
    </div>
    <form method="post" action="/saisie/upload_db" enctype="multipart/form-data"
          style="display:flex;gap:10px;align-items:center;flex-wrap:wrap">
      <input type="file" name="dbfile" accept=".db" required
             style="flex:1;min-width:240px;background:var(--bg3);padding:8px;border-radius:6px">
      <button type="submit" class="btn btn-success">📂 Charger cette base</button>
    </form>
  </div>

  <div style="margin-top:32px;border:1px solid #1d4ed8;border-radius:8px;
              background:#172554;padding:16px 20px;display:flex;align-items:center;
              gap:16px;flex-wrap:wrap">
    <div style="flex:1;min-width:220px">
      <div style="font-size:13px;font-weight:600;color:#93c5fd;margin-bottom:4px">
        ⬇ Transmettre la base de données
      </div>
      <div style="font-size:12px;color:#60a5fa">
        Une fois votre saisie terminée, téléchargez la DB et envoyez-la à l'utilisateur 2 (exports) et à l'administrateur.
      </div>
    </div>
    <a href="/saisie/download_db" class="btn btn-primary" style="white-space:nowrap;font-size:13px">
      ⬇ Télécharger sotreg_saisie.db
    </a>
  </div>
</div>

<script>
// Periods known from DB (used to detect new months)
const KNOWN_PERIODS = {{ periods_db | tojson }};
let state = {sotreg: false, lines: [], vehicles: [], tariffs: [], _is_new: false};

function prevMonth(ym){
  const [y,m] = ym.split('-').map(Number);
  const d = new Date(y, m-2, 1);
  return d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0');
}

function onPeriodPick(val){
  const inp = document.getElementById('sel-period');
  const pick = document.getElementById('sel-period-pick');
  if(val === '__custom__'){
    inp.style.display = '';
    inp.focus();
  } else {
    inp.style.display = 'none';
    inp.value = val;
    onSelChange();
  }
}

function getPeriod(){
  const pick = document.getElementById('sel-period-pick').value;
  if(pick === '__custom__') return document.getElementById('sel-period').value.trim();
  return pick;
}

function onSelChange(){
  // Auto-suggest previous month when period is not in known list
  const period = getPeriod();
  document.getElementById('status-bar').textContent =
    period ? `Prêt à charger: ${period}` : 'Sélectionnez une combinaison et cliquez Charger.';
}

async function loadData(){
  const provider = document.getElementById('sel-provider').value;
  const vtype    = document.getElementById('sel-vtype').value;
  const period   = getPeriod();
  const reYM = /^[0-9]{4}-[0-9]{2}$/;
  if(!period || !reYM.test(period)){
    showToast('Période invalide (format YYYY-MM requis)','error'); return;
  }
  const sp = document.getElementById('load-spinner');
  sp.style.display='inline-block';
  try {
    const r = await fetch(`/api/saisie/load?provider=${encodeURIComponent(provider)}&vtype=${encodeURIComponent(vtype)}&period=${encodeURIComponent(period)}`);
    const d = await r.json();
    if(d.error){ showToast(d.error,'error'); return; }

    // Detect if this is a new month (no data existed → came from previous month)
    const isNew = !KNOWN_PERIODS.includes(period);
    d._is_new = isNew;
    state = d;

    renderAll(d, provider, vtype);
    document.getElementById('data-section').style.display='';
    document.getElementById('status-bar').textContent =
      `Chargé: ${period} / ${provider} / ${vtype} — ${d.lines.length} ligne(s)`;
    resetFilters();

    const hint = document.getElementById('prev-hint');
    if(isNew && d.lines.length > 0){
      hint.style.display = 'block';
      hint.textContent = `⚡ Nouveau mois — ${d.lines.length} ligne(s) pré-remplie(s) depuis ${prevMonth(period)}. Modifiez puis sauvegardez.`;
    } else {
      hint.style.display = 'none';
    }
  } catch(e){ showToast('Erreur réseau: '+e,'error'); }
  finally { sp.style.display='none'; }
}

function renderAll(d, provider, vtype){
  renderTarifs(d);
  if(d.sotreg){
    renderLinesSotreg(d.lines);
    document.getElementById('veh-section').style.display='none';
  } else {
    renderLines(d.lines);
    renderVehicles(d.vehicles, provider, vtype);
    document.getElementById('veh-section').style.display='';
  }
}

function renderTarifs(d){
  const c = document.getElementById('tarifs-container');
  const tariffs = Array.isArray(d.tariffs) ? d.tariffs : (d.tariff ? [d.tariff] : []);
  state.tariffs = tariffs;
  if(!tariffs.length){
    c.innerHTML='<span class="text-muted" style="font-size:13px">Aucun tarif défini — cliquez + Tarif pour en ajouter.</span><br>';
  } else {
    const keys = Object.keys(tariffs[0]).filter(k=>k!=='id');
    let html = '<div class="table-wrap"><table><thead><tr>' + keys.map(k=>`<th>${k}</th>`).join('') + '<th></th></tr></thead><tbody>';
    tariffs.forEach((t,i)=>{
      html += '<tr>' + keys.map(k=>`<td><input data-tarif="${i}" data-key="${k}" value="${t[k]??''}" oninput="updateTarif(this)"></td>`).join('') +
              `<td><button class="btn btn-danger btn-sm btn-icon" onclick="removeTarif(${i})">✕</button></td></tr>`;
    });
    html += '</tbody></table></div>';
    c.innerHTML = html;
  }
  c.innerHTML += `<button class="btn btn-secondary btn-sm mt-16" onclick="addTarif()">+ Tarif</button>`;
}

function updateTarif(el){ state.tariffs[+el.dataset.tarif][el.dataset.key]=el.value; }
function addTarif(){
  const keys = state.tariffs.length ? Object.keys(state.tariffs[0]).filter(k=>k!=='id')
    : ['billing_mode','age_cat','price_mise_dispo','price_km','price_km_supp','price_journalier','price_day','km_forfait_value'];
  const blank = {}; keys.forEach(k=>blank[k]='0'); blank.billing_mode=''; blank.age_cat='';
  state.tariffs.push(blank);
  renderTarifs({...state, tariffs: state.tariffs});
}
function removeTarif(i){ state.tariffs.splice(i,1); renderTarifs({...state, tariffs: state.tariffs}); }

// ── Mapping entité → {circuit: km} ───────────────────────────────────────
let _entityCircuits = {};

async function loadEntityCircuits(){
  try {
    const r = await fetch('/api/entity_circuits');
    _entityCircuits = await r.json();
    // Reconstruire le datalist dl-circuit global (tous circuits)
  } catch(e){ console.warn('entity_circuits load failed:', e); }
}

// Quand l'entité change dans une ligne → filtrer le datalist circuit de cette ligne
function onEntityChange(input){
  updateLine(input);
  const tr      = input.closest('tr');
  const entity  = input.value.trim();
  const dlLocal = tr.querySelector('datalist[id^="dl-circ-row"]');
  if(!dlLocal) return;
  const circuits = entity && _entityCircuits[entity]
    ? Object.keys(_entityCircuits[entity])
    : Object.keys(_circuitKm);
  dlLocal.innerHTML = circuits.map(c => `<option value="${c.replace(/"/g,'&quot;')}">`).join('');
}

// Auto-fill km depuis entité+circuit
function onCircuitBlur(input){
  const i   = +input.dataset.row;
  const cir = input.value.trim();
  if(!cir) return;

  const tr       = input.closest('tr');
  const entInput = tr.querySelector('input[data-key="entity"]');
  let entity   = entInput ? entInput.value.trim() : '';
  const kmInput  = tr.querySelector('input[data-key="km_per_rotation"]');
  if(!kmInput) return;

  // Circuit choisi en premier : retrouver automatiquement son entité.
  const matches = [];
  const wanted = cir.toUpperCase();
  for(const [ent, circuits] of Object.entries(_entityCircuits)){
    for(const known of Object.keys(circuits)){
      if(known.trim().toUpperCase() === wanted){ matches.push(ent); break; }
    }
  }
  if(entInput && (!entity || !matches.includes(entity)) && matches.length){
    // S'il existe plusieurs entités, utiliser la correspondance saisie le plus récemment.
    entity = matches[0];
    entInput.value = entity;
    if(state.lines[i]) state.lines[i].entity = entity;
    onEntityChange(entInput);
    entInput.style.background = 'rgba(16,185,129,0.2)';
    entInput.style.border = '1px solid #10b981';
    setTimeout(()=>{ entInput.style.background='transparent'; entInput.style.border='none'; },1200);
  }

  let km = '';
  // Priorité 1 : règle NAVETTE
  if(cir.toUpperCase().includes('NAVETTE')) {
    km = 18;
  }
  // Priorité 2 : mapping entité → circuit
  else if(entity && _entityCircuits[entity] && _entityCircuits[entity][cir] !== undefined){
    km = _entityCircuits[entity][cir];
  }
  // Priorité 3 : mapping global 2026-03
  else {
    const cu = cir.toUpperCase();
    for(const [k,v] of Object.entries(_circuitKm)){
      if(k.trim().toUpperCase() === cu){ km = v; break; }
    }
  }

  // Remplir seulement si champ vide
  if(km !== '' && (!kmInput.value || +kmInput.value === 0)){
    kmInput.value = km;
    if(state.lines[i]) state.lines[i].km_per_rotation = km;
    // Highlight bref
    kmInput.style.background = 'rgba(59,130,246,0.2)';
    kmInput.style.border = '1px solid var(--accent)';
    setTimeout(()=>{ kmInput.style.background='transparent'; kmInput.style.border='none'; }, 1200);
  }
}
let _circuitKm = {};

async function loadCircuitKm(){
  try {
    const r = await fetch('/api/circuit_km');
    _circuitKm = await r.json();
  } catch(e){ console.warn('circuit_km load failed:', e); }
}

let _providerVtypes = {};

async function loadProviderVtypes(){
  try {
    const r = await fetch('/api/provider_vtypes');
    _providerVtypes = await r.json();
    // Appliquer le filtre initial au chargement
    filterVtypesByProvider(document.getElementById('sel-provider').value);
  } catch(e){ console.warn('provider_vtypes load failed:', e); }
}

function filterVtypesByProvider(provider){
  const sel   = document.getElementById('sel-vtype');
  const avail = _providerVtypes[provider] || [];
  const prev  = sel.value;

  // Mettre à jour les options
  [...sel.options].forEach(opt => {
    const match = avail.length === 0 || avail.includes(opt.value);
    opt.hidden   = !match;
    opt.disabled = !match;
    if(!match && opt.selected) opt.selected = false;
  });

  // Sélectionner le premier disponible si l'ancien n'est plus valide
  const stillValid = [...sel.options].find(o => o.value === prev && !o.hidden);
  if(!stillValid){
    const first = [...sel.options].find(o => !o.hidden);
    if(first) sel.value = first.value;
  }

  onSelChange();
}

function onProviderChange(){
  filterVtypesByProvider(document.getElementById('sel-provider').value);
}
let _lists = { entities:[], circuits:[], mles:[], chauffeurs:[] };

async function loadLists(){
  try {
    const r = await fetch('/api/lists');
    _lists  = await r.json();
    // Injecter les datalists dans le DOM
    const dl = (id, vals) => `<datalist id="${id}">${vals.map(v=>`<option value="${v.replace(/"/g,'&quot;')}">`).join('')}</datalist>`;
    document.body.insertAdjacentHTML('beforeend',
      dl('dl-entity',   _lists.entities)  +
      dl('dl-circuit',  _lists.circuits)  +
      dl('dl-mle',      _lists.mles)      +
      dl('dl-chauffeur',_lists.chauffeurs)
    );
  } catch(e){ console.warn('Lists load failed:', e); }
}

// Cellule circuit : datalist local à la ligne (filtré par entité) + onblur km auto-fill
function inputCellCircuit(i, k, val, dlId=''){
  const style = 'background:transparent;border:none;color:var(--text);width:100%;font-family:inherit;font-size:13px;padding:2px 4px;outline:none';
  const localDlId = `dl-circ-row-${i}`;
  // Circuits initiaux : tous (seront filtrés si entité sélectionnée)
  const allCircuits = Object.keys(_circuitKm);
  const dlHtml = `<datalist id="${localDlId}">${allCircuits.map(c=>`<option value="${c.replace(/"/g,'&quot;')}">`).join('')}</datalist>`;
  return `<td>${dlHtml}<input data-row="${i}" data-key="${k}" value="${(val??'').toString().replace(/"/g,'&quot;')}"
    list="${localDlId}" oninput="updateLine(this)" onchange="onCircuitBlur(this)"
    style="${style}"
    onfocus="this.style.background='var(--bg)';this.style.border='1px solid var(--accent)';this.style.borderRadius='4px'"
    onblur="this.style.background='transparent';this.style.border='none';onCircuitBlur(this)"></td>`;
}

// Cellule entité avec filtre circuit
function inputCellEntity(i, val){
  const style = 'background:transparent;border:none;color:var(--text);width:100%;font-family:inherit;font-size:13px;padding:2px 4px;outline:none';
  return `<td><input data-row="${i}" data-key="entity" value="${(val??'').toString().replace(/"/g,'&quot;')}"
    list="dl-entity" oninput="updateLine(this)" onchange="onEntityChange(this)"
    style="${style}"
    onfocus="this.style.background='var(--bg)';this.style.border='1px solid var(--accent)';this.style.borderRadius='4px'"
    onblur="this.style.background='transparent';this.style.border='none'"></td>`;
}

// Cellule input générique avec datalist optionnel
function inputCell(i, k, val, dlId='', extra=''){
  const list  = dlId ? `list="${dlId}"` : '';
  const style = 'background:transparent;border:none;color:var(--text);width:100%;font-family:inherit;font-size:13px;padding:2px 4px;outline:none';
  return `<td><input data-row="${i}" data-key="${k}" value="${(val??'').toString().replace(/"/g,'&quot;')}"
    ${list} ${extra} oninput="updateLine(this)"
    style="${style}" onfocus="this.style.background='var(--bg)';this.style.border='1px solid var(--accent)';this.style.borderRadius='4px'"
    onblur="this.style.background='transparent';this.style.border='none'"></td>`;
}

// Cellule select (dispo, age)
function selectCell(i, k, val, options, onChange='', isVeh=false){
  const handler = isVeh ? `onchange="updateVeh(this)${onChange?';'+onChange+'(this)':''}"` : `onchange="updateLine(this)"`;
  const selStyle = 'background:var(--bg3);border:1px solid var(--border);color:var(--text);padding:4px 8px;border-radius:4px;font-size:12px;font-family:inherit;width:100%';
  const opts = options.map(([v,l]) => `<option value="${v}" ${val===v?'selected':''}>${l}</option>`).join('');
  const attr = isVeh ? `data-veh="${i}" data-key="${k}"` : `data-row="${i}" data-key="${k}"`;
  return `<td><select ${attr} ${handler} style="${selStyle}">${opts}</select></td>`;
}

function renderLinesSotreg(lines){
  document.getElementById('lines-head').innerHTML = '<tr><th>Entité</th><th>Circuit</th><th>Nb véhicules</th><th>KM/Rotation</th><th>Rotation total</th><th></th></tr>';
  document.getElementById('lines-body').innerHTML = lines.map((r,i) =>
    `<tr>
      ${inputCellEntity(i, r.entity)}
      ${inputCellCircuit(i,'circuit', r.circuit, 'dl-circuit')}
      ${inputCell(i,'nb_vehicles', r.nb_vehicles)}
      ${inputCell(i,'km_per_rotation', r.km_per_rotation)}
      ${inputCell(i,'rotation_total',  r.rotation_total)}
      <td><button class="btn btn-danger btn-sm btn-icon" onclick="removeLine(this)">✕</button></td>
    </tr>`
  ).join('');
  const el=document.getElementById('filter-lines-count'); if(el) el.textContent=`${lines.length} ligne(s)`;
}

function renderLines(lines){
  document.getElementById('lines-head').innerHTML = '<tr><th>Chauffeur</th><th>MLE CAR</th><th>Entité</th><th>Circuit</th><th>KM/Rotation</th><th>Rotation total</th><th></th></tr>';
  document.getElementById('lines-body').innerHTML = lines.map((r,i) =>
    `<tr>
      ${inputCell(i,'chauffeur',      r.chauffeur,       'dl-chauffeur')}
      ${inputCell(i,'mle_car',        r.mle_car,         'dl-mle')}
      ${inputCellEntity(i, r.entity)}
      ${inputCellCircuit(i,'circuit', r.circuit,         'dl-circuit')}
      ${inputCell(i,'km_per_rotation',r.km_per_rotation)}
      ${inputCell(i,'rotation_total', r.rotation_total)}
      <td><button class="btn btn-danger btn-sm btn-icon" onclick="removeLine(this)">✕</button></td>
    </tr>`
  ).join('');
  const el=document.getElementById('filter-lines-count'); if(el) el.textContent=`${lines.length} ligne(s)`;
}

function renderVehicles(vehs, provider, vtype){
  const p  = (provider||'').toUpperCase();
  const vt = (vtype||'').toLowerCase();
  const isAutocar = vt.includes('autocar') && ['STCR','S.TOURISME','S TOURISME'].includes(p);
  const isMinicar = vt.includes('minicar') || vt.includes('minibus');
  const SEL_STYLE = 'background:var(--bg3);border:1px solid var(--border);color:var(--text);padding:4px 8px;border-radius:4px;font-size:12px;font-family:inherit';

  const head = document.getElementById('veh-head');
  const body = document.getElementById('veh-body');

  // Header
  let hdrs = ['MLE CAR','KM compteur'];
  if(isAutocar)         hdrs = ['MLE CAR','KM compteur','Age'];
  else if(isMinicar)    hdrs = ['MLE CAR','KM compteur','KM forfait','KM supp'];
  if(p === 'MANAVETTE' && vt.includes('minibus')) hdrs.push('Zone');
  head.innerHTML = '<tr>' + hdrs.map(l=>`<th>${l}</th>`).join('') +
                   '<th>Disponibilité</th><th>Nb jours</th><th></th></tr>';

  body.innerHTML = vehs.map((v,i) => {
    const dispo  = (v.dispo_type||'permanent').toLowerCase();
    const isTemp = dispo === 'temporaire';
    const age    = v.age_cat || '';

    // MLE + KM compteur (commun)
    const mleCell = `<td><input data-veh="${i}" data-key="mle_car" value="${(v.mle_car||'').replace(/"/g,'&quot;')}"
      list="dl-mle" oninput="updateVeh(this)"
      style="background:transparent;border:none;color:var(--text);width:100%;font-family:inherit;font-size:13px;padding:2px 4px;outline:none"
      onfocus="this.style.background='var(--bg)';this.style.border='1px solid var(--accent)';this.style.borderRadius='4px'"
      onblur="this.style.background='transparent';this.style.border='none'"></td>`;
    const kmcCell = `<td><input data-veh="${i}" data-key="km_compteur" value="${v.km_compteur??''}"
      type="number" oninput="updateVeh(this)"
      style="background:transparent;border:none;color:var(--text);width:90px;font-family:'IBM Plex Mono',monospace;font-size:12px;padding:2px 4px;outline:none"
      onfocus="this.style.background='var(--bg)';this.style.border='1px solid var(--accent)';this.style.borderRadius='4px'"
      onblur="this.style.background='transparent';this.style.border='none'"></td>`;

    // Age select (autocar only)
    const ageCell = `<td><select data-veh="${i}" data-key="age_cat" onchange="updateVeh(this)" style="${SEL_STYLE}">
      <option value=""     ${age===''    ?'selected':''}>—</option>
      <option value="<5ans" ${age==='<5ans'?'selected':''}>&lt;5ans</option>
      <option value=">5ans" ${age==='>5ans'?'selected':''}>&gt;5ans</option>
    </select></td>`;

    // Extra km cells (minicar)
    const kmfCell = `<td><input data-veh="${i}" data-key="km_forfait" value="${v.km_forfait??''}" type="text" inputmode="decimal" oninput="updateVeh(this)"
      style="background:transparent;border:none;color:var(--text);width:90px;font-family:'IBM Plex Mono',monospace;font-size:12px;padding:2px 4px;outline:none"
      onfocus="this.style.background='var(--bg)';this.style.border='1px solid var(--accent)';this.style.borderRadius='4px'"
      onblur="this.style.background='transparent';this.style.border='none'"></td>`;
    const kmsCell = `<td><input data-veh="${i}" data-key="km_supp" value="${v.km_supp??''}" type="text" inputmode="decimal" oninput="updateVeh(this)"
      style="background:transparent;border:none;color:var(--text);width:90px;font-family:'IBM Plex Mono',monospace;font-size:12px;padding:2px 4px;outline:none"
      onfocus="this.style.background='var(--bg)';this.style.border='1px solid var(--accent)';this.style.borderRadius='4px'"
      onblur="this.style.background='transparent';this.style.border='none'"></td>`;

    // Dispo select
    const dispoCell = `<td><select data-veh="${i}" data-key="dispo_type" onchange="onDispoChange(this)" style="${SEL_STYLE}">
      <option value="permanent"  ${!isTemp?'selected':''}>Permanent</option>
      <option value="temporaire" ${isTemp?'selected':''}>Temporaire</option>
    </select></td>`;

    // Nb jours (conditionnel)
    const daysCell = `<td><input data-veh="${i}" data-key="nb_days" type="text" inputmode="decimal" min="0"
      value="${isTemp?(v.nb_days||0):''}" ${isTemp?'':'disabled placeholder="—"'}
      style="width:70px;background:var(--bg3);border:1px solid ${isTemp?'var(--warning)':'var(--border)'};
             color:${isTemp?'var(--warning)':'var(--muted)'};padding:4px 8px;border-radius:4px;
             font-size:12px;font-family:'IBM Plex Mono',monospace;opacity:${isTemp?1:0.4}"
      oninput="updateVeh(this)"></td>`;

    // Zone KH/OZ (MANAVETTE minibus uniquement)
    const isManaMinibus = p === 'MANAVETTE' && vt.includes('minibus');
    const zone = v.zone || '';
    const zoneCell = isManaMinibus ? `<td><select data-veh="${i}" data-key="zone" onchange="updateVeh(this)" style="${SEL_STYLE}">
      <option value=""   ${zone===''  ?'selected':''}>—</option>
      <option value="KH" ${zone==='KH'?'selected':''}>KH — Khouribga</option>
      <option value="OZ" ${zone==='OZ'?'selected':''}>OZ — Oued Zem</option>
    </select></td>` : '';

    // Assemble row selon type véhicule
    let cells = mleCell + kmcCell;
    if(isAutocar)      cells += ageCell;
    else if(isMinicar) cells += kmfCell + kmsCell;
    if(isManaMinibus)  cells += zoneCell;
    cells += dispoCell + daysCell;
    cells += `<td><button class="btn btn-danger btn-sm btn-icon" onclick="removeVeh(this)">✕</button></td>`;
    return `<tr>${cells}</tr>`;
  }).join('');

  const el=document.getElementById('filter-veh-count'); if(el) el.textContent=`${vehs.length} véhicule(s)`;
}

function onDispoChange(sel){
  const i = +sel.dataset.veh;
  const val = sel.value;
  state.vehicles[i].dispo_type = val;
  // update the nb_days input in the same row
  const row = sel.closest('tr');
  const daysInput = row.querySelector('input[data-key="nb_days"]');
  if(!daysInput) return;
  const isTemp = val === 'temporaire';
  daysInput.disabled    = !isTemp;
  daysInput.placeholder = isTemp ? '' : '—';
  daysInput.value       = isTemp ? (state.vehicles[i].nb_days || 0) : '';
  daysInput.style.borderColor = isTemp ? 'var(--warning)' : 'var(--border)';
  daysInput.style.color       = isTemp ? 'var(--warning)' : 'var(--muted)';
  daysInput.style.opacity     = isTemp ? '1' : '0.4';
  if(!isTemp) state.vehicles[i].nb_days = 0;
}

function updateLine(el){ state.lines[+el.dataset.row][el.dataset.key]=el.value; }
function updateVeh(el) { state.vehicles[+el.dataset.veh][el.dataset.key]=el.value; }

function removeLine(btn){
  // Trouver l'index réel dans le tbody (visible ou non)
  const tr  = btn.closest('tr');
  const all = [...document.querySelectorAll('#lines-body tr')];
  const i   = all.indexOf(tr);
  if(i < 0) return;
  state.lines.splice(i, 1);
  const filterVal = document.getElementById('filter-lines')?.value || '';
  state.sotreg ? renderLinesSotreg(state.lines) : renderLines(state.lines);
  if(filterVal) applyFilter('lines-body', filterVal);
}

function removeVeh(btn){
  const tr  = btn.closest('tr');
  const all = [...document.querySelectorAll('#veh-body tr')];
  const i   = all.indexOf(tr);
  if(i < 0) return;
  state.vehicles.splice(i, 1);
  const filterVal = document.getElementById('filter-veh')?.value || '';
  renderVehicles(state.vehicles,
    document.getElementById('sel-provider').value,
    document.getElementById('sel-vtype').value);
  if(filterVal) applyFilter('veh-body', filterVal);
}

function addLine(){
  // Pré-remplissage intelligent depuis le contexte du filtre
  let prefill = {chauffeur:'', mle_car:'', entity:'', circuit:'', km_per_rotation:0, rotation_total:0};

  if(!state.sotreg){
    const filterVal = (document.getElementById('filter-lines')?.value || '').trim();
    if(filterVal){
      // Chercher les lignes visibles pour extraire chauffeur + mle_car + entité + km_rotation
      const visibleRows = [...document.querySelectorAll('#lines-body tr')]
        .filter(tr => tr.style.display !== 'none');
      if(visibleRows.length > 0){
        // Lire les valeurs du premier input de chaque colonne de la première ligne visible
        const inputs = visibleRows[0].querySelectorAll('input');
        // cols: chauffeur, mle_car, entity, circuit, km_per_rotation, rotation_total
        const vals = [...inputs].map(i => i.value);
        prefill.chauffeur       = vals[0] || '';
        prefill.mle_car         = vals[1] || '';
        prefill.entity          = vals[2] || '';
        // circuit vide — c'est ce que l'utilisateur veut saisir
        prefill.circuit         = '';
        prefill.km_per_rotation = vals[4] || 0;
        prefill.rotation_total  = 0;  // à saisir
      }
    }
  }

  if(state.sotreg){
    state.lines.unshift({entity:prefill.entity, circuit:'', nb_vehicles:0,
                          km_per_rotation: prefill.km_per_rotation, rotation_total:0});
    renderLinesSotreg(state.lines);
  } else {
    state.lines.unshift(prefill);
    renderLines(state.lines);
  }

  // Réappliquer le filtre pour garder la nouvelle ligne visible
  const filterVal = document.getElementById('filter-lines')?.value || '';
  if(filterVal) applyFilter('lines-body', filterVal);

  // Focus sur le champ Circuit (index 3) de la nouvelle ligne
  const newRow = document.querySelector('#lines-body tr:first-child');
  if(newRow){
    newRow.scrollIntoView({behavior:'smooth', block:'nearest'});
    // Focus circuit (4e input, index 3) si mle pré-rempli, sinon 1er input
    const inputs = newRow.querySelectorAll('input');
    const focusIdx = prefill.mle_car ? 3 : 0;
    if(inputs[focusIdx]) inputs[focusIdx].focus();
    // Highlight visuel de la nouvelle ligne
    newRow.style.background = 'rgba(59,130,246,0.15)';
    setTimeout(() => newRow.style.background = '', 1500);
  }
}

function addVeh(){
  state.vehicles.unshift({mle_car:'',km_compteur:0,km_forfait:0,km_supp:0,age_cat:'',dispo_type:'permanent',nb_days:0});
  renderVehicles(state.vehicles,
    document.getElementById('sel-provider').value,
    document.getElementById('sel-vtype').value);
  // Réappliquer le filtre véhicules
  const filterVal = document.getElementById('filter-veh')?.value || '';
  if(filterVal) applyFilter('veh-body', filterVal);
  // Focus sur le premier input de la nouvelle ligne
  const firstInput = document.querySelector('#veh-body tr:first-child input');
  if(firstInput){ firstInput.scrollIntoView({behavior:'smooth',block:'nearest'}); firstInput.focus(); }
}

async function saveData(){
  const provider = document.getElementById('sel-provider').value;
  const vtype    = document.getElementById('sel-vtype').value;
  const period   = getPeriod();
  const sp = document.getElementById('save-spinner');
  sp.style.display='inline-block';
  try {
    const r = await fetch('/api/saisie/save',{
      method:'POST', headers:{'Content-Type':'application/json'},
      body: JSON.stringify({provider, vtype, period,
                             tariffs:  state.tariffs,
                             lines:    state.lines,
                             vehicles: state.vehicles})
    });
    const d = await r.json();
    if(d.ok){
      showToast(d.msg,'success');
      document.getElementById('save-status').textContent = 'Sauvegardé à '+new Date().toLocaleTimeString();
      document.getElementById('prev-hint').style.display = 'none';
      // Add period to known list so hint won't re-appear
      if(!KNOWN_PERIODS.includes(period)) KNOWN_PERIODS.push(period);
    } else { showToast(d.error||'Erreur','error'); }
  } catch(e){ showToast('Erreur: '+e,'error'); }
  finally { sp.style.display='none'; }
}

async function calcExcel(){
  const period = getPeriod();
  if(!period){ showToast('Saisir une période','error'); return false; }
  window.location = `/exports/excel_mois?period=${encodeURIComponent(period)}`;
  return false;
}

// ── Filtre de tableau ──────────────────────────────────────────────────────────
function applyFilter(tbodyId, text){
  const tbody  = document.getElementById(tbodyId);
  if(!tbody) return;
  const rows   = tbody.querySelectorAll('tr');
  const q      = text.trim().toLowerCase();
  let visible  = 0;
  rows.forEach(tr => {
    const match = !q || [...tr.querySelectorAll('input,select')].some(el =>
      (el.value||'').toLowerCase().includes(q)
    );
    tr.style.display = match ? '' : 'none';
    if(match) visible++;
  });
  // update count badge
  const countId = tbodyId==='lines-body' ? 'filter-lines-count' : 'filter-veh-count';
  const el = document.getElementById(countId);
  if(el) el.textContent = q ? `${visible} / ${rows.length} ligne(s)` : `${rows.length} ligne(s)`;
}

function clearFilter(inputId, tbodyId){
  const inp = document.getElementById(inputId);
  if(inp){ inp.value=''; applyFilter(tbodyId,''); inp.focus(); }
}

function resetFilters(){
  clearFilter('filter-lines','lines-body');
  clearFilter('filter-veh','veh-body');
}

// Init
onPeriodPick(document.getElementById('sel-period-pick').value);
loadLists();
loadProviderVtypes();
loadCircuitKm();
loadEntityCircuits();
</script>
</body></html>"""

EXPORTS_HTML = """<!doctype html><html><head>""" + _BASE_STYLE + """
<title>SOTREG · Exports</title>
<style>
.export-grid { display:grid; grid-template-columns:repeat(auto-fill,minmax(320px,1fr)); gap:16px; }
.export-card { background:var(--bg2); border:1px solid var(--border); border-radius:var(--radius);
               padding:20px; }
.export-card h3 { font-size:14px; font-weight:600; margin-bottom:6px; }
.export-card p  { font-size:12px; color:var(--muted); margin-bottom:14px; }
.export-card .form-group { margin-bottom:10px; }
.export-card select, .export-card input { width:100%; }
.icon { font-size:24px; margin-bottom:10px; display:block; }
</style>
</head><body>
""" + _NAV.replace("{{ active }}", "") + """
<div class="container">
  <div class="page-title">Exports</div>
  {% with msgs = get_flashed_messages(with_categories=true) %}{% for cat,msg in msgs %}
    <div class="alert alert-{{ cat }}">{{ msg }}</div>{% endfor %}{% endwith %}

  <!-- DB Source -->
  {% if db_name == 'sotreg_work.db' %}
  <div style="border:1px solid #92400e;border-radius:8px;background:#431407;
              padding:14px 18px;margin-bottom:16px;display:flex;align-items:center;gap:14px;flex-wrap:wrap">
    <div style="flex:1;min-width:200px">
      <div style="font-size:13px;font-weight:600;color:#fbbf24;margin-bottom:3px">
        ⚠ Aucune DB importée
      </div>
      <div style="font-size:12px;color:#d97706">
        Importez la base de données envoyée par l'utilisateur 1 avant de générer les exports.
      </div>
    </div>
    <form method="post" action="/exports/upload_db" enctype="multipart/form-data"
          style="display:flex;gap:8px;align-items:center">
      <input type="file" name="dbfile" accept=".db"
             style="font-size:12px;color:var(--muted);min-width:180px">
      <button type="submit" class="btn btn-primary btn-sm">Charger la DB</button>
    </form>
  </div>
  {% else %}
  <div style="border:1px solid #166534;border-radius:8px;background:#052e16;
              padding:12px 18px;margin-bottom:16px;display:flex;align-items:center;
              gap:12px;flex-wrap:wrap">
    <div style="flex:1">
      <span style="font-size:12px;color:#86efac">✓ DB active : </span>
      <span class="badge badge-green text-mono">{{ db_name }}</span>
    </div>
    <form method="post" action="/exports/upload_db" enctype="multipart/form-data"
          style="display:flex;gap:8px;align-items:center">
      <input type="file" name="dbfile" accept=".db" style="font-size:12px;color:var(--muted)">
      <button type="submit" class="btn btn-secondary btn-sm">Changer de DB</button>
    </form>
    <a href="/exports/reset_db" class="btn btn-secondary btn-sm">↺ DB locale</a>
  </div>
  {% endif %}

  <div class="export-grid">

    <!-- Excel Mois -->
    <div class="export-card">
      <span class="icon">📊</span>
      <h3>Excel — Mois complet</h3>
      <p>Génère 4 feuilles: synthèse, entity_circuit, detail_prestataire, sotreg_lines.</p>
      <div class="form-group">
        <label>Période</label>
        <select id="em-period">
          {% for p in periods %}<option>{{ p }}</option>{% endfor %}
          <option value="">Autre…</option>
        </select>
      </div>
      <div class="form-group" id="em-custom" style="display:none">
        <label>Période (YYYY-MM)</label>
        <input type="text" id="em-period-custom" placeholder="2026-03">
      </div>
      <button class="btn btn-success" onclick="dlExcelMois()">⬇ Télécharger</button>
    </div>

    <!-- Attachement Prestataire -->
    <div class="export-card">
      <span class="icon">📎</span>
      <h3>Attachement Prestataire (Excel)</h3>
      <p>Annexe + détail de vérification pour un prestataire et une entité.</p>
      <div class="form-group"><label>Période</label>
        <select id="ap-period">{% for p in periods %}<option>{{ p }}</option>{% endfor %}</select></div>
      <div class="form-group"><label>Entité</label>
        <select id="ap-entity">{% for e in entities %}<option>{{ e }}</option>{% endfor %}</select></div>
      <div class="form-group"><label>Prestataire</label>
        <select id="ap-provider">{% for p in providers %}<option>{{ p }}</option>{% endfor %}</select></div>
      <button class="btn btn-success" onclick="dlAttachExcel()">⬇ Télécharger</button>
    </div>

    <!-- Attachement Global -->
    <div class="export-card">
      <span class="icon">🗂</span>
      <h3>Attachement Global (Excel)</h3>
      <p>Tous prestataires pour une entité et une période — 4 feuilles brutes.</p>
      <div class="form-group"><label>Période</label>
        <select id="ag-period">{% for p in periods %}<option>{{ p }}</option>{% endfor %}</select></div>
      <div class="form-group"><label>Entité</label>
        <select id="ag-entity">{% for e in entities %}<option>{{ e }}</option>{% endfor %}</select></div>
      <button class="btn btn-success" onclick="dlAttachGlobal()">⬇ Télécharger</button>
    </div>

    <!-- PDF Attachement -->
    <div class="export-card">
      <span class="icon">📄</span>
      <h3>Attachement PDF</h3>
      <p>PDF 2 pages : facture + détail de vérification + grille tarifaire.</p>
      <div class="form-group"><label>Période</label>
        <select id="pp-period">{% for p in periods %}<option>{{ p }}</option>{% endfor %}</select></div>
      <div class="form-group"><label>Entité</label>
        <select id="pp-entity">{% for e in entities %}<option>{{ e }}</option>{% endfor %}</select></div>
      <div class="form-group"><label>Prestataire</label>
        <select id="pp-provider">{% for p in providers %}<option>{{ p }}</option>{% endfor %}</select></div>
      <button class="btn btn-danger" style="background:#7c2d12;color:#fed7aa;border-color:#7c2d12" onclick="dlPDF()">⬇ PDF Attachement</button>
    </div>

    <!-- PDF Global sans détail -->
    <div class="export-card">
      <span class="icon">📋</span>
      <h3>PDF Global (sans détail)</h3>
      <p>PDF multi-prestataires pour une entité — facture uniquement, sans page de vérification.</p>
      <div class="form-group"><label>Période</label>
        <select id="pg-period">{% for p in periods %}<option>{{ p }}</option>{% endfor %}</select></div>
      <div class="form-group"><label>Entité</label>
        <select id="pg-entity">{% for e in entities %}<option>{{ e }}</option>{% endfor %}</select></div>
      <button class="btn btn-danger" style="background:#7c2d12;color:#fed7aa;border-color:#7c2d12" onclick="dlPDFGlobal()">⬇ PDF Global</button>
    </div>

  </div>
</div>

<script>
function gv(id){ return document.getElementById(id)?.value||''; }

document.getElementById('em-period').addEventListener('change',function(){
  document.getElementById('em-custom').style.display = this.value==='' ? '' : 'none';
});

function dlExcelMois(){
  let p = gv('em-period');
  if(!p) p = gv('em-period-custom');
  if(!p){ showToast('Sélectionner une période','error'); return; }
  window.location = `/exports/excel_mois?period=${encodeURIComponent(p)}`;
  showToast('Génération en cours…','success');
}
function dlAttachExcel(){
  const p=gv('ap-period'), e=gv('ap-entity'), pr=gv('ap-provider');
  if(!p||!e||!pr){ showToast('Tous les champs requis','error'); return; }
  window.location = `/exports/attachment_excel?period=${encodeURIComponent(p)}&entity=${encodeURIComponent(e)}&provider=${encodeURIComponent(pr)}`;
  showToast('Génération en cours…','success');
}
function dlAttachGlobal(){
  const p=gv('ag-period'), e=gv('ag-entity');
  if(!p||!e){ showToast('Période et entité requises','error'); return; }
  window.location = `/exports/attachment_global?period=${encodeURIComponent(p)}&entity=${encodeURIComponent(e)}`;
  showToast('Génération en cours…','success');
}
function dlPDF(){
  const p=gv('pp-period'), e=gv('pp-entity'), pr=gv('pp-provider');
  if(!p||!e||!pr){ showToast('Tous les champs requis','error'); return; }
  window.location = `/exports/pdf?period=${encodeURIComponent(p)}&entity=${encodeURIComponent(e)}&provider=${encodeURIComponent(pr)}`;
  showToast('Génération PDF en cours…','success');
}
function dlPDFGlobal(){
  const p=gv('pg-period'), e=gv('pg-entity');
  if(!p||!e){ showToast('Période et entité requises','error'); return; }
  window.location = `/exports/pdf_global?period=${encodeURIComponent(p)}&entity=${encodeURIComponent(e)}`;
  showToast('Génération PDF Global en cours…','success');
}
</script>
</body></html>"""

# ── Entrée principale ────────────────────────────────────────────────────────────
if __name__ == "__main__":
    if not WORK_DB.exists():
        if TMPL_DB.exists():
            shutil.copy(str(TMPL_DB), str(WORK_DB))
            print(f"[SOTREG] DB initialisée depuis template: {WORK_DB}")
        else:
            print(f"[SOTREG] ATTENTION: {WORK_DB} introuvable.")
    import socket
    try:
        local_ip = socket.gethostbyname(socket.gethostname())
    except:
        local_ip = "127.0.0.1"
    print("=" * 55)
    print("  SOTREG Web — Application de facturation")
    print("=" * 55)
    print(f"  Accès local  : http://127.0.0.1:5000")
    print(f"  Accès réseau : http://{local_ip}:5000")
    print()
    print("  Comptes utilisateurs:")
    print("  user1 / pass1  →  Saisie")
    print("  user2 / pass2  →  Exports")
    print("  admin / admin  →  Administration")
    print("=" * 55)
    app.run(debug=False, host="0.0.0.0", port=5000)
