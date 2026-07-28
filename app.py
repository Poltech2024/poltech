# -*- coding: utf-8 -*-
"""
POLTECH - Sistema de Control de Nomina
Version 1 (cimiento): Login + Roles, Obras, Catalogo de sueldos,
Personal (con validacion de NSS y CURP) y Cuentas bancarias (sin duplicados).

Hecho para correr en Flask + gunicorn (Render). Base de datos: SQLite.
NOTA: en el plan gratis de Render los datos de SQLite NO son permanentes.
Usa datos de prueba hasta conectar una base de datos permanente.
"""
import os
import re
import sqlite3
from datetime import date
from functools import wraps

from flask import (Flask, request, redirect, url_for, session,
                   flash, render_template, g, abort)
from werkzeug.security import generate_password_hash, check_password_hash

APP_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(APP_DIR, "poltech.db")

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "dev-cambia-esta-clave-en-render")

# ---------------------------------------------------------------------------
# Roles y jerarquia (a mayor numero, mas permisos)
# ---------------------------------------------------------------------------
ROLES = {
    "admin":        ("Administrador", 100),
    "direccion":    ("Direccion General", 90),
    "gerencia":     ("Gerencia General", 85),
    "gerente_obra": ("Gerente de obra", 70),
    "autorizador":  ("Autorizador", 50),
    "capturista":   ("Capturista (Residente/Contable)", 30),
}
GERENTE_RANK = 70  # de gerente de obra hacia arriba

def role_label(r): return ROLES.get(r, (r, 0))[0]
def role_rank(r):  return ROLES.get(r, (r, 0))[1]

ESTADOS = ["Vallarta", "Oaxaca"]
TIPOS_CUENTA = ["Tarjeta", "CLABE interbancaria", "Numero de cuenta"]

# ---------------------------------------------------------------------------
# Base de datos
# ---------------------------------------------------------------------------
SCHEMA = """
CREATE TABLE IF NOT EXISTS users (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    username TEXT UNIQUE NOT NULL,
    password_hash TEXT NOT NULL,
    nombre TEXT NOT NULL,
    role TEXT NOT NULL DEFAULT 'capturista',
    obra_id INTEGER
);
CREATE TABLE IF NOT EXISTS obras (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    proyecto TEXT,
    contrato TEXT,
    nombre TEXT NOT NULL,
    estado TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS puestos (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    obra_id INTEGER NOT NULL,
    nombre TEXT NOT NULL,
    sueldo_semanal REAL NOT NULL DEFAULT 0,
    viaticos_semanales REAL NOT NULL DEFAULT 0,
    FOREIGN KEY (obra_id) REFERENCES obras(id)
);
CREATE TABLE IF NOT EXISTS empleados (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    ref TEXT,
    nombre TEXT NOT NULL,
    primer_apellido TEXT NOT NULL,
    segundo_apellido TEXT,
    curp TEXT,
    rfc TEXT,
    cp_fiscal TEXT,
    nss TEXT,
    sexo TEXT,
    fecha_nacimiento TEXT,
    puesto_id INTEGER,
    fecha_alta TEXT,
    importe_alta_imss REAL DEFAULT 0,
    infonavit_monto REAL DEFAULT 0,
    exime_docs INTEGER DEFAULT 0,
    autoriza_tercero TEXT,
    estatus TEXT DEFAULT 'activo',
    FOREIGN KEY (puesto_id) REFERENCES puestos(id)
);
CREATE TABLE IF NOT EXISTS cuentas_bancarias (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    empleado_id INTEGER NOT NULL,
    institucion TEXT NOT NULL,
    tipo_cuenta TEXT NOT NULL,
    numero TEXT UNIQUE NOT NULL,
    FOREIGN KEY (empleado_id) REFERENCES empleados(id)
);
"""

def get_db():
    if "db" not in g:
        g.db = sqlite3.connect(DB_PATH)
        g.db.row_factory = sqlite3.Row
        g.db.execute("PRAGMA foreign_keys = ON")
    return g.db

@app.teardown_appcontext
def close_db(exc):
    db = g.pop("db", None)
    if db is not None:
        db.close()

def init_db():
    db = sqlite3.connect(DB_PATH)
    db.executescript(SCHEMA)
    # Migracion suave: agrega obra_id a users si la base venia de una version previa
    try:
        db.execute("ALTER TABLE users ADD COLUMN obra_id INTEGER")
    except sqlite3.OperationalError:
        pass
    # Crear usuario administrador la primera vez
    row = db.execute("SELECT COUNT(*) FROM users").fetchone()
    if row[0] == 0:
        pw = os.environ.get("ADMIN_PASSWORD", "cambiar123")
        db.execute(
            "INSERT INTO users(username, password_hash, nombre, role) VALUES(?,?,?,?)",
            ("admin", generate_password_hash(pw), "Administrador", "admin"),
        )
    db.commit()
    db.close()

init_db()

# ---------------------------------------------------------------------------
# Autenticacion y permisos
# ---------------------------------------------------------------------------
def login_required(f):
    @wraps(f)
    def wrapper(*a, **k):
        if not session.get("user_id"):
            return redirect(url_for("login"))
        return f(*a, **k)
    return wrapper

def min_rank(rank):
    """Decorador: exige un rango minimo de rol."""
    def deco(f):
        @wraps(f)
        def wrapper(*a, **k):
            if not session.get("user_id"):
                return redirect(url_for("login"))
            if role_rank(session.get("role", "")) < rank:
                abort(403)
            return f(*a, **k)
        return wrapper
    return deco

@app.context_processor
def inject_user():
    return {
        "current_user": {
            "id": session.get("user_id"),
            "nombre": session.get("nombre"),
            "role": session.get("role"),
            "role_label": role_label(session.get("role", "")),
            "rank": role_rank(session.get("role", "")),
        },
        "GERENTE_RANK": GERENTE_RANK,
    }

# ---------------------------------------------------------------------------
# Validaciones (NSS y CURP)  -- sin librerias externas
# ---------------------------------------------------------------------------
def solo_digitos(s):
    return re.sub(r"\D", "", s or "")

def luhn_ok(num):
    total, alt = 0, False
    for ch in reversed(num):
        d = int(ch)
        if alt:
            d *= 2
            if d > 9:
                d -= 9
        total += d
        alt = not alt
    return total % 10 == 0

def nss_valido(nss):
    n = solo_digitos(nss)
    return len(n) == 11 and luhn_ok(n)

CURP_RE = re.compile(
    r"^[A-Z][AEIOUX][A-Z]{2}\d{2}(0[1-9]|1[0-2])(0[1-9]|[12]\d|3[01])"
    r"[HM][A-Z]{2}[B-DF-HJ-NP-TV-Z]{3}[A-Z0-9]\d$"
)

def curp_valida(curp):
    return bool(CURP_RE.match((curp or "").strip().upper()))

def curp_sexo(curp):
    return "H" if curp[10].upper() == "H" else "M"

def curp_fecha(curp):
    try:
        yy = int(curp[4:6]); mm = int(curp[6:8]); dd = int(curp[8:10])
        siglo = 2000 if not curp[16].isdigit() else 1900
        return date(siglo + yy, mm, dd).isoformat()
    except Exception:
        return None

def revisar_curp(curp, nombre, ap1, ap2, sexo, fecha_nac):
    """Devuelve una lista de advertencias (no bloquea)."""
    avisos = []
    curp = (curp or "").strip().upper()
    if not curp_valida(curp):
        return avisos  # el formato ya se valida aparte
    def inicial(x):
        x = (x or "").strip().upper()
        return x[0] if x else ""
    if ap1 and curp[0] != inicial(ap1):
        avisos.append("La CURP no coincide con la inicial del primer apellido.")
    if ap2 and curp[2] not in (inicial(ap2), "X"):
        avisos.append("La CURP no coincide con la inicial del segundo apellido.")
    if nombre and curp[3] not in (inicial(nombre), "X"):
        avisos.append("La CURP no coincide con la inicial del nombre.")
    if sexo and curp_sexo(curp) != sexo:
        avisos.append("El sexo no coincide con el que indica la CURP.")
    if fecha_nac and curp_fecha(curp) and curp_fecha(curp) != fecha_nac:
        avisos.append("La fecha de nacimiento no coincide con la que indica la CURP.")
    return avisos

# ---------------------------------------------------------------------------
# Rutas: sesion
# ---------------------------------------------------------------------------
@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        u = request.form.get("username", "").strip()
        p = request.form.get("password", "")
        row = get_db().execute("SELECT * FROM users WHERE username=?", (u,)).fetchone()
        if row and check_password_hash(row["password_hash"], p):
            session["user_id"] = row["id"]
            session["nombre"] = row["nombre"]
            session["role"] = row["role"]
            return redirect(url_for("dashboard"))
        flash("Usuario o contrasena incorrectos.", "danger")
    return render_template("login.html")

@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))

# ---------------------------------------------------------------------------
# Dashboard
# ---------------------------------------------------------------------------
@app.route("/")
@login_required
def dashboard():
    db = get_db()
    stats = {
        "empleados": db.execute("SELECT COUNT(*) FROM empleados WHERE estatus='activo'").fetchone()[0],
        "obras": db.execute("SELECT COUNT(*) FROM obras").fetchone()[0],
        "puestos": db.execute("SELECT COUNT(*) FROM puestos").fetchone()[0],
        "cuentas": db.execute("SELECT COUNT(*) FROM cuentas_bancarias").fetchone()[0],
    }
    usa_default = os.environ.get("ADMIN_PASSWORD") is None
    return render_template("dashboard.html", stats=stats, usa_default=usa_default)

# ---------------------------------------------------------------------------
# Obras
# ---------------------------------------------------------------------------
@app.route("/obras", methods=["GET", "POST"])
@login_required
def obras():
    db = get_db()
    if request.method == "POST":
        if role_rank(session["role"]) < GERENTE_RANK:
            abort(403)
        db.execute(
            "INSERT INTO obras(proyecto, contrato, nombre, estado) VALUES(?,?,?,?)",
            (request.form.get("proyecto", "").strip(),
             request.form.get("contrato", "").strip(),
             request.form.get("nombre", "").strip(),
             request.form.get("estado", "")),
        )
        db.commit()
        flash("Obra registrada.", "success")
        return redirect(url_for("obras"))
    filas = db.execute("SELECT * FROM obras ORDER BY nombre").fetchall()
    return render_template("obras_list.html", obras=filas, estados=ESTADOS)

# ---------------------------------------------------------------------------
# Catalogo de sueldos por obra (solo gerente de obra hacia arriba)
# ---------------------------------------------------------------------------
@app.route("/catalogo", methods=["GET", "POST"])
@login_required
def catalogo():
    db = get_db()
    if request.method == "POST":
        if role_rank(session["role"]) < GERENTE_RANK:
            abort(403)
        db.execute(
            "INSERT INTO puestos(obra_id, nombre, sueldo_semanal, viaticos_semanales) VALUES(?,?,?,?)",
            (request.form.get("obra_id"),
             request.form.get("nombre", "").strip(),
             float(request.form.get("sueldo_semanal") or 0),
             float(request.form.get("viaticos_semanales") or 0)),
        )
        db.commit()
        flash("Puesto agregado al catalogo.", "success")
        return redirect(url_for("catalogo"))
    obras_l = db.execute("SELECT * FROM obras ORDER BY nombre").fetchall()
    puestos = db.execute(
        "SELECT p.*, o.nombre AS obra, o.estado FROM puestos p "
        "JOIN obras o ON o.id=p.obra_id ORDER BY o.nombre, p.nombre"
    ).fetchall()
    return render_template("catalogo_list.html", obras=obras_l, puestos=puestos)

# ---------------------------------------------------------------------------
# Personal
# ---------------------------------------------------------------------------
@app.route("/personal")
@login_required
def personal():
    db = get_db()
    filas = db.execute(
        "SELECT e.*, p.nombre AS puesto, o.nombre AS obra, o.estado AS plaza "
        "FROM empleados e "
        "LEFT JOIN puestos p ON p.id=e.puesto_id "
        "LEFT JOIN obras o ON o.id=p.obra_id "
        "ORDER BY e.primer_apellido, e.nombre"
    ).fetchall()
    return render_template("personal_list.html", empleados=filas)

@app.route("/personal/nuevo", methods=["GET", "POST"])
@login_required
def personal_nuevo():
    db = get_db()
    puestos = db.execute(
        "SELECT p.id, p.nombre AS puesto, p.sueldo_semanal, p.viaticos_semanales, "
        "o.nombre AS obra, o.estado FROM puestos p JOIN obras o ON o.id=p.obra_id "
        "ORDER BY o.nombre, p.nombre"
    ).fetchall()

    if request.method == "POST":
        f = request.form
        nombre = f.get("nombre", "").strip()
        ap1 = f.get("primer_apellido", "").strip()
        ap2 = f.get("segundo_apellido", "").strip()
        curp = f.get("curp", "").strip().upper()
        nss = f.get("nss", "").strip()
        sexo = f.get("sexo", "")
        fecha_nac = f.get("fecha_nacimiento", "")
        exime = 1 if f.get("exime_docs") else 0
        autoriza = f.get("autoriza_tercero", "").strip()

        errores = []
        if not nombre: errores.append("El nombre es obligatorio.")
        if not ap1: errores.append("El primer apellido es obligatorio.")
        if not f.get("puesto_id"): errores.append("Selecciona un puesto (obra + categoria).")
        if not f.get("fecha_alta"): errores.append("La fecha de alta es obligatoria.")

        # CURP: formato obligatorio
        if not curp_valida(curp):
            errores.append("La CURP no tiene un formato valido (18 caracteres).")

        # NSS: obligatorio salvo que se exima con autorizacion de un tercero
        if exime:
            if not autoriza:
                errores.append("Para eximir documentos, indica quien autoriza (tercero).")
        else:
            if not nss:
                errores.append("El NSS es obligatorio (o marca 'eximir' con autorizacion).")
            elif not nss_valido(nss):
                errores.append("El NSS no es valido (deben ser 11 digitos con digito verificador correcto).")

        # Cuenta bancaria opcional
        cta_num = solo_digitos(f.get("cta_numero", ""))
        cta_inst = f.get("cta_institucion", "").strip()
        cta_tipo = f.get("cta_tipo", "")
        if cta_num:
            existe = db.execute("SELECT 1 FROM cuentas_bancarias WHERE numero=?", (cta_num,)).fetchone()
            if existe:
                errores.append("Esa cuenta bancaria ya esta registrada (no se permiten duplicados).")
            if not cta_inst or not cta_tipo:
                errores.append("Para la cuenta bancaria, captura institucion y tipo de cuenta.")

        if errores:
            for e in errores:
                flash(e, "danger")
            return render_template("personal_form.html", puestos=puestos,
                                   estados=ESTADOS, tipos=TIPOS_CUENTA, datos=f)

        cur = db.execute(
            """INSERT INTO empleados
               (ref, nombre, primer_apellido, segundo_apellido, curp, rfc, cp_fiscal,
                nss, sexo, fecha_nacimiento, puesto_id, fecha_alta, importe_alta_imss,
                infonavit_monto, exime_docs, autoriza_tercero)
               VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
            (f.get("ref", "").strip(), nombre, ap1, ap2, curp,
             f.get("rfc", "").strip().upper(), f.get("cp_fiscal", "").strip(),
             solo_digitos(nss), sexo, fecha_nac, f.get("puesto_id"),
             f.get("fecha_alta"), float(f.get("importe_alta_imss") or 0),
             float(f.get("infonavit_monto") or 0), exime, autoriza),
        )
        emp_id = cur.lastrowid
        if cta_num:
            db.execute(
                "INSERT INTO cuentas_bancarias(empleado_id, institucion, tipo_cuenta, numero) VALUES(?,?,?,?)",
                (emp_id, cta_inst, cta_tipo, cta_num),
            )
        db.commit()

        for a in revisar_curp(curp, nombre, ap1, ap2, sexo, fecha_nac):
            flash("Aviso: " + a, "warning")
        flash("Empleado dado de alta correctamente.", "success")
        return redirect(url_for("personal"))

    return render_template("personal_form.html", puestos=puestos,
                           estados=ESTADOS, tipos=TIPOS_CUENTA, datos={})

# ---------------------------------------------------------------------------
# Cuentas bancarias
# ---------------------------------------------------------------------------
@app.route("/cuentas")
@login_required
def cuentas():
    db = get_db()
    filas = db.execute(
        "SELECT c.*, e.nombre, e.primer_apellido, e.segundo_apellido "
        "FROM cuentas_bancarias c JOIN empleados e ON e.id=c.empleado_id "
        "ORDER BY e.primer_apellido"
    ).fetchall()
    return render_template("cuentas_list.html", cuentas=filas)

# ---------------------------------------------------------------------------
# Cambiar contrasena
# ---------------------------------------------------------------------------
@app.route("/cuenta/password", methods=["GET", "POST"])
@login_required
def cambiar_password():
    db = get_db()
    if request.method == "POST":
        actual = request.form.get("actual", "")
        nueva = request.form.get("nueva", "")
        row = db.execute("SELECT * FROM users WHERE id=?", (session["user_id"],)).fetchone()
        if not check_password_hash(row["password_hash"], actual):
            flash("La contrasena actual no es correcta.", "danger")
        elif len(nueva) < 6:
            flash("La nueva contrasena debe tener al menos 6 caracteres.", "danger")
        else:
            db.execute("UPDATE users SET password_hash=? WHERE id=?",
                       (generate_password_hash(nueva), session["user_id"]))
            db.commit()
            flash("Contrasena actualizada.", "success")
            return redirect(url_for("dashboard"))
    return render_template("cambiar_password.html")

# ---------------------------------------------------------------------------
# Usuarios (solo el administrador puede crear cuentas y asignar roles)
# ---------------------------------------------------------------------------
ADMIN_RANK = 100  # solo administrador gestiona usuarios

@app.route("/usuarios", methods=["GET", "POST"])
@min_rank(ADMIN_RANK)
def usuarios():
    db = get_db()
    if request.method == "POST":
        u = request.form.get("username", "").strip().lower()
        nombre = request.form.get("nombre", "").strip()
        pw = request.form.get("password", "")
        role = request.form.get("role", "")
        obra_id = request.form.get("obra_id") or None

        errores = []
        if not u: errores.append("El usuario es obligatorio.")
        if not nombre: errores.append("El nombre es obligatorio.")
        if len(pw) < 6: errores.append("La contrasena debe tener al menos 6 caracteres.")
        if role not in ROLES: errores.append("Selecciona un rol valido.")
        if db.execute("SELECT 1 FROM users WHERE username=?", (u,)).fetchone():
            errores.append("Ese usuario ya existe.")

        if errores:
            for e in errores:
                flash(e, "danger")
        else:
            db.execute(
                "INSERT INTO users(username, password_hash, nombre, role, obra_id) VALUES(?,?,?,?,?)",
                (u, generate_password_hash(pw), nombre, role, obra_id),
            )
            db.commit()
            flash("Usuario creado: " + u, "success")
        return redirect(url_for("usuarios"))

    filas = db.execute(
        "SELECT us.*, o.nombre AS obra FROM users us "
        "LEFT JOIN obras o ON o.id=us.obra_id ORDER BY us.nombre"
    ).fetchall()
    obras_l = db.execute("SELECT * FROM obras ORDER BY nombre").fetchall()
    return render_template("usuarios_list.html", usuarios=filas, obras=obras_l, roles=ROLES)

@app.route("/usuarios/<int:uid>/rol", methods=["POST"])
@min_rank(ADMIN_RANK)
def usuario_rol(uid):
    db = get_db()
    role = request.form.get("role", "")
    obra_id = request.form.get("obra_id") or None
    if uid == session["user_id"]:
        flash("No puedes cambiar tu propio rol.", "warning")
    elif role in ROLES:
        db.execute("UPDATE users SET role=?, obra_id=? WHERE id=?", (role, obra_id, uid))
        db.commit()
        flash("Rol actualizado.", "success")
    return redirect(url_for("usuarios"))

@app.route("/usuarios/<int:uid>/password", methods=["POST"])
@min_rank(ADMIN_RANK)
def usuario_password(uid):
    db = get_db()
    nueva = request.form.get("nueva", "")
    if len(nueva) < 6:
        flash("La contrasena debe tener al menos 6 caracteres.", "danger")
    else:
        db.execute("UPDATE users SET password_hash=? WHERE id=?",
                   (generate_password_hash(nueva), uid))
        db.commit()
        flash("Contrasena restablecida.", "success")
    return redirect(url_for("usuarios"))

@app.route("/usuarios/<int:uid>/eliminar", methods=["POST"])
@min_rank(ADMIN_RANK)
def usuario_eliminar(uid):
    db = get_db()
    if uid == session["user_id"]:
        flash("No puedes eliminar tu propia cuenta.", "warning")
    else:
        db.execute("DELETE FROM users WHERE id=?", (uid,))
        db.commit()
        flash("Usuario eliminado.", "success")
    return redirect(url_for("usuarios"))

@app.errorhandler(403)
def forbidden(e):
    return render_template("403.html"), 403

if __name__ == "__main__":
    app.run(debug=True, port=5000)
