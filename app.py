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
import io
import json
import sqlite3
import smtplib
import urllib.request
import urllib.error
import urllib.parse
from datetime import date, datetime, timedelta
from functools import wraps
from email.message import EmailMessage

import openpyxl
from flask import (Flask, request, redirect, url_for, session,
                   flash, render_template, g, abort, send_file, jsonify)
from werkzeug.security import generate_password_hash, check_password_hash

APP_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.environ.get("DB_PATH", os.path.join(APP_DIR, "poltech.db"))

APP_VERSION = "1.15"   # version del sistema (visible en el menu)

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "dev-cambia-esta-clave-en-render")
app.config["MAX_CONTENT_LENGTH"] = 6 * 1024 * 1024  # 6 MB por archivo subido

# ---------------------------------------------------------------------------
# Roles y jerarquia (a mayor numero, mas permisos)
# ---------------------------------------------------------------------------
ROLES = {
    "admin":           ("Administrador", 100),
    "superintendente": ("Superintendente de obra", 70),
    "residente":       ("Residente", 30),
}
GERENTE_RANK = 70  # de gerente de obra hacia arriba
ADMIN_RANK = 100   # solo administrador

def role_label(r): return ROLES.get(r, (r, 0))[0]
def role_rank(r):  return ROLES.get(r, (r, 0))[1]

def obras_del_usuario(db=None):
    """None = ve todas las obras (administrador).
    Lista de obra_ids = solo esas obras (superintendente/residente). [] = ninguna."""
    if role_rank(session.get("role", "")) >= ADMIN_RANK:
        return None
    db = db or get_db()
    rows = db.execute("SELECT obra_id FROM user_obras WHERE user_id=?",
                      (session.get("user_id"),)).fetchall()
    return [r["obra_id"] for r in rows]

def enviar_correo(asunto, cuerpo, destinatarios, adjuntos=None):
    """Envia un correo a una lista de destinatarios, con adjuntos opcionales.
    adjuntos: lista de (nombre_archivo, bytes, mimetype). Nunca detiene el sistema."""
    host = os.environ.get("SMTP_HOST")
    user = os.environ.get("SMTP_USER")
    pw   = os.environ.get("SMTP_PASS")
    dest = [d for d in (destinatarios or []) if d and "@" in d]
    if not (host and user and pw and dest):
        app.logger.warning("CORREO (no configurado o sin destinatarios): %s | para %s", asunto, dest)
        return False
    try:
        msg = EmailMessage()
        msg["Subject"] = asunto
        msg["From"] = user
        msg["To"] = ", ".join(dest)
        msg.set_content(cuerpo)
        for nombre, datos, mime in (adjuntos or []):
            maintype, _, subtype = (mime or "application/octet-stream").partition("/")
            msg.add_attachment(datos, maintype=maintype, subtype=subtype, filename=nombre)
        port = int(os.environ.get("SMTP_PORT", "587"))
        with smtplib.SMTP(host, port, timeout=15) as s:
            s.starttls()
            s.login(user, pw)
            s.send_message(msg)
        return True
    except Exception as e:
        app.logger.error("No se pudo enviar el correo: %s", e)
        return False

def send_admin_alert(asunto, cuerpo, extra=None):
    """Alerta al administrador (y opcionalmente a mas destinatarios)."""
    dest = [os.environ.get("ADMIN_EMAIL", "")]
    if extra:
        dest += extra
    return enviar_correo(asunto, cuerpo, dest)

def siguiente_cedula(db):
    """Cedula del checador: A001, A002 ... A999, B001 ...
    Unica por trabajador, independiente de la obra."""
    row = db.execute(
        "SELECT cedula FROM empleados WHERE cedula IS NOT NULL AND cedula<>'' "
        "ORDER BY cedula DESC LIMIT 1").fetchone()
    if not row or not row["cedula"]:
        return "A001"
    letra = row["cedula"][0]
    num = int(row["cedula"][1:])
    if num >= 999:
        letra = chr(ord(letra) + 1)
        num = 1
    else:
        num += 1
    return f"{letra}{num:03d}"


def resolver_puesto(db, obra_nom, puesto_nom):
    """Devuelve el id del puesto y el estado de su obra, coincidiendo con obra + categoria
    del catalogo (solo puestos activos), o None si esa combinacion no existe."""
    return db.execute(
        "SELECT p.id, o.estado FROM puestos p JOIN obras o ON o.id=p.obra_id "
        "WHERE lower(o.nombre)=lower(?) AND lower(p.nombre)=lower(?) AND p.activo=1",
        ((obra_nom or "").strip(), (puesto_nom or "").strip())).fetchone()


ESTADOS = [
    "Aguascalientes", "Baja California", "Baja California Sur", "Campeche",
    "Chiapas", "Chihuahua", "Ciudad de Mexico", "Coahuila", "Colima",
    "Durango", "Estado de Mexico", "Guanajuato", "Guerrero", "Hidalgo",
    "Jalisco", "Michoacan", "Morelos", "Nayarit", "Nuevo Leon", "Oaxaca",
    "Puebla", "Queretaro", "Quintana Roo", "San Luis Potosi", "Sinaloa",
    "Sonora", "Tabasco", "Tamaulipas", "Tlaxcala", "Veracruz", "Yucatan",
    "Zacatecas",
]
TIPOS_CUENTA = ["Tarjeta", "CLABE interbancaria", "Numero de cuenta"]

# Clasificaciones del catalogo de sueldos (para totalizar la nomina por grupo)
CLASIFICACIONES = [
    "Gerencia de obra", "Residencia de obra", "Administracion", "Montadores",
    "Soldadores", "Laminadores", "Ayudantes generales", "Operadores",
    "Seguridad", "Almacen", "Topografia",
]

MOTIVOS_BAJA = [
    "Renuncia voluntaria", "Termino de la obra o contrato", "Rescision de contrato",
    "Abandono de empleo", "Ausentismo", "Defuncion", "Otro",
]

def titulo(s):
    """Primera letra de cada palabra en mayuscula y el resto en minuscula.
    Mantiene los acentos y no rompe apostrofes."""
    if not s:
        return s
    return " ".join(w[:1].upper() + w[1:].lower() for w in str(s).split())

SIGLAS_CONOCIDAS = {"ISSTE", "ISSSTE", "IMSS", "INFONAVIT", "SAT", "CFE", "PEMEX",
                    "CDMX", "SEP", "STPS", "UMF", "SA", "CV", "RFC", "CURP", "NSS",
                    "SAPI", "SC", "SRL"}

def titulo_obra(s):
    """Como titulo(), pero conserva las SIGLAS conocidas (ISSTE, IMSS, SAT, ...).
    Sirve para nombres de obra y puesto sin arruinar palabras normales como OBRA."""
    if not s:
        return s
    out = []
    for w in str(s).split():
        if w.upper().strip(".,") in SIGLAS_CONOCIDAS:
            out.append(w.upper())
        else:
            out.append(w[:1].upper() + w[1:].lower())
    return " ".join(out)
# NSS generico para personal pensionado que no requiere alta en el IMSS.
# Se exime de la validacion normal (digito verificador) y de la regla de
# NSS unico, para que varios pensionados puedan compartirlo.
NSS_GENERICO = "00000000000"

# Catalogo de bancos disponibles en Mexico (los mas usados para nomina).
BANCOS = [
    "BBVA Mexico", "Banorte", "Santander", "Citibanamex", "HSBC",
    "Scotiabank", "Inbursa", "Banco Azteca", "BanBajio", "Banregio",
    "Afirme", "Banca Mifel", "Multiva", "CIBanco", "Intercam Banco",
    "Banco Ve por Mas", "Compartamos Banco", "Banco del Bienestar",
    "BanCoppel", "STP", "Klar", "Nu Mexico", "Hey Banco", "Mercado Pago",
    "Otro",
]

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
CREATE TABLE IF NOT EXISTS documentos (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    empleado_id INTEGER NOT NULL,
    tipo TEXT NOT NULL,
    filename TEXT,
    contenido BLOB,
    subido_en TEXT,
    subido_por TEXT,
    UNIQUE(empleado_id, tipo),
    FOREIGN KEY (empleado_id) REFERENCES empleados(id)
);
CREATE TABLE IF NOT EXISTS user_obras (
    user_id INTEGER NOT NULL,
    obra_id INTEGER NOT NULL,
    UNIQUE(user_id, obra_id)
);
-- Fase 2: asistencia semanal (viernes a jueves, base de 6 dias)
CREATE TABLE IF NOT EXISTS asistencia_semanas (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    obra_id INTEGER NOT NULL,
    fecha_inicio TEXT NOT NULL,   -- viernes (ISO)
    fecha_fin TEXT NOT NULL,      -- jueves (ISO)
    semana_num INTEGER,
    anio INTEGER,
    estatus TEXT DEFAULT 'borrador',   -- borrador / cerrada
    creada_en TEXT,
    creada_por TEXT,
    UNIQUE(obra_id, fecha_inicio),
    FOREIGN KEY (obra_id) REFERENCES obras(id)
);
CREATE TABLE IF NOT EXISTS asistencia (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    semana_id INTEGER NOT NULL,
    empleado_id INTEGER NOT NULL,
    d1 TEXT, d2 TEXT, d3 TEXT, d4 TEXT, d5 TEXT, d6 TEXT, d7 TEXT,  -- VIE..JUE
    he_150 REAL DEFAULT 0,   -- horas extra al x1.5
    he_200 REAL DEFAULT 0,   -- horas extra al x2.0
    observaciones TEXT,
    UNIQUE(semana_id, empleado_id),
    FOREIGN KEY (semana_id) REFERENCES asistencia_semanas(id),
    FOREIGN KEY (empleado_id) REFERENCES empleados(id)
);
CREATE TABLE IF NOT EXISTS parametros (
    clave TEXT PRIMARY KEY,
    valor TEXT
);
CREATE TABLE IF NOT EXISTS nominas (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    obra_id INTEGER NOT NULL,
    fecha_inicio TEXT NOT NULL,
    fecha_fin TEXT NOT NULL,
    semana_num INTEGER,
    anio INTEGER,
    jornada REAL,
    aplico_retardos INTEGER DEFAULT 0,
    total_neto REAL DEFAULT 0,
    estatus TEXT DEFAULT 'pendiente',
    autorizada_por TEXT,
    autorizada_en TEXT,
    creada_en TEXT,
    creada_por TEXT,
    FOREIGN KEY (obra_id) REFERENCES obras(id)
);
CREATE TABLE IF NOT EXISTS nomina_detalle (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    nomina_id INTEGER NOT NULL,
    empleado_id INTEGER,
    cedula TEXT,
    nombre TEXT,
    dias REAL DEFAULT 0,
    faltas REAL DEFAULT 0,
    retardos REAL DEFAULT 0,
    vacaciones REAL DEFAULT 0,
    sueldo REAL DEFAULT 0,
    viaticos REAL DEFAULT 0,
    he_horas REAL DEFAULT 0,
    he_importe REAL DEFAULT 0,
    infonavit REAL DEFAULT 0,
    desc_nomina REAL DEFAULT 0,
    desc_otra REAL DEFAULT 0,
    desc_retardos REAL DEFAULT 0,
    neto REAL DEFAULT 0,
    baja_fecha TEXT,
    nota TEXT,
    FOREIGN KEY (nomina_id) REFERENCES nominas(id)
);
CREATE TABLE IF NOT EXISTS clasificaciones (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    nombre TEXT UNIQUE NOT NULL
);
CREATE TABLE IF NOT EXISTS bitacora (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    fecha TEXT,
    usuario TEXT,
    rol TEXT,
    accion TEXT,
    detalle TEXT
);
CREATE TABLE IF NOT EXISTS contratos (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    empleado_id INTEGER NOT NULL,
    nombre_archivo TEXT,
    generado_en TEXT,
    generado_por TEXT,
    estatus TEXT DEFAULT 'generado',   -- generado / revisado
    revisado_por TEXT,
    revisado_en TEXT,
    contenido BLOB,
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
    # Migracion suave: cedula del checador y observaciones en empleados
    for col, tipo in (("cedula", "TEXT"), ("observaciones", "TEXT"),
                      ("estatus_docs", "TEXT DEFAULT 'Pendiente de carga'"),
                      ("fecha_registro", "TEXT"), ("fecha_solicitud", "TEXT"),
                      ("fecha_baja", "TEXT"), ("motivo_baja", "TEXT"),
                      ("viaticos_semanales", "REAL"), ("bono_semanal", "REAL DEFAULT 0"),
                      ("nss_generico", "INTEGER DEFAULT 0"), ("autoriza_nss_generico", "TEXT")):
        try:
            db.execute(f"ALTER TABLE empleados ADD COLUMN {col} {tipo}")
        except sqlite3.OperationalError:
            pass
    # Migracion suave: clasificacion en el catalogo de puestos y en el detalle de nomina
    try:
        db.execute("ALTER TABLE puestos ADD COLUMN clasificacion TEXT")
    except sqlite3.OperationalError:
        pass
    # Migracion suave: puesto activo/inactivo (se puede ocultar del catalogo sin borrarlo)
    try:
        db.execute("ALTER TABLE puestos ADD COLUMN activo INTEGER NOT NULL DEFAULT 1")
    except sqlite3.OperationalError:
        pass
    try:
        db.execute("ALTER TABLE nomina_detalle ADD COLUMN clasificacion TEXT")
    except sqlite3.OperationalError:
        pass
    # Migracion suave: autorizacion de nomina
    for col, tipo in (("estatus", "TEXT DEFAULT 'pendiente'"),
                      ("autorizada_por", "TEXT"), ("autorizada_en", "TEXT")):
        try:
            db.execute(f"ALTER TABLE nominas ADD COLUMN {col} {tipo}")
        except sqlite3.OperationalError:
            pass
    # Migracion suave: sueldo y viaticos contratados en el detalle de nomina
    for col in ("sueldo_contratado", "viaticos_contratado"):
        try:
            db.execute(f"ALTER TABLE nomina_detalle ADD COLUMN {col} REAL DEFAULT 0")
        except sqlite3.OperationalError:
            pass
    try:
        db.execute("ALTER TABLE nomina_detalle ADD COLUMN bono REAL DEFAULT 0")
    except sqlite3.OperationalError:
        pass
    # Migracion suave: sembrar la tabla de clasificaciones con la lista fija anterior
    if db.execute("SELECT COUNT(*) FROM clasificaciones").fetchone()[0] == 0:
        for nombre in CLASIFICACIONES:
            try:
                db.execute("INSERT INTO clasificaciones(nombre) VALUES(?)", (nombre,))
            except sqlite3.IntegrityError:
                pass
    # Crear usuario administrador la primera vez
    row = db.execute("SELECT COUNT(*) FROM users").fetchone()
    if row[0] == 0:
        pw = os.environ.get("ADMIN_PASSWORD", "cambiar123")
        db.execute(
            "INSERT INTO users(username, password_hash, nombre, role) VALUES(?,?,?,?)",
            ("admin", generate_password_hash(pw), "Administrador", "admin"),
        )
    # Migracion: mapear roles antiguos a los 3 nuevos
    db.execute("UPDATE users SET role='superintendente' "
               "WHERE role IN ('direccion','gerencia','gerente_obra','autorizador')")
    db.execute("UPDATE users SET role='residente' WHERE role='capturista'")
    # Migracion: pasar la obra unica (obra_id) a la tabla user_obras
    try:
        for r in db.execute("SELECT id, obra_id FROM users WHERE obra_id IS NOT NULL").fetchall():
            try:
                db.execute("INSERT INTO user_obras(user_id, obra_id) VALUES(?,?)", (r[0], r[1]))
            except sqlite3.IntegrityError:
                pass
    except sqlite3.OperationalError:
        pass
    # Parametros por defecto (salario minimo 2026 CONASAMI, editable por el admin)
    for clave, valor in (("sm_general", "315.04"),
                         ("sm_zlfn", "440.87"),
                         ("sm_vigencia", "2026-01-01"),
                         ("propuesta_extra", "15"),
                         ("jornada_horas", "8"),
                         ("porcentaje_despacho", "4"),
                         ("email_contador", "acerospoltech@gmail.com"),
                         ("empresa_razon_social", "POLTECH ACERO Y CONSTRUCCION, S.A. DE C.V."),
                         ("empresa_rfc", "PAC2408281N9"),
                         ("empresa_representante", "HUMBERTO FLORES PRADO"),
                         ("empresa_domicilio", "Calle Sur 27, Manzana 27 Lote 254, Col. Leyes de Reforma 1ra Seccion, Iztapalapa, Ciudad de Mexico, C.P. 09310"),
                         ("empresa_instrumento", "52,465"),
                         ("empresa_volumen", "1,086"),
                         ("empresa_fecha_escritura", "28/08/2024"),
                         ("empresa_notario_num", "10"),
                         ("empresa_notario_nombre", "ROBERTO MENDOZA NAVA"),
                         ("empresa_notario_ciudad", "Chalco, Estado de Mexico"),
                         ("apimarket_token", ""),
                         ("apimarket_url_nss", "https://apimarket.mx/api/imss/grupo/localizar-nss"),
                         ("apimarket_url_vigencia", "https://apimarket.mx/api/imss/grupo/consultar-vigencia")):
        db.execute("INSERT OR IGNORE INTO parametros(clave, valor) VALUES(?,?)", (clave, valor))
    # Rellenar datos de empresa en bases ya existentes (sin pisar lo que ya se edito)
    def _param_actual(clave):
        r = db.execute("SELECT valor FROM parametros WHERE clave=?", (clave,)).fetchone()
        return (r[0] if r else "") or ""
    if not _param_actual("empresa_rfc").strip():
        db.execute("INSERT INTO parametros(clave, valor) VALUES('empresa_rfc', ?) "
                   "ON CONFLICT(clave) DO UPDATE SET valor=excluded.valor", ("PAC2408281N9",))
    _dom_old = "Colonia Leyes de Reforma, Alcaldia de Iztapalapa, Ciudad de Mexico"
    _dom_new = ("Calle Sur 27, Manzana 27 Lote 254, Col. Leyes de Reforma 1ra Seccion, "
                "Iztapalapa, Ciudad de Mexico, C.P. 09310")
    if _param_actual("empresa_domicilio").strip() in ("", _dom_old):
        db.execute("INSERT INTO parametros(clave, valor) VALUES('empresa_domicilio', ?) "
                   "ON CONFLICT(clave) DO UPDATE SET valor=excluded.valor", (_dom_new,))
    if not _param_actual("apimarket_url_nss").strip():
        db.execute("INSERT INTO parametros(clave, valor) VALUES('apimarket_url_nss', ?) "
                   "ON CONFLICT(clave) DO UPDATE SET valor=excluded.valor",
                   ("https://apimarket.mx/api/imss/grupo/localizar-nss",))
    if not _param_actual("apimarket_url_vigencia").strip():
        db.execute("INSERT INTO parametros(clave, valor) VALUES('apimarket_url_vigencia', ?) "
                   "ON CONFLICT(clave) DO UPDATE SET valor=excluded.valor",
                   ("https://apimarket.mx/api/imss/grupo/consultar-vigencia",))
    # Normalizacion unica: nombre/apellidos de empleados, y nombres de puesto y obra
    if _param_actual("norm_nombres_hecho") != "1":
        for eid, nom, a1, a2 in db.execute(
                "SELECT id, nombre, primer_apellido, segundo_apellido FROM empleados").fetchall():
            db.execute("UPDATE empleados SET nombre=?, primer_apellido=?, segundo_apellido=? WHERE id=?",
                       (titulo(nom), titulo(a1), titulo(a2), eid))
        for pid, pnom in db.execute("SELECT id, nombre FROM puestos").fetchall():
            db.execute("UPDATE puestos SET nombre=? WHERE id=?", (titulo_obra(pnom), pid))
        for oid, onom in db.execute("SELECT id, nombre FROM obras").fetchall():
            db.execute("UPDATE obras SET nombre=? WHERE id=?", (titulo_obra(onom), oid))
        db.execute("INSERT INTO parametros(clave, valor) VALUES('norm_nombres_hecho','1') "
                   "ON CONFLICT(clave) DO UPDATE SET valor='1'")
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
        "ADMIN_RANK": ADMIN_RANK,
        "APP_VERSION": APP_VERSION,
    }


@app.template_filter("money")
def money(v):
    """Formatea numeros como $1,234.56 (coma en miles, punto en centavos)."""
    try:
        return "${:,.2f}".format(float(v or 0))
    except (ValueError, TypeError):
        return v


# Estados del proceso de documentos de cada trabajador
ESTATUS_DOCS = ["Pendiente de carga", "Informacion incompleta",
                "Carga completa", "Validado"]
DIAS_RETENCION = 15  # dias para validar la documentacion antes de retener pago

# Documentos que se suben escaneados por trabajador
TIPOS_DOC = [
    ("INE", "INE / Identificacion oficial (PDF)"),
    ("Comprobante", "Comprobante de domicilio (PDF)"),
    ("Contrato", "Contrato firmado (PDF)"),
    ("Fiscal", "Constancia de Situacion Fiscal / RFC (PDF)"),
]
DOCS_REQUERIDOS = ["INE", "Comprobante", "Contrato"]  # necesarios para "Carga completa"


def recomputar_estatus_docs(db, emp_id):
    """Actualiza solo el estatus de documentos segun lo que ya se subio."""
    tipos = [r["tipo"] for r in db.execute(
        "SELECT tipo FROM documentos WHERE empleado_id=?", (emp_id,)).fetchall()]
    presentes = [t for t in DOCS_REQUERIDOS if t in tipos]
    fila = db.execute("SELECT estatus_docs FROM empleados WHERE id=?", (emp_id,)).fetchone()
    actual = fila["estatus_docs"] if fila else None
    if len(presentes) == 0:
        nuevo = "Pendiente de carga"
    elif len(presentes) < len(DOCS_REQUERIDOS):
        nuevo = "Informacion incompleta"
    else:
        nuevo = "Validado" if actual == "Validado" else "Carga completa"
    db.execute("UPDATE empleados SET estatus_docs=? WHERE id=?", (nuevo, emp_id))


def validar_solicitud_imss(fecha_solicitud, fecha_alta):
    """La solicitud de alta IMSS debe ser <= fecha de alta y como maximo 5 dias antes."""
    if not fecha_solicitud or not fecha_alta:
        return None
    try:
        fs = date.fromisoformat(fecha_solicitud[:10])
        fa = date.fromisoformat(fecha_alta[:10])
    except ValueError:
        return None
    if fs > fa:
        return "La fecha de solicitud de alta IMSS no puede ser posterior a la fecha de alta."
    if (fa - fs).days > 5:
        return "La fecha de solicitud de alta IMSS debe ser como maximo 5 dias antes de la fecha de alta."
    return None


def dias_para_retencion(emp):
    """Dias que faltan para que se retenga la nomina. Se ancla a la fecha de alta
    (en la Fase 2 se re-anclara a la primera nomina). Devuelve (texto, clase_color)."""
    if (emp["estatus_docs"] or "") == "Validado":
        return ("Validado", "success")
    base = emp["fecha_alta"]
    if not base:
        return ("-", "secondary")
    try:
        d0 = date.fromisoformat(str(base)[:10])
    except ValueError:
        return ("-", "secondary")
    restan = DIAS_RETENCION - (date.today() - d0).days
    if restan <= 0:
        return ("Retenido", "danger")
    clase = "warning" if restan <= 5 else "secondary"
    return (f"{restan} dias", clase)

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

def clabe_valida(num):
    """CLABE interbancaria: 18 digitos con digito verificador (pesos 3,7,1)."""
    n = solo_digitos(num)
    if len(n) != 18:
        return False
    pesos = [3, 7, 1]
    suma = sum((int(n[i]) * pesos[i % 3]) % 10 for i in range(17))
    control = (10 - (suma % 10)) % 10
    return control == int(n[17])

def validar_cuenta(tipo, numero):
    """Candado de formato para la cuenta bancaria. Devuelve un mensaje de error o None."""
    n = solo_digitos(numero)
    if not n:
        return None  # cuenta opcional; si no hay numero, no se valida aqui
    t = (tipo or "").lower()
    if "clabe" in t:
        if len(n) != 18:
            return f"La CLABE debe tener 18 digitos (tiene {len(n)}). Revisa si falta o sobra algun numero."
        if not clabe_valida(n):
            return "La CLABE no pasa la validacion del digito verificador. Revisa que no haya un numero mal."
    elif "tarjeta" in t:
        if len(n) not in (15, 16):
            return f"El numero de tarjeta debe tener 16 digitos (tiene {len(n)})."
        if not luhn_ok(n):
            return "El numero de tarjeta no pasa la validacion (digito verificador). Revisa los numeros."
    else:  # numero de cuenta
        if not (8 <= len(n) <= 20):
            return f"El numero de cuenta debe tener entre 8 y 20 digitos (tiene {len(n)})."
    return None

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
# Fase 2: asistencia semanal (viernes a jueves, base de 6 dias)
# ---------------------------------------------------------------------------
# d1..d7 = de viernes a jueves. El domingo (d3) es visible pero NO cuenta
# para la base de 6 dias.
DIAS_SEMANA = [("d1", "VIE"), ("d2", "SAB"), ("d3", "DOM"), ("d4", "LUN"),
               ("d5", "MAR"), ("d6", "MIE"), ("d7", "JUE")]
DIAS_BASE = ["d1", "d2", "d4", "d5", "d6", "d7"]   # 6 dias, sin domingo
CODIGOS_ASIS = [
    ("A", "Asistio"),
    ("F", "Falta"),
    ("R", "Retardo"),
    ("D", "Descanso"),
]

def viernes_de(fecha):
    """Viernes que inicia la semana de nomina (viernes-jueves) que contiene 'fecha'."""
    offset = (fecha.weekday() - 4) % 7   # lunes=0 ... viernes=4
    return fecha - timedelta(days=offset)

def fechas_de_dias(viernes):
    """Etiqueta corta dd/mm para cada dia de la semana."""
    return {f"d{i+1}": (viernes + timedelta(days=i)).strftime("%d/%m") for i in range(7)}

def contar_asistencia(cods):
    """cods: dict d1..d7 con codigos. Cuenta solo los 6 dias base (domingo no).
    Devuelve (dias_trabajados, faltas, retardos). Un retardo cuenta como dia trabajado."""
    dias = faltas = retardos = 0
    for k in DIAS_BASE:
        c = (cods.get(k) or "").upper()
        if c in ("A", "R"):
            dias += 1
        if c == "F":
            faltas += 1
        if c == "R":
            retardos += 1
    return dias, faltas, retardos

def obras_visibles(db):
    """Lista de obras que el usuario puede ver (respeta rol/asignacion)."""
    vis = obras_del_usuario(db)
    if vis is None:
        return db.execute("SELECT * FROM obras ORDER BY nombre").fetchall()
    if vis:
        ph = ",".join("?" * len(vis))
        return db.execute(f"SELECT * FROM obras WHERE id IN ({ph}) ORDER BY nombre", vis).fetchall()
    return []

def puede_ver_obra(db, obra_id):
    vis = obras_del_usuario(db)
    return vis is None or obra_id in (vis or [])


# --- Parametros del sistema (salario minimo, etc.) ---
def get_param(db, clave, default=None):
    row = db.execute("SELECT valor FROM parametros WHERE clave=?", (clave,)).fetchone()
    return row["valor"] if row else default

def set_param(db, clave, valor):
    db.execute("INSERT INTO parametros(clave, valor) VALUES(?,?) "
               "ON CONFLICT(clave) DO UPDATE SET valor=excluded.valor", (clave, str(valor)))

def registrar_bitacora(db, accion, detalle=""):
    """Registra un movimiento en la bitacora (quien, cuando, que)."""
    db.execute(
        "INSERT INTO bitacora(fecha, usuario, rol, accion, detalle) VALUES(?,?,?,?,?)",
        (datetime.now().isoformat(timespec="seconds"),
         session.get("nombre", ""), session.get("role", ""), accion, detalle))


# ---------------------------------------------------------------------------
# Contratos: numero a letra y generacion del docx
# ---------------------------------------------------------------------------
MESES_ES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio",
            "agosto", "septiembre", "octubre", "noviembre", "diciembre"]

def _centenas_letra(n):
    UNIDADES = ["", "UNO", "DOS", "TRES", "CUATRO", "CINCO", "SEIS", "SIETE", "OCHO", "NUEVE",
                "DIEZ", "ONCE", "DOCE", "TRECE", "CATORCE", "QUINCE", "DIECISEIS",
                "DIECISIETE", "DIECIOCHO", "DIECINUEVE", "VEINTE"]
    DECENAS = ["", "", "VEINTI", "TREINTA", "CUARENTA", "CINCUENTA", "SESENTA",
               "SETENTA", "OCHENTA", "NOVENTA"]
    CENTENAS = ["", "CIENTO", "DOSCIENTOS", "TRESCIENTOS", "CUATROCIENTOS", "QUINIENTOS",
                "SEISCIENTOS", "SETECIENTOS", "OCHOCIENTOS", "NOVECIENTOS"]
    if n == 0:
        return ""
    if n == 100:
        return "CIEN"
    c, resto = divmod(n, 100)
    out = CENTENAS[c]
    if resto:
        if resto <= 20:
            out = (out + " " + UNIDADES[resto]).strip()
        else:
            d, u = divmod(resto, 10)
            if d == 2:
                out = (out + " VEINTI" + UNIDADES[u].lower()).strip() if u else (out + " VEINTE").strip()
            else:
                seg = DECENAS[d] + (" Y " + UNIDADES[u] if u else "")
                out = (out + " " + seg).strip()
    return out.strip()

def numero_letra(n):
    """Entero a letras en mayusculas (hasta millones)."""
    n = int(n)
    if n == 0:
        return "CERO"
    partes = []
    millones, resto = divmod(n, 1_000_000)
    miles, cientos = divmod(resto, 1000)
    if millones:
        partes.append("UN MILLON" if millones == 1 else _centenas_letra(millones) + " MILLONES")
    if miles:
        partes.append("MIL" if miles == 1 else _centenas_letra(miles) + " MIL")
    if cientos:
        partes.append(_centenas_letra(cientos))
    return " ".join(p for p in partes if p).strip()

def pesos_letra(monto):
    """Ej. 330.04 -> 'TRESCIENTOS TREINTA PESOS 04/100 M.N.'"""
    try:
        monto = float(monto)
    except (TypeError, ValueError):
        monto = 0.0
    entero = int(monto)
    centavos = int(round((monto - entero) * 100))
    return f"{numero_letra(entero)} PESOS {centavos:02d}/100 M.N."

def edad_de(fecha_nac):
    try:
        f = date.fromisoformat(str(fecha_nac)[:10])
    except (ValueError, TypeError):
        return ""
    hoy = date.today()
    return hoy.year - f.year - ((hoy.month, hoy.day) < (f.month, f.day))

def generar_contrato_docx(db, emp):
    """Arma el contrato individual de trabajo (docx) con los datos del empleado
    y de la empresa. Devuelve (bytes, nombre_archivo). Los datos que el sistema
    no captura quedan como linea para llenar a mano."""
    from docx import Document
    from docx.shared import Pt
    from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_COLOR_INDEX

    P = lambda k, d="": (get_param(db, k, d) or d)
    BL = "________________________"      # linea para completar a mano

    razon = P("empresa_razon_social")
    rep = P("empresa_representante")
    dom = P("empresa_domicilio")
    rfc_emp = P("empresa_rfc") or BL
    instr = P("empresa_instrumento")
    vol = P("empresa_volumen")
    fesc = P("empresa_fecha_escritura")
    not_num = P("empresa_notario_num")
    not_nom = P("empresa_notario_nombre")
    not_ciu = P("empresa_notario_ciudad")

    nombre = titulo(" ".join(x for x in [emp["nombre"], emp["primer_apellido"], emp["segundo_apellido"]] if x))
    puesto = titulo_obra(emp["puesto"]) if emp["puesto"] else BL
    estado_obra = emp["estado"] or ""
    salario = _num(emp["importe_alta_imss"])
    sexo_txt = "MASCULINO" if (emp["sexo"] or "").upper().startswith("H") else ("FEMENINO" if (emp["sexo"] or "").upper().startswith("M") else BL)
    edad = edad_de(emp["fecha_nacimiento"])
    hoy = date.today()
    fecha_txt = f"{hoy.day} de {MESES_ES[hoy.month-1]} de {hoy.year}"

    doc = Document()
    base = doc.styles["Normal"]; base.font.name = "Arial"; base.font.size = Pt(10)

    def par(texto="", *, bold=False, center=False, size=None, space=6):
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(space)
        if center: p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        else: p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        r = p.add_run(texto)
        r.bold = bold
        if size: r.font.size = Pt(size)
        return p

    LINEA = "  ____________________________  "   # espacio para escribir a mano

    def par_seg(segmentos, *, center=False, space=6):
        """Arma un parrafo con segmentos; un segmento (texto, True) va resaltado en amarillo."""
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(space)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER if center else WD_ALIGN_PARAGRAPH.JUSTIFY
        for seg in segmentos:
            if isinstance(seg, tuple):
                r = p.add_run(seg[0]); r.font.highlight_color = WD_COLOR_INDEX.YELLOW
            else:
                p.add_run(seg)
        return p

    def blank():
        """Espacio resaltado para llenar a mano."""
        return (LINEA, True)

    def dato(v):
        """Dato normal; si falta, queda como espacio resaltado para llenar a mano."""
        return v if (v is not None and str(v).strip()) else (LINEA, True)

    par("CONTRATO INDIVIDUAL DE TRABAJO POR TIEMPO INDETERMINADO", bold=True, center=True, size=12)
    par(f"CONTRATO INDIVIDUAL DE TRABAJO POR TIEMPO INDETERMINADO QUE CELEBRAN POR UNA PARTE "
        f"LA EMPRESA {razon}, REPRESENTADA EN ESTE ACTO POR EL (LA) C. {rep} A QUIEN EN LO "
        f"SUCESIVO SE LE DENOMINARA \u201cEL PATRON\u201d, Y POR LA OTRA PARTE, EL (LA) C. {nombre} "
        f"A QUIEN EN LO SUCESIVO SE LE DENOMINARA COMO \u201cEL TRABAJADOR\u201d, Y A QUIENES EN SU "
        f"CONJUNTO SE LES DENOMINARA \u201cLAS PARTES\u201d; QUE CONTIENE LAS CONDICIONES GENERALES "
        f"DE TRABAJO BAJO LAS SIGUIENTES:")
    par("D E C L A R A C I O N E S:", bold=True, center=True)
    par_seg([("Nota: los espacios resaltados en amarillo deben llenarse (a mano o a maquina) "
              "y ser verificados ANTES de que las partes firmen el contrato.", True)], space=10)
    par(f"PRIMERA.- Para los efectos de los articulos 10, 11, 16, 24 y 25 de la Ley Federal del "
        f"Trabajo, el C. {rep} declara que su representada es una sociedad mexicana, con domicilio "
        f"ubicado en {dom}, con Registro Federal de Contribuyentes {rfc_emp}, constituida bajo el "
        f"Instrumento Notarial numero {instr}, Volumen {vol}, de fecha {fesc}, levantado ante la fe "
        f"del Notario Publico numero {not_num} Lic. {not_nom}, con ejercicio en {not_ciu}.")
    par_seg([
        "SEGUNDA.- \u201cEL TRABAJADOR\u201d en terminos de los articulos 8, 24 y 25 de la Ley Federal "
        "del Trabajo declara llamarse ", dato(nombre),
        ", con numero de seguridad social ", dato(emp["nss"]),
        ", Clave Unica de Registro de Poblacion ", dato(emp["curp"]),
        ", Registro Federal de Contribuyentes ", dato(emp["rfc"]),
        ", sexo ", dato(None if sexo_txt == BL else sexo_txt),
        ", escolaridad ", blank(),
        ", nacido en ", blank(),
        " que cuenta con ", dato(str(edad) if edad != "" else None),
        " anos de edad, de nacionalidad Mexicana, estado civil ", blank(),
        " y con ultimo domicilio en ", blank(), ".",
    ])
    par("TERCERA.- \u201cLAS PARTES\u201d se reconocen expresamente la personalidad juridica con la que "
        "se ostentan para todos los efectos legales a que haya lugar.")
    par("\u201cLAS PARTES\u201d acuerdan sujetarse al tenor de las siguientes:")
    par("C L A U S U L A S:", bold=True, center=True)
    par(f"PRIMERA.- \u201cEL TRABAJADOR\u201d se obliga a prestar sus servicios personales a "
        f"\u201cEL PATRON\u201d, subordinandose juridicamente para ocupar el puesto de {puesto}, "
        f"conviniendo que este trabajo debera ejecutarlo con cuidado, esmero, eficiencia y en la forma, "
        f"tiempo y lugar convenido, acatando el Reglamento Interior de Trabajo y las disposiciones que "
        f"dicte \u201cEL PATRON\u201d, sobremanera lo senalado por el articulo 27 de la Ley Federal del "
        f"Trabajo; las actividades que ejecutara son las propias del puesto de {puesto} y las que se "
        f"relacionen directa e indirectamente, de manera enunciativa y no limitativa.")
    par("SEGUNDA.- \u201cEL TRABAJADOR\u201d debera ejecutar su trabajo en las oficinas, "
        "establecimientos, talleres, bodegas y en general en cualquier lugar donde \u201cEL PATRON\u201d "
        "ordene desempenar las actividades, solo las que correspondan con su puesto y demas relacionadas.")
    par("TERCERA.- Este contrato se celebra por tiempo indeterminado, conforme al articulo 35 de la Ley "
        "Federal del Trabajo.")
    par(f"CUARTA.- Por los servicios contratados, \u201cEL PATRON\u201d pagara a \u201cEL TRABAJADOR\u201d "
        f"un salario diario de ${salario:,.2f} (*** {pesos_letra(salario)} ***); el cual bajo ninguna "
        f"circunstancia sera inferior al salario minimo del area geografica donde preste sus servicios. "
        f"El salario se fijara de manera semanal, en moneda de curso legal, y podra pagarse en efectivo, "
        f"deposito en cuenta bancaria, tarjeta de debito, transferencias o cualquier otro medio "
        f"electronico, lo cual \u201cEL TRABAJADOR\u201d autoriza a la firma del presente contrato.")
    par("\u201cEL TRABAJADOR\u201d autoriza a \u201cEL PATRON\u201d para que deduzca de su salario los "
        "impuestos a su cargo, las cuotas obreras al IMSS y cualquier otra cantidad conforme al articulo "
        "110 de la Ley Federal del Trabajo.")
    par_seg(["QUINTA.- \u201cEL PATRON\u201d entregara los recibos de nomina de \u201cEL TRABAJADOR\u201d "
             "al correo electronico ", blank(),
             ", conforme al articulo 101 de la Ley Federal del Trabajo."])
    par("SEXTA.- La duracion de la jornada sera de 48 horas a la semana, pudiendo las partes repartir las "
        "horas de trabajo conforme a las necesidades del centro de trabajo, en terminos del articulo 59 "
        "de la Ley Federal del Trabajo.")
    par("SEPTIMA.- El tiempo extraordinario se pagara conforme a la Ley Federal del Trabajo y solo se "
        "laborara previa solicitud expresa y por escrito de \u201cEL PATRON\u201d y aceptacion por escrito "
        "de \u201cEL TRABAJADOR\u201d, sin exceder de tres horas diarias ni tres veces por semana.")
    par("OCTAVA.- \u201cEL TRABAJADOR\u201d esta obligado a registrar su asistencia a la entrada y salida "
        "de sus labores en la forma que establezca \u201cEL PATRON\u201d.")
    par("NOVENA.- \u201cEL TRABAJADOR\u201d tendra derecho a un dia de descanso semanal con goce de salario "
        "integro, preferentemente el domingo, segun las condiciones del centro de trabajo.")
    par("DECIMA.- Despues de un ano de servicios \u201cEL TRABAJADOR\u201d disfrutara del periodo anual de "
        "vacaciones pagadas y prima vacacional conforme a la Ley Federal del Trabajo.")
    par("DECIMA PRIMERA.- \u201cEL TRABAJADOR\u201d percibira un aguinaldo anual, pagadero antes del 20 de "
        "diciembre, equivalente a 15 dias de salario, conforme al articulo 87 de la Ley Federal del Trabajo.")
    par("DECIMA SEGUNDA.- \u201cEL TRABAJADOR\u201d conviene en someterse a los reconocimientos medicos que "
        "ordene \u201cEL PATRON\u201d en terminos del articulo 134 fraccion X de la Ley Federal del Trabajo.")
    par("DECIMA TERCERA.- \u201cEL TRABAJADOR\u201d sera capacitado y adiestrado conforme a los planes y "
        "programas que establezca \u201cEL PATRON\u201d, comprometiendose a participar en ellos.")
    par_seg(["DECIMA CUARTA.- \u201cEL TRABAJADOR\u201d, en terminos del articulo 25 fraccion X y del "
             "diverso 501 de la Ley Federal del Trabajo, designa como beneficiario(s) a ", blank(), "."])
    par("DECIMA QUINTA.- \u201cEL TRABAJADOR\u201d se obliga a guardar discrecion y estricta "
        "confidencialidad de la informacion, datos o documentos confidenciales y reservados de "
        "\u201cEL PATRON\u201d, durante la vigencia del contrato y hasta por cinco anos posteriores.")
    par(f"DECIMA SEXTA.- El presente contrato anula cualquier otro contrato o convenio anterior entre "
        f"\u201cLAS PARTES\u201d. Para todo lo relativo a su interpretacion, cumplimiento y ejecucion, "
        f"\u201cLAS PARTES\u201d se someten a la jurisdiccion de los Tribunales del estado de "
        f"{estado_obra or BL}, renunciando a cualquier otra.")
    par(f"Leido que fue el presente contrato por \u201cLAS PARTES\u201d, e impuestas de su contenido y "
        f"fuerza legal, lo firmaron, quedando un tanto en poder de cada una. {estado_obra}, a {fecha_txt}.",
        space=24)
    par("_________________________________          _________________________________", center=True, space=2)
    par(f"C. {rep}                              C. {nombre}", center=True, space=2)
    par(f"En representacion legal de: {razon}", center=True, space=12)
    par("\u201cEL PATRON\u201d                                        \u201cEL TRABAJADOR\u201d",
        bold=True, center=True)

    buf = io.BytesIO(); doc.save(buf); buf.seek(0)
    slug = re.sub(r"[^A-Za-z0-9]+", "_", nombre).strip("_")
    return buf.getvalue(), f"Contrato_{slug or 'trabajador'}.docx"

def generar_baja_docx(db, emp):
    """Escrito/aviso de baja de trabajador (docx) para tramite ante el IMSS."""
    from docx import Document
    from docx.shared import Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_COLOR_INDEX

    P = lambda k, d="": (get_param(db, k, d) or d)
    razon = P("empresa_razon_social"); rfc_emp = P("empresa_rfc")
    dom = P("empresa_domicilio"); rep = P("empresa_representante")

    nombre = titulo(" ".join(x for x in [emp["nombre"], emp["primer_apellido"], emp["segundo_apellido"]] if x))
    puesto = titulo_obra(emp["puesto"]) if emp["puesto"] else ""
    obra = emp["obra"] or ""
    fbaja = emp["fecha_baja"] or ""
    motivo = emp["motivo_baja"] or ""
    hoy = date.today()
    fecha_txt = f"{hoy.day} de {MESES_ES[hoy.month-1]} de {hoy.year}"

    doc = Document()
    doc.styles["Normal"].font.name = "Arial"; doc.styles["Normal"].font.size = Pt(11)

    def par(txt="", *, bold=False, center=False, size=None, color=None, space=6, just=False):
        p = doc.add_paragraph(); p.paragraph_format.space_after = Pt(space)
        p.alignment = (WD_ALIGN_PARAGRAPH.CENTER if center else
                       WD_ALIGN_PARAGRAPH.JUSTIFY if just else WD_ALIGN_PARAGRAPH.LEFT)
        r = p.add_run(txt); r.bold = bold
        if size: r.font.size = Pt(size)
        if color: r.font.color.rgb = RGBColor(*color)
        return p

    def par_hl(pre, valor, post=".", **kw):
        """Parrafo con un dato; si el dato falta, queda resaltado para llenar a mano."""
        p = doc.add_paragraph(); p.paragraph_format.space_after = Pt(kw.get("space", 6))
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.add_run(pre)
        if valor and str(valor).strip():
            r = p.add_run(str(valor)); r.bold = True
        else:
            r = p.add_run("  ____________________  "); r.font.highlight_color = WD_COLOR_INDEX.YELLOW
        p.add_run(post)
        return p

    if razon:
        par(razon, bold=True, center=True, size=13, color=(0x16, 0x23, 0x3C), space=2)
    if rfc_emp or dom:
        par(f"RFC: {rfc_emp}   {dom}", center=True, size=9, color=(0x55, 0x55, 0x55), space=12)
    par("AVISO DE BAJA DE TRABAJADOR", bold=True, center=True, size=13, color=(0xE1, 0x28, 0x1A), space=14)
    par(f"Ciudad de México, a {fecha_txt}.", space=14)

    par("A QUIEN CORRESPONDA / INSTITUTO MEXICANO DEL SEGURO SOCIAL:", bold=True, space=10)
    par(f"Por medio del presente, la empresa {razon} hace constar la BAJA del trabajador "
        f"que a continuación se detalla, para los efectos administrativos y de seguridad "
        f"social correspondientes:", just=True, space=12)

    par_hl("Nombre del trabajador: ", nombre)
    par_hl("Número de Seguridad Social (NSS): ", emp["nss"])
    par_hl("CURP: ", emp["curp"])
    par_hl("RFC: ", emp["rfc"])
    par_hl("Puesto: ", puesto)
    par_hl("Obra / centro de trabajo: ", obra)
    par_hl("Fecha de ingreso: ", emp["fecha_alta"])
    par_hl("Fecha de baja (último día laborado): ", fbaja)
    par_hl("Motivo de la baja: ", motivo, space=14)

    par("Se solicita realizar los trámites correspondientes ante el IMSS y demás "
        "instancias aplicables. Sin otro particular, quedo a sus órdenes.", just=True, space=28)

    par("_________________________________________", center=True, space=2)
    par_hl2 = par(f"{rep}", center=True, bold=True, space=0) if rep else par("____________________", center=True, space=0)
    par(f"En representación de {razon}", center=True, size=10, space=0)

    buf = io.BytesIO(); doc.save(buf); buf.seek(0)
    slug = re.sub(r"[^A-Za-z0-9]+", "_", nombre).strip("_")
    return buf.getvalue(), f"Aviso_baja_{slug or emp['cedula']}.docx"


def crear_contrato_para(db, empleado_id, avisar_a=None):
    """Genera y guarda el contrato de un empleado. Si 'avisar_a' es un correo,
    lo envia como adjunto. Devuelve el id del contrato o None."""
    emp = db.execute(
        "SELECT e.*, p.nombre AS puesto, o.estado, o.nombre AS obra "
        "FROM empleados e JOIN puestos p ON p.id=e.puesto_id "
        "JOIN obras o ON o.id=p.obra_id WHERE e.id=?", (empleado_id,)).fetchone()
    if not emp:
        return None
    datos, archivo = generar_contrato_docx(db, emp)
    cur = db.execute(
        "INSERT INTO contratos(empleado_id, nombre_archivo, generado_en, generado_por, "
        "estatus, contenido) VALUES(?,?,?,?, 'generado', ?)",
        (empleado_id, archivo, datetime.now().isoformat(timespec="seconds"),
         session.get("nombre", ""), datos))
    registrar_bitacora(db, "Contrato generado",
                       f"{emp['nombre']} {emp['primer_apellido']} (cedula {emp['cedula']})")
    if avisar_a:
        enviar_correo(
            f"POLTECH - Contrato generado: {emp['primer_apellido']} {emp['nombre']}",
            (f"Se genero el contrato de {emp['nombre']} {emp['primer_apellido']} "
             f"{emp['segundo_apellido'] or ''} (cedula {emp['cedula']}, puesto {emp['puesto']}).\n"
             f"Se adjunta para su revision. Los datos que el sistema no captura quedan como linea "
             f"para completar a mano."),
            [avisar_a],
            adjuntos=[(archivo, datos,
                       "application/vnd.openxmlformats-officedocument.wordprocessingml.document")])
    return cur.lastrowid

# Estados dentro de la Zona Libre de la Frontera Norte (para proponer el SM correcto)
ESTADOS_ZLFN = {"Baja California", "Sonora", "Chihuahua", "Coahuila",
                "Nuevo Leon", "Tamaulipas"}

def propuesta_salario_alta(db, estado=None):
    """Propone el salario diario de alta IMSS: salario minimo del dia + extra.
    Usa el SM de la Zona Libre de la Frontera Norte si el estado pertenece a ella."""
    try:
        extra = float(get_param(db, "propuesta_extra", "15") or 15)
    except (ValueError, TypeError):
        extra = 15.0
    clave = "sm_zlfn" if (estado in ESTADOS_ZLFN) else "sm_general"
    try:
        sm = float(get_param(db, clave, "0") or 0)
    except (ValueError, TypeError):
        sm = 0.0
    return round(sm + extra, 2)


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
            session["username"] = row["username"]
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
    vis = obras_del_usuario(db)
    if vis is None:
        stats = {
            "empleados": db.execute("SELECT COUNT(*) FROM empleados WHERE estatus='activo'").fetchone()[0],
            "obras": db.execute("SELECT COUNT(*) FROM obras").fetchone()[0],
            "puestos": db.execute("SELECT COUNT(*) FROM puestos").fetchone()[0],
            "cuentas": db.execute("SELECT COUNT(*) FROM cuentas_bancarias").fetchone()[0],
        }
    elif vis:
        ph = ",".join("?" * len(vis))
        stats = {
            "empleados": db.execute(f"SELECT COUNT(*) FROM empleados e JOIN puestos p ON p.id=e.puesto_id "
                                    f"WHERE e.estatus='activo' AND p.obra_id IN ({ph})", vis).fetchone()[0],
            "obras": len(vis),
            "puestos": db.execute(f"SELECT COUNT(*) FROM puestos WHERE obra_id IN ({ph})", vis).fetchone()[0],
            "cuentas": db.execute(f"SELECT COUNT(*) FROM cuentas_bancarias c JOIN empleados e ON e.id=c.empleado_id "
                                  f"JOIN puestos p ON p.id=e.puesto_id WHERE p.obra_id IN ({ph})", vis).fetchone()[0],
        }
    else:
        stats = {"empleados": 0, "obras": 0, "puestos": 0, "cuentas": 0}
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
    vis = obras_del_usuario(db)
    if vis is None:
        filas = db.execute("SELECT * FROM obras ORDER BY nombre").fetchall()
    elif vis:
        ph = ",".join("?" * len(vis))
        filas = db.execute(f"SELECT * FROM obras WHERE id IN ({ph}) ORDER BY nombre", vis).fetchall()
    else:
        filas = []
    return render_template("obras_list.html", obras=filas, estados=ESTADOS)


@app.route("/obra/<int:obra_id>/editar", methods=["GET", "POST"])
@min_rank(GERENTE_RANK)
def obra_editar(obra_id):
    db = get_db()
    obra = db.execute("SELECT * FROM obras WHERE id=?", (obra_id,)).fetchone()
    if not obra:
        abort(404)
    if not puede_ver_obra(db, obra_id):
        abort(403)
    if request.method == "POST":
        db.execute(
            "UPDATE obras SET proyecto=?, contrato=?, nombre=?, estado=? WHERE id=?",
            (request.form.get("proyecto", "").strip(),
             request.form.get("contrato", "").strip(),
             request.form.get("nombre", "").strip() or obra["nombre"],
             request.form.get("estado", ""), obra_id))
        db.commit()
        registrar_bitacora(db, "Edicion de obra", f"Obra {obra_id}")
        db.commit()
        flash("Obra actualizada.", "success")
        return redirect(url_for("obras"))
    return render_template("obra_form.html", obra=obra, estados=ESTADOS)

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
            "INSERT INTO puestos(obra_id, nombre, sueldo_semanal, viaticos_semanales, clasificacion) "
            "VALUES(?,?,?,?,?)",
            (request.form.get("obra_id"),
             titulo(request.form.get("nombre", "").strip()),
             float(request.form.get("sueldo_semanal") or 0),
             float(request.form.get("viaticos_semanales") or 0),
             request.form.get("clasificacion", "").strip() or None),
        )
        db.commit()
        flash("Puesto agregado al catalogo.", "success")
        return redirect(url_for("catalogo"))
    vis = obras_del_usuario(db)
    filtro_obra = (request.args.get("obra_id") or "").strip()
    filtro_clasif = (request.args.get("clasificacion") or "").strip()
    where, args = [], []
    if vis is None:
        obras_l = db.execute("SELECT * FROM obras ORDER BY nombre").fetchall()
    elif vis:
        ph = ",".join("?" * len(vis))
        obras_l = db.execute(f"SELECT * FROM obras WHERE id IN ({ph}) ORDER BY nombre", vis).fetchall()
        where.append(f"o.id IN ({ph})"); args += vis
    else:
        obras_l = []
        where.append("0")
    if filtro_obra:
        where.append("o.id = ?"); args.append(filtro_obra)
    if filtro_clasif:
        where.append("p.clasificacion = ?"); args.append(filtro_clasif)
    sql = ("SELECT p.*, p.nombre AS puesto, o.nombre AS obra, o.estado FROM puestos p "
           "JOIN obras o ON o.id=p.obra_id")
    if where:
        sql += " WHERE " + " AND ".join(where)
    sql += " ORDER BY o.nombre, p.nombre"
    puestos = db.execute(sql, args).fetchall()
    clasif_rows = db.execute(
        "SELECT c.id, c.nombre, "
        "(SELECT COUNT(*) FROM puestos p WHERE p.clasificacion = c.nombre) AS usos "
        "FROM clasificaciones c ORDER BY c.nombre"
    ).fetchall()
    return render_template("catalogo_list.html", obras=obras_l, puestos=puestos,
                           filtro_obra=filtro_obra, filtro_clasif=filtro_clasif,
                           clasificaciones=[c["nombre"] for c in clasif_rows],
                           clasif_rows=clasif_rows)


@app.route("/clasificacion/agregar", methods=["POST"])
@min_rank(GERENTE_RANK)
def clasificacion_agregar():
    db = get_db()
    nombre = titulo(request.form.get("nombre", "").strip())
    if not nombre:
        flash("Escribe un nombre para la clasificacion.", "warning")
    else:
        try:
            db.execute("INSERT INTO clasificaciones(nombre) VALUES(?)", (nombre,))
            db.commit()
            flash(f"Clasificacion '{nombre}' agregada.", "success")
        except sqlite3.IntegrityError:
            flash(f"La clasificacion '{nombre}' ya existe.", "warning")
    return redirect(url_for("catalogo"))


@app.route("/clasificacion/<int:clasificacion_id>/eliminar", methods=["POST"])
@min_rank(GERENTE_RANK)
def clasificacion_eliminar(clasificacion_id):
    db = get_db()
    c = db.execute("SELECT * FROM clasificaciones WHERE id=?", (clasificacion_id,)).fetchone()
    if not c:
        abort(404)
    usos = db.execute("SELECT COUNT(*) FROM puestos WHERE clasificacion=?", (c["nombre"],)).fetchone()[0]
    if usos > 0:
        flash(f"No se puede eliminar '{c['nombre']}': esta asignada a {usos} puesto(s). "
              "Cambia esos puestos a otra clasificacion primero.", "danger")
    else:
        db.execute("DELETE FROM clasificaciones WHERE id=?", (clasificacion_id,))
        db.commit()
        flash(f"Clasificacion '{c['nombre']}' eliminada.", "success")
    return redirect(url_for("catalogo"))


@app.route("/puesto/<int:puesto_id>/editar", methods=["GET", "POST"])
@min_rank(GERENTE_RANK)
def puesto_editar(puesto_id):
    db = get_db()
    p = db.execute("SELECT p.*, o.nombre AS obra FROM puestos p JOIN obras o ON o.id=p.obra_id "
                   "WHERE p.id=?", (puesto_id,)).fetchone()
    if not p:
        abort(404)
    if not puede_ver_obra(db, p["obra_id"]):
        abort(403)
    if request.method == "POST":
        db.execute(
            "UPDATE puestos SET nombre=?, sueldo_semanal=?, viaticos_semanales=?, clasificacion=? "
            "WHERE id=?",
            (titulo(request.form.get("nombre", "").strip()),
             float(request.form.get("sueldo_semanal") or 0),
             float(request.form.get("viaticos_semanales") or 0),
             request.form.get("clasificacion", "").strip() or None, puesto_id))
        db.commit()
        flash("Puesto actualizado.", "success")
        return redirect(url_for("catalogo"))
    clasificaciones = [r["nombre"] for r in
                       db.execute("SELECT nombre FROM clasificaciones ORDER BY nombre").fetchall()]
    return render_template("puesto_form.html", p=p, clasificaciones=clasificaciones)


@app.route("/puesto/<int:puesto_id>/desactivar", methods=["POST"])
@min_rank(GERENTE_RANK)
def puesto_desactivar(puesto_id):
    db = get_db()
    p = db.execute("SELECT * FROM puestos WHERE id=?", (puesto_id,)).fetchone()
    if not p:
        abort(404)
    if not puede_ver_obra(db, p["obra_id"]):
        abort(403)
    db.execute("UPDATE puestos SET activo=0 WHERE id=?", (puesto_id,))
    db.commit()
    flash(f"Puesto '{p['nombre']}' desactivado. Ya no aparecera para dar de alta gente nueva, "
          "pero los trabajadores que ya lo tienen siguen funcionando normal.", "success")
    return redirect(url_for("catalogo"))


@app.route("/puesto/<int:puesto_id>/reactivar", methods=["POST"])
@min_rank(GERENTE_RANK)
def puesto_reactivar(puesto_id):
    db = get_db()
    p = db.execute("SELECT * FROM puestos WHERE id=?", (puesto_id,)).fetchone()
    if not p:
        abort(404)
    if not puede_ver_obra(db, p["obra_id"]):
        abort(403)
    db.execute("UPDATE puestos SET activo=1 WHERE id=?", (puesto_id,))
    db.commit()
    flash(f"Puesto '{p['nombre']}' reactivado.", "success")
    return redirect(url_for("catalogo"))

# ---------------------------------------------------------------------------
# Personal
# ---------------------------------------------------------------------------
@app.route("/personal")
@login_required
def personal():
    db = get_db()
    q = request.args.get("q", "").strip()
    f_obra = request.args.get("obra", "").strip()
    f_puesto = request.args.get("puesto", "").strip()
    desde = request.args.get("desde", "").strip()
    hasta = request.args.get("hasta", "").strip()

    sql = ("SELECT e.*, p.nombre AS puesto, o.nombre AS obra, o.estado AS plaza "
           "FROM empleados e "
           "LEFT JOIN puestos p ON p.id=e.puesto_id "
           "LEFT JOIN obras o ON o.id=p.obra_id WHERE 1=1")
    args = []
    if q:
        sql += (" AND (e.nombre LIKE ? OR e.primer_apellido LIKE ? OR "
                "e.segundo_apellido LIKE ? OR e.nss LIKE ? OR e.cedula LIKE ?)")
        like = f"%{q}%"
        args += [like, like, like, like, like]
    if f_obra:
        sql += " AND o.nombre = ?"; args.append(f_obra)
    if f_puesto:
        sql += " AND p.nombre = ?"; args.append(f_puesto)
    if desde:
        sql += " AND e.fecha_alta >= ?"; args.append(desde)
    if hasta:
        sql += " AND e.fecha_alta <= ?"; args.append(hasta)
    vis = obras_del_usuario(db)
    if vis is not None:
        if vis:
            ph = ",".join("?" * len(vis))
            sql += f" AND o.id IN ({ph})"; args += vis
        else:
            sql += " AND 1=0"
    sql += " ORDER BY e.primer_apellido, e.nombre"

    filas = []
    for r in db.execute(sql, args).fetchall():
        d = dict(r)
        texto, clase = dias_para_retencion(r)
        d["dias_texto"] = texto
        d["dias_clase"] = clase
        filas.append(d)

    obras = db.execute("SELECT DISTINCT nombre FROM obras ORDER BY nombre").fetchall()
    puestos = db.execute("SELECT DISTINCT nombre FROM puestos ORDER BY nombre").fetchall()
    return render_template("personal_list.html", empleados=filas,
                           obras=obras, puestos=puestos,
                           filtros={"q": q, "obra": f_obra, "puesto": f_puesto,
                                    "desde": desde, "hasta": hasta})


@app.route("/personal/<int:emp_id>/editar", methods=["GET", "POST"])
@min_rank(ADMIN_RANK)
def personal_editar(emp_id):
    db = get_db()
    emp = db.execute("SELECT * FROM empleados WHERE id=?", (emp_id,)).fetchone()
    if not emp:
        abort(404)
    puestos = db.execute(
        "SELECT p.id, p.nombre AS puesto, p.sueldo_semanal, p.viaticos_semanales, "
        "p.obra_id, o.nombre AS obra, o.estado FROM puestos p JOIN obras o ON o.id=p.obra_id "
        "WHERE p.activo=1 OR p.id=? ORDER BY o.nombre, p.nombre", (emp["puesto_id"],)).fetchall()
    obras_form = db.execute("SELECT * FROM obras ORDER BY nombre").fetchall()
    propuesta_por_obra = {o["id"]: propuesta_salario_alta(db, o["estado"]) for o in obras_form}

    if request.method == "POST":
        f = request.form
        nombre = titulo(f.get("nombre", "").strip())
        ap1 = titulo(f.get("primer_apellido", "").strip())
        curp = f.get("curp", "").strip().upper()
        pensionado = 1 if f.get("nss_generico") else 0
        autoriza_pension = f.get("autoriza_nss_generico", "").strip()
        errores = []
        if not nombre: errores.append("El nombre es obligatorio.")
        if not ap1: errores.append("El primer apellido es obligatorio.")
        if not f.get("puesto_id"): errores.append("Selecciona un puesto.")
        if not curp_valida(curp):
            errores.append("La CURP no tiene un formato valido.")
        if pensionado and role_rank(session.get("role", "")) < GERENTE_RANK:
            errores.append("Solo el superintendente o el administrador pueden marcar NSS generico de pensionados.")

        if pensionado:
            if not autoriza_pension:
                errores.append("Para el NSS generico de pensionados, indica que administrador autorizo.")
            nss = NSS_GENERICO
        else:
            nss = solo_digitos(f.get("nss", ""))
            if nss and not nss_valido(nss):
                errores.append("El NSS no es valido.")

        err_sol = validar_solicitud_imss(f.get("fecha_solicitud", ""), f.get("fecha_alta", ""))
        if err_sol: errores.append(err_sol)
        # NSS unico excluyendo al propio empleado (el NSS generico si se puede compartir)
        if nss and nss != NSS_GENERICO:
            otro = db.execute(
                "SELECT cedula FROM empleados WHERE nss=? AND id<>?",
                (nss, emp_id)).fetchone()
            if otro:
                errores.append(f"Otro trabajador ya tiene ese NSS (cedula {otro['cedula']}).")

        if errores:
            for e in errores: flash(e, "danger")
            datos = dict(emp); datos.update(f)
            return render_template("personal_form.html", puestos=puestos,
                                   obras=obras_form, propuesta_por_obra=propuesta_por_obra,
                                   estados=ESTADOS, tipos=TIPOS_CUENTA, bancos=BANCOS,
                                   estatus_docs=ESTATUS_DOCS, datos=datos,
                                   editar=True, emp_id=emp_id)

        db.execute(
            """UPDATE empleados SET nombre=?, primer_apellido=?, segundo_apellido=?,
               curp=?, rfc=?, cp_fiscal=?, nss=?, sexo=?, fecha_nacimiento=?,
               puesto_id=?, fecha_alta=?, fecha_solicitud=?, importe_alta_imss=?,
               infonavit_monto=?, viaticos_semanales=?, bono_semanal=?,
               nss_generico=?, autoriza_nss_generico=?,
               estatus_docs=?, observaciones=? WHERE id=?""",
            (nombre, ap1, titulo(f.get("segundo_apellido", "").strip()), curp,
             f.get("rfc", "").strip().upper(), f.get("cp_fiscal", "").strip(),
             nss, f.get("sexo", ""), f.get("fecha_nacimiento", ""),
             f.get("puesto_id"), f.get("fecha_alta"), f.get("fecha_solicitud", ""),
             float(f.get("importe_alta_imss") or 0), float(f.get("infonavit_monto") or 0),
             (float(f.get("viaticos_semanales")) if (f.get("viaticos_semanales") or "").strip() != "" else None),
             float(f.get("bono_semanal") or 0),
             pensionado, autoriza_pension,
             f.get("estatus_docs", "Pendiente de carga"),
             f.get("observaciones", "").strip(), emp_id))
        if pensionado and not emp["nss_generico"]:
            registrar_bitacora(db, "NSS generico de pensionado",
                                f"Empleado {emp['cedula']} ({nombre} {ap1}), autorizo: {autoriza_pension}")
        db.commit()
        flash(f"Empleado actualizado (cedula {emp['cedula']}).", "success")
        return redirect(url_for("personal"))

    return render_template("personal_form.html", puestos=puestos,
                           obras=obras_form, propuesta_por_obra=propuesta_por_obra,
                           estados=ESTADOS, tipos=TIPOS_CUENTA, bancos=BANCOS,
                           estatus_docs=ESTATUS_DOCS, datos=dict(emp),
                           editar=True, emp_id=emp_id)


# ---------------------------------------------------------------------------
# Documentos del empleado (subida de PDFs)  -- superintendente / residente
# ---------------------------------------------------------------------------
@app.route("/personal/<int:emp_id>/documentos")
@login_required
def documentos_empleado(emp_id):
    db = get_db()
    emp = db.execute(
        "SELECT e.*, p.nombre AS puesto, o.nombre AS obra FROM empleados e "
        "LEFT JOIN puestos p ON p.id=e.puesto_id LEFT JOIN obras o ON o.id=p.obra_id "
        "WHERE e.id=?", (emp_id,)).fetchone()
    if not emp:
        abort(404)
    subidos = {r["tipo"]: r for r in db.execute(
        "SELECT id, tipo, filename, subido_en, subido_por FROM documentos "
        "WHERE empleado_id=?", (emp_id,)).fetchall()}
    return render_template("documentos_empleado.html", emp=emp,
                           tipos_doc=TIPOS_DOC, requeridos=DOCS_REQUERIDOS,
                           subidos=subidos)


@app.route("/personal/<int:emp_id>/documentos/subir", methods=["POST"])
@login_required
def documento_subir(emp_id):
    db = get_db()
    tipo = request.form.get("tipo", "")
    archivo = request.files.get("archivo")
    if tipo not in [t[0] for t in TIPOS_DOC]:
        flash("Tipo de documento no valido.", "danger")
        return redirect(url_for("documentos_empleado", emp_id=emp_id))
    if not archivo or not archivo.filename.lower().endswith(".pdf"):
        flash("Sube un archivo en formato PDF.", "danger")
        return redirect(url_for("documentos_empleado", emp_id=emp_id))
    contenido = archivo.read()
    db.execute("DELETE FROM documentos WHERE empleado_id=? AND tipo=?", (emp_id, tipo))
    db.execute(
        "INSERT INTO documentos(empleado_id, tipo, filename, contenido, subido_en, subido_por) "
        "VALUES(?,?,?,?,?,?)",
        (emp_id, tipo, archivo.filename, contenido, date.today().isoformat(),
         session.get("nombre", "")))
    recomputar_estatus_docs(db, emp_id)
    db.commit()
    flash("Documento subido.", "success")
    return redirect(url_for("documentos_empleado", emp_id=emp_id))


@app.route("/documento/<int:doc_id>")
@login_required
def documento_ver(doc_id):
    db = get_db()
    d = db.execute("SELECT filename, contenido FROM documentos WHERE id=?", (doc_id,)).fetchone()
    if not d:
        abort(404)
    return send_file(io.BytesIO(d["contenido"]), mimetype="application/pdf",
                     download_name=d["filename"] or "documento.pdf")


@app.route("/personal/<int:emp_id>/documentos/<tipo>/eliminar", methods=["POST"])
@login_required
def documento_eliminar(emp_id, tipo):
    db = get_db()
    db.execute("DELETE FROM documentos WHERE empleado_id=? AND tipo=?", (emp_id, tipo))
    recomputar_estatus_docs(db, emp_id)
    db.commit()
    flash("Documento eliminado.", "success")
    return redirect(url_for("documentos_empleado", emp_id=emp_id))


@app.route("/personal/<int:emp_id>/documentos/validar", methods=["POST"])
@min_rank(GERENTE_RANK)
def documentos_validar(emp_id):
    db = get_db()
    tipos = [r["tipo"] for r in db.execute(
        "SELECT tipo FROM documentos WHERE empleado_id=?", (emp_id,)).fetchall()]
    if not all(t in tipos for t in DOCS_REQUERIDOS):
        flash("Faltan documentos requeridos para poder validar.", "danger")
    else:
        db.execute("UPDATE empleados SET estatus_docs='Validado' WHERE id=?", (emp_id,))
        db.commit()
        flash("Documentacion validada.", "success")
    return redirect(url_for("documentos_empleado", emp_id=emp_id))

@app.route("/personal/nuevo", methods=["GET", "POST"])
@login_required
def personal_nuevo():
    db = get_db()
    vis = obras_del_usuario(db)
    if vis is None:
        puestos = db.execute(
            "SELECT p.id, p.nombre AS puesto, p.sueldo_semanal, p.viaticos_semanales, "
            "p.obra_id, o.nombre AS obra, o.estado FROM puestos p JOIN obras o ON o.id=p.obra_id "
            "WHERE p.activo=1 ORDER BY o.nombre, p.nombre").fetchall()
    elif vis:
        ph = ",".join("?" * len(vis))
        puestos = db.execute(
            "SELECT p.id, p.nombre AS puesto, p.sueldo_semanal, p.viaticos_semanales, "
            "p.obra_id, o.nombre AS obra, o.estado FROM puestos p JOIN obras o ON o.id=p.obra_id "
            f"WHERE p.activo=1 AND p.obra_id IN ({ph}) ORDER BY o.nombre, p.nombre", vis).fetchall()
    else:
        puestos = []
    obras_form = obras_visibles(db)
    # propuesta de salario de alta por obra (SM de su zona + extra)
    propuesta_por_obra = {o["id"]: propuesta_salario_alta(db, o["estado"]) for o in obras_form}
    propuesta_salario = propuesta_salario_alta(db)   # SM general + extra (default)

    if request.method == "POST":
        f = request.form
        nombre = titulo(f.get("nombre", "").strip())
        ap1 = titulo(f.get("primer_apellido", "").strip())
        ap2 = titulo(f.get("segundo_apellido", "").strip())
        curp = f.get("curp", "").strip().upper()
        nss = f.get("nss", "").strip()
        sexo = f.get("sexo", "")
        fecha_nac = f.get("fecha_nacimiento", "")
        exime = 1 if f.get("exime_docs") else 0
        autoriza = f.get("autoriza_tercero", "").strip()
        pensionado = 1 if f.get("nss_generico") else 0
        autoriza_pension = f.get("autoriza_nss_generico", "").strip()

        errores = []
        if not nombre: errores.append("El nombre es obligatorio.")
        if not ap1: errores.append("El primer apellido es obligatorio.")
        if not f.get("puesto_id"): errores.append("Selecciona un puesto (obra + categoria).")
        elif f.get("puesto_id") not in {str(p["id"]) for p in puestos}:
            errores.append("El puesto seleccionado no pertenece a una obra que puedas gestionar.")
        if not f.get("fecha_alta"): errores.append("La fecha de alta es obligatoria.")
        err_sol = validar_solicitud_imss(f.get("fecha_solicitud", ""), f.get("fecha_alta", ""))
        if err_sol: errores.append(err_sol)
        if exime and role_rank(session.get("role", "")) < GERENTE_RANK:
            errores.append("Solo el superintendente o el administrador pueden eximir documentos.")
        if pensionado and role_rank(session.get("role", "")) < GERENTE_RANK:
            errores.append("Solo el superintendente o el administrador pueden marcar NSS generico de pensionados.")

        # CURP: formato obligatorio
        if not curp_valida(curp):
            errores.append("La CURP no tiene un formato valido (18 caracteres).")

        # NSS: obligatorio, salvo que se exima con autorizacion de un tercero, o que sea
        # personal pensionado con NSS generico (autorizado por un administrador).
        if pensionado:
            if not autoriza_pension:
                errores.append("Para el NSS generico de pensionados, indica que administrador autorizo.")
            nss_norm = NSS_GENERICO
        elif exime:
            if not autoriza:
                errores.append("Para eximir documentos, indica quien autoriza (tercero).")
            nss_norm = ""
        else:
            if not nss:
                errores.append("El NSS es obligatorio (o marca 'eximir' con autorizacion).")
            elif not nss_valido(nss):
                errores.append("El NSS no es valido (deben ser 11 digitos con digito verificador correcto).")
            nss_norm = solo_digitos(nss)

        # NSS unico: no puede repetirse (salvo el NSS generico de pensionados, que si se comparte)
        if nss_norm and nss_norm != NSS_GENERICO:
            ya = db.execute(
                "SELECT cedula, nombre, primer_apellido FROM empleados WHERE nss=?",
                (nss_norm,)).fetchone()
            if ya:
                errores.append(
                    f"Ya existe un trabajador con ese NSS (cedula {ya['cedula']}: "
                    f"{ya['nombre']} {ya['primer_apellido']}).")

        # Cuenta bancaria opcional
        cta_num = solo_digitos(f.get("cta_numero", ""))
        cta_inst = f.get("cta_institucion", "").strip()
        cta_tipo = f.get("cta_tipo", "")
        if cta_num:
            dup = db.execute(
                "SELECT e.nombre, e.primer_apellido, e.segundo_apellido, e.estatus, "
                "o.nombre AS obra FROM cuentas_bancarias c "
                "JOIN empleados e ON e.id=c.empleado_id "
                "LEFT JOIN puestos p ON p.id=e.puesto_id "
                "LEFT JOIN obras o ON o.id=p.obra_id "
                "WHERE c.numero=?", (cta_num,)).fetchone()
            if dup:
                titular = " ".join(x for x in [dup["nombre"], dup["primer_apellido"],
                                               dup["segundo_apellido"]] if x)
                obra_dup = dup["obra"] or "sin obra"
                estatus_dup = dup["estatus"] or "activo"
                errores.append(
                    f"Esa cuenta ya esta registrada a nombre de {titular} "
                    f"(obra: {obra_dup}, estatus: {estatus_dup}). No se permiten duplicados.")
                # Avisar al administrador por correo (no bloquea si el correo no esta configurado)
                nuevo = " ".join(x for x in [nombre, ap1, ap2] if x)
                send_admin_alert(
                    "POLTECH - Intento de cuenta bancaria duplicada",
                    (f"Se intento dar de alta a '{nuevo}' con la cuenta {cta_num}, "
                     f"que ya pertenece a '{titular}' (obra: {obra_dup}, "
                     f"estatus: {estatus_dup}).\n\nRevisar antes de continuar."))
            if not cta_inst or not cta_tipo:
                errores.append("Para la cuenta bancaria, captura institucion y tipo de cuenta.")
            err_cta = validar_cuenta(cta_tipo, cta_num)
            if err_cta:
                errores.append(err_cta)

        if errores:
            for e in errores:
                flash(e, "danger")
            return render_template("personal_form.html", puestos=puestos,
                                   obras=obras_form, propuesta_por_obra=propuesta_por_obra,
                                   propuesta_salario=propuesta_salario,
                                   estados=ESTADOS, tipos=TIPOS_CUENTA, bancos=BANCOS, datos=f)

        cedula = siguiente_cedula(db)
        observaciones = f.get("observaciones", "").strip()
        cur = db.execute(
            """INSERT INTO empleados
               (cedula, nombre, primer_apellido, segundo_apellido, curp, rfc, cp_fiscal,
                nss, sexo, fecha_nacimiento, puesto_id, fecha_alta, fecha_solicitud,
                fecha_registro, importe_alta_imss, infonavit_monto, viaticos_semanales,
                bono_semanal, exime_docs, autoriza_tercero, nss_generico,
                autoriza_nss_generico, observaciones, estatus_docs)
               VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
            (cedula, nombre, ap1, ap2, curp,
             f.get("rfc", "").strip().upper(), f.get("cp_fiscal", "").strip(),
             nss_norm, sexo, fecha_nac, f.get("puesto_id"),
             f.get("fecha_alta"), f.get("fecha_solicitud", ""),
             date.today().isoformat(), float(f.get("importe_alta_imss") or 0),
             float(f.get("infonavit_monto") or 0),
             (float(f.get("viaticos_semanales")) if (f.get("viaticos_semanales") or "").strip() != "" else None),
             float(f.get("bono_semanal") or 0),
             exime, autoriza, pensionado, autoriza_pension,
             observaciones, "Pendiente de carga"),
        )
        emp_id = cur.lastrowid
        if cta_num:
            db.execute(
                "INSERT INTO cuentas_bancarias(empleado_id, institucion, tipo_cuenta, numero) VALUES(?,?,?,?)",
                (emp_id, cta_inst, cta_tipo, cta_num),
            )
        if pensionado:
            registrar_bitacora(db, "NSS generico de pensionado",
                                f"Empleado {cedula} ({nombre} {ap1}), autorizo: {autoriza_pension}")
        db.commit()

        # Generar el contrato automaticamente y avisar por correo a quien dio el alta
        try:
            crear_contrato_para(db, emp_id, avisar_a=session.get("username"))
            db.commit()
        except Exception as e:
            app.logger.error("No se pudo generar el contrato: %s", e)

        for a in revisar_curp(curp, nombre, ap1, ap2, sexo, fecha_nac):
            flash("Aviso: " + a, "warning")
        flash(f"Empleado dado de alta con cedula {cedula}. Se genero su contrato para revision.", "success")
        return redirect(url_for("personal"))

    return render_template("personal_form.html", puestos=puestos,
                           obras=obras_form, propuesta_por_obra=propuesta_por_obra,
                           propuesta_salario=propuesta_salario,
                           estados=ESTADOS, tipos=TIPOS_CUENTA, bancos=BANCOS, datos={})

# ---------------------------------------------------------------------------
# Carga masiva de trabajadores (por Excel)
# ---------------------------------------------------------------------------
COLS_CARGA = ["NOMBRE(S)", "PRIMER APELLIDO", "SEGUNDO APELLIDO", "CURP", "RFC",
              "CP FISCAL", "NSS", "SEXO (H/M)", "FECHA NACIMIENTO (DD-MM-AAAA)",
              "OBRA", "PUESTO", "FECHA ALTA (DD-MM-AAAA)", "SALARIO ALTA IMSS",
              "INFONAVIT SEMANAL", "BANCO", "TIPO DE CUENTA", "NUMERO DE CUENTA",
              "OBSERVACIONES"]

@app.route("/personal/plantilla")
@login_required
def personal_plantilla():
    db = get_db()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Trabajadores"
    ws.append(COLS_CARGA)
    ws.append(["JUAN", "PEREZ", "LOPEZ", "PELJ800101HDFRPN09", "PELJ800101AB1",
               "54000", "12345678903", "H", "01-01-1980", "Torre Reforma",
               "Soldador", "28-07-2026", "3000", "0", "BBVA Mexico",
               "CLABE interbancaria", "012180001234567890", ""])
    from openpyxl.styles import Font, PatternFill
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="16233C")
    for i in range(1, len(COLS_CARGA) + 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = 20

    # -----------------------------------------------------------------
    # Hoja oculta "Listas": aqui viven los valores de los menus desplegables
    # (Sexo, Obra, Banco, Tipo de cuenta) y, para cada obra, la lista de sus
    # puestos (para que el desplegable de Puesto dependa de la Obra elegida).
    # -----------------------------------------------------------------
    from openpyxl.utils import get_column_letter
    from openpyxl.workbook.defined_name import DefinedName
    from openpyxl.worksheet.datavalidation import DataValidation

    obras_l = obras_visibles(db)
    lst = wb.create_sheet("Listas")
    lst.sheet_state = "hidden"

    def nombrar_rango(nombre, ref):
        wb.defined_names[nombre] = DefinedName(nombre, attr_text=ref)

    # Columna A: nombre de la obra. Columna B: clave interna (para ubicar sus puestos).
    for idx, o in enumerate(obras_l, start=1):
        lst.cell(row=idx, column=1, value=o["nombre"])
        lst.cell(row=idx, column=2, value=f"OBRA_{o['id']}")
    if obras_l:
        nombrar_rango("ListaObras", f"Listas!$A$1:$A${len(obras_l)}")
        nombrar_rango("MapaObraKey", f"Listas!$A$1:$B${len(obras_l)}")

    # Columna D: bancos. Columna E: tipo de cuenta.
    for idx, b in enumerate(BANCOS, start=1):
        lst.cell(row=idx, column=4, value=b)
    nombrar_rango("ListaBancos", f"Listas!$D$1:$D${len(BANCOS)}")
    for idx, t in enumerate(TIPOS_CUENTA, start=1):
        lst.cell(row=idx, column=5, value=t)
    nombrar_rango("ListaTipoCuenta", f"Listas!$E$1:$E${len(TIPOS_CUENTA)}")

    # A partir de la columna G, una columna por obra con sus puestos.
    col = 7
    for o in obras_l:
        filas_puesto = db.execute(
            "SELECT nombre FROM puestos WHERE obra_id=? AND activo=1 ORDER BY nombre",
            (o["id"],)).fetchall()
        nombres_puesto = [r["nombre"] for r in filas_puesto]
        if not nombres_puesto:
            continue
        for r, pnom in enumerate(nombres_puesto, start=1):
            lst.cell(row=r, column=col, value=pnom)
        letra = get_column_letter(col)
        nombrar_rango(f"OBRA_{o['id']}", f"Listas!${letra}$1:${letra}${len(nombres_puesto)}")
        col += 1

    # -----------------------------------------------------------------
    # Menus desplegables en la hoja "Trabajadores"
    # -----------------------------------------------------------------
    ultima_fila = 500  # deja espacio para pegar/escribir muchos trabajadores

    dv_sexo = DataValidation(type="list", formula1='"H,M"', allow_blank=True,
                             showErrorMessage=True, errorTitle="Valor no valido",
                             error="Escribe H (hombre) o M (mujer).")
    ws.add_data_validation(dv_sexo)
    dv_sexo.add(f"H2:H{ultima_fila}")

    if obras_l:
        dv_obra = DataValidation(type="list", formula1="ListaObras", allow_blank=True,
                                  showErrorMessage=True, errorTitle="Obra no valida",
                                  error="Elige una obra de la lista (ya dada de alta en Obras).")
        ws.add_data_validation(dv_obra)
        dv_obra.add(f"J2:J{ultima_fila}")

        # El Puesto depende de la Obra escrita en la misma fila (columna J).
        dv_puesto = DataValidation(
            type="list", formula1="INDIRECT(VLOOKUP($J2,MapaObraKey,2,0))",
            allow_blank=True, showErrorMessage=False)
        ws.add_data_validation(dv_puesto)
        dv_puesto.add(f"K2:K{ultima_fila}")

    dv_banco = DataValidation(type="list", formula1="ListaBancos", allow_blank=True,
                              showErrorMessage=False)
    ws.add_data_validation(dv_banco)
    dv_banco.add(f"O2:O{ultima_fila}")

    dv_tipo = DataValidation(type="list", formula1="ListaTipoCuenta", allow_blank=True,
                             showErrorMessage=False)
    ws.add_data_validation(dv_tipo)
    dv_tipo.add(f"P2:P{ultima_fila}")

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    nombre = f"Plantilla de alta de trabajadores_{datetime.now():%Y-%m-%d_%H%M}.xlsx"
    return send_file(buf, as_attachment=True, download_name=nombre,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.route("/personal/carga", methods=["GET", "POST"])
@login_required
def personal_carga():
    if request.method == "POST":
        archivo = request.files.get("archivo")
        if not archivo or not archivo.filename.lower().endswith(".xlsx"):
            flash("Sube un archivo de Excel (.xlsx) usando la plantilla.", "danger")
            return render_template("personal_carga.html", resumen=None)

        db = get_db()
        wb = openpyxl.load_workbook(archivo, data_only=True)
        ws = wb.active
        creados = 0
        contratos_gen = 0
        errores = []
        duplicadas = []   # cuentas bancarias repetidas

        for i, fila in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
            if fila is None or all(c is None or str(c).strip() == "" for c in fila):
                continue
            vals = (list(fila) + [None] * len(COLS_CARGA))[:len(COLS_CARGA)]
            (nombre, ap1, ap2, curp, rfc, cp, nss, sexo, fnac, obra_nom, puesto_nom,
             falta, salario, infonavit, banco, tipo_cta, num_cta, obs) = vals

            def txt(x): return "" if x is None else str(x).strip()
            def fecha_txt(x):
                """Convierte la fecha a AAAA-MM-DD (formato interno) sin importar como
                haya llegado: fecha nativa de Excel, DD-MM-AAAA o DD/MM/AAAA (formato
                mexicano, el que se pide en la plantilla) o ya en AAAA-MM-DD.
                Si no se reconoce el formato, se regresa tal cual para que la fila
                se marque con error en vez de guardar una fecha incorrecta."""
                if isinstance(x, (datetime, date)):
                    return x.strftime("%Y-%m-%d")
                s = txt(x)
                if not s:
                    return ""
                if re.fullmatch(r"\d{4}-\d{2}-\d{2}", s):
                    return s
                m = re.fullmatch(r"(\d{1,2})[/-](\d{1,2})[/-](\d{4})", s)
                if m:
                    d, mo, y = (int(g) for g in m.groups())
                    try:
                        return date(y, mo, d).isoformat()
                    except ValueError:
                        return s
                return s
            def fecha_valida(s): return bool(re.fullmatch(r"\d{4}-\d{2}-\d{2}", s or ""))
            nombre, ap1, ap2 = titulo(txt(nombre)), titulo(txt(ap1)), titulo(txt(ap2))
            curp = txt(curp).upper()
            nss = solo_digitos(txt(nss))
            # el NSS a veces llega como numero con ".0"
            if nss.endswith(".0"):
                nss = nss[:-2]
            fnac = fecha_txt(fnac)
            falta = fecha_txt(falta)
            nombre_completo = " ".join(x for x in [nombre, ap1, ap2] if x) or "sin nombre"

            errs = []
            if not nombre: errs.append("nombre vacio")
            if not ap1: errs.append("primer apellido vacio")
            if not curp_valida(curp): errs.append("CURP invalida")
            if not nss: errs.append("NSS vacio")
            elif not nss_valido(nss): errs.append("NSS invalido")
            if fnac and not fecha_valida(fnac):
                errs.append(f"fecha de nacimiento '{fnac}' no reconocida (usa DD-MM-AAAA)")
            if not falta:
                errs.append("fecha de alta vacia")
            elif not fecha_valida(falta):
                errs.append(f"fecha de alta '{falta}' no reconocida (usa DD-MM-AAAA)")
            puesto = resolver_puesto(db, obra_nom, puesto_nom)
            if not puesto:
                errs.append(f"la obra/puesto '{txt(obra_nom)} / {txt(puesto_nom)}' no existe en el catalogo")
            # NSS duplicado
            if nss and db.execute("SELECT 1 FROM empleados WHERE nss=?", (nss,)).fetchone():
                errs.append("ya existe un trabajador con ese NSS")

            if errs:
                errores.append(f"Fila {i} ({nombre_completo}): " + "; ".join(errs))
                continue

            # Si no viene el salario de alta, se sugiere el minimo del estado de la obra + extra
            # (igual que en el alta individual), en vez de dejarlo en $0.
            salario_txt = txt(salario)
            if salario_txt:
                salario_final = float(salario_txt)
            else:
                salario_final = propuesta_salario_alta(db, puesto["estado"])

            cedula = siguiente_cedula(db)
            cur = db.execute(
                """INSERT INTO empleados
                   (cedula, nombre, primer_apellido, segundo_apellido, curp, rfc, cp_fiscal,
                    nss, sexo, fecha_nacimiento, puesto_id, fecha_alta, fecha_registro,
                    importe_alta_imss, infonavit_monto, exime_docs, autoriza_tercero,
                    observaciones, estatus_docs)
                   VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                (cedula, nombre, ap1, ap2, curp, txt(rfc).upper(), txt(cp), nss,
                 txt(sexo).upper()[:1], fnac, puesto["id"], falta,
                 date.today().isoformat(), salario_final, float(infonavit or 0),
                 0, "", txt(obs), "Pendiente de carga"))
            emp_id = cur.lastrowid
            creados += 1
            try:
                crear_contrato_para(db, emp_id, avisar_a=None)
                contratos_gen += 1
            except Exception as e:
                app.logger.error("Contrato en carga (fila %s): %s", i, e)

            # cuenta bancaria (opcional): revisar duplicado antes de insertar
            num_cta = solo_digitos(txt(num_cta))
            if num_cta:
                dup = db.execute(
                    "SELECT e.cedula, e.nombre, e.primer_apellido, e.estatus, o.nombre AS obra "
                    "FROM cuentas_bancarias c JOIN empleados e ON e.id=c.empleado_id "
                    "LEFT JOIN puestos p ON p.id=e.puesto_id "
                    "LEFT JOIN obras o ON o.id=p.obra_id WHERE c.numero=?",
                    (num_cta,)).fetchone()
                if dup:
                    titular = f"{dup['nombre']} {dup['primer_apellido']}".strip()
                    duplicadas.append(
                        f"Fila {i}: la cuenta {num_cta} de {nombre} {ap1} ya pertenece a "
                        f"{titular} (cedula {dup['cedula']}, obra {dup['obra'] or '-'}, "
                        f"estatus {dup['estatus'] or 'activo'}). No se guardo la cuenta.")
                else:
                    db.execute(
                        "INSERT INTO cuentas_bancarias(empleado_id, institucion, tipo_cuenta, numero) "
                        "VALUES(?,?,?,?)", (emp_id, txt(banco), txt(tipo_cta), num_cta))

        db.commit()

        if duplicadas:
            send_admin_alert(
                "POLTECH - Cuentas bancarias duplicadas en carga masiva",
                "Durante una carga masiva se detectaron cuentas repetidas:\n\n" +
                "\n".join(duplicadas))

        resumen = {"creados": creados, "contratos": contratos_gen,
                   "errores": errores, "duplicadas": duplicadas}
        return render_template("personal_carga.html", resumen=resumen)

    return render_template("personal_carga.html", resumen=None)

# ---------------------------------------------------------------------------
# Contratos generados
# ---------------------------------------------------------------------------
@app.route("/contratos")
@login_required
def contratos():
    db = get_db()
    vis = obras_del_usuario(db)
    sql = ("SELECT c.id, c.nombre_archivo, c.generado_en, c.generado_por, c.estatus, "
           "c.revisado_por, c.revisado_en, e.cedula, e.nombre, e.primer_apellido, "
           "e.segundo_apellido, p.nombre AS puesto, o.nombre AS obra, o.id AS obra_id "
           "FROM contratos c JOIN empleados e ON e.id=c.empleado_id "
           "JOIN puestos p ON p.id=e.puesto_id JOIN obras o ON o.id=p.obra_id WHERE 1=1")
    args = []
    if vis is not None:
        if vis:
            sql += " AND o.id IN (%s)" % ",".join("?" * len(vis)); args += vis
        else:
            sql += " AND 0"
    sql += " ORDER BY c.generado_en DESC"
    filas = db.execute(sql, args).fetchall()
    return render_template("contratos_list.html", contratos=filas)


@app.route("/contratos/<int:cid>/descargar")
@login_required
def contrato_descargar(cid):
    db = get_db()
    c = db.execute("SELECT c.*, p.obra_id FROM contratos c "
                   "JOIN empleados e ON e.id=c.empleado_id "
                   "JOIN puestos p ON p.id=e.puesto_id WHERE c.id=?", (cid,)).fetchone()
    if not c:
        abort(404)
    if not puede_ver_obra(db, c["obra_id"]):
        abort(403)
    return send_file(io.BytesIO(c["contenido"]), as_attachment=True,
                     download_name=c["nombre_archivo"],
                     mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document")


@app.route("/contratos/<int:cid>/revisado", methods=["POST"])
@min_rank(GERENTE_RANK)
def contrato_revisado(cid):
    db = get_db()
    c = db.execute("SELECT c.*, p.obra_id FROM contratos c "
                   "JOIN empleados e ON e.id=c.empleado_id "
                   "JOIN puestos p ON p.id=e.puesto_id WHERE c.id=?", (cid,)).fetchone()
    if not c:
        abort(404)
    if not puede_ver_obra(db, c["obra_id"]):
        abort(403)
    db.execute("UPDATE contratos SET estatus='revisado', revisado_por=?, revisado_en=? WHERE id=?",
               (session.get("nombre"), datetime.now().isoformat(timespec="seconds"), cid))
    registrar_bitacora(db, "Contrato revisado", f"Contrato {cid}")
    db.commit()
    flash("Contrato marcado como revisado.", "success")
    return redirect(url_for("contratos"))


@app.route("/personal/<int:emp_id>/contrato", methods=["POST"])
@login_required
def contrato_regenerar(emp_id):
    db = get_db()
    emp = db.execute("SELECT e.id, p.obra_id FROM empleados e JOIN puestos p ON p.id=e.puesto_id "
                     "WHERE e.id=?", (emp_id,)).fetchone()
    if not emp:
        abort(404)
    if not puede_ver_obra(db, emp["obra_id"]):
        abort(403)
    try:
        crear_contrato_para(db, emp_id, avisar_a=session.get("username"))
        db.commit()
        flash("Contrato regenerado y enviado a tu correo.", "success")
    except Exception as e:
        app.logger.error("No se pudo regenerar el contrato: %s", e)
        flash("No se pudo generar el contrato.", "danger")
    return redirect(url_for("contratos"))


# ---------------------------------------------------------------------------
# Integracion API market (consulta de NSS por CURP y vigencia)
# ---------------------------------------------------------------------------
def _apimarket_call(db, url, params):
    """POST a un endpoint de API market. Los parametros (curp, nss) viajan en la
    URL como query string (?curp=...), con Authorization: Bearer <token>.
    Devuelve (ok, datos) donde datos es dict/lista (JSON) o texto/mensaje de error."""
    token = (get_param(db, "apimarket_token", "") or os.environ.get("APIMARKET_TOKEN", "")).strip()
    if not url:
        return False, "Falta configurar la URL del servicio en Parametros (API market)."
    if not token:
        return False, "Falta configurar el token de API market en Parametros."
    try:
        full = url
        limpios = {k: v for k, v in (params or {}).items() if v}
        if limpios:
            sep = "&" if "?" in url else "?"
            full = url + sep + urllib.parse.urlencode(limpios)
        req = urllib.request.Request(full, method="POST")
        req.add_header("Authorization", "Bearer " + token)
        req.add_header("Accept", "application/json")
        with urllib.request.urlopen(req, timeout=25) as resp:
            body = resp.read().decode("utf-8", "ignore")
        try:
            return True, json.loads(body)
        except ValueError:
            return True, body
    except urllib.error.HTTPError as e:
        detalle = e.read().decode("utf-8", "ignore")[:400] if hasattr(e, "read") else ""
        return False, f"HTTP {e.code}: {detalle}"
    except Exception as e:
        return False, str(e)


def _buscar_nss(obj):
    """Busca recursivamente un NSS (11 digitos) en la respuesta del API."""
    if isinstance(obj, dict):
        for k, v in obj.items():
            if isinstance(v, (str, int)) and re.fullmatch(r"\d{11}", str(v)):
                if any(t in k.lower() for t in ("nss", "seguro", "social")):
                    return str(v)
        for v in obj.values():
            r = _buscar_nss(v)
            if r:
                return r
    elif isinstance(obj, list):
        for v in obj:
            r = _buscar_nss(v)
            if r:
                return r
    elif isinstance(obj, (str, int)) and re.fullmatch(r"\d{11}", str(obj)):
        return str(obj)
    return None


def _vigencia_datos(resp):
    """Extrae los campos utiles de la respuesta de vigencia (dentro de 'data')."""
    d = resp.get("data") if isinstance(resp, dict) else None
    if not isinstance(d, dict):
        d = resp if isinstance(resp, dict) else {}
    def si_no(v):
        s = str(v).strip().upper()
        if s in ("SI", "SÍ", "TRUE", "1"): return "SI"
        if s in ("NO", "FALSE", "0"): return "NO"
        return str(v)
    return {
        "nss": d.get("nss") or d.get("NSS") or "",
        "servicio_medico": si_no(d.get("conDerechoSm", "")),
        "incapacidad": si_no(d.get("conDerechoInc", "")),
        "vigente_hasta": d.get("vigenteHasta") or "",
        "codigo": (resp.get("codigoValidacion") if isinstance(resp, dict) else "") or "",
        "mensaje": (resp.get("message") if isinstance(resp, dict) else "") or "",
    }


def generar_vigencia_docx(db, vig, curp):
    """Comprobante de vigencia de derechos (IMSS) en Word."""
    from docx import Document
    from docx.shared import Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH

    razon = get_param(db, "empresa_razon_social", "") or ""
    doc = Document()
    doc.styles["Normal"].font.name = "Arial"; doc.styles["Normal"].font.size = Pt(11)

    def par(txt, *, bold=False, center=False, size=None, color=None, space=6):
        p = doc.add_paragraph(); p.paragraph_format.space_after = Pt(space)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER if center else WD_ALIGN_PARAGRAPH.LEFT
        r = p.add_run(txt); r.bold = bold
        if size: r.font.size = Pt(size)
        if color: r.font.color.rgb = RGBColor(*color)
        return p

    par("COMPROBANTE DE VIGENCIA DE DERECHOS (IMSS)", bold=True, center=True, size=14, color=(0x0A, 0x56, 0xC0))
    if razon:
        par(razon, center=True, size=10, color=(0x16, 0x23, 0x3C), space=12)
    par(f"Fecha de consulta: {date.today().strftime('%d/%m/%Y')}", space=12)
    par(f"Número de Seguridad Social (NSS): {vig['nss']}", bold=True, space=4)
    par(f"CURP: {curp}", bold=True, space=12)

    par(f"Con derecho a servicio médico: {vig['servicio_medico']}", space=4)
    par(f"Con derecho a incapacidades: {vig['incapacidad']}", space=4)
    if vig["vigente_hasta"]:
        par(f"Vigente hasta: {vig['vigente_hasta']}", space=12)
    else:
        par("", space=8)

    if vig["codigo"]:
        par(f"Código de validación: {vig['codigo']}", size=9, color=(0x55, 0x55, 0x55), space=4)
    par("Consulta realizada a través de API Market (fuente: IMSS). Documento informativo.",
        size=8, color=(0x88, 0x88, 0x88), space=0)

    buf = io.BytesIO(); doc.save(buf); buf.seek(0)
    return buf.getvalue(), f"Vigencia_{vig['nss'] or curp}.docx"


@app.route("/api/consulta-nss", methods=["POST"])
@login_required
def api_consulta_nss():
    db = get_db()
    curp = ((request.form.get("curp") or (request.json or {}).get("curp") or "")).strip().upper()
    if not re.fullmatch(r"[A-Z0-9]{18}", curp):
        return jsonify({"ok": False, "error": "CURP invalido (deben ser 18 caracteres)."})
    ok, data = _apimarket_call(db, get_param(db, "apimarket_url_nss", ""), {"curp": curp})
    if not ok:
        return jsonify({"ok": False, "error": data})
    return jsonify({"ok": True, "nss": _buscar_nss(data), "raw": data})


@app.route("/apimarket", methods=["GET", "POST"])
@login_required
def apimarket_tool():
    db = get_db()
    resultado = None
    if request.method == "POST":
        tipo = request.form.get("tipo")
        curp = (request.form.get("curp") or "").strip().upper()
        nss = solo_digitos(request.form.get("nss") or "")
        if tipo == "nss":
            ok, data = _apimarket_call(db, get_param(db, "apimarket_url_nss", ""), {"curp": curp})
            resultado = {"ok": ok, "tipo": "nss", "titulo": "Consulta de NSS por CURP",
                         "nss": _buscar_nss(data) if ok else None,
                         "data": data if ok else None, "error": None if ok else data}
        else:
            ok, data = _apimarket_call(db, get_param(db, "apimarket_url_vigencia", ""),
                                       {"nss": nss, "curp": curp})
            vig = _vigencia_datos(data) if ok else None
            resultado = {"ok": ok, "tipo": "vigencia", "titulo": "Consulta de vigencia (NSS y CURP)",
                         "vig": vig, "curp": curp, "data": data if ok else None,
                         "error": None if ok else data,
                         "vig_json": json.dumps({"vig": vig, "curp": curp}) if ok and vig else ""}
        resultado["raw"] = json.dumps(resultado.get("data"), indent=2, ensure_ascii=False) \
            if resultado.get("data") is not None else ""
    configurado = bool((get_param(db, "apimarket_token", "") or "").strip())
    return render_template("apimarket.html", resultado=resultado, configurado=configurado)


@app.route("/apimarket/vigencia/comprobante", methods=["POST"])
@login_required
def apimarket_vigencia_comprobante():
    db = get_db()
    try:
        payload = json.loads(request.form.get("vig_json") or "{}")
        vig = payload.get("vig") or {}
        curp = payload.get("curp") or ""
    except (ValueError, TypeError):
        flash("No se pudo generar el comprobante.", "danger")
        return redirect(url_for("apimarket_tool"))
    datos, archivo = generar_vigencia_docx(db, vig, curp)
    return send_file(io.BytesIO(datos), as_attachment=True, download_name=archivo,
                     mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document")


@app.route("/contratos/generar-todos", methods=["POST"])
@login_required
def contratos_generar_todos():
    db = get_db()
    vis = obras_del_usuario(db)
    sql = ("SELECT e.id FROM empleados e JOIN puestos p ON p.id=e.puesto_id "
           "WHERE (e.estatus IS NULL OR e.estatus != 'baja') "
           "AND e.id NOT IN (SELECT empleado_id FROM contratos)")
    args = []
    if vis is not None:
        if vis:
            sql += " AND p.obra_id IN (%s)" % ",".join("?" * len(vis)); args += vis
        else:
            sql += " AND 0"
    faltantes = [r["id"] for r in db.execute(sql, args).fetchall()]
    generados = 0
    for eid in faltantes:
        try:
            crear_contrato_para(db, eid, avisar_a=None)
            generados += 1
        except Exception as e:
            app.logger.error("Contrato masivo (emp %s): %s", eid, e)
    db.commit()
    if generados:
        flash(f"Se generaron {generados} contratos del personal dado de alta.", "success")
    else:
        flash("No habia contratos por generar (todo el personal activo ya tiene contrato).", "info")
    return redirect(url_for("contratos"))


# ---------------------------------------------------------------------------
# Bajas de personal (aviso/escrito de baja)
# ---------------------------------------------------------------------------
def _emp_para_baja(db, emp_id):
    return db.execute(
        "SELECT e.*, p.nombre AS puesto, o.nombre AS obra, o.id AS obra_id "
        "FROM empleados e JOIN puestos p ON p.id=e.puesto_id "
        "JOIN obras o ON o.id=p.obra_id WHERE e.id=?", (emp_id,)).fetchone()


@app.route("/personal/<int:emp_id>/baja", methods=["GET", "POST"])
@login_required
def personal_baja(emp_id):
    db = get_db()
    emp = _emp_para_baja(db, emp_id)
    if not emp:
        abort(404)
    if not puede_ver_obra(db, emp["obra_id"]):
        abort(403)
    if request.method == "POST":
        fecha = (request.form.get("fecha_baja") or "").strip()
        motivo = (request.form.get("motivo_baja") or "").strip()
        if not fecha:
            flash("Indica la fecha de baja.", "danger")
            return render_template("personal_baja.html", emp=emp, motivos=MOTIVOS_BAJA)
        db.execute("UPDATE empleados SET estatus='baja', fecha_baja=?, motivo_baja=? WHERE id=?",
                   (fecha, motivo, emp_id))
        registrar_bitacora(db, "Baja de trabajador",
                           f"{emp['nombre']} {emp['primer_apellido']} (cedula {emp['cedula']}) "
                           f"baja {fecha} - {motivo}")
        db.commit()
        send_admin_alert(
            "POLTECH - Aviso de baja de trabajador",
            (f"El trabajador {emp['nombre']} {emp['primer_apellido']} {emp['segundo_apellido'] or ''} "
             f"causa baja desde {fecha}.\nMotivo: {motivo or 'no especificado'}.\n"
             f"CURP: {emp['curp']}  RFC: {emp['rfc']}  NSS: {emp['nss']}  Puesto: {emp['puesto']}.\n\n"
             f"Favor de tramitar la baja ante el IMSS."),
            extra=[get_param(db, "email_contador", "")])
        flash("Trabajador dado de baja. Ya puedes descargar el aviso de baja.", "success")
        return redirect(url_for("bajas"))
    return render_template("personal_baja.html", emp=emp, motivos=MOTIVOS_BAJA)


@app.route("/personal/<int:emp_id>/escrito-baja")
@login_required
def escrito_baja(emp_id):
    db = get_db()
    emp = _emp_para_baja(db, emp_id)
    if not emp:
        abort(404)
    if not puede_ver_obra(db, emp["obra_id"]):
        abort(403)
    datos, archivo = generar_baja_docx(db, emp)
    return send_file(io.BytesIO(datos), as_attachment=True, download_name=archivo,
                     mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document")


@app.route("/personal/<int:emp_id>/reactivar", methods=["POST"])
@min_rank(GERENTE_RANK)
def personal_reactivar(emp_id):
    db = get_db()
    emp = _emp_para_baja(db, emp_id)
    if not emp:
        abort(404)
    if not puede_ver_obra(db, emp["obra_id"]):
        abort(403)
    db.execute("UPDATE empleados SET estatus='activo', fecha_baja=NULL, motivo_baja=NULL WHERE id=?", (emp_id,))
    registrar_bitacora(db, "Reingreso de trabajador",
                       f"{emp['nombre']} {emp['primer_apellido']} (cedula {emp['cedula']}) reactivado")
    db.commit()
    flash("Trabajador reactivado.", "success")
    return redirect(url_for("bajas"))


@app.route("/bajas")
@login_required
def bajas():
    db = get_db()
    vis = obras_del_usuario(db)
    f_obra = (request.args.get("obra_id") or "").strip()
    f_desde = (request.args.get("desde") or "").strip()
    f_hasta = (request.args.get("hasta") or "").strip()
    sql = ("SELECT e.id, e.cedula, e.nombre, e.primer_apellido, e.segundo_apellido, e.nss, "
           "e.curp, e.rfc, e.fecha_alta, e.fecha_baja, e.motivo_baja, p.nombre AS puesto, "
           "o.nombre AS obra, o.id AS obra_id FROM empleados e "
           "JOIN puestos p ON p.id=e.puesto_id JOIN obras o ON o.id=p.obra_id "
           "WHERE e.estatus='baja'")
    args = []
    if vis is not None:
        if vis:
            sql += " AND o.id IN (%s)" % ",".join("?" * len(vis)); args += vis
        else:
            sql += " AND 0"
    if f_obra:
        sql += " AND o.id=?"; args.append(f_obra)
    if f_desde:
        sql += " AND e.fecha_baja>=?"; args.append(f_desde)
    if f_hasta:
        sql += " AND e.fecha_baja<=?"; args.append(f_hasta)
    sql += " ORDER BY e.fecha_baja DESC, e.primer_apellido"
    filas = db.execute(sql, args).fetchall()
    obras_l = obras_visibles(db)
    return render_template("bajas.html", bajas=filas, obras=obras_l,
                           f_obra=f_obra, f_desde=f_desde, f_hasta=f_hasta)


@app.route("/bajas/excel")
@login_required
def bajas_excel():
    db = get_db()
    vis = obras_del_usuario(db)
    sql = ("SELECT e.cedula, e.nombre, e.primer_apellido, e.segundo_apellido, e.nss, e.curp, "
           "e.rfc, e.fecha_alta, e.fecha_baja, e.motivo_baja, p.nombre AS puesto, o.nombre AS obra "
           "FROM empleados e JOIN puestos p ON p.id=e.puesto_id JOIN obras o ON o.id=p.obra_id "
           "WHERE e.estatus='baja'")
    args = []
    if vis is not None:
        if vis:
            sql += " AND o.id IN (%s)" % ",".join("?" * len(vis)); args += vis
        else:
            sql += " AND 0"
    sql += " ORDER BY e.fecha_baja DESC"
    filas = db.execute(sql, args).fetchall()
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    wb = Workbook(); ws = wb.active; ws.title = "Bajas"
    cab = ["CEDULA", "NOMBRE", "NSS", "CURP", "RFC", "PUESTO", "OBRA",
           "FECHA INGRESO", "FECHA BAJA", "MOTIVO"]
    ws.append(cab)
    for c in ws[1]:
        c.font = Font(name="Arial", bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="16233C")
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for e in filas:
        nom = titulo(" ".join(x for x in [e["primer_apellido"], e["segundo_apellido"], e["nombre"]] if x))
        ws.append([e["cedula"], nom, str(e["nss"] or ""), e["curp"], e["rfc"],
                   titulo_obra(e["puesto"] or ""), e["obra"], e["fecha_alta"],
                   e["fecha_baja"], e["motivo_baja"]])
    ws.column_dimensions["B"].width = 30
    for L2 in ("A", "C", "D", "E", "F", "G", "H", "I", "J"):
        ws.column_dimensions[L2].width = 16
    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    return send_file(buf, as_attachment=True, download_name="Reporte_bajas.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# ---------------------------------------------------------------------------
# Cuentas bancarias
# ---------------------------------------------------------------------------
@app.route("/cuentas")
@login_required
def cuentas():
    db = get_db()
    q = request.args.get("q", "").strip()
    sql = ("SELECT c.*, e.nombre, e.primer_apellido, e.segundo_apellido "
           "FROM cuentas_bancarias c JOIN empleados e ON e.id=c.empleado_id "
           "LEFT JOIN puestos p ON p.id=e.puesto_id "
           "LEFT JOIN obras o ON o.id=p.obra_id WHERE 1=1")
    args = []
    if q:
        sql += (" AND (e.nombre LIKE ? OR e.primer_apellido LIKE ? OR "
                "c.numero LIKE ? OR c.institucion LIKE ?)")
        like = f"%{q}%"; args += [like, like, like, like]
    vis = obras_del_usuario(db)
    if vis is not None:
        if vis:
            ph = ",".join("?" * len(vis))
            sql += f" AND o.id IN ({ph})"; args += vis
        else:
            sql += " AND 1=0"
    sql += " ORDER BY e.primer_apellido"
    filas = db.execute(sql, args).fetchall()
    return render_template("cuentas_list.html", cuentas=filas, q=q)


@app.route("/cuenta/<int:cid>/editar", methods=["GET", "POST"])
@min_rank(ADMIN_RANK)
def cuenta_editar(cid):
    db = get_db()
    c = db.execute(
        "SELECT c.*, e.nombre, e.primer_apellido, e.segundo_apellido "
        "FROM cuentas_bancarias c JOIN empleados e ON e.id=c.empleado_id WHERE c.id=?",
        (cid,)).fetchone()
    if not c:
        abort(404)
    if request.method == "POST":
        f = request.form
        inst = f.get("institucion", "").strip()
        tipo = f.get("tipo_cuenta", "").strip()
        numero = solo_digitos(f.get("numero", ""))
        errores = []
        if not numero:
            errores.append("El numero de cuenta es obligatorio.")
        err_cta = validar_cuenta(tipo, numero)
        if err_cta:
            errores.append(err_cta)
        dup = db.execute("SELECT id FROM cuentas_bancarias WHERE numero=? AND id<>?",
                         (numero, cid)).fetchone()
        if dup:
            errores.append("Ese numero de cuenta ya esta registrado en otra cuenta.")
        if errores:
            for e in errores:
                flash(e, "danger")
            datos = dict(c); datos.update({"institucion": inst, "tipo_cuenta": tipo, "numero": numero})
            return render_template("cuenta_form.html", c=datos, bancos=BANCOS, tipos=TIPOS_CUENTA)
        db.execute("UPDATE cuentas_bancarias SET institucion=?, tipo_cuenta=?, numero=? WHERE id=?",
                   (inst, tipo, numero, cid))
        db.commit()
        flash("Cuenta bancaria actualizada.", "success")
        return redirect(url_for("cuentas"))
    return render_template("cuenta_form.html", c=c, bancos=BANCOS, tipos=TIPOS_CUENTA)

# ---------------------------------------------------------------------------
# Asistencia semanal (Fase 2)
# ---------------------------------------------------------------------------
@app.route("/asistencia")
@login_required
def asistencia_inicio():
    db = get_db()
    obras = obras_visibles(db)
    vdef = viernes_de(date.today()).isoformat()
    # semanas capturadas recientes (solo obras visibles)
    filas = db.execute(
        "SELECT s.*, o.nombre AS obra, "
        "  (SELECT COUNT(*) FROM asistencia a WHERE a.semana_id=s.id) AS capturados "
        "FROM asistencia_semanas s JOIN obras o ON o.id=s.obra_id "
        "ORDER BY s.fecha_inicio DESC, o.nombre LIMIT 30").fetchall()
    if obras_del_usuario(db) is not None:
        visibles = {o["id"] for o in obras}
        filas = [f for f in filas if f["obra_id"] in visibles]
    return render_template("asistencia.html", obras=obras,
                           viernes_default=vdef, semanas=filas)


@app.route("/asistencia/captura", methods=["GET", "POST"])
@login_required
def asistencia_captura():
    db = get_db()
    if request.method == "POST":
        obra_id = int(request.form.get("obra_id") or 0)
        finicio = request.form.get("fecha_inicio", "")
    else:
        obra_id = int(request.args.get("obra_id") or 0)
        finicio = request.args.get("fecha_inicio", "")

    if not obra_id or not finicio:
        flash("Selecciona una obra y una semana.", "warning")
        return redirect(url_for("asistencia_inicio"))
    if not puede_ver_obra(db, obra_id):
        abort(403)
    try:
        f = date.fromisoformat(finicio[:10])
    except ValueError:
        flash("Fecha invalida.", "danger")
        return redirect(url_for("asistencia_inicio"))

    viernes = viernes_de(f)
    jueves = viernes + timedelta(days=6)
    obra = db.execute("SELECT * FROM obras WHERE id=?", (obra_id,)).fetchone()
    if not obra:
        flash("La obra no existe.", "danger")
        return redirect(url_for("asistencia_inicio"))

    sem = db.execute(
        "SELECT * FROM asistencia_semanas WHERE obra_id=? AND fecha_inicio=?",
        (obra_id, viernes.isoformat())).fetchone()

    empleados = db.execute(
        "SELECT e.id, e.cedula, e.nombre, e.primer_apellido, e.segundo_apellido, "
        "       p.nombre AS puesto "
        "FROM empleados e JOIN puestos p ON p.id=e.puesto_id "
        "WHERE p.obra_id=? AND e.estatus='activo' "
        "ORDER BY e.primer_apellido, e.nombre", (obra_id,)).fetchall()

    if request.method == "POST":
        if not sem:
            cur = db.execute(
                "INSERT INTO asistencia_semanas(obra_id, fecha_inicio, fecha_fin, "
                "semana_num, anio, estatus, creada_en, creada_por) "
                "VALUES(?,?,?,?,?, 'borrador', ?, ?)",
                (obra_id, viernes.isoformat(), jueves.isoformat(),
                 viernes.isocalendar()[1], viernes.year,
                 datetime.now().isoformat(timespec="seconds"), session.get("nombre")))
            semana_id = cur.lastrowid
        else:
            semana_id = sem["id"]

        for e in empleados:
            eid = e["id"]
            cods = {k: (request.form.get(f"{k}_{eid}") or "").upper()[:1] for k, _ in DIAS_SEMANA}
            def _num(name):
                try:
                    return max(0.0, float(request.form.get(name) or 0))
                except (ValueError, TypeError):
                    return 0.0
            he150 = _num(f"he150_{eid}")
            he200 = _num(f"he200_{eid}")
            obs = (request.form.get(f"obs_{eid}") or "").strip()
            db.execute(
                "INSERT INTO asistencia(semana_id, empleado_id, d1,d2,d3,d4,d5,d6,d7, "
                "he_150, he_200, observaciones) VALUES(?,?,?,?,?,?,?,?,?,?,?,?) "
                "ON CONFLICT(semana_id, empleado_id) DO UPDATE SET "
                "d1=excluded.d1, d2=excluded.d2, d3=excluded.d3, d4=excluded.d4, "
                "d5=excluded.d5, d6=excluded.d6, d7=excluded.d7, "
                "he_150=excluded.he_150, he_200=excluded.he_200, "
                "observaciones=excluded.observaciones",
                (semana_id, eid, cods["d1"], cods["d2"], cods["d3"], cods["d4"],
                 cods["d5"], cods["d6"], cods["d7"], he150, he200, obs))
        db.commit()
        flash("Asistencia guardada.", "success")
        return redirect(url_for("asistencia_captura", obra_id=obra_id,
                                fecha_inicio=viernes.isoformat()))

    # GET: cargar lo ya guardado (si existe) y armar la cuadricula
    guardado = {}
    if sem:
        for r in db.execute("SELECT * FROM asistencia WHERE semana_id=?", (sem["id"],)).fetchall():
            guardado[r["empleado_id"]] = dict(r)

    filas = []
    for e in empleados:
        g = guardado.get(e["id"])
        cods = {}
        for k, _ in DIAS_SEMANA:
            if g:
                cods[k] = g.get(k) or ""
            else:
                cods[k] = "D" if k == "d3" else "A"   # domingo descanso, resto asistio
        dias, faltas, retardos = contar_asistencia(cods)
        filas.append({
            "emp": e, "cods": cods,
            "he150": (g["he_150"] if g else 0) or 0,
            "he200": (g["he_200"] if g else 0) or 0,
            "obs": (g["observaciones"] if g else "") or "",
            "dias": dias, "faltas": faltas, "retardos": retardos,
        })

    return render_template(
        "asistencia_captura.html", obra=obra,
        viernes=viernes.isoformat(), jueves=jueves.isoformat(),
        semana_num=viernes.isocalendar()[1], anio=viernes.year,
        dias_semana=DIAS_SEMANA, dias_base=DIAS_BASE, codigos=CODIGOS_ASIS,
        fechas=fechas_de_dias(viernes), filas=filas,
        ya_existe=bool(sem))


@app.route("/asistencia/plantilla")
@login_required
def asistencia_plantilla():
    """Descarga el Excel de asistencia de una obra con el formato de 3 bloques:
    ASISTENCIA (codigos A/F/R/V/D), HORAS EXTRAS (horas por dia + TIPO + VALOR) y
    RETARDOS (por dia), mas columnas de descuentos. Hoja BAJAS de referencia."""
    db = get_db()
    obra_id = int(request.args.get("obra_id") or 0)
    finicio = request.args.get("fecha_inicio", "")
    if not obra_id or not puede_ver_obra(db, obra_id):
        flash("Selecciona una obra valida.", "warning")
        return redirect(url_for("asistencia_inicio"))
    obra = db.execute("SELECT * FROM obras WHERE id=?", (obra_id,)).fetchone()
    if not obra:
        abort(404)
    try:
        viernes = viernes_de(date.fromisoformat(finicio[:10]))
    except ValueError:
        viernes = viernes_de(date.today())
    jueves = viernes + timedelta(days=6)

    activos = db.execute(
        "SELECT e.cedula, e.nombre, e.primer_apellido, e.segundo_apellido, p.nombre AS puesto "
        "FROM empleados e JOIN puestos p ON p.id=e.puesto_id "
        "WHERE p.obra_id=? AND e.estatus='activo' "
        "ORDER BY e.primer_apellido, e.nombre", (obra_id,)).fetchall()
    bajas = db.execute(
        "SELECT e.cedula, e.nombre, e.primer_apellido, e.segundo_apellido, p.nombre AS puesto, "
        "e.fecha_alta FROM empleados e JOIN puestos p ON p.id=e.puesto_id "
        "WHERE p.obra_id=? AND e.estatus<>'activo' "
        "ORDER BY e.primer_apellido, e.nombre", (obra_id,)).fetchall()

    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.worksheet.datavalidation import DataValidation
    from openpyxl.utils import get_column_letter as L
    NAVY, ROJO, AMAR, AZUL = "16233C", "E1281A", "FFF3CD", "0A56C0"
    thin = Side(style="thin", color="BBBBBB")
    borde = Border(left=thin, right=thin, top=thin, bottom=thin)
    def nom(r):
        return " ".join(x for x in [r["primer_apellido"], r["segundo_apellido"], r["nombre"]] if x)
    def F(**k): return Font(name="Arial", **k)
    def hdr(cell, fill=NAVY):
        cell.font = F(bold=True, size=9, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor=fill)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = borde

    fechas = [f"{(viernes+timedelta(days=i)):%d/%m}" for i in range(7)]
    etqs = ["VIE", "SAB", "DOM", "LUN", "MAR", "MIE", "JUE"]

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Asistencia"
    ws["A1"] = "POLTECH ACERO Y CONSTRUCCION, S.A. DE C.V."
    ws["A1"].font = F(bold=True, size=13, color=NAVY)
    ws["A2"] = "CONTROL SEMANAL DE ASISTENCIA  -  Semana laboral: Viernes a Jueves"
    ws["A2"].font = F(size=10, italic=True, color="555555")
    for lc, lbl, vc, val in [
        ("A4", "OBRA:", "C4", obra["nombre"]),
        ("A5", "ESTADO:", "C5", obra["estado"]),
        ("H4", "SEMANA No.:", "J4", viernes.isocalendar()[1]),
        ("H5", "PERIODO:", "J5", f"{viernes:%d/%m/%Y} al {jueves:%d/%m/%Y}"),
        ("L4", "GERENTE DE OBRA:", "N4", ""),
        ("L5", "FECHA DE CAPTURA:", "N5", ""),
    ]:
        ws[lc] = lbl; ws[lc].font = F(bold=True, size=9, color=NAVY)
        ws[vc] = val; ws[vc].font = F(size=9)
        if val == "":
            ws[vc].fill = PatternFill("solid", fgColor=AMAR)

    # Fila 6: titulos de bloque (combinados)
    ws.merge_cells("D6:O6"); ws["D6"] = "ASISTENCIA"
    ws.merge_cells("P6:Y6"); ws["P6"] = "HORAS EXTRAS"
    ws.merge_cells("Z6:AF6"); ws["Z6"] = "RETARDOS"
    for cc, fill in (("D6", NAVY), ("P6", AZUL), ("Z6", ROJO)):
        ws[cc].font = F(bold=True, size=10, color="FFFFFF")
        ws[cc].fill = PatternFill("solid", fgColor=fill)
        ws[cc].alignment = Alignment(horizontal="center", vertical="center")

    # Fila 7: encabezados
    HROW = 7
    cab = ["No.", "CEDULA", "NOMBRE DEL TRABAJADOR"]
    cab += [f"{etqs[i]}\n{fechas[i]}" for i in range(7)]                 # D-J asistencia
    cab += ["DIAS", "FALTAS", "RET.", "VAC.", "BAJA\n(Ultimo dia trabajado)"]  # K-O
    cab += [f"{etqs[i]}\n{fechas[i]}" for i in range(7)]                 # P-V horas extra
    cab += ["TOTAL\nHORAS EXTRAS", "TIPO", "VALOR"]                      # W-Y
    cab += [f"{etqs[i]}\n{fechas[i]}" for i in range(7)]                 # Z-AF retardos
    cab += ["DESCUENTOS\nNOMINA", "DESCUENTOS A\nOTRA CUENTA"]           # AG-AH
    cab += ["OBSERVACIONES\n(a quien se deposita el desc. a otra cuenta)"]  # AI
    for i, t in enumerate(cab, start=1):
        hdr(ws.cell(HROW, i, t))

    r = HROW + 1
    for idx, emp in enumerate(activos, start=1):
        ws.cell(r, 1, idx)
        ws.cell(r, 2, emp["cedula"] or "")
        ws.cell(r, 3, nom(emp))
        # Asistencia (D-J): base=A, domingo(F col6)=D
        for col in (4, 5, 7, 8, 9, 10):
            ws.cell(r, col, "A")
        ws.cell(r, 6, "D")
        # Formulas de conteo
        ws.cell(r, 11).value = (f'=COUNTIF(D{r}:E{r},"A")+COUNTIF(G{r}:J{r},"A")'
                                f'+COUNTIF(D{r}:E{r},"R")+COUNTIF(G{r}:J{r},"R")'
                                f'+COUNTIF(D{r}:E{r},"V")+COUNTIF(G{r}:J{r},"V")')   # DIAS
        ws.cell(r, 12).value = f'=COUNTIF(D{r}:E{r},"F")+COUNTIF(G{r}:J{r},"F")'     # FALTAS
        ws.cell(r, 13).value = f'=SUM(Z{r}:AF{r})'                                   # RET.
        ws.cell(r, 14).value = f'=COUNTIF(D{r}:E{r},"V")+COUNTIF(G{r}:J{r},"V")'     # VAC
        ws.cell(r, 15, "")                                                          # BAJA
        # Horas extra (P-V) = 0, TOTAL, TIPO, VALOR
        for col in range(16, 23):
            ws.cell(r, col, 0)
        ws.cell(r, 23).value = f"=SUM(P{r}:V{r})"                                    # TOTAL HE
        ws.cell(r, 24, "")                                                          # TIPO
        ws.cell(r, 25, "")                                                          # VALOR
        # Retardos (Z-AF) = 0
        for col in range(26, 33):
            ws.cell(r, col, 0)
        ws.cell(r, 33, 0)                                                           # DESC NOMINA
        ws.cell(r, 34, 0)                                                           # DESC OTRA CUENTA
        ws.cell(r, 35, "")                                                          # OBSERVACIONES
        for col in range(1, 36):
            cell = ws.cell(r, col)
            cell.font = F(size=9); cell.border = borde
            if 4 <= col <= 34: cell.alignment = Alignment(horizontal="center")
        ws.cell(r, 35).fill = PatternFill("solid", fgColor=AMAR)
        ws.cell(r, 15).fill = PatternFill("solid", fgColor=AMAR)
        r += 1
    ult = r - 1

    if activos:
        def dv(kind, f1, rng, prompt=None, title=None):
            d = DataValidation(type=kind, formula1=f1, allow_blank=True)
            if kind == "decimal": d.operator = "greaterThanOrEqual"
            if prompt: d.prompt = prompt
            if title: d.promptTitle = title
            ws.add_data_validation(d); d.add(rng)
        dv("list", '"A,F,R,V,D"', f"D{HROW+1}:J{ult}",
           "A=Asistio  F=Falta  R=Retardo  V=Vacaciones  D=Descanso", "Asistencia")
        dv("decimal", "0", f"P{HROW+1}:V{ult}")                 # horas extra
        dv("list", '"Factor,Precio"', f"X{HROW+1}:X{ult}",
           "Factor = 1.5 o 2 veces el sueldo por hora.  Precio = precio por hora pactado.", "Tipo H.E.")
        dv("decimal", "0", f"Y{HROW+1}:Y{ult}")                 # valor
        dv("list", '"0,1"', f"Z{HROW+1}:AF{ult}",
           "1 = ese dia tuvo retardo.  0 = sin retardo.", "Retardos")
        dv("decimal", "0", f"AG{HROW+1}:AH{ult}")               # descuentos

    anchos = ([5, 10, 30] + [6]*7 + [7, 8, 6, 6, 12] + [6]*7 + [9, 9, 8] + [6]*7 + [12, 13, 32])
    for i, w in enumerate(anchos, start=1):
        ws.column_dimensions[L(i)].width = w
    ws.freeze_panes = "D8"

    lr = ult + 2
    ws.cell(lr, 3, "Codigos de ASISTENCIA:  A = Asistio    F = Falta    R = Retardo    V = Vacaciones    D = Descanso").font = F(size=9, bold=True, color=NAVY)
    ws.cell(lr+1, 3, "RETARDOS: escribe 1 en el dia que llego tarde.  HORAS EXTRAS: horas por dia + TIPO (Factor/Precio) y VALOR.").font = F(size=8, italic=True, color="555555")
    ws.cell(lr+2, 3, "BAJA: escribe la fecha del ultimo dia trabajado si el trabajador causa baja.  El DOMINGO no cuenta para la base de 6 dias.").font = F(size=8, italic=True, color="555555")

    # ---- Hoja BAJAS (referencia) ----
    wb2 = wb.create_sheet("Bajas")
    wb2.append(["No.", "CEDULA", "NOMBRE DEL TRABAJADOR", "PUESTO", "FECHA DE ALTA"])
    for c in wb2[1]:
        c.font = F(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor=ROJO)
    for idx, emp in enumerate(bajas, start=1):
        wb2.append([idx, emp["cedula"] or "", nom(emp), emp["puesto"], emp["fecha_alta"] or ""])
    for i, w in enumerate([5, 10, 30, 18, 14], start=1):
        wb2.column_dimensions[L(i)].width = w

    # ---- Hoja INSTRUCCIONES ----
    ins = wb.create_sheet("Instrucciones")
    lineas = [
        ("INSTRUCCIONES DE USO", True),
        ("", False),
        ("1. Solo se captura la hoja ASISTENCIA. La hoja BAJAS es solo de referencia.", False),
        ("2. Completa el encabezado (celdas amarillas): Gerente de obra y Fecha de captura.", False),
        ("3. BLOQUE ASISTENCIA: en cada dia escribe el codigo", False),
        ("     A = Asistio     F = Falta     R = Retardo     V = Vacaciones     D = Descanso", False),
        ("   Cada dia trae su fecha (VIE 24/07, etc.). El DOMINGO no cuenta para la base de 6 dias.", False),
        ("4. DIAS, FALTAS, RET. y VAC. se calculan solos. No los edites.", False),
        ("   DIAS pagables = dias con A, R o V (el retardo y la vacacion cuentan como dia).", False),
        ("5. BLOQUE HORAS EXTRAS: escribe las horas por dia. TOTAL se suma solo.", False),
        ("   TIPO = 'Factor' (1.5 o 2 veces el sueldo por hora) o 'Precio' (precio por hora pactado).", False),
        ("   VALOR = el factor (1.5 / 2) o el precio por hora en pesos.", False),
        ("6. BLOQUE RETARDOS: escribe 1 en el dia que el trabajador llego tarde (RET. los suma).", False),
        ("7. BAJA (ultimo dia trabajado): si ya no regreso, escribe la fecha de su ultimo dia.", False),
        ("   Con eso el sistema lo marca de baja y prepara el aviso para el IMSS.", False),
        ("8. DESCUENTOS NOMINA: importe a descontar (prestamo, herramienta, etc.).", False),
        ("   DESCUENTOS A OTRA CUENTA: importe que se le quita para depositar a otra persona.", False),
        ("9. Guarda el archivo y subelo en el sistema (Asistencia semanal > Calcular nomina).", False),
    ]
    for i, (txt, boldl) in enumerate(lineas, start=1):
        c = ins.cell(i, 1, txt)
        c.font = F(bold=boldl, size=(12 if boldl else 10), color=(NAVY if boldl else "333333"))
    ins.column_dimensions["A"].width = 108

    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    obra_limpia = re.sub(r"[^0-9A-Za-zÁÉÍÓÚÑáéíóúñ ._-]+", "", obra["nombre"]).strip()
    nombre = f"Lista de asistencia semana {viernes.isocalendar()[1]} - {obra_limpia}.xlsx"
    return send_file(buf, as_attachment=True, download_name=nombre,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# ---------------------------------------------------------------------------
# Nomina: subir el Excel de asistencia lleno y calcular
# ---------------------------------------------------------------------------
DIAS_BASE_COLS = [4, 5, 7, 8, 9, 10]   # D,E,G,H,I,J (sin domingo=F/col6)

def _num(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return 0.0

def _cod(v):
    return (str(v).strip().upper()[:1]) if v is not None else ""

def calcular_nomina(db, obra_id, ws, jornada, aplicar_retardos):
    """Lee la hoja 'Asistencia' (formato de 3 bloques) y calcula la nomina.
    Devuelve (detalles, avisos_baja, errores)."""
    detalles, avisos_baja, errores = [], [], []
    HROW = 7
    r = HROW + 1
    while True:
        ced = ws.cell(r, 2).value
        nombre = ws.cell(r, 3).value
        if (ced is None or str(ced).strip() == "") and (nombre is None or str(nombre).strip() == ""):
            break
        cedula = str(ced).strip() if ced else ""
        # buscar empleado por cedula en la obra
        emp = db.execute(
            "SELECT e.id, e.nombre, e.primer_apellido, e.segundo_apellido, e.curp, e.rfc, "
            "e.nss, e.infonavit_monto, e.viaticos_semanales AS viat_emp, e.bono_semanal AS bono_emp, "
            "p.nombre AS puesto, p.clasificacion, "
            "p.sueldo_semanal, p.viaticos_semanales "
            "FROM empleados e JOIN puestos p ON p.id=e.puesto_id "
            "WHERE e.cedula=? AND p.obra_id=?", (cedula, obra_id)).fetchone()
        if not emp:
            errores.append(f"Cedula '{cedula}' ({nombre}) no se encontro en esta obra; se omitio.")
            r += 1
            continue

        codigos = [_cod(ws.cell(r, c).value) for c in DIAS_BASE_COLS]
        dias_A = codigos.count("A")
        dias_R = codigos.count("R")
        dias_V = codigos.count("V")
        faltas = codigos.count("F")
        dias_pagables = dias_A + dias_R + dias_V
        dias_presente = dias_A + dias_R

        he_horas = sum(_num(ws.cell(r, c).value) for c in range(16, 23))   # P-V
        tipo = (str(ws.cell(r, 24).value or "").strip().lower())
        valor = _num(ws.cell(r, 25).value)
        retardos = sum(_num(ws.cell(r, c).value) for c in range(26, 33))   # Z-AF
        desc_nomina = _num(ws.cell(r, 33).value)
        desc_otra = _num(ws.cell(r, 34).value)
        obs = str(ws.cell(r, 35).value or "").strip()   # a quien se deposita el desc. a otra cuenta
        baja_fecha = ws.cell(r, 15).value
        baja_fecha = str(baja_fecha).strip() if baja_fecha not in (None, "") else ""

        sueldo_semanal = _num(emp["sueldo_semanal"])
        # viaticos: si el trabajador tiene su propio valor se usa; si no, el del puesto
        viaticos_semanal = _num(emp["viat_emp"]) if emp["viat_emp"] is not None else _num(emp["viaticos_semanales"])
        bono_semanal = _num(emp["bono_emp"])
        sueldo_diario = sueldo_semanal / 6.0
        viatico_diario = viaticos_semanal / 6.0
        bono_diario = bono_semanal / 6.0

        sueldo = round(sueldo_diario * dias_pagables, 2)
        viaticos = round(viatico_diario * dias_presente, 2)
        bono = round(bono_diario * dias_presente, 2)
        if tipo == "factor":
            he_importe = round(he_horas * (sueldo_diario / jornada) * valor, 2)
        elif tipo == "precio":
            he_importe = round(he_horas * valor, 2)
        else:
            he_importe = 0.0
        infonavit = _num(emp["infonavit_monto"])
        desc_retardos = 0.0
        if aplicar_retardos and retardos >= 3:
            desc_retardos = round((int(retardos) // 3) * sueldo_diario, 2)
        neto = round(sueldo + viaticos + bono + he_importe - infonavit - desc_nomina - desc_otra - desc_retardos, 2)

        nom_full = " ".join(x for x in [emp["primer_apellido"], emp["segundo_apellido"], emp["nombre"]] if x)
        detalles.append({
            "empleado_id": emp["id"], "cedula": cedula, "nombre": nom_full,
            "puesto": emp["puesto"], "clasificacion": emp["clasificacion"] or "Sin clasificar",
            "sueldo_contratado": round(sueldo_semanal, 2), "viaticos_contratado": round(viaticos_semanal, 2),
            "dias": dias_pagables, "faltas": faltas,
            "retardos": retardos, "vacaciones": dias_V, "sueldo": sueldo,
            "viaticos": viaticos, "bono": bono, "he_horas": he_horas, "he_importe": he_importe,
            "infonavit": infonavit, "desc_nomina": desc_nomina, "desc_otra": desc_otra,
            "desc_retardos": desc_retardos, "neto": neto, "baja_fecha": baja_fecha, "nota": obs,
        })
        if baja_fecha:
            avisos_baja.append({
                "empleado_id": emp["id"], "cedula": cedula, "nombre": nom_full,
                "curp": emp["curp"], "rfc": emp["rfc"], "nss": emp["nss"],
                "puesto": emp["puesto"], "baja_fecha": baja_fecha,
            })
        r += 1
    return detalles, avisos_baja, errores


@app.route("/nomina", methods=["GET", "POST"])
@login_required
def nomina():
    db = get_db()
    obras = obras_visibles(db)
    if request.method == "POST":
        obra_id = int(request.form.get("obra_id") or 0)
        finicio = request.form.get("fecha_inicio", "")
        aplicar_retardos = 1 if request.form.get("aplicar_retardos") else 0
        archivo = request.files.get("archivo")
        if not obra_id or not puede_ver_obra(db, obra_id):
            flash("Selecciona una obra valida.", "warning")
            return redirect(url_for("nomina"))
        if not archivo or not archivo.filename.lower().endswith(".xlsx"):
            flash("Sube el Excel de asistencia (.xlsx) que descargaste del sistema.", "danger")
            return redirect(url_for("nomina"))
        try:
            viernes = viernes_de(date.fromisoformat(finicio[:10]))
        except ValueError:
            viernes = viernes_de(date.today())
        jueves = viernes + timedelta(days=6)
        try:
            jornada = float(get_param(db, "jornada_horas", "8") or 8) or 8
        except (ValueError, TypeError):
            jornada = 8.0

        wb = openpyxl.load_workbook(archivo, data_only=True)
        ws = wb["Asistencia"] if "Asistencia" in wb.sheetnames else wb.active
        detalles, avisos_baja, errores = calcular_nomina(db, obra_id, ws, jornada, aplicar_retardos)
        if not detalles:
            for e in errores:
                flash(e, "warning")
            flash("No se pudo calcular: revisa que el archivo tenga trabajadores y el formato correcto.", "danger")
            return redirect(url_for("nomina"))

        total = round(sum(d["neto"] for d in detalles), 2)
        cur = db.execute(
            "INSERT INTO nominas(obra_id, fecha_inicio, fecha_fin, semana_num, anio, jornada, "
            "aplico_retardos, total_neto, creada_en, creada_por) VALUES(?,?,?,?,?,?,?,?,?,?)",
            (obra_id, viernes.isoformat(), jueves.isoformat(), viernes.isocalendar()[1],
             viernes.year, jornada, aplicar_retardos, total,
             datetime.now().isoformat(timespec="seconds"), session.get("nombre")))
        nomina_id = cur.lastrowid
        for d in detalles:
            db.execute(
                "INSERT INTO nomina_detalle(nomina_id, empleado_id, cedula, nombre, clasificacion, "
                "sueldo_contratado, viaticos_contratado, dias, faltas, "
                "retardos, vacaciones, sueldo, viaticos, bono, he_horas, he_importe, infonavit, "
                "desc_nomina, desc_otra, desc_retardos, neto, baja_fecha, nota) "
                "VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                (nomina_id, d["empleado_id"], d["cedula"], d["nombre"], d["clasificacion"],
                 d["sueldo_contratado"], d["viaticos_contratado"], d["dias"], d["faltas"],
                 d["retardos"], d["vacaciones"], d["sueldo"], d["viaticos"], d["bono"], d["he_horas"],
                 d["he_importe"], d["infonavit"], d["desc_nomina"], d["desc_otra"],
                 d["desc_retardos"], d["neto"], d["baja_fecha"], d["nota"]))
        # marcar bajas reportadas
        for b in avisos_baja:
            db.execute("UPDATE empleados SET estatus='baja', fecha_baja=COALESCE(fecha_baja, ?) WHERE id=?",
                       (b["baja_fecha"], b["empleado_id"]))
            registrar_bitacora(db, "Baja de trabajador",
                               f"{b['nombre']} (cedula {b['cedula']}, NSS {b['nss']}) baja desde {b['baja_fecha']}")
            send_admin_alert(
                "POLTECH - Aviso de baja de trabajador",
                (f"El trabajador {b['nombre']} causa baja desde {b['baja_fecha']}.\n"
                 f"CURP: {b['curp']}  RFC: {b['rfc']}  NSS: {b['nss']}  Puesto: {b['puesto']}.\n\n"
                 f"Favor de tramitar la baja ante el IMSS."),
                extra=[get_param(db, "email_contador", "")])
        registrar_bitacora(db, "Calculo de nomina",
                           f"Obra {obra_id}, semana {viernes.isoformat()}, {len(detalles)} trabajadores, total {total:.2f}")
        db.commit()
        for e in errores:
            flash(e, "warning")
        return redirect(url_for("nomina_resultado", nomina_id=nomina_id))

    vdef = viernes_de(date.today()).isoformat()
    recientes = db.execute(
        "SELECT n.*, o.nombre AS obra FROM nominas n JOIN obras o ON o.id=n.obra_id "
        "ORDER BY n.creada_en DESC LIMIT 15").fetchall()
    if obras_del_usuario(db) is not None:
        vis = {o["id"] for o in obras}
        recientes = [n for n in recientes if n["obra_id"] in vis]
    return render_template("nomina.html", obras=obras, viernes_default=vdef, recientes=recientes)


def _con_contratado(db, det):
    """Convierte el detalle a dicts y, si el sueldo/viaticos contratado viene en 0
    (nominas calculadas antes de esta funcion), lo toma del catalogo actual del puesto."""
    filas = [dict(d) for d in det]
    faltan = [d for d in filas if not _num(d.get("sueldo_contratado"))
              and not _num(d.get("viaticos_contratado")) and d.get("empleado_id")]
    if faltan:
        ids = list({d["empleado_id"] for d in faltan})
        ph = ",".join("?" * len(ids))
        cat = {r["id"]: r for r in db.execute(
            f"SELECT e.id, p.sueldo_semanal, p.viaticos_semanales FROM empleados e "
            f"JOIN puestos p ON p.id=e.puesto_id WHERE e.id IN ({ph})", ids).fetchall()}
        for d in faltan:
            r = cat.get(d["empleado_id"])
            if r:
                d["sueldo_contratado"] = _num(r["sueldo_semanal"])
                d["viaticos_contratado"] = _num(r["viaticos_semanales"])
    return filas


@app.route("/nomina/<int:nomina_id>")
@login_required
def nomina_resultado(nomina_id):
    db = get_db()
    n = db.execute("SELECT n.*, o.nombre AS obra, o.estado FROM nominas n "
                   "JOIN obras o ON o.id=n.obra_id WHERE n.id=?", (nomina_id,)).fetchone()
    if not n:
        abort(404)
    if not puede_ver_obra(db, n["obra_id"]):
        abort(403)
    det = _con_contratado(db, db.execute(
        "SELECT * FROM nomina_detalle WHERE nomina_id=? ORDER BY nombre", (nomina_id,)).fetchall())
    tot = {k: sum(_num(d[k]) for d in det) for k in
           ("sueldo", "viaticos", "bono", "he_importe", "infonavit", "desc_nomina",
            "desc_otra", "desc_retardos", "neto")}
    bajas = [d for d in det if d["baja_fecha"]]
    # totales por clasificacion (cuanto se paga de cada grupo)
    por_clasif = {}
    for d in det:
        k = d["clasificacion"] or "Sin clasificar"
        g = por_clasif.setdefault(k, {"num": 0, "neto": 0.0})
        g["num"] += 1
        g["neto"] += _num(d["neto"])
    por_clasif = dict(sorted(por_clasif.items(), key=lambda kv: kv[1]["neto"], reverse=True))
    # Nomina real vs. lo que sale de caja
    caja_rows = []
    caja = {"devengado": 0.0, "infonavit": 0.0, "desc_nomina": 0.0, "desc_retardos": 0.0,
            "se_queda": 0.0, "traspaso": 0.0, "neto": 0.0, "sale": 0.0}
    for d in det:
        devengado = _num(d["sueldo"]) + _num(d["viaticos"]) + _num(d.get("bono")) + _num(d["he_importe"])
        se_queda = _num(d["infonavit"]) + _num(d["desc_nomina"]) + _num(d["desc_retardos"])
        traspaso = _num(d["desc_otra"])
        neto = _num(d["neto"])
        sale = round(neto + traspaso, 2)
        caja_rows.append({"nombre": d["nombre"], "cedula": d["cedula"],
                          "devengado": round(devengado, 2), "se_queda": round(se_queda, 2),
                          "traspaso": traspaso, "neto": neto, "sale": sale, "nota": d["nota"]})
        caja["devengado"] += devengado
        caja["infonavit"] += _num(d["infonavit"])
        caja["desc_nomina"] += _num(d["desc_nomina"])
        caja["desc_retardos"] += _num(d["desc_retardos"])
        caja["se_queda"] += se_queda
        caja["traspaso"] += traspaso
        caja["neto"] += neto
        caja["sale"] += sale
    caja = {k: round(v, 2) for k, v in caja.items()}
    pct_despacho = _num(get_param(db, "porcentaje_despacho", "0"))
    despacho = round(caja["neto"] * pct_despacho / 100.0, 2)
    total_a_pagar = round(caja["neto"] + despacho, 2)
    return render_template("nomina_resultado.html", n=n, det=det, tot=tot,
                           bajas=bajas, por_clasif=por_clasif, caja=caja, caja_rows=caja_rows,
                           pct_despacho=pct_despacho, despacho=despacho, total_a_pagar=total_a_pagar)


@app.route("/nomina/<int:nomina_id>/eliminar", methods=["POST"])
@min_rank(GERENTE_RANK)
def nomina_eliminar(nomina_id):
    db = get_db()
    n = db.execute("SELECT * FROM nominas WHERE id=?", (nomina_id,)).fetchone()
    if not n:
        abort(404)
    if not puede_ver_obra(db, n["obra_id"]):
        abort(403)
    db.execute("DELETE FROM nomina_detalle WHERE nomina_id=?", (nomina_id,))
    db.execute("DELETE FROM nominas WHERE id=?", (nomina_id,))
    registrar_bitacora(db, "Eliminacion de nomina",
                       f"Nomina {nomina_id} (obra {n['obra_id']}, semana {n['fecha_inicio']})")
    db.commit()
    flash("Nomina eliminada.", "success")
    return redirect(url_for("nomina"))


@app.route("/nomina/<int:nomina_id>/autorizar", methods=["POST"])
@min_rank(GERENTE_RANK)
def nomina_autorizar(nomina_id):
    db = get_db()
    n = db.execute("SELECT * FROM nominas WHERE id=?", (nomina_id,)).fetchone()
    if not n:
        abort(404)
    if not puede_ver_obra(db, n["obra_id"]):
        abort(403)
    if (n["estatus"] or "pendiente") == "autorizada":
        flash("Esta nomina ya estaba autorizada.", "info")
        return redirect(url_for("nomina_resultado", nomina_id=nomina_id))
    db.execute("UPDATE nominas SET estatus='autorizada', autorizada_por=?, autorizada_en=? WHERE id=?",
               (session.get("nombre"), datetime.now().isoformat(timespec="seconds"), nomina_id))
    registrar_bitacora(db, "Autorizacion de nomina",
                       f"Nomina {nomina_id} (obra {n['obra_id']}, semana {n['fecha_inicio']}) autorizada y liberada")
    db.commit()
    # Avisar por correo a residentes, superintendentes y administradores de la obra
    obra = db.execute("SELECT nombre FROM obras WHERE id=?", (n["obra_id"],)).fetchone()
    correos = [r["username"] for r in db.execute(
        "SELECT DISTINCT u.username FROM users u "
        "LEFT JOIN user_obras uo ON uo.user_id=u.id "
        "WHERE u.role='admin' OR (u.role IN ('residente','superintendente') AND uo.obra_id=?)",
        (n["obra_id"],)).fetchall()]
    correos.append(os.environ.get("ADMIN_EMAIL", ""))
    enviar_correo(
        f"POLTECH - Nomina liberada: {obra['nombre'] if obra else ''} semana {n['semana_num']}",
        (f"La nomina de la obra {obra['nombre'] if obra else n['obra_id']} "
         f"(semana {n['semana_num']}/{n['anio']}, del {n['fecha_inicio']} al {n['fecha_fin']}) "
         f"fue autorizada y liberada por {session.get('nombre')}.\n"
         f"Total neto: {n['total_neto']:.2f}."),
        correos)
    flash("Nomina autorizada y liberada.", "success")
    return redirect(url_for("nomina_resultado", nomina_id=nomina_id))


@app.route("/nomina/<int:nomina_id>/reabrir", methods=["POST"])
@min_rank(ADMIN_RANK)
def nomina_reabrir(nomina_id):
    db = get_db()
    n = db.execute("SELECT * FROM nominas WHERE id=?", (nomina_id,)).fetchone()
    if not n:
        abort(404)
    db.execute("UPDATE nominas SET estatus='pendiente', autorizada_por=NULL, autorizada_en=NULL WHERE id=?",
               (nomina_id,))
    registrar_bitacora(db, "Reapertura de nomina",
                       f"Nomina {nomina_id} (obra {n['obra_id']}, semana {n['fecha_inicio']}) regresada a pendiente")
    db.commit()
    flash("Se quito la autorizacion; la nomina quedo pendiente de nuevo.", "success")
    return redirect(url_for("nomina_resultado", nomina_id=nomina_id))


@app.route("/nomina/<int:nomina_id>/excel")
@login_required
def nomina_excel(nomina_id):
    db = get_db()
    n = db.execute("SELECT n.*, o.nombre AS obra FROM nominas n JOIN obras o ON o.id=n.obra_id "
                   "WHERE n.id=?", (nomina_id,)).fetchone()
    if not n:
        abort(404)
    if not puede_ver_obra(db, n["obra_id"]):
        abort(403)
    det = _con_contratado(db, db.execute(
        "SELECT * FROM nomina_detalle WHERE nomina_id=? ORDER BY nombre", (nomina_id,)).fetchall())
    # datos bancarios por empleado (para el pago)
    cuentas = {}
    ids = [d["empleado_id"] for d in det if d["empleado_id"]]
    if ids:
        ph = ",".join("?" * len(ids))
        for c in db.execute(
                f"SELECT empleado_id, institucion, tipo_cuenta, numero FROM cuentas_bancarias "
                f"WHERE empleado_id IN ({ph})", ids).fetchall():
            cuentas.setdefault(c["empleado_id"], c)   # la primera cuenta del trabajador

    from openpyxl.styles import Font, PatternFill, Alignment
    NAVY = "16233C"
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Nomina"
    ws["A1"] = "POLTECH ACERO Y CONSTRUCCION, S.A. DE C.V."
    ws["A1"].font = Font(name="Arial", bold=True, size=13, color=NAVY)
    ws["A2"] = f"NOMINA SEMANAL  -  {n['obra']}  -  {n['fecha_inicio']} al {n['fecha_fin']}"
    ws["A2"].font = Font(name="Arial", size=10, italic=True, color="555555")
    if (n["estatus"] or "pendiente") == "autorizada":
        ws["A3"] = f"AUTORIZADA Y LIBERADA por {n['autorizada_por']} el {n['autorizada_en']}"
        ws["A3"].font = Font(name="Arial", size=9, bold=True, color="1E7E34")
    else:
        ws["A3"] = "PENDIENTE DE AUTORIZAR"
        ws["A3"].font = Font(name="Arial", size=9, bold=True, color="B8860B")
    cab = ["No.", "CEDULA", "NOMBRE", "DIAS", "FALTAS", "RET.", "VAC.", "SUELDO",
           "VIATICOS", "TOTAL HORAS EXTRAS", "IMPORTE HORAS EXTRAS", "INFONAVIT", "DESC. NOMINA",
           "DESC. OTRA", "DESC. RET.", "NETO A PAGAR", "BAJA",
           "TIPO CUENTA", "BANCO", "No. CUENTA", "CLASIFICACION",
           "SUELDO CONTRATADO", "VIATICOS CONTRATADO", "DEPOSITO A (obs)", "BONO"]
    HROW = 4
    for i, t in enumerate(cab, start=1):
        c = ws.cell(HROW, i, t)
        c.font = Font(name="Arial", bold=True, size=9, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=NAVY)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    r = HROW + 1
    for idx, d in enumerate(det, start=1):
        cta = cuentas.get(d["empleado_id"])
        tipo_cta = cta["tipo_cuenta"] if cta else ""
        banco_cta = cta["institucion"] if cta else ""
        num_cta = str(cta["numero"]) if cta and cta["numero"] is not None else ""
        ws.append([idx, d["cedula"], d["nombre"], d["dias"], d["faltas"], d["retardos"],
                   d["vacaciones"], d["sueldo"], d["viaticos"], d["he_horas"], d["he_importe"],
                   d["infonavit"], d["desc_nomina"], d["desc_otra"], d["desc_retardos"],
                   d["neto"], d["baja_fecha"] or "", tipo_cta, banco_cta, num_cta,
                   d["clasificacion"] or "Sin clasificar",
                   _num(d["sueldo_contratado"]), _num(d["viaticos_contratado"]), d["nota"] or "",
                   _num(d["bono"])])
        ws.cell(r, 20).number_format = "@"   # No. de cuenta como texto
        r += 1
    # totales generales
    ws.append([])
    fila_tot = ["", "", "TOTALES", "", "", "", ""]
    for col in ("sueldo", "viaticos", "he_importe", "infonavit", "desc_nomina", "desc_otra", "desc_retardos", "neto"):
        fila_tot.append(round(sum(_num(d[col]) for d in det), 2))
    fila_tot.append("")
    ws.append(fila_tot)
    for col in list(range(8, 17)) + [22, 23, 25]:
        for row in range(HROW + 1, r + 2):
            cell = ws.cell(row, col)
            if isinstance(cell.value, (int, float)):
                cell.number_format = '#,##0.00'
    ws.column_dimensions["C"].width = 30
    for i in [2] + list(range(4, 26)):
        Lc = openpyxl.utils.get_column_letter(i)
        ws.column_dimensions[Lc].width = max(ws.column_dimensions[Lc].width or 8, 12)

    # ---- Resumen por clasificacion ----
    por = {}
    for d in det:
        k = d["clasificacion"] or "Sin clasificar"
        g = por.setdefault(k, {"num": 0, "neto": 0.0})
        g["num"] += 1; g["neto"] += _num(d["neto"])
    rs = wb.create_sheet("Por clasificacion")
    rs.append(["CLASIFICACION", "TRABAJADORES", "NETO PAGADO"])
    for c in rs[1]:
        c.font = Font(name="Arial", bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor=NAVY)
    for k, g in sorted(por.items(), key=lambda kv: kv[1]["neto"], reverse=True):
        rs.append([k, g["num"], round(g["neto"], 2)])
    rs.append(["TOTAL", sum(g["num"] for g in por.values()), round(sum(g["neto"] for g in por.values()), 2)])
    for row in range(2, rs.max_row + 1):
        rs.cell(row, 3).number_format = '#,##0.00'
    rs.column_dimensions["A"].width = 24; rs.column_dimensions["B"].width = 14; rs.column_dimensions["C"].width = 16

    # ---- Hoja: Flujo de caja (nomina real vs. lo que sale) ----
    cj = wb.create_sheet("Flujo de caja")
    cab_cj = ["NOMBRE", "NOMINA REAL", "SE QUEDA (INFONAVIT+DESC+RET)",
              "TRASPASO A 3RO", "NETO AL TRABAJADOR", "SALE DE CAJA", "DEPOSITO A"]
    cj.append(cab_cj)
    for c in cj[1]:
        c.font = Font(name="Arial", bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor=NAVY)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    tot_cj = {"dev": 0.0, "sq": 0.0, "tr": 0.0, "net": 0.0, "sal": 0.0}
    for d in det:
        dev = _num(d["sueldo"]) + _num(d["viaticos"]) + _num(d["bono"]) + _num(d["he_importe"])
        sq = _num(d["infonavit"]) + _num(d["desc_nomina"]) + _num(d["desc_retardos"])
        tr = _num(d["desc_otra"]); net = _num(d["neto"]); sal = round(net + tr, 2)
        cj.append([d["nombre"], round(dev, 2), round(sq, 2), tr, net, sal, d["nota"] or ""])
        tot_cj["dev"] += dev; tot_cj["sq"] += sq; tot_cj["tr"] += tr
        tot_cj["net"] += net; tot_cj["sal"] += sal
    cj.append(["TOTALES", round(tot_cj["dev"], 2), round(tot_cj["sq"], 2), round(tot_cj["tr"], 2),
               round(tot_cj["net"], 2), round(tot_cj["sal"], 2), ""])
    fila_tot = cj.max_row
    for c in cj[fila_tot]:
        c.font = Font(name="Arial", bold=True)
    # Despacho (% sobre el neto) y Total a pagar
    pct = _num(get_param(db, "porcentaje_despacho", "0"))
    despacho = round(tot_cj["net"] * pct / 100.0, 2)
    cj.append([f"DESPACHO ({pct:g}%)", "", "", "", despacho, "", ""])
    cj.append(["TOTAL A PAGAR", "", "", "", round(tot_cj["net"] + despacho, 2), "", ""])
    for rr in (cj.max_row - 1, cj.max_row):
        cj.cell(rr, 1).font = Font(name="Arial", bold=True)
        cj.cell(rr, 5).font = Font(name="Arial", bold=True)
    for row in range(2, cj.max_row + 1):
        for col in range(2, 7):
            cj.cell(row, col).number_format = '#,##0.00'
    cj.column_dimensions["A"].width = 28
    for L2 in ("B", "C", "D", "E", "F"):
        cj.column_dimensions[L2].width = 16
    cj.column_dimensions["G"].width = 22

    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    nombre = f"Nomina_{re.sub(r'[^A-Za-z0-9]+','_',n['obra']).strip('_')}_{n['fecha_inicio']}.xlsx"
    return send_file(buf, as_attachment=True, download_name=nombre,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


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

        errores = []
        if not u: errores.append("El usuario (correo) es obligatorio.")
        if not nombre: errores.append("El nombre es obligatorio.")
        if len(pw) < 6: errores.append("La contrasena debe tener al menos 6 caracteres.")
        if role not in ROLES: errores.append("Selecciona un rol valido.")
        if db.execute("SELECT 1 FROM users WHERE username=?", (u,)).fetchone():
            errores.append("Ese usuario ya existe.")

        if errores:
            for e in errores:
                flash(e, "danger")
        else:
            cur = db.execute(
                "INSERT INTO users(username, password_hash, nombre, role) VALUES(?,?,?,?)",
                (u, generate_password_hash(pw), nombre, role),
            )
            nuevo_id = cur.lastrowid
            for oid in request.form.getlist("obras"):
                try:
                    db.execute("INSERT INTO user_obras(user_id, obra_id) VALUES(?,?)", (nuevo_id, oid))
                except sqlite3.IntegrityError:
                    pass
            db.commit()
            flash("Usuario creado: " + u, "success")
        return redirect(url_for("usuarios"))

    filas = db.execute("SELECT * FROM users ORDER BY nombre").fetchall()
    obras_l = db.execute("SELECT * FROM obras ORDER BY nombre").fetchall()
    asignadas = {}
    for r in db.execute(
        "SELECT uo.user_id, o.id, o.nombre FROM user_obras uo "
        "JOIN obras o ON o.id=uo.obra_id ORDER BY o.nombre").fetchall():
        asignadas.setdefault(r["user_id"], []).append({"id": r["id"], "nombre": r["nombre"]})
    return render_template("usuarios_list.html", usuarios=filas, obras=obras_l,
                           roles=ROLES, asignadas=asignadas)


@app.route("/usuarios/<int:uid>/obras", methods=["POST"])
@min_rank(ADMIN_RANK)
def usuario_obras(uid):
    db = get_db()
    db.execute("DELETE FROM user_obras WHERE user_id=?", (uid,))
    for oid in request.form.getlist("obras"):
        try:
            db.execute("INSERT INTO user_obras(user_id, obra_id) VALUES(?,?)", (uid, oid))
        except sqlite3.IntegrityError:
            pass
    db.commit()
    flash("Obras del usuario actualizadas.", "success")
    return redirect(url_for("usuarios"))

@app.route("/usuarios/<int:uid>/rol", methods=["POST"])
@min_rank(ADMIN_RANK)
def usuario_rol(uid):
    db = get_db()
    role = request.form.get("role", "")
    if uid == session["user_id"]:
        flash("No puedes cambiar tu propio rol.", "warning")
    elif role in ROLES:
        db.execute("UPDATE users SET role=? WHERE id=?", (role, uid))
        db.commit()
        flash("Rol actualizado.", "success")
    return redirect(url_for("usuarios"))

@app.route("/usuarios/<int:uid>/nombre", methods=["POST"])
@min_rank(ADMIN_RANK)
def usuario_nombre(uid):
    db = get_db()
    nombre = request.form.get("nombre", "").strip()
    if not nombre:
        flash("El nombre no puede quedar vacio.", "danger")
    else:
        db.execute("UPDATE users SET nombre=? WHERE id=?", (nombre, uid))
        db.commit()
        if uid == session.get("user_id"):
            session["nombre"] = nombre
        flash("Nombre actualizado.", "success")
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

@app.route("/parametros", methods=["GET", "POST"])
@min_rank(ADMIN_RANK)
def parametros():
    db = get_db()
    if request.method == "POST":
        for clave in ("sm_general", "sm_zlfn", "propuesta_extra", "jornada_horas", "porcentaje_despacho"):
            if clave not in request.form:
                continue
            val = (request.form.get(clave, "") or "").replace(",", ".").strip()
            try:
                float(val)
                set_param(db, clave, val)
            except ValueError:
                flash(f"Valor invalido para {clave}.", "danger")
        if request.form.get("sm_vigencia", "").strip():
            set_param(db, "sm_vigencia", request.form.get("sm_vigencia", "").strip())
        if "email_contador" in request.form:
            set_param(db, "email_contador", request.form.get("email_contador", "").strip())
        # datos de la empresa (para el contrato) y API market: solo si vienen en el envio
        for clave in ("empresa_razon_social", "empresa_rfc", "empresa_representante",
                      "empresa_domicilio", "empresa_instrumento", "empresa_volumen",
                      "empresa_fecha_escritura", "empresa_notario_num", "empresa_notario_nombre",
                      "empresa_notario_ciudad", "apimarket_token", "apimarket_url_nss",
                      "apimarket_url_vigencia"):
            if clave in request.form:
                set_param(db, clave, request.form.get(clave, "").strip())
        db.commit()
        flash("Parametros actualizados.", "success")
        return redirect(url_for("parametros"))
    claves = ("sm_general", "sm_zlfn", "sm_vigencia", "propuesta_extra", "jornada_horas",
              "porcentaje_despacho", "email_contador", "empresa_razon_social", "empresa_rfc",
              "empresa_representante", "empresa_domicilio", "empresa_instrumento", "empresa_volumen",
              "empresa_fecha_escritura", "empresa_notario_num", "empresa_notario_nombre",
              "empresa_notario_ciudad", "apimarket_token", "apimarket_url_nss",
              "apimarket_url_vigencia")
    datos = {k: get_param(db, k) for k in claves}
    return render_template("parametros.html", p=datos,
                           propuesta=propuesta_salario_alta(db))

@app.errorhandler(403)
def forbidden(e):
    return render_template("403.html"), 403

if __name__ == "__main__":
    app.run(debug=True, port=5000)
