
import os
import smtplib
from datetime import datetime, timedelta
from email.message import EmailMessage
from urllib.parse import quote_plus
from io import BytesIO
from openpyxl import Workbook

from flask import Flask, render_template_string, request, redirect, flash, session, send_from_directory, Response, send_file
from flask_sqlalchemy import SQLAlchemy
from sqlalchemy import text
from werkzeug.security import generate_password_hash, check_password_hash
from werkzeug.utils import secure_filename

from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

try:
    import cloudinary
    import cloudinary.uploader
except Exception:
    cloudinary = None


app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "ALESI_CAMBIAR_SECRET_KEY")

DB = os.environ.get("DATABASE_URL", "sqlite:///alesi.db")
if DB.startswith("postgres://"):
    DB = DB.replace("postgres://", "postgresql://", 1)

app.config["SQLALCHEMY_DATABASE_URI"] = DB
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False
db = SQLAlchemy(app)

os.makedirs("uploads", exist_ok=True)


MATERIAS = [
    "Familiar", "Laboral", "Civil", "Mercantil", "Penal", "Administrativo",
    "Fiscal", "Amparo", "Agrario", "Corporativo", "Notarial", "Municipal", "Otro"
]

TIPOS_INICIALES = [
    ("Familiar", "Divorcio voluntario"), ("Familiar", "Divorcio incausado"),
    ("Familiar", "Pensión alimenticia"), ("Familiar", "Aumento de pensión"),
    ("Familiar", "Reducción de pensión"), ("Familiar", "Guarda y custodia"),
    ("Familiar", "Régimen de convivencia"), ("Familiar", "Patria potestad"),
    ("Familiar", "Violencia familiar"), ("Familiar", "Sucesión testamentaria"),
    ("Familiar", "Sucesión intestamentaria"), ("Familiar", "Juicio hereditario"),
    ("Laboral", "Despido injustificado"), ("Laboral", "Pago de prestaciones"),
    ("Laboral", "Horas extras"), ("Laboral", "Reinstalación"),
    ("Laboral", "Convenio laboral"), ("Laboral", "Accidente de trabajo"),
    ("Civil", "Cumplimiento de contrato"), ("Civil", "Incumplimiento de contrato"),
    ("Civil", "Daños y perjuicios"), ("Civil", "Cobro de pesos"),
    ("Civil", "Arrendamiento"), ("Civil", "Prescripción positiva"),
    ("Civil", "Otorgamiento y firma de escritura"),
    ("Mercantil", "Cobro de pagaré"), ("Mercantil", "Juicio ejecutivo mercantil"),
    ("Mercantil", "Juicio oral mercantil"), ("Mercantil", "Cobro de factura"),
    ("Penal", "Robo"), ("Penal", "Fraude"), ("Penal", "Abuso de confianza"),
    ("Penal", "Daño en propiedad ajena"), ("Penal", "Despojo"),
    ("Penal", "Amenazas"), ("Penal", "Lesiones"),
    ("Penal", "Incumplimiento de obligaciones alimentarias"),
    ("Administrativo", "Responsabilidad administrativa"),
    ("Administrativo", "Procedimiento administrativo"), ("Administrativo", "Revocación de multa"),
    ("Administrativo", "Clausura"), ("Administrativo", "Auditoría"),
    ("Fiscal", "Crédito fiscal"), ("Fiscal", "Multa fiscal"), ("Fiscal", "Recurso de revocación"),
    ("Fiscal", "Juicio contencioso administrativo"),
    ("Amparo", "Amparo indirecto"), ("Amparo", "Amparo directo"),
    ("Amparo", "Suspensión provisional"), ("Amparo", "Suspensión definitiva"),
    ("Agrario", "Conflicto parcelario"), ("Agrario", "Posesión agraria"),
    ("Corporativo", "Constitución de sociedad"), ("Corporativo", "Acta de asamblea"),
    ("Corporativo", "Poderes"), ("Corporativo", "Modificación estatutaria"),
    ("Notarial", "Notario"), ("Notarial", "Copias certificadas"),
    ("Municipal", "Permiso municipal"), ("Municipal", "Uso de suelo"),
    ("Otro", "Convenio extrajudicial"), ("Otro", "Carta poder"), ("Otro", "Regularización de predio")
]


class Usuario(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    nombre = db.Column(db.String(180), nullable=False)
    usuario = db.Column(db.String(80), unique=True, nullable=False)
    correo = db.Column(db.String(180))
    password_hash = db.Column(db.String(255), nullable=False)
    rol = db.Column(db.String(80), default="Usuario")
    area = db.Column(db.String(80), default="")
    foto_url = db.Column(db.Text)
    activo = db.Column(db.Boolean, default=True)
    creado_en = db.Column(db.DateTime, default=datetime.now)


class TipoAsunto(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    materia = db.Column(db.String(80), nullable=False)
    nombre = db.Column(db.String(180), nullable=False)
    activo = db.Column(db.Boolean, default=True)


class Cliente(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    nombre = db.Column(db.String(200), nullable=False)
    telefono = db.Column(db.String(80))
    correo = db.Column(db.String(180))
    rfc = db.Column(db.String(80))
    curp = db.Column(db.String(80))
    materia = db.Column(db.String(80))
    tipo_asunto = db.Column(db.String(180))
    cobro_total = db.Column(db.Float, default=0)
    moneda_cobro = db.Column(db.String(10), default="MXN")
    direccion = db.Column(db.Text)
    observaciones = db.Column(db.Text)
    creado_en = db.Column(db.DateTime, default=datetime.now)


class Expediente(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    propietario_id = db.Column(db.Integer, db.ForeignKey("usuario.id"), nullable=False)
    cliente_id = db.Column(db.Integer, db.ForeignKey("cliente.id"))
    numero = db.Column(db.String(120), nullable=False)
    materia = db.Column(db.String(80))
    tipo_asunto = db.Column(db.String(180))
    autoridad = db.Column(db.String(200))
    actor = db.Column(db.String(200))
    demandado = db.Column(db.String(200))
    estado = db.Column(db.String(80), default="En trámite")
    prioridad = db.Column(db.String(40), default="Media")
    responsable = db.Column(db.String(180))
    fecha_inicio = db.Column(db.String(20))
    observaciones = db.Column(db.Text)
    creado_en = db.Column(db.DateTime, default=datetime.now)

    propietario = db.relationship("Usuario")
    cliente = db.relationship("Cliente")


class Compartido(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    expediente_id = db.Column(db.Integer, db.ForeignKey("expediente.id"))
    usuario_id = db.Column(db.Integer, db.ForeignKey("usuario.id"))
    permiso = db.Column(db.String(40), default="Lectura")
    creado_en = db.Column(db.DateTime, default=datetime.now)
    expediente = db.relationship("Expediente")
    usuario = db.relationship("Usuario")


class Movimiento(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    expediente_id = db.Column(db.Integer, db.ForeignKey("expediente.id"))
    usuario_id = db.Column(db.Integer, db.ForeignKey("usuario.id"))
    titulo = db.Column(db.String(220), nullable=False)
    fecha = db.Column(db.String(20))
    estatus = db.Column(db.String(80), default="Pendiente")
    proxima_accion = db.Column(db.String(250))
    fecha_limite = db.Column(db.String(20))
    hora_limite = db.Column(db.String(10), default="09:00")
    observaciones = db.Column(db.Text)
    archivo_url = db.Column(db.Text)
    notificar = db.Column(db.Boolean, default=True)
    creado_en = db.Column(db.DateTime, default=datetime.now)
    expediente = db.relationship("Expediente")
    usuario = db.relationship("Usuario")


class Audiencia(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    expediente_id = db.Column(db.Integer, db.ForeignKey("expediente.id"))
    usuario_id = db.Column(db.Integer, db.ForeignKey("usuario.id"))
    titulo = db.Column(db.String(220), nullable=False)
    fecha = db.Column(db.String(20), nullable=False)
    hora = db.Column(db.String(10), nullable=False)
    autoridad = db.Column(db.String(220))
    sala = db.Column(db.String(120))
    modalidad = db.Column(db.String(80), default="Presencial")
    enlace = db.Column(db.Text)
    observaciones = db.Column(db.Text)
    notificar = db.Column(db.Boolean, default=True)
    creado_en = db.Column(db.DateTime, default=datetime.now)
    expediente = db.relationship("Expediente")
    usuario = db.relationship("Usuario")


class Cobranza(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    folio = db.Column(db.String(40), unique=True)
    cliente_id = db.Column(db.Integer, db.ForeignKey("cliente.id"))
    expediente_id = db.Column(db.Integer, db.ForeignKey("expediente.id"))
    usuario_id = db.Column(db.Integer, db.ForeignKey("usuario.id"))
    concepto = db.Column(db.String(120))
    descripcion = db.Column(db.Text)
    monto_total = db.Column(db.Float, default=0)
    abono = db.Column(db.Float, default=0)
    saldo = db.Column(db.Float, default=0)
    moneda = db.Column(db.String(10), default="MXN")
    forma_pago = db.Column(db.String(50), default="Efectivo")
    fecha_pago = db.Column(db.String(20))
    proximo_pago = db.Column(db.String(20))
    estatus = db.Column(db.String(40), default="Pendiente")
    comprobante_url = db.Column(db.Text)
    creado_en = db.Column(db.DateTime, default=datetime.now)
    cliente = db.relationship("Cliente")
    expediente = db.relationship("Expediente")
    usuario = db.relationship("Usuario")


class NotificacionEnviada(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    tipo = db.Column(db.String(40))
    referencia_id = db.Column(db.Integer)
    anticipacion = db.Column(db.String(20))
    usuario_id = db.Column(db.Integer)
    enviado_en = db.Column(db.DateTime, default=datetime.now)


class Bitacora(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    expediente_id = db.Column(db.Integer, db.ForeignKey("expediente.id"))
    usuario_id = db.Column(db.Integer, db.ForeignKey("usuario.id"))
    accion = db.Column(db.String(220), nullable=False)
    detalle = db.Column(db.Text)
    creado_en = db.Column(db.DateTime, default=datetime.now)
    expediente = db.relationship("Expediente")
    usuario = db.relationship("Usuario")


class MensajeExpediente(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    expediente_id = db.Column(db.Integer, db.ForeignKey("expediente.id"))
    usuario_id = db.Column(db.Integer, db.ForeignKey("usuario.id"))
    mensaje = db.Column(db.Text, nullable=False)
    creado_en = db.Column(db.DateTime, default=datetime.now)
    expediente = db.relationship("Expediente")
    usuario = db.relationship("Usuario")


class NotificacionInterna(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    usuario_id = db.Column(db.Integer, db.ForeignKey("usuario.id"))
    titulo = db.Column(db.String(220))
    mensaje = db.Column(db.Text)
    enlace = db.Column(db.String(250))
    leida = db.Column(db.Boolean, default=False)
    creado_en = db.Column(db.DateTime, default=datetime.now)
    usuario = db.relationship("Usuario")


class AvisoExpediente(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    expediente_id = db.Column(db.Integer, db.ForeignKey("expediente.id"))
    usuario_id = db.Column(db.Integer, db.ForeignKey("usuario.id"))
    titulo = db.Column(db.String(220))
    mensaje = db.Column(db.Text)
    fecha_aviso = db.Column(db.String(20))
    creado_en = db.Column(db.DateTime, default=datetime.now)
    expediente = db.relationship("Expediente")
    usuario = db.relationship("Usuario")


def ensure_schema_updates():
    dialect = db.engine.dialect.name

    def add_column(table, column, definition):
        try:
            if dialect == "postgresql":
                db.session.execute(text(f"ALTER TABLE {table} ADD COLUMN IF NOT EXISTS {column} {definition}"))
            elif dialect == "sqlite":
                cols = [row[1] for row in db.session.execute(text(f"PRAGMA table_info({table})")).fetchall()]
                if column not in cols:
                    db.session.execute(text(f"ALTER TABLE {table} ADD COLUMN {column} {definition}"))
        except Exception as exc:
            print("Schema update skipped:", table, column, exc)

    add_column("usuario", "area", "VARCHAR(80)")
    add_column("usuario", "foto_url", "TEXT")
    add_column("cliente", "rfc", "VARCHAR(80)")
    add_column("cliente", "curp", "VARCHAR(80)")
    add_column("cliente", "materia", "VARCHAR(80)")
    add_column("cliente", "tipo_asunto", "VARCHAR(180)")
    add_column("cliente", "cobro_total", "FLOAT DEFAULT 0")
    add_column("cliente", "moneda_cobro", "VARCHAR(10) DEFAULT 'MXN'")
    add_column("expediente", "materia", "VARCHAR(80)")
    add_column("expediente", "tipo_asunto", "VARCHAR(180)")
    add_column("cobranza", "folio", "VARCHAR(40)")
    db.session.commit()


def init_db():
    with app.app_context():
        db.create_all()
        ensure_schema_updates()

        usuarios_iniciales = [
            ("Rosa Isela Vázquez Medina", "ROUSSVAM", "lic.rosavazquezm@gmail.com", "RIVaM061", "Administrador", "Administración"),
            ("Lic. Carmen Acosta C.", "lic.carmenacostac", "mrsmunoz68@gmail.com", "LIC.CARMENAC68", "Usuario", "Jurídico"),
            ("Computadora Oficina", "oficina.alesi", "", "Oficina2026*", "Usuario Oficina", "Oficina"),
        ]

        for nombre, usuario, correo, password, rol, area in usuarios_iniciales:
            existente = Usuario.query.filter_by(usuario=usuario).first()
            if not existente:
                db.session.add(Usuario(nombre=nombre, usuario=usuario, correo=correo, password_hash=generate_password_hash(password), rol=rol, area=area))
            else:
                existente.area = existente.area or area

        for materia, nombre in TIPOS_INICIALES:
            if not TipoAsunto.query.filter_by(materia=materia, nombre=nombre).first():
                db.session.add(TipoAsunto(materia=materia, nombre=nombre))

        db.session.commit()


def actual():
    if not session.get("usuario_id"):
        return None
    return db.session.get(Usuario, session["usuario_id"])


def req():
    return actual() is not None


def admin():
    u = actual()
    return bool(u and u.rol == "Administrador")


def oficina_o_admin():
    u = actual()
    return bool(u and (u.rol == "Administrador" or u.rol == "Usuario Oficina" or u.area == "Oficina"))


def subir(archivo):
    if not archivo or not archivo.filename:
        return ""
    if cloudinary and os.getenv("CLOUDINARY_CLOUD_NAME") and os.getenv("CLOUDINARY_API_KEY") and os.getenv("CLOUDINARY_API_SECRET"):
        cloudinary.config(
            cloud_name=os.getenv("CLOUDINARY_CLOUD_NAME"),
            api_key=os.getenv("CLOUDINARY_API_KEY"),
            api_secret=os.getenv("CLOUDINARY_API_SECRET"),
            secure=True,
        )
        resultado = cloudinary.uploader.upload(archivo, resource_type="auto", folder="alesi_grupo_juridico")
        return resultado.get("secure_url", "")
    nombre_seguro = secure_filename(archivo.filename)
    ruta = "uploads/" + datetime.now().strftime("%Y%m%d%H%M%S_") + nombre_seguro
    archivo.save(ruta)
    return "/" + ruta


def registrar_bitacora(expediente_id, accion, detalle=""):
    if actual():
        db.session.add(Bitacora(expediente_id=expediente_id, usuario_id=actual().id, accion=accion, detalle=detalle))


def permiso_expediente(expediente_id):
    u = actual()
    if not u:
        return ""
    exp = db.session.get(Expediente, expediente_id)
    if not exp:
        return ""
    if u.rol == "Administrador" or exp.propietario_id == u.id:
        return "Administración"
    comp = Compartido.query.filter_by(expediente_id=expediente_id, usuario_id=u.id).first()
    return comp.permiso if comp else ""


def puede_ver(expediente_id):
    return permiso_expediente(expediente_id) in ["Lectura", "Edición", "Administración"]


def puede_editar(expediente_id):
    return permiso_expediente(expediente_id) in ["Edición", "Administración"]


def puede_administrar(expediente_id):
    return permiso_expediente(expediente_id) == "Administración"


def usuarios_acceso(expediente_id):
    expediente = db.session.get(Expediente, expediente_id)
    ids = set()
    if expediente:
        ids.add(expediente.propietario_id)
    for compartido in Compartido.query.filter_by(expediente_id=expediente_id).all():
        ids.add(compartido.usuario_id)
    if not ids:
        return []
    return Usuario.query.filter(Usuario.id.in_(ids), Usuario.activo == True).all()


def correo(destinatario, asunto, cuerpo):
    if not all([os.getenv("SMTP_HOST"), os.getenv("SMTP_USER"), os.getenv("SMTP_PASS"), destinatario]):
        print("[SIMULADO]", destinatario, asunto)
        return False
    mensaje = EmailMessage()
    mensaje["Subject"] = asunto
    mensaje["From"] = os.getenv("FROM_EMAIL", os.getenv("SMTP_USER"))
    mensaje["To"] = destinatario
    mensaje.set_content(cuerpo)
    with smtplib.SMTP(os.getenv("SMTP_HOST"), int(os.getenv("SMTP_PORT", "587"))) as servidor:
        servidor.starttls()
        servidor.login(os.getenv("SMTP_USER"), os.getenv("SMTP_PASS"))
        servidor.send_message(mensaje)
    return True


def notificar_interna(usuario_id, titulo, mensaje, enlace=""):
    db.session.add(NotificacionInterna(usuario_id=usuario_id, titulo=titulo, mensaje=mensaje, enlace=enlace))


def crear_aviso_para_accesos(expediente_id, titulo, mensaje):
    expediente = db.session.get(Expediente, expediente_id)
    if not expediente:
        return
    db.session.add(AvisoExpediente(expediente_id=expediente_id, usuario_id=actual().id if actual() else None, titulo=titulo, mensaje=mensaje, fecha_aviso=datetime.now().strftime("%Y-%m-%d")))
    for usuario in usuarios_acceso(expediente_id):
        notificar_interna(usuario.id, titulo, mensaje, f"/expediente/{expediente_id}")


def revisar_y_enviar_notificaciones():
    enviados = 0
    ahora = datetime.now()
    ventana_minutos = 12
    items = []
    for audiencia in Audiencia.query.filter_by(notificar=True).all():
        items.append(("audiencia", audiencia.id, audiencia.expediente_id, audiencia.fecha, audiencia.hora, f"Audiencia: {audiencia.titulo}", audiencia.expediente.numero if audiencia.expediente else ""))
    for movimiento in Movimiento.query.filter_by(notificar=True).all():
        if movimiento.fecha_limite:
            items.append(("vencimiento", movimiento.id, movimiento.expediente_id, movimiento.fecha_limite, movimiento.hora_limite or "09:00", f"Vencimiento: {movimiento.titulo}", movimiento.expediente.numero if movimiento.expediente else ""))
    for tipo, referencia_id, expediente_id, fecha, hora, titulo, numero in items:
        try:
            fecha_hora = datetime.strptime(fecha + " " + hora, "%Y-%m-%d %H:%M")
        except Exception:
            continue
        for etiqueta, horas in [("24h", 24), ("2h", 2)]:
            momento_aviso = fecha_hora - timedelta(hours=horas)
            diferencia = abs((ahora - momento_aviso).total_seconds()) / 60
            if diferencia <= ventana_minutos:
                for usuario in usuarios_acceso(expediente_id):
                    ya_enviada = NotificacionEnviada.query.filter_by(tipo=tipo, referencia_id=referencia_id, anticipacion=etiqueta, usuario_id=usuario.id).first()
                    if ya_enviada:
                        continue
                    notificar_interna(usuario.id, f"{titulo} en {etiqueta}", f"Expediente {numero}. Fecha: {fecha} {hora}", f"/expediente/{expediente_id}")
                    if usuario.correo:
                        correo(usuario.correo, f"ALESI - {titulo} en {etiqueta}", f"Expediente: {numero}\n{titulo}\nFecha: {fecha} {hora}\n\nALESI GRUPO JURÍDICO")
                    db.session.add(NotificacionEnviada(tipo=tipo, referencia_id=referencia_id, anticipacion=etiqueta, usuario_id=usuario.id))
                    enviados += 1
    db.session.commit()
    return enviados


BASE = """
<!doctype html>
<html lang="es">
<head>
<meta charset="utf-8">
<title>ALESI GRUPO JURÍDICO</title>
<style>
:root{--n:#111827;--t:#14B8A6;--d:#B38B2E}
body{margin:0;font-family:Arial;background:#F8FAFC;color:#374151}
.top{background:var(--n);color:white;padding:8px 28px;display:flex;justify-content:space-between;gap:20px;align-items:center}
.clock{font-weight:bold;color:#f8fafc}
.userbar{display:flex;align-items:center;gap:10px;justify-content:flex-end}
.userbar img{width:34px;height:34px;border-radius:50%;object-fit:cover;border:2px solid var(--d);background:white}
.userbar a{color:white;text-decoration:none;background:#374151;padding:7px 10px;border-radius:8px;font-weight:bold}
.userbar a.logout{background:#991b1b}
.hero{background:linear-gradient(120deg,#111827,#0f766e);color:white;padding:28px 36px;border-bottom:6px solid var(--d);display:flex;align-items:center}
.logo{width:100px;height:100px;border-radius:50%;background:white;object-fit:cover;border:4px solid var(--d);margin-right:20px}
nav{background:#0f172a;padding:12px 30px;line-height:3;display:flex;flex-wrap:wrap;gap:10px}
nav a{color:white;text-decoration:none;font-weight:bold;background:#1f2937;padding:9px 13px;border-radius:10px;border:1px solid #334155;box-shadow:0 2px 6px #0004}
nav a:hover{background:var(--d);color:white}
main{padding:25px 30px}
.card{background:white;border-radius:14px;padding:22px;margin-bottom:20px;box-shadow:0 3px 12px #0002;border-top:4px solid var(--d)}
input,select,textarea{width:100%;padding:10px;margin:6px 0 12px;border:1px solid #cbd5e1;border-radius:8px;box-sizing:border-box}
button,.btn{background:var(--d);color:white;border:0;border-radius:8px;padding:10px 16px;text-decoration:none;font-weight:bold;display:inline-block;cursor:pointer}
.btn2{background:var(--t)}.btnDark{background:var(--n)}.btnDanger{background:#991b1b}.btnWarn{background:#b45309}
table{width:100%;border-collapse:collapse} th,td{padding:10px;border-bottom:1px solid #e5e7eb;text-align:left;vertical-align:top} th{background:#e5e7eb}
.grid{display:grid;grid-template-columns:repeat(4,1fr);gap:15px}.grid2{display:grid;grid-template-columns:repeat(2,1fr);gap:15px}
.stat{text-align:center;background:white;border-radius:14px;padding:20px;box-shadow:0 3px 12px #0002;border-bottom:4px solid var(--t)}
.flash{background:#dcfce7;padding:10px;border-left:5px solid #166534;margin-bottom:12px}.badge{background:var(--t);color:white;padding:4px 8px;border-radius:999px;font-size:12px;font-weight:bold}
.small{font-size:13px;color:#6b7280}.chat{background:#f8fafc;border-left:4px solid var(--t);padding:10px;margin-bottom:8px;border-radius:8px}.bit{background:#fff7ed;border-left:4px solid var(--d);padding:9px;margin-bottom:7px;border-radius:8px}
@media print{nav,.top,button,.btn,.no-print,.no-imprimir{display:none!important}.card{box-shadow:none;border:0}.hero{display:none!important}main{padding:0}body{background:white}.print-list{display:block!important}.solo-imprimir{display:block!important}}
@media(max-width:900px){.grid,.grid2{grid-template-columns:1fr}}
</style>
<script>
function relojAlesi(){
    const el=document.getElementById("clock");
    if(!el) return;
    const d=new Date();
    el.textContent=d.toLocaleDateString("es-MX",{weekday:"long",year:"numeric",month:"long",day:"numeric"})+" | "+d.toLocaleTimeString("es-MX");
}
setInterval(relojAlesi,1000);
window.onload=relojAlesi;
</script>
</head>
<body>
<div class="top">
    <div>Sistema Jurídico + Planeador + Cobranza</div>
    <div class="clock" id="clock"></div>
    <div class="userbar">
        {% if usuario %}
            {% if usuario.foto_url %}<img src="{{usuario.foto_url}}">{% else %}<img src="/logo">{% endif %}
            <span>{{usuario.nombre}} | {{usuario.rol}}</span>
            <a href="/perfil">Mi Perfil</a>
            <a class="logout" href="/logout">Cerrar sesión</a>
        {% endif %}
    </div>
</div>
<section class="hero">
    <img class="logo" src="/logo">
    <div><h1>ALESI GRUPO JURÍDICO</h1><p>Sistema Integral de Gestión Jurídica y Control de Expedientes</p></div>
</section>
{% if usuario %}
<nav>
    <a href="/">Inicio</a>
    <a href="/perfil">Mi Perfil</a>
    <a href="/mis-expedientes">Mis Expedientes</a>
    <a href="/compartidos-conmigo">Expedientes Compartidos</a>
    <a href="/planeador">Planeador</a>
    <a href="/clientes">Clientes</a>
    <a href="/audiencias">Audiencias</a>
    <a href="/vencimientos">Vencimientos</a>
    <a href="/cobranza">Cobranza</a>
    {% if usuario.rol == 'Administrador' %}<a href="/respaldos">Respaldos</a>{% endif %}
    <a href="/notificaciones-internas">Avisos</a>
    <a href="/jurisprudencia">Jurisprudencia</a>
    <a href="/boletin-judicial">Boletín Judicial</a>
    <a href="/teja">TEJA</a>
    {% if usuario.rol == 'Administrador' %}
        <a href="/usuarios">Usuarios</a>
        <a href="/catalogo-asuntos">Catálogo</a>
        <a href="/notificaciones">Notificaciones</a>
    {% endif %}
</nav>
{% endif %}
<main>
{% with messages=get_flashed_messages() %}{% for m in messages %}<div class="flash">{{m}}</div>{% endfor %}{% endwith %}
{{contenido|safe}}
</main>
</body>
</html>
"""


def render(contenido, **kw):
    return render_template_string(BASE, contenido=render_template_string(contenido, **kw), usuario=actual())


@app.route("/logo")
def logo():
    if os.path.exists("logo_alesi.png"):
        return send_from_directory(".", "logo_alesi.png")
    return ("", 404)


@app.route("/uploads/<path:archivo>")
def uploads(archivo):
    return send_from_directory("uploads", archivo)


@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        usuario = Usuario.query.filter_by(usuario=request.form["usuario"], activo=True).first()
        if usuario and check_password_hash(usuario.password_hash, request.form["password"]):
            session["usuario_id"] = usuario.id
            return redirect("/")
        flash("Usuario o contraseña incorrectos.")
    return render("""<div class="card" style="max-width:420px;margin:auto"><h2>Iniciar sesión</h2><form method="post"><label>Usuario</label><input name="usuario" required><label>Contraseña</label><input type="password" name="password" required><button>Entrar</button></form></div>""")


@app.route("/logout")
def logout():
    session.clear()
    return redirect("/login")


@app.route("/perfil", methods=["GET", "POST"])
def perfil():
    if not req():
        return redirect("/login")
    u = actual()
    if request.method == "POST":
        u.nombre = request.form.get("nombre", u.nombre)
        u.correo = request.form.get("correo", u.correo)
        nueva = request.form.get("password", "").strip()
        if nueva:
            u.password_hash = generate_password_hash(nueva)
        foto = subir(request.files.get("foto"))
        if foto:
            u.foto_url = foto
        db.session.commit()
        flash("Perfil actualizado.")
        return redirect("/perfil")
    return render("""<div class="card" style="max-width:760px;margin:auto"><h2>Mi Perfil</h2>{% if u.foto_url %}<p><img src="{{u.foto_url}}" style="width:120px;height:120px;border-radius:50%;object-fit:cover;border:3px solid #B38B2E"></p>{% endif %}<form method="post" enctype="multipart/form-data"><label>Nombre</label><input name="nombre" value="{{u.nombre or ''}}" required><label>Usuario</label><input value="{{u.usuario}}" disabled><label>Correo</label><input type="email" name="correo" value="{{u.correo or ''}}"><label>Nueva contraseña</label><input type="password" name="password" placeholder="Déjalo vacío si no deseas cambiarla"><label>Cambiar fotografía</label><input type="file" name="foto"><button>Guardar cambios</button></form></div>""", u=u)


@app.route("/")
def inicio():
    if not req():
        return redirect("/login")
    hoy = datetime.now().date()
    ayer = hoy - timedelta(days=1)
    manana = hoy + timedelta(days=1)
    fechas = [ayer.strftime("%Y-%m-%d"), hoy.strftime("%Y-%m-%d"), manana.strftime("%Y-%m-%d")]
    labels = {ayer.strftime("%Y-%m-%d"): "Ayer", hoy.strftime("%Y-%m-%d"): "Hoy", manana.strftime("%Y-%m-%d"): "Mañana"}

    pendientes = []
    for a in Audiencia.query.filter(Audiencia.fecha.in_(fechas)).order_by(Audiencia.fecha, Audiencia.hora).all():
        pendientes.append({"fecha": a.fecha, "label": labels.get(a.fecha, ""), "tipo": "Audiencia", "texto": f"{a.hora} - {a.expediente.numero if a.expediente else ''} - {a.titulo}"})
    for v in Movimiento.query.filter(Movimiento.fecha_limite.in_(fechas)).order_by(Movimiento.fecha_limite, Movimiento.hora_limite).all():
        pendientes.append({"fecha": v.fecha_limite, "label": labels.get(v.fecha_limite, ""), "tipo": "Vencimiento", "texto": f"{v.hora_limite or ''} - {v.expediente.numero if v.expediente else ''} - {v.proxima_accion or v.titulo}"})
    for p in Cobranza.query.filter(Cobranza.proximo_pago.in_(fechas), Cobranza.saldo > 0, Cobranza.estatus != "Liquidado").order_by(Cobranza.proximo_pago).all():
        pendientes.append({"fecha": p.proximo_pago, "label": labels.get(p.proximo_pago, ""), "tipo": "Pago", "texto": f"{p.cliente.nombre if p.cliente else ''} - {p.moneda} {p.saldo}"})

    inicio_semana = hoy - timedelta(days=hoy.weekday())
    fin_semana = inicio_semana + timedelta(days=6)
    cobranza_semana = Cobranza.query.filter(Cobranza.proximo_pago >= inicio_semana.strftime("%Y-%m-%d"), Cobranza.proximo_pago <= fin_semana.strftime("%Y-%m-%d"), Cobranza.saldo > 0, Cobranza.estatus != "Liquidado").count()
    avisos = NotificacionInterna.query.filter_by(usuario_id=actual().id, leida=False).count()

    return render("""<div class="grid"><div class="stat"><h2>{{clientes}}</h2><p>Clientes</p></div><div class="stat"><h2>{{expedientes}}</h2><p>Expedientes</p></div><div class="stat"><h2>{{audiencias}}</h2><p>Audiencias</p></div><div class="stat"><h2>{{vencimientos}}</h2><p>Vencimientos</p></div></div><br><div class="grid"><div class="stat"><h2>{{avisos}}</h2><p>Avisos internos no leídos</p></div><div class="stat"><h2>{{cobranza_semana}}</h2><p>Cobranza pendiente esta semana</p></div></div><br><div class="card"><h2>Pendientes de ayer, hoy y mañana</h2><table><tr><th>Día</th><th>Fecha</th><th>Tipo</th><th>Detalle</th></tr>{% for p in pendientes %}<tr><td>{{p.label}}</td><td>{{p.fecha}}</td><td>{{p.tipo}}</td><td>{{p.texto}}</td></tr>{% endfor %}</table></div>""", clientes=Cliente.query.count(), expedientes=Expediente.query.count(), audiencias=Audiencia.query.count(), vencimientos=Movimiento.query.filter(Movimiento.fecha_limite != "").count(), avisos=avisos, cobranza_semana=cobranza_semana, pendientes=pendientes)


@app.route("/usuarios", methods=["GET", "POST"])
def usuarios():
    if not req() or not admin():
        return redirect("/")
    if request.method == "POST":
        if Usuario.query.count() >= 15:
            flash("Límite de 15 usuarios.")
            return redirect("/usuarios")
        db.session.add(Usuario(nombre=request.form["nombre"], usuario=request.form["usuario"], correo=request.form["correo"], rol=request.form["rol"], area=request.form["area"], foto_url=subir(request.files.get("foto")), password_hash=generate_password_hash(request.form["password"])))
        db.session.commit()
        flash("Usuario creado.")
        return redirect("/usuarios")
    return render("""<div class="grid2"><div class="card"><h2>Crear usuario</h2><form method="post" enctype="multipart/form-data"><label>Nombre</label><input name="nombre" required><label>Usuario</label><input name="usuario" required><label>Correo</label><input type="email" name="correo"><label>Contraseña</label><input type="password" name="password" required><label>Rol</label><select name="rol"><option>Usuario</option><option>Administrador</option><option>Usuario Oficina</option></select><label>Área</label><select name="area"><option>Jurídico</option><option>Administración</option><option>Oficina</option><option>Penal</option><option>Laboral</option><option>Administrativo</option><option>Civil</option><option>Mercantil</option><option>Familiar</option></select><label>Foto</label><input type="file" name="foto"><button>Guardar</button></form></div><div class="card"><h2>Usuarios</h2><table><tr><th>Nombre</th><th>Usuario</th><th>Correo</th><th>Rol</th><th>Área</th></tr>{% for u in datos %}<tr><td>{{u.nombre}}</td><td>{{u.usuario}}</td><td>{{u.correo}}</td><td>{{u.rol}}</td><td>{{u.area}}</td></tr>{% endfor %}</table></div></div>""", datos=Usuario.query.order_by(Usuario.id).all())


@app.route("/catalogo-asuntos", methods=["GET", "POST"])
def catalogo_asuntos():
    if not req() or not admin():
        return redirect("/")
    if request.method == "POST":
        db.session.add(TipoAsunto(materia=request.form["materia"], nombre=request.form["nombre"]))
        db.session.commit()
        flash("Tipo de asunto agregado.")
        return redirect("/catalogo-asuntos")
    return render("""<div class="grid2"><div class="card"><h2>Agregar tipo de asunto</h2><form method="post"><label>Materia</label><select name="materia">{% for m in materias %}<option>{{m}}</option>{% endfor %}</select><label>Tipo de asunto</label><input name="nombre" required><button>Guardar</button></form></div><div class="card"><h2>Catálogo de asuntos</h2><table><tr><th>Materia</th><th>Tipo de asunto</th></tr>{% for t in tipos %}<tr><td>{{t.materia}}</td><td>{{t.nombre}}</td></tr>{% endfor %}</table></div></div>""", materias=MATERIAS, tipos=TipoAsunto.query.order_by(TipoAsunto.materia, TipoAsunto.nombre).all())


@app.route("/clientes", methods=["GET", "POST"])
def clientes():
    if not req():
        return redirect("/login")
    if request.method == "POST":
        db.session.add(Cliente(
            nombre=request.form["nombre"], telefono=request.form["telefono"], correo=request.form["correo"],
            rfc=request.form.get("rfc", ""), curp=request.form.get("curp", ""), direccion=request.form["direccion"],
            materia=request.form["materia"], tipo_asunto=request.form["tipo_asunto"],
            cobro_total=float(request.form.get("cobro_total") or 0), moneda_cobro=request.form.get("moneda_cobro", "MXN"),
            observaciones=request.form["observaciones"]
        ))
        db.session.commit()
        flash("Cliente guardado.")
        return redirect("/clientes")
    return render("""<div class="card no-print no-imprimir"><h2>Registrar cliente</h2><form method="post"><label>Nombre</label><input name="nombre" required><label>Teléfono</label><input name="telefono"><label>Correo</label><input name="correo"><label>RFC</label><input name="rfc"><label>CURP</label><input name="curp"><label>Materia jurídica</label><select name="materia">{% for m in materias %}<option>{{m}}</option>{% endfor %}</select><label>Tipo de asunto que le llevo</label><select name="tipo_asunto">{% for t in tipos %}<option>{{t.nombre}}</option>{% endfor %}</select><label>Cobro total pactado</label><input type="number" step="0.01" name="cobro_total"><label>Moneda</label><select name="moneda_cobro"><option>MXN</option><option>USD</option></select><label>Dirección</label><textarea name="direccion"></textarea><label>Observaciones</label><textarea name="observaciones"></textarea><button>Guardar</button></form></div><div class="card print-list"><h2>Lista de clientes <button class="no-print" onclick="window.print()">Imprimir lista</button></h2><table><tr><th>Cliente</th><th>Teléfono</th><th>Materia</th><th>Tipo asunto</th><th>Cobro total</th></tr>{% for c in datos %}<tr><td><a href="/cliente/{{c.id}}">{{c.nombre}}</a></td><td>{{c.telefono}}</td><td>{{c.materia}}</td><td>{{c.tipo_asunto}}</td><td>{{c.moneda_cobro}} {{c.cobro_total}}</td></tr>{% endfor %}</table></div>""", materias=MATERIAS, tipos=TipoAsunto.query.order_by(TipoAsunto.materia, TipoAsunto.nombre).all(), datos=Cliente.query.order_by(Cliente.nombre).all())


@app.route("/cliente/<int:id>")
def cliente(id):
    if not req():
        return redirect("/login")
    c = Cliente.query.get_or_404(id)
    pagos = Cobranza.query.filter_by(cliente_id=id).order_by(Cobranza.id.desc()).all()
    exps = Expediente.query.filter_by(cliente_id=id).order_by(Expediente.id.desc()).all()
    return render("""<div class="card"><h2>Cliente <button onclick="window.print()">Imprimir</button></h2><p><b>Nombre:</b> {{c.nombre}}</p><p><b>Teléfono:</b> {{c.telefono}}</p><p><b>Correo:</b> {{c.correo}}</p><p><b>RFC:</b> {{c.rfc}}</p><p><b>CURP:</b> {{c.curp}}</p><p><b>Materia:</b> {{c.materia}}</p><p><b>Tipo de asunto:</b> {{c.tipo_asunto}}</p><p><b>Cobro total pactado:</b> {{c.moneda_cobro}} {{c.cobro_total}}</p><p><b>Dirección:</b> {{c.direccion}}</p><p>{{c.observaciones}}</p></div><div class="card"><h2>Expedientes</h2><table><tr><th>Número</th><th>Materia</th><th>Tipo</th><th>Estado</th></tr>{% for e in exps %}<tr><td>{{e.numero}}</td><td>{{e.materia}}</td><td>{{e.tipo_asunto}}</td><td>{{e.estado}}</td></tr>{% endfor %}</table></div><div class="card"><h2>Pagos</h2><table><tr><th>Folio</th><th>Concepto</th><th>Abono</th><th>Saldo</th><th>Recibo</th></tr>{% for p in pagos %}<tr><td>{{p.folio}}</td><td>{{p.concepto}}</td><td>{{p.moneda}} {{p.abono}}</td><td>{{p.saldo}}</td><td><a href="/recibo/{{p.id}}">Ver</a> | <a href="/recibo-pdf/{{p.id}}">PDF</a></td></tr>{% endfor %}</table></div>""", c=c, pagos=pagos, exps=exps)


@app.route("/expedientes", methods=["GET", "POST"])
def expedientes():
    if not req():
        return redirect("/login")
    if request.method == "POST":
        nuevo = Expediente(
            propietario_id=actual().id,
            cliente_id=request.form.get("cliente_id") or None,
            numero=request.form["numero"],
            materia=request.form["materia"],
            tipo_asunto=request.form["tipo_asunto"],
            autoridad=request.form["autoridad"],
            actor=request.form["actor"],
            demandado=request.form["demandado"],
            estado=request.form["estado"],
            prioridad=request.form["prioridad"],
            responsable=request.form["responsable"],
            fecha_inicio=request.form["fecha_inicio"],
            observaciones=request.form["observaciones"],
        )
        db.session.add(nuevo)
        db.session.flush()
        registrar_bitacora(nuevo.id, "Creó expediente", f"Expediente {nuevo.numero}")
        db.session.commit()
        flash("Expediente creado.")
        return redirect("/mis-expedientes")
    datos = Expediente.query.order_by(Expediente.id.desc()).all() if oficina_o_admin() else Expediente.query.filter_by(propietario_id=actual().id).order_by(Expediente.id.desc()).all()
    return render("""<div class="card no-print"><h2>Nuevo expediente</h2><form method="post"><label>Cliente</label><select name="cliente_id"><option value="">Sin cliente</option>{% for cliente in clientes %}<option value="{{cliente.id}}">{{cliente.nombre}}</option>{% endfor %}</select><label>Número</label><input name="numero" required><label>Materia</label><select name="materia">{% for m in materias %}<option>{{m}}</option>{% endfor %}</select><label>Tipo de asunto</label><select name="tipo_asunto">{% for t in tipos %}<option>{{t.nombre}}</option>{% endfor %}</select><label>Autoridad</label><input name="autoridad"><label>Actor</label><input name="actor"><label>Demandado</label><input name="demandado"><label>Estado</label><select name="estado"><option>Nuevo</option><option>En trámite</option><option>Pendiente de acuerdo</option><option>Pendiente de audiencia</option><option>Concluido</option><option>Archivado</option><option>Urgente</option></select><label>Prioridad</label><select name="prioridad"><option>Baja</option><option selected>Media</option><option>Alta</option></select><label>Responsable</label><input name="responsable"><label>Fecha inicio</label><input type="date" name="fecha_inicio"><label>Observaciones</label><textarea name="observaciones"></textarea><button>Guardar</button></form></div><div class="card print-list"><h2>Lista de expedientes cargados <button class="no-print" onclick="window.print()">Imprimir lista</button></h2><table><tr><th>Número</th><th>Cliente</th><th>Materia</th><th>Tipo asunto</th><th>Estado</th><th class="no-print">Acción</th></tr>{% for expediente in datos %}<tr><td>{{expediente.numero}}</td><td>{{expediente.cliente.nombre if expediente.cliente else ""}}</td><td>{{expediente.materia}}</td><td>{{expediente.tipo_asunto}}</td><td>{{expediente.estado}}</td><td class="no-print"><a class="btn btn2" href="/expediente/{{expediente.id}}">Abrir</a></td></tr>{% endfor %}</table></div>""", clientes=Cliente.query.order_by(Cliente.nombre).all(), materias=MATERIAS, tipos=TipoAsunto.query.order_by(TipoAsunto.materia, TipoAsunto.nombre).all(), datos=datos)


@app.route("/mis-expedientes")
def mis_expedientes():
    if not req():
        return redirect("/login")
    u = actual()
    if oficina_o_admin():
        datos = Expediente.query.order_by(Expediente.id.desc()).all()
        titulo = "Expedientes del despacho"
    else:
        propios = Expediente.query.filter_by(propietario_id=u.id).all()
        compartidos = [c.expediente for c in Compartido.query.filter_by(usuario_id=u.id).all() if c.expediente]
        mapa = {e.id: e for e in propios + compartidos}
        datos = sorted(mapa.values(), key=lambda x: x.id, reverse=True)
        titulo = "Mis Expedientes"
    return render("""<div class="card no-print"><h2>Nuevo expediente</h2><p>Para capturar un nuevo expediente usa el botón siguiente.</p><a class="btn" href="/expedientes">Nuevo expediente</a></div><div class="card print-list"><h2>{{titulo}} <button class="no-print" onclick="window.print()">Imprimir lista</button></h2><table><tr><th>Número</th><th>Cliente</th><th>Materia</th><th>Tipo</th><th>Estado</th><th>Propietario</th><th class="no-print">Acción</th></tr>{% for e in datos %}<tr><td>{{e.numero}}</td><td>{{e.cliente.nombre if e.cliente else ""}}</td><td>{{e.materia}}</td><td>{{e.tipo_asunto}}</td><td>{{e.estado}}</td><td>{{e.propietario.nombre if e.propietario else ""}}</td><td class="no-print"><a class="btn btn2" href="/expediente/{{e.id}}">Abrir</a></td></tr>{% endfor %}</table></div>""", datos=datos, titulo=titulo)


@app.route("/compartidos-conmigo")
def compartidos_conmigo():
    if not req():
        return redirect("/login")
    compartidos = Compartido.query.filter_by(usuario_id=actual().id).order_by(Compartido.id.desc()).all()
    return render("""<div class="card"><h2>Expedientes compartidos conmigo</h2><table><tr><th>Expediente</th><th>Cliente</th><th>Propietario</th><th>Permiso</th><th>Acción</th></tr>{% for c in compartidos %}<tr><td>{{c.expediente.numero}}</td><td>{{c.expediente.cliente.nombre if c.expediente.cliente else ""}}</td><td>{{c.expediente.propietario.nombre}}</td><td><span class="badge">{{c.permiso}}</span></td><td><a class="btn btn2" href="/expediente/{{c.expediente.id}}">Abrir</a></td></tr>{% endfor %}</table></div>""", compartidos=compartidos)


@app.route("/expediente/<int:id>")
def expediente(id):
    if not req():
        return redirect("/login")
    if not puede_ver(id):
        flash("No tienes acceso a este expediente.")
        return redirect("/")
    expediente = Expediente.query.get_or_404(id)
    return render("""<div class="card"><h2>Expediente {{expediente.numero}}{% if puede_administrar %}
<a class="btn btnDanger"
   href="/eliminar-expediente/{{expediente.id}}"
   onclick="return confirm('¿Eliminar expediente completo?')">
   Eliminar Expediente
</a>
{% endif %} <button onclick="window.print()">Imprimir</button></h2><p><b>Cliente:</b> {{expediente.cliente.nombre if expediente.cliente else ""}}</p><p><b>Materia:</b> {{expediente.materia}}</p><p><b>Tipo de asunto:</b> {{expediente.tipo_asunto}}</p><p><b>Autoridad:</b> {{expediente.autoridad}}</p><p><b>Actor:</b> {{expediente.actor}}</p><p><b>Demandado:</b> {{expediente.demandado}}</p><p><b>Estado:</b> {{expediente.estado}}</p>{% if puede_editar %}<form method="post" action="/actualizar-estatus/{{expediente.id}}"><label>Cambiar estatus</label><select name="estado"><option {% if expediente.estado == 'Nuevo' %}selected{% endif %}>Nuevo</option><option {% if expediente.estado == 'En trámite' %}selected{% endif %}>En trámite</option><option {% if expediente.estado == 'Pendiente de acuerdo' %}selected{% endif %}>Pendiente de acuerdo</option><option {% if expediente.estado == 'Pendiente de audiencia' %}selected{% endif %}>Pendiente de audiencia</option><option {% if expediente.estado == 'Concluido' %}selected{% endif %}>Concluido</option><option {% if expediente.estado == 'Archivado' %}selected{% endif %}>Archivado</option><option {% if expediente.estado == 'Urgente' %}selected{% endif %}>Urgente</option></select><button>Actualizar estatus</button></form>{% endif %}<p><b>Tu permiso:</b> <span class="badge">{{permiso}}</span></p><p>{{expediente.observaciones}}</p>{% if puede_editar %}<a class="btn" href="/movimiento/{{expediente.id}}">Agregar promoción/vencimiento</a> <a class="btn btn2" href="/audiencia/{{expediente.id}}">Programar audiencia</a> <a class="btn" href="/mensaje/{{expediente.id}}">Mensaje interno</a>{% endif %} {% if puede_administrar %}<a class="btn btnDark" href="/compartir/{{expediente.id}}">Compartir / Accesos</a>{% endif %} {% if puede_editar %}<a class="btn" href="/aviso-expediente/{{expediente.id}}">Agregar aviso</a>{% endif %}</div><div class="grid2"><div class="card"><h2>Promociones / vencimientos</h2><table><tr><th>Título</th><th>Usuario</th><th>Límite</th><th>Archivo</th></tr>{% for movimiento in movimientos %}<tr><td>{{movimiento.titulo}}</td><td>{{movimiento.usuario.nombre}}</td><td>{{movimiento.fecha_limite}} {{movimiento.hora_limite}}</td><td>{% if movimiento.archivo_url %}<a href="{{movimiento.archivo_url}}" target="_blank">Ver</a>{% endif %}</td></tr>{% endfor %}</table></div><div class="card"><h2>Audiencias</h2><table><tr><th>Fecha</th><th>Hora</th><th>Audiencia</th><th>Usuario</th></tr>{% for audiencia in audiencias %}<tr><td>{{audiencia.fecha}}</td><td>{{audiencia.hora}}</td><td>{{audiencia.titulo}}</td><td>{{audiencia.usuario.nombre}}</td></tr>{% endfor %}</table></div></div><div class="grid2"><div class="card"><h2>Chat interno</h2>{% for m in mensajes %}<div class="chat"><b>{{m.usuario.nombre}}</b><div class="small">{{m.creado_en.strftime("%d/%m/%Y %H:%M")}}</div><p>{{m.mensaje}}</p></div>{% endfor %}</div><div class="card"><h2>Avisos del expediente</h2>{% for av in avisos_exp %}<div class="chat"><b>{{av.titulo}}</b> - {{av.creado_en.strftime('%d/%m/%Y %H:%M')}}<br>{{av.mensaje}}<br>{% if puede_administrar or av.usuario_id == usuario_actual_id %}<a class="btn btnDanger" href="/eliminar-aviso/{{av.id}}" onclick="return confirm('¿Eliminar este aviso del expediente?')">Eliminar aviso</a>{% endif %}</div>{% endfor %}</div></div><div class="card"><h2>Bitácora</h2>{% for b in bitacora %}<div class="bit"><b>{{b.accion}}</b><div class="small">{{b.usuario.nombre if b.usuario else ""}} | {{b.creado_en.strftime("%d/%m/%Y %H:%M")}}</div><p>{{b.detalle}}</p></div>{% endfor %}</div>""", expediente=expediente, permiso=permiso_expediente(id), puede_editar=puede_editar(id), puede_administrar=puede_administrar(id), movimientos=Movimiento.query.filter_by(expediente_id=id).order_by(Movimiento.id.desc()).all(), audiencias=Audiencia.query.filter_by(expediente_id=id).order_by(Audiencia.fecha, Audiencia.hora).all(), mensajes=MensajeExpediente.query.filter_by(expediente_id=id).order_by(MensajeExpediente.id.desc()).all(), bitacora=Bitacora.query.filter_by(expediente_id=id).order_by(Bitacora.id.desc()).all(), avisos_exp=AvisoExpediente.query.filter_by(expediente_id=id).order_by(AvisoExpediente.id.desc()).all(), usuario_actual_id=actual().id)

@app.route("/eliminar-expediente/<int:id>")
def eliminar_expediente(id):
    if not req():
        return redirect("/login")

    expediente = Expediente.query.get_or_404(id)

    if not (admin() or expediente.propietario_id == actual().id):
        flash("No tienes permiso para eliminar este expediente.")
        return redirect("/expediente/" + str(id))

    Movimiento.query.filter_by(expediente_id=id).delete()
    Audiencia.query.filter_by(expediente_id=id).delete()
    MensajeExpediente.query.filter_by(expediente_id=id).delete()
    AvisoExpediente.query.filter_by(expediente_id=id).delete()
    Bitacora.query.filter_by(expediente_id=id).delete()
    Compartido.query.filter_by(expediente_id=id).delete()
    Cobranza.query.filter_by(expediente_id=id).delete()

    db.session.delete(expediente)
    db.session.commit()

    flash("Expediente eliminado correctamente.")
    return redirect("/mis-expedientes")
    
@app.route("/actualizar-estatus/<int:id>", methods=["POST"])
def actualizar_estatus(id):
    if not req():
        return redirect("/login")

    if not puede_editar(id):
        flash("No tienes permiso para cambiar el estatus.")
        return redirect("/expediente/" + str(id))

    expediente = Expediente.query.get_or_404(id)
    nuevo_estatus = request.form["estado"]

    expediente.estado = nuevo_estatus
    registrar_bitacora(id, "Cambió estatus", f"Nuevo estatus: {nuevo_estatus}")
    crear_aviso_para_accesos(id, "Estatus actualizado", f"El expediente {expediente.numero} cambió a: {nuevo_estatus}")

    db.session.commit()
    flash("Estatus actualizado.")
    return redirect("/expediente/" + str(id))

@app.route("/mensaje/<int:expediente_id>", methods=["GET", "POST"])
def mensaje(expediente_id):
    if not req():
        return redirect("/login")
    if not puede_editar(expediente_id):
        flash("No tienes permiso para agregar mensajes.")
        return redirect("/expediente/" + str(expediente_id))
    if request.method == "POST":
        texto = request.form["mensaje"]
        db.session.add(MensajeExpediente(expediente_id=expediente_id, usuario_id=actual().id, mensaje=texto))
        registrar_bitacora(expediente_id, "Agregó mensaje interno", texto[:120])
        crear_aviso_para_accesos(expediente_id, "Nuevo mensaje interno", f"{actual().nombre}: {texto[:220]}")
        db.session.commit()
        flash("Mensaje agregado y avisado.")
        return redirect("/expediente/" + str(expediente_id))
    return render("""<div class="card"><h2>Mensaje interno</h2><form method="post"><label>Mensaje</label><textarea name="mensaje" required></textarea><button>Enviar</button></form></div>""")

@app.route("/eliminar-aviso/<int:id>")
def eliminar_aviso(id):
    if not req():
        return redirect("/login")

    aviso = AvisoExpediente.query.get_or_404(id)

    if not (admin() or aviso.usuario_id == actual().id):
        flash("No tienes permiso para eliminar este aviso.")
        return redirect("/expediente/" + str(aviso.expediente_id))

    expediente_id = aviso.expediente_id

    db.session.delete(aviso)
    db.session.commit()

    flash("Aviso eliminado.")
    return redirect("/expediente/" + str(expediente_id))
    
@app.route("/aviso-expediente/<int:expediente_id>", methods=["GET", "POST"])
def aviso_expediente(expediente_id):
    if not req():
        return redirect("/login")
    if not puede_editar(expediente_id):
        flash("No tienes permiso para agregar avisos.")
        return redirect("/expediente/" + str(expediente_id))
    if request.method == "POST":
        titulo = request.form["titulo"]
        mensaje = request.form["mensaje"]
        crear_aviso_para_accesos(expediente_id, titulo, mensaje)
        registrar_bitacora(expediente_id, "Aviso agregado", titulo)
        db.session.commit()
        flash("Aviso agregado.")
        return redirect("/expediente/" + str(expediente_id))
    return render("""<div class="card"><h2>Agregar aviso al expediente</h2><form method="post"><label>Título</label><input name="titulo" required><label>Mensaje</label><textarea name="mensaje" required></textarea><button>Guardar aviso</button></form></div>""")


@app.route("/movimiento/<int:expediente_id>", methods=["GET", "POST"])
def movimiento(expediente_id):
    if not req():
        return redirect("/login")
    if not puede_editar(expediente_id):
        flash("No tienes permiso para agregar movimientos.")
        return redirect("/expediente/" + str(expediente_id))
    if request.method == "POST":
        db.session.add(Movimiento(
            expediente_id=expediente_id, usuario_id=actual().id, titulo=request.form["titulo"], fecha=request.form["fecha"],
            estatus=request.form["estatus"], proxima_accion=request.form["proxima_accion"], fecha_limite=request.form["fecha_limite"],
            hora_limite=request.form["hora_limite"], observaciones=request.form["observaciones"],
            archivo_url=subir(request.files.get("archivo")), notificar=bool(request.form.get("notificar"))
        ))
        registrar_bitacora(expediente_id, "Agregó promoción/vencimiento", request.form["titulo"])
        crear_aviso_para_accesos(expediente_id, "Nuevo movimiento", f"{actual().nombre} agregó: {request.form['titulo']}")
        db.session.commit()
        flash("Guardado.")
        return redirect("/expediente/" + str(expediente_id))
    return render("""<div class="card"><h2>Agregar promoción / vencimiento</h2><form method="post" enctype="multipart/form-data"><label>Título</label><input name="titulo" required><label>Fecha</label><input type="date" name="fecha"><label>Estatus</label><select name="estatus"><option>Elaborado</option><option>Presentado</option><option>Acordado</option><option>Pendiente</option><option>Concluido</option></select><label>Próxima acción</label><input name="proxima_accion"><label>Fecha límite</label><input type="date" name="fecha_limite"><label>Hora límite</label><input type="time" name="hora_limite" value="09:00"><label><input type="checkbox" name="notificar" checked style="width:auto"> Notificar 24h y 2h antes</label><label>Observaciones</label><textarea name="observaciones"></textarea><label>Archivo</label><input type="file" name="archivo"><button>Guardar</button></form></div>""")


@app.route("/audiencia/<int:expediente_id>", methods=["GET", "POST"])
def audiencia(expediente_id):
    if not req():
        return redirect("/login")
    if not puede_editar(expediente_id):
        flash("No tienes permiso para programar audiencias.")
        return redirect("/expediente/" + str(expediente_id))
    if request.method == "POST":
        db.session.add(Audiencia(
            expediente_id=expediente_id, usuario_id=actual().id, titulo=request.form["titulo"], fecha=request.form["fecha"],
            hora=request.form["hora"], autoridad=request.form["autoridad"], sala=request.form["sala"],
            modalidad=request.form["modalidad"], enlace=request.form["enlace"], observaciones=request.form["observaciones"],
            notificar=bool(request.form.get("notificar"))
        ))
        registrar_bitacora(expediente_id, "Programó audiencia", f'{request.form["titulo"]} - {request.form["fecha"]} {request.form["hora"]}')
        crear_aviso_para_accesos(expediente_id, "Audiencia programada", f"{request.form['titulo']} - {request.form['fecha']} {request.form['hora']}")
        db.session.commit()
        flash("Audiencia programada.")
        return redirect("/expediente/" + str(expediente_id))
    return render("""<div class="card"><h2>Programar audiencia</h2><form method="post"><label>Audiencia</label><input name="titulo" required><label>Fecha</label><input type="date" name="fecha" required><label>Hora</label><input type="time" name="hora" required><label>Autoridad</label><input name="autoridad"><label>Sala</label><input name="sala"><label>Modalidad</label><select name="modalidad"><option>Presencial</option><option>Virtual</option><option>Mixta</option></select><label>Enlace</label><input name="enlace"><label><input type="checkbox" name="notificar" checked style="width:auto"> Notificar 24h y 2h antes</label><label>Observaciones</label><textarea name="observaciones"></textarea><button>Guardar</button></form></div>""")


@app.route("/compartir/<int:expediente_id>", methods=["GET", "POST"])
def compartir(expediente_id):
    if not req():
        return redirect("/login")
    if not puede_administrar(expediente_id):
        flash("No tienes permiso para compartir este expediente.")
        return redirect("/expediente/" + str(expediente_id))
    expediente = Expediente.query.get_or_404(expediente_id)
    if request.method == "POST":
        usuario_id = int(request.form["usuario_id"])
        permiso = request.form["permiso"]
        existente = Compartido.query.filter_by(expediente_id=expediente_id, usuario_id=usuario_id).first()
        if existente:
            existente.permiso = permiso
            accion = "Actualizó permiso"
        else:
            db.session.add(Compartido(expediente_id=expediente_id, usuario_id=usuario_id, permiso=permiso))
            accion = "Compartió expediente"
        destino = db.session.get(Usuario, usuario_id)
        registrar_bitacora(expediente_id, accion, f"Usuario: {destino.nombre} | Permiso: {permiso}")
        notificar_interna(usuario_id, "Expediente compartido", f"{actual().nombre} compartió contigo el expediente {expediente.numero} con permiso {permiso}.", f"/expediente/{expediente_id}")
        if destino and destino.correo:
            correo(destino.correo, "ALESI - Expediente compartido", f"{actual().nombre} compartió contigo el expediente {expediente.numero}.\nPermiso: {permiso}\n\nALESI GRUPO JURÍDICO")
        db.session.commit()
        flash("Expediente compartido.")
        return redirect("/compartir/" + str(expediente_id))
    accesos = Compartido.query.filter_by(expediente_id=expediente_id).all()
    usuarios = Usuario.query.filter(Usuario.activo == True, Usuario.id != expediente.propietario_id).order_by(Usuario.nombre).all()
    return render("""<div class="grid2"><div class="card"><h2>Compartir expediente {{expediente.numero}}</h2><form method="post"><label>Usuario registrado</label><select name="usuario_id">{% for usuario in usuarios %}<option value="{{usuario.id}}">{{usuario.nombre}} - {{usuario.usuario}} - {{usuario.area}}</option>{% endfor %}</select><label>Permiso</label><select name="permiso"><option>Lectura</option><option>Edición</option><option>Administración</option></select><button>Compartir</button></form><br><a class="btn btn2" href="/compartir-area/{{expediente.id}}">Compartir por área</a></div><div class="card"><h2>Usuarios con acceso</h2><p><b>Propietario:</b> {{expediente.propietario.nombre}} <span class="badge">Administración</span></p><table><tr><th>Usuario</th><th>Permiso</th><th>Acción</th></tr>{% for a in accesos %}<tr><td>{{a.usuario.nombre}}</td><td><span class="badge">{{a.permiso}}</span></td><td><a class="btn btnDark" href="/quitar-acceso/{{a.id}}">Quitar</a></td></tr>{% endfor %}</table></div></div>""", expediente=expediente, usuarios=usuarios, accesos=accesos)


@app.route("/compartir-area/<int:expediente_id>", methods=["GET", "POST"])
def compartir_area(expediente_id):
    if not req():
        return redirect("/login")
    if not puede_administrar(expediente_id):
        flash("No tienes permiso para compartir por área.")
        return redirect("/expediente/" + str(expediente_id))
    expediente = Expediente.query.get_or_404(expediente_id)
    areas = sorted({u.area for u in Usuario.query.filter_by(activo=True).all() if u.area})
    if request.method == "POST":
        area = request.form["area"]
        permiso = request.form["permiso"]
        usuarios_area = Usuario.query.filter_by(area=area, activo=True).all()
        for usuario in usuarios_area:
            if usuario.id == expediente.propietario_id:
                continue
            existente = Compartido.query.filter_by(expediente_id=expediente_id, usuario_id=usuario.id).first()
            if existente:
                existente.permiso = permiso
            else:
                db.session.add(Compartido(expediente_id=expediente_id, usuario_id=usuario.id, permiso=permiso))
            notificar_interna(usuario.id, "Expediente compartido por área", f"Se compartió contigo el expediente {expediente.numero} por pertenecer al área {area}.", f"/expediente/{expediente_id}")
        registrar_bitacora(expediente_id, "Compartió expediente por área", f"Área: {area} | Permiso: {permiso}")
        db.session.commit()
        flash("Expediente compartido por área.")
        return redirect("/compartir/" + str(expediente_id))
    return render("""<div class="card"><h2>Compartir expediente por área</h2><form method="post"><label>Área</label><select name="area">{% for area in areas %}<option>{{area}}</option>{% endfor %}</select><label>Permiso</label><select name="permiso"><option>Lectura</option><option>Edición</option><option>Administración</option></select><button>Compartir por área</button></form></div>""", areas=areas)


@app.route("/quitar-acceso/<int:compartido_id>")
def quitar_acceso(compartido_id):
    if not req():
        return redirect("/login")
    compartido = Compartido.query.get_or_404(compartido_id)
    expediente_id = compartido.expediente_id
    if not puede_administrar(expediente_id):
        flash("No tienes permiso para quitar accesos.")
        return redirect("/expediente/" + str(expediente_id))
    nombre = compartido.usuario.nombre
    db.session.delete(compartido)
    registrar_bitacora(expediente_id, "Quitó acceso", f"Usuario: {nombre}")
    db.session.commit()
    flash("Acceso retirado.")
    return redirect("/compartir/" + str(expediente_id))


@app.route("/audiencias")
def audiencias():
    if not req():
        return redirect("/login")
    datos = Audiencia.query.order_by(Audiencia.fecha, Audiencia.hora).all()
    return render("""<div class="card"><h2>Audiencias <button onclick="window.print()">Imprimir</button></h2><table><tr><th>Fecha</th><th>Hora</th><th>Expediente</th><th>Audiencia</th><th>Autoridad</th></tr>{% for a in datos %}<tr><td>{{a.fecha}}</td><td>{{a.hora}}</td><td>{{a.expediente.numero}}</td><td>{{a.titulo}}</td><td>{{a.autoridad}}</td></tr>{% endfor %}</table></div>""", datos=datos)


@app.route("/vencimientos")
def vencimientos():
    if not req():
        return redirect("/login")
    datos = Movimiento.query.filter(Movimiento.fecha_limite != "").order_by(Movimiento.fecha_limite, Movimiento.hora_limite).all()
    return render("""<div class="card"><h2>Vencimientos <button onclick="window.print()">Imprimir</button></h2><table><tr><th>Fecha</th><th>Hora</th><th>Expediente</th><th>Movimiento</th><th>Acción</th></tr>{% for m in datos %}<tr><td>{{m.fecha_limite}}</td><td>{{m.hora_limite}}</td><td>{{m.expediente.numero}}</td><td>{{m.titulo}}</td><td>{{m.proxima_accion}}</td></tr>{% endfor %}</table></div>""", datos=datos)


@app.route("/cobranza", methods=["GET", "POST"])
def cobranza():
    if not req():
        return redirect("/login")
    if request.method == "POST":
        cliente_id = request.form.get("cliente_id") or None
        cliente = db.session.get(Cliente, int(cliente_id)) if cliente_id else None
        total = float(request.form["monto_total"] or 0)
        abono = float(request.form["abono"] or 0)
        pagos_previos = db.session.query(db.func.sum(Cobranza.abono)).filter_by(cliente_id=cliente.id).scalar() or 0 if cliente else 0
        saldo = max(total - pagos_previos - abono, 0)
        folio = "ALESI-" + datetime.now().strftime("%Y%m%d%H%M%S")
        p = Cobranza(
            folio=folio, cliente_id=cliente_id, expediente_id=request.form.get("expediente_id") or None,
            usuario_id=actual().id, concepto=request.form["concepto"], descripcion=request.form["descripcion"],
            monto_total=total, abono=abono, saldo=saldo, moneda=request.form["moneda"],
            forma_pago=request.form["forma_pago"], fecha_pago=request.form["fecha_pago"],
            proximo_pago="" if saldo <= 0 else request.form["proximo_pago"], estatus="Liquidado" if saldo <= 0 else "Pendiente",
            comprobante_url=subir(request.files.get("comprobante"))
        )
        db.session.add(p)
        db.session.commit()
        flash("Cobranza guardada. Recibo generado.")
        return redirect("/recibo/" + str(p.id))
    clientes_info = []
    for c in Cliente.query.order_by(Cliente.nombre).all():
        abonado = db.session.query(db.func.sum(Cobranza.abono)).filter_by(cliente_id=c.id).scalar() or 0
        total = c.cobro_total or 0
        clientes_info.append({"id": c.id, "nombre": c.nombre, "cobro_total": total, "moneda": c.moneda_cobro or "MXN", "abonado": abonado, "deuda": max(total - abonado, 0)})
    return render("""<div class="card no-print no-imprimir"><h2>Registrar cobranza</h2><form method="post" enctype="multipart/form-data"><label>Cliente</label><select name="cliente_id" onchange="llenarCobro(this)"><option value="">Seleccione</option>{% for c in clientes_info %}<option value="{{c.id}}" data-total="{{c.cobro_total}}" data-moneda="{{c.moneda}}" data-deuda="{{c.deuda}}">{{c.nombre}} | Cobro: {{c.moneda}} {{c.cobro_total}} | Debe: {{c.deuda}}</option>{% endfor %}</select><label>Expediente</label><select name="expediente_id"><option value="">Sin expediente</option>{% for e in expedientes %}<option value="{{e.id}}">{{e.numero}}</option>{% endfor %}</select><label>Concepto</label><select name="concepto"><option>Honorarios</option><option>Abono</option><option>Pago trámite institucional</option><option>Copias certificadas</option><option>Notario</option><option>Viáticos</option><option>Derechos</option><option>Gastos administrativos</option><option>Otro</option></select><label>Descripción</label><textarea name="descripcion"></textarea><label>Total cobrado / pactado</label><input id="monto_total" type="number" step="0.01" name="monto_total"><label>Abono que pagará</label><input type="number" step="0.01" name="abono"><label>Deuda actual antes del abono</label><input id="deuda_actual" disabled><label>Moneda</label><select id="moneda" name="moneda"><option>MXN</option><option>USD</option></select><label>Forma de pago</label><select name="forma_pago"><option>Efectivo</option><option>Tarjeta</option><option>Transferencia</option></select><label>Fecha de pago</label><input type="date" name="fecha_pago"><label>Próximo pago</label><input type="date" name="proximo_pago"><label>Comprobante</label><input type="file" name="comprobante"><button>Guardar y generar recibo PDF</button></form><script>function llenarCobro(sel){var o=sel.options[sel.selectedIndex];document.getElementById('monto_total').value=o.dataset.total||'';document.getElementById('deuda_actual').value=o.dataset.deuda||'';if(o.dataset.moneda){document.getElementById('moneda').value=o.dataset.moneda;}}</script></div><div class="card print-list"><h2>Lista de pagos realizados <button class="no-print" onclick="window.print()">Imprimir lista</button></h2><table><tr><th>Folio</th><th>Cliente</th><th>Concepto</th><th>Abono</th><th>Saldo</th><th>Moneda</th><th>Próximo pago</th><th class="no-print">Recibo</th></tr>{% for p in datos %}<tr><td>{{p.folio}}</td><td>{{p.cliente.nombre if p.cliente else ""}}</td><td>{{p.concepto}}</td><td>{{p.abono}}</td><td>{{p.saldo}}</td><td>{{p.moneda}}</td><td>{{p.proximo_pago}}</td><td class="no-print"><a href="/recibo/{{p.id}}">Ver</a> | <a href="/recibo-pdf/{{p.id}}">PDF</a></td></tr>{% endfor %}</table></div>""", clientes_info=clientes_info, expedientes=Expediente.query.order_by(Expediente.numero).all(), datos=Cobranza.query.order_by(Cobranza.id.desc()).all())


def recibo_html(p):
    return render_template_string("""<!doctype html><html><head><meta charset="utf-8"><title>Recibo {{p.folio}}</title><style>body{font-family:Arial;margin:35px;color:#111827}.recibo{border:2px solid #111827;padding:25px;border-radius:12px}.head{display:flex;align-items:center;border-bottom:3px solid #B38B2E;padding-bottom:15px}.logo{width:90px;height:90px;border-radius:50%;object-fit:cover;margin-right:18px}.grid{display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-top:18px}.box{border:1px solid #ddd;padding:12px;border-radius:8px}.firma{margin-top:60px;border-top:1px solid #111;width:320px;text-align:center;padding-top:8px}.no-print{margin-bottom:20px}@media print{.no-print{display:none}}</style></head><body><div class="no-print"><button onclick="window.print()">Imprimir</button> <a href="/recibo-pdf/{{p.id}}">Descargar PDF</a></div><div class="recibo"><div class="head"><img class="logo" src="/logo"><div><h1>ALESI GRUPO JURÍDICO</h1><p>Recibo de pago</p></div></div><h2>Folio: {{p.folio}}</h2><div class="grid"><div class="box"><b>Fecha:</b> {{p.fecha_pago or p.creado_en.strftime('%Y-%m-%d')}}</div><div class="box"><b>Cliente:</b> {{p.cliente.nombre if p.cliente else ''}}</div><div class="box"><b>Expediente:</b> {{p.expediente.numero if p.expediente else ''}}</div><div class="box"><b>Concepto:</b> {{p.concepto}}</div><div class="box"><b>Monto total:</b> {{p.moneda}} {{'%.2f'|format(p.monto_total or 0)}}</div><div class="box"><b>Abono recibido:</b> {{p.moneda}} {{'%.2f'|format(p.abono or 0)}}</div><div class="box"><b>Saldo pendiente:</b> {{p.moneda}} {{'%.2f'|format(p.saldo or 0)}}</div><div class="box"><b>Forma de pago:</b> {{p.forma_pago}}</div><div class="box"><b>Próximo pago:</b> {{p.proximo_pago}}</div><div class="box"><b>Estatus:</b> {{p.estatus}}</div></div><p><b>Descripción / observaciones:</b><br>{{p.descripcion}}</p><p><b>Recibió:</b> {{p.usuario.nombre if p.usuario else ''}}</p><div class="firma">Firma de recibido</div></div></body></html>""", p=p)


@app.route("/recibo/<int:id>")
def recibo(id):
    if not req():
        return redirect("/login")
    p = Cobranza.query.get_or_404(id)
    return recibo_html(p)


def crear_recibo_pdf(p):
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=40, leftMargin=40, topMargin=40, bottomMargin=40)
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle("AlesiTitle", parent=styles["Title"], textColor=colors.HexColor("#111827"), fontSize=20, spaceAfter=8)
    subtitle_style = ParagraphStyle("AlesiSub", parent=styles["Normal"], textColor=colors.HexColor("#B38B2E"), fontSize=12, spaceAfter=14)
    story = [Paragraph("ALESI GRUPO JURÍDICO", title_style), Paragraph("Recibo de pago", subtitle_style), Spacer(1, 10)]
    data = [
        ["Folio", p.folio or ""], ["Fecha", p.fecha_pago or p.creado_en.strftime("%Y-%m-%d")],
        ["Cliente", p.cliente.nombre if p.cliente else ""], ["Expediente", p.expediente.numero if p.expediente else ""],
        ["Concepto", p.concepto or ""], ["Monto total", f"{p.moneda or ''} {p.monto_total or 0:,.2f}"],
        ["Abono recibido", f"{p.moneda or ''} {p.abono or 0:,.2f}"], ["Saldo pendiente", f"{p.moneda or ''} {p.saldo or 0:,.2f}"],
        ["Forma de pago", p.forma_pago or ""], ["Próximo pago", p.proximo_pago or ""], ["Estatus", p.estatus or ""],
        ["Recibió", p.usuario.nombre if p.usuario else ""],
    ]
    table = Table(data, colWidths=[1.8*inch, 4.8*inch])
    table.setStyle(TableStyle([("BACKGROUND",(0,0),(0,-1),colors.HexColor("#111827")),("TEXTCOLOR",(0,0),(0,-1),colors.white),("BACKGROUND",(1,0),(1,-1),colors.whitesmoke),("BOX",(0,0),(-1,-1),1,colors.HexColor("#B38B2E")),("INNERGRID",(0,0),(-1,-1),0.5,colors.lightgrey),("FONTNAME",(0,0),(0,-1),"Helvetica-Bold"),("VALIGN",(0,0),(-1,-1),"TOP"),("PADDING",(0,0),(-1,-1),8)]))
    story.append(table)
    story.append(Spacer(1, 18))
    story.append(Paragraph("<b>Descripción / observaciones:</b>", styles["Normal"]))
    story.append(Paragraph((p.descripcion or "").replace("\\n","<br/>"), styles["Normal"]))
    story.append(Spacer(1, 60))
    firma = Table([["Firma de recibido"]], colWidths=[3.5*inch])
    firma.setStyle(TableStyle([("LINEABOVE",(0,0),(-1,0),1,colors.black),("ALIGN",(0,0),(-1,-1),"CENTER"),("TOPPADDING",(0,0),(-1,-1),8)]))
    story.append(firma)
    doc.build(story)
    buffer.seek(0)
    return buffer


@app.route("/recibo-pdf/<int:id>")
def recibo_pdf(id):
    if not req():
        return redirect("/login")
    p = Cobranza.query.get_or_404(id)
    return send_file(crear_recibo_pdf(p), mimetype="application/pdf", as_attachment=True, download_name=f"recibo_{p.folio or p.id}.pdf")


@app.route("/planeador")
def planeador():
    if not req():
        return redirect("/login")
    eventos = []
    dias_semana = ["Lunes","Martes","Miércoles","Jueves","Viernes","Sábado","Domingo"]
    for a in Audiencia.query.all():
        if a.fecha:
            eventos.append({"fecha": a.fecha, "hora": a.hora or "", "tipo": "Audiencia", "texto": f"{a.expediente.numero if a.expediente else ''} - {a.titulo}"})
    for v in Movimiento.query.filter(Movimiento.fecha_limite != "").all():
        if v.fecha_limite:
            eventos.append({"fecha": v.fecha_limite, "hora": v.hora_limite or "", "tipo": "Vencimiento", "texto": f"{v.expediente.numero if v.expediente else ''} - {v.proxima_accion or v.titulo}"})
    for p in Cobranza.query.filter(Cobranza.proximo_pago != "", Cobranza.saldo > 0, Cobranza.estatus != "Liquidado").all():
        if p.proximo_pago:
            eventos.append({"fecha": p.proximo_pago, "hora": "", "tipo": "Pago", "texto": f"{p.cliente.nombre if p.cliente else ''} - {p.moneda} {p.saldo}"})
    eventos.sort(key=lambda x: (x["fecha"], x["hora"]))
    agrupado = {}
    for ev in eventos:
        agrupado.setdefault(ev["fecha"], []).append(ev)
    dias = {f: dias_semana[datetime.strptime(f, "%Y-%m-%d").weekday()] if f else "" for f in agrupado}
    return render("""<div class="card no-print"><h2>Planeador jurídico</h2><p>Este módulo imprime únicamente el calendario de audiencias, vencimientos y pagos.</p><button onclick="window.print()">Imprimir calendario</button></div><div class="card solo-imprimir"><h2>Calendario ALESI Grupo Jurídico</h2><table><tr><th style="width:170px">Día</th><th>Actividades</th></tr>{% for fecha, lista in agrupado.items() %}<tr><td><b>{{dias[fecha]}}</b><br>{{fecha}}</td><td>{% for ev in lista %}<div style="margin-bottom:8px"><b>{{ev.tipo}}</b> {% if ev.hora %}{{ev.hora}}{% endif %}<br>{{ev.texto}}</div>{% endfor %}</td></tr>{% endfor %}</table></div>""", agrupado=agrupado, dias=dias)


@app.route("/notificaciones-internas", methods=["GET", "POST"])
def notificaciones_internas():
    if not req():
        return redirect("/login")
    if request.method == "POST":
        titulo = request.form["titulo"]
        mensaje = request.form["mensaje"]
        usuario_id = request.form.get("usuario_id")
        area = request.form.get("area", "")
        expediente_id = request.form.get("expediente_id") or None
        destinos = []
        if usuario_id:
            u = db.session.get(Usuario, int(usuario_id))
            if u:
                destinos.append(u)
        if area:
            destinos += Usuario.query.filter_by(area=area, activo=True).all()
        if not destinos:
            destinos = [actual()]
        for u in destinos:
            notificar_interna(u.id, titulo, mensaje, f"/expediente/{expediente_id}" if expediente_id else "")
        if expediente_id:
            db.session.add(AvisoExpediente(expediente_id=expediente_id, usuario_id=actual().id, titulo=titulo, mensaje=mensaje, fecha_aviso=datetime.now().strftime("%Y-%m-%d")))
            registrar_bitacora(int(expediente_id), "Aviso agregado", titulo + " - " + mensaje[:150])
        db.session.commit()
        flash("Aviso enviado.")
        return redirect("/notificaciones-internas")
    avisos = NotificacionInterna.query.filter_by(usuario_id=actual().id).order_by(NotificacionInterna.id.desc()).all()
    return render("""<div class="grid2"><div class="card"><h2>Agregar aviso interno</h2><form method="post"><label>Título</label><input name="titulo" required><label>Mensaje</label><textarea name="mensaje" required></textarea><label>Enviar a usuario</label><select name="usuario_id"><option value="">Ninguno</option>{% for u in usuarios %}<option value="{{u.id}}">{{u.nombre}} - {{u.usuario}}</option>{% endfor %}</select><label>O enviar a área</label><select name="area"><option value="">Ninguna</option><option>Jurídico</option><option>Administración</option><option>Oficina</option><option>Penal</option><option>Laboral</option><option>Administrativo</option><option>Civil</option><option>Mercantil</option><option>Familiar</option></select><label>Relacionar con expediente</label><select name="expediente_id"><option value="">Sin expediente</option>{% for e in expedientes %}<option value="{{e.id}}">{{e.numero}}</option>{% endfor %}</select><button>Enviar aviso</button></form></div><div class="card"><h2>Avisos internos</h2><p><a class="btn btn2" href="/avisos-marcar-leidos">Marcar todos como leídos</a> <a class="btn btnDanger" href="/eliminar-avisos-leidos" onclick="return confirm('¿Eliminar todos tus avisos leídos?')">Eliminar avisos leídos</a></p>{% for a in avisos %}<div class="chat"><b>{{a.titulo}}</b> <span class="small">{{a.creado_en.strftime('%d/%m/%Y %H:%M')}}</span> {% if a.leida %}<span class="badge">Leído</span>{% else %}<span class="badge">Nuevo</span>{% endif %}<br>{{a.mensaje}}<br>{% if a.enlace %}<a class="btn btn2" href="{{a.enlace}}">Abrir</a>{% endif %} <a class="btn btnDanger" href="/eliminar-notificacion/{{a.id}}" onclick="return confirm('¿Eliminar este aviso?')">Eliminar</a></div>{% endfor %}</div></div>""", avisos=avisos, usuarios=Usuario.query.filter_by(activo=True).order_by(Usuario.nombre).all(), expedientes=Expediente.query.order_by(Expediente.numero).all())


@app.route("/avisos-marcar-leidos")
def avisos_marcar_leidos():
    if not req():
        return redirect("/login")
    for aviso in NotificacionInterna.query.filter_by(usuario_id=actual().id).all():
        aviso.leida = True
    db.session.commit()
    flash("Avisos marcados como leídos.")
    return redirect("/notificaciones-internas")


@app.route("/eliminar-notificacion/<int:id>")
def eliminar_notificacion(id):
    if not req():
        return redirect("/login")
    aviso = NotificacionInterna.query.get_or_404(id)
    if aviso.usuario_id != actual().id and not admin():
        flash("No tienes permiso para eliminar este aviso.")
        return redirect("/notificaciones-internas")
    db.session.delete(aviso)
    db.session.commit()
    flash("Aviso eliminado.")
    return redirect("/notificaciones-internas")


@app.route("/eliminar-avisos-leidos")
def eliminar_avisos_leidos():
    if not req():
        return redirect("/login")
    NotificacionInterna.query.filter_by(usuario_id=actual().id, leida=True).delete()
    db.session.commit()
    flash("Avisos leídos eliminados.")
    return redirect("/notificaciones-internas")


@app.route("/jurisprudencia")
def jurisprudencia():
    if not req():
        return redirect("/login")
    q = request.args.get("q", "").strip()
    enlace = ""
    if q:
        enlace = "https://www.google.com/search?q=" + quote_plus(q + " jurisprudencia Suprema Corte Semanario Judicial site:scjn.gob.mx OR site:sjf2.scjn.gob.mx")
    return render("""<div class="card"><h2>Jurisprudencia</h2><form><label>Concepto jurídico</label><input name="q" value="{{q}}" placeholder="Ejemplo: pensión alimenticia, despido injustificado"><button>Preparar búsqueda ampliada</button></form>{% if q %}<br><a class="btn btn2" target="_blank" href="{{enlace}}">Abrir búsqueda ampliada</a>{% endif %}</div>""", q=q, enlace=enlace)


@app.route("/boletin-judicial")
def boletin_judicial():
    if not req():
        return redirect("/login")
    return render("""<div class="card"><h2>Boletín Judicial de Baja California</h2><p>Acceso directo al Boletín Judicial.</p><a class="btn btn2" target="_blank" href="https://www.pjbc.gob.mx/boletin_judicial.aspx">Abrir Boletín Judicial BC</a></div>""")


@app.route("/teja")
def teja():
    if not req():
        return redirect("/login")
    return render("""<div class="card"><h2>TEJA Baja California</h2><p>Accesos directos al Tribunal Estatal de Justicia Administrativa de Baja California.</p><a class="btn" target="_blank" href="https://tejabc.mx/">Abrir TEJA BC</a> <a class="btn btn2" target="_blank" href="https://tejabc.mx/buscarlistas">Buscar listas TEJA</a></div>""")


@app.route("/notificaciones")
def notificaciones():
    if not req() or not admin():
        return redirect("/")
    total = revisar_y_enviar_notificaciones()
    return render("""<div class="card"><h2>Notificaciones</h2><p>Procesadas: <b>{{total}}</b></p></div>""", total=total)


# =========================
# RESPALDOS MANUALES EN EXCEL
# =========================

@app.route("/respaldos")
def respaldos():
    if not req():
        return redirect("/login")
    if not admin():
        flash("Solo el administrador puede descargar respaldos.")
        return redirect("/")

    return render("""
    <div class="card">
        <h2>Respaldos del Sistema</h2>
        <p>Descarga una copia completa de la información del sistema en Excel para guardarla en tu computadora.</p>
        <br>
        <a class="btn" href="/descargar-respaldo">Descargar respaldo completo en Excel</a>
    </div>
    """)


@app.route("/descargar-respaldo")
def descargar_respaldo():
    if not req():
        return redirect("/login")
    if not admin():
        flash("Solo el administrador puede descargar respaldos.")
        return redirect("/")

    wb = Workbook()

    # CLIENTES
    ws = wb.active
    ws.title = "Clientes"
    ws.append(["ID", "Nombre", "Teléfono", "Correo", "RFC", "CURP", "Materia", "Tipo asunto", "Cobro total", "Moneda", "Dirección", "Observaciones", "Creado en"])
    for c in Cliente.query.order_by(Cliente.id).all():
        ws.append([
            c.id,
            c.nombre,
            c.telefono,
            c.correo,
            c.rfc,
            c.curp,
            c.materia,
            c.tipo_asunto,
            c.cobro_total,
            c.moneda_cobro,
            c.direccion,
            c.observaciones,
            c.creado_en.strftime("%d/%m/%Y %H:%M") if c.creado_en else "",
        ])

    # EXPEDIENTES
    ws = wb.create_sheet("Expedientes")
    ws.append(["ID", "Número", "Cliente", "Materia", "Tipo asunto", "Autoridad", "Actor", "Demandado", "Estado", "Prioridad", "Responsable", "Fecha inicio", "Observaciones", "Propietario", "Creado en"])
    for e in Expediente.query.order_by(Expediente.id).all():
        ws.append([
            e.id,
            e.numero,
            e.cliente.nombre if e.cliente else "",
            e.materia,
            e.tipo_asunto,
            e.autoridad,
            e.actor,
            e.demandado,
            e.estado,
            e.prioridad,
            e.responsable,
            e.fecha_inicio,
            e.observaciones,
            e.propietario.nombre if e.propietario else "",
            e.creado_en.strftime("%d/%m/%Y %H:%M") if e.creado_en else "",
        ])

    # COBRANZA
    ws = wb.create_sheet("Cobranza")
    ws.append(["ID", "Folio", "Cliente", "Expediente", "Usuario", "Concepto", "Descripción", "Monto total", "Abono", "Saldo", "Moneda", "Forma de pago", "Fecha pago", "Próximo pago", "Estatus", "Comprobante", "Creado en"])
    for p in Cobranza.query.order_by(Cobranza.id).all():
        ws.append([
            p.id,
            p.folio,
            p.cliente.nombre if p.cliente else "",
            p.expediente.numero if p.expediente else "",
            p.usuario.nombre if p.usuario else "",
            p.concepto,
            p.descripcion,
            p.monto_total,
            p.abono,
            p.saldo,
            p.moneda,
            p.forma_pago,
            p.fecha_pago,
            p.proximo_pago,
            p.estatus,
            p.comprobante_url,
            p.creado_en.strftime("%d/%m/%Y %H:%M") if p.creado_en else "",
        ])

    # AUDIENCIAS
    ws = wb.create_sheet("Audiencias")
    ws.append(["ID", "Expediente", "Usuario", "Título", "Fecha", "Hora", "Autoridad", "Sala", "Modalidad", "Enlace", "Observaciones", "Notificar", "Creado en"])
    for a in Audiencia.query.order_by(Audiencia.id).all():
        ws.append([
            a.id,
            a.expediente.numero if a.expediente else "",
            a.usuario.nombre if a.usuario else "",
            a.titulo,
            a.fecha,
            a.hora,
            a.autoridad,
            a.sala,
            a.modalidad,
            a.enlace,
            a.observaciones,
            "Sí" if a.notificar else "No",
            a.creado_en.strftime("%d/%m/%Y %H:%M") if a.creado_en else "",
        ])

    # VENCIMIENTOS / MOVIMIENTOS
    ws = wb.create_sheet("Vencimientos")
    ws.append(["ID", "Expediente", "Usuario", "Título", "Fecha", "Estatus", "Próxima acción", "Fecha límite", "Hora límite", "Observaciones", "Archivo", "Notificar", "Creado en"])
    for m in Movimiento.query.order_by(Movimiento.id).all():
        ws.append([
            m.id,
            m.expediente.numero if m.expediente else "",
            m.usuario.nombre if m.usuario else "",
            m.titulo,
            m.fecha,
            m.estatus,
            m.proxima_accion,
            m.fecha_limite,
            m.hora_limite,
            m.observaciones,
            m.archivo_url,
            "Sí" if m.notificar else "No",
            m.creado_en.strftime("%d/%m/%Y %H:%M") if m.creado_en else "",
        ])

    # AVISOS INTERNOS
    ws = wb.create_sheet("Avisos internos")
    ws.append(["ID", "Usuario", "Título", "Mensaje", "Enlace", "Leída", "Creado en"])
    for n in NotificacionInterna.query.order_by(NotificacionInterna.id).all():
        ws.append([
            n.id,
            n.usuario.nombre if n.usuario else "",
            n.titulo,
            n.mensaje,
            n.enlace,
            "Sí" if n.leida else "No",
            n.creado_en.strftime("%d/%m/%Y %H:%M") if n.creado_en else "",
        ])

    # AVISOS DE EXPEDIENTE
    ws = wb.create_sheet("Avisos expediente")
    ws.append(["ID", "Expediente", "Usuario", "Título", "Mensaje", "Fecha aviso", "Creado en"])
    for av in AvisoExpediente.query.order_by(AvisoExpediente.id).all():
        ws.append([
            av.id,
            av.expediente.numero if av.expediente else "",
            av.usuario.nombre if av.usuario else "",
            av.titulo,
            av.mensaje,
            av.fecha_aviso,
            av.creado_en.strftime("%d/%m/%Y %H:%M") if av.creado_en else "",
        ])

    # USUARIOS
    ws = wb.create_sheet("Usuarios")
    ws.append(["ID", "Nombre", "Usuario", "Correo", "Rol", "Área", "Foto", "Activo", "Creado en"])
    for u in Usuario.query.order_by(Usuario.id).all():
        ws.append([
            u.id,
            u.nombre,
            u.usuario,
            u.correo,
            u.rol,
            u.area,
            u.foto_url,
            "Sí" if u.activo else "No",
            u.creado_en.strftime("%d/%m/%Y %H:%M") if u.creado_en else "",
        ])

    # EXPEDIENTES COMPARTIDOS
    ws = wb.create_sheet("Compartidos")
    ws.append(["ID", "Expediente", "Usuario con acceso", "Permiso", "Creado en"])
    for comp in Compartido.query.order_by(Compartido.id).all():
        ws.append([
            comp.id,
            comp.expediente.numero if comp.expediente else "",
            comp.usuario.nombre if comp.usuario else "",
            comp.permiso,
            comp.creado_en.strftime("%d/%m/%Y %H:%M") if comp.creado_en else "",
        ])

    # BITÁCORA
    ws = wb.create_sheet("Bitacora")
    ws.append(["ID", "Expediente", "Usuario", "Acción", "Detalle", "Creado en"])
    for b in Bitacora.query.order_by(Bitacora.id).all():
        ws.append([
            b.id,
            b.expediente.numero if b.expediente else "",
            b.usuario.nombre if b.usuario else "",
            b.accion,
            b.detalle,
            b.creado_en.strftime("%d/%m/%Y %H:%M") if b.creado_en else "",
        ])

    # Ajuste visual básico de columnas
    for hoja in wb.worksheets:
        for columna in hoja.columns:
            letra = columna[0].column_letter
            maximo = 12
            for celda in columna:
                if celda.value is not None:
                    maximo = max(maximo, min(len(str(celda.value)) + 2, 45))
            hoja.column_dimensions[letra].width = maximo

    archivo = BytesIO()
    wb.save(archivo)
    archivo.seek(0)

    nombre = "ALESI_RESPALDO_" + datetime.now().strftime("%Y%m%d_%H%M") + ".xlsx"

    return send_file(
        archivo,
        as_attachment=True,
        download_name=nombre,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


@app.route("/health")
def health():
    return "OK", 200


init_db()


if __name__ == "__main__":
    import sys
    if len(sys.argv) > 1 and sys.argv[1] == "notificaciones":
        print("Notificaciones procesadas:", revisar_y_enviar_notificaciones())
    else:
        app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)), debug=False)
