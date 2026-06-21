
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
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image, PageBreak
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
    cobro_total = db.Column(db.Float, default=0)
    moneda_cobro = db.Column(db.String(10), default="MXN")
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


class ContratoCliente(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    cliente_id = db.Column(db.Integer, db.ForeignKey("cliente.id"))
    expediente_id = db.Column(db.Integer, db.ForeignKey("expediente.id"))
    usuario_id = db.Column(db.Integer, db.ForeignKey("usuario.id"))
    cliente_nombre = db.Column(db.String(200))
    tipo_asunto = db.Column(db.String(180))
    monto_total = db.Column(db.Float, default=0)
    monto_letra = db.Column(db.Text)
    forma_pago = db.Column(db.Text)
    fecha_contrato = db.Column(db.String(40))
    personal_atendio = db.Column(db.String(200))
    telefono_cliente = db.Column(db.String(80))
    correo_cliente = db.Column(db.String(180))
    creado_en = db.Column(db.DateTime, default=datetime.now)
    cliente = db.relationship("Cliente")
    expediente = db.relationship("Expediente")
    usuario = db.relationship("Usuario")


class CitaProspecto(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    cliente_id = db.Column(db.Integer, db.ForeignKey("cliente.id"))
    nombre = db.Column(db.String(200), nullable=False)
    telefono = db.Column(db.String(80))
    correo = db.Column(db.String(180))
    materia = db.Column(db.String(80))
    tipo_asunto = db.Column(db.String(180))
    fecha = db.Column(db.String(20))
    hora = db.Column(db.String(10))
    documentacion_solicitada = db.Column(db.Text)
    documentacion_original = db.Column(db.Text)
    estatus = db.Column(db.String(80), default="Agendada")
    observaciones = db.Column(db.Text)
    creado_en = db.Column(db.DateTime, default=datetime.now)
    cliente = db.relationship("Cliente")


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
    add_column("expediente", "cobro_total", "FLOAT DEFAULT 0")
    add_column("expediente", "moneda_cobro", "VARCHAR(10) DEFAULT 'MXN'")
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
.hero{background:linear-gradient(120deg,#111827,#0f766e);color:white;padding:28px 36px;border-bottom:6px solid var(--d);display:flex;align-items:center}
.logo{width:100px;height:100px;border-radius:50%;background:white;object-fit:cover;border:4px solid var(--d);margin-right:20px}
nav{background:#0f172a;padding:12px 30px;line-height:2.2}
nav a{color:white;text-decoration:none;margin:4px 5px 4px 0;font-weight:bold;background:#1f2937;border:1px solid #334155;border-radius:999px;padding:9px 13px;display:inline-block} nav a:hover{background:var(--d);color:white}
main{padding:25px 30px}
.card{background:white;border-radius:14px;padding:22px;margin-bottom:20px;box-shadow:0 3px 12px #0002;border-top:4px solid var(--d)}
input,select,textarea{width:100%;padding:10px;margin:6px 0 12px;border:1px solid #cbd5e1;border-radius:8px;box-sizing:border-box}
button,.btn{background:var(--d);color:white;border:0;border-radius:8px;padding:10px 16px;text-decoration:none;font-weight:bold;display:inline-block;cursor:pointer}
.btn2{background:var(--t)}.btnDark{background:var(--n)}
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
    <div style="display:flex;align-items:center;gap:10px">{% if usuario %}{% if usuario.foto_url %}<img src="{{usuario.foto_url}}" style="width:34px;height:34px;border-radius:50%;object-fit:cover;border:2px solid #B38B2E">{% endif %}<span>{{usuario.nombre}} | {{usuario.rol}}</span><a class="btn btn2" href="/perfil">Mi Perfil</a><a class="btn btnDark" href="/logout">Cerrar sesión</a>{% endif %}</div>
</div>
<section class="hero">
    <img class="logo" src="/logo">
    <div><h1>ALESI GRUPO JURÍDICO</h1><p>Sistema Integral de Gestión Jurídica y Control de Expedientes</p></div>
</section>
{% if usuario %}
<nav>
    <a href="/">Inicio</a>
    <a href="/mis-expedientes">Mis Expedientes</a>
    <a href="/compartidos-conmigo">Expedientes Compartidos</a>
    <a href="/planeador">Planeador</a>
    <a href="/clientes">Clientes</a>
    <a href="/citas">Citas / Prospectos</a>
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
    for p in Cobranza.query.filter(Cobranza.proximo_pago.in_(fechas)).order_by(Cobranza.proximo_pago).all():
        pendientes.append({"fecha": p.proximo_pago, "label": labels.get(p.proximo_pago, ""), "tipo": "Pago", "texto": f"{p.cliente.nombre if p.cliente else ''} - {p.moneda} {p.saldo}"})

    inicio_semana = hoy - timedelta(days=hoy.weekday())
    fin_semana = inicio_semana + timedelta(days=6)
    cobranza_semana = Cobranza.query.filter(Cobranza.proximo_pago >= inicio_semana.strftime("%Y-%m-%d"), Cobranza.proximo_pago <= fin_semana.strftime("%Y-%m-%d"), Cobranza.saldo > 0).count()
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
    citas = CitaProspecto.query.filter_by(cliente_id=id).order_by(CitaProspecto.fecha.desc(), CitaProspecto.hora.desc()).all()
    return render("""<div class="card no-print"><h2>Cliente</h2><a class="btn" href="/editar-cliente/{{c.id}}">Editar datos del cliente</a> <a class="btn btn2" href="/contrato-cliente/{{c.id}}">Crear contrato de prestación de servicios</a> <a class="btn btnDark" href="/expedientes?cliente_id={{c.id}}">Agregar nuevo asunto / expediente</a> <button onclick="window.print()">Imprimir reporte</button></div><div class="card print-list"><h2>Datos del cliente</h2><p><b>Nombre:</b> {{c.nombre}}</p><p><b>Teléfono:</b> {{c.telefono}}</p><p><b>Correo:</b> {{c.correo}}</p><p><b>RFC:</b> {{c.rfc}}</p><p><b>CURP:</b> {{c.curp}}</p><p><b>Materia principal:</b> {{c.materia}}</p><p><b>Tipo de asunto principal:</b> {{c.tipo_asunto}}</p><p><b>Cobro principal pactado:</b> {{c.moneda_cobro}} {{c.cobro_total}}</p><p><b>Dirección:</b> {{c.direccion}}</p><p>{{c.observaciones}}</p></div><div class="card print-list"><h2>Expedientes iniciados o concluidos</h2><table><tr><th>Número</th><th>Materia</th><th>Tipo</th><th>Estado</th><th>Cobro pactado</th><th>Abonado</th><th>Saldo</th></tr>{% for e in exps %}{% set pagos_exp = pagos|selectattr('expediente_id', 'equalto', e.id)|list %}{% set abonado = pagos_exp|sum(attribute='abono') %}{% set ultimo = pagos_exp[0] if pagos_exp else None %}<tr><td><a href="/expediente/{{e.id}}">{{e.numero}}</a></td><td>{{e.materia}}</td><td>{{e.tipo_asunto}}</td><td>{{e.estado}}</td><td>{{e.moneda_cobro or c.moneda_cobro}} {{e.cobro_total or c.cobro_total}}</td><td>{{abonado}}</td><td>{{ultimo.saldo if ultimo else (e.cobro_total or c.cobro_total or 0)}}</td></tr>{% endfor %}</table></div><div class="card print-list"><h2>Pagos realizados desglosados por expediente</h2><table><tr><th>Fecha</th><th>Folio</th><th>Expediente</th><th>Concepto</th><th>Forma de pago</th><th>Monto total</th><th>Abono</th><th>Saldo</th><th>Recibo</th></tr>{% for p in pagos %}<tr><td>{{p.fecha_pago}}</td><td>{{p.folio}}</td><td>{{p.expediente.numero if p.expediente else ''}}</td><td>{{p.concepto}}</td><td>{{p.forma_pago}}</td><td>{{p.moneda}} {{p.monto_total}}</td><td>{{p.moneda}} {{p.abono}}</td><td>{{p.moneda}} {{p.saldo}}</td><td class="no-print"><a href="/recibo/{{p.id}}">Ver</a> | <a href="/recibo-pdf/{{p.id}}">PDF</a></td></tr>{% endfor %}</table></div><div class="card print-list"><h2>Citas / prospectos relacionados</h2><table><tr><th>Fecha</th><th>Hora</th><th>Tipo asunto</th><th>Documentación solicitada</th><th>Documentación original recibida</th><th>Estatus</th><th class="no-print">Acción</th></tr>{% for ci in citas %}<tr><td>{{ci.fecha}}</td><td>{{ci.hora}}</td><td>{{ci.tipo_asunto}}</td><td>{{ci.documentacion_solicitada}}</td><td>{{ci.documentacion_original}}</td><td>{{ci.estatus}}</td><td class="no-print"><a class="btn btn2" href="/cita/{{ci.id}}">Abrir</a></td></tr>{% endfor %}</table></div>""", c=c, pagos=pagos, exps=exps, citas=citas)


@app.route("/editar-cliente/<int:id>", methods=["GET", "POST"])
def editar_cliente(id):
    if not req():
        return redirect("/login")
    c = Cliente.query.get_or_404(id)
    if request.method == "POST":
        c.nombre = request.form["nombre"]
        c.telefono = request.form.get("telefono", "")
        c.correo = request.form.get("correo", "")
        c.rfc = request.form.get("rfc", "")
        c.curp = request.form.get("curp", "")
        c.materia = request.form.get("materia", "")
        c.tipo_asunto = request.form.get("tipo_asunto", "")
        c.cobro_total = float(request.form.get("cobro_total") or 0)
        c.moneda_cobro = request.form.get("moneda_cobro", "MXN")
        c.direccion = request.form.get("direccion", "")
        c.observaciones = request.form.get("observaciones", "")
        db.session.commit()
        flash("Cliente actualizado.")
        return redirect("/cliente/" + str(id))
    return render("""<div class="card"><h2>Editar cliente</h2><form method="post"><label>Nombre</label><input name="nombre" value="{{c.nombre or ''}}" required><label>Teléfono</label><input name="telefono" value="{{c.telefono or ''}}"><label>Correo</label><input name="correo" value="{{c.correo or ''}}"><label>RFC</label><input name="rfc" value="{{c.rfc or ''}}"><label>CURP</label><input name="curp" value="{{c.curp or ''}}"><label>Materia</label><select name="materia">{% for m in materias %}<option {% if c.materia == m %}selected{% endif %}>{{m}}</option>{% endfor %}</select><label>Tipo de asunto</label><select name="tipo_asunto">{% for t in tipos %}<option {% if c.tipo_asunto == t.nombre %}selected{% endif %}>{{t.nombre}}</option>{% endfor %}</select><label>Cobro total pactado</label><input type="number" step="0.01" name="cobro_total" value="{{c.cobro_total or 0}}"><label>Moneda</label><select name="moneda_cobro"><option {% if c.moneda_cobro == 'MXN' %}selected{% endif %}>MXN</option><option {% if c.moneda_cobro == 'USD' %}selected{% endif %}>USD</option></select><label>Dirección</label><textarea name="direccion">{{c.direccion or ''}}</textarea><label>Observaciones</label><textarea name="observaciones">{{c.observaciones or ''}}</textarea><button>Guardar cambios</button></form></div>""", c=c, materias=MATERIAS, tipos=TipoAsunto.query.order_by(TipoAsunto.materia, TipoAsunto.nombre).all())


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
            cobro_total=float(request.form.get("cobro_total") or 0),
            moneda_cobro=request.form.get("moneda_cobro", "MXN"),
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
        return redirect("/expedientes")
    datos = Expediente.query.order_by(Expediente.id.desc()).all() if admin() else Expediente.query.filter_by(propietario_id=actual().id).order_by(Expediente.id.desc()).all()
    return render("""<div class="card no-print"><h2>Nuevo expediente</h2><form method="post"><label>Cliente</label><select name="cliente_id"><option value="">Sin cliente</option>{% for cliente in clientes %}<option value="{{cliente.id}}">{{cliente.nombre}}</option>{% endfor %}</select><label>Número</label><input name="numero" required><label>Materia</label><select name="materia">{% for m in materias %}<option>{{m}}</option>{% endfor %}</select><label>Tipo de asunto</label><select name="tipo_asunto">{% for t in tipos %}<option>{{t.nombre}}</option>{% endfor %}</select><label>Cobro pactado para este asunto</label><input type="number" step="0.01" name="cobro_total"><label>Moneda</label><select name="moneda_cobro"><option>MXN</option><option>USD</option></select><label>Autoridad</label><input name="autoridad"><label>Actor</label><input name="actor"><label>Demandado</label><input name="demandado"><label>Estado</label><select name="estado"><option>Nuevo</option><option>En trámite</option><option>Pendiente de acuerdo</option><option>Pendiente de audiencia</option><option>Concluido</option><option>Archivado</option><option>Urgente</option></select><label>Prioridad</label><select name="prioridad"><option>Baja</option><option selected>Media</option><option>Alta</option></select><label>Responsable</label><input name="responsable"><label>Fecha inicio</label><input type="date" name="fecha_inicio"><label>Observaciones</label><textarea name="observaciones"></textarea><button>Guardar</button></form></div><div class="card print-list"><h2>Lista de expedientes cargados <button class="no-print" onclick="window.print()">Imprimir lista</button></h2><table><tr><th>Número</th><th>Cliente</th><th>Materia</th><th>Tipo asunto</th><th>Estado</th><th class="no-print">Acción</th></tr>{% for expediente in datos %}<tr><td>{{expediente.numero}}</td><td>{{expediente.cliente.nombre if expediente.cliente else ""}}</td><td>{{expediente.materia}}</td><td>{{expediente.tipo_asunto}}</td><td>{{expediente.estado}}</td><td class="no-print"><a class="btn btn2" href="/expediente/{{expediente.id}}">Abrir</a></td></tr>{% endfor %}</table></div>""", clientes=Cliente.query.order_by(Cliente.nombre).all(), materias=MATERIAS, tipos=TipoAsunto.query.order_by(TipoAsunto.materia, TipoAsunto.nombre).all(), datos=datos)


@app.route("/mis-expedientes")
def mis_expedientes():
    if not req():
        return redirect("/login")
    u = actual()
    if admin() or u.rol == "Usuario Oficina" or u.area == "Oficina":
        datos = Expediente.query.order_by(Expediente.id.desc()).all()
    else:
        ids = {e.id for e in Expediente.query.filter_by(propietario_id=u.id).all()}
        ids.update([c.expediente_id for c in Compartido.query.filter_by(usuario_id=u.id).all()])
        datos = Expediente.query.filter(Expediente.id.in_(ids)).order_by(Expediente.id.desc()).all() if ids else []
    return render("""<div class="card no-print"><h2>Nuevo expediente / asunto</h2><p>Desde aquí se captura un nuevo expediente. La pestaña única para consultar expedientes es Mis Expedientes.</p><a class="btn" href="/expedientes">Nuevo expediente</a></div><div class="card print-list"><h2>Mis Expedientes / Expedientes del despacho <button class="no-print" onclick="window.print()">Imprimir lista</button></h2><table><tr><th>Número</th><th>Cliente</th><th>Materia</th><th>Tipo</th><th>Estado</th><th>Cobro pactado</th><th>Propietario</th><th class="no-print"></th></tr>{% for e in datos %}<tr><td>{{e.numero}}</td><td>{{e.cliente.nombre if e.cliente else ""}}</td><td>{{e.materia}}</td><td>{{e.tipo_asunto}}</td><td>{{e.estado}}</td><td>{{e.moneda_cobro}} {{e.cobro_total}}</td><td>{{e.propietario.nombre if e.propietario else ''}}</td><td class="no-print"><a class="btn btn2" href="/expediente/{{e.id}}">Abrir</a></td></tr>{% endfor %}</table></div>""", datos=datos)


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
    return render("""<div class="card"><h2>Expediente {{expediente.numero}} <button onclick="window.print()">Imprimir</button></h2><p><b>Cliente:</b> {{expediente.cliente.nombre if expediente.cliente else ""}}</p><p><b>Materia:</b> {{expediente.materia}}</p><p><b>Tipo de asunto:</b> {{expediente.tipo_asunto}}</p><p><b>Autoridad:</b> {{expediente.autoridad}}</p><p><b>Actor:</b> {{expediente.actor}}</p><p><b>Demandado:</b> {{expediente.demandado}}</p><p><b>Estado:</b> {{expediente.estado}}</p>{% if puede_editar %}<form method="post" action="/actualizar-estatus/{{expediente.id}}"><label>Cambiar estatus</label><select name="estado"><option {% if expediente.estado == 'Nuevo' %}selected{% endif %}>Nuevo</option><option {% if expediente.estado == 'En trámite' %}selected{% endif %}>En trámite</option><option {% if expediente.estado == 'Pendiente de acuerdo' %}selected{% endif %}>Pendiente de acuerdo</option><option {% if expediente.estado == 'Pendiente de audiencia' %}selected{% endif %}>Pendiente de audiencia</option><option {% if expediente.estado == 'Concluido' %}selected{% endif %}>Concluido</option><option {% if expediente.estado == 'Archivado' %}selected{% endif %}>Archivado</option><option {% if expediente.estado == 'Urgente' %}selected{% endif %}>Urgente</option></select><button>Actualizar estatus</button></form>{% endif %}<p><b>Tu permiso:</b> <span class="badge">{{permiso}}</span></p><p>{{expediente.observaciones}}</p>{% if puede_editar %}<a class="btn" href="/movimiento/{{expediente.id}}">Agregar promoción/vencimiento</a> <a class="btn btn2" href="/audiencia/{{expediente.id}}">Programar audiencia</a> <a class="btn" href="/mensaje/{{expediente.id}}">Mensaje interno</a>{% endif %} {% if puede_administrar %}<a class="btn btnDark" href="/compartir/{{expediente.id}}">Compartir / Accesos</a>{% endif %} {% if puede_editar %}<a class="btn" href="/aviso-expediente/{{expediente.id}}">Agregar aviso</a>{% endif %}</div><div class="grid2"><div class="card"><h2>Promociones / vencimientos</h2><table><tr><th>Título</th><th>Usuario</th><th>Límite</th><th>Archivo</th></tr>{% for movimiento in movimientos %}<tr><td>{{movimiento.titulo}}</td><td>{{movimiento.usuario.nombre}}</td><td>{{movimiento.fecha_limite}} {{movimiento.hora_limite}}</td><td>{% if movimiento.archivo_url %}<a href="{{movimiento.archivo_url}}" target="_blank">Ver</a>{% endif %}</td></tr>{% endfor %}</table></div><div class="card"><h2>Audiencias</h2><table><tr><th>Fecha</th><th>Hora</th><th>Audiencia</th><th>Usuario</th></tr>{% for audiencia in audiencias %}<tr><td>{{audiencia.fecha}}</td><td>{{audiencia.hora}}</td><td>{{audiencia.titulo}}</td><td>{{audiencia.usuario.nombre}}</td></tr>{% endfor %}</table></div></div><div class="grid2"><div class="card"><h2>Chat interno</h2>{% for m in mensajes %}<div class="chat"><b>{{m.usuario.nombre}}</b><div class="small">{{m.creado_en.strftime("%d/%m/%Y %H:%M")}}</div><p>{{m.mensaje}}</p></div>{% endfor %}</div><div class="card"><h2>Avisos del expediente</h2>{% for av in avisos_exp %}<div class="chat"><b>{{av.titulo}}</b> - {{av.creado_en.strftime('%d/%m/%Y %H:%M')}}<br>{{av.mensaje}}</div>{% endfor %}</div></div><div class="card"><h2>Bitácora</h2>{% for b in bitacora %}<div class="bit"><b>{{b.accion}}</b><div class="small">{{b.usuario.nombre if b.usuario else ""}} | {{b.creado_en.strftime("%d/%m/%Y %H:%M")}}</div><p>{{b.detalle}}</p></div>{% endfor %}</div>""", expediente=expediente, permiso=permiso_expediente(id), puede_editar=puede_editar(id), puede_administrar=puede_administrar(id), movimientos=Movimiento.query.filter_by(expediente_id=id).order_by(Movimiento.id.desc()).all(), audiencias=Audiencia.query.filter_by(expediente_id=id).order_by(Audiencia.fecha, Audiencia.hora).all(), mensajes=MensajeExpediente.query.filter_by(expediente_id=id).order_by(MensajeExpediente.id.desc()).all(), bitacora=Bitacora.query.filter_by(expediente_id=id).order_by(Bitacora.id.desc()).all(), avisos_exp=AvisoExpediente.query.filter_by(expediente_id=id).order_by(AvisoExpediente.id.desc()).all())

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
    return redirect("/expedientes")
    
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
    return render("""<div class="card"><h2>Programar audiencia</h2><form method="post"><label>Audiencia</label><input name="titulo" required><label>Fecha</label><input type="date" name="fecha" required><label>Hora</label><input type="time" name="hora" required><label>Cobro pactado para este asunto</label><input type="number" step="0.01" name="cobro_total"><label>Moneda</label><select name="moneda_cobro"><option>MXN</option><option>USD</option></select><label>Autoridad</label><input name="autoridad"><label>Sala</label><input name="sala"><label>Modalidad</label><select name="modalidad"><option>Presencial</option><option>Virtual</option><option>Mixta</option></select><label>Enlace</label><input name="enlace"><label><input type="checkbox" name="notificar" checked style="width:auto"> Notificar 24h y 2h antes</label><label>Observaciones</label><textarea name="observaciones"></textarea><button>Guardar</button></form></div>""")


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
            proximo_pago=request.form["proximo_pago"], estatus="Liquidado" if saldo <= 0 else "Pendiente",
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
    for p in Cobranza.query.filter(Cobranza.proximo_pago != "").all():
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
    return render("""<div class="grid2"><div class="card"><h2>Agregar aviso interno</h2><form method="post"><label>Título</label><input name="titulo" required><label>Mensaje</label><textarea name="mensaje" required></textarea><label>Enviar a usuario</label><select name="usuario_id"><option value="">Ninguno</option>{% for u in usuarios %}<option value="{{u.id}}">{{u.nombre}} - {{u.usuario}}</option>{% endfor %}</select><label>O enviar a área</label><select name="area"><option value="">Ninguna</option><option>Jurídico</option><option>Administración</option><option>Oficina</option><option>Penal</option><option>Laboral</option><option>Administrativo</option><option>Civil</option><option>Mercantil</option><option>Familiar</option></select><label>Relacionar con expediente</label><select name="expediente_id"><option value="">Sin expediente</option>{% for e in expedientes %}<option value="{{e.id}}">{{e.numero}}</option>{% endfor %}</select><button>Enviar aviso</button></form></div><div class="card"><h2>Avisos internos</h2>{% for a in avisos %}<div class="chat"><b>{{a.titulo}}</b> <span class="small">{{a.creado_en.strftime('%d/%m/%Y %H:%M')}}</span><br>{{a.mensaje}}<br>{% if a.enlace %}<a href="{{a.enlace}}">Abrir</a>{% endif %}</div>{% endfor %}</div></div>""", avisos=avisos, usuarios=Usuario.query.filter_by(activo=True).order_by(Usuario.nombre).all(), expedientes=Expediente.query.order_by(Expediente.numero).all())


@app.route("/avisos-marcar-leidos")
def avisos_marcar_leidos():
    if not req():
        return redirect("/login")
    for aviso in NotificacionInterna.query.filter_by(usuario_id=actual().id).all():
        aviso.leida = True
    db.session.commit()
    flash("Avisos marcados como leídos.")
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


@app.route("/citas", methods=["GET", "POST"])
def citas():
    if not req():
        return redirect("/login")
    if request.method == "POST":
        cliente_id = request.form.get("cliente_id") or None
        cliente = db.session.get(Cliente, int(cliente_id)) if cliente_id else None
        nombre = cliente.nombre if cliente else request.form.get("nombre", "")
        telefono = cliente.telefono if cliente else request.form.get("telefono", "")
        correo = cliente.correo if cliente else request.form.get("correo", "")
        cita = CitaProspecto(
            cliente_id=cliente_id,
            nombre=nombre,
            telefono=telefono,
            correo=correo,
            materia=request.form.get("materia", ""),
            tipo_asunto=request.form.get("tipo_asunto", ""),
            fecha=request.form.get("fecha", ""),
            hora=request.form.get("hora", ""),
            documentacion_solicitada=request.form.get("documentacion_solicitada", ""),
            documentacion_original=request.form.get("documentacion_original", ""),
            estatus=request.form.get("estatus", "Agendada"),
            observaciones=request.form.get("observaciones", ""),
        )
        db.session.add(cita)
        db.session.commit()
        flash("Cita registrada.")
        return redirect("/citas")
    datos = CitaProspecto.query.order_by(CitaProspecto.fecha.desc(), CitaProspecto.hora.desc()).all()
    return render("""<div class="grid2"><div class="card"><h2>Registrar cita / posible cliente</h2><form method="post"><label>Cliente existente opcional</label><select name="cliente_id"><option value="">Posible cliente nuevo</option>{% for c in clientes %}<option value="{{c.id}}">{{c.nombre}}</option>{% endfor %}</select><label>Nombre del posible cliente</label><input name="nombre"><label>Teléfono</label><input name="telefono"><label>Correo</label><input name="correo"><label>Materia</label><select name="materia">{% for m in materias %}<option>{{m}}</option>{% endfor %}</select><label>Tipo de asunto</label><select name="tipo_asunto">{% for t in tipos %}<option>{{t.nombre}}</option>{% endfor %}</select><label>Fecha de cita</label><input type="date" name="fecha"><label>Hora</label><input type="time" name="hora"><label>Documentación solicitada</label><textarea name="documentacion_solicitada" placeholder="Ej. acta de matrimonio, actas de nacimiento, INE, comprobantes, etc."></textarea><label>Documentación original que dejó el cliente</label><textarea name="documentacion_original"></textarea><label>Estatus</label><select name="estatus"><option>Agendada</option><option>Atendida</option><option>Reprogramada</option><option>Convertida en cliente</option><option>Cancelada</option></select><label>Observaciones</label><textarea name="observaciones"></textarea><button>Guardar cita</button></form></div><div class="card"><h2>Citas registradas</h2><table><tr><th>Fecha</th><th>Hora</th><th>Nombre</th><th>Tipo asunto</th><th>Documentación solicitada</th><th>Originales recibidos</th><th>Estatus</th><th>Acción</th></tr>{% for c in datos %}<tr><td>{{c.fecha}}</td><td>{{c.hora}}</td><td>{{c.nombre}}</td><td>{{c.tipo_asunto}}</td><td>{{c.documentacion_solicitada}}</td><td>{{c.documentacion_original}}</td><td>{{c.estatus}}</td><td><a class="btn btn2" href="/cita/{{c.id}}">Abrir</a></td></tr>{% endfor %}</table></div></div>""", datos=datos, clientes=Cliente.query.order_by(Cliente.nombre).all(), materias=MATERIAS, tipos=TipoAsunto.query.order_by(TipoAsunto.materia, TipoAsunto.nombre).all())


@app.route("/cita/<int:id>", methods=["GET", "POST"])
def cita(id):
    if not req():
        return redirect("/login")
    ci = CitaProspecto.query.get_or_404(id)
    if request.method == "POST":
        ci.cliente_id = request.form.get("cliente_id") or None
        cliente = db.session.get(Cliente, int(ci.cliente_id)) if ci.cliente_id else None
        ci.nombre = cliente.nombre if cliente else request.form.get("nombre", "")
        ci.telefono = cliente.telefono if cliente else request.form.get("telefono", "")
        ci.correo = cliente.correo if cliente else request.form.get("correo", "")
        ci.materia = request.form.get("materia", "")
        ci.tipo_asunto = request.form.get("tipo_asunto", "")
        ci.fecha = request.form.get("fecha", "")
        ci.hora = request.form.get("hora", "")
        ci.documentacion_solicitada = request.form.get("documentacion_solicitada", "")
        ci.documentacion_original = request.form.get("documentacion_original", "")
        ci.estatus = request.form.get("estatus", "Agendada")
        ci.observaciones = request.form.get("observaciones", "")
        db.session.commit()
        flash("Cita actualizada.")
        return redirect("/cita/" + str(id))
    return render("""<div class="card"><h2>Actualizar cita</h2><form method="post"><label>Cliente existente opcional</label><select name="cliente_id"><option value="">Posible cliente nuevo</option>{% for cl in clientes %}<option value="{{cl.id}}" {% if ci.cliente_id == cl.id %}selected{% endif %}>{{cl.nombre}}</option>{% endfor %}</select><label>Nombre</label><input name="nombre" value="{{ci.nombre or ''}}"><label>Teléfono</label><input name="telefono" value="{{ci.telefono or ''}}"><label>Correo</label><input name="correo" value="{{ci.correo or ''}}"><label>Materia</label><select name="materia">{% for m in materias %}<option {% if ci.materia == m %}selected{% endif %}>{{m}}</option>{% endfor %}</select><label>Tipo de asunto</label><select name="tipo_asunto">{% for t in tipos %}<option {% if ci.tipo_asunto == t.nombre %}selected{% endif %}>{{t.nombre}}</option>{% endfor %}</select><label>Fecha</label><input type="date" name="fecha" value="{{ci.fecha or ''}}"><label>Hora</label><input type="time" name="hora" value="{{ci.hora or ''}}"><label>Documentación solicitada</label><textarea name="documentacion_solicitada">{{ci.documentacion_solicitada or ''}}</textarea><label>Documentación original que dejó el cliente</label><textarea name="documentacion_original">{{ci.documentacion_original or ''}}</textarea><label>Estatus</label><select name="estatus"><option {% if ci.estatus == 'Agendada' %}selected{% endif %}>Agendada</option><option {% if ci.estatus == 'Atendida' %}selected{% endif %}>Atendida</option><option {% if ci.estatus == 'Reprogramada' %}selected{% endif %}>Reprogramada</option><option {% if ci.estatus == 'Convertida en cliente' %}selected{% endif %}>Convertida en cliente</option><option {% if ci.estatus == 'Cancelada' %}selected{% endif %}>Cancelada</option></select><label>Observaciones</label><textarea name="observaciones">{{ci.observaciones or ''}}</textarea><button>Guardar cambios</button></form></div>""", ci=ci, clientes=Cliente.query.order_by(Cliente.nombre).all(), materias=MATERIAS, tipos=TipoAsunto.query.order_by(TipoAsunto.materia, TipoAsunto.nombre).all())


UNIDADES = ["", "UNO", "DOS", "TRES", "CUATRO", "CINCO", "SEIS", "SIETE", "OCHO", "NUEVE"]
ESPECIALES = {
    10: "DIEZ", 11: "ONCE", 12: "DOCE", 13: "TRECE", 14: "CATORCE", 15: "QUINCE",
    16: "DIECISÉIS", 17: "DIECISIETE", 18: "DIECIOCHO", 19: "DIECINUEVE",
    20: "VEINTE", 21: "VEINTIUNO", 22: "VEINTIDÓS", 23: "VEINTITRÉS", 24: "VEINTICUATRO",
    25: "VEINTICINCO", 26: "VEINTISÉIS", 27: "VEINTISIETE", 28: "VEINTIOCHO", 29: "VEINTINUEVE"
}
DECENAS = {30: "TREINTA", 40: "CUARENTA", 50: "CINCUENTA", 60: "SESENTA", 70: "SETENTA", 80: "OCHENTA", 90: "NOVENTA"}
CENTENAS = {100: "CIEN", 200: "DOSCIENTOS", 300: "TRESCIENTOS", 400: "CUATROCIENTOS", 500: "QUINIENTOS", 600: "SEISCIENTOS", 700: "SETECIENTOS", 800: "OCHOCIENTOS", 900: "NOVECIENTOS"}
MESES = ["", "ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"]


def numero_entero_letra(n):
    n = int(n)
    if n == 0:
        return "CERO"
    if n < 10:
        return UNIDADES[n]
    if n < 30:
        return ESPECIALES[n]
    if n < 100:
        d = (n // 10) * 10
        r = n % 10
        return DECENAS[d] if r == 0 else DECENAS[d] + " Y " + UNIDADES[r]
    if n < 1000:
        if n in CENTENAS:
            return CENTENAS[n]
        c = (n // 100) * 100
        r = n % 100
        return ("CIENTO" if c == 100 else CENTENAS[c]) + " " + numero_entero_letra(r)
    if n < 1000000:
        miles = n // 1000
        r = n % 1000
        texto = "MIL" if miles == 1 else numero_entero_letra(miles) + " MIL"
        return texto if r == 0 else texto + " " + numero_entero_letra(r)
    millones = n // 1000000
    r = n % 1000000
    texto = "UN MILLÓN" if millones == 1 else numero_entero_letra(millones) + " MILLONES"
    return texto if r == 0 else texto + " " + numero_entero_letra(r)


def cantidad_en_letra(valor):
    try:
        numero = float(str(valor).replace(",", ""))
    except Exception:
        numero = 0
    entero = int(numero)
    centavos = int(round((numero - entero) * 100))
    return f"{numero_entero_letra(entero)} PESOS {centavos:02d}/100 MONEDA NACIONAL"


def fecha_contrato_texto(fecha_iso):
    try:
        f = datetime.strptime(fecha_iso, "%Y-%m-%d")
        return str(f.day), MESES[f.month], str(f.year), f"{f.day} DE {MESES[f.month]} DEL {f.year}"
    except Exception:
        hoy = datetime.now()
        return str(hoy.day), MESES[hoy.month], str(hoy.year), f"{hoy.day} DE {MESES[hoy.month]} DEL {hoy.year}"


def ptxt(texto, style):
    return Paragraph(str(texto or "").replace("&", "&amp;").replace("\n", "<br/>"), style)


def encabezado_pie(canvas, doc):
    canvas.saveState()
    w, h = letter
    logo_path = "logo_alesi.png"
    if os.path.exists(logo_path):
        try:
            canvas.drawImage(logo_path, 40, h - 90, width=95, height=70, preserveAspectRatio=True, mask='auto')
        except Exception:
            pass
    canvas.setStrokeColor(colors.HexColor("#B38B2E"))
    canvas.setLineWidth(1)
    canvas.line(38, 52, w - 38, 52)
    canvas.setFont("Helvetica", 7)
    canvas.setFillColor(colors.HexColor("#6b7280"))
    canvas.drawCentredString(w/2, 38, "ALESI GRUPO JURÍDICO | Cuauhtémoc 5048, Col. Ampliación Gabriel Rodríguez, 22207, Tijuana, Baja California | 664 565 7298 | lic.rosavazquezm@gmail.com")
    canvas.drawRightString(w - 38, 24, f"Página {doc.page}")
    canvas.restoreState()


def contrato_pdf_response(cliente, expediente, datos):
    fecha_iso = datos.get("fecha_contrato") or datetime.now().strftime("%Y-%m-%d")
    dia, mes, anio, fecha_larga = fecha_contrato_texto(fecha_iso)
    cliente_nombre = datos.get("cliente_nombre") or cliente.nombre
    cliente_tel = datos.get("telefono") or cliente.telefono or ""
    cliente_correo = datos.get("correo") or cliente.correo or ""
    asunto = datos.get("tipo_asunto") or (expediente.tipo_asunto if expediente else cliente.tipo_asunto) or ""
    monto_total = float(datos.get("monto_total") or 0)
    monto_inicio = float(datos.get("monto_inicio") or 0)
    monto_abono = float(datos.get("monto_abono") or 0)
    forma_pago = datos.get("forma_pago") or ""
    periodicidad = datos.get("periodicidad_abono") or ""
    prestador = datos.get("prestador") or "Lic. Rosa Isela Vázquez Medina"
    personal = datos.get("personal_atendio") or prestador
    domicilio = datos.get("domicilio") or cliente.direccion or ""
    curp = datos.get("curp") or cliente.curp or ""
    rfc = datos.get("rfc") or cliente.rfc or ""
    fecha_nacimiento = datos.get("fecha_nacimiento") or ""
    estado_civil = datos.get("estado_civil") or ""
    ocupacion = datos.get("ocupacion") or ""
    identificacion = datos.get("identificacion") or ""
    numero_identificacion = datos.get("numero_identificacion") or ""

    registro = ContratoCliente(
        cliente_id=cliente.id,
        expediente_id=expediente.id if expediente else None,
        usuario_id=actual().id if actual() else None,
        cliente_nombre=cliente_nombre,
        tipo_asunto=asunto,
        monto_total=monto_total,
        monto_letra=cantidad_en_letra(monto_total),
        forma_pago=forma_pago,
        fecha_contrato=fecha_larga,
        personal_atendio=personal,
        telefono_cliente=cliente_tel,
        correo_cliente=cliente_correo,
    )
    db.session.add(registro)
    db.session.commit()

    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=55, leftMargin=55, topMargin=105, bottomMargin=70)
    styles = getSampleStyleSheet()
    normal = ParagraphStyle("ContratoNormal", parent=styles["Normal"], fontName="Helvetica", fontSize=10.5, leading=15, alignment=4, spaceAfter=8)
    bold = ParagraphStyle("ContratoBold", parent=normal, fontName="Helvetica-Bold")
    title = ParagraphStyle("ContratoTitle", parent=styles["Title"], fontName="Helvetica-Bold", fontSize=15, leading=18, alignment=1, spaceAfter=14)
    subtitle = ParagraphStyle("ContratoSub", parent=styles["Heading2"], fontName="Helvetica-Bold", fontSize=12.5, leading=15, alignment=1, spaceBefore=12, spaceAfter=8)
    heading = ParagraphStyle("ContratoHeading", parent=styles["Heading2"], fontName="Helvetica-Bold", fontSize=12, leading=15, spaceBefore=10, spaceAfter=7)
    small = ParagraphStyle("ContratoSmall", parent=styles["Normal"], fontName="Helvetica", fontSize=9, leading=12, spaceAfter=5)

    story = []
    story.append(ptxt("CONTRATO DE PRESTACIÓN DE SERVICIOS JURÍDICOS PROFESIONALES", title))
    story.append(ptxt(f"Tijuana, Baja California, a {dia} de {mes} del {anio}.", normal))
    story.append(ptxt(f"CONTRATO DE PRESTACIÓN DE SERVICIOS JURÍDICOS PROFESIONALES QUE CELEBRAN POR UNA PARTE EL (LA) LICENCIADA EN DERECHO <b>{prestador}</b>, A QUIEN EN LO SUCESIVO SE LE DENOMINARÁ <b>\"EL PRESTADOR\"</b>; Y POR LA OTRA PARTE EL(LA) C. <b>{cliente_nombre}</b>, A QUIEN EN LO SUCESIVO SE LE DENOMINARÁ <b>\"EL CLIENTE\"</b>, QUIENES MANIFIESTAN SU VOLUNTAD DE OBLIGARSE AL TENOR DE LAS SIGUIENTES:", normal))
    story.append(ptxt("DECLARACIONES", subtitle))
    story.append(ptxt("I. DECLARA EL PRESTADOR:", heading))
    story.append(ptxt("a) Que es profesionista legalmente autorizada para ejercer la profesión de Licenciada en Derecho, contando con la preparación académica, experiencia y capacidad técnica necesarias para prestar servicios de asesoría, consultoría, representación y litigio en materia jurídica.", normal))
    story.append(ptxt("b) Que tiene la capacidad legal suficiente para celebrar el presente contrato y obligarse en los términos del mismo.", normal))
    story.append(ptxt("c) Que señala como domicilio para oír y recibir notificaciones el ubicado en calle Cuauhtémoc 5048, colonia Ampliación Gabriel Rodríguez, 22207, Tijuana, Baja California.", normal))
    story.append(ptxt("d) Que su correo electrónico y teléfono de contacto son: correo electrónico: lic.rosavazquezm@gmail.com; teléfono: 664 565 7298.", normal))
    story.append(ptxt("II. DECLARA EL CLIENTE:", heading))
    story.append(ptxt("a) Que es una persona física con capacidad legal suficiente para contratar y obligarse en los términos del presente instrumento.", normal))
    story.append(ptxt("b) Que requiere los servicios profesionales de EL PRESTADOR para la atención de un asunto jurídico de su interés.", normal))
    story.append(ptxt("c) Que la información y documentación que proporcione para la atención del asunto es auténtica, veraz y vigente.", normal))
    story.append(ptxt(f"d) Que señala como domicilio para oír y recibir notificaciones el ubicado en: <b>{domicilio}</b>.", normal))
    story.append(ptxt("e) Que proporciona los siguientes datos de identificación:", normal))
    datos_identificacion = [
        ["Fecha de nacimiento", fecha_nacimiento], ["CURP", curp], ["RFC", rfc], ["Estado civil", estado_civil],
        ["Ocupación", ocupacion], ["Teléfono", cliente_tel], ["Correo electrónico", cliente_correo],
        ["Identificación oficial", identificacion], ["Número de identificación", numero_identificacion]
    ]
    tabla = Table(datos_identificacion, colWidths=[150, 310])
    tabla.setStyle(TableStyle([('GRID',(0,0),(-1,-1),0.25,colors.lightgrey),('BACKGROUND',(0,0),(0,-1),colors.HexColor('#f3f4f6')),('FONTNAME',(0,0),(0,-1),'Helvetica-Bold'),('VALIGN',(0,0),(-1,-1),'TOP'),('FONTSIZE',(0,0),(-1,-1),8.8),('LEADING',(0,0),(-1,-1),11)]))
    story.append(tabla)
    story.append(ptxt("III. DECLARAN AMBAS PARTES:", heading))
    story.append(ptxt("ÚNICA. Que se reconocen mutuamente la personalidad con la que comparecen, así como la capacidad legal necesaria para obligarse conforme a los términos y condiciones establecidos en el presente contrato.", normal))
    story.append(ptxt("CLÁUSULAS", subtitle))
    story.append(ptxt("PRIMERA. OBJETO.", heading))
    story.append(ptxt(f"EL PRESTADOR se obliga a proporcionar a EL CLIENTE servicios jurídicos profesionales consistentes en asesoría, orientación legal, elaboración de documentos jurídicos, representación, gestión administrativa, negociación, mediación, conciliación, seguimiento procesal y, en su caso, litigio ante autoridades administrativas, judiciales o de cualquier otra naturaleza, respecto del siguiente asunto: <b>{asunto}</b>.", normal))
    story.append(ptxt("La prestación de servicios comprenderá únicamente las actuaciones relacionadas con el asunto anteriormente descrito.", normal))
    story.append(ptxt("SEGUNDA. ALCANCE DE LOS SERVICIOS.", heading))
    story.append(ptxt("Los servicios objeto del presente contrato comprenden todas aquellas actividades profesionales que resulten razonablemente necesarias para la adecuada atención del asunto encomendado. Cualquier procedimiento, recurso, incidente, juicio diverso, apelación, amparo, ejecución de sentencia o trámite distinto al originalmente contratado requerirá autorización expresa del cliente y podrá generar honorarios adicionales.", normal))
    story.append(ptxt("TERCERA. HONORARIOS PROFESIONALES.", heading))
    story.append(ptxt(f"Por los servicios profesionales objeto del presente contrato, EL CLIENTE se obliga a pagar a EL PRESTADOR la cantidad de: <b>${monto_total:,.2f} PESOS M.N. ({cantidad_en_letra(monto_total)})</b> por concepto de HONORARIOS.", normal))
    story.append(ptxt(f"Los honorarios serán cubiertos de la siguiente forma: <b>${monto_inicio:,.2f} PESOS M.N. ({cantidad_en_letra(monto_inicio)})</b> por concepto DE INICIO, pagaderos a la firma del presente contrato.", normal))
    story.append(ptxt(f"<b>${monto_abono:,.2f} PESOS M.N. ({cantidad_en_letra(monto_abono)})</b> por concepto DE ABONO, pagaderos <b>{periodicidad}</b> y hasta la liquidación. Forma de pago pactada: <b>{forma_pago}</b>.", normal))
    story.append(ptxt("La falta de pago de cualquiera de las parcialidades en las fechas convenidas constituirá mora automática, sin necesidad de requerimiento judicial o extrajudicial.", normal))
    story.append(ptxt("CUARTA. GASTOS Y COSTAS.", heading))
    story.append(ptxt("Los honorarios pactados no incluyen gastos extraordinarios necesarios para la atención del asunto, incluyendo de manera enunciativa más no limitativa: notificaciones, emplazamientos, edictos, derechos gubernamentales, certificaciones, copias certificadas, honorarios de fedatarios públicos, traducciones, apostillas, peritajes, gastos de ejecución, publicaciones, viáticos, traslados y diligencias fuera de la ciudad. Dichos gastos serán cubiertos por EL CLIENTE previa solicitud o comprobación correspondiente.", normal))
    story.append(ptxt("QUINTA. OBLIGACIONES DEL CLIENTE.", heading))
    story.append(ptxt("EL CLIENTE se obliga a proporcionar información completa, verdadera y oportuna; entregar la documentación requerida; comparecer personalmente cuando sea necesaria su presencia; cubrir oportunamente los honorarios y gastos; mantener actualizados sus datos de localización e informar cualquier circunstancia que pueda afectar el desarrollo del asunto.", normal))
    story.append(ptxt("SEXTA. OBLIGACIONES DEL PRESTADOR.", heading))
    story.append(ptxt("EL PRESTADOR se obliga a actuar con diligencia, ética profesional y apego a derecho; mantener informado al cliente sobre el estado procesal del asunto; guardar absoluta confidencialidad respecto de la información proporcionada y realizar las actuaciones profesionales necesarias para la adecuada atención del asunto.", normal))
    story.append(ptxt("SÉPTIMA. NATURALEZA DE LA OBLIGACIÓN.", heading))
    story.append(ptxt("Las partes reconocen que los servicios objeto del presente contrato corresponden a una prestación de servicios profesionales de carácter jurídico, por lo que las obligaciones asumidas por EL PRESTADOR son de medios y no de resultados. En consecuencia, EL PRESTADOR se compromete a emplear sus conocimientos, experiencia, capacidad profesional y los recursos jurídicos legalmente procedentes para la adecuada atención del asunto encomendado; sin embargo, no garantiza el sentido de las resoluciones, sentencias, acuerdos o determinaciones que emitan las autoridades administrativas, jurisdiccionales o de cualquier otra naturaleza que intervengan en el asunto.", normal))
    story.append(ptxt("OCTAVA. CONFIDENCIALIDAD.", heading))
    story.append(ptxt("Toda información, documentación, datos personales y antecedentes proporcionados por EL CLIENTE tendrán carácter estrictamente confidencial y serán utilizados exclusivamente para la atención del asunto objeto del presente contrato.", normal))
    story.append(ptxt("NOVENA. TERMINACIÓN ANTICIPADA.", heading))
    story.append(ptxt("El presente contrato podrá darse por terminado anticipadamente por cualquiera de las partes mediante aviso por escrito. En caso de terminación por decisión de EL CLIENTE, los honorarios ya pagados no serán reembolsables y deberán cubrirse aquellos servicios efectivamente prestados a la fecha de terminación. La terminación anticipada no extingue las obligaciones de pago pendientes a cargo de EL CLIENTE.", normal))
    story.append(ptxt("DÉCIMA. RESCISIÓN.", heading))
    story.append(ptxt("EL PRESTADOR podrá rescindir el presente contrato sin responsabilidad alguna cuando EL CLIENTE proporcione información falsa, oculte información relevante para el asunto, incumpla con el pago de honorarios o gastos, o se generen conflictos que imposibiliten la adecuada representación profesional.", normal))
    story.append(ptxt("DÉCIMA PRIMERA. INCUMPLIMIENTO DE PAGO.", heading))
    story.append(ptxt("Si EL CLIENTE incumple con el pago de los honorarios pactados por un periodo mayor a treinta días naturales, EL PRESTADOR podrá suspender la prestación de los servicios hasta que el adeudo sea cubierto en su totalidad, sin perjuicio de las acciones legales procedentes para el cobro de los adeudos existentes.", normal))
    story.append(ptxt("DÉCIMA SEGUNDA. DOCUMENTACIÓN.", heading))
    story.append(ptxt("EL CLIENTE autoriza a EL PRESTADOR para utilizar la documentación proporcionada exclusivamente para la atención del asunto descrito en este contrato. Los documentos originales podrán ser devueltos previa solicitud por escrito del cliente. EL CLIENTE manifiesta bajo protesta de decir verdad que la documentación entregada es auténtica.", normal))
    story.append(ptxt("DÉCIMA TERCERA. PROTECCIÓN DE DATOS PERSONALES.", heading))
    story.append(ptxt("Las partes reconocen que los datos personales recabados con motivo de la celebración del presente contrato serán tratados conforme a la legislación aplicable en materia de protección de datos personales.", normal))
    story.append(ptxt("DÉCIMA CUARTA. JURISDICCIÓN.", heading))
    story.append(ptxt("Para la interpretación, cumplimiento y ejecución del presente contrato, las partes se someten expresamente a las leyes y tribunales competentes de la ciudad de Tijuana, Baja California, renunciando a cualquier otro fuero que pudiera corresponderles por razón de sus domicilios presentes o futuros.", normal))
    story.append(Spacer(1, 16))
    story.append(ptxt(f"Leído que fue el presente contrato y enteradas las partes de su contenido, alcance y consecuencias legales, lo firman por duplicado en la ciudad de Tijuana, Baja California, a {dia} de {mes} del {anio}.", normal))
    story.append(Spacer(1, 38))
    firmas = Table([
        [ptxt("<b>EL PRESTADOR</b>", small), ptxt("<b>EL CLIENTE</b>", small)],
        [ptxt("______________________________", small), ptxt("______________________________", small)],
        [ptxt(f"LIC. {personal}", small), ptxt(f"Nombre: {cliente_nombre}", small)],
        [ptxt("", small), ptxt(f"Teléfono: {cliente_tel}", small)],
        [ptxt("", small), ptxt(f"Identificación Oficial No. {numero_identificacion}", small)],
    ], colWidths=[240, 240])
    firmas.setStyle(TableStyle([('ALIGN',(0,0),(-1,-1),'CENTER'),('VALIGN',(0,0),(-1,-1),'TOP'),('TOPPADDING',(0,0),(-1,-1),5)]))
    story.append(firmas)
    doc.build(story, onFirstPage=encabezado_pie, onLaterPages=encabezado_pie)
    buffer.seek(0)
    nombre = "contrato_" + cliente_nombre.replace(" ", "_").replace("/", "_") + ".pdf"
    return send_file(buffer, as_attachment=True, download_name=nombre, mimetype="application/pdf")


@app.route("/monto-letra")
def monto_letra_api():
    return Response(cantidad_en_letra(request.args.get("monto", "0")), mimetype="text/plain; charset=utf-8")


@app.route("/contrato-cliente/<int:id>", methods=["GET", "POST"])
def contrato_cliente(id):
    if not req():
        return redirect("/login")
    c = Cliente.query.get_or_404(id)
    exps = Expediente.query.filter_by(cliente_id=id).order_by(Expediente.id.desc()).all()
    if request.method == "POST":
        expediente = db.session.get(Expediente, int(request.form.get("expediente_id"))) if request.form.get("expediente_id") else None
        return contrato_pdf_response(c, expediente, request.form)
    fecha_hoy = datetime.now().strftime("%Y-%m-%d")
    return render("""
    <div class="card">
        <h2>Crear contrato de prestación de servicios jurídicos profesionales</h2>
        <p>Llena los datos en blanco y al guardar se descargará el contrato en PDF listo para imprimir.</p>
        <form method="post">
            <label>Cliente</label>
            <input name="cliente_nombre" value="{{c.nombre or ''}}" required>
            <label>Teléfono del cliente</label>
            <input name="telefono" value="{{c.telefono or ''}}">
            <label>Correo del cliente</label>
            <input name="correo" value="{{c.correo or ''}}">
            <label>Domicilio del cliente</label>
            <textarea name="domicilio">{{c.direccion or ''}}</textarea>
            <label>Fecha de nacimiento</label>
            <input type="date" name="fecha_nacimiento">
            <label>CURP</label>
            <input name="curp" value="{{c.curp or ''}}">
            <label>RFC</label>
            <input name="rfc" value="{{c.rfc or ''}}">
            <label>Estado civil</label>
            <input name="estado_civil">
            <label>Ocupación</label>
            <input name="ocupacion">
            <label>Identificación oficial</label>
            <input name="identificacion" placeholder="INE, pasaporte, licencia, etc.">
            <label>Número de identificación</label>
            <input name="numero_identificacion">
            <label>Expediente / asunto relacionado</label>
            <select name="expediente_id">
                <option value="">Usar datos generales del cliente</option>
                {% for e in exps %}<option value="{{e.id}}">{{e.numero}} - {{e.tipo_asunto}} - {{e.moneda_cobro}} {{e.cobro_total}}</option>{% endfor %}
            </select>
            <label>Tipo de asunto / servicio contratado</label>
            <input name="tipo_asunto" value="{{c.tipo_asunto or ''}}" required>
            <label>Fecha del contrato</label>
            <input type="date" name="fecha_contrato" value="{{fecha_hoy}}" required>
            <label>Total de honorarios</label>
            <input type="number" step="0.01" name="monto_total" id="monto_total" value="{{c.cobro_total or 0}}" oninput="actualizarLetra()" required>
            <label>Cantidad en letra automática</label>
            <input id="monto_letra_preview" value="" readonly>
            <label>Pago inicial</label>
            <input type="number" step="0.01" name="monto_inicio" value="0">
            <label>Abono o parcialidad</label>
            <input type="number" step="0.01" name="monto_abono" value="0">
            <label>Periodicidad del abono</label>
            <input name="periodicidad_abono" placeholder="Ej. semanal, quincenal, mensual, en cada recuperación, etc.">
            <label>Forma de cobro / observación de pago</label>
            <textarea name="forma_pago" placeholder="Ej. pago inicial a la firma y abonos quincenales hasta liquidar; 50% de lo recuperado; etc."></textarea>
            <label>Licenciada / prestador</label>
            <input name="prestador" value="Lic. Rosa Isela Vázquez Medina">
            <label>Personal que atendió y firma</label>
            <input name="personal_atendio" value="{{usuario.nombre if usuario else 'Lic. Rosa Isela Vázquez Medina'}}">
            <button>Guardar y descargar contrato PDF</button>
        </form>
    </div>
    <script>
    function actualizarLetra(){
        const monto = document.getElementById('monto_total').value || '0';
        fetch('/monto-letra?monto=' + encodeURIComponent(monto))
          .then(r => r.text())
          .then(t => { document.getElementById('monto_letra_preview').value = t; })
          .catch(() => {});
    }
    window.onload = actualizarLetra;
    </script>
    """, c=c, exps=exps, fecha_hoy=fecha_hoy)


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
    if not req() or not admin():
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
    if not req() or not admin():
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

    # CITAS / PROSPECTOS
    ws = wb.create_sheet("Citas")
    ws.append(["ID", "Cliente", "Nombre", "Teléfono", "Correo", "Materia", "Tipo asunto", "Fecha", "Hora", "Documentación solicitada", "Documentación original", "Estatus", "Observaciones", "Creado en"])
    for ci in CitaProspecto.query.order_by(CitaProspecto.id).all():
        ws.append([ci.id, ci.cliente.nombre if ci.cliente else "", ci.nombre, ci.telefono, ci.correo, ci.materia, ci.tipo_asunto, ci.fecha, ci.hora, ci.documentacion_solicitada, ci.documentacion_original, ci.estatus, ci.observaciones, ci.creado_en.strftime("%d/%m/%Y %H:%M") if ci.creado_en else ""])

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
