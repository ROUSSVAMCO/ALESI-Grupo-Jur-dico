ALESI GRUPO JURÍDICO - V6 RENDER TESTED

Versión reconstruida y probada con Flask test client.

Incluye:
- Todas las rutas del menú funcionando.
- Mi Perfil: cambiar nombre, correo, contraseña y foto.
- Reloj de fecha y hora en vivo.
- Inicio con conteo de clientes, expedientes, audiencias, vencimientos, avisos internos y cobranza semanal.
- Pendientes de ayer, hoy y mañana.
- Clientes con materia jurídica, tipo de asunto, cobro total pactado y moneda.
- Catálogo editable de tipos de asunto.
- Expedientes, Mis expedientes y Expedientes compartidos.
- Compartir por usuario o área con permisos.
- Chat interno que también crea aviso interno.
- Avisos internos manuales y avisos por expediente.
- Planeador imprimible como calendario.
- Audiencias y vencimientos.
- Cobranza con cálculo de deuda por cliente.
- Recibo HTML imprimible y PDF descargable.
- Jurisprudencia solo con búsqueda ampliada.
- Botones independientes de Boletín Judicial y TEJA.

Usuarios iniciales:
ROUSSVAM / RIVaM061
lic.carmenacostac / LIC.CARMENAC68
oficina.alesi / Oficina2026*

Render:
Build Command: pip install -r requirements.txt
Start Command: gunicorn app:app

Cron Job:
Command: python notificar.py
Schedule: */10 * * * *
