# Backlog - Nomina Poltech

Este archivo lleva el registro de tareas pendientes para no perderlas entre sesiones.

## Fase A - Persistencia en Render (COMPLETADA)
Disco persistente + DB_PATH para que la base de datos ya no se borre en cada deploy.

## Quick wins de diseno (COMPLETADA)
Menu agrupado (Personal / Nomina / Configuracion), quitar color dorado sin marca,
columna fija en tabla de nomina, acentos en Aviso de baja y Comprobante de vigencia.

## Correcciones (COMPLETADA - Fase A cerrada)
- Menu lateral desplegable (acordeon, cerrado al entrar): hecho.

## Pendientes del documento "Pendientes_Nomina_Poltech_4.docx" (30-jul-2026)
- Configurar SMTP en Render: hecho, correo confirmado funcionando en produccion.
- Baja / reingreso por NSS conservando cedula: hecho (modulo de solicitud/autorizacion
  de reingreso, preserva la cedula original).
- Reportes (backlog, ya identificados antes): calidad de datos e historial de nomina ya
  existen (modulo Reportes). Siguen pendientes: retardos por obra, horas extra, fechas de
  ingreso (aguinaldos), bitacora exportable.
- Obras y Catalogo: se queda restringido de superintendente hacia arriba (decidido, sin cambio).

## Fase B - Migrar SQLite a PostgreSQL (pendiente)
Solucion definitiva de persistencia (mas robusta que el disco de Render).

## Punto 2 - Seguridad (pendiente)
Mover el token de API Market a variable de entorno, limitar intentos de login.

## Pruebas de calculo de nomina (pendiente)
Pruebas automaticas para evitar que un cambio futuro rompa el calculo de sueldos.

## Dividir app.py en modulos/blueprints (pendiente)
Separar personal / nomina / api / reportes en archivos distintos (hoy todo esta en un solo app.py).

## Contrato individual de trabajo (COMPLETADA)
Se reemplazo por completo con el formato nuevo (28 clausulas + 3 anexos), con acentos
correctos, siguiendo el documento que aporto el usuario. Pendiente futuro: reemplazar el
Anexo 1 generico por el detallado de cada puesto (ver "Manual de operaciones" abajo).

## Tarea 2 - Obra alterna / cobro cruzado (pendiente, va en Reportes)
Personal del taller que a veces trabaja horas en otra obra. Esas horas se deben cobrar
financieramente a la obra alterna (con un % de indirectos), pero el pago real / salida de
caja sigue siendo por la nomina del taller. Es un cobro interno, no un pago duplicado.
El usuario confirmo que esto es parte de la seccion de Reportes, se atiende mas adelante.

## Fase C - Expediente digital de obra / cliente (pendiente, NUEVA)
Conforme el sistema crece, agregar un lugar para subir y llevar el expediente de cada
obra/contrato con la informacion que piden los clientes:
- Fianzas
- Contratos firmados
- Pagos de IMSS
- Reportes de IMSS / INFONAVIT
- Altas del personal
- Avances por obra
- Cobros por obra

Falta definir en una sesion futura: si son solo archivos adjuntos (PDF/Excel) ligados a
cada obra, o si tambien llevan datos estructurados (montos, fechas de vigencia, folios)
para poder generar reportes y alertas de vencimiento.

## Comprobantes de pago de nomina (pendiente, NUEVA - va en Fase C)
- Subir los PDF de los comprobantes/pagos de nomina (evidencia de que se pago).
- Generar consultas de CEP (Comprobante Electronico de Pago) de Banxico para corroborar
  que el pago realmente se ejecuto (no solo que se calculo/libero en el sistema).

## Rol de descanso de los sabados (pendiente, NUEVA)
Hacer el rol/calendario de quien descansa cada sabado (rotacion del personal).
Falta definir en una sesion futura como se organiza (por obra, por trabajador, etc.)
y si se necesita una pantalla en el sistema o solo un documento/plantilla.

## Manual de operaciones: funciones detalladas por puesto (pendiente, NUEVA)
El Anexo 1 del contrato (funciones del puesto) hoy usa una version generica y corta por
clasificacion (Soldadores, Montadores, etc.), redactada rapido para poder generar contratos
de inmediato. Falta escribir las funciones detalladas y especificas de cada puesto para
un manual de operaciones real, que tambien viva en el sistema (por ejemplo, reemplazando
el texto generico del Anexo 1 con el detallado de cada puesto).

## Modulo de vacaciones (COMPLETADA)
Calculo de dias por antiguedad (tabla LFT vigente, sin prima vacacional), dashboard de
tomados/disponibles por periodo (Vacaciones en el menu Personal), flujo de solicitud
(cualquier rol con acceso a la obra) y autorizacion (superintendente/administrador), y
hoja de referencia "Vacaciones" en la plantilla de asistencia para saber a quien marcarle
"V" esa semana (el pago de esos dias sin prima ya existia en el calculo de nomina).
Ademas: si se marcan mas dias "V" de los que le quedan de derecho, el excedente no se paga
(aviso al calcular la nomina); y hay un ajuste manual (solo administrador) para capturar
cuantos dias ya tomo un trabajador antiguo antes de usar este sistema.

## Pago de fletes a choferes (pendiente, NUEVA)
Pago de fletes a los choferes en base a los envios realizados, y generar precios de
fletes de embarques (area de logistica). Falta definir en una sesion futura como se
capturan los envios/embarques, quien fija el precio de cada flete, y como se conecta
con el pago semanal del chofer.

## Reportes de KPIs por gerencia (pendiente, NUEVA)
Hacer reportes de indicadores clave (KPIs) por cada una de las gerencias. Falta definir en
una sesion futura que gerencias son, que KPIs le importan a cada una y de donde salen los
datos (nomina, asistencia, catalogo, etc.).
