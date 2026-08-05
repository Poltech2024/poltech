# Backlog - Nomina Poltech

Este archivo lleva el registro de tareas pendientes para no perderlas entre sesiones.

## Fase A - Persistencia en Render (COMPLETADA)
Disco persistente + DB_PATH para que la base de datos ya no se borre en cada deploy.

## Quick wins de diseno (COMPLETADA)
Menu agrupado (Personal / Nomina / Configuracion), quitar color dorado sin marca,
columna fija en tabla de nomina, acentos en Aviso de baja y Comprobante de vigencia.

## Correcciones (pendiente - para cerrar la Fase A)
- Menu lateral desplegable: hoy el menu agrupado (Personal / Nomina / Configuracion)
  siempre muestra todos los enlaces. Falta que cada grupo se pueda colapsar/expandir
  (acordeon), para que no se vea todo junto todo el tiempo.

## Pendientes del documento "Pendientes_Nomina_Poltech_4.docx" (30-jul-2026)
- Configurar SMTP en Render: faltan las variables de correo para que SALGAN los avisos
  (baja al admin/contador, nomina liberada). El correo del contador ya esta cargado.
  Estado: pendiente de configuracion (no de codigo).
- Baja / reingreso por NSS: permitir reingresar a un trabajador dado de baja conservando
  la misma cedula.
- Reportes (backlog, ya identificados antes): retardos por obra, horas extra, fechas de
  ingreso (aguinaldos), nomina y bitacora exportables.
- Obras y Catalogo: se queda restringido de superintendente hacia arriba (decidido, sin cambio).

## Fase B - Migrar SQLite a PostgreSQL (pendiente)
Solucion definitiva de persistencia (mas robusta que el disco de Render).

## Punto 2 - Seguridad (pendiente)
Mover el token de API Market a variable de entorno, limitar intentos de login.

## Pruebas de calculo de nomina (pendiente)
Pruebas automaticas para evitar que un cambio futuro rompa el calculo de sueldos.

## Dividir app.py en modulos/blueprints (pendiente)
Separar personal / nomina / api / reportes en archivos distintos (hoy todo esta en un solo app.py).

## Acentos del Contrato individual de trabajo (pendiente)
Falta corregir los acentos de todo el texto legal (~150 lineas) en `generar_contrato_docx`.
Se dejo fuera de los quick wins por ser un documento legal largo; requiere su propia sesion con cuidado.

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
