# Sistema de Nómina POLTECH — versión 1 (cimiento)

Aplicación web hecha con Flask (Python). Incluye:
- **Login con roles** (admin, dirección, gerencia, gerente de obra, autorizador, capturista)
- **Usuarios** (solo el admin): crear cuentas del personal, asignar rol y obra, resetear contraseñas
- **Obras**
- **Catálogo de sueldos por obra** (puestos con sueldo y viáticos semanales; solo gerente de obra o superior)
- **Personal** (alta manual, con validación de NSS y CURP; eximir documentos con autorización de un tercero)
- **Cuentas bancarias** (sin duplicados; tipo tarjeta / CLABE / número de cuenta)

> Pendiente para las siguientes rondas: cálculo de nómina semanal, reportes, contratos, préstamos y layout de dispersión bancaria.

## Archivos
```
app.py               <- la aplicación
requirements.txt     <- dependencias (Flask, gunicorn)
render.yaml          <- configuración de Render
templates/           <- páginas HTML
static/logo.jpeg     <- logo de la empresa
```

## Cómo poner esto en tu proyecto
1. Copia **todos** estos archivos dentro de tu carpeta `poltech` (reemplaza el `app.py` de prueba).
2. Sube los cambios:
   ```
   git add .
   git commit -m "Sistema de nomina v1"
   git push
   ```
3. Render actualiza la web solo en ~2 minutos.

## Variables en Render (recomendado)
En Render → tu servicio → **Environment**:
- `ADMIN_PASSWORD` = la contraseña que quieras para el usuario **admin**.
- `SECRET_KEY` = se genera sola (ya está en render.yaml).

## Primer acceso
- Usuario: **admin**
- Contraseña: la de `ADMIN_PASSWORD` (o `cambiar123` si no la definiste — cámbiala de inmediato).

## Importante
En el plan gratis de Render la base de datos (SQLite) **no es permanente**: se borra al reiniciar/redeployar.
Usa datos de prueba hasta conectar una base de datos permanente (siguiente etapa).
