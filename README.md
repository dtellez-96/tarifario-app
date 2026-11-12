# Cotizador FCL: análisis y punto de partida web

Se analizó estáticamente el XLSM original, sin ejecutar macros, actualizar Power Query ni modificar el libro.

- [Análisis funcional, arquitectura y plan de migración](docs/PLAN_MIGRACION.md).
- [Inventario y perfil de datos extraídos](docs/evidence/inventory.json).
- [VBA realmente instalado](docs/evidence/Module1.bas).
- [Consultas Power Query extraídas](docs/evidence/PowerQuery.m).
- [DDL inicial PostgreSQL](sql/001_initial.sql).
- [Pruebas de reglas](tests/test_pricing.py) y [pruebas contra el libro](tests/test_workbook.py).

## Alcance entregado

Prototipo **local**, Flask + HTML/Jinja + JavaScript + DaisyUI/Tailwind, con:

- Importación transaccional del libro consolidado; SHA-256, staging e incidencias.
- PostgreSQL como destino empresarial; SQLite para demostración local.
- Motor estricto por defecto y compatibilidad VBA solo para auditoría.
- Comparativo de hasta cinco agentes, desglose, exclusiones y advertencias.
- Ajustes de origen y llenado por oferta, con recálculo de alternativas y auditoría.
- Histórico inmutable de estimaciones provisionales, sin autorización comercial.
- Catálogos de consulta e importaciones por lote; ninguna escritura sobre Excel.
- Interfaz DaisyUI con sidebar responsive, sección activa, header y footer compartidos.
  En móvil, el menú se abre con el botón de navegación y se cierra con Escape, el botón
  de cierre o el fondo. Incluye navegación por teclado y enlace para saltar al contenido.
  El formulario organiza los datos de cotización, la ruta y los equipos en tres secciones.
  En escritorio, el botón de doble flecha reduce la barra lateral a iconos o la expande;
  la preferencia se guarda en el navegador. Los iconos conservan nombres accesibles
  y etiquetas al pasar el cursor. En móvil se mantiene el menú completo.
- Paginación DaisyUI compartida en histórico (25 registros por página) y catálogos
  (50 registros por página), con rango, total, página activa y controles deshabilitados
  en los extremos. Conserva el filtro por lote; en móvil muestra anterior, página/total
  y siguiente. Las tablas del comparativo y sus desgloses se mantienen completas.

**No es una aplicación empresarial terminada.** No tiene autenticación, permisos por usuario,
aprobación comercial, cargas semanales de archivos individuales, exportación PDF ni tareas programadas.
El acceso remoto está bloqueado. No debe publicarse mediante proxy o túnel.
Consultar fases y criterios de salida del [plan](docs/PLAN_MIGRACION.md).

## Ejecutar en Windows / PowerShell

El entorno `.venv` y la base local ya fueron preparados en este espacio de trabajo.
Para una instalación nueva:

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements-dev.txt
npm ci
npm run build:css
.\.venv\Scripts\python.exe -m flask --app wsgi init-db
.\.venv\Scripts\python.exe -m flask --app wsgi import-workbook .\Cotizador_FCL_Automatico_Visual_Final_v6.xlsm
```

La importación repetida del mismo archivo se rechaza explícitamente; no volver a importar el lote
ya cargado. `init-db` crea tablas inexistentes, **no actualiza** un esquema anterior.

Arranque local (el proceso permanece en esa terminal):

```powershell
.\.venv\Scripts\waitress-serve.exe --listen=127.0.0.1:5000 wsgi:app
```

Abrir `http://127.0.0.1:5000`. Los estilos están compilados localmente; no se usan CDN ni
servicios externos para procesar cotizaciones.

SQLite se guarda en `instance\tarifario.db`. La clave de sesión local se genera una sola vez
en `instance\secret-key`, excluida del control de versiones; también se admite `SECRET_KEY`.
No compartir esa carpeta. Respaldar la base detenida o con la API de backup SQLite.
Los archivos de evidencia también contienen datos comerciales: tratarlos como confidenciales.

## Caso reproducible

Seleccionar lote 1, fecha **2026-11-10**, SHANGHAI → CARTAGENA, 20 ST = 0,
40 ST = 0, 40 HC = 2, BL = 1, FOB = 100000:

| Modo | Mejor agente | Total nominal |
|---|---|---:|
| Compatibilidad VBA | DSV / COSCO | 1.300,00 |
| Estricto | CONSOLCARGO / COSCO | 19.317,00 |

No es un incremento de precio producido por la migración: el modo VBA acepta importes
incompletos, documentación ausente y ofertas que el estricto excluye. La hoja Guía ya
advierte sobre el valor 450 de DSV. Incluso el resultado estricto es **provisional**:
los cargos de destino no declaran moneda y los costos manuales requieren revisión.
Después del 16 de noviembre de 2026, las ofertas fechadas de este archivo dejan de
ser válidas para cotizaciones nuevas; no se deben extender sus fechas automáticamente.

## Verificación y regeneración

```powershell
.\.venv\Scripts\python.exe -m pytest -q
.\.venv\Scripts\python.exe tools\inspect_workbook.py
.\.venv\Scripts\python.exe tools\export_schema.py
npm run build:css
```

La prueba diferencial recorre **900 escenarios** sobre todas las rutas presentes:
10 orígenes × 9 destinos × 5 combinaciones de equipos × 2 cantidades de BL.
Compara contra una transcripción independiente del VBA, **no contra Excel ejecutándose**.
La aceptación final necesita una sesión controlada en Excel y firma de pricing.

Resultado de la verificación de esta entrega:

- 52 pruebas aprobadas; los 900 escenarios están incluidos en una prueba diferencial.
- Compilación Python, sintaxis JavaScript y build DaisyUI/Tailwind correctos.
- Servidor Waitress local y formulario real con CSRF: cálculo y guardado comprobados.
- Se conservó una estimación de demostración en el histórico local:
  `3bf9f2e7-ab90-4623-adc1-8edae499e373` (19.317,00 nominales).
- El navegador permitió comprobar DOM y estilos calculados; la captura integrada
  devolvió una imagen en blanco, por lo que la revisión visual completa sigue pendiente.
- Servidor temporal de comprobación detenido; utilizar el comando de arranque anterior.
- XLSM original intacto: hash final idéntico al inicial.

## PostgreSQL

Crear una base vacía y un usuario con permisos mínimos mediante la administración de la empresa.
Aplicar [el SQL inicial](sql/001_initial.sql) una sola vez con `psql` o la herramienta corporativa.
Después, configurar en la misma terminal del proceso Flask:

```powershell
$env:DATABASE_URL = "postgresql+psycopg://USUARIO:CLAVE@localhost:5432/tarifario"
$env:SECRET_KEY = "REEMPLAZAR_POR_UN_SECRETO_ALEATORIO"
```

Los valores anteriores son marcadores, no credenciales reales. En producción, usar gestor de secretos.
El DDL PostgreSQL fue generado desde los modelos; su ejecución real y la batería de integración
PostgreSQL están pendientes de disponer de un servidor. No confundir las pruebas SQLite con esa validación.
Alembic, SSO, RBAC, TLS, backups/restauración y control de concurrencia son puertas obligatorias
antes de habilitar acceso multiusuario.

## Organización

```text
app\
  __init__.py          Factory, configuración, CSRF, CLI y bloqueo remoto
  extensions.py       SQLAlchemy y CSRF
  models.py           Nueve entidades persistentes
  routes.py           Blueprint web y manejo explícito de errores
  services\
    imports.py        Lectura segura sin ejecutar macros, staging y promoción
    pricing.py        Cálculo Decimal, modos, validaciones y ranking
  templates\          Formularios, comparativo, histórico, importaciones y catálogos
  static\             JavaScript y CSS local compilado con DaisyUI/Tailwind
docs\
  PLAN_MIGRACION.md
  evidence\           VBA, Power Query, XML e inventario completo por hoja
sql\001_initial.sql   Esquema PostgreSQL inicial
tests\                Pruebas unitarias, integración y caracterización del libro
tools\                Extracción reproducible y generación del DDL
wsgi.py               Entrada WSGI
```
