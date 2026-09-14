# Sistema de Fotocopias - Albergue Universitario

Aplicacion de escritorio para administrar el cupo mensual de hojas de los becados del Albergue Universitario.

## Objetivo

Cada becado dispone de **100 hojas por mes**. El sistema registra impresiones, donaciones entre becados, correcciones de impresiones fallidas y el historial completo de movimientos por mes.

## Funciones incluidas en el MVP

- Alta, baja logica y busqueda de becados.
- Cupo base configurable (por defecto: 100 hojas por becado y mes).
- Selector de mes y anio.
- Saldo mensual calculado automaticamente:
  - cupo base;
  - hojas impresas;
  - hojas donadas;
  - hojas recibidas;
  - correcciones/reversiones.
- Donacion de hojas entre becados con fecha, donante, receptor y cantidad.
- Historial auditable: los movimientos no se borran silenciosamente.
- Registro de ultima fecha de retiro/impresion.
- Seleccion de archivos para imprimir desde el sistema.
- Conteo automatico de paginas en PDF.
- Conteo de paginas de Word cuando Microsoft Word esta instalado (COM/pywin32); si no, permite indicar manualmente la cantidad.
- Confirmacion de cantidad de hojas antes de descontar saldo.
- Boton **Corregir ultima impresion** para revertir una impresion fallida sin eliminar el movimiento original.
- Editor basico integrado para redactar documentos y exportarlos a `.docx` o `.pdf` antes de imprimir.
- Base de datos local SQLite.

## Criterio de contabilizacion

La unidad descontada es la **hoja fisica**. Antes de imprimir se puede indicar:

- cantidad de paginas del documento;
- cantidad de copias;
- impresion simple faz o doble faz.

Ejemplo: 10 paginas, 2 copias, doble faz = 10 hojas fisicas (ceil(10 / 2) x 2).

La aplicacion nunca descuenta automaticamente solo por abrir un archivo: primero muestra una confirmacion. Si la orden de impresion falla, no registra el consumo. Si la impresora imprime mal despues de enviada la orden, se usa la funcion de correccion.

## Requisitos

- Windows 10/11.
- Python 3.11 o superior para desarrollo.
- Para imprimir `.docx` y obtener su cantidad real de paginas se recomienda Microsoft Word instalado.
- Para el `.exe`, las dependencias Python se empaquetan con PyInstaller.

## Ejecutar en desarrollo

```bat
py -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
python app.py
```

## Generar el `.exe`

Ejecutar:

```bat
build_exe.bat
```

El resultado quedara en:

```text
dist\SistemaFotocopias.exe
```

La base SQLite se guarda en `%LOCALAPPDATA%\SistemaFotocopias\sistema_fotocopias.db`, por lo que actualizar el `.exe` no elimina los datos.

## Importar becados

El MVP permite altas manuales desde la interfaz. La siguiente etapa prevista es importar masivamente desde CSV/Excel con las columnas `Apellido`, `Nombre` y opcionalmente `DNI/Legajo`.

## Arquitectura

- `app.py`: interfaz grafica y flujo principal.
- `database.py`: persistencia SQLite y reglas de saldo/historial.
- `printing.py`: conteo de paginas y envio al sistema de impresion de Windows.
- `document_editor.py`: redaccion y exportacion basica a Word/PDF.
- `tests/test_database.py`: pruebas de la logica de cupos, donaciones y reversiones.

## Proximas mejoras sugeridas

- Importacion masiva de becados desde Excel/CSV.
- Usuarios/operadores con inicio de sesion y nombre en cada movimiento.
- Copias de seguridad automaticas de la base.
- Reportes mensuales exportables a Excel/PDF.
- Configuracion de impresoras y bandejas.
- Vista previa PDF integrada.
- Integracion con lector de codigo/credencial si en el futuro se utiliza identificacion del becado.
