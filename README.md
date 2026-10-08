# UDEC | Seguimiento de PAD

Aplicación Flask que sirve el tablero original sin modificar su estructura,
estilos ni interacciones. Si existe `data/seguimiento_pad.xlsx`, sus datos se
cargan al abrir el tablero y prevalecen sobre una copia guardada anteriormente
en el navegador. Si el archivo todavía no está, se conserva la información
incluida originalmente en el HTML.

## Preparar el Excel

Deja el archivo Excel en esta ruta y con este nombre:

```text
data/seguimiento_pad.xlsx
```

El archivo debe tener la hoja `Control`, como los archivos que admite el
tablero. También puedes actualizar los datos después desde el botón
**Actualizar desde Excel** del propio tablero.

## Ejecutar

Requiere Python 3.9 o posterior.

```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
python app.py
```

Abre <http://127.0.0.1:5000>.

El contenido original del tablero se encuentra en `templates/index.html`.
