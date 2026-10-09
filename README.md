# UDEC | Seguimiento de PAD

Aplicación Flask que sirve el tablero original y comparte el Excel entre todos
los visitantes usando Supabase Storage. Si existe el archivo compartido, sus
datos reemplazan los datos iniciales del tablero al abrir o recargar la página.

## Preparar el Excel

El tablero espera un Excel compatible con su estructura y la hoja `Control`.
En producción se guarda en el bucket privado de Supabase con la ruta fija
`seguimiento_pad.xlsx`. El botón **Actualizar desde Excel** reemplaza ese
archivo; los visitantes verán los datos nuevos al abrir o recargar el tablero.

## Desplegar gratis en Render con Supabase

### 1. Crear el proyecto y el bucket de Supabase

1. Crea una cuenta en [Supabase](https://supabase.com/) y un proyecto en el
   plan Free. Guarda la contraseña de la base de datos en un lugar seguro.
2. En el proyecto, abre **Storage** y crea un bucket llamado exactamente
   `seguimiento-pad`.
3. Deja el bucket **privado**. La aplicación accede desde el servidor con una
   clave secreta; los navegadores de los visitantes no reciben esa clave.
4. En **Project Settings > API Keys**, copia la **Project URL** y la clave
   heredada `service_role` (JWT). No uses la clave `anon` o `publishable` para
   este propósito.

### 2. Crear el servicio gratuito en Render

1. Sube estos cambios a un repositorio GitHub.
2. En [Render](https://render.com/), selecciona **New > Blueprint** y conecta
   ese repositorio. Render leerá el archivo `render.yaml` y creará el servicio
   web con plan **Free**.
3. Cuando Render pida las variables, introduce:
   - `SUPABASE_URL`: la Project URL copiada de Supabase.
   - `SUPABASE_SERVICE_ROLE_KEY`: la clave secreta de servidor. Guárdala solo
     como variable secreta en Render; nunca la publiques en GitHub ni en el
     código del navegador.
4. Confirma el despliegue y espera a que termine.

### 3. Publicar el Excel compartido

Abre la URL pública que Render asignó al servicio y usa
**Actualizar desde Excel** para seleccionar el archivo. La aplicación lo
subirá a Supabase y reemplazará `seguimiento_pad.xlsx`. La siguiente vez que
cualquier visitante abra o recargue el tablero, recibirá esos datos. No hace
falta subir una copia en cada navegador ni compartir la clave de Supabase.

Para comprobar la conexión después del despliegue, abre
`https://TU-SERVICIO.onrender.com/api/status`. Debe responder JSON indicando
`"storage": "supabase"`, `"configured": true` y, después de subir un archivo,
`"workbook_exists": true`. Si indica `local` o `configured: false`, confirma
las variables del servicio en **Render > Environment** y vuelve a desplegar.
Si devuelve HTTP 502, revisa que la URL, la clave secreta y el bucket sean
correctos.

**Aviso:** la carga desde el tablero no pide clave ni inicio de sesión, como se
solicitó. Por tanto, cualquier persona que conozca la URL puede reemplazar el
Excel compartido. La clave de Supabase permanece en el servidor, pero el
endpoint de carga es público. Si más adelante quieres limitar quién publica,
habrá que proteger la carga con autenticación.

El plan Free de Render puede dormir tras un periodo sin visitas. Supabase Free
tiene límites de almacenamiento/tráfico y puede pausar proyectos inactivos;
consulta los límites vigentes en [Render](https://render.com/docs/free) y
[Supabase](https://supabase.com/pricing).

## Ejecutar localmente

Requiere Python 3.10 o posterior.

```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
python app.py
```

Abre <http://127.0.0.1:5000>. Sin `SUPABASE_URL` y
`SUPABASE_SERVICE_ROLE_KEY`, se usa el archivo local
`data/seguimiento_pad.xlsx`, si existe.

El contenido original del tablero se encuentra en `templates/index.html`.
