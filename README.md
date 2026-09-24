# 💳 Control de Cupos Temporales de Tarjetas

**Entrar a la app:** <https://alejomorales360-dev.github.io/control-cupos-tarjetas/> (pide usuario y clave)

App web en Google Apps Script para llevar el registro de clientes, su **cupo base** y los **aumentos temporales de cupo** (monto + rango de fechas). Todos los días revisa los vencimientos y te **envía una alerta** para que hagas el trámite de corte con el banco.

> Ejemplo: a Pedrito (cupo base $5.000) se le aumentan **$2.000** del **29-09-2026 al 15-10-2026**.
> - 12-10 → aviso previo "faltan 3 días"
> - 15-10 → alerta "**vence hoy**, hacer trámite con el banco"
> - 16-10 en adelante → alerta diaria "**vencido sin cortar**" hasta que lo marques como *Cortado* en la app

## Arquitectura

```
 Navegador / celular                    Google (tu cuenta)
┌──────────────────────────┐  POST    ┌────────────────────────────┐    ┌──────────────┐
│ site/index.html          │ ───────▶ │ Apps Script  (Codigo.gs)   │ ──▶│ Google Sheets│
│ GitHub Pages (github.io) │ ◀─────── │ API JSON + login + alertas │    │ (base datos) │
└──────────────────────────┘  JSON    └────────────────────────────┘    └──────────────┘
                                         │ disparador diario → 📧 correo · ✈️ Telegram · 📅 Calendar
```

| Parte | Dónde vive | Cómo se actualiza |
|---|---|---|
| **Interfaz** (`site/index.html`) | GitHub Pages | Automático: cada push a `main` que toque `site/` se publica en 1-2 min. |
| **Backend / API** (`Codigo.gs`) | Apps Script | Pegar `Codigo.gs` en el editor → *Gestionar implementaciones* → ✏️ → *Nueva versión*. |
| **Base de datos** | Google Sheets (hojas `CLIENTES`, `AUMENTOS`, `HISTORIAL`) | La crea `setup()`. |
| **Alertas** | Disparador diario de Apps Script | Configuración desde la app. |

La interfaz llama a la API con `fetch` POST (`Content-Type: text/plain`, para evitar el preflight CORS). La URL de la API está en `API_URL` dentro de `site/index.html`; si alguna vez creas una *implementación nueva* (no *nueva versión*), actualiza esa constante.

## Estados de un aumento

| Estado | Significado |
|---|---|
| Programado | Todavía no empieza |
| Vigente | En curso |
| Por vencer | Termina dentro de los próximos *N* días (configurable, por defecto 3) |
| Vence hoy | Hoy es la fecha de fin → hacer el corte |
| Vencido · cortar | Ya pasó la fecha de fin y no está marcado como cortado (alerta todos los días) |
| Cortado / Anulado | Cerrado; ya no genera alertas |

## Instalación (una sola vez)

1. Abre tu planilla de Google Sheets → **Extensiones → Apps Script** (así el script queda vinculado a ella).
   *Alternativa:* un proyecto suelto en <https://script.google.com>, pegando la URL de la planilla en la constante `PLANILLA_URL` de `Codigo.gs`.
2. Reemplaza el contenido de `Código.gs` por [`Codigo.gs`](Codigo.gs). ⚙️ *Configuración del proyecto* → mostrar `appsscript.json` y pega [`appsscript.json`](appsscript.json) (ajusta `timeZone`).
3. Elige la función **`crearBaseDeDatos`** → **▶ Ejecutar** → acepta los permisos. Crea en la planilla:
   - Hojas `CLIENTES`, `AUMENTOS`, `HISTORIAL`, `USUARIOS`, `CONFIG`, `SESIONES` con encabezados, notas explicativas en cada columna, anchos, formatos, listas desplegables (SI/NO, ADMIN/OPERADOR, ACTIVO/CORTADO/ANULADO) y colores (en AUMENTOS: rojo = vencido sin cortar, naranjo = vence hoy).
   - `CONFIG` con todos los ajustes y valores por defecto (correo de alertas = tu cuenta).
   - Usuario **`admin`** con clave aleatoria → aparece en el **Registro de ejecución**.
   - Disparador diario de alertas.

   Se puede ejecutar de nuevo cuando quieras: **nunca borra datos**, solo agrega lo que falte.
4. **Implementar → Nueva implementación → Aplicación web** · *Ejecutar como:* Yo · *Quién tiene acceso:* Cualquier persona. Copia la URL `/exec` en `API_URL_DEFECTO` de `site/index.html` (o úsala desde el enlace **Servidor** del login).
5. Para cambios futuros de `Codigo.gs`: **Gestionar implementaciones → ✏️ → Nueva versión** (la URL no cambia).

## Opción A — Sin instalar nada (copiar y pegar)
1. Entra a <https://script.google.com> → **Nuevo proyecto**. Ponle nombre, p. ej. *Control de Cupos*.
2. Reemplaza el contenido de `Código.gs` por el de [`Codigo.gs`](Codigo.gs).
3. La interfaz no va en Apps Script: está en [`site/index.html`](site/index.html) y la publica GitHub Pages.
4. ⚙️ **Configuración del proyecto** → marca *"Mostrar el archivo de manifiesto appsscript.json"* y pega el contenido de [`appsscript.json`](appsscript.json). Ajusta `timeZone` a tu país si no es Ecuador (p. ej. `America/Bogota`, `America/Lima`, `America/Mexico_City`, `America/Santiago`).
5. En el editor, selecciona la función **`setup`** y pulsa **Ejecutar**. Acepta los permisos. Esto crea la planilla, el disparador diario y la clave de acceso inicial (usuario, clave y URL de la planilla aparecen en el registro).
6. **Implementar → Nueva implementación → Aplicación web**. *Ejecutar como:* Yo. *Quién tiene acceso:* Cualquier persona (la app se protege con su propio login). Copia la URL `/exec` en `API_URL` de `site/index.html`.

### Opción B — Con `clasp`
```bash
npm install -g @google/clasp
clasp login
clasp create --type standalone --title "Control de Cupos"   # o copia .clasp.json.example a .clasp.json con un scriptId existente
clasp push
clasp open    # ejecuta setup() desde el editor y luego implementa como Web App (paso 6)
```

## Importar clientes desde Excel

**Clientes → ⬆ Importar Excel**. La primera fila debe tener los encabezados; se reconocen (sin importar mayúsculas/tildes):

| Dato | Encabezados aceptados |
|---|---|
| Nombre (obligatorio) | `Cardholder name`, `Nombre`, `Cliente`, `Titular` |
| Cupo base | `cantidad USD base`, `Cupo base` |
| Inicio del aumento | `fecha aumento`, `Desde`, `Fecha inicio` |
| Fin del aumento | `fecha termino`, `Hasta`, `Fecha fin` |
| Monto del aumento | `cantidad USD solicitada`, `Monto`, `Monto solicitado` |
| Otros (opcionales) | `Banco`, `Tarjeta` (últimos 4), `Documento`, `Motivo` |

Antes de guardar muestra una vista previa (nuevos, actualizados, errores por fila). Si una fila trae monto + fechas se crea también el aumento. Los clientes se identifican por nombre, así que **reimportar el mismo archivo no duplica** clientes ni aumentos: sirve para ir completando la planilla y volver a subirla. El archivo se lee en el navegador; solo se envían las filas a la API.

## Configurar las alertas

En la app → **Configuración**. Cada canal tiene su botón **Probar** (guarda lo que haya en pantalla y envía un mensaje de prueba solo por ese canal).

| Canal | Qué recibes | Cómo se activa |
|---|---|---|
| 📧 **Correo** | Resumen diario con tabla por estado | Escribe uno o varios correos (Gmail, Outlook, etc.). |
| 💬 **WhatsApp** | Mensaje a tu propio WhatsApp | Gratis vía [CallMeBot](https://www.callmebot.com/blog/free-api-whatsapp-messages/): agrega **+34 644 51 95 23**, envíale `I allow callmebot to send me messages`, pega la API key que te responde + tu número con código de país. |
| 🟪 **Microsoft Teams** | Tarjeta en un canal o chat | Teams → canal/chat → **•••** → **Workflows** → plantilla *“Send webhook alerts to a channel”* (o *“…to a chat”*) → pega la URL que entrega. |
| 📅 **Outlook / Teams calendario** | Invitación `.ics` por correo al crear cada aumento (evento de día completo en la fecha de corte, recordatorios 09:00 del día anterior y del mismo día) | Marca la casilla. Al importar desde Excel se envía un solo correo con todos los eventos. |
| 📅 **Google Calendar** | Evento creado directamente en el calendario de la cuenta del proyecto | Marca la casilla. |
| ✈️ **Telegram** | Mensaje de un bot propio | @BotFather → `/newbot` → token → escribirle al bot → **Detectar chat**. |

Las claves (API key de WhatsApp, URL de Teams, token de Telegram) se guardan en las *Script Properties* y nunca vuelven a la pantalla.

**Cuota de correos:** una cuenta Gmail gratuita puede enviar ~100 correos al día desde Apps Script (la app muestra cuántos quedan). El resumen diario usa 1; cada invitación `.ics` usa 1.

## Base de datos

Por defecto `setup()` crea una planilla nueva. Para usar una propia: **Configuración → Base de datos** → pega la URL de la planilla → **Usar esta planilla**. Se crean en ella las hojas `CLIENTES`, `AUMENTOS` e `HISTORIAL` (las demás hojas no se tocan). La planilla debe ser de la misma cuenta de Google que el proyecto de Apps Script.

## Seguridad, usuarios y dónde se guarda todo

**Todo vive en la planilla de Google Sheets:**

| Hoja | Contenido |
|---|---|
| `CLIENTES`, `AUMENTOS` | Los datos del negocio |
| `HISTORIAL` | Cada acción, con el usuario de la app que la hizo |
| `USUARIOS` | Usuario, nombre, rol (`ADMIN`/`OPERADOR`), activo, `HASH`+`SAL` de la clave, `CLAVE_NUEVA`, último acceso |
| `CONFIG` | Correos, días de aviso, hora, WhatsApp, Teams, Telegram, calendario (editable también a mano: `SI`/`NO`) |
| `SESIONES` | Sesiones abiertas (hash del token + vencimiento). Borrar una fila cierra esa sesión. |

Lo único fuera de la planilla es `SPREADSHEET_ID` en las *Script Properties* (hace falta para encontrarla). Al actualizar desde una versión anterior, el login, las sesiones y la configuración que estaban en *Script Properties* se mueven solos a la planilla la primera vez.

- **Claves:** nunca en texto; se guarda `HASH` (SHA-256 iterado con sal). Para **resetear** la clave de alguien desde la planilla, escribe la clave nueva en su celda `CLAVE_NUEVA`: sirve en su próximo ingreso y se convierte en hash automáticamente.
- **Roles:** `ADMIN` ve todo (configuración, usuarios, base de datos). `OPERADOR` solo clientes y aumentos. Se administran en **Configuración → Usuarios**.
- **Primer usuario / olvidé la clave del admin:** en el editor de Apps Script ejecuta **`generarClaveAcceso`** → crea `admin` (o resetea el primer ADMIN) y muestra la clave en el *Registro de ejecución*.
- **Implementación:** *Ejecutar como:* **Yo** · *Quién tiene acceso:* **Cualquier persona**. La pantalla es pública pero ninguna función entrega datos sin sesión válida.
- Sesiones de 12 h, o 30 días con "Mantener sesión iniciada" (sin marcarla, se cierra al cerrar la pestaña). Bloqueo de 15 min tras 8 intentos fallidos.
- **Quien tenga acceso a la planilla ve la configuración** (incluidas la API key de WhatsApp y la URL de Teams): no la compartas.
- Solo se guardan los **últimos 4 dígitos** de la tarjeta.

## Archivos

- `Codigo.gs` — backend (Apps Script): API JSON, login, base de datos, revisión diaria y alertas.
- `site/index.html` — interfaz web (GitHub Pages).
- `.github/workflows/pages.yml` — publica `site/` en GitHub Pages.
- `appsscript.json` — manifest (zona horaria, acceso de la web app).
- `.clasp.json.example` — plantilla para `clasp`.

## Actualizar la app publicada

Editar el código no cambia la versión que está en la URL `/exec`. Después de pegar los cambios (o hacer `clasp push`):
**Implementar → Gestionar implementaciones →** lápiz ✏️ → *Versión:* **Nueva versión** → **Implementar**. La URL sigue siendo la misma.
