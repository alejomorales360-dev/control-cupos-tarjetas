# 💳 Control de Cupos Temporales de Tarjetas

**Entrar a la app:** <https://alejomorales360-dev.github.io/control-cupos-tarjetas/> (redirige a la web app de Apps Script, que pide iniciar sesión con la cuenta dueña)

App web en Google Apps Script para llevar el registro de clientes, su **cupo base** y los **aumentos temporales de cupo** (monto + rango de fechas). Todos los días revisa los vencimientos y te **envía una alerta** para que hagas el trámite de corte con el banco.

> Ejemplo: a Pedrito (cupo base $5.000) se le aumentan **$2.000** del **29-09-2026 al 15-10-2026**.
> - 12-10 → aviso previo "faltan 3 días"
> - 15-10 → alerta "**vence hoy**, hacer trámite con el banco"
> - 16-10 en adelante → alerta diaria "**vencido sin cortar**" hasta que lo marques como *Cortado* en la app

## Qué incluye

| Parte | Descripción |
|---|---|
| **Base de datos** | Una planilla de Google Sheets que se crea sola, con las hojas `CLIENTES`, `AUMENTOS` e `HISTORIAL` (registro de cada acción). |
| **App web** (`index.html`) | Panel con lo pendiente, alta/edición de clientes y aumentos, botón **"Marcar cortado"**, cupo actual de cada cliente. Funciona en el celular. |
| **Revisión diaria** | Disparador de Apps Script que corre todos los días a la hora que elijas. |
| **Alertas** | 📧 Correo (Gmail, sin configurar nada) · ✈️ Telegram (opcional, notificación push al celular) · 📅 Evento en Google Calendar el día del corte (opcional). |

## Estados de un aumento

| Estado | Significado |
|---|---|
| Programado | Todavía no empieza |
| Vigente | En curso |
| Por vencer | Termina dentro de los próximos *N* días (configurable, por defecto 3) |
| Vence hoy | Hoy es la fecha de fin → hacer el corte |
| Vencido · cortar | Ya pasó la fecha de fin y no está marcado como cortado (alerta todos los días) |
| Cortado / Anulado | Cerrado; ya no genera alertas |

## Instalación (una sola vez, ~5 minutos)

### Opción A — Sin instalar nada (copiar y pegar)
1. Entra a <https://script.google.com> → **Nuevo proyecto**. Ponle nombre, p. ej. *Control de Cupos*.
2. Reemplaza el contenido de `Código.gs` por el de [`Codigo.gs`](Codigo.gs).
3. **Archivo → Nuevo → HTML**, nómbralo `index` y pega el contenido de [`index.html`](index.html).
4. ⚙️ **Configuración del proyecto** → marca *"Mostrar el archivo de manifiesto appsscript.json"* y pega el contenido de [`appsscript.json`](appsscript.json). Ajusta `timeZone` a tu país si no es Ecuador (p. ej. `America/Bogota`, `America/Lima`, `America/Mexico_City`, `America/Santiago`).
5. En el editor, selecciona la función **`setup`** y pulsa **Ejecutar**. Acepta los permisos. Esto crea la planilla y el disparador diario (la URL de la planilla aparece en el registro).
6. **Implementar → Nueva implementación → Aplicación web**. *Ejecutar como:* Yo. *Quién tiene acceso:* Solo yo. Copia la URL y guárdala en favoritos / pantalla de inicio del celular.

### Opción B — Con `clasp`
```bash
npm install -g @google/clasp
clasp login
clasp create --type standalone --title "Control de Cupos"   # o copia .clasp.json.example a .clasp.json con un scriptId existente
clasp push
clasp open    # ejecuta setup() desde el editor y luego implementa como Web App (paso 6)
```

## Configurar las alertas

En la app → pestaña **Configuración**:

- **Correo(s)**: por defecto tu propia cuenta de Google. Puedes poner varios separados por coma.
- **Días de aviso previo** y **hora** de la revisión diaria.
- **Google Calendar**: al activarlo, cada aumento nuevo crea un evento de día completo "⚠️ Cortar cupo: …" en la fecha de fin, con recordatorio el día anterior. Al marcarlo como cortado el evento cambia a "✅ Cupo cortado".
- **Telegram** (recomendado para tener la notificación en el celular):
  1. En Telegram busca **@BotFather**, envía `/newbot` y sigue los pasos. Copia el *token*.
  2. Pégalo en la app y guarda.
  3. Abre tu bot nuevo y envíale cualquier mensaje (ej. "hola").
  4. Pulsa **Detectar chat** y luego **Enviar alerta de prueba**.

Usa **Revisar vencimientos ahora** para forzar la revisión sin esperar al día siguiente.

## Seguridad

- La web app se publica con acceso **"Solo yo"**: nadie más puede abrirla. Si necesitas que un compañero la use, cambia `webapp.access` a `DOMAIN` (misma organización de Google Workspace) y comparte la planilla con esa persona.
- Solo se guardan los **últimos 4 dígitos** de la tarjeta. No registres números completos, CVV ni fechas de expiración.
- El token de Telegram se guarda en las *Script Properties* del proyecto, no en la planilla.

## Archivos

- `Codigo.gs` — backend: base de datos, lógica de estados, revisión diaria y envío de alertas.
- `index.html` — interfaz web.
- `appsscript.json` — manifest (zona horaria, acceso de la web app).
- `.clasp.json.example` — plantilla para `clasp`.

## Actualizar la app publicada

Editar el código no cambia la versión que está en la URL `/exec`. Después de pegar los cambios (o hacer `clasp push`):
**Implementar → Gestionar implementaciones →** lápiz ✏️ → *Versión:* **Nueva versión** → **Implementar**. La URL sigue siendo la misma.

## Enlace de GitHub Pages

`site/index.html` es una página mínima publicada en GitHub Pages que redirige a la URL `/exec` de Apps Script. La app en sí (y los datos) siguen protegidos por el inicio de sesión de Google.

- Se publica sola con el workflow `.github/workflows/pages.yml` cada vez que cambia `site/`.
- Activación (una sola vez): **Settings → Pages → Build and deployment → Source: GitHub Actions**.
- Si algún día cambia la URL `/exec` (nueva *implementación*, no nueva *versión*), actualízala en `site/index.html`.
