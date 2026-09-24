# 💳 Control de Cupos Temporales de Tarjetas

**Entrar a la app:** <https://alejomorales360-dev.github.io/control-cupos-tarjetas/> (redirige a la web app, que pide usuario y clave)

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
5. En el editor, selecciona la función **`setup`** y pulsa **Ejecutar**. Acepta los permisos. Esto crea la planilla, el disparador diario y la clave de acceso inicial (usuario, clave y URL de la planilla aparecen en el registro).
6. **Implementar → Nueva implementación → Aplicación web**. *Ejecutar como:* Yo. *Quién tiene acceso:* Cualquier persona (la app se protege con su propio login). Copia la URL y guárdala en favoritos / pantalla de inicio del celular.

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

## Seguridad y acceso (login)

La app tiene **login propio con usuario y clave**, para entrar desde cualquier equipo o celular sin depender de una cuenta de Google.

- **Implementación:** *Ejecutar como:* **Yo** · *Quién tiene acceso:* **Cualquier persona**. La pantalla de login es pública, pero **ningún dato se entrega sin sesión válida**: cada función del servidor exige el token de sesión.
- **Clave inicial:** en el editor de Apps Script ejecuta **`generarClaveAcceso`**. En *Registro de ejecución* aparecen el usuario (`admin`) y una clave aleatoria. Entra y cámbiala en **Configuración → Acceso**.
- **¿Olvidaste la clave?** Ejecuta de nuevo `generarClaveAcceso` en el editor: genera una nueva y cierra todas las sesiones. Solo funciona desde el editor (cuenta dueña), nunca desde la web.
- La clave se guarda como **hash SHA-256 iterado con sal** en las *Script Properties*; nunca en texto plano ni en la planilla.
- Sesiones de **12 horas**, o **30 días** con "Mantener sesión iniciada". **Salir** cierra la sesión en el servidor. Cambiar la clave cierra todas las sesiones.
- **Anti fuerza bruta:** tras 8 intentos fallidos el login se bloquea 15 minutos.
- Usa una clave larga y única (idealmente de un gestor de contraseñas): quien la tenga ve todos los datos.
- Solo se guardan los **últimos 4 dígitos** de la tarjeta. No registres números completos, CVV ni fechas de expiración.
- El token de Telegram se guarda en las *Script Properties*, no en la planilla ni en el repositorio.

## Archivos

- `Codigo.gs` — backend: base de datos, lógica de estados, revisión diaria y envío de alertas.
- `index.html` — interfaz web.
- `appsscript.json` — manifest (zona horaria, acceso de la web app).
- `.clasp.json.example` — plantilla para `clasp`.

## Actualizar la app publicada

Editar el código no cambia la versión que está en la URL `/exec`. Después de pegar los cambios (o hacer `clasp push`):
**Implementar → Gestionar implementaciones →** lápiz ✏️ → *Versión:* **Nueva versión** → **Implementar**. La URL sigue siendo la misma.

## Enlace de GitHub Pages

`site/index.html` es una página mínima publicada en GitHub Pages que redirige a la URL `/exec` de Apps Script. Los datos siguen protegidos por el login de la app.

- Se publica sola con el workflow `.github/workflows/pages.yml` cada vez que cambia `site/`.
- Activación (una sola vez): **Settings → Pages → Build and deployment → Source: GitHub Actions**.
- Si algún día cambia la URL `/exec` (nueva *implementación*, no nueva *versión*), actualízala en `site/index.html`.
