# Asistencia QR — Liceo Tiuquilemu
### Cómo dejar la plataforma funcionando de forma permanente, gratis, y accesible desde cualquier computador

Esta carpeta contiene una plataforma web completa. Para que funcione necesitas hacer
**dos pasos, una sola vez**. Después de eso, queda funcionando por años sin volver a tocarlo.

---

## Paso 1 — Crear la base de datos (Google Apps Script)

1. Ve a **sheets.google.com** y crea una hoja de cálculo nueva. Llámala, por ejemplo,
   "Asistencia Liceo Tiuquilemu".
2. Ve al menú **Extensiones → Apps Script**.
3. Borra el código de ejemplo que aparece y pega **todo** el contenido del archivo
   `Code.gs` (está en esta misma carpeta).
4. Haz clic en **Implementar → Nueva implementación**.
   - Tipo: **Aplicación web**.
   - Ejecutar como: **tu cuenta**.
   - Quién tiene acceso: **Cualquier usuario**.
5. Google te pedirá autorizar permisos — son permisos sobre tu propia cuenta, acéptalos.
6. Copia la **URL de la aplicación web** que te entrega (empieza con
   `https://script.google.com/macros/s/.../exec`). La vas a necesitar en el Paso 3.

Esta hoja de cálculo es ahora la **única base de datos real** del sistema. Todos los
computadores que usen la plataforma leen y escriben ahí mismo.

---

## Paso 2 — Publicar la plataforma en internet (GitHub Pages, gratis)

1. Crea una cuenta gratuita en **github.com** (si no tienes una).
2. Crea un repositorio nuevo, por ejemplo llamado `asistencia-qr`. Puede ser público.
3. Dentro del repositorio, sube **todos los archivos de esta carpeta** (arrástralos con
   el botón "Add file → Upload files"): `index.html`, `app.js`, `manifest.webmanifest`,
   `service-worker.js`, la carpeta `vendor` y la carpeta `icons`. (El archivo `Code.gs`
   y esta guía no hace falta subirlos, son solo para ti).
4. Ve a **Settings → Pages** dentro del repositorio.
5. En "Source", elige la rama `main` y la carpeta `/ (root)`, y guarda.
6. GitHub te entrega una dirección fija, algo como:
   `https://tu-usuario.github.io/asistencia-qr/`
   Esa es la dirección permanente de tu plataforma — la misma para siempre, desde
   cualquier computador.

---

## Paso 3 — Conectar y configurar (una vez por computador)

1. Abre la dirección del Paso 2 en el navegador (Chrome, Edge, Firefox…).
2. La primera vez te va a pedir la **URL de tu Google Apps Script** (la del Paso 1).
   Pégala y presiona "Conectar". Esto se guarda en ese navegador y no se vuelve a
   pedir en ese computador.
3. Luego te va a pedir **crear la clave de administrador** (y, si quieres, también la
   clave de usuario común en ese mismo momento).
4. ¡Listo! Ya puedes agregar alumnos y empezar a escanear.

Para usarla en otro computador, repite solo el paso 3.1 y 3.2 con la misma URL de
Google Apps Script — así todos comparten la misma información.

---

## "Instalarla" como una app (opcional pero recomendado)

Con la plataforma abierta en Chrome o Edge, busca en la barra de direcciones un ícono
de "Instalar" (o en el menú ⋮ → "Instalar Asistencia QR"). Esto la deja con su propio
ícono en el escritorio o el celular, y se abre como una aplicación aparte, sin barra
del navegador.

---

## Roles de acceso

- **Administrador** (con clave): agrega, edita y borra alumnos; entra a Ajustes;
  configura el informe automático y las claves.
- **Usuario común** (con otra clave): puede escanear, ver historial/estadísticas y
  **agregar alumnos nuevos**, pero no puede editar ni borrar datos existentes, ni
  entrar a Ajustes.

## Por qué esto va a seguir funcionando por años

- La base de datos es tu propia cuenta de Google (gratuita, sin fecha de vencimiento).
- La página está publicada en GitHub Pages (gratuito, sin fecha de vencimiento).
- No depende de Claude, de Anthropic, ni de ningún servicio de pago.
- Todas las librerías que usa (para leer QR, generar QR, gráficos y Excel) están
  incluidas dentro de esta misma carpeta — no dependen de enlaces externos que puedan
  romperse con el tiempo.
