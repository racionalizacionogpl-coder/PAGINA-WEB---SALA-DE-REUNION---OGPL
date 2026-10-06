# Reserva de Salas – OGPL UNMSM (v4.3: diseño corporativo según el Manual de Marca Gestión)

Página web publicada desde Google Apps Script. Cada persona ingresa con su **correo y contraseña**, completa el formulario y la reserva queda en Google Sheets. Al reservar o cancelar se envía un correo con el encabezado **"Oficina General de Planificación – UNMSM"** y el escudo de la universidad.

```
asistente-reservas-ogpp/
├── apps-script/                 ← lo único que se copia a Apps Script
│   ├── Code.gs                  Backend: sesiones, usuarios, reservas, correos, setup()
│   ├── Logo.gs                  Escudo UNMSM en base64 (para correos y página)
│   ├── Fotos.gs                 Fotos de la Sala 1 y la Sala 2 en base64 (nuevo en v4.3)
│   ├── Index.html               La página: ingreso, inicio, reservas, calendario, administración
│   └── appsscript.json          Manifiesto
├── recursos/                    Logo y fotos originales de las salas (recursos/fotos/)
├── tests/                       Pruebas automáticas (44 de backend + 24 pasos en navegador)
└── pospuesto-whatsapp/          Versión con WhatsApp y n8n (archivada, no se usa)
```

---

## Novedades de la v4.3: diseño corporativo

La página y los correos siguen ahora el **Manual de Marca Gestión** (DGP · OGPL-UNMSM), con el escudo de la UNMSM como firma:

- **Colores:** navy `#0B1D3A` / `#14315F` y acento teal `#1BB0C4` / `#0A7E90`, en proporción 60 · 30 · 10.
- **Tipografías:** Montserrat (títulos), Source Sans 3 (texto) e IBM Plex Mono (códigos y horas).
- **Estilo:** iconos de línea, sin emojis, y los estados siempre con palabra + color.
- **Modo oscuro:** se activa solo si el equipo lo usa.

| Pantalla | Qué cambió |
| --- | --- |
| **Ingreso / Crear cuenta / Contraseña** | Panel de marca a la izquierda: foto de la Sala 1 con velo navy, escudo y "Reserva tu sala en tiempo real". El formulario va a la derecha. En el celular el panel se reduce a una franja. |
| **Inicio (nuevo, vista por defecto)** | **Ahora mismo:** panel con el estado de cada sala (Libre / Ocupada hasta…).<br>**Indicadores:** cinco bloques (reservas del día, salas libres, tus próximas reservas, ocupación del día, horas de la semana).<br>**Tarjetas de sala:** carrusel con las 3 fotos de cada sala, capacidad, ubicación, equipamiento, barra con la agenda del día y botón **Reservar Sala N**.<br>**Cuenta regresiva:** hasta tu próxima reunión. |
| **Reservar** | Proceso en 3 pasos. En el paso 1 la sala se elige con **tarjetas con foto** (incluye "Cualquier sala libre"). A la derecha hay un **resumen fijo** con la foto, los datos elegidos, el resultado de la verificación y el botón **Reservar**. |
| **Calendario, Mis reservas, Administración** | Misma lógica de antes, con la franja de título, tarjetas y tablas del manual. |
| **Barra superior y pie** | Barra navy de 64 px: pestaña activa subrayada en teal, botón **Nueva reserva**, iniciales del usuario, cambiar contraseña y salir.<br>Pie navy en 4 columnas: horario, salas, oficinas y versión, con botón para volver arriba. |
| **Correos** | Cabecera navy con el escudo y "Oficina General de Planificación – UNMSM", línea teal, estado en etiqueta de color ("Reserva confirmada", "Reserva cancelada"…) y pie navy. |

**Fotos.** Las tres primeras imágenes son de la **Sala 1** y las tres siguientes de la **Sala 2**. Van dentro de `Fotos.gs` (recortadas a 4:3, 800×600 px) y la página las pide una vez con `apiFotos()`; luego quedan guardadas en la pestaña del navegador.
- Las fotos se asocian por el **nombre de la sala**, sin distinguir mayúsculas ni espacios. Si renombras una sala en la pestaña `Salas`, renombra también su clave en `Fotos.gs`.
- **Para cambiar o agregar fotos:** convierte cada JPG (de unos 800×600 px y menos de 100 KB) a base64 con cualquier conversor "imagen a base64". Pega el texto en el arreglo de la sala dentro de `FOTOS_SALAS`, por ejemplo `'Sala 3': ['…', '…']`.
- Las fotos originales están en `recursos/fotos/`.

**Completa la pestaña `Salas`:** las tarjetas muestran las columnas `Ubicacion` (por ejemplo, "2.º piso – OGPL") y `Equipamiento` (por ejemplo, "Proyector, TV, 6 laptops"). Si dicen "Por definir", reemplaza ese texto en la hoja.

---

## Novedades de la v4.2

| Cambio | Detalle |
| --- | --- |
| **Proyector y laptops** | Al reservar se debe indicar **Sí / No** para el proyector y para las laptops (obligatorio). Se guarda en `Reservas` (columnas `Usa_Proyector`, `Usa_Laptops`) y aparece en la confirmación, en el aviso a los encargados, en el detalle y en *Administración* (columna "Equipos"). |
| **Lista de oficinas** | Nueva pestaña `Oficinas` con las 5 oficinas de la OGPL. Al crear una cuenta (o un administrador al crear o editar un usuario) la oficina se elige de una lista desplegable. |
| **Prioridad de la Oficina de Presupuesto** | Si alguien de una oficina con `Prioridad = SI` intenta reservar una sala ocupada, además del aviso "Sala ocupada" ve el botón **"Reservar … con prioridad"**. Al confirmarlo (puede escribir un motivo), la reserva anterior queda cancelada y se envía un correo a quien la hizo **y al correo de su oficina**, explicando que fue ocupada por la Oficina de Presupuesto y **sugiriendo horarios disponibles** de la misma duración. |

### Pestaña `Oficinas`

| Oficina | Prioridad | Correo | Activa |
| --- | --- | --- | --- |
| Oficina de Presupuesto | **SI** | *(completar)* | SI |
| Oficina de Racionalización | NO | *(completar)* | SI |
| Oficina de Estadística e Informática | NO | *(completar)* | SI |
| Oficina de Planes y Programas | NO | *(completar)* | SI |
| Oficina de Coordinación de Producción | NO | *(completar)* | SI |

- **Correo:** el correo institucional de cada oficina. Recibe copia cuando una reserva de esa oficina es desplazada. Si se deja vacío, solo se avisa a la persona.
- **Prioridad:** puede tenerla más de una oficina. Una oficina con prioridad **no puede desplazar a otra que también la tenga**.
- **Reuniones en curso:** nunca se desplaza una reunión que ya comenzó.
- **Cambiar o agregar oficinas:** se hace en esta pestaña. Los usuarios cuya oficina no esté en la lista pueden seguir reservando, pero sin prioridad.

### Reglas de la reserva con prioridad

1. La prioridad solo se usa si la persona pulsa **"Reservar con prioridad"** y confirma. Nunca es automática.
2. Si eligió "Cualquier sala libre", el sistema propone la sala que afecta a menos reservas.
3. La reserva desplazada queda **Cancelada**, con el motivo "Desplazada por reserva prioritaria RES-… de la Oficina de Presupuesto". La nueva reserva registra a quién reemplazó (columna `Reemplaza_A`).
4. Los horarios sugeridos **no quedan apartados**. El correo incluye un botón para entrar al sistema y reservar.

> **Al actualizar:** revisa en *Administración → Usuarios* que la oficina de cada persona coincida con un nombre de la lista. Por ejemplo, quien deba tener prioridad tiene que figurar exactamente como "Oficina de Presupuesto". Corrígela con **Editar**.

## ¿La pantalla de ingreso no muestra el escudo ni "Crear una cuenta"?

Pasa cuando el servidor tiene un `Code.gs` anterior al de la página, o cuando no se publicó una **nueva versión**. Desde la v4.1 la página lo detecta y muestra un aviso en rojo. Solución:

1. Comprueba que en el proyecto existan exactamente **un** `Code`, **un** `Logo`, **un** `Fotos` y **un** `Index`, y que no queden copias viejas con otro nombre (por ejemplo `Código.gs` además de `Code.gs`).
2. *Implementar → Gestionar implementaciones → ✏️ → Versión: **Nueva versión** → Implementar*.
3. Recarga la página con Ctrl+F5.

## A. Si ya tienes la versión anterior funcionando (actualización, ≈ 15 min)

1. **Abre el proyecto** de Apps Script desde la hoja: *Extensiones → Apps Script*.
2. **Reemplaza `Code.gs`** por el nuevo contenido.
3. **Reemplaza `Index`** (el archivo HTML) por el nuevo `Index.html`.
4. **Crea dos archivos de script nuevos** (*+ → Secuencia de comandos*) si aún no existen:
   - **`Logo`**, con el contenido de `Logo.gs`.
   - **`Fotos`**, con el contenido de `Fotos.gs`. Este archivo es grande (≈ 380 KB): ábrelo con un editor de texto, selecciona todo y cópialo.
5. **Reemplaza `appsscript.json`** por el nuevo.
6. Guarda (Ctrl+S). Elige la función **`setup`** y pulsa **Ejecutar**:
   - Agrega a la pestaña `Usuarios` las columnas de contraseña, sin tocar los datos existentes.
   - Agrega a `Config` las claves nuevas (`NOMBRE_OFICINA`, `MINUTOS_SESION`, `PERMITIR_SOLICITUDES`, `REGISTRO_REQUIERE_APROBACION`, `DOMINIOS_PERMITIDOS`).
   - Crea la pestaña `Oficinas` con las 5 oficinas, y agrega a `Reservas` las columnas `Usa_Proyector`, `Usa_Laptops` y `Reemplaza_A`.
   - Genera una **contraseña temporal para el administrador**. Ábrela en *Registro de ejecución* (abajo en el editor): verás `Usuario administrador: …` y `Contraseña temporal: …`. **Anótala.**
7. **Actualiza la implementación:**
   1. Ve a *Implementar → Gestionar implementaciones → ✏️ Editar*.
   2. Configúrala así:
      - Versión: **Nueva versión**.
      - Ejecutar como: **Yo**.
      - Quién tiene acceso: **Cualquier persona**. Ahora el control de acceso lo hace la propia página con usuario y contraseña.
   3. Pulsa *Implementar*. La URL no cambia.
8. **Entra a la URL** con el correo del administrador y la contraseña temporal. El sistema te pedirá crear una contraseña propia.
9. **Da contraseña a los usuarios que ya estaban** en la hoja. Las cuentas antiguas no tienen contraseña, así que no pueden entrar hasta que hagas esto:
   1. Ve a *Administración → Usuarios*.
   2. En cada persona, pulsa **Contraseña**. El sistema genera una temporal y, si marcas la casilla, se la envía por correo.

## B. Instalación desde cero

1. Crea la hoja **BD_Reservas_Salas_OGPL** y abre *Extensiones → Apps Script*.
2. Pega los archivos:
   - `Code.gs` en el archivo principal.
   - `Logo.gs` en un archivo de script nuevo llamado `Logo`, y `Fotos.gs` en otro llamado `Fotos`.
   - `Index.html` en un archivo HTML llamado `Index`.
   - `appsscript.json` en el manifiesto. Para verlo, ve a *Configuración del proyecto* y activa la opción de mostrar el manifiesto.
3. Ejecuta **`setup`**, acepta los permisos y anota la contraseña temporal del administrador que aparece en el registro de ejecución.
4. Completa las pestañas `Salas` (con el correo del encargado de cada una), `Config` y `Feriados`.
5. Publica: *Implementar → Nueva implementación → Aplicación web*, con "Ejecutar como: **Yo**" y "Quién tiene acceso: **Cualquier persona**". Copia la URL `/exec`.
6. Entra con el administrador, cambia la contraseña y crea los usuarios en *Administración → Usuarios*.

---

## Cómo funciona el acceso

| Situación | Qué pasa |
| --- | --- |
| **Ingreso** | Correo + contraseña. La sesión se cierra tras 120 min sin actividad (`Config › MINUTOS_SESION`, máximo 360) o al pulsar **Salir**. |
| **Contraseña olvidada** | Un administrador pulsa **Contraseña** en *Administración → Usuarios* y asigna una temporal (puede enviarla por correo). |
| **Administrador que olvidó la suya** | En la hoja: menú *Reservas OGPL → Asignar contraseña a un usuario…* (solo el propietario del script). |
| **Usuario creado por un administrador** | Entra con la contraseña temporal y **debe cambiarla** antes de usar el sistema. |
| **Crear una cuenta** | Botón en la pantalla de ingreso. La persona registra nombre, oficina, correo y contraseña. **Todos los administradores activos reciben un correo "Nueva cuenta registrada"**, y la persona recibe una confirmación. |
| **¿Aprobación o acceso inmediato?** | `Config › REGISTRO_REQUIERE_APROBACION`. Con `SI` (valor inicial), la cuenta queda **Pendiente** hasta que un administrador la apruebe en *Administración → Usuarios*, y al aprobarla la persona recibe un aviso. Con `NO`, la cuenta queda **activa de inmediato** y el correo a los administradores es solo informativo; pueden desactivarla si no corresponde. |
| **Quitar el botón "Crear una cuenta"** | `Config › PERMITIR_SOLICITUDES = NO`. Así solo los administradores crean usuarios. |
| **Restringir dominios** | `Config › DOMINIOS_PERMITIDOS = unmsm.edu.pe` hace que solo se acepten cuentas con ese dominio. Vacío = cualquiera. |
| **Bloqueo por intentos** | 5 contraseñas incorrectas seguidas bloquean ese correo 15 minutos. |
| **Desactivar a alguien** | *Administración → Usuarios → Desactivar*. Se cierran sus sesiones abiertas; sus reservas se conservan. |

**Seguridad.** Las contraseñas nunca se guardan: la hoja solo tiene un *hash* con sal (SHA-256 iterado) que no permite recuperarlas. Cambiar o restablecer una contraseña cierra las sesiones anteriores. La oficina y el nombre de cada reserva salen de la cuenta, no del formulario. Las funciones de mantenimiento (`setup`, asignar contraseña) solo las puede ejecutar el propietario desde el editor o la hoja.

## Correos

Todos los correos (confirmación, cancelación, aviso a encargados, datos de acceso, nueva cuenta registrada, cuenta creada o aprobada) llevan una cabecera navy con el escudo de la UNMSM en un círculo blanco y una línea teal y el texto de `Config › NOMBRE_OFICINA` (**Oficina General de Planificación – UNMSM**). Debajo va el estado: "Reserva confirmada", "Reserva cancelada", etc.

- **Cambiar el texto:** edita `NOMBRE_OFICINA` en la pestaña `Config`.
- **Cambiar el logo:** reemplaza el texto de `Logo.gs` por el base64 de otro PNG cuadrado de unos 144 px, usando cualquier conversor "imagen a base64". El original está en `recursos/`.

## Cambios comunes

| Quiero… | Dónde |
| --- | --- |
| Crear, editar, desactivar usuarios o resetear contraseñas | Página → *Administración → Usuarios* |
| Hacer administrador a alguien | *Administración → Usuarios → Editar → Rol* |
| Agregar una sala o su encargado | Pestaña `Salas` |
| Dar o quitar prioridad a una oficina, o poner su correo | Pestaña `Oficinas` |
| Cambiar quién recibe todos los avisos | `Config › CORREOS_NOTIFICACION` |
| Cambiar horario, duración, anticipación | Pestaña `Config` |
| Agregar feriados | Pestaña `Feriados` |

> No edites a mano las columnas `Hash_Contrasena` ni `Sal` de la pestaña `Usuarios`. Si una cuenta queda inservible, asígnale una contraseña nueva desde la página o desde el menú de la hoja.

## Límites

- **Correos:** Apps Script permite unos 100 envíos diarios en cuentas Gmail y 1500 en Google Workspace. Cada reserva usa hasta 2.
- **Cuenta propietaria:** la página funciona con la cuenta que la publicó. Si esa cuenta pierde acceso a la hoja, el sistema deja de funcionar. Da acceso de editor del proyecto a una segunda persona de confianza.
- **Datos personales (Ley N.° 29733):** restringe quién puede abrir la hoja y haz copias periódicas (*Archivo → Hacer una copia*).

## Pruebas automáticas

```bash
node tests/test_backend.js     # 44 pruebas: acceso, registro, usuarios, reservas, prioridad, correos
node tests/test_ui.js          # 24 pasos en Chromium (requiere: npm i playwright)
```

Simulan los servicios de Google. Lo que solo se comprueba en Google es la entrega real de correos (y cómo muestra cada cliente de correo la imagen incrustada) y el comportamiento de la caché de sesiones.
