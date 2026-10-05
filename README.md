# Reserva de Salas – OGPP UNMSM

Aplicación web en Google Apps Script con Google Sheets como base de datos.

| Archivo | Función |
|---|---|
| `Code.gs` | Backend: reglas de reserva, hojas, correos, identidad |
| `Index.html` | Interfaz web (reservar, calendario, mis reservas, administración) |
| `appsscript.json` | Manifiesto: zona horaria, permisos y acceso de la web app |

## Puesta en marcha

1. **Crear la hoja**: en Google Sheets (cuenta institucional) crea una hoja, p. ej. "Reservas de Salas OGPP".
2. **Abrir Apps Script desde la hoja**: Extensiones → Apps Script (debe ser un script vinculado a la hoja).
3. **Cargar archivos**:
   - Pega `Code.gs` en el archivo de código.
   - Crea un archivo HTML llamado exactamente `Index` y pega `Index.html`.
   - En Configuración del proyecto activa "Mostrar el archivo de manifiesto appsscript.json" y reemplaza su contenido por `appsscript.json`.
4. **Inicializar**: ejecuta `setup()` y acepta los permisos. Crea las pestañas `Reservas`, `Usuarios`, `Salas`, `Config`, `Feriados` y `Log`, y te registra como ADMIN.
5. **Completar datos**:
   - `Usuarios`: correo (minúsculas), nombre, oficina, rol (`ADMIN`/`USUARIO`), activo (`SI`).
   - `Salas`: nombre, capacidad, ubicación, equipamiento, correo del responsable, activa (`SI`).
   - `Config`: horario, días hábiles, duraciones, correos de notificación, administradores.
   - `Feriados`: agrega los días no laborables decretados.
6. **Probar**: ejecuta `probar()` y revisa el registro de ejecución.
7. **Publicar**: Implementar → Nueva implementación → Aplicación web.
   - Ejecutar como: yo.
   - Acceso: cualquier persona de la organización.
   - Comparte la URL `/exec`.
8. **Actualizar**: tras cambiar código, Implementar → Administrar implementaciones → Editar → Nueva versión.

## Notas

- Requiere Google Workspace: `Session.getActiveUser()` solo devuelve el correo a usuarios del mismo dominio que quien publica.
- Cuota de correos: ~1.500/día en Workspace (100 en cuentas personales).
- Las celdas de fecha y hora son texto plano; no cambies su formato.
- `LockService` evita dobles reservas simultáneas.
