# Programador de Turnos 📅

Sistema web de programación y gestión de turnos para empleados (especificamente para la compañia Doria), integrado con Google Sheets para almacenamiento centralizado de datos.

## Características ✨

- **Autenticación segura**: Sistema de login con validación de credenciales
- **Gestión de turnos**: Programa turnos para empleados (T1, T2, T3 y variantes con extras)
- **Cálculo automático de horas**: Calcula automáticamente las horas trabajadas por semana
- **Integración con Google Sheets**: Sincronización de datos con hojas de cálculo de Google
- **Interfaz moderna**: Diseño responsivo con gradientes y efectos visuales
- **Buscar y contacto**: Secciones dedicadas para búsquedas y contacto

## Tecnologías utilizadas 🛠️

- **Frontend**: HTML5, CSS3, JavaScript vanilla
- **Backend**: Google Apps Script (GAS)
- **Almacenamiento**: Google Sheets
- **Fonts**: Google Fonts (Urbanist, Lato, Open Sans)
- **Iconos**: Material Symbols Outlined

## Estructura de turnos 🔄

El sistema soporta los siguientes tipos de turnos:

| Turno | Lunes | Martes | Miércoles | Jueves | Viernes | Sábado | Domingo | Total Horas |
|-------|-------|--------|-----------|--------|---------|--------|---------|-------------|
| T1 | 6.5 | 7.5 | 7.5 | 7.5 | 7.5 | 7.5 | 0 | 44 |
| T2 | 7.5 | 7.5 | 7.5 | 7.5 | 7.5 | 6.5 | 0 | 44 |
| T3 | 7.5 | 7.5 | 7.5 | 7.5 | 7.5 | 0 | 0 | 37.5 |
| T1 EXTRA | 7.5 | 7.5 | 7.5 | 7.5 | 7.5 | 7.5 | 8 | 53 |
| T2 EXTRA | 7.5 | 7.5 | 7.5 | 7.5 | 7.5 | 8 | 8 | 53.5 |
| T3 EXTRA | 7.5 | 7.5 | 7.5 | 7.5 | 7.5 | 8 | 8 | 53.5 |

## Archivos del proyecto 📁

```
Programador de Turnos/
├── index.html          # Interfaz principal
├── CodigoGs.js         # Scripts de Google Apps Script (backend)
├── style.css           # Estilos CSS
└── README.md           # Este archivo
```

## Instalación y despliegue 🚀

### Requisitos previos
- Cuenta de Google
- Google Sheets con datos de usuarios y empleados
- Acceso a Google Apps Script

### Pasos de instalación

1. **Crear un proyecto de Apps Script**
   - Ve a [Google Apps Script](https://script.google.com)
   - Crea un nuevo proyecto

2. **Configurar los IDs de hojas**
   - En `CodigoGs.js`, actualiza las constantes con tus IDs de Google Sheets:
   ```javascript
   const USERS_SHEET_ID = 'tu_id_aqui';
   const EMPLOYEES_SHEET_ID = 'tu_id_aqui';
   const BRIGADAS_DORIA_SHEET_ID = 'tu_id_aqui';
   ```

3. **Subir archivos**
   - Copia el contenido de `index.html` a un archivo HTML en Apps Script
   - Copia el contenido de `CodigoGs.js` al archivo `.gs` principal
   - Copia el contenido de `style.css` a un archivo CSS en Apps Script

4. **Desplegar como aplicación web**
   - Click en "Implementar" → "Nueva implementación"
   - Tipo: "Aplicación web"
   - Ejecutar como: Tu cuenta de Google
   - Acceso: "Cualquiera que tenga el enlace"

5. **Obtener el enlace público**
   - Copia la URL de despliegue

## Estructura de datos de Google Sheets 📊

### Hoja de Usuarios (`usuario`)
| Email | Contraseña | Cargo |
|-------|-----------|-------|
| usuario@email.com | contraseña | Gerente |

### Hoja de Empleados
- Contiene datos de empleados y asignaciones de turnos semanales

## Uso 📖

1. **Acceder a la aplicación**
   - Abre la URL del despliegue en tu navegador

2. **Iniciar sesión**
   - Ingresa tu email y contraseña registrados
   - El sistema validará tus credenciales contra Google Sheets

3. **Gestionar turnos**
   - Busca empleados
   - Asigna turnos según las necesidades
   - El sistema calcula automáticamente las horas

4. **Contacto**
   - Utiliza la sección de contacto para consultas

## Características de seguridad 🔒

- Validación de credenciales contra Google Sheets
- Autenticación basada en email y contraseña
- Roles de usuario con diferentes permisos

## Funciones principales 🔧

### `verificarCredenciales(correo, clave)`
Valida el login del usuario

### `obtenerInfoUsuario(email)`
Obtiene información del perfil del usuario

### `obtenerEmpleados()`
Recupera lista de empleados

### `asignarTurno(empleadoId, turno, semana)`
Asigna turno a un empleado

## Personalización 🎨

- Colores: Edita las variables CSS en `style.css` (azul oscuro, blanco)
- Logo: Actualiza la URL en la sección de header de `index.html`
- Fuentes: Modifica las importaciones de Google Fonts en `<head>`

## Notas importantes ⚠️

- El proyecto requiere conexión a internet para funcionar
- Los datos se sincronizan en tiempo real con Google Sheets
- Asegúrate de tener los permisos necesarios en las hojas de cálculo
- Guarda regularmente respaldos de tus datos en Google Sheets

## Soporte 📞

Para reportar bugs o sugerir mejoras, contactar con siendokevi@gmail.com  3144110953 Kevin Camilo Delgado R. 

## Licencia 📄

Este proyecto es propiedad de KEVIN CAMILO DELGADO RESTREPO. Todos los derechos reservados.

---

**Última actualización**: Diciembre 2026
