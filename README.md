# GEP Group · Certificados de Formaciones

Aplicación web para agilizar la generación de certificados de formaciones a partir de la
información almacenada en Pipedrive. El objetivo de esta primera versión es contar con
una interfaz amigable que permita recuperar los datos de un presupuesto (deal),
desglosar los alumnos asociados y preparar una tabla editable que actúa como base para
los siguientes pasos del flujo (creación de certificados, envío, etc.).

## Características principales

- **Experiencia UX optimizada**: interfaz responsive basada en Bootstrap 5 y tipografía
  Poppins para mantener consistencia visual con la identidad de la empresa.
- **Persistencia durante la sesión**: los datos introducidos se guardan en `sessionStorage`,
  permitiendo refrescar o navegar sin perder el trabajo hasta que se cierre la pestaña.
- **Integración con Pipedrive**: función serverless (Netlify) que consulta deals, notas,
  organizaciones y personas para recuperar todos los campos relevantes.
- **Normalización de alumnos**: extracción inteligente de la nota “Alumnos del Deal”,
  separación por columnas (nombre, apellidos, documento) y detección automática de
  DNI/NIE.
- **Gestión flexible de filas**: posibilidad de añadir filas manuales, editar cualquier
  campo directamente en tabla y eliminar filas erróneas. Incluye limpieza completa con
  confirmación.
- **Preparado para siguientes fases**: el back-end serverless devuelve también los datos
  de contacto (persona y email) y normaliza ubicaciones para facilitar el posterior uso
  con plantillas PDF y almacenamiento en Google Workspace.

## Estructura del proyecto

```
.
├── netlify
│   └── functions
│       └── fetch-deal.js    # Función serverless que habla con la API de Pipedrive
├── public
│   ├── assets
│   │   ├── css
│   │   │   └── styles.css   # Estilos propios (tipografía Poppins + refinamientos UX)
│   │   └── js
│   │       └── app.js       # Lógica del frontend, tabla editable y persistencia
│   └── index.html           # Landing principal con formulario y tabla
├── netlify.toml             # Configuración de despliegue en Netlify
├── package.json             # Scripts y metadatos del proyecto
├── .gitignore
└── README.md
```

## Requisitos previos

- Node.js 18 o superior (Netlify ejecuta actualmente Node 18/20 en funciones).
- [Netlify CLI](https://docs.netlify.com/cli/get-started/) instalado globalmente para el
  trabajo local (`npm install -g netlify-cli`).
- Credenciales de Pipedrive con permisos para leer deals, notas, organizaciones y
  personas.

## Configuración de variables de entorno

Define las siguientes variables en tu proyecto de Netlify (UI o `netlify env:set`):

| Variable                 | Descripción                                              |
| ------------------------ | -------------------------------------------------------- |
| `PIPEDRIVE_API_URL`      | URL base de la API de Pipedrive (`https://api.pipedrive.com/v1`). |
| `PIPEDRIVE_API_TOKEN`    | Token privado de la API de Pipedrive.                    |

> **Importante**: No publiques el token en el repositorio. Las variables se inyectan de
> forma segura en el runtime de Netlify y nunca quedan expuestas en el cliente.

## Ejecución en local

1. Clona el repositorio y entra en la carpeta del proyecto.
2. Instala dependencias (no hay dependencias adicionales en esta fase, pero crea
   `node_modules` si añades más adelante):

   ```bash
   npm install
   ```

3. Autentícate con Netlify (`netlify login`) si aún no lo has hecho.
4. Arranca el entorno local:

   ```bash
   netlify dev
   ```

   El comando levantará un servidor local y expondrá la función `fetch-deal` bajo
   `/.netlify/functions/fetch-deal`.

5. Abre `http://localhost:8888` en el navegador y utiliza el formulario introduciendo un
   número de presupuesto válido.

## Flujo de datos

1. **Formulario**: el usuario introduce el número de presupuesto y pulsa “Rellenar”.
2. **Función serverless**: se llama a `/.netlify/functions/fetch-deal?dealId=<ID>`.
3. **Consulta a Pipedrive**:
   - Recupera los datos del deal.
   - Obtiene organización y persona de contacto asociadas.
   - Lee las notas, detecta la última con el prefijo “Alumnos del Deal” y extrae cada
     alumno.
4. **Respuesta normalizada**: devuelve fecha, sede, nombre de la formación, cliente,
   persona/email de contacto y listado de alumnos.
5. **Render en tabla**: la tabla añade una fila por alumno (o una fila vacía si no hay
   nota). Se guardan los cambios en `sessionStorage` para mantenerlos hasta cerrar pestaña.

## Próximos pasos sugeridos

- Incorporar autenticación ligera antes de exponer la landing pública.
- Añadir generación de certificados PDF con `pdfmake` usando los datos de cada fila.
- Integrar subida automática a Google Workspace y disparo de correos vía Google Cloud.
- Añadir tests automatizados (por ejemplo, Jest para funciones y Playwright para UX).
- Crear un diseño específico para imprimir/exportar la tabla a PDF.

## Licencia

Distribuido bajo licencia MIT. Consulta el archivo `LICENSE` si se añade en futuras
iteraciones.
