# Integración de Google Sheets con Base de Datos Relacional SQL

Este proyecto tiene como objetivo conectar un documento de Google Sheets con una base de datos relacional SQL utilizando la API de Google Sheets. Esto permite realizar consultas SQL desde una celda en el documento y obtener los resultados en celdas específicas.

## Funcionalidades
- Conexión entre Google Sheets y una base de datos relacional (SQL).
- Ejecución de consultas SQL ingresadas en una celda del documento.
- Visualización de los resultados de las consultas directamente en otras celdas del documento.

## Requisitos
1. **Cuenta de Servicio de Google Cloud**: Necesaria para configurar la API de Google Sheets.
2. **Permisos en Google Sheets**: El archivo debe compartir permisos con la cuenta de servicio.
3. **Base de Datos Relacional**: Compatible con SQL y accesible desde el proyecto.
4. **Bibliotecas y Dependencias**:
   - [Google APIs Client Library](https://developers.google.com/sheets/api/quickstart/python).
   - Conector para la base de datos relacional (por ejemplo, `pyodbc`, `psycopg2`, etc.).

## Configuración
### 1. Crear una cuenta de servicio
- Accede a [Google Cloud Console](https://console.cloud.google.com/).
- Habilita la API de Google Sheets.
- Crea una cuenta de servicio y descarga el archivo de credenciales en formato JSON.
- Comparte el documento de Google Sheets con el correo electrónico de la cuenta de servicio.

### 2. Configurar la base de datos
- Asegúrate de que la base de datos esté configurada para aceptar conexiones externas.
- Configura un usuario con los permisos necesarios para ejecutar consultas.

### 3. Configurar el proyecto
- Instala las dependencias:
  ```bash
  pip install --upgrade google-api-python-client google-auth-httplib2 google-auth-oauthlib
  pip install <conector SQL específico>
