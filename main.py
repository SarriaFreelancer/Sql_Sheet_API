import mysql.connector
import time
from datetime import date, datetime
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

# Configuración de Google Sheets
SHEET_ID = '1NhBJj_e41PewFirIZ9-pneVlaYDcNqAW4VfGOjO7SCU'  # ID de tu hoja de cálculo
RANGE_QUERY = 'inicio!B1'  # Celda donde está la consulta
RANGE_RESULT = 'inicio!B2'  # Comienza en la columna B (sin especificar el rango final)

# Ruta al archivo JSON de la cuenta de servicio
SERVICE_ACCOUNT_FILE = 'D:/informacion/Desarrollo/learny/hobbie/prueba-sql-447904-6cfd1a0685dd.json'

# Alcances requeridos por la API
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# Crear credenciales
credentials = Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE, scopes=SCOPES
)

# Construir el servicio de Google Sheets
service = build('sheets', 'v4', credentials=credentials)

# Configuración de la base de datos
db_config = {
    "host": "localhost",
    "user": "root",
    "password": "root",
    "database": "academia",
    "port": 3307
}

# Función para leer la consulta desde Google Sheets
def leer_consulta_de_google_sheets():
    try:
        result = service.spreadsheets().values().get(
            spreadsheetId=SHEET_ID,
            range=RANGE_QUERY
        ).execute()
        consulta = result.get('values', [['']])[0][0]  # Obtener la consulta de la celda A1
        print(f"Consulta leída desde Google Sheets: {consulta}")
        return consulta
    except Exception as e:
        print(f"Error al leer la consulta de Google Sheets: {e}")
        return None

# Función para ejecutar una consulta en la base de datos
def ejecutar_consulta(query):
    try:
        connection = mysql.connector.connect(**db_config)
        cursor = connection.cursor()
        cursor.execute(query)
        resultados = cursor.fetchall()
        columnas = [desc[0] for desc in cursor.description]  # Obtener los nombres de las columnas
        connection.close()
        return columnas, resultados  # Regresar las columnas y los resultados
    except mysql.connector.Error as e:
        print(f"Error al conectar a la base de datos: {e}")
        return None, None

def formatear_datos(data):
    """
    Convierte objetos no serializables como fechas en cadenas de texto.
    """
    formatted_data = []
    for fila in data:
        formatted_row = []
        for valor in fila:
            if isinstance(valor, (date, datetime)):  # Si es tipo fecha
                formatted_row.append(valor.strftime('%Y-%m-%d'))  # Formatea como AAAA-MM-DD
            elif valor is None:  # Si el valor es None
                formatted_row.append("")  # Sustitúyelo por una cadena vacía
            else:
                formatted_row.append(valor)  # Deja el valor como está
        formatted_data.append(formatted_row)
    return formatted_data

# Función para enviar datos a Google Sheets# Función para enviar datos a Google Sheets
def enviar_a_google_sheets(data, range_name):
    try:
        body = {
            'values': data
        }
        service.spreadsheets().values().update(
            spreadsheetId=SHEET_ID,
            range=range_name,
            valueInputOption='RAW',  # Usar "RAW" para que no haya formateo automático
            body=body
        ).execute()
        print("Datos enviados a Google Sheets correctamente.")
    except Exception as e:
        print(f"Error al enviar datos a Google Sheets: {e}")


# Función para limpiar la hoja (eliminar bordes, colores y datos)
def limpiar_hoja():
    try:
        # Limpiar el contenido de las celdas desde B2 en adelante
        service.spreadsheets().values().clear(
            spreadsheetId=SHEET_ID,
            range="inicio!B2:Z1000"  # Rango amplio que cubre los posibles datos
        ).execute()

        # Limpiar bordes y colores (restaurar a estado limpio)
        requests = [
            {
                'updateCells': {
                    'range': {
                        'sheetId': 0,
                        'startRowIndex': 0,
                        'startColumnIndex': 0,
                        'endRowIndex': 1000,  # Ajusta según el número de filas que esperas
                        'endColumnIndex': 26  # Hasta la columna Z (puedes ajustar si es necesario)
                    },
                    'fields': 'userEnteredFormat.backgroundColor, userEnteredFormat.borders'
                }
            }
        ]
        service.spreadsheets().batchUpdate(
            spreadsheetId=SHEET_ID,
            body={'requests': requests}
        ).execute()

        print("Hoja limpiada correctamente.")
    except Exception as e:
        print(f"Error al limpiar la hoja: {e}")

# Función para enviar mensajes de error o resultados vacíos
def enviar_mensaje_error_o_vacio(mensaje):
    try:
        body = {
            'values': [[mensaje]]  # El mensaje que quieres mostrar en la celda
        }
        service.spreadsheets().values().update(
            spreadsheetId=SHEET_ID,
            range='inicio!B2',  # Rango donde quieres mostrar el mensaje
            valueInputOption='RAW',  # Usar "RAW" para que no haya formateo automático
            body=body
        ).execute()
        print("Mensaje de error o vacío enviado a Google Sheets.")
    except Exception as e:
        print(f"Error al enviar el mensaje a Google Sheets: {e}")

# Función para aplicar bordes y color de fondo verde solo en el encabezado desde la columna B
# Función para aplicar bordes y color de fondo verde solo en el encabezado
def aplicar_bordes_y_color(range_name, num_columns, num_filas):
    requests = [
        # Bordes en todo el rango de datos
        {
            'updateBorders': {
                'range': {
                    'sheetId': 0,  # ID de la hoja, puedes obtenerlo si no es la primera
                    'startRowIndex': 1,  # Comienza desde la fila 2 (B2)
                    'startColumnIndex': 1,  # Comienza desde la columna B (índice 1)
                    'endColumnIndex': num_columns + 1 ,  # Número de columnas con datos
                    'endRowIndex': num_filas + 2  # Ajusta según el número de filas con datos
                },
                'top': {'style': 'SOLID', 'width': 1},
                'bottom': {'style': 'SOLID', 'width': 1},
                'left': {'style': 'SOLID', 'width': 1},
                'right': {'style': 'SOLID', 'width': 1},
            }
        },
        # Bordes en la línea inferior del encabezado (solo desde la columna B)
        {
            'updateBorders': {
                'range': {
                    'sheetId': 0,
                    'startRowIndex': 1,  # La primera fila de datos (fila 2)
                    'startColumnIndex': 1,  # Comienza desde la columna B
                    'endColumnIndex': num_columns + 1  # Número de columnas de encabezado
                },
                'bottom': {'style': 'SOLID', 'width': 2},
            }
        },
        # Color verde para el encabezado (solo desde la columna B en adelante)
        {
            'updateCells': {
                'range': {
                    'sheetId': 0,
                    'startRowIndex': 1,  # La fila 2 (B2) en adelante
                    'startColumnIndex': 1,  # Comienza desde la columna B (índice 1)
                    'endColumnIndex': num_columns + 1  # Número de columnas de encabezado
                },
                'rows': [{
                    'values': [{
                        'userEnteredFormat': {
                            'backgroundColor': {
                                'red': 0.0,  # Rojo (0.0 es sin color)
                                'green': 1.0,  # Verde
                                'blue': 0.0   # Azul (0.0 es sin color)
                            }
                        }
                    } for _ in range(num_columns)]  # Número de columnas en los encabezados
                }] ,  # Solo una fila para los encabezados
                'fields': 'userEnteredFormat.backgroundColor'
            }
        }
    ]
    
    # Aplicar las solicitudes de bordes y color de fondo verde solo en el encabezado
    try:
        service.spreadsheets().batchUpdate(
            spreadsheetId=SHEET_ID,
            body={'requests': requests}
        ).execute()
        print("Bordes y color de fondo aplicados correctamente.")
    except Exception as e:
        print(f"Error al aplicar bordes y color de fondo: {e}")


import time

# Función principal para manejar la ejecución del script
def ejecutar_consulta_y_actualizar():
    # Limpiar la hoja antes de insertar nuevos datos
    limpiar_hoja()

    # Leer consulta desde Google Sheets
    consulta = leer_consulta_de_google_sheets()

    if consulta:
        # Ejecutar consulta
        columnas, resultados = ejecutar_consulta(consulta)

        if resultados:
            # Obtener el número de filas de resultados
            num_filas = len(resultados)

            # Convertir los resultados a un formato adecuado para Google Sheets (listas de listas)
            data_to_send = [list(fila) for fila in resultados]

            # Agregar encabezados automáticamente desde las columnas de la consulta
            data_to_send.insert(0, columnas)  # Insertar los encabezados al principio de los datos

            # Formatea los datos para evitar errores de serialización
            data_to_send = formatear_datos(data_to_send)

            # Enviar resultados a Google Sheets, comenzando en la celda B2
            enviar_a_google_sheets(data_to_send, RANGE_RESULT)

            # Aplicar bordes y color de fondo verde solo al encabezado
            aplicar_bordes_y_color(RANGE_RESULT, len(columnas), num_filas)

        else:
            mensaje = "No se obtuvieron resultados de la consulta."
            print(mensaje)
            enviar_mensaje_error_o_vacio(mensaje)  # Enviar mensaje de error a Google Sheets
    else:
        mensaje = "La consulta leída desde Google Sheets está vacía o no se pudo leer."
        print(mensaje)
        enviar_mensaje_error_o_vacio(mensaje)  # Enviar mensaje de error a Google Sheets

# Función para ejecutar el ciclo de revisión cada cierto tiempo
def ejecutar_en_ciclo():
    while True:
        print("Revisando la consulta en Google Sheets...")

        # Llamar a la función para ejecutar la consulta y actualizar Google Sheets
        ejecutar_consulta_y_actualizar()

        # Esperar 30 segundos antes de revisar nuevamente
        print("Esperando 30 segundos antes de revisar nuevamente la consulta...")
        time.sleep(30)  # Ajusta el tiempo de espera según lo necesites

# Ejecución principal
if __name__ == "__main__":
    # Ejecutar el ciclo de revisión continuo
    ejecutar_en_ciclo()
