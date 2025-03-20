import json
import requests
import urllib3
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# Cargar configuraciones
settings = json.load(open("settings/settings.json"))
user_config = json.load(open("settings/user_id.json"))
sheet_config = json.load(open("settings/sheet_config.json"))

BITRIX_BASE_URL = settings["BITRIX_BASE_URL"].replace("{USER_ID}", user_config["USER_ID"])
BITRIX_URLS = {key: f"{BITRIX_BASE_URL}{endpoint}" for key, endpoint in settings["ENDPOINTS"].items()}

SERVICE_ACCOUNT_FILE = "fit-guide-433118-p4-193f4862b36c.json"
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

READ_RANGE_NAME = sheet_config["READ_RANGE_NAME"]
WRITE_BITRIX_URL = sheet_config["WRITE_BITRIX_URL"]

def get_credentials():
    return Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)

def read_from_sheet(sheet_id, range_name):
    creds = get_credentials()
    service = build("sheets", "v4", credentials=creds)
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=sheet_id, range=range_name).execute()
    return result.get("values", [])

def get_task_id_from_bitrix(task_name):
    url = BITRIX_URLS["BITRIX_GET_TASKS_URL"]
    params = {"FILTER": {"%TITLE": task_name}}  # Búsqueda más flexible

    # print(f"Buscando tarea en Bitrix: {task_name}")

    try:
        response = requests.post(url, json=params, verify=False)

        if response.status_code == 200:
            data = response.json()
            tasks = data.get("result", {}).get("tasks", [])

            # Ver todas las tareas encontradas
            # print(f"Todas las tareas encontradas: {tasks}")

            # Si hay varias tareas, buscar la que mejor coincida
            for task in tasks:
                if task_name.lower() in task["title"].lower():
                    print(f"Tarea encontrada: {task['id']} - {task['title']}")
                    return task["id"]

            print(f"No se encontró coincidencia exacta para '{task_name}', usando la primera disponible.")
            return tasks[0]["id"] if tasks else None

        print(f"Error en Bitrix: {response.status_code} - {response.text}")
    except Exception as e:
        print(f"Error en la API de Bitrix: {e}")

    return None

def write_link_to_sheet(sheet_id):
    creds = get_credentials()
    if not creds:
        print("Credenciales no válidas.")
        return

    service = build("sheets", "v4", credentials=creds)
    sheet = service.spreadsheets()

    # Leer los datos de la hoja
    result = sheet.values().get(spreadsheetId=sheet_id, range=READ_RANGE_NAME).execute()
    existing_values = result.get("values", [])

    if not existing_values:
        print("La hoja está vacía.")
        return

    values_to_update = []
    row_updates = []

    for i, row in enumerate(existing_values, start=2):  # Comenzamos en la fila 2 porque la 1 es encabezado
        if len(row) < 4:  # Asegurar que existan al menos 4 columnas (C y D)
            continue

        task_name = f"{row[2]} {row[3]}"  # Concatenar C y D (índices 2 y 3)
        task_id = get_task_id_from_bitrix(task_name)

        if task_id:
            task_link = f'=HYPERLINK("https://bitrix.kernotek.mx/company/personal/user/410/tasks/task/view/{task_id}/", "{task_id}")'
            values_to_update.append([task_link])
            row_updates.append(f"M{i}")  # Guardar la celda a actualizar (Ej: M2, M3, etc.)

    # Verificar si hay datos antes de actualizar
    if values_to_update:
        # Escribir valores en la hoja, pero asegurándonos de que se haga celda por celda
        for cell, value in zip(row_updates, values_to_update):
            body = {"values": [value]}
            result = sheet.values().update(
                spreadsheetId=sheet_id,
                range=f"Hoja 9!{cell}",
                valueInputOption="USER_ENTERED",
                body=body
            ).execute()
            print(f"Celda {cell} actualizada con: {value}")

        print(f"Se han actualizado {len(values_to_update)} celdas en total.")
    else:
        print("No se encontraron tareas para actualizar.")

def main():
    SHEET_ID = "tu_sheet_id_aqui"
    write_link_to_sheet(SHEET_ID)

if __name__ == "__main__":
    main()
