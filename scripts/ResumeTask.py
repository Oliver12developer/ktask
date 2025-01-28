import requests
import json
import re
import urllib3
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

# Desactivar advertencias SSL (opcional, solo si Bitrix usa HTTPS sin certificado válido)
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# Función para cargar JSONs de configuración de forma segura
def load_json(filename):
    try:
        with open(filename, "r", encoding="utf-8") as file:
            return json.load(file)
    except (FileNotFoundError, json.JSONDecodeError):
        print(f"❌ Error al cargar {filename}")
        exit(1)

# Cargar configuraciones desde los archivos JSON
settings = load_json("settings/settings.json")
user_config = load_json("settings/user_id.json")
sheet_config = load_json("settings/SheetURL.json")

# Extraer ID de la hoja desde la URL en `SheetURL.json`
def extract_sheet_id(sheet_url):
    match = re.search(r"/d/([a-zA-Z0-9-_]+)/", sheet_url)
    return match.group(1) if match else None

SHEET_ID = extract_sheet_id(sheet_config["SHEET_URL"])
if not SHEET_ID:
    print("No se pudo extraer el ID de la hoja de cálculo.")
    exit(1)

# Construcción de URLs dinámicas de Bitrix
BITRIX_BASE_URL = settings["BITRIX_BASE_URL"].replace("{USER_ID}", user_config["USER_ID"])
BITRIX_URLS = {key: f"{BITRIX_BASE_URL}{endpoint}" for key, endpoint in settings["ENDPOINTS"].items()}

# Configuración de Google Sheets
SERVICE_ACCOUNT_FILE = "fit-guide-433118-p4-193f4862b36c.json"
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

credentials = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
service = build("sheets", "v4", credentials=credentials)

WRITE_RANGE_NAME = "Hoja 2!R2:R"

# 🔹 Función para obtener tareas desde Bitrix
def get_tasks_from_bitrix():
    response = requests.post(BITRIX_URLS["BITRIX_GET_TASKS_URL"], verify=False)

    if response.status_code == 200:
        data = response.json()
        tasks = data.get("result", {}).get("tasks", [])
        return tasks if tasks else []

    print(f"❌ Error al obtener tareas: {response.status_code} - {response.text}")
    return []

# 🔹 Función para obtener el Resumen del Estado de la tarea
def get_resume_task(taskId):
    params = {"taskId": taskId}
    headers = {"Content-Type": "application/json"}

    response = requests.post(BITRIX_URLS["GET_RESUME"], json=params, headers=headers, verify=False)

    if response.status_code == 200:
        data = response.json()

        if "result" in data and isinstance(data["result"], list) and len(data["result"]) > 0:
            return data["result"][0].get("text", "Sin comentarios disponibles").strip()

    print(f"❌ Error al obtener resumen de la tarea {taskId}: {response.status_code} - {response.text}")
    return "Sin resúmen"

# 🔹 Función para actualizar Google Sheets con los resúmenes de estado
def update_sheet_with_summaries(sheet_id, tasks):
    if not tasks:
        print("⚠️ No hay tareas para actualizar en Google Sheets.")
        return

    try:
        values = []
        for task in tasks:
            task_id = task.get("id")
            resumen = get_resume_task(task_id)
            print(f"📌 Tarea {task_id}: {resumen}")  # Para depuración
            values.append([resumen])

        body = {"values": values}

        # Actualiza Google Sheets
        response = service.spreadsheets().values().update(
            spreadsheetId=sheet_id,
            range=WRITE_RANGE_NAME,
            valueInputOption="USER_ENTERED",
            body=body
        ).execute()

        print(f"✅ Se han actualizado {response.get('updatedCells', 0)} celdas en Google Sheets.")
    except Exception as e:
        print(f"❌ Error al actualizar Google Sheets: {e}")

# 🔹 Ejecutar el script
if __name__ == "__main__":
    print("📥 Obteniendo resúmenes de tareas desde Bitrix...")
    tasks = get_tasks_from_bitrix()

    if tasks:
        update_sheet_with_summaries(SHEET_ID, tasks)
    else:
        print("⚠️ No se encontraron tareas en Bitrix.")
