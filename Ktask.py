import json
import re
import subprocess
import urllib3
from services.google_sheets import read_from_sheet, write_tasks_to_sheet
from services.bitrix import get_user_id_by_name, create_task_in_bitrix, get_tasks_from_bitrix, get_group_id_by_name
from scripts.ResumeTask import get_resume_task
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

# 🔹 Función para cargar JSON de manera segura
def load_json(filename):
    try:
        with open(filename, 'r', encoding='utf-8') as file:
            return json.load(file)
    except (FileNotFoundError, json.JSONDecodeError):
        print(f"Error al cargar {filename}")
        exit(1)

# 🔹 Cargar configuraciones desde archivos JSON
sheet_config = load_json("settings/SheetURL.json")
settings_config = load_json("settings/settings.json")
user_config = load_json("user_config/user_id.json")

# 🔹 Extraer SHEET_ID desde la URL
def extract_sheet_id(sheet_url):
    match = re.search(r"/d/([a-zA-Z0-9-_]+)/", sheet_url)
    return match.group(1) if match else None

SHEET_ID = extract_sheet_id(sheet_config["SHEET_URL"])
if not SHEET_ID:
    print("Error: No se pudo extraer el ID de la hoja de cálculo.")
    exit(1)

# 🔹 Obtener USER_ID
USER_ID = user_config.get("USER_ID")
if not USER_ID:
    print("Error: No se encontró 'USER_ID' en user_id.json")
    exit(1)

# 🔹 Construcción de URLs de Bitrix con USER_ID dinámico
BITRIX_BASE_URL = settings_config["BITRIX_BASE_URL"].replace("{USER_ID}", USER_ID)
BITRIX_URLS = {key: f"{BITRIX_BASE_URL}{endpoint}" for key, endpoint in settings_config["ENDPOINTS"].items()}

# Ruta al archivo JSON de tu cuenta de servicio
SERVICE_ACCOUNT_FILE = "fit-guide-433118-p4-193f4862b36c.json"

# Alcances (scopes) necesarios
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

# Autenticación
credentials = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)

# Construir el servicio de Google Sheets
service = build("sheets", "v4", credentials=credentials)

# Desactivar advertencias de SSL
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# Rango de celdas para lectura y escritura
READ_RANGE_NAME = "Hoja 2!A2:R"
WRITE_RANGE_NAME = "Hoja 2!A2"

headers = {"Content-Type": "application/json", "Accept": "application/json"}

# Cache para almacenar usuarios y grupos para poder obtener su id y nombre
user_cache = {}
group_cache = {}


# 🔹 Función principal del menú
def main():
    while True:
        print("\n📌 Menú de Opciones:")
        print("1. Obtener tareas")
        print("2. Crear tareas")
        print("3. Obtener Resumen del Estado de las Tareas")
        print("4. Salir de Ktask")

        opcion = input("Ingrese el número de la opción: ")

        if opcion == "1":
            print("\n📥 Obteniendo tareas desde Bitrix...")
            tasks = get_tasks_from_bitrix()
            if tasks:
                write_tasks_to_sheet(SHEET_ID, tasks)
            else:
                print("⚠️ No se encontraron tareas en Bitrix.")

        elif opcion == "2":
            print("\n📤 Creando tareas en Bitrix...")
            data = read_from_sheet(SHEET_ID, READ_RANGE_NAME)
            if data:
                process_sheet_data(data)
            else:
                print("⚠️ No se encontraron datos en la hoja.")

        elif opcion == "3":
            print("\n📋 Ejecutando el script scripts/ResumeTask.py...")

            result = subprocess.run(["python3", "scripts/ResumeTask.py"], capture_output=True, text=True)

            print(result.stdout)  # Mostrar salida del script
            print(result.stderr)  # Mostrar errores (si hay)

        elif opcion == "4":
            print("✅ Saliendo del programa...")
            break

        else:
            print("❌ Opción no válida, intenta de nuevo.")

# 🔹 Función para procesar las tareas y enviarlas a Bitrix
def process_sheet_data(data):
    for row in data:
        if len(row) < 13:
            print(f"⚠️ Fila incompleta, se omite: {row}")
            continue

        if row[12].strip():
            print(f"⚠️ La celda 'M' contiene datos, se omite la fila: {row}")
            continue

        task_name = f"{row[2].strip()} {row[3].strip()}"
        task_description = row[4].strip() if len(row) > 4 else ""

        responsible_name = row[6].strip() if len(row) > 6 else ""
        creator_name = row[7].strip() if len(row) > 7 else ""

        responsible_id = get_user_id_by_name(responsible_name)
        creator_id = get_user_id_by_name(creator_name)

        if not responsible_id or not creator_id:
            print(f"⚠️ Error con los IDs de Responsable o Creador: {responsible_name}, {creator_name}")
            continue

        participants = []
        if len(row) > 8 and row[8].strip():
            participants = [
                get_user_id_by_name(name.strip())
                for name in row[8].strip().split(",")
                if name.strip()
            ]
            participants = [id for id in participants if id]

        observers = []
        if len(row) > 9 and row[9].strip():
            observers = [
                get_user_id_by_name(name.strip())
                for name in row[9].strip().split(",")
                if name.strip()
            ]
            observers = [id for id in observers if id]

        tags = []
        if len(row) > 10 and row[10].strip():
            tags = [tag.strip() for tag in row[10].strip().split(",") if tag.strip()]

        group_name = row[11].strip() if len(row) > 11 else ""
        group_id = get_group_id_by_name(group_name)

        task_data = {
            "TITLE": task_name,
            "DESCRIPTION": task_description,
            "RESPONSIBLE_ID": responsible_id,
            "CREATED_BY": creator_id,
            "ACCOMPLICES": participants,
            "AUDITORS": observers,
            "TAGS": tags,
            "GROUP_ID": group_id,
        }

        print(f"📌 Tarea preparada: {task_data}")
        create_task_in_bitrix(task_data)


# 🔹 Iniciar el menú primero
if __name__ == "__main__":
    main()
