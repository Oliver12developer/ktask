import json
import re
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


#Librería para desactivar la autenticación de certificado
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)


# Cargar configuraciones desde settings.json
with open("settings/settings.json", "r") as file:
    config = json.load(file)

#Cargar URL desde SheetURL.json
with open("settings/SheetURL.json", "r") as file:
    urlSheet = json.load(file)

#Rango de celdas para lectura y escritura
READ_RANGE_NAME = "Hoja 2!A2:R"
WRITE_RANGE_NAME = "Hoja 2!A2"


headers = {"Content-Type": "application/json", "Accept": "application/json"}

#Cache para almacenar usuarios y grupos para poder obtener su id y nombre
user_cache = {}
group_cache = {}


def process_sheet_data(data):
    for row in data:
        if len(row) < 13:  # Validar que la fila tenga al menos 13 columnas
            print(f"Fila incompleta, se omite: {row}")
            continue

        # Validar que la celda `M` (índice 12) esté vacía
        if row[12].strip():  # Si contiene datos, omitir la fila
            print(f"La celda 'M' contiene datos, se omite la fila: {row}")
            continue

        # Obtener y preparar datos para crear la tarea
        task_name = f"{row[2].strip()} {row[3].strip()}"  # Concatenar C y D
        task_description = row[4].strip() if len(row) > 4 else ""  # Celda E

        # Convertir nombres en IDs con `get_user_id_by_name`
        responsible_name = row[6].strip() if len(row) > 6 else ""  # Columna G
        creator_name = row[7].strip() if len(row) > 7 else ""  # Columna H

        responsible_id = get_user_id_by_name(responsible_name)
        creator_id = get_user_id_by_name(creator_name)

        # Validar que los IDs sean válidos
        if not responsible_id:
            print(f"Error: No se encontró el ID del Responsable '{responsible_name}'")
            continue

        if not creator_id:
            print(f"Error: No se encontró el ID del Creador '{creator_name}'")
            continue

        # Obtener IDs para participantes y observadores
        participants = []
        if len(row) > 8 and row[8].strip():  # Columna I
            participants = [
                get_user_id_by_name(name.strip())
                for name in row[8].strip().split(",")
                if name.strip()
            ]
            # Filtrar IDs inválidos
            participants = [id for id in participants if id]

        observers = []
        if len(row) > 9 and row[9].strip():  # Columna J
            observers = [
                get_user_id_by_name(name.strip())
                for name in row[9].strip().split(",")
                if name.strip()
            ]
            # Filtrar IDs inválidos
            observers = [id for id in observers if id]

        # Obtener etiquetas (tags)
        tags = []
        if len(row) > 10 and row[10].strip():  # Columna K
            tags = [tag.strip() for tag in row[10].strip().split(",") if tag.strip()]

        # Obtener ID del grupo
        group_name = row[11].strip() if len(row) > 11 else ""  # Columna L
        group_id = get_group_id_by_name(group_name)

        # Crear la tarea en Bitrix
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

        # Verificar datos enviados
        print(f"Tarea preparada: {task_data}")

        create_task_in_bitrix(task_data)

def main():
    while True:
        print("\nSeleccione una opción:")
        print("1. Iniciar programa (solo si es la primera vez que ejecutas Ktask)")
        print("2. Obtener tareas")
        print("3. Crear tareas")
        print("4. Resumen del estado de tarea")
        print("5. Salir")

        opcion = input("Ingrese el número de la opción: ")

        if opcion == "1":
            print("Ejecutando la configuración inicial de Ktask...")
            process_sheet_data()  # Se puede modificar según la inicialización real
        elif opcion == "2":
            process_sheet_data()
        elif opcion == "3":
            create_task_in_bitrix()
        elif opcion == "4":
            get_resume_task()
        elif opcion == "5":
            print("Saliendo del programa...")
            break
        else:
            print("Opción no válida, intenta de nuevo.")


if __name__ == "__main__":
    # Prueba de la función get_user_id_by_name
    user_name = "Oliver Suárez"  # Cambia esto por el nombre real
    user_id = get_user_id_by_name(user_name)
    if user_id:
        print(f"El ID de '{user_name}' es {user_id}")
    else:
        print(f"No se encontró el usuario con el nombre '{user_name}'")

    data = read_from_sheet(SHEET_ID, READ_RANGE_NAME)
    if data:
        for row in data:
            if len(row) < 13:  # Validar que la fila tenga al menos 13 columnas
                print(f"Fila incompleta, se omite: {row}")
                continue

            # Validar que la celda `M` (índice 12) contenga datos
            if row[12].strip():  # Celda M
                print(f"La celda 'M' contiene datos, se omite la fila: {row}")
                continue

            # Obtener y preparar datos para crear la tarea
            task_name = f"{row[2].strip()} {row[3].strip()}"  # Concatenar C y D
            task_description = row[4].strip() if len(row) > 4 else ""  # Celda E

            # Convertir nombres en IDs con `get_user_id_by_name`
            responsible_name = row[6].strip() if len(row) > 6 else ""  # Columna G
            creator_name = row[7].strip() if len(row) > 7 else ""  # Columna H

            responsible_id = get_user_id_by_name(responsible_name)
            creator_id = get_user_id_by_name(creator_name)

            # Validar que los IDs sean válidos
            if not responsible_id:
                print(f"Error: No se encontró el ID del Responsable '{responsible_name}'")
                continue

            if not creator_id:
                print(f"Error: No se encontró el ID del Creador '{creator_name}'")
                continue

            # Obtener IDs para participantes y observadores
            participants = []
            if len(row) > 8 and row[8].strip():  # Columna I
                participants = [
                    get_user_id_by_name(name.strip())
                    for name in row[8].strip().split(",")
                    if name.strip()
                ]
                # Filtrar IDs inválidos
                participants = [id for id in participants if id]

            observers = []
            if len(row) > 9 and row[9].strip():  # Columna J
                observers = [
                    get_user_id_by_name(name.strip())
                    for name in row[9].strip().split(",")
                    if name.strip()
                ]
                # Filtrar IDs inválidos
                observers = [id for id in observers if id]

            # Obtener etiquetas (tags)
            tags = []
            if len(row) > 10 and row[10].strip():  # Columna K
                tags = [tag.strip() for tag in row[10].strip().split(",") if tag.strip()]

            # Obtener ID del grupo
            group_name = row[11].strip() if len(row) > 11 else ""  # Columna L
            group_id = get_group_id_by_name(group_name)

            # Crear la tarea en Bitrix
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

            # Verificar datos enviados
            print(f"Tarea preparada: {task_data}")

            create_task_in_bitrix(task_data)
    else:
        print("No se encontraron datos en la hoja.")

    # Obtener tareas de Bitrix y escribirlas en la hoja de cálculo
    tasks = get_tasks_from_bitrix()
    if tasks:
        write_tasks_to_sheet(SHEET_ID, tasks)
    else:
        print("No se encontraron tareas.")
