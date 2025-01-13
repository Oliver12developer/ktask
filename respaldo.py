import requests
import json
import urllib3
from services.google_sheets import read_from_sheet, write_tasks_to_sheet
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

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

# Acceder a las configuraciones
BITRIX_CREATE_TASK_URL = config["BITRIX_CREATE_TASK_URL"]
BITRIX_GET_TASKS_URL = config["BITRIX_GET_TASKS_URL"]
GET_COMMENTS_LIST_URL = config["GET_COMMENTS_LIST_URL"]
GET_COMMENT_DETAILS_URL = config["GET_COMMENT_DETAILS_URL"]
BITRIX_USER_INFO_URL = config["BITRIX_USER_INFO_URL"]
BITRIX_GROUP_INFO_URL = config["BITRIX_GROUP_INFO_URL"]
SHEET_ID = urlSheet["SHEET_ID"]


READ_RANGE_NAME = "Hoja 2!A2:R"
WRITE_RANGE_NAME = "Hoja 2!A2"


headers = {"Content-Type": "application/json", "Accept": "application/json"}


user_cache = {}
group_cache = {}


def get_credentials():
    SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
    SERVICE_ACCOUNT_FILE = "fit-guide-433118-p4-193f4862b36c.json"  # Ruta a tu archivo JSON
    return Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)


def read_from_sheet(sheet_id, range_name):
    creds = get_credentials()
    if not creds:
        print("Credenciales no válidas.")
        return None

    service = build("sheets", "v4", credentials=creds)
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=sheet_id, range=range_name).execute()
    values = result.get("values", [])
    return values if values else None


def create_task_in_bitrix(task_data):
    required_fields = ["TITLE", "RESPONSIBLE_ID", "CREATED_BY"]

    # Verifica si faltan campos obligatorios
    for field in required_fields:
        if not task_data.get(field):
            print(f"Falta el campo obligatorio '{field}' en la tarea: {task_data}")
            return

    # Log detallado: Mostrar los datos que se enviarán
    print(f"Datos enviados a Bitrix: {task_data}")

    # Extraer los IDs para validación adicional y mostrarlos
    responsible_id = task_data.get("RESPONSIBLE_ID")
    creator_id = task_data.get("CREATED_BY")
    print(f"RESPONSIBLE_ID: {responsible_id}, CREATED_BY: {creator_id}")

    # Realizar la solicitud al API
    response = requests.post(
        BITRIX_CREATE_TASK_URL,
        json={"fields": task_data},
        headers=headers,
    )
    if response.status_code == 200:
        print("Tarea creada exitosamente:", response.json())
    else:
        print(f"Error al crear la tarea: {response.status_code} - {response.text}")


def get_user_id_by_name(full_name):
    if full_name in user_cache:
        return user_cache[full_name]

    name_parts = full_name.split(" ", 1)
    first_name = name_parts[0]
    last_name = name_parts[1] if len(name_parts) > 1 else ""

    response = requests.post(
        BITRIX_USER_INFO_URL,
        json={"FILTER": {"NAME": first_name, "LAST_NAME": last_name}},
        verify=False,
    )

    if response.status_code == 200:
        user_info = response.json().get("result", [])
        if user_info:
            user_id = user_info[0].get("ID")
            user_cache[full_name] = user_id
            return user_id
    else:
        print(
            f"Error al obtener el ID del usuario '{full_name}': {response.status_code} - {response.text}"
        )

    return None


def get_group_id_by_name(group_name):
    if group_name in group_cache:
        return group_cache[group_name]

    response = requests.post(
        BITRIX_GROUP_INFO_URL,
        json={"FILTER": {"NAME": group_name}},  # Cambia "NAME" si es necesario
        verify=False,
    )

    print(
        f"Respuesta completa para el grupo '{group_name}': {response.json()}"
    )  # Agregado

    if response.status_code == 200:
        group_info = response.json().get("result", [])
        if group_info:
            group_id = group_info[0].get("ID")
            group_cache[group_name] = group_id
            return group_id
    else:
        print(
            f"Error al obtener el ID del grupo '{group_name}': {response.status_code} - {response.text}"
        )

    return None


def get_tasks_from_bitrix():
    response = requests.post(BITRIX_GET_TASKS_URL, verify=False)
    if response.status_code == 200:
        tasks = response.json().get("result", {}).get("tasks", [])
        return tasks
    else:
        print(f"Error al obtener tareas: {response.status_code} - {response.text}")
        return []


def get_last_comment(task_id):
    params = [task_id, {"POST_DATE": "desc"}]
    response = requests.post(GET_COMMENTS_LIST_URL, json=params, verify=False)
    if response.status_code == 200:
        comments = response.json().get("result", [])
        return comments[0].get("ID") if comments else None
    else:
        print(
            f"Error al obtener comentarios de la tarea {task_id}: {response.status_code} - {response.text}"
        )
        return None


def get_comment_details(task_id, comment_id):
    params = [task_id, comment_id]
    response = requests.post(GET_COMMENT_DETAILS_URL, json=params, verify=False)
    if response.status_code == 200:
        comment = response.json().get("result", {})
        return comment.get("POST_MESSAGE", "Sin contenido")
    else:
        print(
            f"Error al obtener detalles del comentario {comment_id}: {response.status_code} - {response.text}"
        )
        return "Sin contenido"


def get_user_name(user_id):
    if user_id in user_cache:
        return user_cache[user_id]

    response = requests.post(BITRIX_USER_INFO_URL, json={"ID": user_id}, verify=False)
    if response.status_code == 200:
        user_info = response.json().get("result", [])
        if user_info:
            user = user_info[0]
            name = f"{user.get('NAME', 'Sin nombre')} {user.get('LAST_NAME', 'Sin apellido')}"
            user_cache[user_id] = name
            return name
    else:
        print(
            f"Error al obtener información del usuario {user_id}: {response.status_code} - {response.text}"
        )
    return "Sin información"


def get_group_name(group_id):
    if not group_id:
        return "Sin grupo"

    if group_id in group_cache:
        return group_cache[group_id]

    response = requests.post(
        BITRIX_GROUP_INFO_URL, json={"FILTER": {"ID": group_id}}, verify=False
    )
    if response.status_code == 200:
        group_info = response.json().get("result", [])
        if group_info:
            group_name = group_info[0].get("NAME", "Sin nombre")
            group_cache[group_id] = group_name
            return group_name
    else:
        print(
            f"Error al obtener información del grupo {group_id}: {response.status_code} - {response.text}"
        )
    return "Sin información"


def write_tasks_to_sheet(sheet_id, tasks):
    creds = get_credentials()
    if not creds:
        print("Credenciales no válidas.")
        return

    service = build("sheets", "v4", credentials=creds)
    sheet = service.spreadsheets()

    # Leer los datos existentes en la hoja
    result = sheet.values().get(spreadsheetId=sheet_id, range=READ_RANGE_NAME).execute()
    existing_values = result.get("values", [])

    # Crear un diccionario para facilitar la búsqueda de nombres existentes
    existing_tasks = {
        f"{row[2]} {row[3]}": index
        for index, row in enumerate(existing_values)
        if len(row) >= 4  # Asegurar que haya al menos 4 columnas (C y D)
    }

    # Valores para actualizar la hoja
    values = existing_values  # Partimos de los valores existentes
    for task in tasks:
        task_id = task["id"]
        task_title = task["title"].split(" ", 1)
        task_name = (
            f"{task_title[0]} {task_title[1]}"
            if len(task_title) > 1
            else f"{task_title[0]} Sin nombre"
        )
        task_link = f'=HYPERLINK("https://bitrix.kernotek.mx/company/personal/user/410/tasks/task/view/{task_id}/", "{task_id}")'
        last_comment_id = get_last_comment(task_id)
        last_comment_message = (
            get_comment_details(task_id, last_comment_id)
            if last_comment_id
            else "Sin comentarios"
        )

        # Validar si la tarea ya existe
        if task_name in existing_tasks:
            # Actualizar solo las columnas necesarias
            row_index = existing_tasks[task_name]
            while len(values[row_index]) < 17:
                values[row_index].append("")  # Asegurar que la fila tenga al menos 17 columnas
            values[row_index][12] = task_link  # Columna M
            values[row_index][16] = last_comment_message  # Columna Q
        else:
            # Agregar nueva fila con todos los datos
            new_row = [""] * 17  # Inicializar fila con 17 columnas vacías
            new_row[2] = task_title[0]  # Columna C
            new_row[3] = task_title[1] if len(task_title) > 1 else "Sin nombre"  # Columna D
            new_row[4] = task.get("description", "")  # Columna E
            new_row[6] = get_user_name(task.get("responsibleId", ""))  # Columna G
            new_row[7] = get_user_name(task.get("createdBy", ""))  # Columna H
            new_row[8] = ", ".join(
                [get_user_name(user_id) for user_id in task.get("accomplices", [])]
            )  # Columna I
            new_row[9] = ", ".join(
                [get_user_name(user_id) for user_id in task.get("auditors", [])]
            )  # Columna J
            new_row[11] = get_group_name(task.get("groupId", ""))  # Columna L
            new_row[12] = task_link  # Columna M
            new_row[16] = last_comment_message  # Columna Q

            values.append(new_row)

    # Actualizar los datos en la hoja
    body = {"values": values}
    result = (
        sheet.values()
        .update(
            spreadsheetId=sheet_id,
            range=READ_RANGE_NAME,
            valueInputOption="USER_ENTERED",
            body=body,
        )
        .execute()
    )
    print(f"Se han actualizado {result.get('updatedCells')} celdas.")
import requests
import json
import urllib3
from services.google_sheets import read_from_sheet, write_tasks_to_sheet
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

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

# Acceder a las configuraciones
BITRIX_CREATE_TASK_URL = config["BITRIX_CREATE_TASK_URL"]
BITRIX_GET_TASKS_URL = config["BITRIX_GET_TASKS_URL"]
GET_COMMENTS_LIST_URL = config["GET_COMMENTS_LIST_URL"]
GET_COMMENT_DETAILS_URL = config["GET_COMMENT_DETAILS_URL"]
BITRIX_USER_INFO_URL = config["BITRIX_USER_INFO_URL"]
BITRIX_GROUP_INFO_URL = config["BITRIX_GROUP_INFO_URL"]
SHEET_ID = urlSheet["SHEET_ID"]


READ_RANGE_NAME = "Hoja 2!A2:R"
WRITE_RANGE_NAME = "Hoja 2!A2"


headers = {"Content-Type": "application/json", "Accept": "application/json"}


user_cache = {}
group_cache = {}


def get_credentials():
    SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
    SERVICE_ACCOUNT_FILE = "fit-guide-433118-p4-193f4862b36c.json"  # Ruta a tu archivo JSON
    return Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)


def read_from_sheet(sheet_id, range_name):
    creds = get_credentials()
    if not creds:
        print("Credenciales no válidas.")
        return None

    service = build("sheets", "v4", credentials=creds)
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=sheet_id, range=range_name).execute()
    values = result.get("values", [])
    return values if values else None


def create_task_in_bitrix(task_data):
    required_fields = ["TITLE", "RESPONSIBLE_ID", "CREATED_BY"]

    # Verifica si faltan campos obligatorios
    for field in required_fields:
        if not task_data.get(field):
            print(f"Falta el campo obligatorio '{field}' en la tarea: {task_data}")
            return

    # Log detallado: Mostrar los datos que se enviarán
    print(f"Datos enviados a Bitrix: {task_data}")

    # Extraer los IDs para validación adicional y mostrarlos
    responsible_id = task_data.get("RESPONSIBLE_ID")
    creator_id = task_data.get("CREATED_BY")
    print(f"RESPONSIBLE_ID: {responsible_id}, CREATED_BY: {creator_id}")

    # Realizar la solicitud al API
    response = requests.post(
        BITRIX_CREATE_TASK_URL,
        json={"fields": task_data},
        headers=headers,
    )
    if response.status_code == 200:
        print("Tarea creada exitosamente:", response.json())
    else:
        print(f"Error al crear la tarea: {response.status_code} - {response.text}")


def get_user_id_by_name(full_name):
    if full_name in user_cache:
        return user_cache[full_name]

    name_parts = full_name.split(" ", 1)
    first_name = name_parts[0]
    last_name = name_parts[1] if len(name_parts) > 1 else ""

    response = requests.post(
        BITRIX_USER_INFO_URL,
        json={"FILTER": {"NAME": first_name, "LAST_NAME": last_name}},
        verify=False,
    )

    if response.status_code == 200:
        user_info = response.json().get("result", [])
        if user_info:
            user_id = user_info[0].get("ID")
            user_cache[full_name] = user_id
            return user_id
    else:
        print(
            f"Error al obtener el ID del usuario '{full_name}': {response.status_code} - {response.text}"
        )

    return None


def get_group_id_by_name(group_name):
    if group_name in group_cache:
        return group_cache[group_name]

    response = requests.post(
        BITRIX_GROUP_INFO_URL,
        json={"FILTER": {"NAME": group_name}},  # Cambia "NAME" si es necesario
        verify=False,
    )

    print(
        f"Respuesta completa para el grupo '{group_name}': {response.json()}"
    )  # Agregado

    if response.status_code == 200:
        group_info = response.json().get("result", [])
        if group_info:
            group_id = group_info[0].get("ID")
            group_cache[group_name] = group_id
            return group_id
    else:
        print(
            f"Error al obtener el ID del grupo '{group_name}': {response.status_code} - {response.text}"
        )

    return None


def get_tasks_from_bitrix():
    response = requests.post(BITRIX_GET_TASKS_URL, verify=False)
    if response.status_code == 200:
        tasks = response.json().get("result", {}).get("tasks", [])
        return tasks
    else:
        print(f"Error al obtener tareas: {response.status_code} - {response.text}")
        return []


def get_last_comment(task_id):
    params = [task_id, {"POST_DATE": "desc"}]
    response = requests.post(GET_COMMENTS_LIST_URL, json=params, verify=False)
    if response.status_code == 200:
        comments = response.json().get("result", [])
        return comments[0].get("ID") if comments else None
    else:
        print(
            f"Error al obtener comentarios de la tarea {task_id}: {response.status_code} - {response.text}"
        )
        return None


def get_comment_details(task_id, comment_id):
    params = [task_id, comment_id]
    response = requests.post(GET_COMMENT_DETAILS_URL, json=params, verify=False)
    if response.status_code == 200:
        comment = response.json().get("result", {})
        return comment.get("POST_MESSAGE", "Sin contenido")
    else:
        print(
            f"Error al obtener detalles del comentario {comment_id}: {response.status_code} - {response.text}"
        )
        return "Sin contenido"


def get_user_name(user_id):
    if user_id in user_cache:
        return user_cache[user_id]

    response = requests.post(BITRIX_USER_INFO_URL, json={"ID": user_id}, verify=False)
    if response.status_code == 200:
        user_info = response.json().get("result", [])
        if user_info:
            user = user_info[0]
            name = f"{user.get('NAME', 'Sin nombre')} {user.get('LAST_NAME', 'Sin apellido')}"
            user_cache[user_id] = name
            return name
    else:
        print(
            f"Error al obtener información del usuario {user_id}: {response.status_code} - {response.text}"
        )
    return "Sin información"


def get_group_name(group_id):
    if not group_id:
        return "Sin grupo"

    if group_id in group_cache:
        return group_cache[group_id]

    response = requests.post(
        BITRIX_GROUP_INFO_URL, json={"FILTER": {"ID": group_id}}, verify=False
    )
    if response.status_code == 200:
        group_info = response.json().get("result", [])
        if group_info:
            group_name = group_info[0].get("NAME", "Sin nombre")
            group_cache[group_id] = group_name
            return group_name
    else:
        print(
            f"Error al obtener información del grupo {group_id}: {response.status_code} - {response.text}"
        )
    return "Sin información"


def write_tasks_to_sheet(sheet_id, tasks):
    creds = get_credentials()
    if not creds:
        print("Credenciales no válidas.")
        return

    service = build("sheets", "v4", credentials=creds)
    sheet = service.spreadsheets()

    # Leer los datos existentes en la hoja
    result = sheet.values().get(spreadsheetId=sheet_id, range=READ_RANGE_NAME).execute()
    existing_values = result.get("values", [])

    # Crear un diccionario para facilitar la búsqueda de nombres existentes
    existing_tasks = {
        f"{row[2]} {row[3]}": index
        for index, row in enumerate(existing_values)
        if len(row) >= 4  # Asegurar que haya al menos 4 columnas (C y D)
    }

    # Valores para actualizar la hoja
    values = existing_values  # Partimos de los valores existentes
    for task in tasks:
        task_id = task["id"]
        task_title = task["title"].split(" ", 1)
        task_name = (
            f"{task_title[0]} {task_title[1]}"
            if len(task_title) > 1
            else f"{task_title[0]} Sin nombre"
        )
        task_link = f'=HYPERLINK("https://bitrix.kernotek.mx/company/personal/user/410/tasks/task/view/{task_id}/", "{task_id}")'
        last_comment_id = get_last_comment(task_id)
        last_comment_message = (
            get_comment_details(task_id, last_comment_id)
            if last_comment_id
            else "Sin comentarios"
        )

        # Validar si la tarea ya existe
        if task_name in existing_tasks:
            # Actualizar solo las columnas necesarias
            row_index = existing_tasks[task_name]
            while len(values[row_index]) < 17:
                values[row_index].append("")  # Asegurar que la fila tenga al menos 17 columnas
            values[row_index][12] = task_link  # Columna M
            values[row_index][16] = last_comment_message  # Columna Q
        else:
            # Agregar nueva fila con todos los datos
            new_row = [""] * 17  # Inicializar fila con 17 columnas vacías
            new_row[2] = task_title[0]  # Columna C
            new_row[3] = task_title[1] if len(task_title) > 1 else "Sin nombre"  # Columna D
            new_row[4] = task.get("description", "")  # Columna E
            new_row[6] = get_user_name(task.get("responsibleId", ""))  # Columna G
            new_row[7] = get_user_name(task.get("createdBy", ""))  # Columna H
            new_row[8] = ", ".join(
                [get_user_name(user_id) for user_id in task.get("accomplices", [])]
            )  # Columna I
            new_row[9] = ", ".join(
                [get_user_name(user_id) for user_id in task.get("auditors", [])]
            )  # Columna J
            new_row[11] = get_group_name(task.get("groupId", ""))  # Columna L
            new_row[12] = task_link  # Columna M
            new_row[16] = last_comment_message  # Columna Q

            values.append(new_row)

    # Actualizar los datos en la hoja
    body = {"values": values}
    result = (
        sheet.values()
        .update(
            spreadsheetId=sheet_id,
            range=READ_RANGE_NAME,
            valueInputOption="USER_ENTERED",
            body=body,
        )
        .execute()
    )
    print(f"Se han actualizado {result.get('updatedCells')} celdas.")


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
