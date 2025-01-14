import json
import re
import requests

# Cargar configuraciones desde settings.json
with open("settings/settings.json", "r") as file:
    config = json.load(file)

# 🔹 Cargar configuraciones
settings = json.load(open("settings/settings.json"))
user_config = json.load(open("user_config/user_id.json"))

# 🔹 Construcción de URLs dinámicas
BITRIX_BASE_URL = settings["BITRIX_BASE_URL"].replace("{USER_ID}", user_config["USER_ID"])
BITRIX_URLS = {key: f"{BITRIX_BASE_URL}{endpoint}" for key, endpoint in settings["ENDPOINTS"].items()}

headers = {"Content-Type": "application/json", "Accept": "application/json"}


user_cache = {}
group_cache = {}


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
        BITRIX_URLS["BITRIX_CREATE_TASK_URL"],
        json={"fields": task_data},
        headers=headers,
    )
    if response.status_code == 200:
        print("Tarea creada exitosamente:", response.json())
    else:
        print(f"Error al crear la tarea: {response.status_code} - {response.text}")


# 🔹 Función para obtener tareas desde Bitrix con filtro por nombre
#def get_tasks_from_bitrix():
#    response = requests.post(BITRIX_URLS["BITRIX_GET_TASKS_URL"], verify=False)
#
#    if response.status_code == 200:
#        data = response.json()
#        tasks = data.get("result", {}).get("tasks", [])
#
#        # 🔹 Filtrar solo las tareas cuyo nombre comienza con "KTK0"
#        filtered_tasks = [task for task in tasks if task.get("title", "").startswith("KTK0")]
#
#        return filtered_tasks if filtered_tasks else []
#
#    print(f"❌ Error al obtener tareas: {response.status_code} - {response.text}")
#    return []

# 🔹 Función para obtener tareas desde Bitrix
def get_tasks_from_bitrix():
    response = requests.post(BITRIX_URLS["BITRIX_GET_TASKS_URL"], verify=False)

    if response.status_code == 200:
        data = response.json()
        tasks = data.get("result", {}).get("tasks", [])
        return tasks if tasks else []

    print(f"❌ Error al obtener tareas: {response.status_code} - {response.text}")
    return []

def get_user_name(user_id):
    if user_id in user_cache:
        return user_cache[user_id]

    response = requests.post(BITRIX_URLS["BITRIX_USER_INFO"], json={"ID": user_id}, verify=False)
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
        BITRIX_URLS["BITRIX_GROUP_INFO"], json={"FILTER": {"ID": group_id}}, verify=False
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



def get_user_id_by_name(full_name):
    if full_name in user_cache:
        return user_cache[full_name]

    name_parts = full_name.split(" ", 1)
    first_name = name_parts[0]
    last_name = name_parts[1] if len(name_parts) > 1 else ""

    response = requests.post(
        BITRIX_URLS["BITRIX_USER_INFO"],
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
        BITRIX_URLS["BITRIX_GROUP_INFO"],
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


def get_last_comment(task_id):
    params = [task_id, {"POST_DATE": "desc"}]
    response = requests.post(BITRIX_URLS["GET_COMMENTS_LIST_URL"], json=params, verify=False)
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
    response = requests.post(BITRIX_URLS["GET_COMMENT_DETAILS_URL"], json=params, verify=False)
    if response.status_code == 200:
        comment = response.json().get("result", {})
        return comment.get("POST_MESSAGE", "Sin contenido")
    else:
        print(
            f"Error al obtener detalles del comentario {comment_id}: {response.status_code} - {response.text}"
        )
        return "Sin contenido"
