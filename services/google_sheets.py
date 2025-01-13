import json
import os
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from services.bitrix import get_user_name, get_group_name, get_user_id_by_name
from scripts.ResumeTask import get_task_summary

# Ruta al archivo JSON de tu cuenta de servicio
SERVICE_ACCOUNT_FILE = "fit-guide-433118-p4-193f4862b36c.json"

# Alcances (scopes) necesarios
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

READ_RANGE_NAME = "Hoja 2!A2:R"
WRITE_RANGE_NAME = "Hoja 2!A2"


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
        resume_task_status = get_task_summary(task_id)  # Obtener resumen de la tarea

        # Validar si la tarea ya existe
        if task_name in existing_tasks:
            # Actualizar solo las columnas necesarias
            row_index = existing_tasks[task_name]
            while len(values[row_index]) < 17:
                values[row_index].append("")  # Asegurar que la fila tenga al menos 17 columnas
            values[row_index][12] = task_link  # Columna M
            values[row_index][16] = resume_task_status  # Columna Q
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
            new_row[16] = resume_task_status  # Columna Q

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
