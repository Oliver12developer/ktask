import requests

BITRIX_USER_INFO_URL = "https://bitrix.kernotek.mx/rest/294/3kvtr1kiz9f8u18d/user.get"

payload = {"FILTER": {"NAME": "Rosalba"}}

response = requests.post(BITRIX_USER_INFO_URL, json=payload, verify=False)

print(response.json())  # Muestra todos los usuarios con nombre "Fernanda"
