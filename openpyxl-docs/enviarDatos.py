import httpx

data = {
    "productos": "Camisa",
    "precios": 145.75,
    "codigo_de_producto": "ABC123",
    "estatus": 1
}

# URL de tu API
url = "http://127.0.0.1:8000/productos"

# Realizar la solicitud POST
response = httpx.post(url, json=data)


# Verificar la respuesta
if response.status_code == 200:
    print("Datos enviados correctamente")
else:
    print("Error al enviar los datos")
    print(response.text)

