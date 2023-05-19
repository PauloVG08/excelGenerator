from fastapi import FastAPI
from pydantic import BaseModel
from typing import Optional


class Productos(BaseModel):
    productos : Optional[str]
    precios : Optional[float]
    codigo_de_producto : Optional[str]
    estatus : Optional[int]

app = FastAPI()

@app.get("/")
def index():
    return {"mensaje": "Hola, mundo"}

@app.get("/productos/{id}")
def mostrarProducto(id: int):
    return {"idProducto": id}

@app.post("/productos")
def insertarProducto(producto: Productos):
    return producto.dict()