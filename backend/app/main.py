from fastapi import FastAPI

from . import models  # noqa: F401 garante que os modelos são registrados antes do create_all
from .database import Base, engine
from .routers import admin, auth, conta, versao

Base.metadata.create_all(bind=engine)

app = FastAPI(title="ViscondeApp - Contas")

app.include_router(auth.router)
app.include_router(conta.router)
app.include_router(admin.router)
app.include_router(versao.router)


@app.get("/saude")
def saude():
    return {"status": "ok"}
