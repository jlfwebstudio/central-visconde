from datetime import datetime

from pydantic import BaseModel, Field


class ContaCreate(BaseModel):
    usuario: str = Field(min_length=3, max_length=80)
    senha: str = Field(min_length=6, max_length=255)


class ContaLogin(BaseModel):
    usuario: str
    senha: str


class TokenResponse(BaseModel):
    token: str
    expira_em: datetime


class CredencialIn(BaseModel):
    url: str = ""
    usuario: str = ""
    senha: str = ""
    ativo: bool = True


class CredencialOut(BaseModel):
    url: str
    usuario: str
    senha: str
    ativo: bool


class ContaConfigIn(BaseModel):
    plataformas: dict[str, CredencialIn] = {}
    config: dict = {}


class ContaConfigOut(BaseModel):
    usuario: str
    is_admin_master: bool
    plataformas: dict[str, CredencialOut]
    config: dict


class AdminContaResumo(BaseModel):
    id: int
    usuario: str
    criado_em: datetime
    is_admin_master: bool
    plataformas_ativas: list[str]


class VersaoIn(BaseModel):
    plataforma: str = Field(pattern="^(mac|windows)$")
    versao: str = Field(min_length=1, max_length=20)
    url_download: str = Field(min_length=1, max_length=500)
    obrigatoria: bool = False
    notas: str = ""


class VersaoOut(BaseModel):
    plataforma: str
    versao: str
    url_download: str
    obrigatoria: bool
    notas: str
    publicada_em: datetime
