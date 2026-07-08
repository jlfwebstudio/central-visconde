from datetime import datetime
from typing import Optional

from fastapi import Depends, HTTPException, status
from fastapi.security import HTTPAuthorizationCredentials, HTTPBearer
from sqlalchemy.orm import Session

from .database import obter_sessao
from .models import Conta, Sessao

_bearer = HTTPBearer(auto_error=False)


def obter_conta_atual(
    credenciais: Optional[HTTPAuthorizationCredentials] = Depends(_bearer),
    sessao_db: Session = Depends(obter_sessao),
) -> Conta:
    if credenciais is None:
        raise HTTPException(status.HTTP_401_UNAUTHORIZED, "Token de sessão ausente.")

    sessao = (
        sessao_db.query(Sessao)
        .filter(Sessao.token == credenciais.credentials)
        .one_or_none()
    )

    if sessao is None:
        raise HTTPException(status.HTTP_401_UNAUTHORIZED, "Sessão inválida.")

    if sessao.expira_em < datetime.utcnow():
        raise HTTPException(status.HTTP_401_UNAUTHORIZED, "Sessão expirada.")

    return sessao.conta


def exigir_admin_master(conta: Conta = Depends(obter_conta_atual)) -> Conta:
    if not conta.is_admin_master:
        raise HTTPException(status.HTTP_403_FORBIDDEN, "Acesso restrito ao admin master.")
    return conta
