from datetime import datetime, timedelta

from fastapi import APIRouter, Depends, HTTPException, status
from sqlalchemy.orm import Session

from ..config import SESSION_TTL_DIAS
from ..database import obter_sessao
from ..models import Conta, ConfigAutomacao, Sessao
from ..schemas import ContaCreate, ContaLogin, TokenResponse
from ..security import gerar_token_sessao, hash_senha, verificar_senha

router = APIRouter(prefix="/auth", tags=["auth"])


def _criar_sessao(conta: Conta, banco: Session) -> TokenResponse:
    expira_em = datetime.utcnow() + timedelta(days=SESSION_TTL_DIAS)
    sessao = Sessao(conta_id=conta.id, token=gerar_token_sessao(), expira_em=expira_em)
    banco.add(sessao)
    banco.commit()
    return TokenResponse(token=sessao.token, expira_em=sessao.expira_em)


@router.post("/signup", response_model=TokenResponse)
def signup(dados: ContaCreate, banco: Session = Depends(obter_sessao)):
    ja_existe = banco.query(Conta).filter(Conta.usuario == dados.usuario).one_or_none()
    if ja_existe:
        raise HTTPException(status.HTTP_409_CONFLICT, "Já existe uma conta com esse usuário.")

    conta = Conta(usuario=dados.usuario, senha_hash=hash_senha(dados.senha))
    banco.add(conta)
    banco.flush()

    banco.add(ConfigAutomacao(conta_id=conta.id, config_json="{}"))
    banco.commit()
    banco.refresh(conta)

    return _criar_sessao(conta, banco)


@router.post("/login", response_model=TokenResponse)
def login(dados: ContaLogin, banco: Session = Depends(obter_sessao)):
    conta = banco.query(Conta).filter(Conta.usuario == dados.usuario).one_or_none()

    if conta is None or not verificar_senha(dados.senha, conta.senha_hash):
        raise HTTPException(status.HTTP_401_UNAUTHORIZED, "Usuário ou senha inválidos.")

    return _criar_sessao(conta, banco)
