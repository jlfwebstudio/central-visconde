import json

from fastapi import APIRouter, Depends
from sqlalchemy.orm import Session

from ..auth import obter_conta_atual
from ..database import obter_sessao
from ..models import Conta, ConfigAutomacao, CredencialPlataforma
from ..schemas import ContaConfigIn, ContaConfigOut, CredencialOut
from ..security import cifrar, decifrar

router = APIRouter(prefix="/conta", tags=["conta"])

PLATAFORMAS_SUPORTADAS = ("MOBYAN", "OGEA")


@router.get("/config", response_model=ContaConfigOut)
def obter_config(conta: Conta = Depends(obter_conta_atual)):
    plataformas = {
        credencial.plataforma: CredencialOut(
            url=credencial.url,
            usuario=credencial.usuario,
            senha=decifrar(credencial.senha_criptografada),
            ativo=credencial.ativo,
        )
        for credencial in conta.credenciais
    }

    config_json = conta.config.config_json if conta.config else "{}"

    return ContaConfigOut(
        usuario=conta.usuario,
        is_admin_master=conta.is_admin_master,
        plataformas=plataformas,
        config=json.loads(config_json),
    )


@router.put("/config", response_model=ContaConfigOut)
def atualizar_config(
    dados: ContaConfigIn,
    conta: Conta = Depends(obter_conta_atual),
    banco: Session = Depends(obter_sessao),
):
    existentes = {credencial.plataforma: credencial for credencial in conta.credenciais}

    for plataforma, credencial_in in dados.plataformas.items():
        plataforma = plataforma.upper()
        if plataforma not in PLATAFORMAS_SUPORTADAS:
            continue

        credencial = existentes.get(plataforma)
        if credencial is None:
            credencial = CredencialPlataforma(conta_id=conta.id, plataforma=plataforma)
            banco.add(credencial)

        credencial.url = credencial_in.url
        credencial.usuario = credencial_in.usuario
        if credencial_in.senha:
            credencial.senha_criptografada = cifrar(credencial_in.senha)
        credencial.ativo = credencial_in.ativo

    if conta.config is None:
        conta.config = ConfigAutomacao(conta_id=conta.id, config_json="{}")

    if dados.config:
        atual = json.loads(conta.config.config_json)
        atual.update(dados.config)
        conta.config.config_json = json.dumps(atual)

    banco.commit()
    banco.refresh(conta)

    return obter_config(conta)
