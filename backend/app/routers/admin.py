from fastapi import APIRouter, Depends, HTTPException, status
from sqlalchemy.orm import Session

from ..auth import exigir_admin_master
from ..database import obter_sessao
from ..models import Conta, VersaoApp
from ..schemas import AdminContaResumo, ContaConfigOut, VersaoIn, VersaoOut
from .conta import obter_config

router = APIRouter(prefix="/admin", tags=["admin"])


@router.get("/contas", response_model=list[AdminContaResumo])
def listar_contas(
    _admin: Conta = Depends(exigir_admin_master),
    banco: Session = Depends(obter_sessao),
):
    contas = banco.query(Conta).order_by(Conta.criado_em).all()

    return [
        AdminContaResumo(
            id=conta.id,
            usuario=conta.usuario,
            criado_em=conta.criado_em,
            is_admin_master=conta.is_admin_master,
            plataformas_ativas=[c.plataforma for c in conta.credenciais if c.ativo],
        )
        for conta in contas
    ]


@router.get("/contas/{conta_id}", response_model=ContaConfigOut)
def detalhar_conta(
    conta_id: int,
    _admin: Conta = Depends(exigir_admin_master),
    banco: Session = Depends(obter_sessao),
):
    conta = banco.get(Conta, conta_id)
    if conta is None:
        raise HTTPException(status.HTTP_404_NOT_FOUND, "Conta não encontrada.")

    return obter_config(conta)


@router.post("/versoes", response_model=VersaoOut)
def publicar_versao(
    dados: VersaoIn,
    _admin: Conta = Depends(exigir_admin_master),
    banco: Session = Depends(obter_sessao),
):
    versao = VersaoApp(
        plataforma=dados.plataforma,
        versao=dados.versao,
        url_download=dados.url_download,
        obrigatoria=dados.obrigatoria,
        notas=dados.notas,
    )
    banco.add(versao)
    banco.commit()
    banco.refresh(versao)

    return versao
