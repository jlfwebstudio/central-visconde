from fastapi import APIRouter, Depends, HTTPException, status
from sqlalchemy.orm import Session

from ..database import obter_sessao
from ..models import VersaoApp
from ..schemas import VersaoOut

router = APIRouter(prefix="/versao", tags=["versao"])


@router.get("/atual", response_model=VersaoOut)
def obter_versao_atual(plataforma: str, banco: Session = Depends(obter_sessao)):
    versao = (
        banco.query(VersaoApp)
        .filter(VersaoApp.plataforma == plataforma)
        .order_by(VersaoApp.publicada_em.desc())
        .first()
    )
    if versao is None:
        raise HTTPException(status.HTTP_404_NOT_FOUND, "Nenhuma versão publicada para essa plataforma.")

    return versao
