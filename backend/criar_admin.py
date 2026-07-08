"""Cria (ou promove) uma conta admin master. Uso:

    .venv/bin/python criar_admin.py <usuario> <senha>
"""
import sys

from app.database import Base, SessionLocal, engine
from app.models import Conta, ConfigAutomacao
from app.security import hash_senha


def main():
    if len(sys.argv) != 3:
        print("Uso: python criar_admin.py <usuario> <senha>")
        raise SystemExit(1)

    usuario, senha = sys.argv[1], sys.argv[2]

    Base.metadata.create_all(bind=engine)
    banco = SessionLocal()

    try:
        conta = banco.query(Conta).filter(Conta.usuario == usuario).one_or_none()

        if conta is None:
            conta = Conta(usuario=usuario, senha_hash=hash_senha(senha), is_admin_master=True)
            banco.add(conta)
            banco.flush()
            banco.add(ConfigAutomacao(conta_id=conta.id, config_json="{}"))
            print(f"Conta admin master '{usuario}' criada.")
        else:
            conta.senha_hash = hash_senha(senha)
            conta.is_admin_master = True
            print(f"Conta '{usuario}' promovida a admin master e senha atualizada.")

        banco.commit()
    finally:
        banco.close()


if __name__ == "__main__":
    main()
