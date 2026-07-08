import secrets

import bcrypt
from cryptography.fernet import Fernet, InvalidToken

from .config import FERNET_KEY

_fernet = Fernet(FERNET_KEY.encode())


def hash_senha(senha: str) -> str:
    return bcrypt.hashpw(senha.encode(), bcrypt.gensalt()).decode()


def verificar_senha(senha: str, senha_hash: str) -> bool:
    return bcrypt.checkpw(senha.encode(), senha_hash.encode())


def cifrar(texto: str) -> str:
    return _fernet.encrypt(texto.encode()).decode()


def decifrar(texto_cifrado: str) -> str:
    if not texto_cifrado:
        return ""
    try:
        return _fernet.decrypt(texto_cifrado.encode()).decode()
    except InvalidToken:
        return ""


def gerar_token_sessao() -> str:
    return secrets.token_urlsafe(32)
