import os
from pathlib import Path

from dotenv import load_dotenv

BASE_DIR = Path(__file__).resolve().parent.parent

load_dotenv(BASE_DIR / ".env")

DATABASE_URL = os.getenv("DATABASE_URL", "sqlite:///./visconde.db")
# Alguns provedores (Railway, Heroku) ainda entregam "postgres://", que o
# SQLAlchemy 2 não reconhece mais como alias de "postgresql://".
if DATABASE_URL.startswith("postgres://"):
    DATABASE_URL = DATABASE_URL.replace("postgres://", "postgresql://", 1)

FERNET_KEY = os.getenv("FERNET_KEY", "").strip()
if not FERNET_KEY:
    raise RuntimeError(
        "FERNET_KEY não foi encontrada no .env do backend. Gere uma com: "
        "python -c \"from cryptography.fernet import Fernet; "
        "print(Fernet.generate_key().decode())\""
    )

SESSION_TTL_DIAS = int(os.getenv("SESSION_TTL_DIAS", "30"))
