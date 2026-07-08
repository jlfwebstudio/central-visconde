from datetime import datetime

from sqlalchemy import Boolean, DateTime, ForeignKey, Integer, String, Text, UniqueConstraint
from sqlalchemy.orm import Mapped, mapped_column, relationship

from .database import Base


def agora():
    # SQLite não preserva tzinfo entre escrita/leitura, então guardamos
    # sempre UTC "naive" pra evitar comparar aware x naive depois.
    return datetime.utcnow()


class Conta(Base):
    __tablename__ = "contas"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    usuario: Mapped[str] = mapped_column(String(80), unique=True, index=True)
    senha_hash: Mapped[str] = mapped_column(String(255))
    is_admin_master: Mapped[bool] = mapped_column(Boolean, default=False)
    criado_em: Mapped[datetime] = mapped_column(DateTime(), default=agora)

    credenciais: Mapped[list["CredencialPlataforma"]] = relationship(
        back_populates="conta", cascade="all, delete-orphan"
    )
    config: Mapped["ConfigAutomacao"] = relationship(
        back_populates="conta", cascade="all, delete-orphan", uselist=False
    )
    sessoes: Mapped[list["Sessao"]] = relationship(
        back_populates="conta", cascade="all, delete-orphan"
    )


class CredencialPlataforma(Base):
    __tablename__ = "credenciais_plataforma"
    __table_args__ = (UniqueConstraint("conta_id", "plataforma", name="uq_conta_plataforma"),)

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    conta_id: Mapped[int] = mapped_column(ForeignKey("contas.id"))
    plataforma: Mapped[str] = mapped_column(String(20))  # "MOBYAN" ou "OGEA"
    url: Mapped[str] = mapped_column(String(255), default="")
    usuario: Mapped[str] = mapped_column(String(255), default="")
    senha_criptografada: Mapped[str] = mapped_column(Text, default="")
    ativo: Mapped[bool] = mapped_column(Boolean, default=True)

    conta: Mapped["Conta"] = relationship(back_populates="credenciais")


class ConfigAutomacao(Base):
    __tablename__ = "config_automacao"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    conta_id: Mapped[int] = mapped_column(ForeignKey("contas.id"), unique=True)
    config_json: Mapped[str] = mapped_column(Text, default="{}")

    conta: Mapped["Conta"] = relationship(back_populates="config")


class Sessao(Base):
    __tablename__ = "sessoes"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    conta_id: Mapped[int] = mapped_column(ForeignKey("contas.id"))
    token: Mapped[str] = mapped_column(String(64), unique=True, index=True)
    criado_em: Mapped[datetime] = mapped_column(DateTime(), default=agora)
    expira_em: Mapped[datetime] = mapped_column(DateTime())

    conta: Mapped["Conta"] = relationship(back_populates="sessoes")


class VersaoApp(Base):
    __tablename__ = "versoes_app"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    plataforma: Mapped[str] = mapped_column(String(20))  # "mac" ou "windows"
    versao: Mapped[str] = mapped_column(String(20))
    url_download: Mapped[str] = mapped_column(String(500))
    obrigatoria: Mapped[bool] = mapped_column(Boolean, default=False)
    notas: Mapped[str] = mapped_column(Text, default="")
    publicada_em: Mapped[datetime] = mapped_column(DateTime(), default=agora)
