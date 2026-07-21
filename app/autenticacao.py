import json
import os
import shutil
import sys
import tkinter as tk
import urllib.error
import urllib.request
from tkinter import messagebox

try:
    from PIL import Image, ImageTk
except Exception:
    Image = None
    ImageTk = None

from caminho_base import BASE_DIR, RECURSOS_DIR
from versao import VERSAO_ATUAL

ARQUIVO_ICONE_PNG = RECURSOS_DIR / "assets" / "logo_visconde_app.png"
ARQUIVO_ICONE_ICO = RECURSOS_DIR / "assets" / "logo_visconde.ico"

PASTA_CONFIG = BASE_DIR / "config"
ARQUIVO_SESSAO = PASTA_CONFIG / "sessao.json"
ARQUIVO_ENV = BASE_DIR / ".env"

# Criado na primeira vez que essa máquina faz login numa conta. Diferente de
# ARQUIVO_SESSAO, não é apagado no logout — é o que permite distinguir "máquina
# antiga que nunca usou contas" (pode pular login) de "já usa contas, só saiu
# agora" (tem que pedir login de novo, mesmo com um .env antigo ainda no disco).
ARQUIVO_MARCADOR_CONTA = PASTA_CONFIG / "conta_ativada"

# Cofre onde ficam guardados os dados de execução de cada conta que já logou
# nessa máquina, enquanto ela não está ativa (ver trocar_para_conta).
PASTA_WORKSPACES = PASTA_CONFIG / "workspaces"
ARQUIVO_WORKSPACE_ATIVO = PASTA_CONFIG / "workspace_ativo.json"

# Pastas com dados de execução específicos de cada conta: relatórios/PDFs
# gerados, downloads, logs, planilhas de referência (prestadores, regras de
# roteirização) e a sessão persistente do WhatsApp Web. Não inclui o .env
# (sempre regenerado do zero a partir da config da conta em materializar_env)
# nem backups/ (snapshots manuais do usuário, tratados como legado — ver
# CLAUDE.md — e não amarrados a nenhuma conta específica).
PASTAS_WORKSPACE = ("outputs", "downloads", "logs", "bases", "whatsapp_profile")

BACKEND_URL = os.getenv(
    "VISCONDE_BACKEND_URL",
    "https://central-visconde-api-production.up.railway.app",
).rstrip("/")

# A OGEA tem uma URL de login estável e igual pra todo mundo (mesmo padrão já
# usado como fallback em baixar_relatorios_roteirizacao.py). A Mobyan, por outro
# lado, tem uma URL com um ID de sessão embutido — não é seguro fixar isso como
# padrão global, então ela é pedida no cadastro junto com usuário/senha.
OGEA_URL_PADRAO = "https://tefti.workfinity.com.br/login/auth"
OGEA_ORDEM_SERVICO_URL_PADRAO = "https://tefti.workfinity.com.br/serviceOrder/index"

COR_FUNDO = "#080808"
COR_FUNDO_2 = "#101010"
COR_CARD = "#171717"
COR_BORDA = "#584A18"
COR_DOURADO = "#F4C430"
COR_DOURADO_HOVER = "#D8AA18"
COR_BRANCO = "#F5F5F5"
COR_TEXTO_SECUNDARIO = "#CFCFCF"
COR_TEXTO_FRACO = "#8E8E8E"
COR_VERDE = "#2EAD68"
COR_VERDE_HOVER = "#248A54"
COR_VERMELHO = "#D14949"


class ErroAutenticacao(Exception):
    pass


class ConexaoBackendError(Exception):
    pass


# ---------------------------------------------------------------------------
# Cliente HTTP do backend de contas
# ---------------------------------------------------------------------------

def _requisitar(metodo, caminho, corpo=None, token=None, timeout=15):
    url = f"{BACKEND_URL}{caminho}"
    dados = json.dumps(corpo).encode("utf-8") if corpo is not None else None

    requisicao = urllib.request.Request(url, data=dados, method=metodo)
    requisicao.add_header("Content-Type", "application/json")
    if token:
        requisicao.add_header("Authorization", f"Bearer {token}")

    try:
        with urllib.request.urlopen(requisicao, timeout=timeout) as resposta:
            bruto = resposta.read()
            return resposta.status, (json.loads(bruto) if bruto else {})
    except urllib.error.HTTPError as erro:
        bruto = erro.read()
        try:
            detalhe = json.loads(bruto).get("detail", "Erro desconhecido.")
        except Exception:
            detalhe = bruto.decode(errors="ignore") or f"Erro HTTP {erro.code}"
        return erro.code, {"detail": detalhe}
    except urllib.error.URLError as erro:
        raise ConexaoBackendError(str(erro.reason)) from erro


def signup(usuario, senha):
    status, corpo = _requisitar("POST", "/auth/signup", {"usuario": usuario, "senha": senha})
    if status != 200:
        raise ErroAutenticacao(corpo.get("detail", "Não consegui criar a conta."))
    return corpo


def login(usuario, senha):
    status, corpo = _requisitar("POST", "/auth/login", {"usuario": usuario, "senha": senha})
    if status != 200:
        raise ErroAutenticacao(corpo.get("detail", "Usuário ou senha inválidos."))
    return corpo


def obter_config(token):
    status, corpo = _requisitar("GET", "/conta/config", token=token)
    if status != 200:
        raise ErroAutenticacao(corpo.get("detail", "Não consegui buscar a configuração da conta."))
    return corpo


def salvar_config(token, plataformas, config_extra=None):
    corpo = {"plataformas": plataformas, "config": config_extra or {}}
    status, resposta = _requisitar("PUT", "/conta/config", corpo, token=token)
    if status != 200:
        raise ErroAutenticacao(resposta.get("detail", "Não consegui salvar a configuração da conta."))
    return resposta


# ---------------------------------------------------------------------------
# Checagem de atualização do app (só relevante em builds empacotados — ver
# app/versao.py e a Fase 4 do plano de multi-tenant)
# ---------------------------------------------------------------------------

def _plataforma_atual():
    return "mac" if sys.platform == "darwin" else "windows"


def _versao_para_tupla(versao):
    partes = []
    for pedaco in str(versao).strip().split("."):
        try:
            partes.append(int(pedaco))
        except ValueError:
            partes.append(0)
    return tuple(partes)


def verificar_atualizacao():
    """Consulta o backend por uma versão mais nova pra essa plataforma.

    Retorna o dict da versão publicada se for mais nova que VERSAO_ATUAL, ou
    None se já está atualizado, se nada foi publicado ainda, ou se a checagem
    falhar por qualquer motivo — nunca deve impedir o uso normal do app."""
    try:
        status, corpo = _requisitar(
            "GET", f"/versao/atual?plataforma={_plataforma_atual()}", timeout=8,
        )
    except ConexaoBackendError:
        return None

    if status != 200:
        return None

    if _versao_para_tupla(corpo.get("versao", "0")) <= _versao_para_tupla(VERSAO_ATUAL):
        return None

    return corpo


def baixar_atualizacao(url_download, destino):
    """Baixa o instalador da atualização pra um caminho local."""
    try:
        urllib.request.urlretrieve(url_download, str(destino))
    except Exception as erro:
        raise ConexaoBackendError(str(erro)) from erro
    return destino


# ---------------------------------------------------------------------------
# Sessão local (evita pedir login em toda abertura do app)
# ---------------------------------------------------------------------------

def carregar_sessao_local():
    if not ARQUIVO_SESSAO.exists():
        return None
    try:
        return json.loads(ARQUIVO_SESSAO.read_text(encoding="utf-8"))
    except Exception:
        return None


def salvar_sessao_local(usuario, token, expira_em):
    PASTA_CONFIG.mkdir(parents=True, exist_ok=True)
    ARQUIVO_SESSAO.write_text(
        json.dumps(
            {"usuario": usuario, "token": token, "expira_em": expira_em},
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )
    try:
        os.chmod(ARQUIVO_SESSAO, 0o600)
    except Exception:
        pass

    if not ARQUIVO_MARCADOR_CONTA.exists():
        ARQUIVO_MARCADOR_CONTA.write_text("", encoding="utf-8")


def limpar_sessao_local():
    """Faz logout: derruba a sessão local, mas mantém o marcador de que essa
    máquina já usa contas — pra não voltar a pular a tela de login."""
    if ARQUIVO_SESSAO.exists():
        ARQUIVO_SESSAO.unlink()


# ---------------------------------------------------------------------------
# Workspace por conta (evita que uma conta rode automação em cima dos dados
# de execução de outra, quando a mesma máquina faz login em contas diferentes
# — ex: suporte entrando na conta de um cliente)
# ---------------------------------------------------------------------------

def _identificador_workspace(usuario):
    limpo = "".join(
        caractere if caractere.isalnum() else "_"
        for caractere in usuario.strip().lower()
    )
    return limpo.strip("_") or "conta"


def _pasta_workspace(identificador):
    return PASTA_WORKSPACES / identificador


def _workspace_ativo():
    if not ARQUIVO_WORKSPACE_ATIVO.exists():
        return None
    try:
        dados = json.loads(ARQUIVO_WORKSPACE_ATIVO.read_text(encoding="utf-8"))
        return dados.get("identificador")
    except Exception:
        return None


def _definir_workspace_ativo(identificador):
    PASTA_CONFIG.mkdir(parents=True, exist_ok=True)
    ARQUIVO_WORKSPACE_ATIVO.write_text(
        json.dumps({"identificador": identificador}, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def _arquivar_workspace_atual(identificador):
    """Move os dados de execução atuais (pertencentes à conta que está saindo)
    para o cofre dela, sem apagar nada."""
    destino = _pasta_workspace(identificador)
    destino.mkdir(parents=True, exist_ok=True)

    for nome in PASTAS_WORKSPACE:
        origem = BASE_DIR / nome
        if not origem.exists():
            continue

        alvo = destino / nome
        if alvo.exists():
            shutil.rmtree(alvo)

        shutil.move(str(origem), str(alvo))


def _restaurar_workspace(identificador):
    """Traz de volta os dados de execução da conta que está entrando, se ela
    já tiver logado nessa máquina antes. Nunca sobrescreve algo que já esteja
    no lugar (não deveria acontecer, mas evita perda de dado por segurança)."""
    origem = _pasta_workspace(identificador)

    for nome in PASTAS_WORKSPACE:
        item_origem = origem / nome
        item_destino = BASE_DIR / nome
        if item_destino.exists() or not item_origem.exists():
            continue

        shutil.move(str(item_origem), str(item_destino))


def trocar_para_conta(usuario):
    """Garante que outputs/downloads/logs/bases/whatsapp_profile em disco
    pertencem à conta que está prestes a ficar ativa.

    Na primeira vez que essa máquina reconhece uma conta (instalação legada
    virando conta, ou o primeiro login de sempre), assume os dados já
    presentes como sendo dessa conta, sem mover nada. Da segunda troca de
    conta em diante, arquiva os dados da conta anterior e restaura os da
    nova (ou começa vazio, se ela nunca logou nessa máquina antes)."""
    identificador = _identificador_workspace(usuario)
    ativo = _workspace_ativo()

    if ativo is None:
        _definir_workspace_ativo(identificador)
        return

    if ativo == identificador:
        return

    _arquivar_workspace_atual(ativo)
    _restaurar_workspace(identificador)
    _definir_workspace_ativo(identificador)


# ---------------------------------------------------------------------------
# Materialização do .env local a partir da config da conta
# ---------------------------------------------------------------------------

def materializar_env(config_conta):
    plataformas = config_conta.get("plataformas", {})
    config_extra = config_conta.get("config", {})

    linhas = [
        "# Gerado automaticamente a partir da conta logada no ViscondeApp.",
        "# Não edite à mão — as credenciais são gerenciadas pela conta.",
        "",
    ]

    mobyan = plataformas.get("MOBYAN")
    if mobyan and mobyan.get("ativo", True):
        linhas += [
            f"MOBYAN_URL={mobyan.get('url', '')}",
            f"MOBYAN_USUARIO={mobyan.get('usuario', '')}",
            f"MOBYAN_SENHA={mobyan.get('senha', '')}",
        ]

    ogea = plataformas.get("OGEA")
    if ogea and ogea.get("ativo", True):
        linhas += [
            f"OGEA_URL={ogea.get('url') or OGEA_URL_PADRAO}",
            f"OGEA_USUARIO={ogea.get('usuario', '')}",
            f"OGEA_SENHA={ogea.get('senha', '')}",
            "OGEA_ORDEM_SERVICO_URL="
            + str(config_extra.get("OGEA_ORDEM_SERVICO_URL", OGEA_ORDEM_SERVICO_URL_PADRAO)),
        ]

    for chave in (
        "OGEA_DIAS_ABERTURA",
        "OGEA_TIMEOUT_CARREGAMENTO",
        "PDF_TIMEOUT_SEGUNDOS",
        "MOBYAN_PRESTADORES",
        "MOBYAN_ESTADO",
        "OGEA_PRESTADOR",
    ):
        valor = config_extra.get(chave)
        if valor is None or valor == "":
            continue
        if isinstance(valor, list):
            valor = ",".join(str(item) for item in valor)
        linhas.append(f"{chave}={valor}")

    ARQUIVO_ENV.write_text("\n".join(linhas) + "\n", encoding="utf-8")
    try:
        os.chmod(ARQUIVO_ENV, 0o600)
    except Exception:
        pass


# ---------------------------------------------------------------------------
# Ponto de entrada: garante uma sessão válida antes de abrir a Central
# ---------------------------------------------------------------------------

def garantir_sessao_valida():
    """Retorna True se o app pode abrir normalmente, False se o usuário
    fechou a tela de login sem autenticar (o app deve encerrar nesse caso)."""

    sessao = carregar_sessao_local()

    if sessao is None:
        if ARQUIVO_ENV.exists() and not ARQUIVO_MARCADOR_CONTA.exists():
            # Instalação existente de antes do sistema de contas (nunca fez
            # login nenhuma vez): mantém funcionando como está, sem forçar
            # login numa conta nova.
            return True
        return _abrir_janela_login()

    try:
        config_conta = obter_config(sessao["token"])
        trocar_para_conta(sessao["usuario"])
        materializar_env(config_conta)
        return True
    except ConexaoBackendError:
        if ARQUIVO_ENV.exists():
            # Sem internet pro backend agora, mas já existe uma configuração
            # local válida de um login anterior — segue operando offline.
            return True
        return _abrir_janela_login()
    except ErroAutenticacao:
        limpar_sessao_local()
        return _abrir_janela_login()


def _configurar_icone_janela(root):
    """Evita que o Windows prenda o ícone genérico na barra de tarefas — essa
    é a primeira janela do processo, então precisa do mesmo ícone custom que
    central_mobyan.py aplica na janela principal (ver configurar_icone_janela
    lá), senão o Windows já "trava" o ícone padrão nesse momento inicial."""
    if ARQUIVO_ICONE_PNG.exists() and Image is not None and ImageTk is not None:
        try:
            imagem = Image.open(ARQUIVO_ICONE_PNG)
            imagem.thumbnail((256, 256), Image.Resampling.LANCZOS)
            icone = ImageTk.PhotoImage(imagem)
            root._icone_visconde = icone
            root.iconphoto(True, icone)
        except Exception:
            pass

    if os.name == "nt" and ARQUIVO_ICONE_ICO.exists():
        try:
            root.iconbitmap(default=str(ARQUIVO_ICONE_ICO))
        except Exception:
            pass


def _abrir_janela_login():
    root = tk.Tk()
    _configurar_icone_janela(root)
    janela = _JanelaAutenticacao(root)
    root.mainloop()
    return janela.autenticado


# ---------------------------------------------------------------------------
# Interface gráfica
# ---------------------------------------------------------------------------

class _TelaFormulario:
    """Widgets de formulário compartilhados entre a tela de login/cadastro e a
    tela de configurações da conta (edição pós-cadastro) — mesma identidade
    visual, sem duplicar os construtores de campo/botão em cada classe."""

    def _limpar_container(self):
        for widget in self.container.winfo_children():
            widget.destroy()

    def _titulo(self, texto, subtitulo=None):
        tk.Label(
            self.container,
            text=texto,
            bg=COR_FUNDO,
            fg=COR_DOURADO,
            font=("Arial", 20, "bold"),
        ).pack(anchor="w", pady=(0, 4))
        if subtitulo:
            tk.Label(
                self.container,
                text=subtitulo,
                bg=COR_FUNDO,
                fg=COR_TEXTO_SECUNDARIO,
                font=("Arial", 11),
                wraplength=400,
                justify="left",
            ).pack(anchor="w", pady=(0, 18))

    def _campo(self, rotulo, mostrar=None):
        tk.Label(
            self.container,
            text=rotulo,
            bg=COR_FUNDO,
            fg=COR_TEXTO_SECUNDARIO,
            font=("Arial", 10),
        ).pack(anchor="w", pady=(10, 2))
        entrada = tk.Entry(
            self.container,
            bg=COR_CARD,
            fg=COR_BRANCO,
            insertbackground=COR_BRANCO,
            relief="flat",
            font=("Arial", 12),
            show=mostrar,
        )
        entrada.pack(fill="x", ipady=6)
        return entrada

    def _botao(self, texto, comando, cor=COR_DOURADO, cor_hover=COR_DOURADO_HOVER, cor_texto="#000000"):
        botao = tk.Label(
            self.container,
            text=texto,
            bg=cor,
            fg=cor_texto,
            font=("Arial", 11, "bold"),
            cursor="hand2",
            pady=10,
        )
        botao.pack(fill="x", pady=(14, 0))
        botao.bind("<Button-1>", lambda _evento: comando())
        botao.bind("<Enter>", lambda _evento: botao.configure(bg=cor_hover))
        botao.bind("<Leave>", lambda _evento: botao.configure(bg=cor))
        return botao

    def _link(self, texto, comando):
        rotulo = tk.Label(
            self.container,
            text=texto,
            bg=COR_FUNDO,
            fg=COR_TEXTO_FRACO,
            font=("Arial", 10, "underline"),
            cursor="hand2",
        )
        rotulo.pack(pady=(16, 0))
        rotulo.bind("<Button-1>", lambda _evento: comando())
        return rotulo


class _JanelaAutenticacao(_TelaFormulario):
    def __init__(self, root):
        self.root = root
        self.autenticado = False
        self.token = None
        self.usuario = None

        self.root.title("ViscondeApp — Entrar")
        self.root.configure(bg=COR_FUNDO)
        self.root.geometry("460x620")
        self.root.minsize(420, 560)
        self.root.protocol("WM_DELETE_WINDOW", self._fechar)

        self.container = tk.Frame(self.root, bg=COR_FUNDO)
        self.container.pack(fill="both", expand=True, padx=28, pady=28)

        self._mostrar_login()

    def _fechar(self):
        self.autenticado = False
        self.root.destroy()

    # -- Tela 1: login/cadastro -------------------------------------------------

    def _mostrar_login(self):
        self._limpar_container()
        self._titulo("ViscondeApp", "Entre com a sua conta ou crie uma nova.")

        self.campo_usuario = self._campo("Usuário")
        self.campo_senha = self._campo("Senha", mostrar="*")

        self._botao("Entrar", self._fazer_login)
        self._botao(
            "Criar conta",
            self._mostrar_cadastro,
            cor=COR_CARD,
            cor_hover=COR_BORDA,
            cor_texto=COR_BRANCO,
        )

    def _fazer_login(self):
        usuario = self.campo_usuario.get().strip()
        senha = self.campo_senha.get()

        if not usuario or not senha:
            messagebox.showwarning("ViscondeApp", "Preencha usuário e senha.")
            return

        try:
            resposta = login(usuario, senha)
            config_conta = obter_config(resposta["token"])
            trocar_para_conta(usuario)
            materializar_env(config_conta)
            salvar_sessao_local(usuario, resposta["token"], resposta["expira_em"])
        except ConexaoBackendError:
            messagebox.showerror(
                "ViscondeApp",
                "Não consegui falar com o servidor de contas. Verifique sua internet e tente de novo.",
            )
            return
        except ErroAutenticacao as erro:
            messagebox.showerror("ViscondeApp", str(erro))
            return

        self.autenticado = True
        self.root.destroy()

    # -- Tela 2: cadastro (usuário/senha da conta) ------------------------------

    def _mostrar_cadastro(self):
        self._limpar_container()
        self._titulo("Criar conta", "Escolha um usuário e uma senha só sua — sem precisar de e-mail.")

        self.campo_novo_usuario = self._campo("Usuário")
        self.campo_nova_senha = self._campo("Senha (mínimo 6 caracteres)", mostrar="*")

        self._botao("Continuar", self._criar_conta)
        self._link("Voltar", self._mostrar_login)

    def _criar_conta(self):
        usuario = self.campo_novo_usuario.get().strip()
        senha = self.campo_nova_senha.get()

        if not usuario or not senha:
            messagebox.showwarning("ViscondeApp", "Preencha usuário e senha.")
            return

        try:
            resposta = signup(usuario, senha)
        except ConexaoBackendError:
            messagebox.showerror(
                "ViscondeApp",
                "Não consegui falar com o servidor de contas. Verifique sua internet e tente de novo.",
            )
            return
        except ErroAutenticacao as erro:
            messagebox.showerror("ViscondeApp", str(erro))
            return

        self.token = resposta["token"]
        self.usuario = usuario
        self.expira_em = resposta["expira_em"]
        self.plataformas = {}
        self.config_extra = {}
        self._mostrar_pergunta_ogea()

    # -- Tela 3: pergunta OGEA ---------------------------------------------------

    def _mostrar_pergunta_ogea(self):
        self._limpar_container()
        self._titulo("OGEA", "Você quer que a automação funcione na OGEA?")

        self._botao("Sim", self._mostrar_credenciais_ogea, cor=COR_VERDE, cor_hover=COR_VERDE_HOVER, cor_texto=COR_BRANCO)
        self._botao(
            "Não, pular",
            self._mostrar_pergunta_mobyan,
            cor=COR_CARD,
            cor_hover=COR_BORDA,
            cor_texto=COR_BRANCO,
        )

    def _mostrar_credenciais_ogea(self):
        self._limpar_container()
        self._titulo("OGEA", "Login que você já usa pra entrar na OGEA (um usuário só cobre todos os adquirentes).")

        self.campo_ogea_usuario = self._campo("Usuário OGEA")
        self.campo_ogea_senha = self._campo("Senha OGEA", mostrar="*")
        self.campo_ogea_base = self._campo(
            "Nome exato da sua base na OGEA (ex: SMART TECNOLOGIA) — é o que aparece "
            "lá dentro do sistema, pode ser diferente do nome da sua empresa"
        )

        self._botao("Continuar", self._salvar_ogea_e_avancar)
        self._link("Pular OGEA", self._mostrar_pergunta_mobyan)

    def _salvar_ogea_e_avancar(self):
        usuario = self.campo_ogea_usuario.get().strip()
        senha = self.campo_ogea_senha.get()
        base = self.campo_ogea_base.get().strip()

        if not usuario or not senha or not base:
            messagebox.showwarning("ViscondeApp", "Preencha usuário, senha e o nome da base da OGEA.")
            return

        self.plataformas["OGEA"] = {
            "url": OGEA_URL_PADRAO,
            "usuario": usuario,
            "senha": senha,
            "ativo": True,
        }
        self.config_extra["OGEA_PRESTADOR"] = base
        self._mostrar_pergunta_mobyan()

    # -- Tela 4: pergunta Mobyan --------------------------------------------------

    def _mostrar_pergunta_mobyan(self):
        self._limpar_container()
        self._titulo("Mobyan", "Você quer que a automação funcione na Mobyan?")

        self._botao("Sim", self._mostrar_credenciais_mobyan, cor=COR_VERDE, cor_hover=COR_VERDE_HOVER, cor_texto=COR_BRANCO)
        self._botao(
            "Não, pular",
            self._mostrar_aviso_fedex,
            cor=COR_CARD,
            cor_hover=COR_BORDA,
            cor_texto=COR_BRANCO,
        )

    def _mostrar_credenciais_mobyan(self):
        self._limpar_container()
        self._titulo("Mobyan", "Login que você já usa pra entrar na Mobyan (cliente Ponto).")

        self.campo_mobyan_url = self._campo("URL de login da Mobyan")
        self.campo_mobyan_usuario = self._campo("Usuário Mobyan")
        self.campo_mobyan_senha = self._campo("Senha Mobyan", mostrar="*")
        self.campo_mobyan_base = self._campo(
            "Nome exato da(s) sua(s) base(s) na Mobyan — separe por vírgula se tiver "
            "mais de uma (ex: RS-SMART, RS-SMART - PELOTAS)"
        )
        self.campo_mobyan_estado = self._campo("Sigla do estado filtrado na Mobyan (ex: RS)")

        self._botao("Continuar", self._salvar_mobyan_e_avancar)
        self._link("Pular Mobyan", self._mostrar_aviso_fedex)

    def _salvar_mobyan_e_avancar(self):
        url = self.campo_mobyan_url.get().strip()
        usuario = self.campo_mobyan_usuario.get().strip()
        senha = self.campo_mobyan_senha.get()
        base = self.campo_mobyan_base.get().strip()
        estado = self.campo_mobyan_estado.get().strip()

        if not url or not usuario or not senha or not base or not estado:
            messagebox.showwarning("ViscondeApp", "Preencha URL, usuário, senha, base e estado da Mobyan.")
            return

        self.plataformas["MOBYAN"] = {
            "url": url,
            "usuario": usuario,
            "senha": senha,
            "ativo": True,
        }
        self.config_extra["MOBYAN_PRESTADORES"] = base
        self.config_extra["MOBYAN_ESTADO"] = estado
        self._mostrar_aviso_fedex()

    # -- Tela 5: aviso FedEx + conclusão ------------------------------------------

    def _mostrar_aviso_fedex(self):
        self._limpar_container()
        self._titulo(
            "Quase lá",
            "Por enquanto ainda não temos suporte para FedEx. "
            "Assim que estiver disponível, você poderá habilitar por aqui.",
        )
        self._botao("Concluir cadastro", self._concluir_cadastro)

    def _concluir_cadastro(self):
        try:
            salvar_config(self.token, self.plataformas, self.config_extra)
            config_conta = obter_config(self.token)
            trocar_para_conta(self.usuario)
            materializar_env(config_conta)
            salvar_sessao_local(self.usuario, self.token, self.expira_em)
        except ConexaoBackendError:
            messagebox.showerror(
                "ViscondeApp",
                "Não consegui falar com o servidor de contas. Verifique sua internet e tente de novo.",
            )
            return
        except ErroAutenticacao as erro:
            messagebox.showerror("ViscondeApp", str(erro))
            return

        self.autenticado = True
        self.root.destroy()


# ---------------------------------------------------------------------------
# Configurações da conta (edição pós-cadastro: credenciais e bases)
# ---------------------------------------------------------------------------

class _JanelaConfiguracaoConta(_TelaFormulario):
    """Formulário único (sem o passo a passo do cadastro) pra editar depois
    o que foi respondido na hora de criar a conta — principalmente o nome
    exato da base em cada plataforma, já que uma conta pode passar a puxar
    bases adicionais depois do cadastro inicial."""

    def __init__(self, master, token, usuario, config_conta, ao_salvar=None):
        self.token = token
        self.usuario = usuario
        self.ao_salvar = ao_salvar

        plataformas = config_conta.get("plataformas", {})
        self.ogea_atual = plataformas.get("OGEA") or {}
        self.mobyan_atual = plataformas.get("MOBYAN") or {}
        self.config_atual = config_conta.get("config", {})

        self.root = tk.Toplevel(master)
        self.root.title("ViscondeApp — Configurações da conta")
        self.root.configure(bg=COR_FUNDO)
        self.root.geometry("520x760")
        self.root.minsize(460, 620)
        self.root.transient(master)
        self.root.grab_set()

        self.container = tk.Frame(self.root, bg=COR_FUNDO)
        self.container.pack(fill="both", expand=True, padx=28, pady=28)

        self._montar_formulario()

    def _secao(self, texto):
        tk.Label(
            self.container,
            text=texto,
            bg=COR_FUNDO,
            fg=COR_DOURADO,
            font=("Arial", 12, "bold"),
        ).pack(anchor="w", pady=(18, 0))

    def _montar_formulario(self):
        self._titulo("Configurações da conta", f"Conta: {self.usuario}")

        self._secao("OGEA")
        self.campo_ogea_usuario = self._campo("Usuário OGEA")
        self.campo_ogea_usuario.insert(0, self.ogea_atual.get("usuario", ""))
        self.campo_ogea_senha = self._campo("Senha OGEA (deixe em branco pra manter a atual)", mostrar="*")
        self.campo_ogea_base = self._campo("Nome exato da sua base na OGEA")
        self.campo_ogea_base.insert(0, self.config_atual.get("OGEA_PRESTADOR", ""))

        self._secao("MOBYAN")
        self.campo_mobyan_url = self._campo("URL de login da Mobyan")
        self.campo_mobyan_url.insert(0, self.mobyan_atual.get("url", ""))
        self.campo_mobyan_usuario = self._campo("Usuário Mobyan")
        self.campo_mobyan_usuario.insert(0, self.mobyan_atual.get("usuario", ""))
        self.campo_mobyan_senha = self._campo("Senha Mobyan (deixe em branco pra manter a atual)", mostrar="*")
        self.campo_mobyan_base = self._campo("Base(s) na Mobyan — separadas por vírgula")
        prestadores = self.config_atual.get("MOBYAN_PRESTADORES", "")
        if isinstance(prestadores, list):
            prestadores = ", ".join(prestadores)
        self.campo_mobyan_base.insert(0, prestadores)
        self.campo_mobyan_estado = self._campo("Sigla do estado filtrado na Mobyan")
        self.campo_mobyan_estado.insert(0, self.config_atual.get("MOBYAN_ESTADO", ""))

        self._botao("Salvar alterações", self._salvar)
        self._link("Fechar sem salvar", self.root.destroy)

    def _salvar(self):
        plataformas = {}

        usuario_ogea = self.campo_ogea_usuario.get().strip()
        base_ogea = self.campo_ogea_base.get().strip()
        senha_ogea = self.campo_ogea_senha.get()
        if usuario_ogea and base_ogea:
            if not senha_ogea and not self.ogea_atual.get("usuario"):
                messagebox.showwarning("ViscondeApp", "Preencha a senha da OGEA (é a primeira vez que essa plataforma é configurada).")
                return
            plataformas["OGEA"] = {
                "url": self.ogea_atual.get("url") or OGEA_URL_PADRAO,
                "usuario": usuario_ogea,
                "senha": senha_ogea,
                "ativo": True,
            }
        elif usuario_ogea or base_ogea:
            messagebox.showwarning("ViscondeApp", "Preencha usuário e base da OGEA (ou deixe os dois em branco).")
            return

        url_mobyan = self.campo_mobyan_url.get().strip()
        usuario_mobyan = self.campo_mobyan_usuario.get().strip()
        base_mobyan = self.campo_mobyan_base.get().strip()
        estado_mobyan = self.campo_mobyan_estado.get().strip()
        senha_mobyan = self.campo_mobyan_senha.get()
        if url_mobyan and usuario_mobyan and base_mobyan and estado_mobyan:
            if not senha_mobyan and not self.mobyan_atual.get("usuario"):
                messagebox.showwarning("ViscondeApp", "Preencha a senha da Mobyan (é a primeira vez que essa plataforma é configurada).")
                return
            plataformas["MOBYAN"] = {
                "url": url_mobyan,
                "usuario": usuario_mobyan,
                "senha": senha_mobyan,
                "ativo": True,
            }
        elif url_mobyan or usuario_mobyan or base_mobyan or estado_mobyan:
            messagebox.showwarning(
                "ViscondeApp",
                "Preencha URL, usuário, base e estado da Mobyan (ou deixe tudo em branco).",
            )
            return

        config_extra = {}
        if base_ogea:
            config_extra["OGEA_PRESTADOR"] = base_ogea
        if base_mobyan:
            config_extra["MOBYAN_PRESTADORES"] = base_mobyan
        if estado_mobyan:
            config_extra["MOBYAN_ESTADO"] = estado_mobyan

        try:
            salvar_config(self.token, plataformas, config_extra)
            config_conta = obter_config(self.token)
            materializar_env(config_conta)
        except ConexaoBackendError:
            messagebox.showerror(
                "ViscondeApp",
                "Não consegui falar com o servidor de contas. Verifique sua internet e tente de novo.",
            )
            return
        except ErroAutenticacao as erro:
            messagebox.showerror("ViscondeApp", str(erro))
            return

        messagebox.showinfo("ViscondeApp", "Configurações salvas. As próximas automações já usam os novos dados.")
        if self.ao_salvar:
            self.ao_salvar()
        self.root.destroy()


def abrir_configuracoes_conta(master, ao_salvar=None):
    """Abre a janela de configurações da conta logada. Retorna a instância da
    janela (ou None se não há sessão local válida ou o backend não respondeu),
    seguindo o mesmo padrão de "nunca travar o app" já usado em
    verificar_atualizacao()."""
    sessao = carregar_sessao_local()
    if not sessao:
        messagebox.showwarning("ViscondeApp", "Faça login novamente para editar as configurações da conta.")
        return None

    try:
        config_conta = obter_config(sessao["token"])
    except ConexaoBackendError:
        messagebox.showerror(
            "ViscondeApp",
            "Não consegui falar com o servidor de contas. Verifique sua internet e tente de novo.",
        )
        return None
    except ErroAutenticacao as erro:
        messagebox.showerror("ViscondeApp", str(erro))
        return None

    return _JanelaConfiguracaoConta(
        master, sessao["token"], sessao.get("usuario", ""), config_conta, ao_salvar=ao_salvar
    )
