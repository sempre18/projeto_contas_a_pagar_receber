"""
github_updater.py
==================
Modulo generico de auto-atualizacao para apps .exe (backend Python) via
GitHub Releases de repositorio PRIVADO, usando um Personal Access Token
guardado em arquivo JSON de configuracao.

Basta copiar este arquivo pra dentro de qualquer projeto e importar:

    from github_updater import GitHubUpdater

    updater = GitHubUpdater("update_config.json")
    updater.run()  # verifica, baixa e aplica atualizacao se houver uma nova

------------------------------------------------------------------------
Arquivo de config (update_config.json) -- fica ao lado do exe:
------------------------------------------------------------------------
{
    "owner": "seu-usuario-ou-org",
    "repo": "seu-repositorio",
    "token": "github_pat_xxxxxxxxxxxxxxxxxxxxx",
    "current_version": "1.0.0",
    "exe_name": "meuapp.exe",
    "asset_suffix": ".exe"
}

- owner / repo   : dono e nome do repositorio no GitHub.
- token          : Fine-grained PAT com permissao "Contents: Read-only"
                    APENAS nesse repositorio. Nao use um token com mais
                    escopo do que isso.
- current_version: versao instalada agora. O modulo atualiza este campo
                    sozinho apos cada atualizacao aplicada com sucesso.
- exe_name       : nome do executavel a ser substituido (usado quando o
                    script nao esta rodando a partir do proprio exe
                    compilado, ex: durante testes em .py).
- asset_suffix   : sufixo do arquivo anexado na Release que deve ser
                    baixado (padrao ".exe").

------------------------------------------------------------------------
Requisitos:
------------------------------------------------------------------------
    pip install requests

Funciona em Windows (usa um .bat para trocar o exe, ja que o Windows
trava o arquivo enquanto o processo esta rodando).

------------------------------------------------------------------------
AVISO DE SEGURANCA (Opcao A - token embutido):
------------------------------------------------------------------------
O token fica em texto plano no JSON, ao lado do exe. Qualquer pessoa com
acesso ao arquivo consegue ler o codigo-fonte do repositorio com ele.
Use um token com escopo minimo (somente leitura de Contents, somente
nesse repo) e trate o JSON como algo sensivel (nao versionar, nao
compartilhar fora do ambiente interno). Para maior seguranca, considere
migrar futuramente para um backend-proxy que guarda o token no servidor.
"""

from __future__ import annotations

import json
import logging
import os
import subprocess
import sys
import tempfile
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Optional

import requests

logger = logging.getLogger("github_updater")


class GitHubUpdaterError(Exception):
    """Erro generico do modulo de atualizacao."""


@dataclass
class ReleaseInfo:
    version: str
    tag_name: str
    asset_id: int
    asset_name: str
    asset_size: int


def _get_base_dir() -> Path:
    """Retorna a pasta onde o app realmente esta rodando, verificada em
    tempo de execucao. Usada para resolver caminhos relativos (config,
    exe_name) de forma dinamica, em vez de depender do diretorio de
    trabalho atual (cwd), que pode ser diferente (ex: app aberto por
    atalho com 'Iniciar em' distinto, ou por outro processo)."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    main_module = sys.modules.get("__main__")
    main_file = getattr(main_module, "__file__", None)
    if main_file:
        return Path(main_file).resolve().parent
    return Path.cwd()


def _resolve_path(path_str: str, base_dir: Path) -> Path:
    """Resolve path_str: se ja for absoluto, usa como esta; se for
    relativo, resolve em relacao a base_dir (nao ao cwd)."""
    p = Path(path_str)
    if not p.is_absolute():
        p = base_dir / p
    return p.resolve()


def _parse_version(v: str) -> tuple:
    """Converte 'v1.2.3' ou '1.2.3' em (1, 2, 3) para comparacao segura,
    sem depender do pacote 'packaging'."""
    v = v.strip().lstrip("vV")
    parts = []
    for p in v.split("."):
        num = ""
        for ch in p:
            if ch.isdigit():
                num += ch
            else:
                break
        parts.append(int(num) if num else 0)
    while len(parts) < 3:
        parts.append(0)
    return tuple(parts)


class GitHubUpdater:
    def __init__(self, config_path: str = "update_config.json"):
        self.base_dir = _get_base_dir()
        self.config_path = _resolve_path(config_path, self.base_dir)
        self.config = self._load_config()

    # ---------------- config ----------------

    def _load_config(self) -> dict:
        if not self.config_path.exists():
            raise GitHubUpdaterError(f"Arquivo de config nao encontrado: {self.config_path}")
        with open(self.config_path, "r", encoding="utf-8") as f:
            data = json.load(f)
        for campo in ("owner", "repo", "token", "current_version"):
            if campo not in data or not data[campo]:
                raise GitHubUpdaterError(f"Campo obrigatorio ausente/vazio no config: '{campo}'")
        data.setdefault("exe_name", None)
        data.setdefault("asset_suffix", ".exe")
        return data

    def _save_config(self):
        with open(self.config_path, "w", encoding="utf-8") as f:
            json.dump(self.config, f, indent=2, ensure_ascii=False)

    # ---------------- API do GitHub ----------------

    def _headers(self, accept: str = "application/vnd.github+json") -> dict:
        return {
            "Authorization": f"Bearer {self.config['token']}",
            "Accept": accept,
            "X-GitHub-Api-Version": "2022-11-28",
        }

    def check_update(self) -> Optional[ReleaseInfo]:
        """Consulta a release mais recente no GitHub e retorna info se for
        mais nova que 'current_version' do config. Retorna None se ja
        esta atualizado."""
        url = f"https://api.github.com/repos/{self.config['owner']}/{self.config['repo']}/releases/latest"
        try:
            r = requests.get(url, headers=self._headers(), timeout=15)
        except requests.RequestException as e:
            raise GitHubUpdaterError(f"Falha ao conectar no GitHub: {e}") from e

        if r.status_code == 401:
            raise GitHubUpdaterError("Token invalido, expirado ou sem permissao (401).")
        if r.status_code == 404:
            raise GitHubUpdaterError(
                "Repositorio ou release nao encontrado (404). Verifique owner/repo/token/permissoes."
            )
        r.raise_for_status()

        data = r.json()
        tag = data.get("tag_name", "")
        remote_version = tag.lstrip("vV")

        assets = data.get("assets", [])
        suffix = self.config["asset_suffix"]
        asset = next((a for a in assets if a["name"].endswith(suffix)), None)
        if asset is None:
            raise GitHubUpdaterError(f"Nenhum asset terminando em '{suffix}' encontrado na release {tag}.")

        atual = _parse_version(self.config["current_version"])
        remoto = _parse_version(remote_version)

        logger.info("Versao atual: %s | Versao remota: %s", self.config["current_version"], remote_version)

        if remoto <= atual:
            return None

        return ReleaseInfo(
            version=remote_version,
            tag_name=tag,
            asset_id=asset["id"],
            asset_name=asset["name"],
            asset_size=asset.get("size", 0),
        )

    # ---------------- download ----------------

    def download(
        self,
        release: ReleaseInfo,
        progress_cb: Optional[Callable[[int, int], None]] = None,
    ) -> Path:
        """Baixa o asset da release para uma pasta temporaria e retorna o
        caminho do arquivo baixado."""
        url = (
            f"https://api.github.com/repos/{self.config['owner']}/{self.config['repo']}"
            f"/releases/assets/{release.asset_id}"
        )
        destino = Path(tempfile.gettempdir()) / release.asset_name
        headers = self._headers(accept="application/octet-stream")

        try:
            with requests.get(url, headers=headers, stream=True, timeout=60) as r:
                r.raise_for_status()
                baixado = 0
                with open(destino, "wb") as f:
                    for chunk in r.iter_content(chunk_size=8192):
                        if chunk:
                            f.write(chunk)
                            baixado += len(chunk)
                            if progress_cb:
                                progress_cb(baixado, release.asset_size)
        except requests.RequestException as e:
            raise GitHubUpdaterError(f"Falha ao baixar atualizacao: {e}") from e

        return destino

    # ---------------- aplicacao ----------------

    def apply(
        self,
        downloaded_path: Path,
        release: ReleaseInfo,
        exe_path: Optional[str] = None,
    ):
        """Gera um .bat que espera este processo encerrar, faz backup do
        exe atual, coloca o novo no lugar, reabre o app e se autodestroi.
        Atualiza current_version no config ANTES de reiniciar, e finaliza
        o processo atual (necessario pra liberar o lock do arquivo)."""

        alvo_raw = exe_path or self.config.get("exe_name") or (
            sys.executable if getattr(sys, "frozen", False) else None
        )
        if not alvo_raw:
            raise GitHubUpdaterError(
                "Nao foi possivel determinar o exe alvo. Rode a partir do exe "
                "compilado, ou informe 'exe_name' no config, ou passe exe_path=... no apply()."
            )
        alvo = _resolve_path(alvo_raw, self.base_dir)

        pid = os.getpid()
        bat_path = Path(tempfile.gettempdir()) / f"update_{pid}.bat"
        backup_path = alvo.with_suffix(alvo.suffix + ".bak")

        # Atualiza a versao no config ANTES de reiniciar: quando o exe novo
        # subir, ele ja le o JSON com a versao correta.
        self.config["current_version"] = release.version
        self._save_config()

        script = f"""@echo off
:loop
tasklist /FI "PID eq {pid}" 2>NUL | find "{pid}" >NUL
if "%ERRORLEVEL%"=="0" (
    timeout /t 1 /nobreak >NUL
    goto loop
)
if exist "{backup_path}" del /F /Q "{backup_path}"
if exist "{alvo}" move /Y "{alvo}" "{backup_path}" >NUL
move /Y "{downloaded_path}" "{alvo}" >NUL
start "" "{alvo}"
del "%~f0"
"""
        with open(bat_path, "w", encoding="utf-8") as f:
            f.write(script)

        logger.info("Atualizando para versao %s. Reiniciando aplicativo...", release.version)
        subprocess.Popen(["cmd", "/c", str(bat_path)], creationflags=subprocess.CREATE_NO_WINDOW)
        sys.exit(0)

    # ---------------- fluxo completo ----------------

    def run(
        self,
        auto_apply: bool = True,
        silent: bool = True,
        exe_path: Optional[str] = None,
        on_update_found: Optional[Callable[[ReleaseInfo], bool]] = None,
        progress_cb: Optional[Callable[[int, int], None]] = None,
    ) -> bool:
        """Fluxo completo: verifica -> (pergunta opcional) -> baixa -> aplica.

        on_update_found: callback que recebe o ReleaseInfo encontrado e
            retorna True/False dizendo se deve prosseguir com a atualizacao.
            Util pra mostrar um dialogo tipo "Nova versao 1.3.0 disponivel,
            atualizar agora?" antes de baixar. Se None, usa 'auto_apply'.

        progress_cb: callback(bytes_baixados, total_bytes) chamado durante
            o download, util pra barra de progresso.

        Retorna False se nao havia atualizacao ou se o usuario recusou.
        Se uma atualizacao for aplicada, o processo atual e encerrado
        dentro de apply() (sys.exit), entao normalmente nao ha retorno.
        """
        try:
            release = self.check_update()
        except GitHubUpdaterError as e:
            if not silent:
                raise
            logger.warning("Verificacao de atualizacao falhou: %s", e)
            return False

        if release is None:
            logger.info("Aplicativo ja esta na versao mais recente.")
            return False

        prosseguir = on_update_found(release) if on_update_found is not None else auto_apply
        if not prosseguir:
            return False

        caminho = self.download(release, progress_cb=progress_cb)
        self.apply(caminho, release, exe_path=exe_path)
        return True
