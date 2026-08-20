"""
main.py
=======
Ponto de entrada do app. Antes de abrir a interface, checa se tem
atualizaçao publicada no repositorio do GitHub (via github_updater.py) e
aplica sozinho se tiver uma versao mais nova - se nao conseguir checar
(sem internet, token invalido etc.) ele so ignora e abre o app normal.
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent / "app"))

from github_updater import GitHubUpdater, GitHubUpdaterError

import ui


def verificar_atualizacao():
    try:
        updater = GitHubUpdater("update_config.json")
        updater.run(auto_apply=True, silent=True)
    except GitHubUpdaterError as erro:
        print(f"Nao foi possivel checar atualizaçoes: {erro}")


if __name__ == "__main__":
    verificar_atualizacao()
    app = ui.App()
    app.mainloop()
