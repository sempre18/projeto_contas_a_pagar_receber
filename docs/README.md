# PROJETO CONTAS A PAGAR E CONTAS A RECEBER

## Conceito do projeto:
este é um simples projeto para a criaçao de um dashboard com contas a pagar e a vender,
nele você tem uma simples interface onde vc importa as planilhas de projeção e de recebimento,
seleciona o periodo e escolhe o lugar de saida, entao ele gera uma planilha excel(.xlsx) com 
o dashboard e o banco de dados

## Estrutura
    app
     |config
     |  |config.py
     |      L> tem as configurações padrão do sistema
     |data
     |  L>modelo inicial do Dashboard e algumas planilhas para testes(sempre que quiser reiniciar o teste exclua o banco de dados em:app\database\RECEBIMENTOS_DESPESAS.db)
     |
     |database
     |  |RECEBIMENTOS_DESPESAS.db
     |      L> Banco de Dados
     |models
     |  |dashboard.py
     |  |   L> Cria o dashboard
     |  |sql.py
     |      L>tem os modelos das tabelas SQL e todo o CRUD do projeto
     |ui.py
        L>cria a interface

### Fora da pasta app
    docs
     |README.md
     |  L> este arquivo
     |requirements.txt
     |  L> bibliotecas usadas no projeto para facilitar a instalaçao
    main.py
     L> ponto de entrada do app: antes de abrir a interface, checa se tem
        atualizaçao nova no GitHub (via github_updater.py) e aplica sozinho
    github_updater.py
     L> biblioteca (arquivo unico) de auto-atualizaçao via GitHub Releases
        de um repositorio privado
    update_config.json
     L> config da auto-atualizaçao (owner, repo, token, versao atual) -
        NÃO é versionado (esta no .gitignore) porque tem o token real
    update_config.example.json
     L> mesmo arquivo acima mas sem token real, só pra mostrar o formato
        pra quem for mexer no projeto

## Stacks
- Python
- Tkinter + tkcalendar - interface
- SQLAlchemy + SQLite - banco de dados
- pandas + openpyxl - leitura/escrita das planilhas .xlsx e montagem do dashboard
- requests - comunicaçao com a API do GitHub na auto-atualizaçao
- PyInstaller - empacotar o app em .exe pra distribuir