from pathlib import Path

 # tabelas que serao ultilizadas
TABELAS = ["COBRANÇA", "PROJEÇÃO"]

 # nome do banc de dados
BANCO_DADOS_NAME = "RECEBIMENTOS_DESPESAS.db"

 # caminho completo do projeto
PROJECT_DIR = Path(__file__).resolve().parent.parent


# conexões planilha/banco de dados
COLUNAS_DESPESA = {
    "SIS": "sis",
    "EMISSAO": "emissao",
    "Documento": "documento",
    "Historico": "historico",
    "Vencimento": "vencimento",
    "Valor": "valor",
    "Pagamento": "pagamento",
    "Valor Pago": "valor_pago",
    "Desconto": "desconto",
    "Por": "por",
    "Descrição Portador": "descrição_portador",
    "N° BANCARIO": "num_bancario",
    "Observação #1": "observação_1",
    "Valor Juros": "valor_juros",
    "Avisos": "avisos",
    "Nome Vendedor": "nome_vendedor",
    "C.Cli": "c_cli",
    "CNPJ do Cliente": "cnpj_cliente",
    "Telefone Cliente": "telefone_cliente",
    "DATA COB": "data_cob",
    "OBS DANI ": "obs_dani",
    "COMISSAO PAGA": "comissao_paga",
}

COLUNAS_RECEBIMENTO = {
    "Emissão": "emissao",
    "Documento": "documento",
    "Histórico": "historico",
    "Vencimento": "vencimento",
    "Vl.Documento": "vl_documento",
    "Pagamento": "pagamento",
    "Valor Pago": "valor_pago",
    "Por": "por",
    "Descrição Portador": "descriçao_portador",
    "Cód.C.Custo": "cod_c_custo",
    "Conta Centro de Custo": "conta_centro_custo",
    "Descrição Centro de Custo": "descricao_centro_custo",
    "Cód.Setor": "cod_setor",
    "Setor": "descricao_setor",
    "Descrição do Setor": "descricao_setor",
    "Observ #1": "observ_1",
    "Observ #2": "observ_2",
    "Observ #3": "observ_3",
}
