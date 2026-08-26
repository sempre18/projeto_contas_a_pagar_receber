import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import hashlib

import pandas as pd
from sqlalchemy import create_engine, String, Float, Column, Date, Integer
from sqlalchemy import func
from sqlalchemy.orm import sessionmaker, declarative_base
from config.config import (
    BANCO_DADOS_NAME,
    PROJECT_DIR,
    COLUNAS_DESPESA,
    COLUNAS_RECEBIMENTO,
)

PASTA_DB = PROJECT_DIR / "database"
PASTA_DB.mkdir(exist_ok=True)

db = create_engine(f"sqlite:///{str(PASTA_DB / BANCO_DADOS_NAME)}")
Session = sessionmaker(bind=db)
Session = Session()

Base = declarative_base()


# ==========================================
#       TABELAS DO BANCO DE DADOS
# ==========================================
# emissao, documento, historico, vencimento, vl. documento, pagamento, valor pago, por, descriçao portador,
# cod.c. custo, conta centro de custo, descricao centro de custo, cod. setor, descricao do setor, observ #1, observ #2, observ #3
class Recebimento(Base):
    __tablename__ = "recebimentos"

    id = Column("id", Integer, primary_key=True, autoincrement=True)
    _hash = Column("_hash", String, unique=True, index=True)
    emissao = Column("emissao", Date)
    documento = Column("documento", String)
    historico = Column("historico", String)
    vencimento = Column("vencimento", Date)
    vl_documento = Column("vl_documento", Float)
    pagamento = Column("pagamento", String)
    valor_pago = Column("valor_pago", Float)
    por = Column("por", String)
    descriçao_portador = Column("descriçao_portador", String)
    cod_c_custo = Column("cod_c_custo", String)
    conta_centro_custo = Column("conta_centro_custo", String)
    descricao_centro_custo = Column("descricao_centro_custo", String)
    cod_setor = Column("cod_setor", String)
    descricao_setor = Column("descricao_setor", String)
    observ_1 = Column("observ_1", String)
    observ_2 = Column("observ_2", String)
    observ_3 = Column("observ_3", String)

    def __init__(
        self,
        emissao,
        documento,
        historico,
        vencimento,
        vl_documento,
        pagamento,
        valor_pago,
        por,
        descriçao_portador,
        cod_c_custo,
        conta_centro_custo,
        descricao_centro_custo,
        cod_setor,
        descricao_setor,
        observ_1,
        observ_2,
        observ_3,
    ):
        self.emissao = emissao
        self.documento = documento
        self.historico = historico
        self.vencimento = vencimento
        self.vl_documento = vl_documento
        self.pagamento = pagamento
        self.valor_pago = valor_pago
        self.por = por
        self.descriçao_portador = descriçao_portador
        self.cod_c_custo = cod_c_custo
        self.conta_centro_custo = conta_centro_custo
        self.descricao_centro_custo = descricao_centro_custo
        self.cod_setor = cod_setor
        self.descricao_setor = descricao_setor
        self.observ_1 = observ_1
        self.observ_2 = observ_2
        self.observ_3 = observ_3


# sis, emissao, documento, historico, vencimento, valor, pagamento, valor_pago, desconto, por, descrição_portador,
# num_bancario, observação_1, valor_juros, avisos, nome_vendedor, c_cli, cnpj_cliente, telefone_cliente, data_cob, obs_dani, comissao_paga


class Despesa(Base):
    __tablename__ = "despesas"

    id = Column("id", Integer, primary_key=True, autoincrement=True)
    _hash = Column("_hash", String, unique=True, index=True)
    sis = Column("sis", String)
    emissao = Column("emissao", Date)
    documento = Column("documento", String)
    historico = Column("historico", String)
    vencimento = Column("vencimento", Date)
    valor = Column("valor", Float)
    pagamento = Column("pagamento", Date)
    valor_pago = Column("valor_pago", Float)
    desconto = Column("desconto", Float)
    por = Column("por", String)
    descrição_portador = Column("descrição_portador", String)
    num_bancario = Column("num_bancario", String)
    observação_1 = Column("observação_1", String)
    valor_juros = Column("valor_juros", Float)
    avisos = Column("avisos", String)
    nome_vendedor = Column("nome_vendedor", String)
    c_cli = Column("c_cli", String)
    cnpj_cliente = Column("cnpj_cliente", String)
    telefone_cliente = Column("telefone_cliente", String)
    data_cob = Column("data_cob", String)
    obs_dani = Column("obs_dani", String)
    comissao_paga = Column("comissao_paga", String)

    def __init__(
        self,
        sis,
        emissao,
        documento,
        historico,
        vencimento,
        valor,
        pagamento,
        valor_pago,
        desconto,
        por,
        descrição_portador,
        num_bancario,
        observação_1,
        valor_juros,
        avisos,
        nome_vendedor,
        c_cli,
        cnpj_cliente,
        telefone_cliente,
        data_cob,
        obs_dani,
        comissao_paga,
    ):
        self.sis = sis
        self.emissao = emissao
        self.documento = documento
        self.historico = historico
        self.vencimento = vencimento
        self.valor = valor
        self.pagamento = pagamento
        self.valor_pago = valor_pago
        self.desconto = desconto
        self.por = por
        self.descrição_portador = descrição_portador
        self.num_bancario = num_bancario
        self.observação_1 = observação_1
        self.valor_juros = valor_juros
        self.avisos = avisos
        self.nome_vendedor = nome_vendedor
        self.c_cli = c_cli
        self.cnpj_cliente = cnpj_cliente
        self.telefone_cliente = telefone_cliente
        self.data_cob = data_cob
        self.obs_dani = obs_dani
        self.comissao_paga = comissao_paga


Base.metadata.create_all(bind=db)


def data_maxima_atualizada():
    data_receb = Session.query(func.max(Recebimento.emissao)).scalar()
    data_desp = Session.query(func.max(Despesa.emissao)).scalar()
    datas = [d for d in (data_receb, data_desp) if d is not None]
    return max(datas) if datas else None


# ==========================================
#       IMPORTAÇÃO DE ARQUIVOS .XLSX
# ==========================================

TABELA_CONFIG = {
    "COBRANÇA": {
        "model": Despesa,
        "colunas_map": COLUNAS_DESPESA,
        "colunas": list(dict.fromkeys(COLUNAS_DESPESA.values())),
        "colunas_data": ["emissao", "vencimento", "pagamento"],
        "colunas_float": ["valor", "valor_pago", "desconto", "valor_juros"],
    },
    "PROJEÇÃO": {
        "model": Recebimento,
        "colunas_map": COLUNAS_RECEBIMENTO,
        "colunas": list(dict.fromkeys(COLUNAS_RECEBIMENTO.values())),
        "colunas_data": ["emissao", "vencimento"],
        "colunas_float": ["vl_documento", "valor_pago"],
    },
}


def _hash_linha(valores):
    """Hash estável de uma linha (mesmos valores em todas as colunas -> mesmo hash), usado para não importar a mesma linha duas vezes."""
    texto = "|".join("" if v is None or pd.isna(v) else str(v) for v in valores)
    return hashlib.sha256(texto.encode("utf-8")).hexdigest()


def _hashes_existentes(model_class):
    return {h for (h,) in Session.query(model_class._hash).all()}


def _importar_dataframe(model_class, df, colunas):
    df = df.copy()
    total_arquivo = len(df)
    df["_hash"] = df[colunas].apply(lambda linha: _hash_linha(linha.tolist()), axis=1)
    df = df.drop_duplicates(
        subset="_hash"
    )  # linhas idênticas dentro do próprio arquivo

    existentes = _hashes_existentes(model_class)
    novas = df[~df["_hash"].isin(existentes)]

    if len(novas) > 0:
        Session.bulk_insert_mappings(
            model_class,
            novas[colunas + ["_hash"]].to_dict("records"),
            render_nulls=True,  # sem isso, colunas None na 1a linha do lote somem do INSERT inteiro
        )
        Session.commit()

    return {
        "total_arquivo": total_arquivo,
        "novas": len(novas),
        "duplicadas": total_arquivo - len(novas),
    }


def listar_abas_compativeis(arquivo, caminho):
    """Lista as abas de `caminho` cujas colunas batem com o formato esperado de `arquivo` (um valor de TABELAS)."""
    config = TABELA_CONFIG.get(arquivo)
    if config is None:
        raise ValueError(
            f"Tabela desconhecida: {arquivo!r}. Esperado um de {list(TABELA_CONFIG)}"
        )
    mapa_colunas = config["colunas_map"]

    xls = pd.ExcelFile(caminho)
    compativeis = []
    for aba in xls.sheet_names:
        colunas_aba = pd.read_excel(xls, sheet_name=aba, nrows=0).columns
        if any(c in mapa_colunas for c in colunas_aba):
            compativeis.append(aba)
    return compativeis


def importar_arquivo(arquivo, caminho, abas=None):
    """
    abas: lista das abas a importar (deve vir de listar_abas_compativeis).
    Se None, usa todas as abas do arquivo que baterem com o formato esperado.
    """
    config = TABELA_CONFIG.get(arquivo)
    if config is None:
        raise ValueError(
            f"Tabela desconhecida: {arquivo!r}. Esperado um de {list(TABELA_CONFIG)}"
        )

    mapa_colunas = config["colunas_map"]
    colunas = config["colunas"]

    xls = pd.ExcelFile(caminho)
    abas_a_processar = abas if abas is not None else xls.sheet_names

    partes = []
    for aba in abas_a_processar:
        if aba not in xls.sheet_names:
            raise ValueError(f"A aba '{aba}' não existe em '{caminho}'.")
        df_aba = pd.read_excel(xls, sheet_name=aba)
        colunas_presentes = [c for c in df_aba.columns if c in mapa_colunas]
        if not colunas_presentes:
            continue
        df_aba = df_aba[colunas_presentes].rename(columns=mapa_colunas)
        partes.append(df_aba.reindex(columns=colunas))

    if not partes:
        raise ValueError(
            f"Nenhuma das abas selecionadas de '{caminho}' bate com o formato esperado de '{arquivo}'."
        )

    df = pd.concat(partes, ignore_index=True)

    for coluna in config["colunas_data"]:
        df[coluna] = pd.to_datetime(df[coluna], errors="coerce")
        df[coluna] = df[coluna].apply(lambda v: v.date() if pd.notna(v) else None)
    for coluna in config["colunas_float"]:
        df[coluna] = pd.to_numeric(df[coluna], errors="coerce")
    df = df.where(pd.notnull(df), None)

    return _importar_dataframe(config["model"], df, colunas)


if __name__ == "__main__":
    pass
