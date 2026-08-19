from sqlalchemy import create_engine, String, Float, Column, Date
from sqlalchemy.orm import sessionmaker, declarative_base
from config.config import BANCO_DADOS_NAME

db = create_engine(f"sqlite:///{BANCO_DADOS_NAME}")
Session = sessionmaker(bind=db)
Session = Session()

Base = declarative_base()


# ==========================================
#       TABELAS DO BANCO DE DADOS
# ==========================================
# emissao, documento, historico, vencimento, vl. documento, pagamento, valor pago, por, descriçao portador,
# cod.c. custo, conta centro de custo, descricao centro de custo, cod. setor, descricao do setor, observ #1, observ #2, observ #3
class Recebimento:
    __tablename__ = "recebimentos"

    id = Column("id", Integer, primary_key=True, autoincrement=True)
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


class Despesa:
    __tablename__ = "despesas"

    id = Column("id", Integer, primary_key=True, autoincrement=True)
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
