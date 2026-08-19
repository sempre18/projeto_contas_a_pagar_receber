"""
Interface gráfica (Tkinter) para o sistema CUSTO COMPLETO:
- importa cada base (.xlsx) para o SQLite (dados_custo.db), sem duplicar linhas;
- gera a planilha de saída para um período escolhido, com as abas brutas e o
  dashboard como opcionais (checkboxes).

Uso:
    python interface.py
"""

import os
import tkinter as tk
from datetime import date
from tkinter import filedialog, messagebox, scrolledtext, ttk

import pandas as pd
from tkcalendar import DateEntry

import banco_dados
import dashboard
import emissao
from custo_completo import (
    aplicar_validacoes_categoria,
    criar_tabela_custo_completo,
    formatar_planilha_custo_completo,
    montar_custo_completo,
)

TABELAS = ["SISREV", "FSIST", "CTE", "COMPRA COMPLETO", "BAIXA ESPECIAL", "FORNECEDORES", "FATURAMENTO"]

# Tabelas filtradas pelo período escolhido (as que definem quais notas/emissões
# existem). As demais (CTE, COMPRA COMPLETO, BAIXA ESPECIAL, FORNECEDORES) são
# sempre carregadas por inteiro, como referência para os cruzamentos.
TABELAS_FILTRADAS_POR_PERIODO = {"SISREV", "FSIST", "FATURAMENTO"}


class Aplicativo:
    def __init__(self, raiz):
        self.raiz = raiz
        raiz.title("CUSTO COMPLETO")

        self.caminho_saida = tk.StringVar()
        self.incluir_bases = tk.BooleanVar(value=True)
        self.incluir_dashboard = tk.BooleanVar(value=True)
        self.incluir_emissao = tk.BooleanVar(value=True)
        self.labels_status = {}

        self._montar_secao_status()

        notebook = ttk.Notebook(self.raiz)
        notebook.pack(fill="both", expand=True, padx=10, pady=(5, 10))
        aba_principal = ttk.Frame(notebook)
        aba_categorias = ttk.Frame(notebook)
        notebook.add(aba_principal, text="Principal")
        notebook.add(aba_categorias, text="Categorias")

        self._montar_secao_importar(aba_principal)
        self._montar_secao_alteracoes(aba_principal)
        self._montar_secao_periodo(aba_principal)
        self._montar_secao_saida(aba_principal)
        self._montar_secao_gerar(aba_principal)

        self._montar_aba_categorias(aba_categorias)

    # ------------------------------------------------------------------
    # Seção "Dados atualizados até" (data mais recente já importada)
    # ------------------------------------------------------------------
    def _montar_secao_status(self):
        moldura = ttk.Frame(self.raiz)
        moldura.pack(fill="x", padx=10, pady=(10, 0))
        self.label_data_atualizada = ttk.Label(moldura, font=("TkDefaultFont", 10, "bold"))
        self.label_data_atualizada.pack(anchor="w")
        self._atualizar_label_data_atualizada()

    def _atualizar_label_data_atualizada(self):
        data_maxima = banco_dados.data_maxima_atualizada()
        if data_maxima is None:
            texto = "Dados atualizados até: (nenhum dado importado ainda)"
        else:
            texto = f"Dados atualizados até: {data_maxima.strftime('%d/%m/%Y')}"
        self.label_data_atualizada.config(text=texto)

    # ------------------------------------------------------------------
    # Seção "Importar bases"
    # ------------------------------------------------------------------
    def _montar_secao_importar(self, pai):
        moldura = ttk.LabelFrame(pai, text="Importar bases (.xlsx) para o banco de dados")
        moldura.pack(fill="x", padx=10, pady=10)

        for linha, nome_tabela in enumerate(TABELAS):
            ttk.Label(moldura, text=nome_tabela, width=18).grid(row=linha, column=0, padx=5, pady=3, sticky="w")

            botao = ttk.Button(
                moldura, text="Selecionar arquivo...",
                command=lambda n=nome_tabela: self._selecionar_e_importar(n),
            )
            botao.grid(row=linha, column=1, padx=5, pady=3)

            label_status = ttk.Label(moldura, text="Nenhum arquivo importado ainda", foreground="gray")
            label_status.grid(row=linha, column=2, padx=5, pady=3, sticky="w")
            self.labels_status[nome_tabela] = label_status

    def _selecionar_e_importar(self, nome_tabela):
        caminhos = filedialog.askopenfilenames(
            title=f"Selecionar arquivo(s) .xlsx para {nome_tabela}",
            filetypes=[("Planilhas Excel", "*.xlsx")],
        )
        if not caminhos:
            return

        total_arquivo = novas = duplicadas = arquivos_ok = 0
        erros = []
        for caminho in caminhos:
            try:
                resultado = banco_dados.importar_arquivo(nome_tabela, caminho)
            except Exception as erro:
                erros.append(f"{caminho}:\n{erro}")
                continue
            arquivos_ok += 1
            total_arquivo += resultado["total_arquivo"]
            novas += resultado["novas"]
            duplicadas += resultado["duplicadas"]

        if erros:
            messagebox.showerror(
                "Erro ao importar",
                f"Falha ao importar {len(erros)} de {len(caminhos)} arquivo(s) para '{nome_tabela}':\n\n"
                + "\n\n".join(erros),
            )

        if arquivos_ok:
            self.labels_status[nome_tabela].config(
                text=(
                    f"{arquivos_ok} arquivo(s) importado(s) — {total_arquivo} linhas — "
                    f"{novas} novas, {duplicadas} duplicadas ignoradas"
                ),
                foreground="black",
            )
            self._atualizar_label_data_atualizada()

    # ------------------------------------------------------------------
    # Seção "Atualizar alterações": reimporta uma planilha CUSTO COMPLETO já
    # gerada (e possivelmente editada à mão - Categoria, Cliente, Fornecedor,
    # valores, o que for) e persiste toda célula que estiver diferente do
    # que sairia numa geração nova, pra reaproveitar na próxima geração em
    # vez de perder a edição.
    # ------------------------------------------------------------------
    def _montar_secao_alteracoes(self, pai):
        moldura = ttk.LabelFrame(pai, text="Atualizar alterações (a partir de um CUSTO COMPLETO editado)")
        moldura.pack(fill="x", padx=10, pady=5)

        ttk.Label(moldura, text="CUSTO COMPLETO", width=18).grid(row=0, column=0, padx=5, pady=3, sticky="w")
        ttk.Button(
            moldura, text="Selecionar arquivo...", command=self._selecionar_e_importar_alteracoes,
        ).grid(row=0, column=1, padx=5, pady=3)

        self.label_status_categorias = ttk.Label(
            moldura, text="Nenhuma alteração importada ainda", foreground="gray",
        )
        self.label_status_categorias.grid(row=0, column=2, padx=5, pady=3, sticky="w")

    def _selecionar_e_importar_alteracoes(self):
        caminho = filedialog.askopenfilename(
            title="Selecionar planilha CUSTO COMPLETO (gerada por este programa) já editada",
            filetypes=[("Planilhas Excel", "*.xlsx")],
        )
        if not caminho:
            return

        try:
            quantidade = banco_dados.importar_alteracoes_custo_completo(caminho)
        except Exception as erro:
            messagebox.showerror("Erro ao importar alterações", str(erro))
            return

        if quantidade == 0:
            messagebox.showwarning(
                "Nada encontrado",
                "Não encontrei a aba 'CUSTO COMPLETO' nesse arquivo (ou as colunas não batem com o esperado).",
            )
        self.label_status_categorias.config(
            text=f"{quantidade} célula(s) salva(s) — serão reaproveitadas na próxima geração",
            foreground="black",
        )

    # ------------------------------------------------------------------
    # Aba "Categorias": gerenciar (adicionar/renomear/excluir) as listas de
    # Categoria e Sub- Categoria usadas no dropdown do CUSTO COMPLETO -
    # persistidas em banco_dados (tabelas "categorias"/"subcategorias"), não
    # são mais fixas no código.
    # ------------------------------------------------------------------
    def _montar_aba_categorias(self, pai):
        moldura_categoria = self._montar_editor_lista(
            pai, "Categoria",
            banco_dados.listar_categorias, banco_dados.adicionar_categoria,
            banco_dados.renomear_categoria, banco_dados.excluir_categoria,
        )
        moldura_categoria.pack(fill="both", expand=True, padx=10, pady=10, side="left")

        moldura_subcategoria = self._montar_editor_lista(
            pai, "Sub- Categoria",
            banco_dados.listar_subcategorias, banco_dados.adicionar_subcategoria,
            banco_dados.renomear_subcategoria, banco_dados.excluir_subcategoria,
        )
        moldura_subcategoria.pack(fill="both", expand=True, padx=10, pady=10, side="left")

    def _montar_editor_lista(self, pai, titulo, listar_fn, adicionar_fn, renomear_fn, excluir_fn):
        """Painel genérico de CRUD (adicionar/renomear/excluir) pra uma lista simples de texto."""
        moldura = ttk.LabelFrame(pai, text=titulo)

        lista = tk.Listbox(moldura, height=14, exportselection=False)
        lista.pack(fill="both", expand=True, padx=5, pady=5)

        def recarregar():
            lista.delete(0, "end")
            for valor in listar_fn():
                lista.insert("end", valor)

        entrada = tk.StringVar()

        def selecionar(_evento=None):
            selecao = lista.curselection()
            if selecao:
                entrada.set(lista.get(selecao[0]))
        lista.bind("<<ListboxSelect>>", selecionar)

        linha_entrada = ttk.Frame(moldura)
        linha_entrada.pack(fill="x", padx=5, pady=(0, 5))
        ttk.Entry(linha_entrada, textvariable=entrada).pack(side="left", fill="x", expand=True)

        def adicionar():
            valor = entrada.get().strip()
            if not valor:
                return
            adicionar_fn(valor)
            entrada.set("")
            recarregar()

        def renomear():
            selecao = lista.curselection()
            novo_valor = entrada.get().strip()
            if not selecao or not novo_valor:
                messagebox.showwarning(
                    "Selecione um item", "Selecione um item da lista e digite o novo nome no campo.",
                )
                return
            renomear_fn(lista.get(selecao[0]), novo_valor)
            recarregar()

        def excluir():
            selecao = lista.curselection()
            if not selecao:
                messagebox.showwarning("Selecione um item", "Selecione um item da lista pra excluir.")
                return
            valor = lista.get(selecao[0])
            if messagebox.askyesno("Excluir", f"Excluir '{valor}'? Notas que já usam esse valor não são afetadas."):
                excluir_fn(valor)
                entrada.set("")
                recarregar()

        linha_botoes = ttk.Frame(moldura)
        linha_botoes.pack(fill="x", padx=5, pady=(0, 5))
        ttk.Button(linha_botoes, text="Adicionar", command=adicionar).pack(side="left", padx=2)
        ttk.Button(linha_botoes, text="Renomear selecionado", command=renomear).pack(side="left", padx=2)
        ttk.Button(linha_botoes, text="Excluir selecionado", command=excluir).pack(side="left", padx=2)

        recarregar()
        return moldura

    # ------------------------------------------------------------------
    # Seção "Período"
    # ------------------------------------------------------------------
    def _montar_secao_periodo(self, pai):
        moldura = ttk.LabelFrame(pai, text="Período (define o que entra na planilha de saída)")
        moldura.pack(fill="x", padx=10, pady=5)

        hoje = date.today()
        inicio_mes = hoje.replace(day=1)

        ttk.Label(moldura, text="Data início:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.data_inicio_widget = DateEntry(
            moldura, date_pattern="dd/mm/yyyy", locale="pt_BR",
            year=inicio_mes.year, month=inicio_mes.month, day=inicio_mes.day,
        )
        self.data_inicio_widget.grid(row=0, column=1, padx=5, pady=5)

        ttk.Label(moldura, text="Data fim:").grid(row=0, column=2, padx=5, pady=5, sticky="w")
        self.data_fim_widget = DateEntry(moldura, date_pattern="dd/mm/yyyy", locale="pt_BR")
        self.data_fim_widget.grid(row=0, column=3, padx=5, pady=5)

    # ------------------------------------------------------------------
    # Seção "Saída"
    # ------------------------------------------------------------------
    def _montar_secao_saida(self, pai):
        moldura = ttk.LabelFrame(pai, text="Saída")
        moldura.pack(fill="x", padx=10, pady=5)

        ttk.Label(moldura, text="Salvar planilha em:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        ttk.Entry(moldura, textvariable=self.caminho_saida, width=50).grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(moldura, text="Salvar como...", command=self._escolher_caminho_saida).grid(
            row=0, column=2, padx=5, pady=5
        )

        ttk.Checkbutton(
            moldura, text="Incluir abas dos bancos de dados", variable=self.incluir_bases
        ).grid(row=1, column=0, columnspan=2, padx=5, pady=3, sticky="w")
        ttk.Checkbutton(
            moldura, text="Incluir dashboard", variable=self.incluir_dashboard
        ).grid(row=2, column=0, columnspan=2, padx=5, pady=3, sticky="w")
        ttk.Checkbutton(
            moldura, text="Incluir tabela de emissão", variable=self.incluir_emissao
        ).grid(row=3, column=0, columnspan=2, padx=5, pady=3, sticky="w")

    def _escolher_caminho_saida(self):
        caminho = filedialog.asksaveasfilename(
            title="Salvar planilha de saída",
            defaultextension=".xlsx",
            filetypes=[("Planilha Excel", "*.xlsx")],
        )
        if caminho:
            self.caminho_saida.set(caminho)

    # ------------------------------------------------------------------
    # Seção "Gerar planilha" + log
    # ------------------------------------------------------------------
    def _montar_secao_gerar(self, pai):
        ttk.Button(pai, text="Gerar planilha", command=self._gerar_planilha).pack(pady=5)

        moldura = ttk.LabelFrame(pai, text="Log")
        moldura.pack(fill="both", expand=True, padx=10, pady=10)
        self.texto_log = scrolledtext.ScrolledText(moldura, height=10, state="disabled")
        self.texto_log.pack(fill="both", expand=True, padx=5, pady=5)

    def _registrar_log(self, mensagem):
        self.texto_log.config(state="normal")
        self.texto_log.insert("end", mensagem + "\n")
        self.texto_log.see("end")
        self.texto_log.config(state="disabled")

    def _gerar_planilha(self):
        caminho_saida = self.caminho_saida.get().strip()
        if not caminho_saida:
            messagebox.showerror("Caminho de saída faltando", "Escolha onde salvar a planilha de saída.")
            return

        data_inicio = self.data_inicio_widget.get_date()
        data_fim = self.data_fim_widget.get_date()

        if data_inicio > data_fim:
            messagebox.showerror("Período inválido", "Data início não pode ser depois da data fim.")
            return

        try:
            planilhas = {}
            for nome_tabela in TABELAS:
                if nome_tabela in TABELAS_FILTRADAS_POR_PERIODO:
                    planilhas[nome_tabela] = banco_dados.carregar_tabela(nome_tabela, data_inicio, data_fim)
                else:
                    planilhas[nome_tabela] = banco_dados.carregar_tabela(nome_tabela)

            alteracoes_manuais = banco_dados.carregar_alteracoes_manuais()
            df_custo_completo, relatorio = montar_custo_completo(planilhas, alteracoes_manuais)

            df_emissao = None
            if self.incluir_emissao.get():
                df_emissao = emissao.montar_tabela_emissao(
                    planilhas["FATURAMENTO"], df_custo_completo, planilhas["BAIXA ESPECIAL"],
                )

            with pd.ExcelWriter(caminho_saida, engine="openpyxl") as writer:
                # As fórmulas FILTER/SUMPRODUCT da AUDITORIA e do dashboard são
                # escritas sem valor em cache (openpyxl não calcula fórmulas) -
                # isso força o Excel a recalcular tudo ao abrir o arquivo, em
                # vez de mostrar células em branco até um F9 manual.
                writer.book.calculation.fullCalcOnLoad = True

                if self.incluir_bases.get():
                    for nome_tabela in TABELAS:
                        planilhas[nome_tabela].to_excel(writer, sheet_name=nome_tabela, index=False)

                categorias = banco_dados.listar_categorias()
                subcategorias = banco_dados.listar_subcategorias()

                df_custo_completo.to_excel(writer, sheet_name="CUSTO COMPLETO", index=False)
                formatar_planilha_custo_completo(writer.sheets["CUSTO COMPLETO"], len(df_custo_completo))
                aplicar_validacoes_categoria(
                    writer.sheets["CUSTO COMPLETO"], len(df_custo_completo), categorias, subcategorias,
                )
                criar_tabela_custo_completo(writer.sheets["CUSTO COMPLETO"], len(df_custo_completo))

                if df_emissao is not None:
                    df_emissao.to_excel(writer, sheet_name="EMISSÃO", index=False)
                    emissao.formatar_planilha_emissao(writer.sheets["EMISSÃO"], len(df_emissao))

                if self.incluir_dashboard.get():
                    dashboard.construir_dashboard(
                        writer, df_custo_completo, relatorio,
                        planilhas["BAIXA ESPECIAL"], data_inicio, data_fim,
                        df_emissao=df_emissao, cte=planilhas["CTE"],
                        subcategorias=subcategorias,
                    )

        except Exception as erro:
            messagebox.showerror("Erro ao gerar planilha", str(erro))
            return

        self._registrar_log(
            f"=== Planilha gerada: {caminho_saida} ===\n"
            f"Período: {data_inicio.strftime('%d/%m/%Y')} a {data_fim.strftime('%d/%m/%Y')}\n"
            f"Total de notas processadas: {relatorio['total_notas']}\n"
            f"Sem CT-e correspondente: {relatorio['sem_cte']}\n"
            f"Sem pedido de compra correspondente: {relatorio['sem_pedido_compra']}\n"
            f"Sem baixa especial correspondente: {relatorio['sem_baixa_especial']}\n"
        )
        if relatorio["total_notas"] == 0:
            messagebox.showwarning(
                "Nenhuma nota encontrada",
                "Não há notas no período escolhido. Confira as datas e se as bases já foram importadas.",
            )

        try:
            os.startfile(caminho_saida)
        except OSError as erro:
            messagebox.showwarning(
                "Não foi possível abrir a planilha",
                f"A planilha foi gerada em {caminho_saida}, mas não consegui abri-la automaticamente:\n{erro}",
            )


def main():
    raiz = tk.Tk()
    Aplicativo(raiz)
    raiz.mainloop()


if __name__ == "__main__":
    main()
