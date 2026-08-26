import os
import tkinter as tk
from datetime import date
from tkinter import filedialog, messagebox, scrolledtext, ttk
from tkcalendar import DateEntry
from config.config import TABELAS
import models.sql as banco_dados
import models.dashboard as M_final


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("RECEBIMENTOS E DESPESAS")

        self.incluir_bases = tk.BooleanVar(value=True)
        self.caminho_saida = tk.StringVar()
        self.labels_status = {}

        self.montar_status()
        self.montar_importacao()
        self.selecao_periodo()
        self.secao_saida()
        self._montar_secao_gerar()

    # =================================================
    #           STATUS DATA DE IMPORTAÇÃO
    # =================================================
    def montar_status(self):
        moldura = ttk.Frame(self)
        moldura.pack(fill="x", padx=10, pady=(10, 0))
        self.label_data_atualizada = ttk.Label(
            moldura, font=("TkDefaultFont", 10, "bold")
        )
        self.label_data_atualizada.pack(anchor="w")
        self._atualizar_label_data_atualizada()

    def _atualizar_label_data_atualizada(self):
        data_maxima = banco_dados.data_maxima_atualizada()
        if data_maxima is None:
            texto = "Dados atualizados até: (nenhum dado importado ainda)"
        else:
            texto = f"Dados atualizados até: {data_maxima.strftime('%d/%m/%Y')}"
        self.label_data_atualizada.config(text=texto)

    # =================================================
    #           SEÇÃO DE IMPORTAÇÃO
    # =================================================
    def montar_importacao(self):
        moldura = ttk.LabelFrame(
            self, text="Importar bases (.xlsx) para o banco de dados"
        )
        moldura.pack(fill="x", padx=10, pady=10)

        for linha, nome_tabela in enumerate(TABELAS):
            ttk.Label(moldura, text=nome_tabela, width=18).grid(
                row=linha, column=0, padx=5, pady=3, sticky="w"
            )

            botao = ttk.Button(
                moldura,
                text="selecionar arquivo...",
                command=lambda n=nome_tabela: self._selecionar_e_importar(n),
            )
            botao.grid(row=linha, column=1, padx=5, pady=3)

            label_status = ttk.Label(
                moldura, text="Nenhum arquivo importado ainda", foreground="gray"
            )
            label_status.grid(row=linha, column=2, padx=5, pady=3, sticky="w")
            self.labels_status[nome_tabela] = label_status

    def _selecionar_e_importar(self, arquivo):
        caminhos = filedialog.askopenfilenames(
            title=f"Selecionar arquivo(s) .xlsx para {arquivo}",
            filetypes=[("Planilhas Excel", "*.xlsx")],
        )
        if not caminhos:
            return

        total_arquivo = novas = duplicadas = arquivos_ok = 0
        erros = []
        for caminho in caminhos:
            try:
                abas_disponiveis = banco_dados.listar_abas_compativeis(arquivo, caminho)
                if not abas_disponiveis:
                    raise ValueError(
                        f"Nenhuma aba de '{caminho}' bate com o formato esperado de '{arquivo}'."
                    )
                if len(abas_disponiveis) > 1:
                    abas_escolhidas = self._escolher_abas(caminho, abas_disponiveis)
                    if not abas_escolhidas:
                        continue  # cancelado ou nenhuma aba marcada
                else:
                    abas_escolhidas = abas_disponiveis
                resultado = banco_dados.importar_arquivo(arquivo, caminho, abas=abas_escolhidas)
            except Exception as erro:
                erros.append(f"{caminho}:\n{erro}")
                continue
            arquivos_ok += 1
            total_arquivo += resultado["total_arquivo"]
            novas += resultado["novas"]
            duplicadas += resultado["duplicadas"]

        if erros:
            messagebox.showerror(
                "erro ao importar",
                f"Falha ao importar {len(erros)} de {len(caminhos)} arquivo(s) para '{arquivo}':\n\n"
                + "\n\n".join(erros),
            )

        if arquivos_ok:
            self.labels_status[arquivo].config(
                text=(
                    f"{arquivos_ok} arquivo(s) importado(s) — {total_arquivo} linhas — "
                    f"{novas} novas, {duplicadas} duplicadas ignoradas"
                ),
                foreground="black",
            )
            self._atualizar_label_data_atualizada()

    def _escolher_abas(self, caminho, abas_disponiveis):
        """Mostra um diálogo com uma checkbox por aba compatível e devolve as marcadas (ou [] se cancelado)."""
        janela = tk.Toplevel(self)
        janela.title(f"Selecionar abas - {os.path.basename(caminho)}")
        janela.transient(self)
        janela.grab_set()
        janela.resizable(False, False)

        ttk.Label(janela, text="Escolha quais abas importar deste arquivo:").pack(
            anchor="w", padx=10, pady=(10, 5)
        )

        variaveis = {}
        for aba in abas_disponiveis:
            var = tk.BooleanVar(value=True)
            ttk.Checkbutton(janela, text=aba, variable=var).pack(anchor="w", padx=20)
            variaveis[aba] = var

        resultado = []

        def confirmar():
            resultado.extend(aba for aba, var in variaveis.items() if var.get())
            janela.destroy()

        def cancelar():
            janela.destroy()

        moldura_botoes = ttk.Frame(janela)
        moldura_botoes.pack(fill="x", padx=10, pady=10)
        ttk.Button(moldura_botoes, text="Cancelar", command=cancelar).pack(side="right", padx=(5, 0))
        ttk.Button(moldura_botoes, text="Importar", command=confirmar).pack(side="right")

        janela.protocol("WM_DELETE_WINDOW", cancelar)
        janela.wait_window()
        return resultado

    # =================================================
    #           SELEÇÃO DE PERIODO
    # =================================================
    def selecao_periodo(self):
        moldura = ttk.LabelFrame(
            self, text="Período (define o que entra na planilha de saída)"
        )
        moldura.pack(fill="x", padx=10, pady=5)

        hoje = date.today()
        inicio_mes = hoje.replace(day=1)

        ttk.Label(moldura, text="Data início:").grid(
            row=0, column=0, padx=5, pady=5, sticky="w"
        )
        self.data_inicio_widget = DateEntry(
            moldura,
            date_pattern="dd/mm/yyyy",
            locale="pt_BR",
            year=inicio_mes.year,
            month=inicio_mes.month,
            day=inicio_mes.day,
        )
        self.data_inicio_widget.grid(row=0, column=1, padx=5, pady=5)

        ttk.Label(moldura, text="Data fim:").grid(
            row=0, column=2, padx=5, pady=5, sticky="w"
        )
        self.data_fim_widget = DateEntry(
            moldura, date_pattern="dd/mm/yyyy", locale="pt_BR"
        )
        self.data_fim_widget.grid(row=0, column=3, padx=5, pady=5)

    # =================================================
    #           SEÇÃO DE SAÍDA
    # =================================================
    def secao_saida(self):
        moldura = ttk.LabelFrame(self, text="saída")
        moldura.pack(fill="x", padx=10, pady=5)

        ttk.Label(moldura, text="Salvar planilha em:").grid(
            row=0, column=0, padx=5, pady=5, sticky="w"
        )
        ttk.Entry(moldura, textvariable=self.caminho_saida, width=50).grid(
            row=0, column=1, padx=5, pady=5
        )
        ttk.Button(
            moldura, text="Salvar como...", command=self._escolher_caminho_saida
        ).grid(row=0, column=2, padx=5, pady=5)

        ttk.Checkbutton(
            moldura,
            text="Incluir abas dos bancos de dados",
            variable=self.incluir_bases,
        ).grid(row=1, column=0, columnspan=2, padx=5, pady=3, sticky="w")

    def _escolher_caminho_saida(self):
            caminho = filedialog.asksaveasfilename(
                title="Salvar planilha de saída",
                defaultextension=".xlsx",
                filetypes=[("Planilha Excel", "*.xlsx")],
            )
            if caminho:
                self.caminho_saida.set(caminho)

    # =================================================
    #           GERAR PLANILHA
    # =================================================
    def _montar_secao_gerar(self):
        ttk.Button(self, text="Gerar planilha", command=self._gerar_planilha).pack(pady=5)

        moldura = ttk.LabelFrame(self, text="Log")
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
            M_final.gerar_dashboard(caminho_saida, data_inicio, data_fim, self.incluir_bases.get())
        except Exception as erro:
            messagebox.showerror("Erro ao gerar planilha", str(erro))
            return

        self._registrar_log(
            f"=== Planilha gerada: {caminho_saida} ===\n"
            f"Período: {data_inicio.strftime('%d/%m/%Y')} a {data_fim.strftime('%d/%m/%Y')}\n"
        )

        try:
            os.startfile(caminho_saida)
        except OSError as erro:
            messagebox.showwarning(
                "Não foi possível abrir a planilha",
                f"A planilha foi gerada em {caminho_saida}, mas não consegui abri-la automaticamente:\n{erro}",
            )


if __name__ == "__main__":
    app = App()
    app.mainloop()
