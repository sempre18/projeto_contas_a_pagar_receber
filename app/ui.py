import os
import tkinter as tk
from datetime import date
from tkinter import filedialog, messagebox, scrolledtext, ttk
import pandas as pd
from tkcalendar import DateEntry
from config.config import TABELAS


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("RECEBIMENTOS E DESPESAS")

        self.montar_status()
        self.montar_importacao()

    def montar_status(self):
        pass

    def montar_importacao(self):
        pass


if __name__ == "__main__":
    app = App()
    app.mainloop()
