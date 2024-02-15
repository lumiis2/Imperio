import tkinter as tk
from tkinter import ttk
import csv

class TabelaCSV(tk.Tk):
    def __init__(self, csv_path):
        super().__init__()

        self.title("Tabela CSV")
        self.geometry("600x400")

        # Criar Treeview
        self.tree = ttk.Treeview(self, columns=("Coluna1", "Coluna2", "Coluna3"), show="headings", selectmode="browse")

        # Definir cabeçalhos
        self.tree.heading("Coluna1", text="Coluna1")
        self.tree.heading("Coluna2", text="Coluna2")
        self.tree.heading("Coluna3", text="Coluna3")

        # Adicionar estilos
        style = ttk.Style(self)
        style.configure("Treeview", rowheight=30, font=("Arial", 12))

        # Adicionar cores de fundo alternadas
        #style.map("Treeview", background=[("!evenrow", "#E1E1E1"), ("!oddrow", "#F0F0F0")])

        # Ler dados do CSV
        self.carregar_csv(csv_path)

        # Empacotar Treeview
        self.tree.pack(expand=True, fill="both")

    def carregar_csv(self, csv_path):
        with open(csv_path, newline="", encoding="utf-8") as file:
            reader = csv.DictReader(file)
            for row in reader:
                self.tree.insert("", "end", values=(row["Coluna1"], row["Coluna2"], row["Coluna3"]))

if __name__ == "__main__":
    # Substitua "seu_arquivo.csv" pelo caminho do seu arquivo CSV
    app = TabelaCSV("Planilhas/teste3.csv")
    app.mainloop()
