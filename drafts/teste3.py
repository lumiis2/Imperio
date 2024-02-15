import tkinter as tk
from tkinter import messagebox

def divide_numbers():
    try:
        result = 10 / 0  # Exemplo de uma divisão por zero que vai lançar uma exceção
        print(result)
    except ZeroDivisionError:
        messagebox.showerror("Erro", "Não é possível dividir por zero!")

class MyApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Minha Aplicação")

        self.button = tk.Button(self, text="Dividir", command=divide_numbers)
        self.button.pack()

if __name__ == "__main__":
    app = MyApp()
    app.mainloop()
