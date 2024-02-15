import tkinter as tk
from tkinter import filedialog
from tkinter import ttk
import pandas as pd

class ExcelViewerApp(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("Excel Viewer")

        self.tree = ttk.Treeview(self)
        self.tree.pack(fill="both", expand=True)

        self.menu_bar = tk.Menu(self)
        self.config(menu=self.menu_bar)

        file_menu = tk.Menu(self.menu_bar, tearoff=False)
        file_menu.add_command(label="Open Excel", command=self.open_excel)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.quit)
        self.menu_bar.add_cascade(label="File", menu=file_menu)

    def open_excel(self):
        file_path = filedialog.askopenfilename(title="Select an Excel file", filetypes=[("Excel files", "*.xlsx *.xls")])
        if file_path:
            excel_data = pd.read_excel(file_path)

            # Clear previous content
            self.tree.delete(*self.tree.get_children())

            # Insert Excel data into the Treeview
            headers = excel_data.columns.tolist()
            self.tree["columns"] = headers
            self.tree.heading("#0", text="Index")
            for i, header in enumerate(headers):
                self.tree.heading(header, text=header)
                self.tree.column(header, anchor="center")
            for index, row in excel_data.iterrows():
                self.tree.insert("", tk.END, text=str(index), values=row.tolist())

if __name__ == "__main__":
    app = ExcelViewerApp()
    app.mainloop()


