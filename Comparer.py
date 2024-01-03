import tkinter
import tkinter.messagebox
import customtkinter
import pandas as pd
import os
import subprocess   
from tkinter import Tk, Button, Label, Toplevel, Menu, filedialog
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

customtkinter.set_appearance_mode("Dark")  # Modes: "System" (standard), "Dark", "Light"
customtkinter.set_default_color_theme("dark-blue")  # Themes: "blue" (standard), "green", "dark-blue"


class App(customtkinter.CTk):
    def __init__(self):
        super().__init__()

        # configure window
        self.title("Comparador de Planilhas")
        self.geometry(f"{550}x{290}")

        self.grid_columnconfigure(0, weight=1)  
        self.grid_rowconfigure(0, weight=0)  
        self.grid_rowconfigure((1, 2, 3), weight=1)

        self.menubar_frame = customtkinter.CTkFrame(self, height=30, corner_radius=0)
        self.menubar_frame.grid(row=0, column=0, columnspan=3, sticky="nsew")
        self.menubar_frame.grid_rowconfigure(0, weight=1)

        self.chooser_label = customtkinter.CTkLabel(self.menubar_frame, text="Escolha uma loja:", anchor="w")
        self.chooser_label.grid(row=0, column=1, columnspan=2, padx=5, pady=(10, 10))
        self.chooser_optionemenu = customtkinter.CTkOptionMenu(self.menubar_frame, values=["Loja Castelo", "Loja Cidade Nova", "Loja Planalto", "Loja Contagem", "Loja Nova Lima", "Loja E-commerce"],
                                                                       command=lambda selected_store: self.change_chooser_event(selected_store))
        self.chooser_optionemenu.grid(row=0, column=3, columnspan=4, padx=5, pady=(10, 10))

        self.button_1 = customtkinter.CTkButton(self, text="Adicionar REDE", command=lambda: self.dowload('xlsx'))
        self.button_1.grid(row=1, column=0, padx=20, pady=10)
        self.button_2 = customtkinter.CTkButton(self, text="Adicionar w3erp", command=lambda: self.dowload('csv'))
        self.button_2.grid(row=2, column=0, padx=20, pady=10)
        





    def change_chooser_event(self, store):
        global var_name
        if store == "Loja Castelo":
            var_name = "CASTELO"
        if store == "Loja Cidade Nova":
            var_name = "CID. NOVA"
        if store == "Loja Planalto":
            var_name = "PLANALTO"
        if store == "Loja Contagem":
            var_name = "CONTAGEM"
        if store =="Loja Nova Lima":
            var_name = "NOVA LIMA"
        if store =="Loja E-commerce":
            var_name = "E-COMM"

    def dowload(self, tipo):
        global var_csv, var_xlsx

        if tipo == 'csv':
            var_csv = filedialog.askopenfilename(title=f"Selecione o arquivo {tipo.upper()}", filetypes=[(f"Arquivos {tipo.upper()}", f"*.{tipo}")])
        elif tipo == 'xlsx':
            var_xlsx = filedialog.askopenfilename(title=f"Selecione o arquivo {tipo.upper()}", filetypes=[(f"Arquivos {tipo.upper()}", f"*.{tipo}")])

        if var_csv and var_xlsx:
            self.button_3 = customtkinter.CTkButton(self, text="Comparar", command=lambda: self.comparer())
            self.button_3.grid(row=3, column=0, padx=20, pady=10)

    def comparer(self):
        self.process()
        self.open_sheets()
        pass

    def process(self):
        global var_csv, var_xlsx
        print(var_xlsx)
        print(var_csv)

        excel_data = self.ler_planilha_excel(var_xlsx)

        csv_data = self.ler_planilha_csv(var_csv)

        Valor_REDE = excel_data[2]
        Valor_w3rp = csv_data['Total']
        Parcelas = excel_data[11]
        

        print("W3 para REDE")
        #check_diff(Valor_w3rp, Valor_REDE)
        #print("REDE para W3")
        #check_diff(Valor_REDE, Valor_w3rp)
        #Checando_pares(Valor_REDE, Valor_w3rp)

        difference_sheet = pd.DataFrame(columns=["Data Recebimento", "Data Original", "Valor_REDE", "Valor_w3rp", "Metodo de Pagamento", "Parcelas", "Diferenca"])

        # Use min() para garantir que você está iterando sobre o menor número de linhas entre os dois DataFrames
        num_linhas = min(excel_data.shape[0], csv_data.shape[0])

        for index in range(num_linhas):
            Valor_REDE_atual = excel_data.iloc[index, 2]
            Valor_w3rp_atual = Valor_w3rp.iloc[index]

            diferenca = abs(Valor_REDE_atual - Valor_w3rp_atual)

            nova_linha = {
                "Data Recebimento": excel_data.iloc[index, 0],
                "Data Original": excel_data.iloc[index, 1],
                "Valor_REDE": Valor_REDE_atual,
                "Valor_w3rp": Valor_w3rp_atual,
                "Metodo de Pagamento": excel_data.iloc[index, 9],
                "Parcelas":  excel_data.iloc[index, 11],  # Certifique-se de entender como você deseja preencher esta coluna
                "Diferenca": diferenca
            }

            difference_sheet = difference_sheet._append(nova_linha, ignore_index=True)

        self.formatar_planilha_diferencas(difference_sheet, 'Planilhas/excel_diferencas_form.xlsx')

        # Imprimir a tabela de diferenças
        print("Tabela de Diferenças:")
        print(difference_sheet)

        # Salvar a tabela de diferenças como XLSX
        difference_sheet.to_excel('Planilhas/excel_diferencas.xlsx', index=False, float_format="%.2f")
    
    

    def formatar_planilha_diferencas(self, diferencas, caminho_arquivo):
        # Criar um novo arquivo Excel
        wb = Workbook()
        ws = wb.active

        # Adicionar os nomes das colunas como a primeira linha (cabeçalho)
        cabeçalhos = list(diferencas.columns)
        ws.append(cabeçalhos)

        # Adicionar os dados ao Excel
        for r_idx, row in enumerate(diferencas.values, start=2):  # Iniciar a partir da segunda linha (após o cabeçalho)
            for c_idx, value in enumerate(row, start=1):
                cell = ws.cell(row=r_idx, column=c_idx, value=value)

        # Ajustar a largura das colunas com base no tamanho do texto nos cabeçalhos
        for col_idx, col_name in enumerate(diferencas.columns, start=1):
            max_length = max(len(str(col_name)), max(len(str(cell.value)) for cell in ws[col_idx]))
            adjusted_width = (max_length + 2)
            ws.column_dimensions[get_column_letter(col_idx)].width = adjusted_width

        # Definir a coluna "Diferenca" em negrito
        for cell in ws['G']:
            cell.font = Font(bold=True)

        # Destacar valores de diferença acima de 60 centavos em vermelho
        for row in ws.iter_rows(min_row=2, max_col=7, max_row=ws.max_row):
            for cell in row:
                if cell.column == 7:  # Coluna "Diferenca"
                    if isinstance(cell.value, (int, float)) and cell.value > 0.6:  # 60 centavos
                        cell.fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")

        # Salvar o arquivo Excel
        wb.save(caminho_arquivo)

    def open_sheets(self):
        arquivo_diferencas = 'Planilhas/excel_diferencas_form.xlsx'

        try:
            # Abre o arquivo com o aplicativo padrão
            subprocess.Popen(['start', 'excel', arquivo_diferencas], shell=True)
        except Exception as e:
            print(f"Erro ao abrir o arquivo: {e}")

    def ler_planilha_excel(self, file_path):
        global var_name
        compare_col = 2
        num_lines = 0
        try:

            excel_data = pd.read_excel(file_path, header=None)

            for index, row in excel_data.iterrows():
                if row[0] == var_name:
                    num_lines = index + 1
                    break  

            print(num_lines)
            excel_data = pd.read_excel(file_path, header=None, skiprows= num_lines)
            
            print(f"Valor da coluna {compare_col} na planilha Excel:")
            print(excel_data.iloc[:, compare_col])

            return excel_data

        except Exception as e:
            print(f"Ocorreu um erro ao ler a planilha Excel: {e}")
            return None
        

    def ler_planilha_csv(self, file_path):
        compare_col = 'Total'
        try:
            csv_data = pd.read_csv(file_path, encoding='utf-8', delimiter=';', header=None, skip_blank_lines=False, names=['Total'])

            csv_data = csv_data.dropna(subset=['Total']).apply(lambda x: x.str.strip() if x.dtype == 'object' else x)

            csv_data.to_excel('Planilhas/w3_01-12.xlsx', index=False)

            csv_data = pd.read_excel('Planilhas/w3_01-12.xlsx', header=None, names=['Total'], skiprows=2)

            print(f"Valores da coluna {compare_col} na planilha convertida:")
            csv_data[compare_col] = pd.to_numeric(csv_data[compare_col].replace({r'\.': '', r',': '.'}, regex=True), errors='coerce').fillna(1)


            print(csv_data[compare_col])

            return csv_data
        
        except Exception as e:
            print(f"Ocorreu um erro ao ler a planilha CSV: {e}")
            return None

if __name__ == "__main__":

    script_directory = os.path.dirname(os.path.abspath(__file__))
    os.chdir(script_directory)

    var_csv = ''
    var_xlsx = ''
    root = None 
    var_name = "CASTELO"
    app = App()
    app.mainloop()