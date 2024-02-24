
import pandas as pd
import subprocess   

def set_var(self):
        self.root = None
        self.tolerance = 0.5
        self.var_name = "CASTELO"
        self.csv_data = pd.DataFrame()
        self.excel_data = pd.DataFrame()
        self.general_data = pd.DataFrame()

        ### QUERO DIMINUIR ESSAS VARS GLOBAIS
        reset_var(self)
        self.csv_path = " "
        self.excel_path = " "
        self.temp_excel_path = " "
        self.temp_csv_path = " "
        self.temp_var_name = " "

def reset_var(self):
    self.rede_storage = []
    self.w3_storage = []
    self.w3_storage_s = []
    self.rede_storage_s = []
    self.pares_encontrados = []
    self.order = 0
    
        

def choose_store(self, store):
        store_mappings = {
            "Loja Castelo": "CASTELO",
            "Loja Cidade Nova": "CID. NOVA",
            "Loja Planalto": "PLANALTO",
            "Loja Contagem": "CONTAGEM",
            "Loja Nova Lima": "NOVA LIMA",
            "Loja E-commerce": "E-COMM"
        }
        self.var_name = store_mappings.get(store, "Loja não encontrada")
        self.order = 1

def adjust_size(self):
    num_rows = max(self.excel_data.shape[0], self.csv_data.shape[0])

    missing_rows = num_rows - self.excel_data.shape[0]
    if missing_rows > 0:
        null_rows = pd.DataFrame(index=range(self.excel_data.shape[0], num_rows))
        self.excel_data = pd.concat([self.excel_data, null_rows])
    return self.excel_data



def check_diff(self, coluna_repetidos, coluna_checagem, storage, verbose=False):
    valores = []
    somas_repetidas = {}

    tolerancia = 0.03

    for elemento in coluna_repetidos:
        elemento_arredondado = round(elemento, 2)
        elemento_str = str(elemento_arredondado)

        encontrado = False
        for chave, valor in somas_repetidas.items():
            chave_arredondada = round(float(chave), 2)
            if abs(elemento_arredondado - chave_arredondada) <= tolerancia:
                somas_repetidas[chave] += elemento
                valores.append(elemento)
                encontrado = True
                break

        if not encontrado:
            somas_repetidas[elemento_str] = elemento

    for valor, soma in somas_repetidas.items():
        repeticoes = coluna_repetidos[abs(coluna_repetidos.astype(float) - float(valor)) <=0.03].count()

        if repeticoes >= 2:
            print(f"Repetição do {valor}: {repeticoes} vezes e gera a soma: {soma}")

            indices_checagem = coluna_checagem[abs(coluna_checagem - soma) <= 0.3].index #+ var_skip
            
            if not indices_checagem.empty:
                print(f"A soma {soma} está nas linhas {indices_checagem.to_list()} da REDE")
                if storage == "w3erp":
                    self.w3_storage.append(float(valor))
                    self.rede_storage_s.append(indices_checagem.to_list()[0])
                elif storage == "REDE":
                    self.rede_storage.append(float(valor))
                    self.w3_storage_s.append(indices_checagem.to_list()[0])
            else:
                print(f"A soma {soma} não foi encontrada na coluna_checagem")

    if verbose:
        print("Rede storage: ", self.rede_storage)
        print("W3 storage: ", self.w3_storage)
        print("Rede soma: ", self.rede_storage_s)
        print("W3 soma:", self.w3_storage_s)
        print("Soma repetidos", self.somas_repetidas)

def Checando_pares(self, coluna_REDE, coluna_w3rp, verbose=False):
    sem_par_REDE = []
    sem_par_w3rp = []

    for i_REDE, valor_REDE in enumerate(coluna_REDE):
        encontrado = False
        start_index_w3rp = self.pares_encontrados[-1][1] + 1 if self.pares_encontrados else 0

        for i_w3rp in range(start_index_w3rp, len(coluna_w3rp)):
            valor_w3rp = coluna_w3rp[i_w3rp]

            if abs(valor_REDE - valor_w3rp) < 0.05:
                self.pares_encontrados.append((i_REDE, i_w3rp))
                encontrado = True
                break

        if not encontrado:
            sem_par_REDE.append((i_REDE, valor_REDE))
            sem_par_w3rp.append(None)
    if verbose:
        print("Pares Encontrados:")
        for par in self.pares_encontrados:
            i_REDE, i_w3rp = par
            valor_REDE = coluna_REDE[i_REDE]
            valor_w3rp = coluna_w3rp[i_w3rp]
            print(f"({i_REDE}: {valor_REDE}, {i_w3rp}: {valor_w3rp})")

        print("\nSem Par na coluna_REDE:")
        for sem_par in sem_par_REDE:
            i_REDE, valor_REDE = sem_par
            print(f"({i_REDE}: {valor_REDE}, None)")

        print("\nSem Par na coluna_w3rp:")
        for sem_par in sem_par_w3rp:
            print(f"None")

        print("####################")
        print(self.pares_encontrados)
        return self.pares_encontrados, sem_par_REDE, sem_par_w3rp

def open_sheets(self):
        arquivo_diferencas = 'excel_diferencas_form.xlsx'
        try:
            subprocess.Popen(['start', 'excel', arquivo_diferencas], shell=True)
        except Exception as e:
            print(f"Erro ao abrir o arquivo: {e}")