def difference_check(coluna_repetidos, coluna_checagem):
    somas_repetidas = {}

    # Iterar sobre os elementos da coluna
    for elemento in coluna_repetidos:
        elemento_str = str(elemento)

        # Adicionar o valor ao dicionário ou somar se já existir
        if elemento_str in somas_repetidas:
            somas_repetidas[elemento_str] += elemento
        else:
            somas_repetidas[elemento_str] = elemento

    # Iterar sobre o dicionário e imprimir as repetições e somas
    for valor, soma in somas_repetidas.items():
        repeticoes = coluna_repetidos[coluna_repetidos.astype(str) == valor].count()

        if repeticoes >= 2:
            print(f"Repetição do {valor}: {repeticoes} vezes e gera a soma: {soma}")

            indices_checagem = coluna_checagem[coluna_checagem == soma].index + 2 #TODO CORRIGIR num_linhas do skiprows
            
            # Imprimir os índices correspondentes
            if not indices_checagem.empty:
                print(f"A soma {soma} está nas linhas {indices_checagem.to_list()} da REDE")
            else:
                print(f"A soma {soma} não foi encontrada na coluna_checagem")