def Checando_pares(coluna_REDE, coluna_w3rp):
    global pares_encontrados

    sem_par_REDE = []
    sem_par_w3rp = []

    for i_REDE, valor_REDE in enumerate(coluna_REDE):
        encontrado = False
        start_index_w3rp = pares_encontrados[-1][1] + 1 if pares_encontrados else 0

        for i_w3rp in range(start_index_w3rp, len(coluna_w3rp)):
            valor_w3rp = coluna_w3rp[i_w3rp]

            # Comparação para números de ponto flutuante
            if abs(valor_REDE - valor_w3rp) < 1e-6:
                pares_encontrados.append((i_REDE, i_w3rp))
                encontrado = True
                break

        if not encontrado:
            sem_par_REDE.append((i_REDE, valor_REDE))
            sem_par_w3rp.append(None)

    print("Pares Encontrados:")
    for par in pares_encontrados:
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
    print(pares_encontrados)
    return pares_encontrados, sem_par_REDE, sem_par_w3rp