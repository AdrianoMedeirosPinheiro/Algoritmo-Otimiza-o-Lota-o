import pandas as pd
from itertools import combinations

def otimizar_alocacao_professores(arquivo):
    df = pd.read_excel(arquivo)
    alocacoes_otimizadas = {}
    professores = df.groupby('SERVIDOR')

    for nome, dados_professor in professores:
        escolas_professor = dados_professor['ESCOLA'].unique()
        combinacoes_escolas = combinations(escolas_professor, 2)

        for comb in combinacoes_escolas:
            professores_potenciais = df[(df['ESCOLA'].isin(comb)) &
                                        ((df['HABILITAÇÃO 1'] == dados_professor['HABILITAÇÃO 1'].iloc[0]) |
                                         (df['HABILITAÇÃO 2'] == dados_professor['HABILITAÇÃO 2'].iloc[0]))]

            for _, prof_potencial in professores_potenciais.iterrows():
                if prof_potencial['SERVIDOR'] != nome:
                    carga_horaria_atual_professor = dados_professor['CH_ESCOLA'].sum()
                    carga_horaria_prof_potencial = df[df['SERVIDOR'] == prof_potencial['SERVIDOR']]['CH_ESCOLA'].sum()

                    if carga_horaria_atual_professor == carga_horaria_prof_potencial or \
                       abs(carga_horaria_atual_professor - carga_horaria_prof_potencial) <= 2:
                        chave = tuple(sorted((nome, prof_potencial['SERVIDOR'])))
                        if chave not in alocacoes_otimizadas:
                            alocacoes_otimizadas[chave] = {
                                'Escola 1': comb[0],
                                'Escola 2': comb[1],
                                'CH Total': carga_horaria_atual_professor
                            }

    return alocacoes_otimizadas

def imprimir_resultados(otimizacoes, df):
    linhas = []

    for (prof1, prof2), dados in otimizacoes.items():
        escola_origem_prof1 = dados['Escola 1']
        escola_destino_prof1 = dados['Escola 2']
        escola_origem_prof2 = dados['Escola 2']
        escola_destino_prof2 = dados['Escola 1']

        # Atualizando as escolas no DataFrame original
        df.loc[df['SERVIDOR'] == prof1, 'ESCOLA'] = escola_destino_prof1
        df.loc[df['SERVIDOR'] == prof2, 'ESCOLA'] = escola_destino_prof2

        # Recalculando a carga horária
        carga_horaria_prof1 = df[df['SERVIDOR'] == prof1]['CH_ESCOLA'].sum()
        carga_horaria_prof2 = df[df['SERVIDOR'] == prof2]['CH_ESCOLA'].sum()

        linha = {
            'Professor': prof1,
            'Escola Origem': escola_origem_prof1,
            'Escola Destino': escola_destino_prof1,
            'CH_Escola': carga_horaria_prof1
        }
        linhas.append(linha)

        linha = {
            'Professor': prof2,
            'Escola Origem': escola_origem_prof2,
            'Escola Destino': escola_destino_prof2,
            'CH_Escola': carga_horaria_prof2
        }
        linhas.append(linha)

    resultados_df = pd.DataFrame(linhas)
    print(f"Total de Professores que conseguiram ficar em uma escola: {resultados_df['Professor'].nunique()}")
    print(resultados_df)
    resultados_df.to_excel('df_otimizado.xlsx', index=False)
    df.to_excel('df_final.xlsx', index=False)

# Exemplo de uso:
df = pd.read_excel('Base_Adriano.xlsx')
otimizacoes = otimizar_alocacao_professores('Base_Adriano.xlsx')
imprimir_resultados(otimizacoes, df)
