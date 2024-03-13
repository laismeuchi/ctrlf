import numpy as np

import config
import pandas

configuration = config.get_config()


if __name__ == '__main__':

    palavras_interesse = [x.lower() for x in configuration["palavras_interesse"]]

    # leitura do arquivo xlsx pulando a primeira linha
    xlsx_file = pandas.read_excel(configuration["bases"]["rag_orcamento"]["caminho"],
                                  configuration["bases"]["rag_orcamento"]["aba"]
                                  , skiprows=[0])

    # carrega só as colunas que são texto
    text_columns = xlsx_file.select_dtypes(include=object).columns

    desired_columns = text_columns.tolist()
    # adiciona a coluna chave
    desired_columns.append('COD. DA AÇÃO')

    # filtra só as colunas de texto e a chave
    search_base = xlsx_file[[col for col in xlsx_file.columns if col in desired_columns]]
    # print(search_base)
    search_base['tem_palavra_interesse'] = False

    # para cada linha da base
    for index, row in search_base.iterrows():
        print("estou na linha: ", index)

        # para cada coluna da busca, verifica se tem a palavra de interesse
        for column_name in text_columns:
            if any(x in search_base.at[index, column_name].lower() for x in palavras_interesse):
                row['tem_palavra_interesse'] = True
                # print("tem!")
                search_base.at[index, 'tem_palavra_interesse'] = row['tem_palavra_interesse']
                break


    result = search_base[search_base['tem_palavra_interesse'] == True]
    result.to_csv('bases/resultado_rag_orcamento.csv', sep=';', encoding='utf-8', index=False)
