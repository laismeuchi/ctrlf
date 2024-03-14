import config
import pandas

configuration = config.get_config()
palavras_interesse = [x.lower() for x in configuration["palavras_interesse"]]


def treate_files(file):
    # rag_orcamento
    # rag_acao
    # rag_programa

    print(configuration["bases"][file]["caminho"])
    print(configuration["bases"][file]["aba"])

    # leitura do arquivo xlsx pulando a primeira linha
    xlsx_file = pandas.read_excel(configuration["bases"][file]["caminho"],
                                  configuration["bases"][file]["aba"]
                                  # , skiprows=[configuration["bases"][file]["quantidade_linhas_pular"]]
                                  , skiprows=configuration["bases"][file]["quantidade_linhas_pular"])

    # carrega só as colunas que são texto
    text_columns = xlsx_file.select_dtypes(include=object).columns

    desired_columns = text_columns.tolist()
    # adiciona a coluna chave
    desired_columns.append(configuration["bases"][file]["coluna_chave"])

    # filtra só as colunas de texto e a chave
    search_base = xlsx_file[[col for col in xlsx_file.columns if col in desired_columns]]
    # print(search_base)
    search_base['tem_palavra_interesse'] = False

    # para cada linha da base
    for index, row in search_base.iterrows():
        print("estou na linha: ", index)

        # para cada coluna da busca, verifica se tem a palavra de interesse
        for column_name in text_columns:
            value_find = str(search_base.at[index, column_name]).lower()
            if any(x in value_find for x in palavras_interesse):
                row['tem_palavra_interesse'] = True
                # print("tem!")
                search_base.at[index, 'tem_palavra_interesse'] = row['tem_palavra_interesse']
                break

    result = search_base[search_base['tem_palavra_interesse'] == True]
    result.to_csv('bases/resultado_' + file + '.csv', sep=';', encoding='utf-8-sig', index=False)


if __name__ == '__main__':
    treate_files("rag_orcamento")
    treate_files("rag_acao")
    treate_files("rag_programa")
