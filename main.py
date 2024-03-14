import config
import pandas

import spacy

configuration = config.get_config()
palavras_interesse = [x.lower() for x in configuration["palavras_interesse"]]


# palavras_interesse_lematizadas =

def get_lemas():
    load_model = spacy.load('pt_core_news_sm', disable=['parser'])
    lemma_tags = {"NNS", "NNPS"}
    lemma = ""

    for palavara in palavras_interesse:
        My_text = palavara
        doc = load_model(My_text)

        for token in doc:
            lemma = token.text
            if token.tag_ in lemma_tags:
                lemma = token.lemma_
            result = lemma
            print(result)


def treate_files(file):
    print(configuration["bases"][file]["caminho"])
    print(configuration["bases"][file]["aba"])

    if configuration["bases"][file]["quantidade_linhas_pular"] == -1:
        skip_rows = None
    else:
        skip_rows = configuration["bases"][file]["quantidade_linhas_pular"]

    # leitura do arquivo xlsx pulando a primeira linha
    xlsx_file = pandas.read_excel(configuration["bases"][file]["caminho"],
                                  configuration["bases"][file]["aba"]
                                  # , skiprows=[configuration["bases"][file]["quantidade_linhas_pular"]]
                                  , skiprows=skip_rows)

    # carrega só as colunas que são texto
    text_columns = xlsx_file.select_dtypes(include=object).columns

    desired_columns = text_columns.tolist()
    # adiciona a coluna chave
    desired_columns.append(configuration["bases"][file]["coluna_chave"])

    # filtra só as colunas de texto e a chave
    search_base = xlsx_file[[col for col in xlsx_file.columns if col in desired_columns]]
    # print(search_base)
    search_base.loc[:, 'tem_palavra_interesse'] = False

    # para cada linha da base
    for index, row in search_base.iterrows():
        # print("estou na linha: ", index)

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

    bases = configuration["bases"]
    print(bases)

    for key in bases:
        print(key)
        treate_files(key)

