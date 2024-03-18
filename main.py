import pandas

from pathlib import Path

import tkinter
from tkinter import filedialog, INSERT, DISABLED, WORD
from tkinter.messagebox import showinfo

import spacy

words_file_path = ""
base_file_path = ""


def browse_words_file():
    global words_file_path
    words_file_path = filedialog.askopenfilename(initialdir="/",
                                                 title="Selecione o Arquivo de Palavras:",
                                                 filetypes=(("Excel Files",
                                                             "*.xlsx*"),
                                                            ("all files",
                                                             "*.*")))

    words_file_label['text'] = "Arquivo de palavras: " + words_file_path
    print(words_file_path)


def browse_base_file():
    global base_file_path
    base_file_path = filedialog.askopenfilename(initialdir="/",
                                                title="Selecione o Arquivo de Base:",
                                                filetypes=(("Excel Files",
                                                            "*.xlsx*"),
                                                           ("all files",
                                                            "*.*")))

    base_file_label['text'] = "Arquivo base: " + base_file_path
    print(base_file_path)


def get_lemas():
    load_model = spacy.load('pt_core_news_sm', disable=['parser'])
    lemma_tags = {"NN", "NNP", "NNS", "NNPS"}
    lemma = ""

    # for palavara in palavras_interesse:
    #     My_text = palavara
    #     doc = load_model(My_text)
    #
    #     for token in doc:
    #         lemma = token.text
    #         if token.tag_ in lemma_tags:
    #             lemma = token.lemma_
    #         result = lemma
    #         print(result)


def treate_files():
    file_serach_words = words_file_path

    xlsx_search_words = pandas.read_excel(file_serach_words)

    xlsx_search_words['Palavras'] = xlsx_search_words['Palavras'].str.lower()

    file = base_file_path
    # print(configuration["bases"][file]["caminho"])
    # print(configuration["bases"][file]["aba"])

    if len(skip_rows_entry.get()) == 0:
        skip_rows = None
    else:
        skip_rows = skip_rows_entry.get()

    # leitura do arquivo xlsx pulando as linhas informadas
    xlsx_file = pandas.read_excel(file, sheet_name_entry.get(), skiprows=skip_rows)

    # carrega só as colunas que são texto
    text_columns = xlsx_file.select_dtypes(include=object).columns

    desired_columns = text_columns.tolist()
    # adiciona a coluna chave
    # desired_columns.append("COD. DA AÇÃO")

    # filtra só as colunas de texto
    search_base = xlsx_file[[col for col in xlsx_file.columns if col in desired_columns]]
    # print(search_base)
    search_base.loc[:, 'tem_palavra_interesse'] = False
    search_base.loc[:, 'coluna_palavra_interesse'] = ""

    # para cada linha da base
    for index, row in search_base.iterrows():
        # print("estou na linha: ", index)
        # para cada coluna da busca, verifica se tem a palavra de interesse
        for column_name in text_columns:
            value_find = str(search_base.at[index, column_name]).lower()
            for word in xlsx_search_words['Palavras']:
                if word in value_find:
                    search_base.at[index, 'tem_palavra_interesse'] = True
                    key_value = column_name + ": " + word
                    search_base.at[index, 'coluna_palavra_interesse'] = (
                            search_base.loc[index, 'coluna_palavra_interesse'] + key_value) \
                        if search_base.loc[index, 'coluna_palavra_interesse'] == "" \
                        else search_base.loc[index, 'coluna_palavra_interesse'] + ", " + key_value

    result = search_base[search_base['tem_palavra_interesse'] == True]
    result_path = str(Path(file).resolve().parent)
    result.to_csv(result_path + '\\\\resultado_' + Path(file).stem + '.csv', sep=';', encoding='utf-8-sig', index=False)

    showinfo(
        message='Finalizado! Resultado está no caminho: ' + result_path + '\\resultado_' + Path(file).stem + '.csv')


if __name__ == '__main__':
    filename = None

    window = tkinter.Tk()
    window.title("Ctrl+F")
    window.iconbitmap('icone.ico')
    window.config(padx=10, pady=10)

    informative_label = tkinter.Label(text="Informe os campos abaixo:", font='Calibri 14 bold')
    informative_label.grid(row=0)

    words_file_label = tkinter.Label(text="Arquivo de palavras selecionado: ", font='Calibri 14')

    words_file_browse_button = tkinter.Button(window,
                                              text="Selecionar arquivo de palavras",
                                              font='Calibri 12',
                                              command=browse_words_file)

    words_file_label.grid(row=1, column=0)
    words_file_browse_button.grid(row=1, column=1)

    base_file_label = tkinter.Label(text="Arquivo base selecionado: ", font='Calibri 14')

    file_browse_button = tkinter.Button(window,
                                        text="Selecionar arquivo base",
                                        font='Calibri 12',
                                        command=browse_base_file)

    base_file_label.grid(row=2, column=0)
    file_browse_button.grid(row=2, column=1)

    skip_rows_label = tkinter.Label(
        text="Linhas para pular:",
        font='Calibri 14')
    skip_rows_label.grid(row=5, column=0)

    skip_rows_entry = tkinter.Entry(window)
    skip_rows_entry.grid(row=5, column=1)

    sheet_name_label = tkinter.Label(text="Nome da aba a ser lida:", font='Calibri 14')
    sheet_name_label.grid(row=6, column=0)

    sheet_name_entry = tkinter.Entry(window)
    sheet_name_entry.grid(row=6, column=1)

    process_file_button = tkinter.Button(text="Processar Arquivo", font='Calibri 12', fg='white', bg='green',
                                         command=treate_files)
    process_file_button.grid(row=8, column=3)

    text_instructions = tkinter.Text(window)
    text_instructions.insert(INSERT, "Instruções:\n"
                                     "1 - Selecione o arquivo com as palavras a serem buscadas, no botão 'Selecionar "
                                     "arquivo de palavras'\n"
                                     "O arquivo deve ter uma unica coluna chamada 'Palavras' e de estar no formato "
                                     "XLSX.\n"
                                     "2 - Selecione o arquivo a base a ser consultada, no botão 'Selecionar arquivo "
                                     "base'\n"
                                     "  O arquivo deve estar no formato XLSX.\n"
                                     "3 - Indique os númveros das linhas a serem puladas na leitura da base, "
                                     "separados por virgula\n"
                                     "Esse campo será utilizadao para arquivos que tem linhas antes da linha de "
                                     "cabeçalho.\n"
                                     "Para remove-las, basta indicar os números da linhas, iniciando a contagem em 0.\n"
                                     "4 - Indique o nome da aba do arquivo base que deve ser lida.\n"
                                     "5 - Clique no botão processar e aguarde a informação de finalização\n"
                                     "O arquivo será gerado na mesma localização do arquivo base, com o mesmo nome "
                                     "adicionado o sufixo '_resultado' ao final"
                             )
    text_instructions.config(state=DISABLED, wrap=WORD)
    text_instructions.grid(row=9, column=0, padx=30, pady=100)

    window.mainloop()
