import config
import pandas

from pathlib import Path

import tkinter
from tkinter import filedialog, ttk
from tkinter.messagebox import showinfo

import spacy


# configuration = config.get_config()
# palavras_interesse = [x.lower() for x in configuration["palavras_interesse"]]


def browseFiles():
    filename = filedialog.askopenfilename(initialdir="/",
                                          title="Select a File",
                                          filetypes=(("Excel Files",
                                                      "*.xlsx*"),
                                                     ("all files",
                                                      "*.*")))

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
    file_serach_words = words_search_file_explorer_label.cget("text")

    xlsx_search_words = pandas.read_excel(file_serach_words)

    palavras_interesse = [x.lower() for x in xlsx_search_words]

    file = file_explorer_label.cget("text")
    # print(configuration["bases"][file]["caminho"])
    # print(configuration["bases"][file]["aba"])

    if len(skip_rows_entry.get()) == 0:
        skip_rows = None
    else:
        skip_rows = skip_rows_entry.get()

    # leitura do arquivo xlsx pulando a primeira linha
    xlsx_file = pandas.read_excel(file,
                                  sheet_name_entry.get()
                                  , skiprows=skip_rows)

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
            for word in palavras_interesse:
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
    window = tkinter.Tk()
    window.title("Ctrl+F")
    window.iconbitmap('icone.ico')
    window.config(padx=100, pady=100)

    informative_label = tkinter.Label(text="Informe os campos abaixo:", font='Calibri 14 bold')
    informative_label.grid(row=0)

    words_search_file_path_label = tkinter.Label(text="Selecione o Arquivo das palavras:", font='Calibri 14')
    words_search_file_path_label.grid(row=1, column=0)

    words_search_browse_button = tkinter.Button(window,
                                                text="Selecionar Arquivo",
                                                font='Calibri 12',
                                                command=browseFiles)
    words_search_browse_button.grid(row=1, column=1)

    words_search_file_explorer_label = tkinter.Label(text="Arquivo de palavras selecionado:", font='Calibri 14')
    words_search_file_explorer_label.grid(row=2, column=0)


    file_path_label = tkinter.Label(text="Selecione o Arquivo:", font='Calibri 14')
    file_path_label.grid(row=2, column=0, ipadx=0)

    file_browse_button = tkinter.Button(window,
                                        text="Selecionar Arquivo",
                                        font='Calibri 12',
                                        command=browseFiles)
    file_browse_button.grid(row=2, column=1)

    file_explorer_label = tkinter.Label(text="Arquivo selecionado:", font='Calibri 14')
    file_explorer_label.grid(row=3, column=0)



    skip_rows_label = tkinter.Label(
        text="Linhas para pular: (colocar separado por ',' e se não precisar pular deixar em branco)",
        font='Calibri 14')
    skip_rows_label.grid(row=4, column=0)

    skip_rows_entry = tkinter.Entry(window)
    skip_rows_entry.grid(row=4, column=1)

    sheet_name_label = tkinter.Label(text="Nome da aba a ser lida:", font='Calibri 14')
    sheet_name_label.grid(row=5, column=0)

    sheet_name_entry = tkinter.Entry(window)
    sheet_name_entry.grid(row=5, column=1)

    process_file_button = tkinter.Button(text="Processar Arquivo", font='Calibri 12', command=treate_files)
    process_file_button.grid(row=6, column=1)

    window.mainloop()
