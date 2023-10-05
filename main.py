import tkinter
from tkinter import filedialog, messagebox
import pandas as pd
import customtkinter
from tkcalendar import DateEntry
import requests
from datetime import datetime




requisicao = requests.get("http://economia.awesomeapi.com.br/json/all")
dicionari_moedas = requisicao.json()

lista_moedas = list(dicionari_moedas.keys())

def pegar_cotacao():
    try:
        moeda = combobox_selecionarmoeda.get()
        data_cotacao = calendario_moedas.get()
        data_atual = datetime.now()
        data_cotacao_date = datetime.strptime(data_cotacao, "%d/%m/%Y")

        if data_cotacao_date > data_atual:
            messagebox.showerror("Erro", "Você não pode pegar a cotação de um dia no futuro.")
            return

        ano = data_cotacao[-4:]
        mes = data_cotacao[3:5]
        dia = data_cotacao[:2]
        link = f"https://economia.awesomeapi.com.br/{moeda}-BRL/10?start_date={ano}{mes}{dia}&end_date={ano}{mes}{dia}"
        requisicao_moeda = requests.get(link)
        requisicao_moeda.raise_for_status()
        cotacao = requisicao_moeda.json()
        if cotacao and 'bid' in cotacao[0]:
            valor_moeda = cotacao[0]['bid']
            data_cotacao_formatada = f"{dia}/{mes}/{ano}"
            novo_texto = f"Moeda: {moeda}\n\n Data da Cotação: {data_cotacao_formatada}\n\n Valor: R${valor_moeda}"
            label_textocotacao.configure(text=novo_texto)
            print(valor_moeda)
        else:
            print("Não foi possível obter a cotação.")

    except requests.exceptions.RequestException as e:
        print(f"Erro na requisição HTTP: {e}")
    except Exception as e:
        print(f"Erro: {e}")

def selecionar_arquivo():
    global label_arquivoselecionado
    try:
        caminho_arquivo = filedialog.askopenfilename(title="Selecione o arquivo de moeda")
        var_caminho_arquivo.set(caminho_arquivo)
        if caminho_arquivo:
            if label_arquivoselecionado is None:
                label_arquivoselecionado = customtkinter.CTkLabel(Janela, text=f"Arquivo Selecionado: {caminho_arquivo}", anchor='e')
                label_arquivoselecionado.grid(row=6, column=0, columnspan=3, padx=10, pady=10, sticky='nsew')
            else:
                label_arquivoselecionado.configure(text=f"Arquivo Selecionado: {caminho_arquivo}")
    except Exception as e:
        print(f"Erro ao selecionar o arquivo: {e}")

def atualizar_cotacoes():
    try:
        arquivo_excel = var_caminho_arquivo.get()
        if not arquivo_excel:
            messagebox.showerror("Erro", "Selecione um arquivo Excel válido antes de atualizar cotações.")
            return

        df = pd.read_excel(arquivo_excel, index_col=0)  # Carrega o arquivo Excel com a primeira coluna como índice
        data_inicial = calendario_data_inicial.get()
        data_final = calendario_data_final.get()
        data_atual = datetime.now()
        data_inicial_date = datetime.strptime(data_inicial, "%d/%m/%Y")
        data_final_date = datetime.strptime(data_final, "%d/%m/%Y")

        if data_inicial_date > data_atual or data_final_date > data_atual:
            messagebox.showerror("Erro", "Você não pode atualizar cotações com datas no futuro.")
            return

        ano_inicial = data_inicial[-4:]
        mes_inicial = data_inicial[3:5]
        dia_inicial = data_inicial[:2]
        ano_final = data_final[-4:]
        mes_final = data_final[3:5]
        dia_final = data_final[:2]

        for moeda in lista_moedas:
            link = f"https://economia.awesomeapi.com.br/{moeda}-BRL/10?start_date={ano_inicial}{mes_inicial}{dia_inicial}" \
                   f"&end_date={ano_final}{mes_final}{dia_final}"
            requisicao_moeda = requests.get(link)
            requisicao_moeda.raise_for_status()
            cotacoes = requisicao_moeda.json()
            for cotacao in cotacoes:
                timestamp = int(cotacao['timestamp'])
                bid = float(cotacao['bid'])
                data = datetime.fromtimestamp(timestamp)
                data = datetime.strftime(data, '%d/%m/%y')
                df.at[moeda, data] = bid

        # Salvar o DataFrame atualizado em um arquivo Excel
        df.to_excel(arquivo_excel)
        label_atualizarcotacoes.configure(text="Cotações atualizadas e salvas com sucesso!")

    except requests.exceptions.RequestException as e:
        print(f"Erro na requisição HTTP: {e}")
    except Exception as e:
        print(f"Erro ao atualizar cotações: {e}")


Janela = customtkinter.CTk()
Janela._set_appearance_mode("system")
Janela.title("Ferramenta de cotação de moedas")
Janela.minsize(800, 600)

# Configuração de peso das linhas e colunas para tornar a janela responsiva
for i in range(10):
    Janela.grid_rowconfigure(i, weight=1)
    Janela.grid_columnconfigure(i, weight=1)




label_cotacao_moeda = customtkinter.CTkLabel(Janela, text=" cotação de uma moeda especifica")
label_cotacao_moeda.grid(row=0, column=0, padx=10, pady=10, sticky='nswe', columnspan=3)

label_selecionarmoeda = customtkinter.CTkLabel(Janela, text=" Selecionar moeda", anchor="e")
label_selecionarmoeda.grid(row=1, column=0, padx=10, pady=10, sticky='nswe', columnspan=2)

combobox_selecionarmoeda = customtkinter.CTkComboBox(Janela, values=lista_moedas)
combobox_selecionarmoeda.grid(row=1, column=2, padx=10, pady=10, sticky='nswe')

label_selecionar_dia = customtkinter.CTkLabel(Janela, text=" Selecione o dia que deseja pegar a cotação", anchor="e")
label_selecionar_dia.grid(row=2, column=0, padx=10, pady=10, sticky='nswe', columnspan=2)

calendario_moedas = DateEntry(Janela, year=2023, locale='pt_BR')
calendario_moedas.grid(row=2, column=2, padx=10, pady=10, sticky='nswe', columnspan=2)

label_textocotacao = customtkinter.CTkLabel(Janela, text="")
label_textocotacao.grid(row=3, column=0, columnspan=2, padx=10, pady=10, sticky='nsew')

botao_pegarcotacao = customtkinter.CTkButton(Janela, text="Pegar Cotação", command=pegar_cotacao)
botao_pegarcotacao.grid(row=3, column=2, padx=10, pady=10, sticky='nsew')

label_data_inicial = customtkinter.CTkLabel(Janela, text="Data Inicial", anchor='e')
label_data_inicial.grid(row=7, column=0, padx=10, pady=10, sticky="nswe")
calendario_data_inicial = DateEntry(Janela, year=2023, locale='pt_BR')
calendario_data_inicial.grid(row=7, column=1, padx=10, pady=10, sticky='nswe')

label_data_final = customtkinter.CTkLabel(Janela, text="Data Final", anchor='e')
label_data_final.grid(row=8, column=0, padx=10, pady=10, sticky='nswe')
calendario_data_final = DateEntry(Janela, year=2023, locale='pt_BR')
calendario_data_final.grid(row=8, column=1, padx=10, pady=10, sticky='nswe')

#cotação varias moedas

label_cotacavariasmoedas = customtkinter.CTkLabel(Janela, text="Cotação de multiplicação moedas")
label_cotacavariasmoedas.grid(row=4, column=0, padx=10, pady=10, sticky='nswe', columnspan=3)

label_selecionar_arquivo = customtkinter.CTkLabel(Janela, text="Selecionar um arquivo em Excel com as moedas na Coluna A")
label_selecionar_arquivo.grid(row=5, column=0, columnspan=2, padx=10, pady=10, sticky='nsew')

var_caminho_arquivo = tkinter.StringVar()

botao_selecionar_arquivo = customtkinter.CTkButton(Janela, text=" clique para selecionar arquivo", command=selecionar_arquivo)
botao_selecionar_arquivo.grid(row=5, column=2, padx=10, pady=10, sticky='nsew')

label_arquivoselecionado = customtkinter.CTkLabel(Janela, text="Nenhum arquivo selecionado", anchor='e')
label_arquivoselecionado.grid(row=6, column=0, columnspan=3, padx=10, pady=10, sticky='nsew')

botao_atualizarcotacoes = customtkinter.CTkButton(Janela, text="atualizar cotações", command=atualizar_cotacoes)
botao_atualizarcotacoes.grid(row=9, column=0, padx=10, pady=10, sticky='nsew')

label_atualizarcotacoes = customtkinter.CTkLabel(Janela, text="")
label_atualizarcotacoes.grid(row=9, column=1, columnspan=2, padx=10, pady=10, sticky='nsew')

botao_fechar = customtkinter.CTkButton(Janela, text="Fechar", command=Janela.quit)
botao_fechar.grid(row=10, column=2, padx=10, pady=10, sticky='nsew')

Janela.mainloop()
