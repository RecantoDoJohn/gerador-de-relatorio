from tkinter import *
from tkinter import ttk
from os import startfile

import meses
import Metodos


def printar():
    global lista_metodos
    valor_lista = lista_metodos.get()

    nome_planilha = caixa.get() + '.xlsx'
    
    if nome_planilha == '.xlsx':
        return retorno.config(text="Insira o nome da planilha")
        
    try:
        indexi = meses.metodos.index(valor_lista)
        valor_planilha = str(Metodos.operacoes(nome_planilha).organizar(indexi))
        retorno.config(text=valor_planilha)
    except ValueError:
        retorno.config(text="Insira o nome da planilha e o metodo")

def abrir_planilha():
    try:
        nome_planilha = caixa.get() + '.xlsx'
        startfile(nome_planilha)
    except:
        retorno.config(text="Insira o nome da planilha")

tela = Tk()

tela.title("Trabalho de estatística")
tela.iconbitmap("iconjojo.ico")
# tela.iconphoto(True, icone)
tela.config(bg=('#78243f'))
tela.geometry("800x450")
tela.resizable(0, 0)


infos = Frame(
    tela,
    width=780,
    height=800,
    bg='#78243f',
)
infos.pack(pady=70)


titulo = Label(
    text="-=| ESTATISTICA:",
    font=("-weight bold", 25),
    bg='#78243f',
    fg='#24061d',
    # wraplength=250
)
titulo.pack()
titulo.place(x=10, y=5)

nsei = Label(
    infos,
    text="NOME DA PLANILHA:",
    font=("-weight bold", 20),
    bg='#78243f',
    fg='#24061d',
)
nsei.pack()

caixa = Entry(
    infos,
    font=("Arial", 17),
    width=15,
)
caixa.pack(pady=10)

listas_frames = Frame(
    infos,
    width=100,
    bg='#78243f',
)
listas_frames.pack(side=TOP)

lista1_frame = Frame(
    listas_frames,
    width=100,
    bg='#78243f',
) 
lista1_frame.pack(padx=8)

lista1_texto = Label(
    lista1_frame,
    text="INFORMAR MEDIDA:",
    font=("-weight bold", 20),
    bg='#78243f',
    fg='#24061d',
)
lista1_texto.pack()

lista_metodos = ttk.Combobox(
    lista1_frame,
    width=15,
    font=("Arial", 17),
    values=meses.metodos
)


lista_metodos.pack()


botoes = Frame(
    infos,
    width=100,
    bg='#78243f',
)
botoes.pack(side=BOTTOM)

botao = Button(
    botoes,
    text="Gerar",
    command=printar,
    font=("Arial", 15),
    bg='#d2632c',
    fg='#24061d',
    width=12,
    height=1
)
botao.pack(pady=20, padx=10, side=LEFT)

botao_plan = Button(
    botoes,
    text="Abrir Planilha",
    command=abrir_planilha,
    font=("Arial", 15),
    bg='#d2632c',
    fg='#24061d',
    width=12,
    height=1
)
botao_plan.pack(pady=20, padx=10, side=RIGHT)


caixa_retorno = Frame(
    infos,
    bg='#78243f',
)
caixa_retorno.pack(side=BOTTOM)

retorno= Label(
    font=("Arial", 15),
    bg='#d2632c',
    fg='#24061d',
    text=":)"
)
retorno.pack()


tela.mainloop()

