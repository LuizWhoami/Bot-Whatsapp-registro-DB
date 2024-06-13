
from tkinter import *
from tkinter import messagebox
import pandas as pd
from openpyxl.workbook import Workbook
from ttkbootstrap.style import Style
import sqlite3


#cores -----------------------------
co0 = "#f0f3f5"  # Preta / black
co1 = "#feffff"  # branca / white
co2 = "#3fb5a3"  # verde / green
co3 = "#38576b"  # valor / value
co4 = "#403d3d"   # letra / letters
co5 = "#4b81a6" # azul

janela = Tk()
janela.title("")
janela.geometry("300x350")
estilo = Style(theme="darkly")
janela.resizable(width=False, height=False)

#dividir tela
Frame_cima = Frame(janela, width=310, height=50, bg=co5, relief="flat")
Frame_cima.grid(row=0, column=0, pady=1, padx=0, sticky=NSEW)

Frame_baixo = Frame(janela, width=310, height=450, bg=co1, relief="flat")
Frame_baixo.grid(row=1, column=0, pady=1, padx=0, sticky=NSEW) 

#frame cima

Label_nome = Label(Frame_cima, text="Registro", anchor=NE,width=10, font=("Ivy 25"), bg=co5, fg=co4)
Label_nome.place(x=5, y=5)

#nome
Label_nome = Label(Frame_baixo, text="Nome", anchor=NW,width=10, font=("Ivy 10"), bg=co1, fg=co4)
Label_nome.place(x=15, y=20)
Entry_nome1 = Entry(Frame_baixo, width=25, justify="left", font=("", 15), highlightthickness=1, relief="solid" )
Entry_nome1.place(x=15, y=45)

Label_telefone = Label(Frame_baixo, text="email", anchor=NW,width=10, font=("Ivy 10"), bg=co1, fg=co4)
Label_telefone.place(x=15, y=80)
Entry_telefone = Entry(Frame_baixo, width=25, justify="left", font=("", 15), highlightthickness=1, relief="solid" )
Entry_telefone.place(x=15, y=105)

def cadastrar_clientes():
    #Copia da criação do banco
    conexao = sqlite3.connect("Banco_Clientes.db")

    cursor = conexao.cursor()
    #tabela
    cursor.execute("INSERT INTO clientes VALUES (:nome, :email)",
                   {
                       'nome':Entry_nome1.get(),
                       'email':Entry_telefone.get()
                   }
                   )
    
    conexao.commit()

    conexao.close()
    Entry_nome1.delete(0, 'end')
    Entry_telefone.delete(0, 'end')

def exportar_clientes():
    #Conectando ao banco
    conexao = sqlite3.connect("Banco_Clientes.db")

    cursor = conexao.cursor()
    #Selecionando tabela
    cursor.execute("SELECT *, oid FROM clientes")
    clientes_cadastrados = cursor.fetchall()
    clientes_cadastrados = pd.DataFrame(clientes_cadastrados, columns=['nome', 'telefone', 'vencimento', 'ID_banco'])
    clientes_cadastrados.to_excel('clientes.xlsx')
    conexao.commit()
    conexao.close()
    
Botão = Button(Frame_baixo, text="Cadastrar", command=cadastrar_clientes, width=39, height=2, font=("Ivy 8 bold"), bg=co2, fg=co0, relief=RAISED, overrelief=RIDGE)
Botão.place(x=15, y=200)
Botão = Button(Frame_baixo, text="Exportar", command=exportar_clientes, width=39, height=2, font=("Ivy 8 bold"), bg=co2, fg=co0, relief=RAISED, overrelief=RIDGE)
Botão.place(x=15, y=250)



janela.mainloop()