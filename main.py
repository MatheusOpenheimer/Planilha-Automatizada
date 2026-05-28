import pandas as pd
import tkinter as tk
from tkinter import filedialog
from openpyxl import load_workbook
import win32com.client as win32
import os
import sys


caminho_planilha = ""
caminho_forms = ""
def selecionar_forms():
    global caminho_forms

    caminho_forms = filedialog.askopenfilename(filetypes=[("Arquivos Excel", "*.xlsx")]) #pega o caminho
    campo_forms.config(text=f"ORIGEM:\n{caminho_forms}")

def selecionar_planilha():
    global caminho_planilha

    caminho_planilha = filedialog.askopenfilename(filetypes=[("Arquivos Excel", "*.xlsm")]) #pega o caminho
    campo_planilha.config(text=f"DESTINO:\n{caminho_planilha}")
    

def executar():
    # LER DADOS
    origem = pd.read_excel(caminho_forms)
    # ABRIR EXCEL REAL
    excel = win32.Dispatch("Excel.Application")
    excel.Visible = True

    workbook = excel.Workbooks.Open(caminho_planilha)
    aba = workbook.Worksheets("4.01")

    
    # PRÓXIMA LINHA
    linha = 9
    while aba.Range(f"G{linha}").Text != "":
        linha += 1
        
    perguntas = int(campo_perguntas.get())
    coluna_final = 11 + perguntas
    pesquisa = campo_pesquisa.get()
    # ESCREVER DADOS
    for _, linha_forms in origem.iterrows(): #linha_forms recebe cada linha da origem percorrida
       
        aba.Cells(linha, 5).Value = linha_forms["Nome completo: (opcional)"] #7 = indice de 8 #percorre a linha que ta no loop, na COLUNA 8 | NOME
        aba.Cells(linha, 7).Value = linha_forms["Setor:"] #percorre a linha que ta no loop, COLUNA 7 | SETOR
        aba.Cells(linha, 8).Value = pesquisa
        
        coluna = 9
        for valor in linha_forms.iloc[11:coluna_final]:
            aba.Cells(linha, coluna).Value = valor
            coluna += 1

        
        
        linha += 1

    # SALVAR
    workbook.Save()
    workbook.Close()
    excel.Quit()

def caminho_recurso(relativo):
    return os.path.join(os.path.abspath("."), relativo)     
#Janela
janela = tk.Tk()
janela.title("Planilha automatizada")
janela.geometry("600x600")
janela.configure(bg="#739ab9")
janela.iconbitmap(caminho_recurso("assets/AutomatizaX.ico"))

#pega a origem FORMS
label_forms = tk.Label(janela, text="INSIRA A PLANILHA DO FORMULÁRIO:", font=("Times New Roman", 14, "bold"), bg="#739ab9")
label_forms.pack(pady=2)
botao_forms = tk.Button(janela, text="SELECIONAR ORIGEM", bg="#113047", fg="white",
    activebackground="#476781",
    activeforeground="white", 
    command= selecionar_forms)
botao_forms.pack(pady=5)

#Pega o destino PLANILHA COMPLETA
label_main = tk.Label(janela, text="INSIRA A PLANILHA PRINCIPAL:", font=("Times New Roman", 14, "bold"), bg="#739ab9")
label_main.pack(pady=5)
botao_main = tk.Button(janela, text="SELECIONAR DESTINO", bg="#113047", fg="white", activebackground="#476781",
    activeforeground="white", command = selecionar_planilha)
botao_main.pack(pady=5)

label_perguntas = tk.Label(janela, text="INSIRA A QUANTIDADE DE QUESTÕES:", font=("Times New Roman", 14, "bold"), bg="#739ab9")
label_perguntas.pack(pady=5)
campo_perguntas = tk.Entry(janela)
campo_perguntas.pack(pady=5)

label_pesquisa = tk.Label(janela, text="INSIRA A PESQUISA:", font=("Times New Roman", 14, "bold"), bg="#739ab9")
label_pesquisa.pack(pady=5)
campo_pesquisa = tk.Entry(janela)
campo_pesquisa.pack(pady=5)


botao_executar = tk.Button(janela,text="EXECUTAR", bg="#6d120b", fg="white",
    activebackground="#af1207",
    command=executar
)
botao_executar.pack(pady=20)

#mostra origem
campo_forms = tk.Label(janela, text="ORIGEM:\nNenhum arquivo selecionado", wraplength=100, bg="#739ab9")
campo_forms.place(x=10, y=395)

#mostra destino
campo_planilha = tk.Label(janela, text="DESTINO:\nNenhum arquivo selecionado", wraplength=100, bg="#739ab9")
campo_planilha.place(x=490, y=395)


janela.mainloop()