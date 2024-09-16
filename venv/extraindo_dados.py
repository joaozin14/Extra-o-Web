#Importando as bilbiotecas
from selenium import webdriver as opcoes_selenium
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from  time import sleep
from tkinter import *
from tkinter import ttk
from tkinter import messagebox
from openpyxl import load_workbook

#Definindo a janela com tkinter e criando a treeview
janela = Tk()
janela.title("Janela dados web")
estilo = ttk.Style()
estilo.theme_use("alt")
estilo.configure(".", font="Arial 20", rowheight=40)
treeview = ttk.Treeview(janela, columns=(1, 2, 3), show="headings")

treeview.column("1", width=100, anchor=CENTER)
treeview.heading("1", text="ID")

treeview.column("2", width=500, anchor=CENTER)
treeview.heading("2", text="Processos")

treeview.column("3", width=200, anchor=CENTER)
treeview.heading("3", text="Data")
treeview.grid(row=0, column=0, columnspan=8, sticky="NSEW")

#Craindo a função que exporta os dados para a planilha do excel
def exportar():
     try:
        workbook = load_workbook(filename="C:\\Users\\joaoh\\OneDrive\\Área de Trabalho\\Tudo de  Python\\extraindo_tabelas\\venv\\Extracao.xlsx")
        sheet = workbook["Extrações"]
        sheet.delete_rows(idx=1, amount=10000)

        for numero_linha in treeview.get_children():
            linha = treeview.item(numero_linha)["values"]
            sheet.append(linha)
        workbook.save(filename="C:\\Users\\joaoh\\OneDrive\\Área de Trabalho\\Tudo de  Python\extraindo_tabelas\\venv\\Dados_extraidos.xlsx")
        messagebox.showinfo("Atenção", "Dados exportados com sucesso")
     except:
        messagebox.showinfo("Atenção", "Deu algum erro na exportação!")

#Criando o botão que exporta
botao_exportar = Button(janela, text="Exportar", font="Arial 15", background="black", foreground="white", command=exportar)
botao_exportar.grid(row=7, sticky="NSEW", column=7, padx=5, pady=5)

#Definindo as a inicialização do navegador e abrindo o link do site
opcoes_chrome = Options()
opcoes_chrome.add_argument("--start-maximized")
#opcoes_chrome.add_argument("--headless")
navegador = opcoes_selenium.Chrome(options=opcoes_chrome)
navegador.get("https://rpachallengeocr.azurewebsites.net/")

sleep(2)

i = 1
#Usando o while para o loop até encerrar o número de páginas do site
while i < 4:
    #Pegando os elemntos com XPATH
    elemento_tabela = navegador.find_element(By.XPATH, '//*[@id="tableSandbox"]')
    linhas = elemento_tabela.find_elements(By.TAG_NAME, "tr")
    colunas = elemento_tabela.find_elements(By.TAG_NAME, "td")
    
    #Usando o for para percorrer as linhas e inserir na treeview
    for linha_atual in linhas:
        texto = linha_atual.text
        if texto[0] != "#":
            texto2 = texto.split(" ")
            print(f"coluna 1: {texto2[0]}")
            print(f"coluna 2: {texto2[1]}")
            print(f"coluna 3: {texto2[2]}")
            treeview.insert("", "end", values=(str(texto2[0]),
                            str(texto2[1]),
                            str(texto2[2])
                            ))

    sleep(2)

    navegador.find_element(By.XPATH, '//*[@id="tableSandbox_next"]').click()
    
    i += 1
    sleep(2)

#Fechando o navegador e abrindo a janela do Tkinter
navegador.close()
janela.mainloop()