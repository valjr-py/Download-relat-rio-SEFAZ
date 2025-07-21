from time import sleep
import openpyxl
from selenium import webdriver
from selenium.webdriver.common.by import By
import tkinter as tk
from tkinter import messagebox
from datetime import datetime
import pyautogui

def format_date(event=None, entry=None):
    text = entry.get()
    if len(text) == 2 or len(text) == 5:
        entry.insert(tk.END, '/')
    elif len(text) > 10:
        entry.delete(10, tk.END)

def submit():
    global start_date_str, end_date_str
    try:
        start_date = datetime.strptime(entry1.get(), '%d/%m/%Y')
        end_date = datetime.strptime(entry2.get(), '%d/%m/%Y')
        if start_date > end_date:
            messagebox.showerror("Erro", "A data inicial não pode ser maior que a data final.")
        else:
            start_date_str = entry1.get()
            end_date_str = entry2.get()
            print(f"Data Inicial: {start_date_str}")
            print(f"Data Final: {end_date_str}")
            root.destroy()  
    except ValueError:
        messagebox.showerror("Erro", "Por favor, insira datas válidas no formato DD/MM/AAAA.")

root = tk.Tk()
root.title("Data da competência")
root.geometry("400x250") 

label1 = tk.Label(root, text="Data Inicial (DD/MM/AAAA):", font=("Arial", 12))
label1.pack(pady=(20, 5))

entry1 = tk.Entry(root, width=30, font=("Arial", 12))
entry1.pack()
entry1.bind('<KeyRelease>', lambda event: format_date(event, entry1))  

label2 = tk.Label(root, text="Data Final (DD/MM/AAAA):", font=("Arial", 12))
label2.pack(pady=(20, 5))

entry2 = tk.Entry(root, width=30, font=("Arial", 12))
entry2.pack()
entry2.bind('<KeyRelease>', lambda event: format_date(event, entry2))  

submit_button = tk.Button(root, text="Enviar", font=("Arial", 12), bg="lightblue", command=submit)
submit_button.pack(pady=30)

root.mainloop()


planilha_empresas = openpyxl.load_workbook('empresas.xlsx')
pagina_empresas = planilha_empresas['Plan1']



for linha in pagina_empresas.iter_rows(min_row=2, values_only=True):

    empresa, caceal, senha = linha

    driver = webdriver.Chrome()
    driver.get('https://contribuinte.sefaz.al.gov.br/malhafiscal/#/')

    conta = driver.find_element(By.XPATH, '//a[@id="account-menu"]')
    conta.click()
    entrar = driver.find_element(By.XPATH, '//a[@id="login"]')
    entrar.click()
    usuario = driver.find_element(By.XPATH, '//input[@id="username"]')
    usuario.send_keys(caceal)
    senha_user = driver.find_element(By.XPATH, '//input[@id="password"]')
    senha_user.send_keys(senha)
    botao_entrar = driver.find_element(By.XPATH, '//button[@type="submit"]')
    botao_entrar.click()
    sleep(10)
    relatorios = driver.find_element(By.XPATH, '//a[@id="relatorio"]')
    relatorios.click()
    contribuinte = driver.find_element(By.XPATH, '//a[@routerlink="relatorio-contribuinte"]')
    contribuinte.click()
    sleep(6)
    data_inicial = driver.find_element(By.XPATH, '//input[@name="dataInicial"]')
    data_inicial.send_keys(start_date_str)
    data_final = driver.find_element(By.XPATH, '//input[@name="dataFinal"]')
    data_final.send_keys(str(end_date_str))
    botao_emititir = driver.find_element(By.XPATH, '(//button[@class="btn btn-info btn-sm"])[5]')
    botao_emititir.click()
    sleep(4)
    pyautogui.click(x=1191, y=209)
    sleep(4)
    pyautogui.click(x=570, y=79)
    sleep(1)
    pyautogui.write(r'G:\Meu Drive\Pcontabilidade - Setor Fiscal\11. Geração de Tarefas\1. Relatorio de Entradas\2025\01.2025')
    sleep(10)
    pyautogui.press('enter')
    sleep(10)
    pyautogui.click(x=767, y=560)
    driver.quit()
