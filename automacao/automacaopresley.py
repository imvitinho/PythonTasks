#!/usr/bin/env python
# coding: utf-8

# In[27]:


import pandas as pd
import pyautogui
import pyperclip
import logging
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

pyautogui.PAUSE = 0.1

# Configuração de logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Definindo a aba da planilha
ABA_PLANILHA = "Produto Detalhado"  # Defina o nome correto da aba

# Coordenadas para automação
COORDENADAS = {
    'F4': (269, 80),
    'primeira_aba': (886, 675),
    'modificar_valor': (649, 889),
    'adicionar_info': (75, 106),
    'avancar': (919, 570),
    'segurado_contato': (518, 894)
}

# Função para carregar a planilha
def carregar_planilha(caminho_planilha):
    try:
        with pd.ExcelFile(caminho_planilha) as xls:
            df = pd.read_excel(xls, sheet_name=ABA_PLANILHA, skiprows=3)
            df.rename(columns={"Unnamed: 1": "Veiculo", "Unnamed: 2": "Valor"}, inplace=True)
        return df
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao carregar a planilha: {str(e)}")
        logging.error(f"Erro ao carregar planilha: {str(e)}")
        return None

# Função para buscar veículo no DataFrame
def buscar_veiculo(event=None):
    veiculo = entrada_veiculo.get()
    if not veiculo:
        messagebox.showwarning("Aviso", "Por favor, insira o nome de um veículo.")
        return

    if caminho_planilha == "":
        messagebox.showerror("Erro", "Por favor, selecione a planilha primeiro.")
        return

    df = carregar_planilha(caminho_planilha)
    if df is None:
        return

    limpar_opcoes()

    try:
        filtrados = df[df["Veiculo"].str.contains(veiculo, na=False, case=False)]
        if filtrados.empty:
            mostrar_feedback("Nenhum veículo encontrado.", "red")
        elif len(filtrados) == 1:
            selecionar_veiculo(filtrados.iloc[0])
        else:
            exibir_opcoes(filtrados)
    except KeyError:
        messagebox.showerror("Erro", "A coluna 'Veiculo' não foi encontrada na planilha.")
        logging.error("A coluna 'Veiculo' não foi encontrada.")

# Função para limpar as opções de veículos
def limpar_opcoes():
    for widget in frame_canvas.winfo_children():
        widget.destroy()

# Função para exibir opções de veículos filtrados centralizados
def exibir_opcoes(filtrados):
    limpar_opcoes()
    
    for i, row in filtrados.iterrows():
        veiculo = row["Veiculo"]
        valor = round(row["Valor"], 0)
        btn = ttk.Button(frame_canvas, text=f"{veiculo} - R${valor}", command=lambda r=row: selecionar_veiculo(r))
        btn.pack(pady=5, padx=10, fill = 'x')

# Função para selecionar veículo
def selecionar_veiculo(row):
    veiculo = row["Veiculo"]
    valor = round(row["Valor"], 0)

    mensagem = definir_mensagem(valor)

    pyperclip.copy(str(valor))
    automatizar_tarefas(valor, mensagem)

    limpar_opcoes()
    btn_buscar.config(state="disabled")
    mostrar_feedback("Seleção realizada. Por favor, pesquise um novo veículo.", "blue")

# Função para definir a mensagem baseada no valor
def definir_mensagem(valor):
    if valor == 130:
        return "#INF INTERNA# - FORNECIMENTO EM TRATATIVA# RECEBIDO AVIÃOZINHO DA LOJA - NAO CABE ALTERACAO ##"
    return "#INF INTERNA# - FORNECIMENTO EM TRATATIVA# RECEBIDO AVIÃOZINHO DA LOJA - FEITA ALTERACAO DE MO ##"

# Função de automação com PyAutoGUI
def automatizar_tarefas(valor_mo, mensagem):
    try: 
        pyautogui.click(*COORDENADAS['F4'])
        pyautogui.click(*COORDENADAS['modificar_valor'])
        pyautogui.hotkey('ctrl', 'v')
        pyautogui.click(*COORDENADAS['adicionar_info'])
        pyautogui.write(mensagem)
        pyautogui.moveTo(*COORDENADAS['avancar'])
        pyautogui.press('enter')
        pyautogui.press('f2')
        pyautogui.click(*COORDENADAS['segurado_contato'])
        pyautogui.press('insert')
        pyautogui.write("#INF INTERNA# CASO SEGURADO ENTRAR EM CONTATO, SOLICITAR AGENDAR SERVICO COM A LOJA - O.S LIBERADA #")
        pyautogui.click(*COORDENADAS['primeira_aba'])
    except Exception as e:
        logging.error(f"Ocorreu um erro durante a automação: {e}")
        messagebox.showerror("Erro", f"Ocorreu um erro durante a automação: {e}")

# Função para mostrar feedback na interface
def mostrar_feedback(mensagem, cor):
    label_feedback.config(text=mensagem, foreground=cor)

# Função para selecionar o arquivo da planilha
def selecionar_planilha():
    global caminho_planilha
    caminho_planilha = filedialog.askopenfilename(filetypes=[("Arquivos Excel", "*.xlsx")])
    if caminho_planilha:
        label_feedback.config(text="Planilha carregada com sucesso!", foreground="green")
    else:
        label_feedback.config(text="Nenhuma planilha selecionada.", foreground="red")

# Variável global para armazenar o caminho da planilha
caminho_planilha = ""

# Configuração da interface gráfica
root = tk.Tk()
root.title("Filtro de Veículos")
root.geometry("600x500")

btn_selecionar_planilha = ttk.Button(root, text="Selecionar Planilha", command=selecionar_planilha, width=20)
btn_selecionar_planilha.pack(pady=15)

tk.Label(root, text="Digite o nome do veículo:").pack(pady=10)
entrada_veiculo = tk.Entry(root, font=("Arial", 12), width=40)
entrada_veiculo.pack(pady=10)
entrada_veiculo.bind("<Return>", buscar_veiculo)

btn_buscar = ttk.Button(root, text="Buscar", command=buscar_veiculo, width=20)
btn_buscar.pack(pady=15)

label_feedback = tk.Label(root, text="", font=("Arial", 10, "italic"))
label_feedback.pack(pady=5)

frame_opcoes = tk.Frame(root)
frame_opcoes.pack(pady=10, fill="both", expand=True)

frame_canvas = tk.Frame(frame_opcoes)
frame_canvas.pack(fill="both", expand=True)

root.mainloop()


# In[ ]:




