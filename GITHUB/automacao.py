import pandas as pd
import pyautogui
import pyperclip
import logging
import tkinter as tk
from tkinter import ttk, messagebox
 
# Configuração de logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
 
# Caminho do arquivo Excel
CAMINHO_PLANILHA = "C:\Users\victor.machado\Downloads\Valor MO 2025 (1).xlsx"
ABA_PLANILHA = "Produto Detalhado"
 
# Coordenadas para automação
COORDENADAS = {
    'linha_venus': (572, 913),
    'primeira_aba': (886, 675),
    'modificar_valor': (649, 889),
    'adicionar_info': (75, 106),
    'avancar': (919, 570),
    'segurado_contato': (518, 894)
}
 
# Função para carregar a planilha
def carregar_planilha():
    try:
        df = pd.read_excel(CAMINHO_PLANILHA, sheet_name=ABA_PLANILHA, skiprows=3)
        df.rename(columns={"Unnamed: 1": "Veiculo", "Unnamed: 2": "Valor"}, inplace=True)
        return df
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao carregar a planilha: {str(e)}")
        return None
 
# Função para buscar veículo no DataFrame
def buscar_veiculo(event=None):
    veiculo = entrada_veiculo.get()
    if not veiculo:
        messagebox.showwarning("Aviso", "Por favor, insira o nome de um veículo.")
        return
 
    # Carregar e buscar veículo
    df = carregar_planilha()
    if df is None:
        return
 
    # Limpar opções anteriores antes de exibir novas
    for widget in frame_canvas.winfo_children():
        widget.destroy()
 
    try:
        filtrados = df[df["Veiculo"].str.contains(veiculo, na=False, case=False)]
        if filtrados.empty:
            messagebox.showinfo("Resultado", "Nenhum veículo encontrado.")
            label_feedback.config(text="Nenhum veículo encontrado.", foreground="red")
            scrollbar.pack_forget()  # Esconder a scrollbar
        elif len(filtrados) == 1:
            selecionar_veiculo(filtrados.iloc[0])
            scrollbar.pack_forget()  # Esconder a scrollbar se apenas um veículo
        else:
            exibir_opcoes(filtrados)
            scrollbar.pack(side="right", fill="y")  # Mostrar a scrollbar se houver mais de um veículo
 
    except KeyError:
        messagebox.showerror("Erro", "A coluna 'Veiculo' não foi encontrada na planilha.")
    finally:
        # Reabilitar o botão de busca
        btn_buscar.config(state="normal")
        label_feedback.config(text="")
 
# Função para exibir opções de veículos filtrados com scroll
def exibir_opcoes(filtrados):
    for i, row in filtrados.iterrows():
        veiculo = row["Veiculo"]
        valor = round(row["Valor"], 0)
        # Adicionar botão com mais detalhes
        btn = ttk.Button(frame_canvas, text=f"{veiculo} - R${valor}", command=lambda r=row: selecionar_veiculo(r))
        btn.grid(row=i, column=0, sticky="w", padx=10, pady=2)
 
    # Atualizar o tamanho da área de rolagem
    canvas.config(scrollregion=canvas.bbox("all"))
 
# Função para selecionar veículo
def selecionar_veiculo(row):
    veiculo = row["Veiculo"]
    valor = round(row["Valor"], 0)
 
    if valor == 130:
        mensagem = "## INF INTERNA FORNECIMENTO - NAO CABE ALTERACAO ##"
    else:
        mensagem = "## INF INTERNA FORNECIMENTO - ALTERACAO DE MO ##"
 
    copiar_para_area_transferencia(valor)
    automatizar_tarefas(valor, mensagem)
 
    # Limpar a lista de opções após a seleção do veículo
    for widget in frame_canvas.winfo_children():
        widget.destroy()
 
    # Desabilitar o botão de busca para forçar o usuário a realizar uma nova pesquisa
    btn_buscar.config(state="disabled")
    label_feedback.config(text="Seleção realizada. Por favor, pesquise um novo veículo.", foreground="blue")
 
# Função para copiar valor para a área de transferência
def copiar_para_area_transferencia(valor):
    pyperclip.copy(str(valor))
 
# Função de automação com PyAutoGUI
def automatizar_tarefas(valor_mo, mensagem):
    try:
        pyautogui.click(*COORDENADAS['linha_venus'])
        pyautogui.press('insert')
        pyautogui.write("# INF INTERNA - FORNECIMENTO EM TRATATIVA# RECEBIDO AVIÃOZINHO DA LOJA")
 
        pyautogui.click(*COORDENADAS['primeira_aba'])
        pyautogui.press('f4')
 
        pyautogui.click(*COORDENADAS['modificar_valor'])
        pyautogui.hotkey('ctrl', 'v')
 
        pyautogui.click(*COORDENADAS['adicionar_info'])
        pyautogui.write(mensagem)
 
        pyautogui.moveTo(*COORDENADAS['avancar'])
        pyautogui.press('enter')
        pyautogui.press('f2')
 
        pyautogui.click(*COORDENADAS['segurado_contato'])
        pyautogui.press('insert')
        pyautogui.write("# CASO SEGURADO ENTRAR EM CONTATO, SOLICITAR AGENDAR SERVICO COM A LOJA - O.S LIBERADA #")
 
        pyautogui.click(*COORDENADAS['primeira_aba'])
    except Exception as e:
        logging.error(f"Ocorreu um erro durante a automação: {e}")
 
# Configuração da interface gráfica
root = tk.Tk()
root.title("Filtro de Veículos")
 
# Definir o tamanho da janela
root.geometry("600x500")  # Aumentar o tamanho da janela principal
 
# Entrada para o nome do veículo
tk.Label(root, text="Digite o nome do veículo:").pack(pady=10)
entrada_veiculo = tk.Entry(root, font=("Arial", 12), width=40)  # Aumentar largura da entrada
entrada_veiculo.pack(pady=10)
 
# Configurar para o Enter ativar o botão Buscar
entrada_veiculo.bind("<Return>", buscar_veiculo)
 
# Botão para buscar veículos
btn_buscar = ttk.Button(root, text="Buscar", command=buscar_veiculo, width=20)  # Aumentar o tamanho do botão
btn_buscar.pack(pady=15)
 
# Label de feedback
label_feedback = tk.Label(root, text="", font=("Arial", 10, "italic"))
label_feedback.pack(pady=5)
 
# Frame para exibir as opções de veículos com scrollbar
frame_opcoes = tk.Frame(root)
frame_opcoes.pack(pady=10, fill="both", expand=True)
 
canvas = tk.Canvas(frame_opcoes)
canvas.pack(side="left", fill="both", expand=True)
 
scrollbar = ttk.Scrollbar(frame_opcoes, orient="vertical", command=canvas.yview)
scrollbar.pack(side="right", fill="y")  # Mostrar a scrollbar no lado direito
 
canvas.config(yscrollcommand=scrollbar.set)
 
frame_canvas = tk.Frame(canvas)
canvas.create_window((0, 0), window=frame_canvas, anchor="nw")
 
# Ajustando o estilo da barra de rolagem para uma aparência mais discreta
scrollbar.config(style="TScrollbar")
 
# Definindo o estilo customizado da barra de rolagem
style = ttk.Style()
style.configure("TScrollbar", thickness=8, gripcount=0)  # Definindo a espessura e aparência da barra
 
# Iniciar o loop da interface gráfica
root.mainloop()