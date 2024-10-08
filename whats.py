import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import pywhatkit as kit
import time


def ler_arquivo_excel(filepath):
    df = pd.read_excel(filepath, usecols=[0], nrows=500)
    numeros = []
    for numero in df.iloc[:, 0].tolist():
        if isinstance(numero, (int, float)):
            numero_str = str(int(numero)).strip() 
            if len(numero_str) == 10:  
                numero_str = '55' + numero_str  
            elif len(numero_str) == 11:  
                numero_str = '55' + numero_str  
            numeros.append(numero_str)
    return numeros


def enviar_mensagens(numeros, mensagem):
    for i in range(0, len(numeros), 5):
        lote = numeros[i:i + 5]
        for numero in lote:
            try:
                
                kit.sendwhatmsg_instantly(f'+55{numero}', mensagem, 5)  
                print(f"Mensagem enviada para {numero}")
                time.sleep(2)  
            except Exception as e:
                print(f"Erro ao enviar mensagem para {numero}: {e}")
        
        if i + 5 < len(numeros):
            print("Aguardando 2 minutos...")
            time.sleep(120)

    messagebox.showinfo("Concluído", "Todas as mensagens foram enviadas!")


def iniciar_envio():
    filepath = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    if filepath:
        numeros = ler_arquivo_excel(filepath)
        mensagem = entry_mensagem.get()
        if mensagem:
            enviar_mensagens(numeros, mensagem)
        else:
            messagebox.showwarning("Aviso", "Por favor, insira uma mensagem para enviar.")


root = tk.Tk()
root.title("Disparador de Mensagens")
root.geometry("400x300")
root.configure(bg="#1e2024")

# Estilo moderno para os widgets
estilo_bg = "#1e2024"
estilo_fg = "#ffffff"
estilo_btn_bg = "#2a2d34"
estilo_btn_fg = "#ffffff"
estilo_entry_bg = "#2a2d34"
estilo_entry_fg = "#ffffff"

label_instrucoes = tk.Label(root, text="Escolha um arquivo Excel com números de telefone:", bg=estilo_bg, fg=estilo_fg)
label_instrucoes.pack(pady=10)

label_mensagem = tk.Label(root, text="Digite a mensagem a ser enviada:", bg=estilo_bg, fg=estilo_fg)
label_mensagem.pack(pady=10)

entry_mensagem = tk.Entry(root, width=50, bg=estilo_entry_bg, fg=estilo_entry_fg, insertbackground=estilo_fg)
entry_mensagem.pack(pady=10)

botao = tk.Button(root, text="Selecionar Arquivo e Enviar Mensagens", command=iniciar_envio, bg=estilo_btn_bg, fg=estilo_btn_fg)
botao.pack(pady=20)

root.mainloop()
