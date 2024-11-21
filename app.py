import tkinter as tk
from tkinter import messagebox
from tkinter import PhotoImage
from threading import Thread
from main import executar_relatorio_completo as executar_relatorio

def run_main():
    path = entry_path.get().strip()
    
    if not path:
        messagebox.showerror("Erro", "Por favor, insira um caminho válido.")
        return

    try:
        def task():
            sucesso = executar_relatorio(path)
            if sucesso:
                messagebox.showinfo("Sucesso", "O script foi executado com sucesso! O relatório gerado inclui:\nItems do Estoque Lote\nItems do Vendas de Produtos\nEstoques\nItens Ordens Compra\nVendas de Produtos\n Curva ABC.")
            else:
                messagebox.showerror("Erro", "Ocorreu um erro ao executar o script. Verifique os logs para mais detalhes.")
        
        Thread(target=task).start()
        
    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro inesperado ao tentar executar o script:\n{e}")

# Criação da janela principal
root = tk.Tk()
root.title("Gerador de relatórios")
root.geometry("350x180")
root.configure(bg="#34495E")

# Carrega o ícone do aplicativo (usando .png)
try:
    icon = PhotoImage(file="3979302.png")  # Certifique-se de ter um arquivo '3979302.png' no mesmo diretório
    root.iconphoto(True, icon)
except Exception as e:
    print("Erro ao carregar o ícone:", e)

# Estilos modernizados
style_font = ("Helvetica Neue", 11)
color_bg = "#34495E"
color_fg = "#ECF0F1"
color_button = "#1ABC9C"
color_button_hover = "#16A085"

# Label para instrução
label_instruction = tk.Label(root, text="Informe o Caminho do Relatório:", bg=color_bg, fg=color_fg, font=style_font)
label_instruction.pack(pady=15)

# Campo de entrada para o caminho
entry_path = tk.Entry(root, width=40, font=style_font, bg="#2C3E50", fg=color_fg, insertbackground=color_fg, relief="flat")
entry_path.pack(pady=5)
entry_path.insert(0, r"\\WKRadar-SRV\Relatórios")

# Função para efeito hover no botão
def on_enter(e):
    button_start['background'] = color_button_hover

def on_leave(e):
    button_start['background'] = color_button

# Botão para executar o script com estilo flat e hover
button_start = tk.Button(root, text="Executar", command=run_main, bg=color_button, fg="#FFFFFF", font=style_font, relief="flat", cursor="hand2")
button_start.pack(pady=30)
button_start.bind("<Enter>", on_enter)
button_start.bind("<Leave>", on_leave)

# Inicia o loop da interface
root.mainloop()
