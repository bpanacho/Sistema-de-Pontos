import tkinter as tk
from tkinter import messagebox, ttk
import subprocess

def capture_and_register():
    employee_name = name_entry.get()
    if not employee_name:
        messagebox.showerror("Erro!", "Por favor, digite seu nome.")
        return
    
    try:
        # Executa o main.py com o nome do funcionário como argumento
        subprocess.run(['python', 'ocr2.py', employee_name], check=True)
        messagebox.showinfo("Sucesso", f"Ponto registrado com sucesso para {employee_name}!")
    except subprocess.CalledProcessError:
        messagebox.showerror("Erro", "Ocorreu um erro ao registrar o ponto.")

# Criando a janela principal
root = tk.Tk()
root.title("Registro de Ponto")
root.configure(bg='white')  # Define o fundo da janela como branco

# Estilo para os widgets ttk (botões)
style = ttk.Style()

# Configuração do botão
style.configure('Purple.TButton',
                foreground='white',  
                background='#127244',  
                font=('Helvetica', 12, 'bold'),  
                padding=10,  
                borderwidth=2, 
                bordercolor='#6ebc82', 
                relief='raised')  

style.map('Purple.TButton', background=[('active', '#093d24')], relief=[('pressed', 'sunken')])  # Estilo de relevo quando pressionado

# Frame para organizar os widgets
frame = tk.Frame(root, bg='white')
frame.pack(padx=20, pady=20)

# Adicionando imagem
image = tk.PhotoImage(file='/home/bianca-panacho/Desktop/ponto/logo.png') 
image_label = tk.Label(frame, image=image, bg='white')
image_label.grid(row=0, columnspan=2, padx=10, pady=10)

# Nome do funcionário
tk.Label(frame, text="Funcionário:", bg='white').grid(row=1, column=0, padx=10, pady=10)
name_entry = tk.Entry(frame, width=27)  # Aumenta a largura da caixa 
name_entry.grid(row=1, column=1, padx=10, pady=10)

# Botão para capturar e registrar ponto
capture_button = ttk.Button(root, text="Registrar Ponto", style='Purple.TButton', command=capture_and_register)
capture_button.pack(pady=10)

root.mainloop()
