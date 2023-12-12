import tkinter as tk
from tkinter import filedialog
import os
import subprocess
import tkinter.messagebox as msgbox

def execute_script():
    selected_script = script_listbox.get(script_listbox.curselection())
    if selected_script.endswith(".py"):
        subprocess.Popen(["python", selected_script])

def open_file_dialog():
    folder_path = filedialog.askdirectory()  # Permite ao usuário escolher o diretório
    script_listbox.delete(0, tk.END)  # Limpa a lista atual
    load_scripts_from_folder(folder_path)

def load_scripts_from_folder(folder_path):
    if not folder_path:
        return
    script_files = [f for f in os.listdir(folder_path) if f.endswith(".py")]
    for script in script_files:
        script_listbox.insert(tk.END, script)

# Crie uma janela
window = tk.Tk()
window.title("Interface de Automações Web - Ativy")

# Adicione uma cor de fundo à janela
window.configure(bg="purple")

# Crie uma barra de menu
menubar = tk.Menu(window)
window.config(menu=menubar)

# Submenu para abrir um diretório
file_menu = tk.Menu(menubar, tearoff=0)
menubar.add_cascade(label="Menu", menu=file_menu)
file_menu.add_separator()
file_menu.add_command(label="Sair", command=window.quit)

# Obtenha a lista de arquivos .py no diretório atual
current_directory = os.getcwd()
script_files = [f for f in os.listdir(current_directory) if f.endswith(".py") and f.startswith("indicador") 
                and "callface" not in f]

# Crie uma lista para exibir os arquivos .py
script_listbox = tk.Listbox(window, bg="purple", fg="white", font=("Arial", 12))
for script in script_files:
    script_listbox.insert(tk.END, script)

# Adicione uma barra de rolagem à lista
scrollbar = tk.Scrollbar(window)
script_listbox.config(yscrollcommand=scrollbar.set)
scrollbar.config(command=script_listbox.yview)

# Crie um botão para executar o script selecionado
execute_button = tk.Button(window, text="Executar Automação .py", command=execute_script)

# Carregue uma imagem como logotipo
logo_image = tk.PhotoImage(file="Logo Ativy.png")  # Substitua "logo.png" pelo caminho da sua imagem
logo_image = logo_image.subsample(4, 4)
logo_label = tk.Label(window, image=logo_image)

# Coloque a lista, o botão e o logotipo na janela
script_listbox.pack(side=tk.LEFT, padx=10, pady=10)
scrollbar.pack(side=tk.LEFT, fill=tk.Y)
execute_button.pack(side=tk.LEFT, padx=10, pady=10)
logo_label.pack(side=tk.RIGHT, padx=10, pady=10)

# Inicie o loop principal da interface gráfica
window.mainloop()