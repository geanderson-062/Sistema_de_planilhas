import tkinter as tk
from tkinter import ttk, messagebox  # Importação adicional
import customtkinter
from tkinter import filedialog
import openpyxl

# Configuração da janela Tkinter usando customtkinter
root = customtkinter.CTk()
root.geometry("1920x1080")
root.title("Sistema de Planilhas")
root.resizable(False, False)

# Definindo o tema
customtkinter.set_appearance_mode("dark")
customtkinter.set_default_color_theme("green")

# Variável para armazenar os dados (fora das funções)
dados_salvos = []


def adicionar_dados():
    # Adiciona dados à lista global e limpa as entradas
    dados = [id_entry.get(), cliente_entry.get(), produto_entry.get(), data_entry.get()]
    dados_salvos.append(dados)
    id_entry.delete(0, tk.END)
    cliente_entry.delete(0, tk.END)
    produto_entry.delete(0, tk.END)
    data_entry.delete(0, tk.END)


def mostrar_tabela():
    # Cria uma nova janela
    janela_tabela = customtkinter.CTkToplevel(root)
    janela_tabela.geometry("800x600")
    janela_tabela.title("Dados Adicionados")

    # Criando a tabela na nova janela
    colunas = ("ID", "Cliente", "Produto", "Data")
    tabela = ttk.Treeview(janela_tabela, columns=colunas, show="headings")
    for col in colunas:
        tabela.heading(col, text=col)
    tabela.pack(expand=True, fill="both")

    # Adicionando dados existentes à tabela
    for dados in dados_salvos:
        tabela.insert("", "end", values=dados)


def salvar_excel():
    # Criando o workbook e adicionando os dados
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.append(["ID", "Cliente", "Produto", "Data"])  # Cabeçalhos
    for item in dados_salvos:
        sheet.append(item)  # Valores
    # Pedindo ao usuário para escolher onde salvar o arquivo
    file_path = filedialog.asksaveasfilename(defaultextension=".xlsx")
    if file_path:
        workbook.save(file_path)
        messagebox.showinfo(
            "Sucesso", "Arquivo Excel salvo com sucesso!"
        )  # Mensagem de sucesso


# Ajustes de centralização para o grid
root.grid_rowconfigure(0, weight=1)
root.grid_rowconfigure(10, weight=1)
root.grid_columnconfigure(0, weight=1)
root.grid_columnconfigure(4, weight=1)

# Título "Cadastrar"
titulo_label = customtkinter.CTkLabel(root, text="Cadastrar", font=("Roboto", 25))
titulo_label.grid(row=1, column=1, columnspan=2, pady=(10, 20))

# Campos de entrada com customtkinter
id_entry = customtkinter.CTkEntry(root, placeholder_text="ID")
id_entry.grid(row=2, column=1, padx=10, pady=5)
cliente_entry = customtkinter.CTkEntry(root, placeholder_text="Cliente")
cliente_entry.grid(row=2, column=2, padx=10, pady=5)
produto_entry = customtkinter.CTkEntry(root, placeholder_text="Produto")
produto_entry.grid(row=3, column=1, padx=10, pady=5)
data_entry = customtkinter.CTkEntry(root, placeholder_text="Data")
data_entry.grid(row=3, column=2, padx=10, pady=5)

# Botões estilizados
add_button = customtkinter.CTkButton(
    root,
    text="Adicionar",
    command=adicionar_dados,
    fg_color="gray",
    hover_color="gray20",
)
add_button.grid(row=4, column=1, padx=10, pady=10)

botao_tabela = customtkinter.CTkButton(
    root,
    text="Mostrar Tabela",
    command=mostrar_tabela,
    fg_color="gray",
    hover_color="gray20",
)
botao_tabela.grid(row=4, column=2, padx=10, pady=10)

save_button = customtkinter.CTkButton(
    root,
    text="Salvar como Excel",
    command=salvar_excel,
    fg_color="gray",
    hover_color="gray20",
)
save_button.grid(row=5, column=1, columnspan=2, pady=10)

# Executa a aplicação
root.mainloop()
