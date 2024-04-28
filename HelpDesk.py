#!/usr/bin/env python
# coding: utf-8

# In[2]:


import tkinter as tk
from tkinter import messagebox, ttk
import openpyxl
import random
import string


def fazer_login():
    """Função para realizar o login."""
    usuario = usuario_entry.get()
    senha = senha_entry.get()

    # Credenciais estáticas pré-definidas (apenas para exemplo)
    if usuario == "admin" and senha == "admin":
        messagebox.showinfo("Login bem-sucedido", "Login realizado com sucesso!")
        janela_login.destroy()  # Fechar a janela de login
        abrir_janela_principal()
    else:
        messagebox.showerror("Erro de login", "Credenciais inválidas. Tente novamente.")


def criar_ticket(assunto, descricao, prioridade, solicitante):
    """Função para criar um novo ticket de helpdesk."""
    try:
        workbook = openpyxl.load_workbook("tickets.xlsx")
    except FileNotFoundError:
        workbook = openpyxl.Workbook()

    sheet = workbook.active
    nova_linha = [assunto, descricao, prioridade, solicitante]  # Adicionar solicitante à nova linha

    sheet.append(nova_linha)
    workbook.save("tickets.xlsx")

    messagebox.showinfo("Sucesso", "Ticket criado com sucesso.")
    log_text.insert(tk.END, "Novo ticket criado.\n")


def abrir_janela_principal():
    """Função para abrir a janela principal do sistema."""
    global janela_principal
    janela_principal = tk.Tk()
    janela_principal.title("Sistema de Helpdesk")

    # Adicionar as etiquetas e as caixas de entrada para os dados do ticket
    id_label = tk.Label(janela_principal, text="ID do Ticket:")
    id_label.grid(row=0, column=0, padx=10, pady=5)
    id_entry = tk.Entry(janela_principal)
    id_entry.grid(row=0, column=1, padx=10, pady=5)

    assunto_label = tk.Label(janela_principal, text="Assunto:")
    assunto_label.grid(row=1, column=0, padx=10, pady=5)
    assunto_entry = tk.Entry(janela_principal)
    assunto_entry.grid(row=1, column=1, padx=10, pady=5)

    descricao_label = tk.Label(janela_principal, text="Descrição:")
    descricao_label.grid(row=2, column=0, padx=10, pady=5)
    descricao_entry = tk.Entry(janela_principal)
    descricao_entry.grid(row=2, column=1, padx=10, pady=5)

    # Caixa de texto para o log de eventos
    global log_text
    log_text = tk.Text(janela_principal, height=10, width=50)
    log_text.grid(row=3, columnspan=2, padx=10, pady=5)

    # Botão Criar Ticket
    botao_criar_ticket = tk.Button(janela_principal, text="Criar Ticket", command=abrir_janela_criar_ticket)
    botao_criar_ticket.grid(row=4, column=0, padx=10, pady=5)

    # Botão Editar
    botao_editar = tk.Button(janela_principal, text="Editar Ticket", command=abrir_janela_edicao)
    botao_editar.grid(row=4, column=1, padx=10, pady=5)

    janela_principal.mainloop()


def abrir_janela_criar_ticket():
    """Função para abrir a janela de criação de ticket."""
    janela_criar_ticket = tk.Toplevel(janela_principal)
    janela_criar_ticket.title("Criar Novo Ticket de Helpdesk")

    # Label e Entry para o Assunto
    assunto_label = tk.Label(janela_criar_ticket, text="Assunto:")
    assunto_label.grid(row=0, column=0, padx=10, pady=5)
    assunto_entry = tk.Entry(janela_criar_ticket)
    assunto_entry.grid(row=0, column=1, padx=10, pady=5)

    # Label e Entry para a Descrição
    descricao_label = tk.Label(janela_criar_ticket, text="Descrição:")
    descricao_label.grid(row=1, column=0, padx=10, pady=5)
    descricao_entry = tk.Entry(janela_criar_ticket)
    descricao_entry.grid(row=1, column=1, padx=10, pady=5)

    # Adicionando uma lista suspensa para selecionar a prioridade do ticket
    prioridade_label = tk.Label(janela_criar_ticket, text="Prioridade:")
    prioridade_label.grid(row=2, column=0, padx=10, pady=5)
    prioridade_var = tk.StringVar()
    prioridade_var.set("Baixa")  # Definir prioridade padrão como "Baixa"
    prioridade_dropdown = tk.OptionMenu(janela_criar_ticket, prioridade_var, "Baixa", "Média", "Alta")
    prioridade_dropdown.grid(row=2, column=1, padx=10, pady=5)

    # Adicionando uma lista suspensa para selecionar o solicitante do ticket
    solicitante_label = tk.Label(janela_criar_ticket, text="Solicitante:")
    solicitante_label.grid(row=3, column=0, padx=10, pady=5)
    solicitante_var = tk.StringVar()
    solicitante_var.set("TI")  # Definir solicitante padrão como "TI"
    solicitante_dropdown = tk.OptionMenu(janela_criar_ticket, solicitante_var, "TI", "RH", "Recepção 1", "Recepção 2", "Administrativo")
    solicitante_dropdown.grid(row=3, column=1, padx=10, pady=5)

    # Botão Criar
    botao_criar = tk.Button(janela_criar_ticket, text="Criar Ticket", command=lambda: criar_ticket(assunto_entry.get(), descricao_entry.get(), prioridade_var.get(), solicitante_var.get()))
    botao_criar.grid(row=4, columnspan=2, pady=10)


def abrir_janela_edicao():
    """Função para abrir a janela de edição."""
    janela_edicao = tk.Toplevel(janela_principal)
    janela_edicao.title("Editar Ticket de Helpdesk")

    # Label e Entry para o ID
    id_label = tk.Label(janela_edicao, text="ID do Ticket:")
    id_label.grid(row=0, column=0, padx=10, pady=5)
    id_entry = tk.Entry(janela_edicao)
    id_entry.grid(row=0, column=1, padx=10, pady=5)

    # Label e Entry para o Assunto
    assunto_label = tk.Label(janela_edicao, text="Assunto:")
    assunto_label.grid(row=1, column=0, padx=10, pady=5)
    assunto_entry = tk.Entry(janela_edicao)
    assunto_entry.grid(row=1, column=1, padx=10, pady=5)

    # Label e Entry para a Descrição
    descricao_label = tk.Label(janela_edicao, text="Descrição:")
    descricao_label.grid(row=2, column=0, padx=10, pady=5)
    descricao_entry = tk.Entry(janela_edicao)
    descricao_entry.grid(row=2, column=1, padx=10, pady=5)

    # Botão Pesquisar
    botao_pesquisar = tk.Button(janela_edicao, text="Pesquisar", command=lambda: pesquisar_ticket(id_entry, assunto_entry, descricao_entry))
    botao_pesquisar.grid(row=3, columnspan=2, pady=10)


# Janela de login
janela_login = tk.Tk()
janela_login.title("Login")
login_label = tk.Label(janela_login, text="Por favor, faça login")
login_label.pack(pady=10)
usuario_label = tk.Label(janela_login, text="Usuário:")
usuario_label.pack()
usuario_entry = tk.Entry(janela_login)
usuario_entry.pack()
senha_label = tk.Label(janela_login, text="Senha:")
senha_label.pack()
senha_entry = tk.Entry(janela_login, show="*")
senha_entry.pack()
botao_login = tk.Button(janela_login, text="Login", command=fazer_login)
botao_login.pack(pady=10)

janela_login.mainloop()


# In[ ]:




