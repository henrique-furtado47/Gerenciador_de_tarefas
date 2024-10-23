import tkinter as tk
from tkinter import messagebox
import xlsxwriter

# Função para adicionar tarefa à lista
def adicionar_tarefa():
    tarefa = entrada_tarefa.get()
    if tarefa:
        lista_tarefas.insert(tk.END, tarefa)
        entrada_tarefa.delete(0, tk.END)
    else:
        messagebox.showwarning("Aviso", "A tarefa não pode estar vazia!")

# Função para salvar tarefas em um arquivo Excel
def salvar_tarefas():
    tarefas = lista_tarefas.get(0, tk.END)
    if tarefas:
        workbook = xlsxwriter.Workbook('tarefas.xlsx')
        worksheet = workbook.add_worksheet()
        
        worksheet.write(0, 0, 'Tarefas')
        for i, tarefa in enumerate(tarefas, 1):
            worksheet.write(i, 0, tarefa)
        
        workbook.close()
        messagebox.showinfo("Sucesso", "Tarefas salvas com sucesso!")
    else:
        messagebox.showwarning("Aviso", "Não há tarefas para salvar!")
def remover_tarefas():
    tarefas = lista_tarefas.get(0, tk.END)
    if tarefas:
        workbook = xlsxwriter.Workbook('tarefas.xlsx')
        worksheet = workbook.add_worksheet()
        
        worksheet.write(0, 0, 'Tarefas')
        for i, tarefa in enumerate(tarefas, 1):
            worksheet.write(i, 0, tarefa)
            
        workbook.close()
        messagebox.showinfo("Sucesso", "Tarefas salvas com sucesso!")
    else:
        messagebox.showwarning("Aviso", "Não há tarefas para salvar!")

# Configuração da interface gráfica
janela = tk.Tk()
janela.title("Gerenciador de Tarefas")

frame = tk.Frame(janela)
frame.pack(pady=10)

entrada_tarefa = tk.Entry(frame, width=40)
entrada_tarefa.pack(side=tk.LEFT, padx=10)

botao_adicionar = tk.Button(frame, text="Adicionar Tarefa", command=adicionar_tarefa)
botao_adicionar.pack(side=tk.LEFT)
botao_remover = tk.Button(frame, text="Remover Tarefa", command=remover_tarefas)
botao_remover.pack(side=tk.LEFT)

lista_tarefas = tk.Listbox(janela, width=50, height=10)
lista_tarefas.pack(pady=10)

botao_salvar = tk.Button(janela, text="Salvar Tarefas", command=salvar_tarefas)
botao_salvar.pack(pady=10)

# Executar o loop principal do Tkinter
janela.mainloop()