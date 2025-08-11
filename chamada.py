import tkinter as tk
from tkinter import messagebox
from openpyxl import Workbook, load_workbook
from openpyxl.utils.exceptions import InvalidFileException
from datetime import datetime
import os

alunos = [
    "OsValdina Francisca", "Paulo Henrique", "João Vitor", "Elza Alves",
    "Antonio Patricio", "Gesmindo Boostel", "Kalahan Boostel", "Geciel Polegario",
    "Diana", "Vanuza Nascimento", "Welington Nascimento", "Welington Ribeiro",
    "Jorge", "Gosmira", "Almir Rodrigues"
]

# normaliza nomes (remove espaços no início/fim)
alunos = [a.strip() for a in alunos]

checkboxes = {}

def validar_data(data_str):
    try:
        return datetime.strptime(data_str, "%d/%m/%Y")
    except ValueError:
        return None

def salvar_chamada():
    data_str = entrada_data.get().strip()
    data_dt = validar_data(data_str)
    if not data_dt:
        messagebox.showerror("Erro", "Digite a data no formato DD/MM/AAAA.")
        return

    data_fmt = data_dt.strftime("%d/%m/%Y")  # formato padronizado para comparar/salvar
    arquivo = "chamada.xlsx"

    # Abre ou cria o arquivo
    if os.path.exists(arquivo):
        try:
            wb = load_workbook(arquivo)
            ws = wb.active
        except InvalidFileException:
            messagebox.showerror("Erro", "O arquivo chamada.xlsx está corrompido ou inválido.")
            return
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = "Chamada"
        ws.append(["Nome", "Presença", "Data"])

    # coleta registros existentes (normalizando nome e data lida)
    registros_existentes = set()
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or row[0] is None:
            continue
        nome_row = str(row[0]).strip()
        # pega a coluna data (index 2) se existir
        data_registro = row[2] if len(row) > 2 else None
        if isinstance(data_registro, datetime):
            data_row = data_registro.strftime("%d/%m/%Y")
        else:
            data_row = str(data_registro).strip() if data_registro is not None else ""
        registros_existentes.add((nome_row, data_row))

    novos = 0
    for nome, var in checkboxes.items():
        status = "Presente" if var.get() else "Ausente"
        if (nome, data_fmt) not in registros_existentes:
            ws.append([nome, status, data_fmt])
            novos += 1

    wb.save(arquivo)

    # limpa seleções
    for var in checkboxes.values():
        var.set(False)

    messagebox.showinfo("Sucesso", f"Chamada registrada com sucesso! ({novos} novos registros adicionados)")

# Interface
root = tk.Tk()
root.title("Chamada Escolar")

tk.Label(root, text="Marque os alunos presentes:", font=("Arial", 14)).pack(pady=10)

frame = tk.Frame(root)
frame.pack()

for nome in alunos:
    var = tk.BooleanVar()
    chk = tk.Checkbutton(frame, text=nome, variable=var)
    chk.pack(anchor='w')
    checkboxes[nome] = var

frame_data = tk.Frame(root)
frame_data.pack(pady=10)

tk.Label(frame_data, text="Data da chamada (DD/MM/AAAA):").pack(side=tk.LEFT)
entrada_data = tk.Entry(frame_data)
entrada_data.pack(side=tk.LEFT)

tk.Button(root, text="Salvar Chamada", command=salvar_chamada, bg="green", fg="white").pack(pady=10)

root.mainloop()
