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
    "Jorge", "Gosmira"," Almir Rodrigues"
]

checkboxes = {}

def validar_data(data_str):
    try:
        datetime.strptime(data_str, "%d/%m/%Y")
        return True
    except ValueError:
        return False

def salvar_chamada():
    data = entrada_data.get().strip()

    if not validar_data(data):
        messagebox.showerror("Erro", "Digite a data no formato DD/MM/AAAA.")
        return

    arquivo = "chamada.xlsx"
    nova_chamada = []

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

    # Coletar registros existentes
    registros_existentes = set()
    for row in ws.iter_rows(min_row=2, values_only=True):
        nome, _, data_registro = row
        registros_existentes.add((nome, data_registro))

    for nome, var in checkboxes.items():
        status = "Presente" if var.get() else "Ausente"
        if (nome, data) not in registros_existentes:
            ws.append([nome, status, data])

    wb.save(arquivo)

    # Limpa as seleções após salvar
    for var in checkboxes.values():
        var.set(False)

    messagebox.showinfo("Sucesso", "Chamada registrada com sucesso!")

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

# Campo de data
frame_data = tk.Frame(root)
frame_data.pack(pady=10)

tk.Label(frame_data, text="Data da chamada (DD/MM/AAAA):").pack(side=tk.LEFT)
entrada_data = tk.Entry(frame_data)
entrada_data.pack(side=tk.LEFT)

tk.Button(root, text="Salvar Chamada", command=salvar_chamada, bg="green", fg="white").pack(pady=10)

root.mainloop()