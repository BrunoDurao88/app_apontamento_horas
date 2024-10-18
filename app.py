import tkinter as tk
from tkinter import ttk, messagebox
import time
import pandas as pd
from datetime import datetime
import os
from tkinter import PhotoImage

# Global variables
start_time = None
running = False
log_data = []

# Initialize the Excel file if it doesn't exist
file_name = "registro_atividades.xlsx"
if not os.path.exists(file_name):
    df = pd.DataFrame(columns=["Nome", "Data", "Nome da Demanda", "Tipo de Demanda", 
                               "Tempo de Processo", "Etapa Crisp-DM", "Descrição"])
    df.to_excel(file_name, index=False)

# Start the timer
def start_timer():
    global start_time, running
    if not running:
        start_time = time.time()
        running = True
        update_timer()

        nome = entry_nome.get()
        nome_demanda = entry_nome_demanda.get()
        tipo_demanda = combo_tipo_demanda.get()
        # Validate required fields
        if not nome or not nome_demanda or not tipo_demanda:
            messagebox.showwarning("Campos obrigatórios", "Preencha os campos Nome, Nome da Demanda e Tipo de Demanda.")
            return

# Stop the timer and save the data to Excel
def stop_timer():
    global running
    if running:
        running = False
        end_time = time.time()
        elapsed_time = round(end_time - start_time, 2)
        
        # Get the data from the fields
        nome = entry_nome.get()
        nome_demanda = entry_nome_demanda.get()
        tipo_demanda = combo_tipo_demanda.get()
        etapa_crisp = combo_etapa_crisp.get()
        descricao = entry_descricao.get("1.0", tk.END).strip()
        
        # Validate required fields
        if not nome or not nome_demanda or not tipo_demanda:
            messagebox.showwarning("Campos obrigatórios", "Preencha os campos Nome, Nome da Demanda e Tipo de Demanda.")
            return
        
        # Save the data to Excel
        data_atual = datetime.now().strftime("%Y-%m-%d")
        df = pd.read_excel(file_name)
        new_entry = {"Nome": nome, "Data": data_atual, "Nome da Demanda": nome_demanda, 
                     "Tipo de Demanda": tipo_demanda, "Tempo de Processo": elapsed_time, 
                     "Etapa Crisp-DM": etapa_crisp, "Descrição": descricao}
        df = df.append(new_entry, ignore_index=True)
        df.to_excel(file_name, index=False)
        
        # Reset the fields
        entry_nome.delete(0, tk.END)
        entry_nome_demanda.delete(0, tk.END)
        combo_tipo_demanda.set("")
        combo_etapa_crisp.set("")
        entry_descricao.delete("1.0", tk.END)
        label_timer.config(text="00:00:00")

# Update the timer display
def update_timer():
    if running:
        elapsed_time = time.time() - start_time
        mins, secs = divmod(elapsed_time, 60)
        hours, mins = divmod(mins, 60)
        time_format = f'{int(hours):02}:{int(mins):02}:{int(secs):02}'
        label_timer.config(text=time_format)
        root.after(1000, update_timer)



# Adjustments based on the user's requests

# Start the Tkinter UI construction with the updated background color, resized fields, and logo placeholder
root = tk.Tk()
root.title("Registro de Tempo")
root.geometry("400x450")
root.config(bg="#D3D3D3")  # Light gray background

# Logo (configurable to accept an image)
logo = PhotoImage(file="Imagem\Logo_SDD_verde_Escuro3.png")  # Replace with your image path
logo_label = tk.Label(root, image=logo, bg="#D3D3D3")
logo_label.place(x=140, y=362)  # Positioning in the top right corner

# Fields
tk.Label(root, text="Nome:", bg="#D3D3D3").pack(anchor="w", padx=10)
entry_nome = tk.Entry(root, width=40)
entry_nome.pack(pady=5)

tk.Label(root, text="Nome da Demanda:", bg="#D3D3D3").pack(anchor="w", padx=10)
entry_nome_demanda = tk.Entry(root, width=40)
entry_nome_demanda.pack(pady=5)

tk.Label(root, text="Tipo de Demanda:", bg="#D3D3D3").pack(anchor="w", padx=10)
combo_tipo_demanda = ttk.Combobox(root, values=["Projeto", "Manutenção", "Acionamento Pontual", "Reunião"], width=37)
combo_tipo_demanda.pack(pady=5)

tk.Label(root, text="Etapa Crisp-DM:", bg="#D3D3D3").pack(anchor="w", padx=10)
combo_etapa_crisp = ttk.Combobox(root, values=["1.Business Understanding", "2.Data Understanding", "3.Data Preparation", 
                                               "4.Modeling", "5.Evaluation", "6.Deployment"], width=37)
combo_etapa_crisp.pack(pady=5)

tk.Label(root, text="Descrição:", bg="#D3D3D3").pack(anchor="w", padx=10)
entry_descricao = tk.Text(root, height=6, width=35, font=("Helvetica", 9))  # Reduced font size
entry_descricao.pack(pady=5)

# Timer display
label_timer = tk.Label(root, text="00:00:00", font=("Helvetica", 16), bg="#D3D3D3")
label_timer.pack(pady=5)  # Reduced space

# Start/Stop buttons
btn_start = tk.Button(root, text="Start", command=start_timer, bg="green", fg="white", width=10)
btn_start.pack(side="left", padx=40, pady=10)

btn_stop = tk.Button(root, text="Stop", command=stop_timer, bg="red", fg="white", width=10)
btn_stop.pack(side="right", padx=40, pady=10)

# Start the Tkinter loop
root.mainloop()
