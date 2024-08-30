import tkinter as tk
from tkinter import filedialog, messagebox
import os

def AbrirInterfaz():
    
    root = tk.Tk()
    root.title("Generador de fórmulas")
    
    texto_ingresado = tk.StringVar()
    Celda1 = tk.StringVar()
    Celda2 = tk.StringVar()
    CeldaIngreso = tk.StringVar()
    HojaIngreso = tk.StringVar()
    archivo_seleccionado = tk.StringVar()
    var_checkbox = tk.IntVar()
    var_checkbox2 = tk.IntVar()

    def aceptar():
        Celda1.set(prompt_entry2.get())
        Celda2.set(prompt_entry3.get())
        CeldaIngreso.set(entry.get())
        texto_ingresado.set(listbox.get(tk.ACTIVE))
        HojaIngreso.set(sheets.get())
        root.destroy() 
   
    def cerrar_ventana():
        root.destroy()

    def seleccionar_archivo():
        filename = filedialog.askopenfilename()
        if filename:
            if os.path.splitext(filename)[1].lower() == ".xlsx":
                archivo_seleccionado.set(filename)
            else:
                messagebox.showerror("Error", "Seleccione un archivo con extensión .xlsx")

    def toggle_entry():
        if var_checkbox.get() == 1: 
            sheets.grid(row=6, column=1, padx=10, pady=5)
        else:
            sheets.grid_forget()
    def hidden_entry():
        if var_checkbox2.get() == 1: 
            prompt_entry3.grid_forget()
            label2.grid_forget()
            label1.grid(row=2, column=0, padx=10, pady=5)
            labelcol.grid(row=2, column=0, padx=10, pady=5)
        else:
            prompt_entry3.grid(row=3, column=1, padx=10, pady=5)
            label2.grid(row=3, column=0, padx=10, pady=5)
            label1.grid(row=2, column=0, padx=10, pady=5)
            labelcol.grid_forget()

    label = tk.Label(root, text="Digite la fórmula que requiere")
    label.grid(row=0, column=0, padx=10, pady=5)
  
    options = ["SUMA POR RANGOS", "SUMA ENTRE DOS CELDAS", "RESTA", "MULTIPLICACION POR RANGOS","MULTIPLICACION ENTRE DOS CELDAS", "DIVISION", "PROMEDIO", "CONTAR"]

    listbox = tk.Listbox(root, width=25)
    for option in options:
        listbox.insert(tk.END, option)
    listbox.grid(row=1, column=0, padx=10, pady=5)

    label1 = tk.Label(root, text="Digite la primera celda")
    label1.grid(row=2, column=0, padx=10, pady=5)

    labelcol = tk.Label(root, text="Digite la columna a operar")
    
    prompt_entry2 = tk.Entry(root, textvariable=Celda1)
    prompt_entry2.grid(row=2, column=1, padx=10, pady=5)

    label2 = tk.Label(root, text="Digite la segunda celda")
    label2.grid(row=3, column=0, padx=10, pady=5)

    prompt_entry3 = tk.Entry(root, textvariable=Celda2)
    prompt_entry3.grid(row=3, column=1, padx=10, pady=5)

    archivo_label = tk.Label(root, text="Seleccione un archivo:")
    archivo_label.grid(row=5, column=0, padx=10, pady=5)

    archivo_entry = tk.Entry(root, textvariable=archivo_seleccionado)
    archivo_entry.grid(row=5, column=1, padx=10, pady=5)

    archivo_button = tk.Button(root, text="Seleccionar archivo", command=seleccionar_archivo)
    archivo_button.grid(row=5, column=2, padx=10, pady=5)

    checkbox = tk.Checkbutton(root, text="Contiene más de 1 una hoja", variable=var_checkbox, command=toggle_entry)
    checkbox.grid(row=6, column=0, padx=10, pady=5)

    checkbox = tk.Checkbutton(root, text="Operar toda la columna", variable=var_checkbox2, command=hidden_entry)
    checkbox.grid(row=1, column=1, padx=10, pady=5)

    sheets = tk.Entry(root, textvariable=HojaIngreso)

    label = tk.Label(root, text="Digite la celda donde requiere ingresar la fórmula")
    label.grid(row=4, column=0, padx=10, pady=5)

    entry = tk.Entry(root, textvariable=CeldaIngreso)
    entry.grid(row=4, column=1, padx=10, pady=5)

    aceptar_button = tk.Button(root, text="Aceptar", command=aceptar)
    aceptar_button.grid(row=7, column=0, columnspan=3, padx=10, pady=5)

    root.protocol("WM_DELETE_WINDOW", cerrar_ventana)
    root.mainloop()
    return texto_ingresado.get(), Celda1.get(), Celda2.get(), archivo_seleccionado.get(), CeldaIngreso.get(), HojaIngreso.get(), "U"