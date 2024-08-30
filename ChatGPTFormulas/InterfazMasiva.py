import tkinter as tk
from tkinter import filedialog, messagebox
import os

def AbrirInterfazMasiva():
    root = tk.Tk()
    root.title("Generador de fórmulas")
    
    texto_ingresado = tk.StringVar()
    Fila1 = tk.StringVar()
    Fila2 = tk.StringVar()
    col1= tk.StringVar()
    col2= tk.StringVar()
    ColIngreso = tk.StringVar()
    HojaIngreso = tk.StringVar()
    archivo_seleccionado = tk.StringVar()
    var_checkbox = tk.IntVar()

    def aceptar():
        Fila1.set(prompt_entry2.get())
        Fila2.set(prompt_entry3.get())
        col1.set(prompt_entry4.get())
        col2.set(prompt_entry5.get())
        ColIngreso.set(entry.get())
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
            sheets.grid(row=8, column=1, padx=10, pady=5)
        else:
            sheets.grid_forget()

    label = tk.Label(root, text="Digite la fórmula que requiere")
    label.grid(row=0, column=0, padx=10, pady=5)
  
    options = ["SUMA", "RESTA", "MULTIPLICACION", "DIVISION", "PROMEDIO", "MAXIMO"]

    listbox = tk.Listbox(root, width=25)
    for option in options:
        listbox.insert(tk.END, option)
    listbox.grid(row=1, column=0, padx=10, pady=5)

    label1 = tk.Label(root, text="Digite la fila inicial")
    label1.grid(row=2, column=0, padx=10, pady=5)
    
    prompt_entry2 = tk.Entry(root, textvariable=Fila1)
    prompt_entry2.grid(row=2, column=1, padx=10, pady=5)

    label2 = tk.Label(root, text="Digite la fila final")
    label2.grid(row=3, column=0, padx=10, pady=5)

    prompt_entry3 = tk.Entry(root, textvariable=Fila2)
    prompt_entry3.grid(row=3, column=1, padx=10, pady=5)

    label2 = tk.Label(root, text="Digite la primera columna")
    label2.grid(row=4, column=0, padx=10, pady=5)
    
    prompt_entry4 = tk.Entry(root, textvariable=col1)
    prompt_entry4.grid(row=4, column=1, padx=10, pady=5)

    label2 = tk.Label(root, text="Digite la segunda columna")
    label2.grid(row=5, column=0, padx=10, pady=5)

    prompt_entry5 = tk.Entry(root, textvariable=col2)
    prompt_entry5.grid(row=5, column=1, padx=10, pady=5)

    archivo_label = tk.Label(root, text="Seleccione un archivo:")
    archivo_label.grid(row=7, column=0, padx=10, pady=5)

    archivo_entry = tk.Entry(root, textvariable=archivo_seleccionado)
    archivo_entry.grid(row=7, column=1, padx=10, pady=5)

    archivo_button = tk.Button(root, text="Seleccionar archivo", command=seleccionar_archivo)
    archivo_button.grid(row=7, column=2, padx=10, pady=5)

    checkbox = tk.Checkbutton(root, text="Contiene más de 1 una hoja", variable=var_checkbox, command=toggle_entry)
    checkbox.grid(row=8, column=0, padx=10, pady=5)

    sheets = tk.Entry(root, textvariable=HojaIngreso)

    label = tk.Label(root, text="Digite la columna donde requiere ingresar las fórmulas")
    label.grid(row=6, column=0, padx=10, pady=5)

    entry = tk.Entry(root, textvariable=ColIngreso)
    entry.grid(row=6, column=1, padx=10, pady=5)

    aceptar_button = tk.Button(root, text="Aceptar", command=aceptar)
    aceptar_button.grid(row=9, column=0, columnspan=3, padx=10, pady=5)

    root.protocol("WM_DELETE_WINDOW", cerrar_ventana)
    root.mainloop()
    return texto_ingresado.get(), Fila1.get(), Fila2.get(), archivo_seleccionado.get(), ColIngreso.get(), HojaIngreso.get(), "M", col1.get(), col2.get()