import tkinter as tk

# Variable global para almacenar la opción seleccionada
opcion_seleccionada = None

def AbrirMain():
    global opcion_seleccionada
    root = tk.Tk()
    root.title("Formulario Principal")
    root.geometry("200x100")
    root.configure(bg="darkgray")

    def obtener_valor(opcion):
        global opcion_seleccionada
        opcion_seleccionada = opcion  # Actualizar la variable global con la opción seleccionada
        root.destroy()  # Cerrar la ventana principal cuando se seleccione una opción

    frame = tk.Frame(root)
    frame.pack(pady=10)

    boton_formulario_1 = tk.Button(frame, text="Formulas unitarias", bg="white", fg="black", command=lambda: obtener_valor("Formulas unitarias"))
    boton_formulario_1.pack(expand=True)

    frame_formulario_2 = tk.Frame(root)
    frame_formulario_2.pack(pady=10)

    boton_formulario_2 = tk.Button(frame_formulario_2, text="Formulas Masivas", bg="white", fg="black", command=lambda: obtener_valor("Formulas Masivas"))
    boton_formulario_2.pack(expand=True)

    root.mainloop()

