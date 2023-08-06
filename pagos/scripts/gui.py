#!/usr/bin/env python
from tkinter import *
from tkinter import filedialog
from tkinter.scrolledtext import *
import subprocess, os

SCRIPTDIR=os.path.dirname(os.path.realpath(__file__))
DIRNAME = ""
 
def browseDirs():
    global DIRNAME
    DIRNAME = filedialog.askdirectory()
    label_top.configure(text="Directorio: "+DIRNAME)
    script_out.insert("end", "Directorio seleccionado: " + DIRNAME+ "\n") 
    script_out.insert("end", "\nAhora puede ejecutar 'Procesar bancos' \n") 
    
def run_script(script):
    global DIRNAME
    if not DIRNAME:
        script_out.insert("end", "El directorio no sido seleccionado. No se ejecutó el proceso.\n", "warning")
        return 1
    script_dir = os.path.join(SCRIPTDIR, script)
    script_out.insert("end", "Ejecutando:" + script_dir + ":\n")
    try:
        cmd_out = subprocess.check_output(["python", script_dir, DIRNAME])
        script_out.insert("end", cmd_out, "console")
        script_out.yview(END)
    except subprocess.CalledProcessError as e:
        script_out.insert("end", str(e), "error")
        script_out.yview(END)
        return 1
    return 0
    
def pbanks():
    if run_script("procesar_bancos.py") == 0:
        script_out.insert("end", "\nRevise el excel generado y luego puede ejecutar 'Procesar facturas' \n") 
    
def pbills():
    if run_script("procesar_facturas.py") == 0:
        script_out.insert("end", "\nRevise el excel generado y luego puede importar el archivo csv a QuickBooks \n")
        script_out.insert("end", "Si todo esta correcto, ya puede salir del programa.\n")

gui = Tk()
gui.title('asoarcosQB')
#window.geometry("500x500")

label_top = Label(gui, text = "Seleccione el directorio de trabajo", fg="green")
button_explore = Button(gui, text = "Buscar directorio", command = browseDirs)
button_exit = Button(gui, text = "Salir", command = exit)
button_banks = Button(gui, text = "Procesar Bancos", command = pbanks)
button_bills = Button(gui, text = "Procesar Facturas", command = pbills)
script_out = ScrolledText(gui,  height=30, width=100) 
script_out.tag_config('console', foreground="blue")
script_out.tag_config('error', foreground="red")
script_out.tag_config('warning', background="yellow")

script_out.insert("end", "Esperando a que el directorio sea seleccionado\n")

#grid
label_top.grid(column = 1, columnspan = 3, row = 1)
button_explore.grid(column = 1,  columnspan = 3, row = 2)
script_out.grid(column = 1, columnspan = 3, row = 3)
button_banks.grid(column = 1,row = 4)
button_bills.grid(column = 2,row = 4)
button_exit.grid(column = 3,row = 4)
  
gui.mainloop()