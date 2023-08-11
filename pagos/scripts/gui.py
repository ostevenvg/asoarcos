#!/usr/bin/env python
from tkinter import *
from tkinter import filedialog
from tkinter.scrolledtext import *
import subprocess, os, sys

SCRIPTDIR=os.path.dirname(os.path.realpath(__file__))
DIRNAME = ""
 
def browseDirs():
    global DIRNAME
    DIRNAME = filedialog.askdirectory()
    label_top.configure(text="Directorio: "+DIRNAME)
    script_out.insert("end", "Directorio seleccionado: " + DIRNAME+ "\n") 
    script_out.insert("end", "\nAhora puede ejecutar 'Procesar bancos',  'Procesar facturas' o 'Generar reporte'\n") 
    
def run_script(script, option=''):
    global DIRNAME
    if not DIRNAME:
        script_out.insert("end", "El directorio no sido seleccionado. No se ejecutó el proceso.\n", "warning")
        return 1

    script_dir = os.path.join(SCRIPTDIR, script)
    cmd = ["python", script_dir, DIRNAME]
    if option:
        cmd += [option]
    cmd = ' '.join(cmd)
    script_out.insert("end", "\nEjecutando:" + cmd + ":\n")
    
    p = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, shell=True)
    (out, err) = p.communicate()
    p_status = p.wait()
    script_out.insert("end", out, "console")
    script_out.yview(END)
    if p_status:
        script_out.insert("end", err.decode() + '\n', "error")
        script_out.insert("end", "El proceso falló, revise los errores.\n", "warning")
        script_out.yview(END)
        return 1    
    return 0
    
def pbanks():
    if run_script("procesar_bancos.py") == 0:
        script_out.insert("end", "\nRevise el excel generado y luego puede ejecutar 'Procesar facturas' para generar el archivo que se importa a Quickbooks.\n") 
    
def pbills():
    if run_script("procesar_facturas.py", "gen_qb_csv") == 0:
        script_out.insert("end", "\nRevise el excel generados, si 'Procesar bancos' había sido ejecutado también se creó el archivo CSV que se importa en Quickbooks. \n")

def grpt():
    if run_script("procesar_facturas.py", "gen_report") == 0:
        script_out.insert("end", "\nRevise el reporte HTML generado.\n")


gui = Tk()
gui.title('asoarcosQB')

label_top = Label(gui, text = "Seleccione el directorio de trabajo")
button_explore = Button(gui, text = "Elegir directorio", command = browseDirs)
button_exit = Button(gui, text = "Salir", command = exit)
button_banks = Button(gui, text = "Procesar Bancos", command = pbanks)
button_bills = Button(gui, text = "Procesar Facturas", command = pbills)
button_rpt = Button(gui, text = "Crear Reporte", command = grpt)
script_out = ScrolledText(gui,  height=30, width=100) 
script_out.tag_config('console', foreground="blue")
script_out.tag_config('error', foreground="red")
script_out.tag_config('warning', background="yellow")

script_out.insert("end", "Esperando a que el directorio sea seleccionado\n")

#grid
label_top.grid(column = 1, columnspan = 4, row = 1)
button_explore.grid(column = 1,  columnspan = 4, row = 2)
script_out.grid(column = 1, columnspan = 4, row = 3)
button_banks.grid(column = 1,row = 4)
button_bills.grid(column = 2,row = 4)
button_rpt.grid(column = 3,row = 4)
button_exit.grid(column = 4,row = 4)
  
gui.mainloop()