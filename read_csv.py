import os
import tkinter as tk
import tkinter.messagebox
from pathlib import Path
from subprocess import *
import time
from tkinter import filedialog

import pandas as pd

def place_after(gauche, droite, dx = 5, dy = 0):
    gauche_info = gauche.place_info()
    x = int(gauche_info['x']) + gauche.winfo_reqwidth() + dx
    y = int(gauche_info['y']) + dy
    droite.place(x = x, y = y)

def place_bellow(haut, bas, dx = 0, dy = 5):
    haut_info = haut.place_info()
    x = int(haut_info['x']) + dx
    y = int(haut_info['y']) + haut.winfo_reqheight() + dy
    bas.place(x = x, y = y)


def find_exe_path(pgm_name):
    
    '''
    Returns the full path`pgm_path` of the `pgm_name` executable file located in "C:\Progra Files"
    '''
    
    for dirpath, dirs, files in os.walk(r'C:\Program Files'):  
        for filename in files: 
            fname = os.path.join(dirpath,filename) 
            if fname.endswith(pgm_name): 
                pgm_path = fname
            
    return pgm_path

def launch_appli(pgm_name,file_name):

    '''
    Launch the launch_pgm.bat batch file located in ~/.bat directory
    '''
   
    bat_path = str(Path.home()) / Path(r".bat\launch_pgm.bat")
    pgm_path = find_exe_path(pgm_name)
    p = Popen([bat_path, pgm_path,file_name], stdout=PIPE, stderr=PIPE)
    output, errors = p.communicate()
    p.wait() # wait for process to terminate

def browseFiles():
    filename = filedialog.askopenfilename(initialdir = Path.home() / Path('download'),
                                          title = "Select a File",
                                          filetypes = (("Fichiers .csv",
                                                        "*.csv"),
                                                       ("all files",
                                                        "*.*")))
    variable_file.set(filename)
    df = pd.read_csv(variable_file.get())
    
    output_file = str(os.path.basename(variable_file.get())).split('.')[0]+".xlsx"

    output_path = Path.home() / Path("AppData\Local\Temp") / Path(output_file)
    df.to_excel(output_path)
    launch_appli('EXCEL.EXE',
                 output_path)
    output_file_save = '~$'+str(os.path.basename(variable_file.get())).split('.')[0]+".xlsx"
    output_path_save = Path.home() / Path("AppData\Local\Temp") / Path(output_file_save)
    time.sleep(5)
    
    while True:
        time.sleep(0.5)
        if not os.path.isfile(output_path_save):
            break
    os.remove(output_path)
    print(output_path_save," : finished")
    
    
root = tk.Tk()
root.geometry("350x200")
root.title("csv2excel")
variable_file = tk.StringVar(root)
label_file_explorer = tk.Label(root,
                              text = "Choisisseez votre fichier",)
label_file_explorer.place(x=50,y=50)
button_explore = tk.Button(root,
                           text = "Browse Files",
                           command = browseFiles)
place_after(label_file_explorer,button_explore,dy=0)

root.mainloop()