import fileinput
import os
import shutil
from zipfile import ZipFile
from tkinter import *
import subprocess
import time
import getpass

# interface checkbox
window = Tk()
window.title("Choix des fichiers")
window.minsize(300,300)
services = []
files = ["BuildingEnergySimulation=C:\\ProgramData\ABBEM\Revit_BuildingEnergyAnalysis\BuildingEnergySimulation.osw, ",
         "AB Besoins Bioclimatiques=C:\\ProgramData\ABBEM\AB Besoins Bioclimatiques.osw, ",
         "AB Distribution Besoins=C:\\ProgramData\ABBEM\AB Distribution Besoins.osw, ",
         "AB ViewData=C:\\ProgramData\ABBEM\ABViewData.osw, ",
         "AB EnergyPlus=C:\\ProgramData\ABBEM\AB EnergyPlus.osw, ",
         "AB Bem=C:\\ProgramData\ABBEM\AB Bem.osw, ",
         "AB ProfilsHoraires=C:\\ProgramData\ABBEM\AB ProfilsHoraires.osw, ",
         "AB Test=C:\\ProgramData\ABBEM\AB Test.osw, "]

selected =[]
a = IntVar()
b = IntVar()
c = IntVar()
d = IntVar()
e = IntVar()
f = IntVar()
g = IntVar()
h = IntVar()

def getInfo(selected):
    if (a.get() == 1):
        selected += files[0]
    if (b.get() == 1):
        selected += files[1]
    if (c.get() == 1):
        selected += files[2]
    if (d.get() == 1):
        selected += files[3]
    if (e.get() == 1):
        selected += files[4]
    if (f.get() == 1):
        selected += files[5]
    if (g.get() == 1):
        selected += files[6]
    if (h.get() == 1):
        selected += files[7]
    selected.pop()
    selected.pop()
    print(selected)

def closeWindow():
    getInfo(selected)
    window.destroy()

def checkBoxSet():
    Checkbutton(window, text = "BuildingEnergySimulation.osw", variable = a, onvalue = 1, offvalue = 0).pack()
    Checkbutton(window, text = "AB Besoins Bioclimatiques.osw", variable = b, onvalue = 1, offvalue = 0).pack()
    Checkbutton(window, text = "AB Distribution Besoins.osw", variable = c, onvalue = 1, offvalue = 0).pack()
    Checkbutton(window, text = "AB ViewData.osw", variable = d, onvalue = 1, offvalue = 0).pack()
    Checkbutton(window, text = "AB EnergyPlus.osw", variable = e, onvalue = 1, offvalue = 0).pack()
    Checkbutton(window, text = "AB Bem.osw", variable = f, onvalue = 1, offvalue = 0).pack()
    Checkbutton(window, text = "AB ProfilsHoraires.osw", variable = g, onvalue = 1, offvalue = 0).pack()
    Checkbutton(window, text = "AB Test.osw", variable = h, onvalue = 1, offvalue = 0).pack()

    Button(window, text = "OK", command = closeWindow).pack()

    window.mainloop()


# find a file
def find(name, path):
    for root, dirs, files in os.walk(path):
        if name in files:
            return root


# ######## main ########
# move dirs
src = os.path.join('C:\\', 'Users',getpass.getuser(),'Downloads','DirTest')
dest = os.path.join('C:\\', 'ProgramData', 'ABBEM') # To complete when have full path to avoid time cost

os.mkdir(dest)

shutil.move(src, dest)

# read and write file
path = find('Revit.ini', dest)
fileInput = open(os.path.join(path,'Revit.ini'), "r", encoding="utf-16")
fileOutput =  open(os.path.join(path,'res.ini'), "w", encoding="utf-16")

checkBoxSet()

var = "SystemsAnalysisWorkflows="
for i in range(len(selected)):
    var+=selected[i]
print(var)

for line in fileInput:
    fileOutput.write(line.replace("OpenStudio=%ProgramFiles%\\NREL\OpenStudio CLI For Revit 2021","OpenStudio=C:\\ProgramData\ABBem\OStudio").replace("SystemsAnalysisWorkflows=BuildingEnergySimulation=E:\Revit_BuildingEnergyAnalysis\BuildingEnergySimulation.osw, AB Besoins Bioclimatiques=E:\ABBem2\AB Besoins Bioclimatiques.osw, AB Distribution Besoins=E:\ABBem2\AB Distribution Besoins.osw, AB ViewData=E:\ABBem2\AB ViewData.osw, AB EnergyPlus=E:\ABBem2\AB EnergyPlus.osw, AB Bem=E:\ABBem2\AB Bem.osw, AB ProfilsHoraires=E:\ABBem2\AB ProfilsHoraires.osw, AB Test=E:\ABBem2\AB Test.osw",var))
fileInput.close()
fileOutput.close()
os.remove(os.path.join(path,'Revit.ini'))
os.rename(os.path.join(path,'res.ini'), os.path.join(path,'Revit.ini'))
