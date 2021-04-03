import fileinput
import os
from zipfile import ZipFile
from tkinter import *


# interface checkbox
window = Tk()
window.title("Choix des fichiers")
window.minsize(300,300)
services = []

def getInfo():
    for i in range(len(services)):
        if services[i].get() >= 1:
            selected = ""
            selected+=str(i)
            print(selected)

def closeWindow():
    getInfo()
    window.destroy()

def checkBoxSet():
    for i in range(8):
        option = IntVar()
        option.set(0)
        services.append(option)

    Checkbutton(window, text = "BuildingEnergySimulation.osw", variable = services[0]).pack()
    Checkbutton(window, text = "AB Besoins Bioclimatiques.osw", variable = services[1]).pack()
    Checkbutton(window, text = "AB Distribution Besoins.osw", variable = services[2]).pack()
    Checkbutton(window, text = "AB ViewData.osw", variable = services[3]).pack()
    Checkbutton(window, text = "AB EnergyPlus.osw", variable = services[4]).pack()
    Checkbutton(window, text = "AB Bem.osw", variable = services[5]).pack()
    Checkbutton(window, text = "AB ProfilsHoraires.osw", variable = services[6]).pack()
    Checkbutton(window, text = "AB Test.osw", variable = services[7]).pack()

    Button(window, text = "OK", command = closeWindow).pack()

    window.mainloop()


# find a file
def find(name, path):
    for root, dirs, files in os.walk(path):
        if name in files:
            return os.path.join(root,name)


# main
dirS = find("zipTest.zip","C:\\Users\AMANI\Downloads")
dirD = "C:\\ProgramData\ABBEM"

# extract zip file and move it to new directory
zf = ZipFile(dirS, 'r')
zf.extractall("C:\\ProgramData\ABBEM")
zf.close()

# read and write file
path = find('test.txt', 'C:\\ProgramData\ABBEM')
print(path)
fileInput = open(path, "r")
fileOutput =  open('res.txt', "w")

checkBoxSet()

for line in fileInput:
    fileOutput.write(line.replace("OpenStudio=%ProgramFiles%\\NREL\OpenStudio CLI For Revit 2021","OpenStudio=C:\\ProgramData\ABBem\OStudio"))
fileInput.close()
fileOutput.close()
os.remove(path)
os.rename(find('res.txt', 'C:\\ProgramData\ABBEM'), path)
