from tkinter import *
from tkinter import ttk
from tkinter.ttk import Progressbar
import time
import os
import shutil
import getpass
import fileinput

# TODO : - add loading bar in splash screen
#        - 

# Window settings
window = Tk()
window.geometry("600x450+470+150")
window.overrideredirect(True)

# Frames
splashScreen = Frame(window, width=window.winfo_screenwidth(), height=window.winfo_screenheight(), bg="#249794")
splashScreen.pack()

instFrame = Frame(window, width=window.winfo_screenwidth(), height=window.winfo_screenheight())
checkFrame = Frame(window, width=window.winfo_screenwidth(), height=window.winfo_screenheight())

# Splash screen labels
l1 = Label(splashScreen, text="ABBEM", fg="white", bg="#249794")
lset1 = ("Calibri (Body)", 18, "bold")
l1.config(font=lset1)
l1.place(x=70, y=50)

l2 = Label(splashScreen, text="FILES", fg="white", bg="#249794")
lset2 = ("Calibri (Body)", 18, "bold")
l2.config(font=lset2)
l2.place(x=100, y=80)

l3 = Label(splashScreen, text="INSTALLER", fg="white", bg="#249794")
lset3 = ("Calibri (Body)", 13, "bold")
l3.config(font=lset3)
l3.place(x=70, y=110)

l5 = Label(splashScreen, text="Some blabla to explain the purpose of this program", fg="white", bg="#249794")
lset5 = ("Calibri (Body)", 13, "italic")
l5.config(font=lset5)
l5.place(x=50, y=180)


# Variables
selected =[]
choice = IntVar()
a = IntVar()
b = IntVar()
c = IntVar()
d = IntVar()
e = IntVar()
f = IntVar()
g = IntVar()
h = IntVar()
files = ["BuildingEnergySimulation=C:\\ProgramData\ABBEM\Revit_BuildingEnergyAnalysis\BuildingEnergySimulation.osw, ",
         "AB Besoins Bioclimatiques=C:\\ProgramData\ABBEM\AB Besoins Bioclimatiques.osw, ",
         "AB Distribution Besoins=C:\\ProgramData\ABBEM\AB Distribution Besoins.osw, ",
         "AB ViewData=C:\\ProgramData\ABBEM\ABViewData.osw, ",
         "AB EnergyPlus=C:\\ProgramData\ABBEM\AB EnergyPlus.osw, ",
         "AB Bem=C:\\ProgramData\ABBEM\AB Bem.osw, ",
         "AB ProfilsHoraires=C:\\ProgramData\ABBEM\AB ProfilsHoraires.osw, ",
         "AB Test=C:\\ProgramData\ABBEM\AB Test.osw, "]


# Get user name
user = getpass.getuser()

# Paths
src = os.path.join('C:\\', 'Users', getpass.getuser(), 'Downloads','ABBEM')
dest = os.path.join('C:\\', 'ProgramData', 'ABBem')
destCheck = os.path.join('C:\\', 'ProgramData')
revitPath = os.path.join('C:\\', 'Users', getpass.getuser(), 'AppData', 'Roaming', 'Autodesk', 'Revit')
# See if have to add 'Autodesk Revit 2021' if too slow


# Get files selected in filesFrame
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
    if (len(selected)>0):
        selected.pop()
        selected.pop()


# Find a file or a directory
def find(name, path, num):
    if num==2:
        for root, dirs, files in os.walk(path):
            if name in files:
                return root
    else:
        for root, dirs, files in os.walk(path):
            if name in dirs:
                return root
    return 'notFound'

abbemPath = find('ABBem', destCheck, 1)


# Close program
def closeWindow():
    getInfo(selected)
    window.destroy()


# Modify file 'Revit.ini'
def changeRevit():
    # Get user selection
    getInfo(selected)

    fileR = 'Revit.ini'
    filer = 'res.ini'

    # Check if user selected files
    if len(selected) == 0:
        # Pop up window
        popUpEr = Toplevel(bg="dimgrey")
        popUpEr.geometry("350x200+590+270")
        popUpEr.overrideredirect(True)

        # Label error
        label = Label(popUpEr, text="ERREUR", fg="red", bg="dimgrey")
        labelset = ('Calibri (Body)', 14, 'bold')
        label.config(font=labelset)
        label.place(x=150, y=30)

        label2 = Label(popUpEr, text="Veuillez sélectionner les fichiers voulus.", bg="dimgrey")
        label2set = ('Calibri (Body)', 10, 'normal')
        label2.config(font=label2set)
        label2.place(x=30, y=90)

        btn = Button(popUpEr, text="OK", width=10, bg="lightgray", activebackground="white", relief=GROOVE, cursor="hand2", command= lambda: popUpEr.destroy())
        btn.place(x=160, y=170)
    else:
        # Pop up window
        popUp = Toplevel(bg="dimgrey")
        popUp.geometry("350x200+590+270")
        popUp.overrideredirect(True)

        # Label program over
        label = Label(popUp, text="Programme terminé. \n Fichiers installés avec succès.", bg="dimgrey")
        labelset = ('Calibri (Body)', 14, 'italic')
        label.config(font=labelset)
        label.place(x=50, y=50)

        popUp.after(1000, lambda: (popUp.destroy(), window.destroy()))
    
        # Get 'Revit.ini' path
        # TODO : real 'Revit.ini' path = revitPath
        path = find(fileR, dest, 2)

        # Open files to read and write
        fileInput = open(os.path.join(path, fileR), "r", encoding="utf-16")
        fileOutput =  open(os.path.join(path, filer), "w", encoding="utf-16")

        # Get variable 'SystemsAnalysisWorkflows'
        var = "SystemsAnalysisWorkflows="
        for i in range(len(selected)):
            var+=selected[i]
        print(var)

        # Rewrite file 'Revit.ini'
        # TODO : automatize
        for line in fileInput:
            fileOutput.write(line.replace("OpenStudio=%ProgramFiles%\\NREL\OpenStudio CLI For Revit 2021","OpenStudio=C:\\ProgramData\ABBem\OStudio").replace("SystemsAnalysisWorkflows=BuildingEnergySimulation=E:\Revit_BuildingEnergyAnalysis\BuildingEnergySimulation.osw, AB Besoins Bioclimatiques=E:\ABBem2\AB Besoins Bioclimatiques.osw, AB Distribution Besoins=E:\ABBem2\AB Distribution Besoins.osw, AB ViewData=E:\ABBem2\AB ViewData.osw, AB EnergyPlus=E:\ABBem2\AB EnergyPlus.osw, AB Bem=E:\ABBem2\AB Bem.osw, AB ProfilsHoraires=E:\ABBem2\AB ProfilsHoraires.osw, AB Test=E:\ABBem2\AB Test.osw",var))

        # Close files
        fileInput.close()
        fileOutput.close()

        # Rename files
        # TODO : rename ONLY 1st old 'Revit.ini' in 'original_Revit.ini'
        os.rename(os.path.join(path, fileR), os.path.join(path, 'original_Revit.ini'))
        os.rename(os.path.join(path, filer), os.path.join(path, fileR))


# Frame choose files
def filesFrame():
    checkFrame.pack()

    # Label instructions
    label = Label(checkFrame, text="Veuillez sélectionner les fichiers à intégrer")
    labelset = ("Calibri (Body)", 14, "bold")
    label.config(font=labelset)
    label.place(x=100, y=30)

    # Checkbuttons to choose files
    ck1 = Checkbutton(checkFrame, text = "BuildingEnergySimulation.osw", variable = a, onvalue = 1, offvalue = 0, cursor="hand2")
    ck1.place(x=200, y=90)
    ck2 = Checkbutton(checkFrame, text = "AB Besoins Bioclimatiques.osw", variable = b, onvalue = 1, offvalue = 0, cursor="hand2")
    ck2.place(x=200, y=120)
    ck3 = Checkbutton(checkFrame, text = "AB Distribution Besoins.osw", variable = c, onvalue = 1, offvalue = 0, cursor="hand2")
    ck3.place(x=200, y=150)
    ck4 = Checkbutton(checkFrame, text = "AB ViewData.osw", variable = d, onvalue = 1, offvalue = 0, cursor="hand2")
    ck4.place(x=200, y=180)
    ck5 = Checkbutton(checkFrame, text = "AB EnergyPlus.osw", variable = e, onvalue = 1, offvalue = 0, cursor="hand2")
    ck5.place(x=200, y=210)
    ck6 = Checkbutton(checkFrame, text = "AB Bem.osw", variable = f, onvalue = 1, offvalue = 0, cursor="hand2")
    ck6.place(x=200, y=240)
    ck7 = Checkbutton(checkFrame, text = "AB ProfilsHoraires.osw", variable = g, onvalue = 1, offvalue = 0, cursor="hand2")
    ck7.place(x=200, y=270)
    ck8 = Checkbutton(checkFrame, text = "AB Test.osw", variable = h, onvalue = 1, offvalue = 0, cursor="hand2")
    ck8.place(x=200, y=300)

    # Buttons 'ok' and 'exit'
    ok = Button(checkFrame, text = "OK", width=10, bg="lightgray", activebackground="white", cursor="hand2", relief=GROOVE, command = lambda: window.destroy())
    ok.place(x=260, y=350)
    ex = Button(checkFrame, text = "Exit", width=10, bg="lightgray", activebackground="white", cursor="hand2", relief=GROOVE, command = changeRevit)
    ex.place(x=500, y=390)


# Window mooving files
def popUpWindow():

    # Create pop up window
    popUp = Toplevel(bg="dimgrey")
    popUp.geometry("350x200+590+270")
    popUp.overrideredirect(True)

    # Label waiting for files to moove
    label = Label(popUp, text="Mooving files...", bg="dimgrey")
    labelset = ('Calibri (Body)', 14, 'italic')
    label.config(font=labelset)
    label.place(x=50, y=50)

    docPath = os.path.join('Documents', 'ABBem')
    
    currentPath = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'ABBem')
    print("CURRENT PATH : " + os.path.join(currentPath, 'ABBEM'))

    # Create new directory 'ABBem' in 'ProgramData'
    if ( abbemPath != 'notFound'):
        shutil.rmtree(abbemPath)
    os.mkdir(dest)


    # Create new directory 'ABBem' in 'Documents'
    os.mkdir(docPath)

    # Moove files and dirs in new directory
    # TODO : change 'src' by 'currentPath'
    shutil.move(src, dest)

    # TODO : change for 'if files mooved'
    popUp.after(2000, lambda: (filesFrame(), popUp.destroy()))


# Mooving files or not
def mooveFiles():
    # Get user choice 'installation' or 'only files'
    install = choice.get()

    if install==0:
        instFrame.pack_forget()
        popUpWindow()
    else:
        if (abbemPath == 'notFound'):
            popUp = Toplevel(bg="dimgrey")
            popUp.geometry("400x250+580+230")
            popUp.overrideredirect(True)

            # Label error
            label = Label(popUp, text="ERREUR", fg="red", bg="dimgrey")
            labelset = ('Calibri (Body)', 14, 'bold')
            label.config(font=labelset)
            label.place(x=150, y=30)

            label2 = Label(popUp, text="L'installation complète n'a pas été effectuée auparavent. \n Vous ne pouvez pas uniquement ajouter les fichiers.", bg="dimgrey")
            label2set = ('Calibri (Body)', 10, 'normal')
            label2.config(font=label2set)
            label2.place(x=30, y=90)

            btn = Button(popUp, text="OK", width=10, bg="lightgray", activebackground="white", relief=GROOVE, cursor="hand2", command= lambda: popUp.destroy())
            btn.place(x=160, y=170)


# Frame choose 'installation' or 'only files'
def installFrame():
    splashScreen.pack_forget()

    instFrame.pack()

    # Label instructions
    label = Label(instFrame, text="Veuillez choisir l'action à effectuer")
    labelset = ("Calibri (Body)", 14, "bold")
    label.config(font=labelset)
    label.place(x=140, y=80)

    # Radio buttons
    r1 = Radiobutton(instFrame, text = "Installation complète", variable = choice, value = 0, cursor="hand2")
    r1.place(x=170, y=150)
    r2 = Radiobutton(instFrame, text = "Ajout des fichiers", variable = choice, value = 1, cursor="hand2")
    r2.place(x=170, y=200)

    # Label mention under 2nd radio btn
    warn = Label(instFrame, text="(uniquement si l'installation complète à déjà été effectuée)", fg="red")
    warn.place(x=190, y=220)

    # Button frame check files
    btn = Button(instFrame, text="Suivant", width=10, bg="lightgray", activebackground="white", cursor="hand2", relief=GROOVE, command=mooveFiles)
    btn.place(x=250, y=290)



# Splash screen timer
splashScreen.after(3000, installFrame)

window.mainloop()