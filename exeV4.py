import tkinter
import tkinter.messagebox
from tkinter import *
from tkinter import ttk
from tkinter.ttk import Progressbar
import time
import os
import stat
import sys
import shutil
import getpass
import fileinput

# TODO : 
#        + handle error message when doesn't find ABBem in commplete installation 
#        + add button get original_Revit.ini back
#        + create raccourci of exe and put in Desktop
#        + add help button



import os, winshell, win32com.client
import pathlib
desktop = winshell.desktop()
#desktop = r"path to where you wanna put your .lnk file"
path2 = os.path.join(desktop, 'exeV4 - raccourci.lnk')
import pathlib

target = os.path.join(pathlib.Path().resolve(),'exeV4.exe')
print(pathlib.Path().resolve())
shell = win32com.client.Dispatch("WScript.Shell")
shortcut = shell.CreateShortCut(path2)
shortcut.Targetpath = target
shortcut.save()

# Window settings
window = Tk()
window.geometry("600x450+470+150")

# Frames
splashScreen = Frame(window, width=window.winfo_screenwidth(), height=window.winfo_screenheight(), bg="#249794")
splashScreen.pack()

instFrame = Frame(window, width=window.winfo_screenwidth(), height=window.winfo_screenheight())
checkFrame = Frame(window, width=window.winfo_screenwidth(), height=window.winfo_screenheight())


# Variables
fileR = 'Revit.ini'
filer = 'res.ini'
selected =[]
choice = IntVar()

# Get user name
user = getpass.getuser()

# Paths
srcABBEM = os.path.join(pathlib.Path().resolve(),'ABBEM')
dest = os.path.join('C:\\', 'ProgramData', 'ABBem')
destCheck = os.path.join('C:\\', 'ProgramData')
revitPath = os.path.join('C:\\', 'Users', user, 'AppData', 'Roaming', 'Autodesk', 'Revit')
pathRtmp = os.path.join('C:\\', 'Users', user, 'Documents', 'Work')

if(str(os.path.exists(srcABBEM))=='False'):
   tkinter.messagebox.showinfo(title='Dossier manquant', message="Dossier ABBem inexistant, \nABBem doit être dans le même répertoire que l'exécutable.")
   window.destroy()
   sys.exit()


# Splash screen labels

# Texts labels
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

l5 = Label(splashScreen, text="explications", fg="white", bg="#249794")
lset5 = ("Calibri (Body)", 13, "italic")
l5.config(font=lset5)
l5.place(x=50, y=180)


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

# Read Revit.ini and copy lines
def getFiles():
    if (find('original_Revit.ini', revitPath, 2) != 'notFound'):
        fileR = "original_Revit.ini"
    else:
        fileR = 'Revit.ini'
    
    path = find(fileR, revitPath, 2)

    fileInput = open(os.path.join(path, fileR), "r", encoding="utf-16")

    for line in fileInput:
            if (line.startswith("SystemsAnalysisWorkflows")):
                files = line
    fileInput.close()
    files = files.replace("SystemsAnalysisWorkflows=", '', 1)
    files = files.replace("\n", '')
    files = files.split(', ')
    return files

def getLines():
    path = find(fileR, revitPath, 2)

    sysLin = ""

    fileInput = open(os.path.join(path, fileR), "r", encoding="utf-16")

    for line in fileInput:
        if (line.startswith("SystemsAnalysisWorkflows")):
                sysLin = line
                sysLin = sysLin.replace("\n", "")
        if (line.startswith("OpenStudio")):
            openLine = line
            openLine = openLine.replace("\n", "")
    return sysLin

def getLinesO():
    path = find(fileR, revitPath, 2)

    openLi = ""

    fileInput = open(os.path.join(path, fileR), "r", encoding="utf-16")
    for line in fileInput:
        if (line.startswith("OpenStudio")):
            openLi = line
            openLi = openLi.replace("\n", "")
    return openLi

sysLine = getLines()
openLine = getLinesO()

files = getFiles()

# Get files names
fileName = files.copy()
ckVar = {}
for i in range (len(fileName)):
    elmt = fileName[i]
    elmt = elmt.rsplit('\\', 1)
    fileName[i] = elmt[1]    
    ckVar[i] = tkinter.IntVar()

# Get files selected in filesFrame
def getInfo(selected):
    for i in range(len(fileName)):
        if (ckVar[i].get() == 1):
            selected+=(files[i]+', ')

if (find('ABBem', destCheck, 1)!='notFound'):
    abbemPath = os.path.join(find('ABBem', destCheck, 1), 'ABBem')
else:
    abbemPath = -1
    


# Close program
def closeWindow():
    getInfo(selected)
    window.destroy()


# Modify file 'Revit.ini' without closing window
def revitNoClose():
    getInfo(selected)

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
        label.place(x=130, y=30)

        label2 = Label(popUpEr, text="Veuillez sélectionner les fichiers voulus.", bg="dimgrey")
        label2set = ('Calibri (Body)', 12, 'italic')
        label2.config(font=label2set)
        label2.place(x=35, y=80)

        btn = Button(popUpEr, text="OK", width=10, bg="lightgray", activebackground="white", relief=GROOVE, cursor="hand2", command= lambda: popUpEr.destroy())
        btn.place(x=140, y=150)
    else:
        selected.pop()
        selected.pop()
        # Pop up window
        popUp = Toplevel(bg="dimgrey")
        popUp.geometry("350x200+590+270")
        popUp.overrideredirect(True)

        # Label program over
        label = Label(popUp, text="Programme terminé. \n Fichiers installés avec succès.", bg="dimgrey")
        labelset = ('Calibri (Body)', 12, 'italic')
        label.config(font=labelset)
        label.place(x=40, y=70)

        popUp.after(1300, popUp.destroy())
    
        # Get 'Revit.ini' path
        # TODO : real 'Revit.ini' path = revitPath
        path = find(fileR, revitPath, 2)

        # Open files to read and write
        fileInput = open(os.path.join(path, fileR), "r", encoding="utf-16")
        fileOutput =  open(os.path.join(path, filer), "w", encoding="utf-16")

        # Get variable 'SystemsAnalysisWorkflows'
        var = "SystemsAnalysisWorkflows="
        for i in range(len(selected)):
            var+=selected[i]

        # Rewrite file 'Revit.ini'
        for line in fileInput:
            fileOutput.write(line.replace(openLine,"OpenStudio=C:\\ProgramData\ABBem\OStudio").replace(sysLine, var))

        # Close files
        fileInput.close()
        fileOutput.close()

        # Rename files
        if (find('original_Revit.ini', revitPath, 2) == 'notFound'):
            os.rename(os.path.join(path, fileR), os.path.join(path, 'original_Revit.ini'))
        else:
            os.remove(os.path.join(path, fileR))
        os.rename(os.path.join(path, filer), os.path.join(path, fileR))
    tkinter.messagebox.showinfo(title='ajout fichiers', message='les fichiers ont été ajoutés avec succès')

def reinitRevit():
    check = revitPath+'\original_Revit.ini'
    toDelete = revitPath+'\Revit.ini'
    if((str(os.path.exists(check)))!="True"):
        tkinter.messagebox.showinfo(title='réinitialisation', message='vos fichiers sont déjà réinitialisés')
    else:
        os.remove(toDelete)
        os.rename(os.path.join(revitPath, 'original_Revit.ini'), os.path.join(revitPath, 'Revit.ini'))
        tkinter.messagebox.showinfo(title='réinitialisation', message='la réinitialisation a été effectuée ')

# Modify file 'Revit.ini'
def changeRevit():
    # Get user selection
    getInfo(selected)

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
        label.place(x=130, y=30)

        label2 = Label(popUpEr, text="Veuillez sélectionner les fichiers voulus.", bg="dimgrey")
        label2set = ('Calibri (Body)', 12, 'italic')
        label2.config(font=label2set)
        label2.place(x=35, y=80)

        btn = Button(popUpEr, text="OK", width=10, bg="lightgray", activebackground="white", relief=GROOVE, cursor="hand2", command= lambda: popUpEr.destroy())
        btn.place(x=140, y=150)
    else:
        selected.pop()
        selected.pop()
        # Pop up window
        popUp = Toplevel(bg="dimgrey")
        popUp.geometry("350x200+590+270")
        popUp.overrideredirect(True)

        # Label program over
        label = Label(popUp, text="Programme terminé. \n Fichiers installés avec succès.", bg="dimgrey")
        labelset = ('Calibri (Body)', 14, 'italic')
        label.config(font=labelset)
        label.place(x=40, y=70)

        popUp.after(1300, lambda: (popUp.destroy(), window.destroy()))
    
        # Get 'Revit.ini' path
        # TODO : real 'Revit.ini' path = revitPath
        path = find(fileR, revitPath, 2)

        # Open files to read and write
        fileInput = open(os.path.join(path, fileR), "r", encoding="utf-16")
        fileOutput = open(os.path.join(path, filer), "w", encoding="utf-16")

        # Get variable 'SystemsAnalysisWorkflows'
        var = "SystemsAnalysisWorkflows="
        for i in range(len(selected)):
            var+=selected[i]

        # Rewrite file 'Revit.ini'
        for line in fileInput:
            fileOutput.write(line.replace(openLine,"OpenStudio=C:\\ProgramData\ABBem\OStudio").replace(sysLine, var))

        # Close files
        fileInput.close()
        fileOutput.close()

        # Rename files
        if (find('original_Revit.ini', revitPath, 2) == 'notFound'):
            os.rename(os.path.join(path, fileR), os.path.join(path, 'original_Revit.ini'))
        else:
            os.remove(os.path.join(path, fileR))
        os.rename(os.path.join(path, filer), os.path.join(path, fileR))

def helpBtn():

    # Pop up window
    popUpEr = Toplevel(bg="lightgrey", borderwidth="2")
    popUpEr.geometry("560x420+500+180")
    popUpEr.overrideredirect(True)
    cpt = 10
    for i in range(len(fileName)):
        label=Label(popUpEr, text="- "+fileName[i], bg="lightgrey")
        labelset = ('Calibri (Body)', 10, 'bold')
        label.config(font=labelset)
        label.place(x=25, y=cpt)
        cpt+=45
   



    #explication de chaque element

    labelExplain=[]
    labelExplain.append("explication 1 ...")
    labelExplain.append("explication 2 ...")
    labelExplain.append("explication 3 ...")
    labelExplain.append("explication 4 ...")
    labelExplain.append("explication 5 ...")
    labelExplain.append("explication 6 ...")
    labelExplain.append("explication 7 ...")
    labelExplain.append("explication 8 ...")

    cpt=30
    for i in range(len(fileName)):
        label=Label(popUpEr, text=labelExplain[i], bg="lightgrey")
        labelset = ('Calibri (Body)', 10)
        label.config(font=labelset)
        label.place(x=25, y=cpt)
        cpt+=45




    btn = Button(popUpEr, text="OK", width=10, bg="lightgray", activebackground="white", relief=GROOVE, cursor="hand2", command= lambda: popUpEr.destroy())
    btn.place(x=235, y=390)


# Frame choose files
def filesFrame():
    checkFrame.pack()

    # Label instructions
    label = Label(checkFrame, text="Veuillez sélectionner les fichiers à intégrer")
    labelset = ("Calibri (Body)", 14, "bold")
    label.config(font=labelset)
    label.place(x=100, y=30)
    yPlace = 60

    # Checkbuttons to choose files
    for i in range (len(fileName)):
        ck = Checkbutton(checkFrame, text=fileName[i], variable=ckVar[i], onvalue=1, offvalue=0, cursor="hand2")
        yPlace+=30
        ck.place(x=200, y=yPlace)

    # Buttons 'ok' and 'exit'
    ex = Button(checkFrame, text = "OK et Quitter", width=10, bg="lightgray", activebackground="white", cursor="hand2", relief=GROOVE, command = changeRevit)
    ex.place(x=240, y=350)
    reinit = Button(checkFrame, text = "Réinitialiser Revit", width=15, bg="lightgray", activebackground="white", cursor="hand2", relief=GROOVE, command = reinitRevit)
    reinit.place(x=25, y=390)
    help = Button(checkFrame, text = "Aide", width=10, bg="lightgray", activebackground="white", cursor="hand2", relief=GROOVE, command = helpBtn)
    help.place(x=25, y=70)

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
    label.place(x=100, y=80)

    docPath = os.path.join('C:\\','Users',user,'Documents','ABBem')
    
    currentPath = pathlib.Path().resolve()

    # Create new directory 'ABBem' in 'ProgramData'
    if (abbemPath != -1):
        shutil.rmtree(abbemPath)
    os.mkdir(dest)

    # Create new directory 'ABBem' in 'Documents'
    if (docPath == 'notFound'):
        os.mkdir(docPath)

    # Moove files and dirs in new directory
    shutil.move(currentPath, dest)

    popUp.after(2000, lambda: (filesFrame(), popUp.destroy()))


# Mooving files or not
def mooveFiles():
    # Get user choice 'installation' or 'only files'
    install = choice.get()

    if install==0:
        instFrame.pack_forget()
        popUpWindow()
    else:
        if (abbemPath == -1):
            popUp = Toplevel(bg="dimgrey")
            popUp.geometry("400x250+580+230")
            popUp.overrideredirect(True)

            # Label error
            label = Label(popUp, text="ERREUR", fg="red", bg="dimgrey")
            labelset = ('Calibri (Body)', 14, 'bold')
            label.config(font=labelset)
            label.place(x=150, y=30)

            label2 = Label(popUp, text="L'installation complète n'a pas été effectuée auparavent. \n Vous ne pouvez pas uniquement ajouter les fichiers.", bg="dimgrey")
            label2set = ('Calibri (Body)', 11, 'italic')
            label2.config(font=label2set)
            label2.place(x=20, y=90)

            btn = Button(popUp, text="OK", width=10, bg="lightgray", activebackground="white", relief=GROOVE, cursor="hand2", command= lambda: popUp.destroy())
            btn.place(x=160, y=170)
        else:
            instFrame.pack_forget()
            filesFrame()


# Frame choose 'installation' or 'only files'
def installFrame():
    splashScreen.pack_forget()

    instFrame.pack()

    # Logo label
    #logoLabel = Label(instFrame, image=logo)
    #logoLabel.place(x=5, y=5)

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

    # Button to frame check files
    btn = Button(instFrame, text="Suivant", width=10, bg="lightgray", activebackground="white", cursor="hand2", relief=GROOVE, command=mooveFiles)
    btn.place(x=250, y=290)

    # Image label
    #imgLabel = Label(instFrame, image=img)
    #imgLabel.place(x=30, y=320)


# Splash screen timer
splashScreen.after(3000, installFrame)

window.mainloop()
