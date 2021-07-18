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
import os, winshell, win32com.client
import pathlib


desktop = winshell.desktop()
#desktop = r"path to where you wanna put your .lnk file"
path2 = os.path.join(desktop, 'exeV4 - raccourci.lnk')
target = os.path.join(pathlib.Path().resolve(),'exeV4.exe')
print("TARGET "+target)
shell = win32com.client.Dispatch("WScript.Shell")
shortcut = shell.CreateShortCut(path2)
shortcut.Targetpath = target
shortcut.save()

# Initialisation de la fenêtre
window = Tk()
window.geometry("600x450+470+150")

# Initialisation frame vert
splashScreen = Frame(window, width=window.winfo_screenwidth(), height=window.winfo_screenheight(), bg="#249794")
splashScreen.pack()

instFrame = Frame(window, width=window.winfo_screenwidth(), height=window.winfo_screenheight())
checkFrame = Frame(window, width=window.winfo_screenwidth(), height=window.winfo_screenheight())


# Variables
fileR = 'Revit.ini'
filer = 'res.ini'
selected =[]
choice = IntVar()

# Nom d'utilisateur du PC
user = getpass.getuser()

# Chemins des fichiers
srcABBEM = os.path.join(pathlib.Path().resolve(),'ABBEM')
dest = os.path.join('C:\\', 'ProgramData', 'ABBem')
destCheck = os.path.join('C:\\', 'ProgramData')
revitPath = os.path.join('C:\\', 'Users', user, 'AppData', 'Roaming', 'Autodesk', 'Revit')
pathRtmp = os.path.join('C:\\', 'Users', user, 'Documents', 'Work')

# Si ABBem n'existe pas à côté de l'exécutable au moment de l'exécution
if(str(os.path.exists(srcABBEM))=='False'):
   tkinter.messagebox.showinfo(title='Dossier manquant', message="Dossier ABBem inexistant, \nABBem doit être dans le même répertoire que l'exécutable.")
   window.destroy()
   sys.exit()


# Textes du frame vert
# Pour modifier : changer ce qu'il y a dans 'text='
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
l3.place(x=120, y=110)

l5 = Label(splashScreen, text="explications", fg="white", bg="#249794")
lset5 = ("Calibri (Body)", 13, "italic")
l5.config(font=lset5)
l5.place(x=50, y=180)


# Fonction qui renvoie le chemin d'un fichier (si on met num=2) ou
# d'un répertoire (si on met num=1)
# elle renvoie 'notFound' si le fichier ou le répertoire n'existe pas
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


# Fonction qui lit Revit.ini et récupère les noms des fichiers de la variable SystemsAnalysisWorkflows
def getFiles():
    if (find('original_Revit.ini', revitPath, 2) != 'notFound'):
        fileR = "original_Revit.ini"
    else:
        fileR = 'Revit.ini'
    
    path = find(fileR, revitPath, 2)
    
    # Lit Revit.ini
    fileInput = open(os.path.join(path, fileR), "r", encoding="utf-16")
    
    # Récupère tout SystemsAnalysisWorkflows dans line
    for line in fileInput:
            if (line.startswith("SystemsAnalysisWorkflows")):
                files = line
    fileInput.close()
    # Supprime de SystemsAnalysisWorkflows les retours à la ligne et les virgules
    # pour ensuite pouvoir récupérer uniquement les noms des fichiers
    # sans tout le reste
    files = files.replace("SystemsAnalysisWorkflows=", '', 1)
    files = files.replace("\n", '')
    files = files.split(', ')
    return files

# Récupère la variable SystemsAnalysisWorkflows de Revit.ini
def getLines():
    path = find(fileR, revitPath, 2)

    sysLin = ""

    fileInput = open(os.path.join(path, fileR), "r", encoding="utf-16")

    for line in fileInput:
        if (line.startswith("SystemsAnalysisWorkflows")):
                sysLin = line
                sysLin = sysLin.replace("\n", "")
    return sysLin

# Récupère la variable OpenStudio de Revit.ini
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

# Récupère le nom des fichiers de SystemsAnalysisWorkflows
fileName = files.copy()
ckVar = {}
# On met dans le tableau fileName les noms des fichiers
for i in range (len(fileName)):
    elmt = fileName[i]
    elmt = elmt.rsplit('\\', 1)
    fileName[i] = elmt[1]    
    ckVar[i] = tkinter.IntVar()

# Récupère les noms des fichiers sélectionnés par l'utilisateur
def getInfo(selected):
    for i in range(len(fileName)):
        if (ckVar[i].get() == 1):
            selected+=(files[i]+', ')

if (find('ABBem', destCheck, 1)!='notFound'):
    abbemPath = os.path.join(find('ABBem', destCheck, 1), 'ABBem')
else:
    abbemPath = -1
    


# Ferme la fenêtre
def closeWindow():
    getInfo(selected)
    window.destroy()

def reinitRevit():
    check = revitPath+'\original_Revit.ini'
    toDelete = revitPath+'\Revit.ini'
    if((str(os.path.exists(check)))!="True"):
        tkinter.messagebox.showinfo(title='réinitialisation', message='vos fichiers sont déjà réinitialisés')
    else:
        os.remove(toDelete)
        os.rename(os.path.join(revitPath, 'original_Revit.ini'), os.path.join(revitPath, 'Revit.ini'))
        tkinter.messagebox.showinfo(title='réinitialisation', message='la réinitialisation a été effectuée ')

# Fonction qui modifie Revit.ini
def changeRevit():
    # Récupère les fichiers sélectionnés par l'utilisateur
    getInfo(selected)

    # Vérifie si l'utilisateur à bien sélectionné des fichiers
    # affiche un message erreur si non
    if len(selected) == 0:
        # Fenêtre erreur
        popUpEr = Toplevel(bg="dimgrey")
        popUpEr.geometry("350x200+590+270")
        popUpEr.overrideredirect(True)

        # Texte erreur
        label = Label(popUpEr, text="ERREUR", fg="red", bg="dimgrey")
        labelset = ('Calibri (Body)', 14, 'bold')
        label.config(font=labelset)
        label.place(x=130, y=30)

        label2 = Label(popUpEr, text="Veuillez sélectionner les fichiers voulus.", bg="dimgrey")
        label2set = ('Calibri (Body)', 12, 'italic')
        label2.config(font=label2set)
        label2.place(x=35, y=80)

        # Boutton 'ok'
        btn = Button(popUpEr, text="OK", width=10, bg="lightgray", activebackground="white", relief=GROOVE, cursor="hand2", command= lambda: popUpEr.destroy())
        btn.place(x=140, y=150)
    else:
        selected.pop()
        selected.pop()
        # Fenêtre qui indique la fin du programme
        popUp = Toplevel(bg="dimgrey")
        popUp.geometry("350x200+590+270")
        popUp.overrideredirect(True)

        # Texte
        label = Label(popUp, text="Programme terminé. \n Fichiers installés avec succès.", bg="dimgrey")
        labelset = ('Calibri (Body)', 14, 'italic')
        label.config(font=labelset)
        label.place(x=40, y=70)

        popUp.after(1300, lambda: (popUp.destroy(), window.destroy()))
    
        # Récupère le chemin de Revit.ini
        path = find(fileR, revitPath, 2)

        # Ouverture et lecture Revit.ini
        fileInput = open(os.path.join(path, fileR), "r", encoding="utf-16")
        fileOutput = open(os.path.join(path, filer), "w", encoding="utf-16")

        # Récupère la variable SystemsAnalysisWorkflows
        var = "SystemsAnalysisWorkflows="
        for i in range(len(selected)):
            var+=selected[i]

        # Réecriture du fichier Revit.ini
        for line in fileInput:
            fileOutput.write(line.replace(openLine,"OpenStudio=C:\\ProgramData\ABBem\OStudio").replace(sysLine, var))

        # Fermeture des fichiers
        fileInput.close()
        fileOutput.close()

        # Renomme l'ancien Revit.ini en 'original_Revit.ini'
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


# Fenêtre du choix des fichiers à ajouter dans Revit.ini
def filesFrame():
    checkFrame.pack()

    # Texte instructions
    label = Label(checkFrame, text="Veuillez sélectionner les fichiers à intégrer")
    labelset = ("Calibri (Body)", 14, "bold")
    label.config(font=labelset)
    label.place(x=100, y=30)
    yPlace = 60

    # Création des Checkbuttons pour le choix des fichiers à ajouter dans Revit.ini
    for i in range (len(fileName)):
        ck = Checkbutton(checkFrame, text=fileName[i], variable=ckVar[i], onvalue=1, offvalue=0, cursor="hand2")
        yPlace+=30
        ck.place(x=200, y=yPlace)

    # Boutton 'ok et quitter'
    ex = Button(checkFrame, text = "OK et Quitter", width=10, bg="lightgray", activebackground="white", cursor="hand2", relief=GROOVE, command = changeRevit)
    ex.place(x=240, y=350)
    # Boutton réintialiser Revit.ini
    reinit = Button(checkFrame, text = "Réinitialiser Revit", width=15, bg="lightgray", activebackground="white", cursor="hand2", relief=GROOVE, command = reinitRevit)
    reinit.place(x=25, y=390)
    # Boutton aide
    help = Button(checkFrame, text = "Aide", width=10, bg="lightgray", activebackground="white", cursor="hand2", relief=GROOVE, command = helpBtn)
    help.place(x=25, y=70)

# Fonction qui déplace ABBem vers ProgramData
def popUpWindow():

    # Fenêtre temporaire indiquant le déplacement de ABBem vers ProgramData
    popUp = Toplevel(bg="dimgrey")
    popUp.geometry("350x200+590+270")
    popUp.overrideredirect(True)

    # Texte
    label = Label(popUp, text="Mooving files...", bg="dimgrey")
    labelset = ('Calibri (Body)', 14, 'italic')
    label.config(font=labelset)
    label.place(x=100, y=80)
    
    # Chemin de la localisation de l'exécutable
    currentPath = os.path.join(pathlib.Path().resolve(),'ABBem')

    # Vérifie si ABBem existe déjà dans ProgramData avant de déplacer le nouveau
    # supprime l'ancien si il existe
    if (abbemPath != -1):
        print("ABBEM PATH DELETE : "+abbemPath)
        shutil.rmtree(abbemPath)
    # Crée un répertoir ABBem dans ProgramData
    os.mkdir(dest)

    # Déplace ABBem dans ABBem/ProgramData
    shutil.move(currentPath, dest)

    popUp.after(2000, lambda: (filesFrame(), popUp.destroy()))


# Fonction qui s'occupe de l'installation
def mooveFiles():
    # Récupère le choix de l'utilisateur (installation complète ou non)
    install = choice.get()

    # Si l'utilisateur choisit 'installation complète'
    if install==0:
        # Ferme la frame
        instFrame.pack_forget()
        popUpWindow()
    else:
        # Si l'utilisateur choisit de juste modifier Revit.ini
        # on vérifie si ABBem existe déjà dans AppData
        # si il n'existe pas on affiche un message d'erreur
        if (abbemPath == -1):
            # Frame de l'erreur
            popUp = Toplevel(bg="dimgrey")
            popUp.geometry("400x250+580+230")
            popUp.overrideredirect(True)

            # Texte erreur
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
            # si AppData existe bien on continue
            instFrame.pack_forget()
            filesFrame()


# Frame où on choisit entre 'installation complète' et 'ajouter uniquement les fichiers'
def installFrame():
    # Ferme la frame verte
    splashScreen.pack_forget()

    # Ouvre cette frame
    instFrame.pack()

    # Texte instructions
    label = Label(instFrame, text="Veuillez choisir l'action à effectuer")
    labelset = ("Calibri (Body)", 14, "bold")
    label.config(font=labelset)
    label.place(x=140, y=80)

    # Bouttons radios
    r1 = Radiobutton(instFrame, text = "Installation complète", variable = choice, value = 0, cursor="hand2")
    r1.place(x=170, y=150)
    r2 = Radiobutton(instFrame, text = "Ajout des fichiers", variable = choice, value = 1, cursor="hand2")
    r2.place(x=170, y=200)


    # Texte en rouge
    warn = Label(instFrame, text="(uniquement si l'installation complète à déjà été effectuée)", fg="red")
    warn.place(x=190, y=220)

    # Boutton Suivant
    btn = Button(instFrame, text="Suivant", width=10, bg="lightgray", activebackground="white", cursor="hand2", relief=GROOVE, command=mooveFiles)
    btn.place(x=250, y=290)

# Timer qui fermer la 1ère fenêtre verte au bout du temps donnée
# (ici 3000ms soit 3s)
splashScreen.after(3000, installFrame)

window.mainloop()
