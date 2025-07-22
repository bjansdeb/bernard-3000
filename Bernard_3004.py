#init
from openpyxl import load_workbook
from urllib.request import urlopen
from tkinter import *
from tkinter import ttk
from tkinter.filedialog import askopenfilename
from tkinter.messagebox import showinfo
import modules

#TK
root = Tk() # création de la fenêtre racine
root.title("Extractor Bernard 3005")
root.geometry("720x720")

frame_home = Frame(root)
frame_prepare = Frame(root)
frame_extractor = Frame(root)
frame_epurator = Frame(root)
frame_regex = Frame(root)
frame_help = Frame(root)

def select_file():
    modules.choosing_file()

def preparator():
    modules.preparator()

def choixFichier():
    modules.choixRep()

def extractor_main():
    modules.extractor()

def choixEpur():
    modules.choixEpurator()
    progBar['value']=0

def epurator_main():
    # progBar['value']=0
    progBar.step(99)
    modules.epurator()

def assign_atom():
    modules.assign_at()

def assign_collective_access():
    modules.assign_ca()


def connector_main():
    if modules.connector():
        showinfo(message ='la connection marche')
    else:
        showinfo(message='la connection ne marche pas')

def exMachina():
    modules.regexMachina()


##Pages du logiciel
frame_home.pack()

def prepare_page():
    frame_prepare.pack()
    frame_home.forget()
    frame_extractor.forget()
    frame_regex.forget()
    frame_epurator.forget()
    frame_help.forget()
 
def extract_page():
    frame_extractor.pack()
    frame_home.forget()
    frame_prepare.forget()
    frame_regex.forget()
    frame_epurator.forget()
    frame_help.forget()
 
def regex_page():
    frame_regex.pack()
    frame_home.forget()
    frame_prepare.forget()
    frame_extractor.forget()
    frame_epurator.forget()
    frame_help.forget()
 
def epur_page():
    frame_epurator.pack()
    frame_home.forget()
    frame_prepare.forget()
    frame_extractor.forget()
    frame_regex.forget()
    frame_help.forget()
 
def help_page():
    frame_help.pack()
    frame_home.forget()
    frame_prepare.forget()
    frame_extractor.forget()
    frame_regex.forget()
    frame_epurator.forget()

###Fenêtre Root
#Menu démarrer
menu_bar = Menu(root)
menu_main = Menu(menu_bar, tearoff=0)
menu_main.add_command(label="Préparer un fichier au balisage", command= prepare_page)
menu_main.add_command(label="Extraire par balise", command=extract_page)
menu_main.add_command(label="Extraire par regex", command=regex_page)
menu_main.add_command(label="Epurer les données", command=epur_page)
menu_main.add_separator() #crée une ligne entre les options et quitter
menu_main.add_command(label="Quitter", command=root.destroy)
menu_bar.add_cascade(label="Démarrer", menu=menu_main)

menu_help = Menu(menu_bar, tearoff=0)
menu_help.add_command(label="A propos", command=help_page)
menu_bar.add_cascade(label="Aide", menu=menu_help) #un menu se termine toujours par l'ajout d'une cascade

root.config(menu=menu_bar)

imgBern = PhotoImage(file='Bern.png')
photo = Label(root, image= imgBern, height = 90, width =900)
photo.image = imgBern
photo.pack(ipadx=10, ipady=110)


####Home
label = Label(frame_home, text="Bonjour, bienvenue dans votre extracteur, le Bernard 3005 !", foreground="#892222", background="#FFAAAA", padx="10", pady="4")
label6 = Label(frame_home, text="Avant de commencer, vérifiez si les modules sont bien connectés\nSi ce n'est pas le cas, relancez le logiciel")
connect = Button (frame_home, text="Vérifier la connection des modules", command = connector_main)
label_home3 = Label(frame_home, text="Bernard vous remercie pour votre visite !")

label.pack(ipadx=10, ipady=20)
label6.pack()
connect.pack()
label_home3.pack()

####Prepare
introPrepare = Label (frame_prepare, text="Le préparateur vous permet d'ajouter automatiquement à votre fichier Word les styles requis pour le balisage. \n Il remplace aussi les tab par des retours à la ligne et peut taguer l'importance matérielle \n Le fichier créé s'appelle Bernard_auto\n Il est enregristré dans le répertoire du logiciel")
choixPrep = Button(frame_prepare, text="Choisir un fichier à préparer", command = select_file)
preparer = Button(frame_prepare, text="Lancer la préparation", command = preparator)

introPrepare.pack()
choixPrep.pack()
preparer.pack()

####Extractor
label2 = Label(frame_extractor, text="Pour commencer, cliquez sur le bouton 'Choisir un fichier' \n Seuls les fichiers htm et html sont autorisés.\n NB: vous pouvez facilement enregistrer vos fichiers doc en htm avec Word\n Assurez-vous également que ce fichier contient bien les styles suivants :\n Cote, Description, Datum, Importance et Note", padx="10", pady="4")
label3 = Label(frame_extractor, text="Lorsque vous avez choisi votre fichier, vous pouvez maintenant lancer l'extraction\n Bonne chance !", padx="10", pady="4")
label4 = Label(frame_extractor, text="Si l'extraction a réussi, un fichier xlxs a été créé dans le fichier courant\n Il porte le nom de Bernard", padx="10", pady="4")
choixRepBut = Button(frame_extractor, text="Choisir un fichier à extraire", command = choixFichier)
extract = Button(frame_extractor, text="Lancer l'extraction", command=extractor_main)

####Epurator
label5 = Label(frame_epurator, text="L'épuration permet de supprimer les caractères invisibles pouvant poser problème lors d'un import\n Le fichier épuré s'intitule Epure et se trouve dans le fichier courant", padx="10", pady="4")
choixEpurBut = Button(frame_epurator, text="Choisir un fichier à épurer", command = choixEpur)
epurer = Button(frame_epurator, text="Lancer l'épuration", command = epurator_main)
assign_at_button= Button(frame_epurator, text="Convertir en template AtoM", command = assign_atom)
assign_ca_button = Button(frame_epurator, text="Convertir en template Collective Access", command = assign_collective_access)

label2.pack()
choixRepBut.pack()
label3.pack()
extract.pack()
label4.pack(ipadx=10, ipady=10)
label5.pack(ipadx=10, ipady=10)
choixEpurBut.pack()
epurer.pack()
#Build Progressbar
progBar = ttk.Progressbar(frame_epurator,orient=HORIZONTAL, length=400,mode="determinate")
progBar.pack()
assign_at_button.pack()
assign_ca_button.pack()

####Regex
label_introRegex = Label(frame_regex, text="Les expressions régulières sont une méthode d'analyse de données reposant sur l'emploi de chaines de caractères typographiques \nExemple, le mot '1 liasse' pourra être reconnu avec 100% de certitude par l'expression '[0-9] l....e'\n Cet extracteur procède de la même manière\n L'extraction par regex est encore expérimentale\n Bernard fonctionne le mieux avec des fichiers word dont les données sont formatées comme suit : \n 1.[tab]Description[tab ou retour à la ligne]Date[tab]Importance matérielle")
choixRegex = Button(frame_regex, text="Choisir un fichier à regex", command = select_file)
regexer = Button(frame_regex, text="Lancer l'analyse regex", command=exMachina)

label_introRegex.pack()
choixRegex.pack()
regexer.pack()


####Help
label5 = Label(frame_help, text="Le Bernard 3000 est un petit logiciel d\'extraction de données \n Il procède à des extractions de deux façons différentes : par des balises (styles Word) et par Regex (expressions régulières) \n Il propose aussi un nettoyage de fichier xlsx avant un import \n Ce logiciel est développé par Benjamin Janssens de Bisthoven depuis 2021\n Il vise à aider tous les archivistes à convertir leurs anciens inventaires dans un esprit d'entraide et d'universalisme\n Il n'a pas vocation à être commercialisé")
label5.pack()

#lancement de la boucle principale
root.mainloop()
