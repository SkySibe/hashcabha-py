import shutil
import os
import time
from os.path import exists
import fileinput
import sys
import PySimpleGUIQt as sg
import pickle
import ctypes, sys

# TODO: check if it rund from exe or py
def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

if not is_admin():
    # Re-run the program with admin rights
    if sys.argv[1:] == []:
        ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, " ".join(sys.argv), None, 1)
    else:
        ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, " ".join(sys.argv[1:]), None, 1)

# getting the hebrew years and months from files
with open('files/modification parts/years.txt', 'r', encoding="utf8") as file:
    YEARS = file.read()
with open('files/modification parts/months.txt', 'r', encoding="utf8") as file:
    MONTHS = file.read()

#variables
delay = 0.05
debug = False
cli = False
niftarimDir = 'niftarim.bin'

#niftar object
class Niftar:
    def __init__(self, name, male, mother, maroqai, rab, day, month, year):
        self.name = name
        self.male = male
        self.mother = mother
        self.maroqai = maroqai
        self.rab = rab
        self.day = day
        self.month = month
        self.year = year
        if rab:
            if male:
                self.rabText = 'הרב'
            else:
                self.rabText = 'הרבנית'
        else:
            self.rabText = ''
        if maroqai:
            if male:
                self.mrqTxt = 'המרוקאי'
            else:
                self.mrqTxt = 'המרוקאית'
        else:
            self.mrqTxt = ''
        if male:
            self.sexText = 'בן'
        else:
            self.sexText = 'בת'
        self.fullDate = day + ' ב' + month + ' ' + year
        self.fullRabName = self.rabText + ' ' + name + ' ' + self.sexText + ' ' + mother
        self.fullName = name + ' ' + self.sexText + ' ' + mother
        self.info = self.fullRabName + ' ' + self.mrqTxt + ', תאריך פטירה: ' + self.fullDate
# -- functions --

# func 4 getting niftrim from file
def getNiftarim(varo):
    if varo == 'fullRabName':
        # checks if it isn't the first time if so gets the objects and returning them
        if exists(niftarimDir):
            with open(niftarimDir, 'rb') as f:
                listOrOne = pickle.load(f)
                if isinstance(listOrOne, list):
                    x = []
                    for w in listOrOne:
                        x = x + [w.fullRabName]
                    return x
                else:
                    return listOrOne.fullRabName
        else:
            return 'no'
    elif varo == 'obj':
        # checks if it isn't the first time if so gets the objects and returning them
        if exists(niftarimDir):
            with open(niftarimDir, 'rb') as f:
                listOrOne = pickle.load(f)
                return listOrOne
        else:
            return 'no'

# func 4 writing a niftar's object to file
def writeNiftar(niftar):
    # if it isn't the 1st time it uses the get func
    if exists(niftarimDir):
        niftarimList = getNiftarim('obj')
        # if there isn't more than one object
        if not isinstance(niftarimList, list):
            # it inserts the one niftar to be in a list
            niftarimList = list([niftarimList])
        # now it adds the new niftar to the list from the file
        niftarimList = niftarimList + [niftar]
        # eventually writes the merged list to the file
        with open(niftarimDir, 'wb') as f:
            pickle.dump(niftarimList, f)
    else:
        # it's the 1st time so it writes it without reading
        with open(niftarimDir, 'wb') as f:
            pickle.dump(niftar, f)

# arrange letters blocks from full niftar's name's letters
def nameLetterSq(strLtrs):
    # letters dictionary for storing the letters blocks within an accessible storage
    letters = {}
    # getting empty line block for separiding letters blocks
    with open('files/modification parts/emptyLine.xml', 'r', encoding="utf8") as file:
        EMPTYLINE = file.read().replace('\n', '')
    # letters blocks files into letters blocks dictionary
    # for over the letters blocks files
    for x in os.listdir('files/modification parts/letters'):
        # open them
        with open('files/modification parts/letters/'+x, 'r', encoding="utf8") as file:
            # saving file context (a block) to letters dictionary with key name like the file name but uppercase and without the file type
            letters[x.split('.', 1)[0].upper()] = file.read().replace('\n','')
    # variable for storing the blocks by name "building"
    nameSq = ''
    # for loop the full name's letters
    for i in strLtrs:
        # match a letter from name
        match i:
            # attaching letter block to the building variable with an empty line block after it
            case 'א':
                nameSq = nameSq + letters['ALEF']
                nameSq = nameSq + EMPTYLINE
            case 'ב':
                nameSq = nameSq + letters['BEYT']
                nameSq = nameSq + EMPTYLINE
            case 'ג':
                nameSq = nameSq + letters['GIMEL']
                nameSq = nameSq + EMPTYLINE
            case 'ד':
                nameSq = nameSq + letters['DALET']
                nameSq = nameSq + EMPTYLINE
            case 'ה':
                nameSq = nameSq + letters['HEY']
                nameSq = nameSq + EMPTYLINE
            case 'ו':
                nameSq = nameSq + letters['WAW']
                nameSq = nameSq + EMPTYLINE
            case 'ז':
                nameSq = nameSq + letters['ZAYIN']
                nameSq = nameSq + EMPTYLINE
            case 'ח':
                nameSq = nameSq + letters['HEYT']
                nameSq = nameSq + EMPTYLINE
            case 'ט':
                nameSq = nameSq + letters['TEYT']
                nameSq = nameSq + EMPTYLINE
            case 'י':
                nameSq = nameSq + letters['YUD']
                nameSq = nameSq + EMPTYLINE
            case 'כ' | 'ך':
                nameSq = nameSq + letters['CAF']
                nameSq = nameSq + EMPTYLINE
            case 'ל':
                nameSq = nameSq + letters['LAMED']
                nameSq = nameSq + EMPTYLINE
            case 'מ' | 'ם':
                nameSq = nameSq + letters['MEM']
                nameSq = nameSq + EMPTYLINE
            case 'נ' | 'ן':
                nameSq = nameSq + letters['NUN']
                nameSq = nameSq + EMPTYLINE
            case 'ס':
                nameSq = nameSq + letters['SAMEC']
                nameSq = nameSq + EMPTYLINE
            case 'ע':
                nameSq = nameSq + letters['AYIN']
                nameSq = nameSq + EMPTYLINE
            case 'פ' | 'ף':
                nameSq = nameSq + letters['PEY']
                nameSq = nameSq + EMPTYLINE
            case 'צ' | 'ץ':
                nameSq = nameSq + letters['SADI']
                nameSq = nameSq + EMPTYLINE
            case 'ק':
                nameSq = nameSq + letters['QUF']
                nameSq = nameSq + EMPTYLINE
            case 'ר':
                nameSq = nameSq + letters['REYS']
                nameSq = nameSq + EMPTYLINE
            case 'ש':
                nameSq = nameSq + letters['SIN']
                nameSq = nameSq + EMPTYLINE
            case 'ת':
                nameSq = nameSq + letters['TAW']
                nameSq = nameSq + EMPTYLINE
    # eventently return the building
    return nameSq
# function for making the modificated(sex+fName) hashcabha block
def haxcaba(nftr):
    if nftr.male:
        # getting hashcabha block for modificating and returning it
        with open('files/modification parts/hascabaLben.xml', 'r', encoding="utf8") as file:
            HASHCABHA = file.read().replace('\n', '')
    else:
        # getting hashcabha block for modificating and returning it
        with open('files/modification parts/hascabaLbat.xml', 'r', encoding="utf8") as file:
            HASHCABHA = file.read().replace('\n', '')
    HASHCABHA = HASHCABHA.replace('{{NAME}}', nftr.fullRabName)
    return HASHCABHA
# function for making Y to True and N to False
def boolear(yesOrNo):
    if yesOrNo == 'Y':
        yesOrNo = True
    else:
        yesOrNo = False
    return yesOrNo
# returning the suitable block if sefaradi or maroqai
def isMrq(nftr):
    if nftr.maroqai:
        with open('files/modification parts/mroqaiQraStn.xml', 'r', encoding="utf8") as file:
            PART = file.read().replace('\n', '')
    else:
        with open('files/modification parts/sfradiNesama.xml', 'r', encoding="utf8") as file:
            PART = file.read().replace('\n', '')
    return PART
if cli:
    if debug :
        theNiftar = Niftar("אהרן", True, "יוכבד", False, False, "א'", 'אב', 'ב`תפ"ז')
        magic()
    else:
        name = input("Enter niftar's name: ")
        male = boolear(input("Is male? (Y/N): "))
        mother = input("Enter niftar's mother name: ")
        rab = boolear(input("Is rab? (Y/N): "))
        maroqai = boolear(input("Is maroqai? (Y/N): "))
        day = input("Enter niftar's day of ptira: ")
        month = input("Enter niftar's month of ptira: ")
        year = input("Enter niftar's year of ptira: ")
        theNiftar = Niftar(name, male, mother, maroqai, rab, day, month, year)
        magic()
        
def magic(dirToSaveAt):
    #verifying if there is allready a copied folder of xmls
    if exists('TempoXMLs'):
        shutil.rmtree(os.getcwd()+'/TempoXMLs')
        time.sleep(delay)
    #copying xmls folder
    shutil.copytree('files/modificative', 'TempoXMLs')
    time.sleep(delay)

    # HERE the magic should happend --->

    # Read in the file
    with open('TempoXMLs/word/document.xml', 'r', encoding="utf8") as file :
        filedata = file.read()

    # writing letters capital 119 sequence by full name & Replace the target string in the main document
    filedata = filedata.replace('{{LETTERS}}', nameLetterSq(theNiftar.fullName))
    # modifinig the hashcabha block by name and picking it by sex & Inserting it
    filedata = filedata.replace('{{HAXCABH}}', haxcaba(theNiftar))
    # getting the right block
    filedata = filedata.replace('{{MRQY}}', isMrq(theNiftar))

    # - Now the Header -
    with open('TempoXMLs/word/header1.xml', 'r', encoding="utf8") as hFile :
        fData = hFile.read()
    # full (rab) name
    fData = fData.replace('{{NAME}}', theNiftar.fullRabName)
    # his/her date
    fData = fData.replace('{{DATE}}', theNiftar.fullDate)
    # Write the file out again
    with open('TempoXMLs/word/document.xml', 'w', encoding="utf8"
            ) as file:
        file.write(filedata)

    with open('TempoXMLs/word/header1.xml', 'w', encoding="utf8"
            ) as hFile:
        hFile.write(fData)

    # <--- End of modifiactions

    #zipping new cpied docx xsmls folder
    shutil.make_archive('newDocx', 'zip', 'TempoXMLs')
    time.sleep(delay)
    #removing folder
    shutil.rmtree(os.getcwd()+'/TempoXMLs')
    time.sleep(delay)
    #verifying if there is allready a modified docx file
    if exists('newDocx.docx'):
        os.remove('newDocx.docx')
        time.sleep(delay)
    #renaming zip file back to a docx file
    if dirToSaveAt is None:
        if os.name == 'nt':
            dirToSaveAt = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
        elif os.name == 'posix':
            dirToSaveAt = os.path.join(os.path.join(os.path.expanduser('~')), 'Desktop') 
        else:
            dirToSaveAt = ''
    dst = theNiftar.fullRabName+".docx"
    os.rename('newDocx.zip',dirToSaveAt+'/'+dst)

if not cli:
    theNiftarim = getNiftarim('obj')
    niftarimList = getNiftarim('fullRabName')
    if theNiftarim == 'no':
        niftarimList = ['אין נפטרים ברשימה']
    elif isinstance(theNiftarim, list):
        niftarimList = ['הוסף נפטר חדש'] + niftarimList
    else:
        niftarimList = ['הוסף נפטר חדש'] + [niftarimList]
    sg.theme('Black')   # Add a touch of color
    # All the stuff inside the window.
    layout = [  [sg.InputText(), sg.Text('שם הנפטר/ת:')],
                [sg.Radio('בן', "sex", key='male', default=True), sg.Radio('בת', "sex", key='female')],
                [sg.InputText(), sg.Text('שם אמו/ה:')],
                [sg.Checkbox('רב/נית'), sg.Checkbox('מרוקאי/ת')],
                [sg.Combo(YEARS.split(',')),sg.Combo(MONTHS.replace('"','').replace('\n','').split(',')),sg.Combo(YEARS.split(',')[:30])],
                [sg.Combo(niftarimList)],
                [sg.FolderBrowse('בחר תקיית יעד לשמירת הקובץ', key='targetDir')],
                [sg.Button('אישור'), sg.Button('ביטול')] ]

    # Create the Window with a title a text modificatin and an icon
    window = sg.Window('השכבה | לעילוי נשמת עמוס פרץ בן מז`לה (מזל)', layout, text_justification="right", icon='files/images/candleT.ico')
    
    # Event Loop to process "events" and get the "values" of the inputs
    while True:
        event, values = window.read()
        if event == sg.WIN_CLOSED or event == 'ביטול': # if user closes window or clicks cancel
            break
        if values[7] == 'אין נפטרים ברשימה' or values[7] == 'הוסף נפטר חדש':
            theNiftar = Niftar(values[0], values["male"], values[1], values[3], values[2], values[6], values[5], values[4])
            writeNiftar(theNiftar)
        else:
            if isinstance(theNiftarim, list):
                theNiftar = theNiftarim[(niftarimList.index(values[7]))-1]
            else:
                theNiftar = theNiftarim
        magic(values['targetDir'])

    window.close()
else:
    input("-<[ ENTER TO EXIT ]>-")