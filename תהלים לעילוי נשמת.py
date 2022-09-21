import shutil
import os
import time
from os.path import exists
import fileinput
import sys
import pickle
import ctypes, sys
from docx2pdf import convert
from configparser import ConfigParser
from pyluach import dates, hebrewcal, parshios

#variables
MODULE = 'qt'
HEBREW_DATE = str(dates.HebrewDate.today()).split('-')
config = ConfigParser()
SETTINGS_FILE = 'config.ini'
DELAY = 0
NIFTARIM_FILE = 'niftarim.bin'
INVALID_FILE_CHARTERS = '<>:"/\|?*'

if MODULE == 'qt':
    import PySimpleGUIQt as sg
elif MODULE == 'tk':
    import PySimpleGUI as sg
    from tkhtmlview import html_parser
    
    with open('files/modification parts/html/ad.html', 'r', encoding="utf8") as file:
        html = file.read()
    
    def set_html(widget, html, strip=True):
        prev_state = widget.cget('state')
        widget.config(state=sg.tk.NORMAL)
        widget.delete('1.0', sg.tk.END)
        widget.tag_delete(widget.tag_names)
        html_parser.w_set_html(widget, html, strip=strip)
        widget.config(state=prev_state)
    
    layout_advertise = [
        [sg.Multiline(
            size=(25, 10),
            border_width=2,
            text_color='white',
            background_color='green',
            disabled=True,
            no_scrollbar=True,
            expand_x=True,
            expand_y=True,
            key='Advertise')],
    ]

# function for getting input from the user
def getInput(title,txt,dt):
    layoutTextIn = [[sg.Text(txt)],
                [sg.InputText(default_text=dt, key='_INPUT_')],
                [sg.CloseButton('אישור', size=(60, 20), bind_return_key=True), sg.CloseButton('בטל', size=(60, 20))]]
    windowTextIn = sg.Window(title, layoutTextIn, text_justification="right", icon='files/images/candleT.ico')
    while True:
        event, values = windowTextIn.read()
        if event == sg.WIN_CLOSED or event == 'ביטול': # if user closes window or clicks cancel
            return None
            break
        elif event == 'אישור':
            return values['_INPUT_']
        break
    windowTextIn.close()

# getting the hebrew years and months from files
with open('files/modification parts/years.txt', 'r', encoding="utf8") as file:
    YEARS = file.read()
with open('files/modification parts/months.txt', 'r', encoding="utf8") as file:
    MONTHS = file.read()

#niftar object
class Niftar:
    def __init__(self, name, lastName, male, mother, eda, rab, day, month, year):
        self.name = name
        self.lastName = lastName
        self.male = male
        self.mother = mother
        self.eda = eda
        self.rab = rab
        self.day = day
        self.month = month
        self.year = year
        if rab:
            if male:
                self.rabText = 'הרב '
            else:
                self.rabText = 'הרבנית '
        else:
            self.rabText = ''
        if eda == 'mrqo':
            if male:
                self.mrqTxt = 'המרוקאי'
            else:
                self.mrqTxt = 'המרוקאית'
        elif eda == 'asknz':
            if male:
                self.mrqTxt = 'האשכנזי'
            else:
                self.mrqTxt = 'האשכנזיה'
        elif eda == 'sfrd':
            if male:
                self.mrqTxt = 'הספרדי'
            else:
                self.mrqTxt = 'הספרדיה'
        else:
            self.mrqTxt = ''
        if male:
            self.sexText = 'בן'
        else:
            self.sexText = 'בת'
        self.fullDate = day + ' ב' + month + ' ' + year
        self.nameWlast = self.name + ' (' + self.lastName + ') '
        self.fullRabNameWithLastName = self.rabText + self.nameWlast + self.sexText + ' ' + mother
        self.fullRabName = self.rabText + self.name + ' ' + self.sexText + ' ' + mother
        self.fullName = self.name + ' ' + self.sexText + ' ' + mother
        self.info = self.fullRabNameWithLastName + ' ' + self.mrqTxt + ', תאריך פטירה: ' + self.fullDate
# -- functions --

# func 4 getting niftrim from file
def getNiftarim(varo):
    if varo == 'fullRabNameWithLastName':
        # checks if it isn't the first time if so gets the objects and returning them
        if exists(NIFTARIM_FILE):
            with open(NIFTARIM_FILE, 'rb') as f:
                listOrOne = pickle.load(f)
                if isinstance(listOrOne, list):
                    x = []
                    for w in listOrOne:
                        # gives it a longer name if it's same as was one before
                        if w.fullRabNameWithLastName in x:
                            x = x + [w.info]
                        else:
                            x = x + [w.fullRabNameWithLastName]
                    return x
                else:
                    return listOrOne.fullRabNameWithLastName
        else:
            return 'no'
    elif varo == 'obj':
        # checks if it isn't the first time if so gets the objects and returning them
        if exists(NIFTARIM_FILE):
            with open(NIFTARIM_FILE, 'rb') as f:
                listOrOne = pickle.load(f)
                return listOrOne
        else:
            return 'no'

# func 4 writing a niftar's object to file
def writeNiftar(niftar):
    
    # if it isn't the 1st time it uses the get func
    if exists(NIFTARIM_FILE):
        theNiftarimAsNames = getNiftarim('obj')
        if isinstance(theNiftarimAsNames, list):
            if niftar in theNiftarimAsNames:
                return 'Niftar is already exists'
        else: # if there isn't more than one object
            if niftar == theNiftarimAsNames:
                return 'Niftar is already exists'
            else: # and it's a new niftarS
                # it inserts the one niftar to be in a list
                theNiftarimAsNames = list([theNiftarimAsNames])
        # now it adds the new niftar to the list from the file
        theNiftarimAsNames = theNiftarimAsNames + [niftar]
        # eventually writes the merged list to the file
        with open(NIFTARIM_FILE, 'wb') as f:
            pickle.dump(theNiftarimAsNames, f)
    else:
        # it's the 1st time so it writes it without reading
        with open(NIFTARIM_FILE, 'wb') as f:
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
    for index,i in enumerate(strLtrs):
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
    HASHCABHA = HASHCABHA.replace('{{NAME}}', nftr.fullName)
    return HASHCABHA
# function for making Y to True and N to False
def boolear(yesOrNo):
    if yesOrNo == 'Y':
        yesOrNo = True
    else:
        yesOrNo = False
    return yesOrNo
# returning the suitable block for its eda
def isMrq(nftr):
    if nftr.eda == 'mrqo':
        with open('files/modification parts/mroqaiQraStn.xml', 'r', encoding="utf8") as file:
            PART = file.read().replace('\n', '')
    else:
        with open('files/modification parts/sfradiNesama.xml', 'r', encoding="utf8") as file:
            PART = file.read().replace('\n', '')
    return PART
        
def magic(dirFromInput,wordT,pdfT,launchAfterCreating):
    # lenOfNameSq = len(theNiftar.fullName)
    # match :
    #     case >
    #verifying if there is allready a copied folder of xmls
    if exists('TempoXMLs'):
        shutil.rmtree(os.getcwd()+'/TempoXMLs')
        time.sleep(DELAY)
    #copying xmls folder
    if theNiftar.eda == 'asknz':
        shutil.copytree('files/modificativeAshkenazi', 'TempoXMLs')
    else:
        shutil.copytree('files/modificative', 'TempoXMLs')
    time.sleep(DELAY)

    # HERE the magic should happend --->

    # Read in the file
    with open('TempoXMLs/word/document.xml', 'r', encoding="utf8") as file :
        filedata = file.read()

    # writing letters capital 119 sequence by full name & Replace the target string in the main document
    filedata = filedata.replace('{{LETTERS}}', nameLetterSq(theNiftar.fullName))
    if theNiftar.eda != 'asknz':
        # modifinig the hashcabha block by name and picking it by sex & Inserting it
        filedata = filedata.replace('{{HAXCABH}}', haxcaba(theNiftar))
    # getting the right block
    filedata = filedata.replace('{{MRQY}}', isMrq(theNiftar))

    # - Now the Header -
    with open('TempoXMLs/word/header1.xml', 'r', encoding="utf8") as hFile :
        fData = hFile.read()
    # full (rab) name
    fData = fData.replace('{{NAME}}', ' ' + theNiftar.fullRabNameWithLastName)
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
    time.sleep(DELAY)
    #removing folder
    shutil.rmtree(os.getcwd()+'/TempoXMLs')
    time.sleep(DELAY)
    #verifying if there is allready a modified docx file
    if exists('newDocx.docx'):
        os.remove('newDocx.docx')
        time.sleep(DELAY)
    
    # is didn't got a path from the user setting it to the Desktop as defaultd
    if dirFromInput is None or dirFromInput == 'None':
        # by OS
        if os.name == 'nt':
            dirFromInput = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
        elif os.name == 'posix':
            dirFromInput = os.path.join(os.path.join(os.path.expanduser('~')), 'Desktop') 
        else:
            dirFromInput = ''
    #renaming zip file back to a docx file
    dst = theNiftar.fullRabNameWithLastName
    if exists(dirFromInput+'/'+dst+".docx"):
        # if there's alrady file named the same asking the use for a name
        inFromUser = getInput('שמור בשם','  קובץ באותו בשם כבר קיים \n אפשר לבחור לו שם או לבטל',theNiftar.fullRabNameWithLastName)
        if inFromUser == None:
            sg.Popup('לא נשמר קובץ',title="פרשת דרכים",background_color='black',button_color=('white', '#5555ff'),custom_text='אישור')
            wordT = False
            pdfT = False
        else:
            for i in INVALID_FILE_CHARTERS:
                inFromUser = inFromUser.replace(i,"") # removing invalid chartres that can't be saved as file on wndows
            x = 1
            while exists(dirFromInput+'/'+inFromUser+".docx"): # adding number to the name if user didn't renamed the file
                if x > 1:
                    inFromUser =inFromUser[:-1]
                print(inFromUser,inFromUser[:-1])
                print(x)
                inFromUser = inFromUser + str(x)
                x = x + 1
            name = inFromUser
    else:
        name = dst
    if pdfT or wordT:
        os.rename('newDocx.zip',dirFromInput+'/'+name+".docx")
    if pdfT:
        convert(dirFromInput+'/'+name+".docx",dirFromInput+'/'+name+".pdf") # coverting the docx file to a pdf
    if not wordT:
       os.remove(dirFromInput+'/'+name+".docx") # removes the docx file if the user didn't marked it
    if launchAfterCreating: # launches the files if asked
        if pdfT:
            os.startfile(dirFromInput+'/'+name+".pdf")
        if wordT:
            os.startfile(dirFromInput+'/'+name+".docx")


niftarimAsObjects = getNiftarim('obj') # getting the existing niftarim as written (as objects) | for creating them again
niftarimAsNames = getNiftarim('fullRabNameWithLastName') # getting the existing niftarim as names | for displaying them at the GUI
if niftarimAsObjects == 'no': # it tells us there is no niftarim saved
    niftarimAsNames = ['אין נפטרים ברשימה']
elif isinstance(niftarimAsObjects, list): # if there are (more then one)
    niftarimAsNames = ['הוסף נפטר חדש'] + niftarimAsNames # adds the list a creating new option
else:
    niftarimAsNames = ['הוסף נפטר חדש'] + [niftarimAsNames] # creating a list with the one niftar and the new option 2.Le.Li
config.read(SETTINGS_FILE) # preperes the file with saved settings (cookies like) in order to O/I
if exists(SETTINGS_FILE): # if it exists
    # information from the settings file to be attached to the GUI inputs
    word = config.getboolean('settings', 'Word')
    pdf = config.getboolean('settings', 'PDF')
    autoLaunch = config.getboolean('settings', 'AL')
    dirFromCookie = config.get('settings', 'DTS')
else: # if there are no settings set
    # Settings Variables
    word = True
    pdf = False
    autoLaunch = False
    dirFromCookie = None
    config.add_section('settings') # adds a section named settings on the file to save the parameters under
years = YEARS.split(',') # the hebrew years string from its file to list of years
# same with months and days
months = MONTHS.replace('"','').replace('\n','').split(',')
days = YEARS.split(',')[:30]

# design
font = ("Raleway", 12, 'bold')
sg.set_options(font=font)
sg.theme('Black')   # Add a touch of color
# sg.theme("systemdefaultforreal")

layout = [  [sg.InputText(key='lastName'), sg.Text('שם משפחתו/ה:') ,sg.InputText(key='name'), sg.Text('שם הנפטר/ת:')],
            [sg.Radio('בן', "sex", key='male', default=True), sg.Radio('בת', "sex", key='female')],
            [sg.InputText(key='mother'), sg.Text('שם אמו/ה:')],
            [sg.Radio('אשכנזי/ה', "eda", key='asknz'), sg.Radio('ספרדי/ה', "eda", key='sfrd', default=True), sg.Radio('מרוקאי/ת', "eda", key='mrqo'), sg.Checkbox('רב/נית', key='rab')],
            [sg.Combo(years, default_value=years[int(HEBREW_DATE[0])-5001], key='year', readonly=True),sg.Combo(months, default_value=months[int(HEBREW_DATE[1])-1], key='month', readonly=True),sg.Combo(days,default_value=days[int(HEBREW_DATE[2])-1], key='day', readonly=True)],
            [sg.Combo(niftarimAsNames, key='niftarFromList', readonly=True)],
            [sg.FolderBrowse('בחר תקיית יעד לשמירת הקובץ', initial_folder=dirFromCookie, key='targetDir')],
            [sg.Checkbox('וורד', default=word, key='word'), sg.Checkbox('פי.די.אף', default=pdf, key='pdf'), sg.Checkbox('פתח קובץ/ים', default=autoLaunch, key='atla')],
            [sg.Button('אישור'), sg.Button('ביטול')]
            ]

if MODULE == 'tk':
    layout = layout + [[ sg.Frame("Advertise",  layout_advertise, expand_x=True, expand_y=True) ]]
    layout = [[sg.Frame(None ,layout ,element_justification="right")]]

window = sg.Window('השכבה | לעילוי נשמת עמוס פרץ בן מז`לה (מזל)', text_justification="right", icon='files/images/candleT.ico',finalize=True, use_default_focus=False).Layout(layout).Finalize()

if MODULE == 'tk':
    advertise = window['Advertise'].Widget
    
    html_parser = html_parser.HTMLTextParser()
    set_html(advertise, html)
    width, height = advertise.winfo_width(), advertise.winfo_height()

# Event Loop to process "events" and get the "values" of the inputs
while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'ביטול': # if user closes window or clicks cancel
        break
    # getting info from the inputs to save cookies
    config.set('settings', 'Word', str(values['word']))
    config.set('settings', 'PDF', str(values['pdf']))
    config.set('settings', 'AL', str(values['atla']))
    if values['targetDir'] is not None:
        dirFromCookie = values['targetDir']
        config.set('settings', 'DTS', str(dirFromCookie))
    elif dirFromCookie is None or dirFromCookie == 'None':
        config.set('settings', 'DTS', str(None))
    with open(SETTINGS_FILE, 'w') as f:
        config.write(f) # writes the setting files with the info from GUI inputs
    if values['asknz']:
        eeda = 'asknz'
    elif values['sfrd']:
        eeda = 'sfrd'
    elif values['mrqo']:
        eeda = 'mrqo'
    if values['niftarFromList'] == 'אין נפטרים ברשימה' or values['niftarFromList'] == 'הוסף נפטר חדש': # if I need to get the information from GUI inputs
        theNiftar = Niftar(values['name'], values['lastName'], values["male"], values['mother'], eeda, values['rab'], values['day'], values['month'], values['year'])
        wn = writeNiftar(theNiftar) # trying to write the new niftar to the niftarim.bin file
        if wn == 'Niftar is already exists': # if it's exist it asks the user if he wants to create the file anyway
            outFromPU = sg.Popup('נפטר זה קיים ברשימה','האם ליצור לו קובץ בכל זאת?','(בלי לשמור אותו ברשימה שוב)',title="פרשת דרכים",background_color='black',button_color=('white', '#5555ff'),custom_text=('כן','לא'))
            if 'כן':
                magic(dirFromCookie, values['word'], values['pdf'], values['atla'])
        else:
            magic(dirFromCookie, values['word'], values['pdf'], values['atla'])
    else:
        # How to treat the var as a sring or as a list?
        if isinstance(niftarimAsObjects, list): # if there are more then one niftar
            theNiftar = niftarimAsObjects[(niftarimAsNames.index(values['niftarFromList']))-1] # minus one because of the add new option
        else:
            theNiftar = niftarimAsObjects # if it's only one
        magic(dirFromCookie, values['word'], values['pdf'], values['atla'])
#input('x')
window.close()