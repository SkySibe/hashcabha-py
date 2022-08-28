import shutil
import os
import time
from os.path import exists
import fileinput
import sys
import fileinput

#variables
delay = 0.05
debug = True

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
        if male:
            self.sexText = 'בן'
        else:
            self.sexText = 'בת'
        self.fullDate = day + ' ב' + month + ' ' + year
        self.fullRabName = self.rabText + ' ' + name + ' ' + self.sexText + ' ' + mother
        self.fullName = name + ' ' + self.sexText + ' ' + mother
# -- functions --
# arrange letters blocks from full niftar's name's letters
def nameLetterSq(strLtrs):
    # letters dictionary for storing the letters blocks within an accessible storage
    letters = {}
    # getting empty line block for separiding letters blocks
    with open('modification parts/emptyLine.xml', 'r', encoding="utf8") as file:
        EMPTYLINE = file.read().replace('\n', '')
    # letters blocks files into letters blocks dictionary
    # for over the letters blocks files
    for x in os.listdir('modification parts/letters'):
        # open them
        with open('modification parts/letters/'+x, 'r', encoding="utf8") as file:
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
        with open('modification parts/hascabaLben.xml', 'r', encoding="utf8") as file:
            HASHCABHA = file.read().replace('\n', '')
    else:
        # getting hashcabha block for modificating and returning it
        with open('modification parts/hascabaLbat.xml', 'r', encoding="utf8") as file:
            HASHCABHA = file.read().replace('\n', '')
    HASHCABHA = HASHCABHA.replace('{{NAME}}', nftr.fullRabName)
    return HASHCABHA
# returning the suitable block if sefaradi or maroqai
def isMrq(nftr):
    if nftr.maroqai:
        with open('modification parts/mroqaiQraStn.xml', 'r', encoding="utf8") as file:
            PART = file.read().replace('\n', '')
    else:
        with open('modification parts/sfradiNesama.xml', 'r', encoding="utf8") as file:
            PART = file.read().replace('\n', '')
    return PART

if debug :
    theNiftar = Niftar("אהרן", True, "יוכבד", False, False, "א'", 'אב', 'ב`תפ"ז')
else:
    name = input("Enter niftar's name: ")
    male = input("Is male? (Y/N): ")
    mother = input("Enter niftar's mother name: ")
    rab = input("Is rab? (Y/N): ")
    maroqai = input("Is maroqai? (Y/N): ")
    day = input("Enter niftar's day of ptira: ")
    month = input("Enter niftar's month of ptira: ")
    year = input("Enter niftar's year of ptira: ")
    theNiftar = Niftar(name, male, mother, maroqai, rab, day, month, year)

#verifying if there is allready a copied folder of xmls
if exists('TempoXMLs'):
    shutil.rmtree(os.getcwd()+'/TempoXMLs')
    time.sleep(delay)
#copying xmls folder
shutil.copytree('modificative', 'TempoXMLs')
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
fData = fData.replace('{{DATE}}', theNiftar.fullRabName)
# his/her date
fData = fData.replace('{{NAME}}', theNiftar.fullDate)
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
os.rename('newDocx.zip','newDocx.docx')
input("-<[ ENTER TO EXIT ]>-")