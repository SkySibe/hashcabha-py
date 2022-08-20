import shutil
import os
import time
from os.path import exists
import fileinput
import sys

#variables
niftar = "יעקב"
male = True
mother = "רבקה"
#maroqai = False
rab = True
delay = 0.2

#verifying if there is allready a copied folder of xmls
if exists('TempoXMLs'):
    shutil.rmtree(os.getcwd()+'/TempoXMLs')
    time.sleep(delay)
#copying xmls folder
shutil.copytree('xmls', 'TempoXMLs')
time.sleep(delay)

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