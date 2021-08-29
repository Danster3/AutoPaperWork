import PySimpleGUI as sg
from PySimpleGUI.PySimpleGUI import Input
import pandas as pd
import os
import sys  # Standard Python Libraries
from docxtpl import DocxTemplate, InlineImage  # pip install docxtpl
from docx.shared import Cm, Inches, Mm, Emu  # pip install python-docx
from datetime import date

# path for Word document same as the program path
os.chdir(sys.path[0])

# Add some color to the window
sg.theme('LightBlue5')

# Create Data
EXCEL_FILE = 'Data_Entry.xlsx'
df = pd.read_excel(EXCEL_FILE)

People = ["Daniel Forbes",
          "David Downer",
          "Alan Lokking",
          "Greg Simmonds"
          ]


# Input box Layout
layout = [
    [sg.Text('Please fill out the following fields:')],
    [sg.T('Name', size=(15, 1)), sg.Combo(People, key='Name')],
    [sg.Text('Site', size=(15, 1)), sg.InputText(key='Site')],
    [sg.Text('System', size=(15, 1)), sg.Combo(
        ['iSTAR', 'mSTAR', 'WXT_WMT'], key='System')],
    [sg.Text('Document', size=(15, 1)), sg.Combo(
        ['Calabration', 'Field Test Sheet', 'Other'], key='Document')],
    # [sg.Text('Befores/Afters', size=(15,1)),
    #                         sg.Checkbox('Befores', key='Befores'),
    #                         sg.Checkbox('Afters', key='Afters')],
    [sg.Text('Befores', size=(15, 1))],
    [sg.Multiline(size=(60, 5), key='BeforesCSD')],
    [sg.Text('Afters', size=(15, 1))],
    [sg.Multiline(size=(60, 5), key='AftersCSD')],
    [sg.Submit(), sg.Button('Clear'), sg.Exit()]
]

window = sg.Window('Paperwork Automation', layout)

# Clear input field


def clear_input():
    for key in values:
        window[key]('')
    return None


# Check values from input box
while True:
    event, values = window.read()

    Name = values['Name']
    Site = values['Site']
    System = values['System']
    Document = values['Document']
    BeforesCSD = values['BeforesCSD']
    AftersCSD = values['AftersCSD']

    if event == sg.WIN_CLOSED or event == 'Exit':
        break
    if event == 'Clear':
        clear_input()
    if event == 'Submit':
        df = df.append(values, ignore_index=True)
        df.to_excel(EXCEL_FILE, index=False)
        sg.popup('Data saved!')
        # clear_input()
window.close()

doctype = "{}".format(Document)

if(doctype == "Calabration"):
    doctype = "Cal"
    print("Cal")
elif (doctype == "Field Test Sheet"):
    doctype = "FTS"
    print("Field Test Sheet")
elif(doctype == "Other"):
    doctype = "Other"
    print("Other")

# Creates txt file to read and use panda libray
f = open("Befores.txt", "w+")
f.write(BeforesCSD)
f.close()

f = open("Afters.txt", "w+")
f.write(AftersCSD)
f.close()
# ////////////////////////////////////////// Parser //////////////////////////////

CsdBefores = pd.read_csv("Befores.txt", skiprows=2, sep=',', na_values=[''], names=[
    "Type", "Status", "Value", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16", "17", "18"])

CsdAfters = pd.read_csv("Afters.txt", skiprows=2, sep=',', na_values=[''], names=[
    "Type", "Status", "Value", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16", "17", "18"])

typeBefores = CsdBefores["Type"]
typeAfters = CsdAfters["Type"]

# Software version
AABeforeVales1 = ""
AABeforeVales2 = ""
AABeforeVales3 = ""
AABeforeVales4_6 = ""
AABeforeVales7_9 = ""

# House Keeping
HKBeforeStatus1 = ""
HKBeforeVales2 = ""
HKBeforeVales4 = ""
HKBeforeVales6 = ""

# Barametric Pressure
BPBeforeStatus1 = ""
BPBeforeValues2 = ""
BPBeforeValues3 = ""
BPBeforeValues5 = ""

# Bright Sunshine
BSBeforeStatus1 = ""
BSBeforeValues2 = ""

# DS
DSBeforeStatus1 = ""
DSBeforeValues2 = ""
DSBeforeValues3 = ""
DSBeforeValues4 = ""

# FX
FXBeforeStatus1 = ""
FXBeforeValues2_3_4 = ""
FXBeforeValues5_6_7 = ""
FXBeforeValues8 = ""

# Percipitation
PRBeforeStatus1 = ""
PRBeforeValues2 = ""

# Road sensor
RCBeforeStatus1 = ""
RCBeforeValues2 = ""
RCBeforeValues4 = ""
RCBeforeValues6 = ""
RCBeforeValues8 = ""
RCBeforeValues10 = ""
RCBeforeValues12 = ""
RCBeforeValues14 = ""
RCBeforeValues16 = ""
RCBeforeValues18 = ""

# Relitive Humidity
RHBeforeStatus1 = ""
RHBeforeValues2 = ""
RHBeforeValues3 = ""

# Snow Depth
SDBeforeStatus1 = ""
SDBeforeValues2 = ""
SDBeforeValues3 = ""
SDBeforeValues4 = ""
SDBeforeValues5 = ""

# Soil Mostiure
SMBeforeValues2 = ""

# Solar Radiation
SRBeforeStatus1 = ""
SRBeforeValues2 = ""

# Sea State
SSBeforeStatus1 = ""
SSBeforeValues3 = ""
SSBeforeValues4_5 = ""
SSBeforeValues6_7 = ""
SSBeforeValues8 = ""
SSBeforeValues10 = ""
SSBeforeValues11 = ""
SSBeforeValues12 = ""
SSBeforeValues13 = ""
SSBeforeValues14 = ""
SSBeforeValues15 = ""

# Temp 100cm
T0BeforeStatus1 = ""
T0BeforeValues2 = ""

# Temp 10cm
T1BeforeStatus1 = ""
T1BeforeValues2 = ""

# Temp 20cm
T2BeforeStatus1 = ""
T2BeforeValues2 = ""

# Temp 50cm
T5BeforeStatus1 = ""
T5BeforeValues2 = ""

# Temp Air
TABeforeStatus1 = ""
TABeforeValues2 = ""

# Temp Road
TEBeforeStatus1 = ""
TEBeforeValues2 = ""

# Temp Surface
TSBeforeStatus1 = ""
TSBeforeValues2 = ""

# Temp 5cm
TUBeforeStatus1 = ""
TUBeforeValues2 = ""

# Temp X
TXBeforeStatus1 = ""
TXBeforeValues2 = ""

# Temp Y
TYBeforeStatus1 = ""
TYBeforeValues2 = ""

# Temp Z
TZBeforeStatus1 = ""
TZBeforeValues2 = ""

# Visability
VBBeforeStatus1 = ""
VBBeforeValues2 = ""

# Wind Direction
WDBeforeStatus1 = ""
WDBeforeValues2 = ""

# Instant Wind Direction and Speed
WRBeforeStatus1 = ""
WRBeforeValues2 = ""
WRBeforeValues3 = ""

# Wind Speed
WSBeforeStatus1 = ""
WSBeforeValues2 = ""

# Present Wind
WXBeforeStatus1 = ""
WXBeforeValues2 = ""

# /////////////////After Values///////////////

# Software version
AAAfterVales1 = ""
AAAfterVales2 = ""
AAAfterVales3 = ""
AAAfterVales4_6 = ""
AAAfterVales7_9 = ""

# House Keeping
HKAfterStatus1 = ""
HKAfterVales2 = ""
HKAfterVales4 = ""
HKAfterVales6 = ""

# Barametric Pressure
BPAfterStatus1 = ""
BPAfterValues2 = ""
BPAfterValues3 = ""
BPAfterValues5 = ""

# Bright Sunshine
BSAfterStatus1 = ""
BSAfterValues2 = ""

# DS
DSAfterStatus1 = ""
DSAfterValues2 = ""
DSAfterValues3 = ""
DSAfterValues4 = ""

# FX
FXAfterStatus1 = ""
FXAfterValues2_3_4 = ""
FXAfterValues5_6_7 = ""
FXAfterValues8 = ""

# Percipitation
PRAfterStatus1 = ""
PRAfterValues2 = ""

# Road sensor
RCAfterStatus1 = ""
RCAfterValues2 = ""
RCAfterValues4 = ""
RCAfterValues6 = ""
RCAfterValues8 = ""
RCAfterValues10 = ""
RCAfterValues12 = ""
RCAfterValues14 = ""
RCAfterValues16 = ""
RCAfterValues18 = ""

# Relitive Humidity
RHAfterStatus1 = ""
RHAfterValues2 = ""
RHAfterValues3 = ""

# Snow Depth
SDAfterStatus1 = ""
SDAfterValues2 = ""
SDAfterValues3 = ""
SDAfterValues4 = ""
SDAfterValues5 = ""

# Soil Mostiure
SMAfterValues2 = ""

# Solar Radiation
SRAfterStatus1 = ""
SRAfterValues2 = ""

# Sea State
SSAfterStatus1 = ""
SSAfterValues3 = ""
SSAfterValues4_5 = ""
SSAfterValues6_7 = ""
SSAfterValues8 = ""
SSAfterValues10 = ""
SSAfterValues11 = ""
SSAfterValues12 = ""
SSAfterValues13 = ""
SSAfterValues14 = ""
SSAfterValues15 = ""

# Temp 100cm
T0AfterStatus1 = ""
T0AfterValues2 = ""

# Temp 10cm
T1AfterStatus1 = ""
T1AfterValues2 = ""

# Temp 20cm
T2AfterStatus1 = ""
T2AfterValues2 = ""

# Temp 50cm
T5AfterStatus1 = ""
T5AfterValues2 = ""

# Temp Air
TAAfterStatus1 = ""
TAAfterValues2 = ""

# Temp Road
TEAfterStatus1 = ""
TEAfterValues2 = ""

# Temp Surface
TSAfterStatus1 = ""
TSAfterValues2 = ""

# Temp 5cm
TUAfterStatus1 = ""
TUAfterValues2 = ""

# Temp X
TXAfterStatus1 = ""
TXAfterValues2 = ""

# Temp Y
TYAfterStatus1 = ""
TYAfterValues2 = ""

# Temp Z
TZAfterStatus1 = ""
TZAfterValues2 = ""

# Visability
VBAfterStatus1 = ""
VBAfterValues2 = ""

# Wind Direction
WDAfterStatus1 = ""
WDAfterValues2 = ""

# Instant Wind Direction and Speed
WRAfterStatus1 = ""
WRAfterValues2 = ""
WRAfterValues3 = ""

# Wind Speed
WSAfterStatus1 = ""
WSAfterValues2 = ""

# Present Wind
WXAfterStatus1 = ""
WXAfterValues2 = ""

i = 0  # indexting for 'for' loop
# Befores Parser
for x in typeBefores:
    print(x)
    if(x == "$AA"):
        AABeforeVales1 = CsdBefores["Status"][i]
        AABeforeVales2 = CsdBefores["Value"][i]
        AABeforeVales3 = CsdBefores["3"][i]
        AABeforeVales4_6 = "{}-{}-{}".format(
            CsdBefores["4"][i], CsdBefores["5"][i], CsdBefores["6"][i])
        AABeforeVales7_9 = "{}-{}-{}".format(
            CsdBefores["7"][i], CsdBefores["8"][i], CsdBefores["9"][i])
    if(x == "$HK"):
        HKBeforeStatus1 = CsdBefores["Status"][i]
        HKBeforeVales2 = CsdBefores["Value"][i]
        HKBeforeVales4 = CsdBefores["4"][i]
        HKBeforeVales6 = CsdBefores["6"][i]
    elif(x == "$BP"):
        BPBeforeStatus1 = CsdBefores["Status"][i]
        BPBeforeValues2 = CsdBefores["Value"][i]
        BPBeforeValues3 = CsdBefores["3"][i]
        BPBeforeValues5 = CsdBefores["5"][i]
    elif(x == "$BS"):
        BSBeforeStatus1 = CsdBefores["Status"][i]
        BSBeforeValues2 = CsdBefores["Value"][i]
    elif(x == "$DS"):
        DSBeforeStatus1 = CsdBefores["Status"][i]
        DSBeforeValues2 = CsdBefores["Value"][i]
        DSBeforeValues3 = CsdBefores["3"][i]
        DSBeforeValues4 = CsdBefores["4"][i]
    elif(x == "$FX"):
        FXBeforeStatus1 = CsdBefores["Status"][i]
        FXBeforeValues2_3_4 = "{}-{}-{}".format(
            CsdBefores["2"][i], CsdBefores["3"][i], CsdBefores["4"][i])
        FXBeforeValues5_6_7 = "{}-{}-{}".format(
            CsdBefores["5"][i], CsdBefores["6"][i], CsdBefores["7"][i])
        FXBeforeValues8 = CsdBefores["8"][i]
    elif(x == "$PR"):
        PRBeforeStatus1 = CsdBefores["Status"][i]
        PRBeforeValues2 = CsdBefores["Value"][i]
    elif(x == "$RC"):
        RCBeforeStatus1 = CsdBefores["Status"][i]
        RCBeforeValues2 = CsdBefores["Value"][i]
        RCBeforeValues4 = CsdBefores["4"][i]
        RCBeforeValues6 = CsdBefores["6"][i]
        RCBeforeValues8 = CsdBefores["8"][i]
        RCBeforeValues10 = CsdBefores["10"][i]
        RCBeforeValues12 = CsdBefores["12"][i]
        RCBeforeValues14 = CsdBefores["14"][i]
        RCBeforeValues16 = CsdBefores["16"][i]
        RCBeforeValues18 = CsdBefores["18"][i]
    elif(x == "$RH"):
        RHBeforeStatus1 = CsdBefores["Status"][i]
        RHBeforeValues2 = CsdBefores["Value"][i]
        RHBeforeValues3 = ""

    elif(x == "$SD"):
        SDBeforeStatus1 = CsdBefores["Status"][i]
        SDBeforeValues2 = CsdBefores["Value"][i]
        SDBeforeValues3 = CsdBefores["3"][i]
        SDBeforeValues4 = CsdBefores["4"][i]
        SDBeforeValues5 = CsdBefores["5"][i]
    elif(x == "$SM"):
        SMBeforeStatus1 = CsdBefores["Status"][i]
        SMBeforeValues2 = CsdBefores["Value"][i]
    elif(x == "$SR"):
        SRBeforeStatus1 = CsdBefores["Status"][i]
        SRBeforeValues2 = CsdBefores["Value"][i]
    elif(x == "$SS"):
        SSBeforeStatus1 = CsdBefores["Status"][i]
        SSBeforeValues3 = CsdBefores["3"][i]
        SSBeforeValues4_5 = "{},{}".format(
            CsdBefores["4"][i], CsdBefores["5"][i])
        SSBeforeValues6_7 = "{},{}".format(
            CsdBefores["6"][i], CsdBefores["7"][i])
        SSBeforeValues8 = CsdBefores["8"][i]
        SSBeforeValues10 = CsdBefores["10"][i]
        SSBeforeValues11 = CsdBefores["11"][i]
        SSBeforeValues12 = CsdBefores["12"][i]
        SSBeforeValues13 = CsdBefores["13"][i]
        SSBeforeValues14 = CsdBefores["14"][i]
        SSBeforeValues15 = CsdBefores["15"][i]
    elif(x == "$T0"):
        T0BeforeStatus1 = CsdBefores["Status"][i]
        T0BeforeValues2 = CsdBefores["Value"][i]
    elif(x == "$T1"):
        T1BeforeStatus1 = CsdBefores["Status"][i]
        T1BeforeValues2 = CsdBefores["Value"][i]
    elif(x == "$T2"):
        T2BeforeStatus1 = CsdBefores["Status"][i]
        T2BeforeValues2 = CsdBefores["Value"][i]
    elif(x == "$T5"):
        T5BeforeStatus1 = CsdBefores["Status"][i]
        T5BeforeValues2 = CsdBefores["Value"][i]
    elif(x == "$TA"):
        TABeforeStatus1 = CsdBefores["Status"][i]
        TABeforeValues2 = CsdBefores["Value"][i]
    elif(x == "$TE"):
        TEBeforeStatus1 = CsdBefores["Status"][i]
        TEBeforeValues2 = CsdBefores["Value"][i]
    elif(x == "$TS"):
        TSBeforeStatus1 = CsdBefores["Status"][i]
        TSBeforeValues2 = CsdBefores["Value"][i]
    elif(x == "$TU"):
        TUBeforeStatus1 = CsdBefores["Status"][i]
        TUBeforeValues2 = CsdBefores["Value"][i]
    elif(x == "$TX"):
        TXBeforeStatus1 = CsdBefores["Status"][i]
        TXBeforeValues2 = CsdBefores["Value"][i]
    elif(x == "$TY"):
        WXBeforeStatus1 = CsdBefores["Status"][i]
        TYBeforeValues2 = CsdBefores["Value"][i]
    elif(x == "$TZ"):
        TZBeforeStatus1 = CsdBefores["Status"][i]
        TZBeforeValues2 = CsdBefores["Value"][i]
    elif(x == "$VB"):
        VBBeforeStatus1 = CsdBefores["Status"][i]
        VBBeforeValues2 = CsdBefores["Value"][i]
    elif(x == "$WD"):
        WDBeforeStatus1 = CsdBefores["Status"][i]
        WDBeforeValues2 = CsdBefores["Value"][i]
    elif(x == "$WR"):
        WRBeforeStatus1 = CsdBefores["Status"][i]
        WRBeforeValues2 = CsdBefores["Value"][i]
    elif(x == "$WS"):
        WSBeforeStatus1 = CsdBefores["Status"][i]
        WSBeforeValues2 = CsdBefores["Value"][i]
    elif(x == "$WX"):
        WXBeforeStatus1 = CsdBefores["Status"][i]
        WXBeforeValues2 = CsdBefores["Value"][i]
    i = i+1

j = 0

for x in typeAfters:

    if(x == "$AA"):
        AAAfterVales1 = CsdAfters["Status"][j]
        AAAfterVales2 = CsdAfters["Value"][j]
        AAAfterVales3 = CsdAfters["3"][j]
        AAAfterVales4_6 = "{}-{}-{}".format(
            CsdAfters["4"][j], CsdAfters["5"][j], CsdAfters["6"][j])
        AAAfterVales7_9 = "{}-{}-{}".format(
            CsdAfters["7"][j], CsdAfters["8"][j], CsdAfters["9"][j])
    if(x == "$HK"):
        HKAfterStatus1 = CsdAfters["Status"][j]
        HKAfterVales2 = CsdAfters["Value"][j]
        HKAfterVales4 = CsdAfters["4"][j]
        HKAfterVales6 = CsdAfters["6"][j]
    elif(x == "$BP"):
        BPAfterStatus1 = CsdAfters["Status"][j]
        BPAfterValues2 = CsdAfters["Value"][j]
        BPAfterValues3 = CsdAfters["3"][j]
        BPAfterValues5 = CsdAfters["5"][j]
    elif(x == "$BS"):
        BSAfterStatus1 = CsdAfters["Status"][j]
        BSAfterValues2 = CsdAfters["Value"][j]
    elif(x == "$DS"):
        DSAfterStatus1 = CsdAfters["Status"][j]
        DSAfterValues2 = CsdAfters["Value"][j]
        DSAfterValues3 = CsdAfters["3"][j]
        DSAfterValues4 = CsdAfters["4"][j]
    elif(x == "$FX"):
        FXAfterStatus1 = CsdAfters["Status"][j]
        FXAfterValues2_3_4 = "{}-{}-{}".format(
            CsdAfters["2"][j], CsdAfters["3"][j], CsdAfters["4"][j])
        FXAfterValues5_6_7 = "{}-{}-{}".format(
            CsdAfters["5"][j], CsdAfters["6"][j], CsdAfters["7"][j])
        FXAfterValues8 = CsdAfters["8"][j]
    elif(x == "$PR"):
        PRAfterStatus1 = CsdAfters["Status"][j]
        PRAfterValues2 = CsdAfters["Value"][j]
    elif(x == "$RC"):
        RCAfterStatus1 = CsdAfters["Status"][j]
        RCAfterValues2 = CsdAfters["Value"][j]
        RCAfterValues4 = CsdAfters["4"][j]
        RCAfterValues6 = CsdAfters["6"][j]
        RCAfterValues8 = CsdAfters["8"][j]
        RCAfterValues10 = CsdAfters["10"][j]
        RCAfterValues12 = CsdAfters["12"][j]
        RCAfterValues14 = CsdAfters["14"][j]
        RCAfterValues16 = CsdAfters["16"][j]
        RCAfterValues18 = CsdAfters["18"][j]
    elif(x == "$RH"):
        RHAfterStatus1 = CsdAfters["Status"][j]
        RHAfterValues2 = CsdAfters["Value"][j]
        RHAfterValues3 = ""

    elif(x == "$SD"):
        SDAfterStatus1 = CsdAfters["Status"][j]
        SDAfterValues2 = CsdAfters["Value"][j]
        SDAfterValues3 = CsdAfters["3"][j]
        SDAfterValues4 = CsdAfters["4"][j]
        SDAfterValues5 = CsdAfters["5"][j]
    elif(x == "$SM"):
        SMAfterStatus1 = CsdAfters["Status"][j]
        SMAfterValues2 = CsdAfters["Value"][j]
    elif(x == "$SR"):
        SRAfterStatus1 = CsdAfters["Status"][j]
        SRAfterValues2 = CsdAfters["Value"][j]
    elif(x == "$SS"):
        SSAfterStatus1 = CsdAfters["Status"][j]
        SSAfterValues3 = CsdAfters["3"][j]
        SSAfterValues4_5 = "{},{}".format(
            CsdAfters["4"][j], CsdAfters["5"][j])
        SSAfterValues6_7 = "{},{}".format(
            CsdAfters["6"][j], CsdAfters["7"][j])
        SSAfterValues8 = CsdAfters["8"][j]
        SSAfterValues10 = CsdAfters["10"][j]
        SSAfterValues11 = CsdAfters["11"][j]
        SSAfterValues12 = CsdAfters["12"][j]
        SSAfterValues13 = CsdAfters["13"][j]
        SSAfterValues14 = CsdAfters["14"][j]
        SSAfterValues15 = CsdAfters["15"][j]
    elif(x == "$T0"):
        T0AfterStatus1 = CsdAfters["Status"][j]
        T0AfterValues2 = CsdAfters["Value"][j]
    elif(x == "$T1"):
        T1AfterStatus1 = CsdAfters["Status"][j]
        T1AfterValues2 = CsdAfters["Value"][j]
    elif(x == "$T2"):
        T2AfterStatus1 = CsdAfters["Status"][j]
        T2AfterValues2 = CsdAfters["Value"][j]
    elif(x == "$T5"):
        T5AfterStatus1 = CsdAfters["Status"][j]
        T5AfterValues2 = CsdAfters["Value"][j]
    elif(x == "$TA"):
        TAAfterStatus1 = CsdAfters["Status"][j]
        TAAfterValues2 = CsdAfters["Value"][j]
    elif(x == "$TE"):
        TEAfterStatus1 = CsdAfters["Status"][j]
        TEAfterValues2 = CsdAfters["Value"][j]
    elif(x == "$TS"):
        TSAfterStatus1 = CsdAfters["Status"][j]
        TSAfterValues2 = CsdAfters["Value"][j]
    elif(x == "$TU"):
        TUAfterStatus1 = CsdAfters["Status"][j]
        TUAfterValues2 = CsdAfters["Value"][j]
    elif(x == "$TX"):
        TXAfterStatus1 = CsdAfters["Status"][j]
        TXAfterValues2 = CsdAfters["Value"][j]
    elif(x == "$TY"):
        WXAfterStatus1 = CsdAfters["Status"][j]
        TYAfterValues2 = CsdAfters["Value"][j]
    elif(x == "$TZ"):
        TZAfterStatus1 = CsdAfters["Status"][j]
        TZAfterValues2 = CsdAfters["Value"][j]
    elif(x == "$VB"):
        VBAfterStatus1 = CsdAfters["Status"][j]
        VBAfterValues2 = CsdAfters["Value"][j]
    elif(x == "$WD"):
        WDAfterStatus1 = CsdAfters["Status"][j]
        WDAfterValues2 = CsdAfters["Value"][j]
    elif(x == "$WR"):
        WRAfterStatus1 = CsdAfters["Status"][j]
        WRAfterValues2 = CsdAfters["Value"][j]
    elif(x == "$WS"):
        WSAfterStatus1 = CsdAfters["Status"][j]
        WSAfterValues2 = CsdAfters["Value"][j]
    elif(x == "$WX"):
        WXAfterStatus1 = CsdAfters["Status"][j]
        WXAfterValues2 = CsdAfters["Value"][j]
    j = j+1


Date = date.today()
TemplateFile = "TemplateFiles/{}/{}_{}.docx".format(System, System, doctype)

print(TemplateFile)
print(AABeforeVales1)
doc = DocxTemplate(TemplateFile)
#placeholder_1 = InlineImage(doc, "Placeholders/Placeholder_1.png", Cm(5))
#placeholder_2 = InlineImage(doc, "Placeholders/Placeholder_2.png", Cm(5))
context = {
    "name": Name,
    "site": Site,
    "date": Date,
    "data": BeforesCSD,
    # ///////////////////Before Values//////////////////////
    "AA_BV1": AABeforeVales1,
    "AA_BV2": AABeforeVales2,
    "AA_BV3": AABeforeVales3,
    "AA_BV4_6": AABeforeVales4_6,
    "AA_BV7_9": AABeforeVales7_9,

    "HK_BS1": HKBeforeStatus1,
    "HK_BV2": HKBeforeVales2,
    "HK_BV4": HKBeforeVales4,
    "HK_BV6": HKBeforeVales6,

    "BP_BS1": BPBeforeStatus1,
    "BP_BV2": BPBeforeValues2,
    "BP_BV3": BPBeforeValues3,
    "BP_BV5": BPBeforeValues5,

    "BS_BS1": BSBeforeStatus1,
    "BS_BV2": BSBeforeValues2,

    # DS
    "DS_BS1": DSBeforeStatus1,
    "DS_BV2": DSBeforeValues2,
    "DS_BV3": DSBeforeValues3,
    "DS_BV4": DSBeforeValues4,

    # FX
    "FX_BS1": FXBeforeStatus1,
    "FX_BV2_4": FXBeforeValues2_3_4,
    "FX_BV5_7": FXBeforeValues5_6_7,
    "FX_BV8": FXBeforeValues8,

    # Percipitation
    "PR_BS1": PRBeforeStatus1,
    "PR_BV2": PRBeforeValues2,

    # Road sensor
    "RC_BS1": RCBeforeStatus1,
    "RC_BV2": RCBeforeValues2,
    "RC_BV4": RCBeforeValues4,
    "RC_BV6": RCBeforeValues6,
    "RC_BV8": RCBeforeValues8,
    "RC_BV10": RCBeforeValues10,
    "RC_BV12": RCBeforeValues12,
    "RC_BV14": RCBeforeValues14,
    "RC_BV16": RCBeforeValues16,
    "RC_BV18": RCBeforeValues18,

    # Relitive Humidity
    "RH_BS1": RHBeforeStatus1,
    "RH_BV2": RHBeforeValues2,
    "RH_BV3": RHBeforeValues3,

    # Snow Depth
    "SD_BV2": SDBeforeStatus1,
    "SD_BV2": SDBeforeValues2,
    "SD_BV3": SDBeforeValues3,
    "SD_BV4": SDBeforeValues4,
    "SD_BV5": SDBeforeValues5,

    # Soil Mostiure
    "SM_BV2": SMBeforeValues2,

    # Solar Radiation
    "SR_BS1": SRBeforeStatus1,
    "SR_BV2": SRBeforeValues2,

    # Sea State
    "SS_BS1": SSBeforeStatus1,
    "SS_BV3": SSBeforeValues3,
    "SS_BV4_5": SSBeforeValues4_5,
    "SS_BV6_7": SSBeforeValues6_7,
    "SS_BV8": SSBeforeValues8,
    "SS_BV10": SSBeforeValues10,
    "SS_BV11": SSBeforeValues11,
    "SS_BV12": SSBeforeValues12,
    "SS_BV13": SSBeforeValues13,
    "SS_BV14": SSBeforeValues14,
    "SS_BV15": SSBeforeValues15,

    # Temp 100cm
    "T0_BS1": T0BeforeStatus1,
    "T0_BV2": T0BeforeValues2,

    # Temp 10cm
    "T1_BS1": T1BeforeStatus1,
    "T1_BV2": T1BeforeValues2,

    # Temp 20cm
    "T2_BS1": T2BeforeStatus1,
    "T2_BV2": T2BeforeValues2,

    # Temp 50cm
    "T5_BS1": T5BeforeStatus1,
    "T5_BV2": T5BeforeValues2,

    # Temp Air
    "TA_BS1": TABeforeStatus1,
    "TA_BV2": TABeforeValues2,

    # Temp Road
    "TR_BS1": TEBeforeStatus1,
    "BS_BV2": TEBeforeValues2,

    # Temp Surface
    "TS_BS1": TSBeforeStatus1,
    "TS_BV2": TSBeforeValues2,

    # Temp 5cm
    "T5_BS1": TUBeforeStatus1,
    "T5_BV2": TUBeforeValues2,

    # Temp X
    "TX_BS1": TXBeforeStatus1,
    "TX_BV2": TXBeforeValues2,

    # Temp Y
    "TY_BS1": TYBeforeStatus1,
    "TY_BV2": TYBeforeValues2,

    # Temp Z
    "TZ_BS1": TZBeforeStatus1,
    "TZ_BV2": TZBeforeValues2,

    # Visability
    "VB_BS1": VBBeforeStatus1,
    "VB_BV2": VBBeforeValues2,

    # Wind Direction
    "WD_BS1": WDBeforeStatus1,
    "WD_BV2": WDBeforeValues2,

    # Instant Wind Direction and Speed
    "WR_BS1": WRBeforeStatus1,
    "WR_BV2": WRBeforeValues2,
    "WR_BV3": WRBeforeValues3,

    # Wind Speed
    "WS_BS1": WSBeforeStatus1,
    "WS_BV2": WSBeforeValues2,

    # Present Wind
    "WX_BS1": WXBeforeStatus1,
    "WX_BV2": WXBeforeValues2,
    # /////////////////////////////////After Values///////////////

    "AA_AV1": AAAfterVales1,
    "AA_AV2": AAAfterVales2,
    "AA_AV3": AAAfterVales3,
    "AA_AV4_6": AAAfterVales4_6,
    "AA_AV7_9": AAAfterVales7_9,

    "HK_AS1": HKAfterStatus1,
    "HK_AV2": HKAfterVales2,
    "HK_AV4": HKAfterVales4,
    "HK_AV6": HKAfterVales6,

    "BP_AS1": BPAfterStatus1,
    "BP_AV2": BPAfterValues2,
    "BP_AV3": BPAfterValues3,
    "BP_AV5": BPAfterValues5,

    "BS_AS1": BSAfterStatus1,
    "BS_AV2": BSAfterValues2,

    # DS
    "DS_AS1": DSAfterStatus1,
    "DS_AV2": DSAfterValues2,
    "DS_AV3": DSAfterValues3,
    "DS_AV4": DSAfterValues4,

    # FX
    "FX_AS1": FXAfterStatus1,
    "FX_AV2_4": FXAfterValues2_3_4,
    "FX_AV5_7": FXAfterValues5_6_7,
    "FX_AV8": FXAfterValues8,

    # Percipitation
    "PR_AS1": PRAfterStatus1,
    "PR_AV2": PRAfterValues2,

    # Road sensor
    "RC_AS1": RCAfterStatus1,
    "RC_AV2": RCAfterValues2,
    "RC_AV4": RCAfterValues4,
    "RC_AV6": RCAfterValues6,
    "RC_AV8": RCAfterValues8,
    "RC_AV10": RCAfterValues10,
    "RC_AV12": RCAfterValues12,
    "RC_AV14": RCAfterValues14,
    "RC_AV16": RCAfterValues16,
    "RC_AV18": RCAfterValues18,

    # Relitive Humidity
    "RH_AS1": RHAfterStatus1,
    "RH_AV2": RHAfterValues2,
    "RH_AV3": RHAfterValues3,

    # Snow Depth
    "SD_AV2": SDAfterStatus1,
    "SD_AV2": SDAfterValues2,
    "SD_AV3": SDAfterValues3,
    "SD_AV4": SDAfterValues4,
    "SD_AV5": SDAfterValues5,

    # Soil Mostiure
    "SM_AV2": SMAfterValues2,

    # Solar Radiation
    "SR_AS1": SRAfterStatus1,
    "SR_AV2": SRAfterValues2,

    # Sea State
    "SS_AS1": SSAfterStatus1,
    "SS_AV3": SSAfterValues3,
    "SS_AV4_5": SSAfterValues4_5,
    "SS_AV6_7": SSAfterValues6_7,
    "SS_AV8": SSAfterValues8,
    "SS_AV10": SSAfterValues10,
    "SS_AV11": SSAfterValues11,
    "SS_AV12": SSAfterValues12,
    "SS_AV13": SSAfterValues13,
    "SS_AV14": SSAfterValues14,
    "SS_AV15": SSAfterValues15,

    # Temp 100cm
    "T0_AS1": T0AfterStatus1,
    "T0_AV2": T0AfterValues2,

    # Temp 10cm
    "T1_AS1": T1AfterStatus1,
    "T1_AV2": T1AfterValues2,

    # Temp 20cm
    "T2_AS1": T2AfterStatus1,
    "T2_AV2": T2AfterValues2,

    # Temp 50cm
    "T5_AS1": T5AfterStatus1,
    "T5_AV2": T5AfterValues2,

    # Temp Air
    "TA_AS1": TAAfterStatus1,
    "TA_AV2": TAAfterValues2,

    # Temp Road
    "TR_AS1": TEAfterStatus1,
    "AS_AV2": TEAfterValues2,

    # Temp Surface
    "TS_AS1": TSAfterStatus1,
    "TS_AV2": TSAfterValues2,

    # Temp 5cm
    "T5_AS1": TUAfterStatus1,
    "T5_AV2": TUAfterValues2,

    # Temp X
    "TX_AS1": TXAfterStatus1,
    "TX_AV2": TXAfterValues2,

    # Temp Y
    "TY_AS1": TYAfterStatus1,
    "TY_AV2": TYAfterValues2,

    # Temp Z
    "TZ_AS1": TZAfterStatus1,
    "TZ_AV2": TZAfterValues2,

    # Visability
    "VB_AS1": VBAfterStatus1,
    "VB_AV2": VBAfterValues2,

    # Wind Direction
    "WD_AS1": WDAfterStatus1,
    "WD_AV2": WDAfterValues2,

    # Instant Wind Direction and Speed
    "WR_AS1": WRAfterStatus1,
    "WR_AV2": WRAfterValues2,
    "WR_AV3": WRAfterValues3,

    # Wind Speed
    "WS_AS1": WSAfterStatus1,
    "WS_AV2": WSAfterValues2,

    # Present Wind
    "WX_AS1": WXAfterStatus1,
    "WX_AV2": WXAfterValues2


    #   "placeholder_1": placeholder_1,
    #   "placeholder_2": placeholder_2,
}

# Genterates file and saves it
FileName = "OutputFiles/{}_{}_{}_{}_{}.docx".format(
    System, doctype, Site, Name, Date)
print(FileName)
doc.render(context)
doc.save(FileName)
