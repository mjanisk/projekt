import os
import easygui
from win32com.propsys import propsys, pscon
import shutil

def valik(): #easygui funktsioon
    
    msg ="Mida sa teha soovid?"
    title = "Failihaldur"
    choices = ["Sorteeri žanri järgi", "2", "3", "4"] #valik, mida kasutaja taha programmiga teha
    global choice
    choice = easygui.choicebox(msg, title, choices)
    

def sorteeri(kaust):
    properties = propsys.SHGetPropertyStoreFromParsingName(kaust) #file path
    title = properties.GetValue(pscon.PKEY_Music_Genre)

    try: #kui Genre on olemas
        if os.path.isdir(kaust1+aa[0]+title.GetValue()[0]) == True:shutil.copy((kaust1+aa[0]+j),(kaust1+aa[0]+title.GetValue()[0]))
        else:os.mkdir(kaust1+aa[0]+title.GetValue()[0]); shutil.copy((kaust1+aa[0]+j),(kaust1+aa[0]+title.GetValue()[0]))
    except TypeError:print("Faili "+kaust+" žanri ei õnnestunud leida")
    os.remove(kaust1+aa[0]+j)

valik()

a = ["\\"]; aa=(''.join(a)) #escape cancer
if choice == "Sorteeri žanri järgi":
    kaust = easygui.diropenbox(msg="Vali kaust, kus soovid sorteerimist läbi viia", title=None, default=None) 
    kaust1 = kaust
    for i in os.walk(kaust):
        for j in i[2]:
            sorteeri(kaust+aa[0]+j) #teeb iga failiga läbi
            
    
#easygui.multchoicebox(msg='', title=' ', choices=())











