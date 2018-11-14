import os
import easygui
from win32com.propsys import propsys, pscon
import shutil

def valik(): #easygui funktsioon
    msg ="Mida sa teha soovid?"
    title = "Failihaldur"
    choices = ["Sorteeri žanri järgi", "Failide ümbernimetamine", "3", "4"] #valik, mida kasutaja taha programmiga teha
    global choice
    choice = easygui.choicebox(msg, title, choices)

def sorteeri(kaust):
    properties = propsys.SHGetPropertyStoreFromParsingName(kaust) #file path
    title = properties.GetValue(pscon.PKEY_Music_Genre)
    try: #kui Genre on olemas
        if os.path.isdir(kaust1+aa[0]+title.GetValue()[0]) == True:
            shutil.copy((kaust1+aa[0]+j),(kaust1+aa[0]+title.GetValue()[0]))
        else:
            os.mkdir(kaust1+aa[0]+title.GetValue()[0]); shutil.copy((kaust1+aa[0]+j),(kaust1+aa[0]+title.GetValue()[0]))
    except TypeError:print("Faili "+kaust+" žanri ei õnnestunud leida")

def nimeta(kaust, positsioon, liide):
    temp_list = os.listdir(kaust3)
    for k in range(len(temp_list)):
        vana_nimi = temp_list[k]
        nime_list = vana_nimi.split('.')
        if var2 == 'Ette':
            os.rename(kaust3 + aa[0] + vana_nimi, kaust3 + aa[0] + liide + nime_list[0] + '.' + nime_list[1])
        else:
            os.rename(kaust3 + aa[0] + vana_nimi, kaust3 + aa[0] + nime_list[0] + liide + '.' + nime_list[1])

valik()
a = ["\\"]; aa=(''.join(a)) #escape cancer
if choice == "Sorteeri žanri järgi":
    kaust = easygui.diropenbox(msg="Vali kaust, kus soovid sorteerimist läbi viia", title=None, default=None) 
    kaust1 = kaust
    for i in os.walk(kaust):
        for j in i[2]:
            sorteeri(kaust+aa[0]+j) #teeb iga failiga läbi
            os.remove(kaust1+aa[0]+j)
elif choice == "Failide ümbernimetamine":
    kaust3 = easygui.diropenbox(msg="Vali kaust, kus soovid ümbernimetamist läbi viia", title=None, default=None)
    var2 = easygui.buttonbox("Kas soovid midagi lisada ette või taha?: ", "", ["Ette","Taha"])
    liide = easygui.enterbox('Mida soovid lisada failinimedele?')
    nimeta(kaust3, var2, liide)
