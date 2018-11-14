import os
import easygui
from win32com.propsys import propsys, pscon
import shutil

def valik(): #easygui funktsioon
    msg ="Mida sa teha soovid?"
    title = "Failihaldur"
    choices = ["Sorteeri žanri järgi", "Failide ümbernimetamine", "Failide ümbernimetamine tunnuse järgi", "4"] #valik, mida kasutaja taha programmiga teha
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
            
def sorteerinimeta(kaust4, valik2):
    if valik2 == "Žanr":
        list1 = []
        for i in os.walk(kaust4):
            for j in i[2]:
                properties1 = propsys.SHGetPropertyStoreFromParsingName(kaust4+aa[0]+j)
                title1 = properties1.GetValue(pscon.PKEY_Music_Genre)
                list1.append([j,title1.GetValue()[0]])
        print(list1)
    elif valik2 == "Album":
        list1 = []
        for i in os.walk(kaust4):
            for j in i[2]:
                properties1 = propsys.SHGetPropertyStoreFromParsingName(kaust4+aa[0]+j)
                title1 = properties1.GetValue(pscon.PKEY_Music_AlbumTitle)
                list1.append([j,title1.GetValue()])
        print(list1)
    elif valik2 == "Faililõpp":
        list1 = []
        for i in os.walk(kaust4):
            for j in i[2]:
                properties1 = propsys.SHGetPropertyStoreFromParsingName(kaust4+aa[0]+j)
                title1 = properties1.GetValue(pscon.PKEY_FileName)
                list1.append([j,title1.GetValue().split(".")[::-1][0]])
        print(list1)
                
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

elif choice == "Failide ümbernimetamine tunnuse järgi":
    kaust4 = easygui.diropenbox(msg="Vali kaust, kus soovid ümbernimetamist tunnuse järgi läbi viia", title=None, default=None)
    valik2 = easygui.buttonbox("Millise tunnuse järgi soovid faili ümber nimetada?", "", ["Žanr","Album","Faililõpp"])
    sorteerinimeta(kaust4, valik2)
    

    
