import os
import easygui
from win32com.propsys import propsys, pscon
import shutil
import tkinter as tk
from tkinter import ttk, filedialog
from functools import partial


def valik():
    msg = "Mida sa teha soovid?"
    title = "Failihaldur"
    choices = ["Sorteeri tunnuse järgi", "Kausta kõigi failide ümbernimetamine",
               "Failide ümbernimetamine tunnuse järgi", "4"]  # valik, mida kasutaja taha programmiga teha
    otsus = easygui.choicebox(msg, title, choices)
    return otsus


def sorteeri_tunnus(kaust, tunnus):
    for l in os.walk(kaust):
        for m in l[2]:
            fail_dir = kaust + aa[0] + m
            properties = propsys.SHGetPropertyStoreFromParsingName(fail_dir)  # file path
            title = properties.GetValue(votmed[tunnus])
            try:
                liide = tunnuse_saamine(title, tunnus)
                uus_dir = kaust + aa[0] + liide
                if not os.path.isdir(uus_dir):
                    os.mkdir(uus_dir)
                shutil.copy(fail_dir, uus_dir)
            except TypeError or NameError:
                print("Faili " + fail_dir + " žanri ei õnnestunud leida")


def nimeta_kaust(kaust, positsioon, liide):
    temp_list = os.listdir(kaust)
    for k in range(len(temp_list)):
        vana_nimi = temp_list[k]
        nime_list = vana_nimi.split('.')
        vana_dir = kaust + aa[0] + vana_nimi
        if positsioon == 'Algus':
            os.rename(vana_dir, kaust + aa[0] + liide + ' ' + nime_list[0] + '.' + nime_list[1])
        else:
            os.rename(vana_dir, kaust + aa[0] + nime_list[0] + ' ' + liide + '.' + nime_list[1])


def tunnuse_saamine(title, tunnus):
    if tunnus == "Žanr":
        liide = title.GetValue()[0]
    elif tunnus == "Album":
        liide = title.GetValue()
    elif tunnus == "Faililõpp":
        liide = title.GetValue().split(".")[::-1][0]
    return liide


def tunnusega_failid(kaust, tunnus):
    jarjend = []
    for i in os.walk(kaust):
        for j in i[2]:
            fail_dir = kaust + aa[0] + j
            properties = propsys.SHGetPropertyStoreFromParsingName(fail_dir)
            title = properties.GetValue(votmed[tunnus])
            try:
                liide = tunnuse_saamine(fail_dir, tunnus)
            except TypeError or NameError:
                continue
            if title.GetValue() is not None:
                jarjend.append([j, liide])
    return jarjend


def nimeta_tunnus(failid, kaust):
    for element in failid:
        nime_list = element[0].split('.')
        vana_dir = kaust + aa[0] + element[0]
        uus_dir = kaust + aa[0] + nime_list[0] + ' [' + element[1] + '].' + nime_list[1]
        os.rename(vana_dir, uus_dir)


aa = (''.join(["\\"]))  # escape cancer
votmed = {'Žanr': pscon.PKEY_Music_Genre, 'Album': pscon.PKEY_Music_AlbumTitle, 'Faililõpp': pscon.PKEY_FileName}

alg_kaust = easygui.diropenbox(msg="Vali kaust, millega tahad muutust läbi viia")
choice = valik()
if choice == "Sorteeri tunnuse järgi":
    # lisand = easygui.buttonbox("Millise tunnuse järgi soovid sorteerida?", "", ["Žanr", "Album", "Faililõpp"])
    # jätsin algse igaks juhuks alles
    base = tk.Tk()     # teeb akna
    base.title("Sorteeri tunnuse järgi")    # annnab aknale pealkirja
    main_frame = tk.Frame(base)    # tekitab raamistiku (window) akna jaoks
    main_frame.grid(column=0, row=0)    # jaotab window mentaalselt tükkideks
    ttk.Label(main_frame, text="Millise tunnuse järgi soovid sorteerida?").grid(column=1, row=0)    # paneb teksti kasti
    zanr_nupp = ttk.Button(main_frame, text='Žanr', command=partial(sorteeri_tunnus, alg_kaust, 'Žanr'))     # teeb nupu
    zanr_nupp.grid(row=2, column=2)     # paigutab nupu
    base.mainloop()     # laseb aknal ilmuda
    # sorteeri_tunnus(alg_kaust, lisand)
    # jätsin igaks juhuks alles
elif choice == "Kausta kõigi failide ümbernimetamine":
    lisand = easygui.enterbox('Mida soovid lisada failinimedele?')
    algus_lopp = easygui.buttonbox("Kas soovid seda lisada nime algusesse või lõppu?: ", "", ["Algus", "Lõpp"])
    nimeta_kaust(alg_kaust, algus_lopp, lisand)
elif choice == "Failide ümbernimetamine tunnuse järgi":
    lisand = easygui.buttonbox("Millise tunnuse soovid failinimele lisada?", "", ["Žanr", "Album", "Faililõpp"])
    nimeta_tunnus(tunnusega_failid(alg_kaust, lisand), alg_kaust)
