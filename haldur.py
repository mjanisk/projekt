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
               "Failide ümbernimetamine tunnuse järgi"]
    for tunnus in choices:
        pass
    otsus = easygui.choicebox(msg, title, choices)
    return otsus


def sorteeri_tunnus(kaust, tunnus):
    for l in os.walk(kaust):
        for m in l[2]:
            fail_dir = kaust + aa[0] + m
            properties = propsys.SHGetPropertyStoreFromParsingName(fail_dir)
            title = properties.GetValue(votmed[tunnus])
            try:
                liide = tunnuse_saamine(title, tunnus)
                uus_dir = kaust + aa[0] + liide
                if not os.path.isdir(uus_dir):
                    os.mkdir(uus_dir)
                shutil.copy(fail_dir, uus_dir)
            except TypeError or NameError:
                print("Faili " + fail_dir + " žanri ei õnnestunud leida")
    main_base.destroy()


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
                liide = tunnuse_saamine(title, tunnus)
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


def kysi_kaust():
    nimi = filedialog.askdirectory()
    temp_list = nimi.split('/')
    parandatud = aa.join(temp_list)
    return parandatud


def kysi_failid():
    list1 = filedialog.askopenfilenames(initialdir="/", title='Vajuta CTRL, et valida mitu faili')
    return list1


def main_funk():    # jagan väiksemateks tükkideks pärast
    alg_kaust = kysi_kaust()
    if alg_kaust != '':
        choice = valik()
        global main_base
        if choice == "Sorteeri tunnuse järgi":
            main_base = tk.Tk()
            main_base.title("Sorteeri tunnuse järgi")
            main_frame = tk.Frame(main_base)
            main_frame.grid(column=0, row=0)
            ttk.Label(main_frame, text="Millise tunnuse järgi soovid sorteerida?").grid(column=1, row=0)
            tunnused = ['Žanr', 'Album', 'Faililõpp']
            for tunnus in tunnused:
                nupp = ttk.Button(main_frame, text=tunnus, command=partial(sorteeri_tunnus, alg_kaust, tunnus))
                nupp.grid(row=2, column=tunnused.index(tunnus))
            main_base.mainloop()
        elif choice == "Kausta kõigi failide ümbernimetamine":
            lisand = easygui.enterbox('Mida soovid lisada failinimedele?')
            algus_lopp = easygui.buttonbox("Kas soovid seda lisada nime algusesse või lõppu?: ", "", ["Algus", "Lõpp"])
            nimeta_kaust(alg_kaust, algus_lopp, lisand)
        elif choice == "Failide ümbernimetamine tunnuse järgi":
            lisand = easygui.buttonbox("Millise tunnuse soovid failinimele lisada?", "", ["Žanr", "Album", "Faililõpp"])
            nimeta_tunnus(tunnusega_failid(alg_kaust, lisand), alg_kaust)


aa = (''.join(["\\"]))  # escape cancer
votmed = {'Žanr': pscon.PKEY_Music_Genre, 'Album': pscon.PKEY_Music_AlbumTitle, 'Faililõpp': pscon.PKEY_FileName}

esiaken_base = tk.Tk()
esiaken_base.title('Failihaldur')
esiaken_frame = tk.Frame(esiaken_base)
esiaken_frame.grid(column=0, row=0)
ttk.Label(esiaken_frame, text="Tere!").grid(column=1, row=0)
ttk.Label(esiaken_frame, text="Alustamiseks vajuta Alusta, lõpetamiseks Lõpeta").grid(column=1, row=1)
alusta_nupp = ttk.Button(esiaken_frame, text="Alusta", command=main_funk)
alusta_nupp.grid(column=0, row=2)
lopeta_nupp = ttk.Button(esiaken_frame, text="Lõpeta", command=quit)
lopeta_nupp.grid(column=2, row=2)
esiaken_base.mainloop()
