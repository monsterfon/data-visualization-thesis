

# CMD:      pip install matplotlib ne dela
# pycharm---- file----settings----interpreter----matplotlib---install package-------DONE
import os
import matplotlib.pyplot as plt
import matplotlib.ticker
from matplotlib.ticker import LogFormatter
import matplotlib.backends.backend_svg
import string
import itertools
import re
import sys
from sys  import exc_info
import numpy as np
from datetime import date







Tabela_Frenkvenc_spremeni_po_potrebi_RI = [20, 80, 200, 1000, 6000]
Tabela_Frenkvenc_spremeni_po_potrebi_BCI = [10,250]
# ob želji dodajanja novig grafov je treba dodati nove funkcije
Izrisi_te_grafe_RI = ['forward', 'immunity', 'net']
Izrisi_te_grafe_BCI = ['forward','forward', 'immunity', 'net', 'sensor']
# Tabela podatkov, uporabljena za generiranje grafov
# result (YEAR_COLUMN, FREQUENCY_COLUMN, AXIS_COLUMN, PATH_COLUMN, TIME_COLUMN)
YEAR_COLUMN = 0
FREQUENCY_COLUMN = 1
AXIS_COLUMN = 2
PATH_COLUMN = 3
TIME_COLUMN = 4
IMMUNITY_COLUMN = 5
#datoteka v kateri so shranjene napake Error messages + warnings
napake = open("Napake EMC PRK.txt", "w+")






#fwd, net, imm, sensor
#fwd 0      3
#net 0     3-4
#imm 0      1
#sen 0      2




# VSE FUNKCIJE
def containsNumber(value):
    for character in value:
        if character.isdigit():
            return True


    return False

def has_result(fname):
    fname = fname + ('\Result Table.Result')
    #Vse kar ni test method =2 reference kalibration ignoriraj mapo
    try:
        with open(fname, 'r', encoding="utf-16") as f:
            return True
    except:
        return False

def is_valid_calibration(fname):
    fname = fname + "\\\\" + os.path.basename(fname) + ('.TestSetup')
    try:
        with open(fname, 'r', encoding="utf-16") as f:
            for line in f.readlines():
                # '[TestTime]\n'
                #MeasClass=2 EMS conducted
                #TestMethod=2  Reference Calibration
                if line.__contains__("TestMethod=2 Reference Calibration"):
                    return True;

        napake.writelines("\n   CHAIN ERROR. WRONG DATA. This is not a calibartion " )
        return False;


    except Exception as e:
        if containsNumber(os.path.basename(fname)):
            napake.writelines("\n   CHAIN VERY WEAK WARNING No  testSetup file in "+fname + str(e))
        return False;




def getAllSubpaths(path):  # razen tiste z številkami
    tabela_path = []
    for root, dirs, files in os.walk(path):
        for name in dirs:
            #has_result(path)
            if is_valid_calibration(    str (os.path.join(root, name))       ):
                tabela_path.append(os.path.join(root, name))
    return tabela_path




# funkcija ki  z regex izrazom  vzame datum iz string
# uporabi se jo v iteritools groupby
def extract_date(path: string):
    year = re.findall(r'\d+_\d+_\d+', path)[-1]
    # YEAR=re.sub(r'.', '', year, count = 1)
    return year


# funkcije ki vzame path datoteke in vrne arrey 2 X very long list ki  ga rabimo za izrisovanje grafov
# kakšne sorte list vrne je odvisno od imena funkcije
# ima tudi error cheacking mehanizem, ki pove katero datoteka ima problem prebrat

#        Immunity_level_file_read       content_array.append((float(line[0]), float(line[1])))
#Ant_fwd_file_read    content_array.append((float(line[0]), float(line[2])))
def graphDATA_file_read(fname,number0, number1, number2): #3 4 net  forward reverse
    content_array = []
    is_data_line = False
    fname = fname + ('\Result Table.Result')
    try:
        with open(fname, 'r', encoding="utf-16") as f:
            for line in f.readlines():
                if is_data_line:
                    #line = line.split()
                    line = re.split(r'\t+', line)
                    #line = re.split(r'\t+', line.rstrip('\t'))
                    if number2 != 0:
                        content_array.append((float(line[number0]), float(line[number1]) - float(line[number2])))
                    else:
                        content_array.append((float(line[number0]), float(line[number1]) ))
                else:
                    if line == "[TableValues]\n":
                        is_data_line = True
            return content_array
    except ValueError:
        input("CHAIN ERROR: pri branju  " + fname + "\n Klikni enter za naprej")
        print("Preveri nastavitve EMC32")
        print("Problem branja podatkov iz result table. Lahko  so ---  ---    ---    ---    namesto številk.")
        print(        napake.writelines("\n     CHAIN ERROR  net_file_read")
)
        napake.writelines("\n     CHAIN ERROR  net_file_read")
        napake.writelines("Problem branja podatkov iz result table. Lahko  so ---  ---    ---    ---    namesto številk.")
    except Exception as e:
        print(exc_info()[0])

        input("   CHAIN ERROR Net_pow_file_read" + str(e) + "\n Klikni enter za naprej")
        napake.writelines("\n     CHAIN ERROR  net_file_read" + str(e))
        napake.writelines("\n     ERROR: pri branju  Ant_fwd_file_read" + fname)
        napake.writelines("\n      Preveri nastavitve EMC32")


def Axis_file_read(fname):  # hor,ver
    content_array = []
    is_data_line = False
    fname = fname + ('\Result Table.Result')
    try:
        with open(fname, 'r', encoding="utf-16") as f:
            for line in f.readlines():
                if is_data_line:
                    line = line.split()
                    for word in line:
                        if word == 'V':
                            return 'ver'
                        if word == 'H':
                            return 'hor'
                elif line == "[TableValues]\n":
                    is_data_line = True

    except Exception as e:
        input("   CHAIN ERROR Axis_file_read" + str(e) + "\nKlikni enter za naprej")
        napake.writelines("\n     CHAIN ERROR  Axis_file_read" + str(e))

def Frequency_file_read(fname,number0):
    content_array = []
    is_data_line = False
    fname = fname + ('\Result Table.Result')
    try:
        with open(fname, 'r', encoding="utf-16") as f:
            for line in f.readlines():
                if is_data_line:
                    line = line.split()
                    bodoci_integer = float(line[number0])
                    content_array.append(int(bodoci_integer))
                elif line == "[TableValues]\n":
                    is_data_line = True
            return content_array
    except Exception as e:
        input("  CHAIN ERROR" + str(e) + "\n Klikni enter za naprej")
        napake.writelines("\n    CHAIN ERROR" + str(e))

def Time_file_read(fname):
    i = 0
    NEXT_LINE = False
    fname = fname + "\\\\" + os.path.basename(fname) + ('.TestSetup')
    try:
        with open(fname, 'r', encoding="utf-16") as f:
            for line in f.readlines():
                # '[TestTime]\n'
                if NEXT_LINE == True:
                    return line.replace('TestStart=', '');
                if line.__contains__("Time"):
                    NEXT_LINE = True
    except Exception as e:
        input("  CHAIN ERROR" + str(e) + "\n Klikni enter za naprej")
        napake.writelines("\n   CHAIN ERROR" + str(e))

def BCI_or_RI(fname):
    fname = fname + "\\\\" + os.path.basename(fname) + ('.TestSetup')
    try:
        with open(fname, 'r', encoding="utf-16") as f:
            for line in f.readlines():
                # '[TestTime]\n'



                if line.__contains__("EMS conducted"):
                    return "BCI";
                if line.__contains__("EMS radiated"):
                    return "RI";
    except Exception as e:
        input("  CHAIN ERROR" + str(e) + "\n Klikni enter za naprej")
        napake.writelines("\n   CHAIN ERROR" + str(e))


def EnotarAmper(fname):
    fname = fname + ('\Result Table.Result')
    try:
        with open(fname, 'r', encoding="utf-16") as f:
            for line in f.readlines():

                if line.__contains__("Unit"):
                    line = line.split()
                    return line
                    #for unit in line:



        napake.writelines("\n   CHAIN ERROR. WRONG DATA. This is not a calibartion " )
        return False

    except Exception as e:
        napake.writelines("\n   CHAIN  WARNING No  result file in "+fname + str(e))
        return False;


def IskanjeStolpcev_in_Enot(fname, searchword):
    fname = fname + ('\Result Table.Result')
    try:
        with open(fname, 'r', encoding="utf-16") as f:
            for line in f.readlines():

                if line.__contains__("Name"):
                    nameENOTE = re.split(r'\t+', line)
                    nameENOTE = re.split(r'\t+', line.rstrip('\t'))

                if line.__contains__("Unit"):
                    unit = re.split(r'\t+', line)
                    unit = re.split(r'\t+', line.rstrip('\t'))


                if line.__contains__("[TableValues]"):
                        break

        for index, enota in enumerate(nameENOTE):
               if enota.__contains__(searchword):
                        return nameENOTE[index], unit[index], index - 1

        napake.writelines("\n   CHAIN ERROR. WRONG DATA. Damaged result file" + fname)


    except Exception as e:
        napake.writelines("\n   CHAIN  WARNING No  result file in "+fname + str(e))






















# MAIN
print("Avtor: Žiga Fon©")
napake.writelines("V datoteki navodila uporabi CTRL+F in najdi besedo Nujno. Tako bos ugotovil kako urediti mape, da program deluje pravilno\n")



# preveri ali je uporabnik path napisal
while True:
    print("Shift + desni klik + copy as path")
    file_path = input("Napisi pot do glavne nadmape (Primerjave_referencnih_kalibracij):")
    file_name = os.path.basename(file_path)
    file_path = file_path.rstrip('\"').lstrip('\"')
    if file_path in ['n', 'N', 'No', 'NO', 'nO']:
        file_path = 'C:\\Users\\M0182965\\PycharmProjects\\EMC2.00\\Primerjave_referencnih_kalibracij'
        # quit
        break

        # če uporabnik ni vnesel path basename funkcija vrne kar string sam
        # če uporabnik ni vnesel path  prva črka ni enaka c
    if file_path in ['m', 'M']:
        file_path = 'C:\\Users\\M0182965\\PycharmProjects\\EMC2.00\\ISO11452-4'
        # quit
        break
    elif file_path == file_name:
        print("Napacno napisana pot do primerjav referencnih kalibracij.")
    else:
        print("Pravilna pot.")
        break


today = str(date.today())
newpath = os.getcwd()  + "\\Primerjava_Referenčnih_Kalibracij_Graf_" + today
pass
if not os.path.exists(newpath):
    os.makedirs(newpath)



while True:
    print("Ali želi odpraviti ver/hor napake avtomatično  ali ročno")
    automatic_or_not = input("A za automatično, R za ročno:")
    if automatic_or_not in ['A', 'a']:
        automatic_or_not = "automatic"
        break
    if automatic_or_not in ['R', 'r']:
        automatic_or_not = "manual"
        break

tabela_path = getAllSubpaths(file_path)


# zanka naredi (dictionary DATUM : vsi path iz DATUMA)
paths_by_date = {}
for key, group in itertools.groupby(tabela_path, extract_date):
    paths_by_date[key] = list(group)







# Ustvarjanje lista z (leto,frekvenčno_območje,hor/ver, path,year)
# če je automatic vzame ver/hor iz imena datoteke
# če je ročno pa se sam odločiš ali boš vzel iz imena datoteke ali iz result table
# frekvenco vedno prebere iz datoteke
SamoEnkrat = True
result123 = []
for key, value in paths_by_date.items():
    for val in value:
        *_, name = val.split('\\')
        # branje iz imena datoteke z regular expressions

        m = re.search(r"(?:ref[-_]cal)?(\d+M?-\d+MH?z?).*(hor|ver)", name, re.I)
        if m:
            freq_range, direction_ime = m.groups()
        else:
            # resitev warning Name 'direction_ime' can be undefined
            direction_ime = " "

        # branje iz datoteke
        try:
            direction_file = Axis_file_read(val)

            Freq_True_values = IskanjeStolpcev_in_Enot(val, "Frequency")
            index_od_frekvence = Freq_True_values[2]

            first = Frequency_file_read(val,index_od_frekvence).pop(0)
            last = Frequency_file_read(val,index_od_frekvence).pop(-1)
            freq_range = str(first) + "M-" + str(last) + "MHz"
            # freq_range izgleda priblizno  tako 80M-200MHz
            Exact_time = Time_file_read(val)
            imunost = BCI_or_RI(val)
            if imunost == "BCI":
                direction_file = "none"
                direction_ime= "none"
        except AttributeError:
            print("WARNING" + "freequency NoneType object has no attribute pop. \nProblem branja direction iz datoteke")
            print(value)
            napake.writelines("\nWARNING direction NoneType object has no attribute pop. \nProblem branja direction iz datoteke")
            napake.writelines("\nValue")
        except Exception as e:
            input("ERROR" + "\nKlikni enter za naprej"+ str(e) )
            napake.writelines("\nERROR" + str(e))

        try:
            # to se izvede samo v primeru problema
            if direction_ime.lower() != direction_file.lower():
                print("\nWARNING pri branju" + val)
                print(
                    "Ime datoteke pravi,  da je smo merili    " + direction_ime.upper() + "    v result table pa pise   " + direction_file.upper() + "\n")
                napake.writelines("\n\n WARNING pri branju" + val)
                napake.writelines("    Ime datoteke pravi,  da smo merili    " + direction_ime.upper() + "    v result table pa pise   " + direction_file.upper() + "\n")

                if automatic_or_not == "automatic":
                    # automatično  da tisto od imena datoteke
                    if direction_ime == None:
                        direction_ime = direction_file
                    if direction_file == None:
                        direction_file = direction_ime
                    result123.append((key, freq_range, direction_ime, val, Exact_time,imunost))
                elif automatic_or_not == "manual":
                    while True:

                        True_Axis = input("Horizontal ali vertical? (H or V):")
                        if True_Axis == 'H' or True_Axis == 'h':
                            result123.append((key, freq_range, 'hor', val, Exact_time,imunost))
                            break
                        if True_Axis == 'V' or True_Axis == 'v':
                            result123.append((key, freq_range, 'ver', val, Exact_time,imunost))
                            break

            else:
                # Ko ni problema:  v tem primeru sta oba direction ista
                result123.append((key, freq_range, direction_ime, val, Exact_time,imunost))

        except AttributeError:
            print("WARNING NoneType object direction has no attribute lower. Problem branja direction iz datoteke")
            print(value)
            napake.writelines("\nWARNING NoneType object direction has no attribute lower. \nProblem branja direction iz datoteke")
        except Exception as e:
            input(e + "\n Klikni enter za naprej")
            napake.writelines("\nERROR" + str(e))
# result123 je duplicated
# result je filtriran
ObstojeciDatumi = list()
# Remove Duplicates from a Python list using a For Loop
result = []
for index, row in enumerate(result123):

    if row[TIME_COLUMN] not in ObstojeciDatumi:
        result.append(result123[index])
        ObstojeciDatumi.append(result123[index][TIME_COLUMN])
    else:
        print()

[print(i) for i in result]
vrstica_prej = result[0][IMMUNITY_COLUMN]
for vrstica in result:

    if vrstica[IMMUNITY_COLUMN] is not vrstica_prej:
        print("ERROR: POSKUŠAŠ HKRATI BCI IN RI DELAT!!! TEGA SE NE SME!!! VSE KAR TI BO PROGRAM DAL VN BO NEUPORABNO!!! ")

    vrstica_prej = vrstica[IMMUNITY_COLUMN]




























katerikoli_ker_so_vsi_enaki = 0
rainbow = ['b', 'r', 'g', 'k', 'dodgerblue', 'orange', 'm', 'lime', 'saddlebrown', 'c']

if result[katerikoli_ker_so_vsi_enaki][IMMUNITY_COLUMN] == "RI":
    #immunity == conducted immunity, BCI, bulk current injection
    """Naredit da če je r naj naredi tko kot zdej. Če je b naj skippa na drugo for zanko , ki je b
    Brez hor ver, ime vsebuje bci namesto ri, zriše naj še graf sensor level, 
   bere naj iz drugih funkcij, izrisi te grafe
     Izrisi_te_grafe = ['forward', 'immunity', 'net', 'sensor']"""
    #ŠE ENO ZANKO RAJŠI


    #immunity == Radiated immunity
    # 4 for zanke
    # primerjamo (graf_type,zacetna_frekvenca, hor ver, za vsako vrstico)


    # indeksi se za razliko od COLUMN spreminjajop pri vrtenju for zanke
    indeks_barve = 0
    letni_indeks = 0
    frekvencni_indeks = 0
    polarizacija_antene_indeks = 0
    pot_indeks = 0
    once = True
    zadnji_indeks = 0

    #naj ignorira ce ni ri
    for graf in Izrisi_te_grafe_RI:
        for zacetna_frekvenca in Tabela_Frenkvenc_spremeni_po_potrebi_RI:
            frekvencni_indeks = 0
            # iskanje elementov v tabeli
            # narejeno na zelo časovno potraten način, a na srečo čas ni faktor
            while True:
                try:
                    if int(re.compile(r'\d+(?:\.\d+)?').findall(result[frekvencni_indeks][FREQUENCY_COLUMN])[
                           0]) != zacetna_frekvenca:
                        frekvencni_indeks += 1
                    else:
                        break
                    #če manjka graf ene frekvence
                    #će manjka graf ene  v tabeli frekvenc
                except IndexError as e:
                    frekvencni_indeks = 0
                    print("WARNING: Could not find frequency:" +  str(zacetna_frekvenca)  + "       Radiated immunity graf maker"  )
                    print(e)
                    napake.writelines("\n       WARNING: Could not find frequency:" +  str(zacetna_frekvenca)  + "       Radiated immunity graf maker")
                    napake.writelines(str(e))
                    break
                except Exception as e:
                    frekvencni_indeks = 0
                    input(e + "\n Klikni enter za naprej")
                    print(exc_info()[0])



            polarizacija_antene_indeks = 0
            for n in range(2):
                # leto je enako prvemu elementu v tabeli. solves undefined issue


                # shranimo in ustvarimo figure in subfigure
                fig = plt.figure(1)
                ax = fig.add_subplot(111)
                for vrstica in range(len(result)):
                    # najprej je search value, potle je semi fixed value
                    result[vrstica]
                    if int(re.compile(r'\d+(?:\.\d+)?').findall(result[vrstica][FREQUENCY_COLUMN])[0]) == int(
                            re.compile(r'\d+(?:\.\d+)?').findall(result[frekvencni_indeks][FREQUENCY_COLUMN])[0]):
                        if result[vrstica][AXIS_COLUMN].lower() == result[polarizacija_antene_indeks][AXIS_COLUMN].lower():
                            leto = result[vrstica][YEAR_COLUMN]

                            if result[vrstica][IMMUNITY_COLUMN] == "RI":


                                Freq_True_values =  IskanjeStolpcev_in_Enot(result[vrstica][PATH_COLUMN], "Frequency")
                                Fwd_True_values = IskanjeStolpcev_in_Enot(result[vrstica][PATH_COLUMN], "Fwd")
                                Rev_True_values = IskanjeStolpcev_in_Enot(result[vrstica][PATH_COLUMN], "Rev")
                                Imm_True_values = IskanjeStolpcev_in_Enot(result[vrstica][PATH_COLUMN], "Imm")

                                # list = (name, unit, index)


                                if graf == 'forward':
                                    data1 =  graphDATA_file_read(    result[vrstica][PATH_COLUMN],  Freq_True_values[2]    , Fwd_True_values[2], 0)

                                if graf == 'immunity':
                                    data1 =  graphDATA_file_read(    result[vrstica][PATH_COLUMN],  Freq_True_values[2]    , Imm_True_values[2], 0)
                                if graf == 'net':
                                    data1 = graphDATA_file_read(result[vrstica][PATH_COLUMN], Freq_True_values[2], Fwd_True_values[2], Rev_True_values[2])
                                try:
                                    x_val = [x[0] for x in data1]
                                    y_val = [x[1] for x in data1]
                                except Exception as e:
                                    print("\n Klikni enter za naprej  ")
                                    input(e)
                                    print(exc_info()[0])
                                    napake.writelines("\nERROR" + str(e))
                                plt.semilogx(x_val, y_val, color=rainbow[indeks_barve], label=leto, linewidth=0.4)
                                # če je letos enako leto past 30, potem čekiraj ali imasta datoteki enak čas in isto frekvenco in hor ver, potem sta enaki

                                indeks_barve += 1
                                if indeks_barve >= 9:
                                    indeks_barve = 0
                                zadnji_indeks = vrstica
                indeks_barve = 0
                # konec in zacetek risanja grafa mora biti temeljito definiran
                plt.xlim(xmin=int(re.compile(r'\d+(?:\.\d+)?').findall(result[zadnji_indeks][FREQUENCY_COLUMN])[0]),
                         xmax=int(re.compile(r'\d+(?:\.\d+)?').findall(result[zadnji_indeks][FREQUENCY_COLUMN])[1]))
                # plt.ylim(ymin=10, ymax=80)
                # grid on
                plt.grid(visible=True, which='major', color='dimgray', linestyle='-')
                plt.grid(visible=True, which='minor', color='grey', linestyle='-', alpha=0.2)
                plt.minorticks_on()
                # zbrise 10^3
                plt.gca().xaxis.set_major_formatter(matplotlib.ticker.ScalarFormatter())
                plt.gca().xaxis.set_minor_formatter(matplotlib.ticker.ScalarFormatter())

                #ODSTRANI IF CE EN DAN HORIZONTALA NE BO MANJKALA PRI 20
                if result[polarizacija_antene_indeks][AXIS_COLUMN].lower() == 'hor' and zacetna_frekvenca == 20:
                    # če horizontala ne manjka zbrisi ta if else in pusti samo plt.legend()
                    lgd = None
                    Freq_True_values = IskanjeStolpcev_in_Enot(result[polarizacija_antene_indeks][PATH_COLUMN], "Frequency")
                    if once == True:
                        print("Manjka horizontala 20-80MHz")
                    once = False
                    Fwd_True_values = IskanjeStolpcev_in_Enot(result[vrstica][PATH_COLUMN], "Fwd")
                    Rev_True_values = IskanjeStolpcev_in_Enot(result[vrstica][PATH_COLUMN], "Rev")
                    Imm_True_values = IskanjeStolpcev_in_Enot(result[vrstica][PATH_COLUMN], "Imm")

                else:
                    lgd = plt.legend(bbox_to_anchor=(1.05, 1), loc='upper left')

                # TUKI DA BERE VSE IZ NET
                # IskanjeStolpcev_in_Enot(result[zadnji_indeks][PATH_COLUMN], "Freq")  # name, unit, index

                plt.xlabel('{} in {}'.format(   Freq_True_values[0], Freq_True_values[1]    ))
                if graf == 'forward':
                    plt.ylabel('{} in {}'.format(   Fwd_True_values[0], Fwd_True_values[1]    ))
                if graf == 'immunity':
                    plt.ylabel('{} in {}'.format(   Imm_True_values[0], Imm_True_values[1]    ))
                if graf == 'net':
                    plt.ylabel('Net Power in {}'.format(Fwd_True_values[1])     )
                i_old = ""
                grafic = ""
                for vrsta in result:
                    # prevrti tabelo do indeks katerega leta iščemo
                    year = vrsta[YEAR_COLUMN]
                    if year != i_old:
                        grafic += year + "..."
                    i_old = year
                # naslov narisat
                plt.title(result[zadnji_indeks][FREQUENCY_COLUMN] + "_" + result[zadnji_indeks][AXIS_COLUMN])

                title_name = newpath +"\\"+ graf + "..." + grafic + result[zadnji_indeks][FREQUENCY_COLUMN] + "_" + result[zadnji_indeks][
                    AXIS_COLUMN] + "_"  + result[zadnji_indeks][IMMUNITY_COLUMN]+ '.svg'
                if lgd is not None:
                    plt.savefig(title_name, bbox_extra_artists=(lgd,), bbox_inches='tight')

                plt.close()

                # toggle hor -> ver   or   ver -> hor
                if result[polarizacija_antene_indeks][AXIS_COLUMN].lower() == 'hor':
                    while result[polarizacija_antene_indeks][AXIS_COLUMN].lower() == 'hor':
                        polarizacija_antene_indeks += 1
                        # print(result[polarizacija_antene_indeks][AXIS_COLUMN].lower())
                else:
                    while result[polarizacija_antene_indeks][AXIS_COLUMN].lower() == 'ver':
                        polarizacija_antene_indeks += 1
                        # print(result[polarizacija_antene_indeks][AXIS_COLUMN].lower())

















if result[katerikoli_ker_so_vsi_enaki][IMMUNITY_COLUMN] == "BCI":
    #conducted immunity. Bci.
    indeks_barve = 0
    letni_indeks = 0
    frekvencni_indeks = 0

    pot_indeks = 0
    once = True
    zadnji_indeks = 0
    brejkic = False
    #naj ignorira ce ni ri
    for graf in Izrisi_te_grafe_BCI:

        for zacetna_frekvenca in Tabela_Frenkvenc_spremeni_po_potrebi_BCI:
            frekvencni_indeks = 0
            # iskanje elementov v tabeli
            # narejeno na zelo časovno potraten način, a na srečo čas ni faktor
            while True:
                try:
                    if int(re.compile(r'\d+(?:\.\d+)?').findall(result[frekvencni_indeks][FREQUENCY_COLUMN])[
                           0]) != zacetna_frekvenca:
                        #išči dokler ne najdeš

                        if frekvencni_indeks  >=  len(result):
                            break
                        frekvencni_indeks += 1
                        pass

                    else:
                        break
                    #če manjka graf ene frekvence
                    #će manjka graf ene  v tabeli frekvenc
                except IndexError as e:
                    frekvencni_indeks = 0
                    print("WARNING: Could not find frequency:" +  str(zacetna_frekvenca)  + "       Radiated immunity graf maker"  )
                    print(e)
                    napake.writelines("\n       WARNING: Could not find frequency:" +  str(zacetna_frekvenca)  + "       Radiated immunity graf maker")
                    napake.writelines(str(e))
                    break
                except Exception as e:
                    frekvencni_indeks = 0
                    input(e + "\n Klikni enter za naprej")
                    print(exc_info()[0])
            #if brejkic == True:
            #    break


                # shranimo in ustvarimo figure in subfigure
                fig = plt.figure(1)
                ax = fig.add_subplot(111)
                for vrstica in range(len(result)):

                    # najprej je search value, potle je semi fixed value

                    if int(re.compile(r'\d+(?:\.\d+)?').findall(result[vrstica][FREQUENCY_COLUMN])[0]) == int(
                            re.compile(r'\d+(?:\.\d+)?').findall(result[frekvencni_indeks][FREQUENCY_COLUMN])[0]):
                            if result[vrstica][IMMUNITY_COLUMN] == "BCI":
                                leto = result[vrstica][YEAR_COLUMN]
                                # TUKI DA BERE VSE IZ NET
                                # IskanjeStolpcev_in_Enot(result[zadnji_indeks][PATH_COLUMN], "Freq")  # name, unit, index

                                Freq_True_values = IskanjeStolpcev_in_Enot(result[vrstica][PATH_COLUMN], "Frequency")
                                Fwd_True_values = IskanjeStolpcev_in_Enot(result[vrstica][PATH_COLUMN], "Fwd")
                                Rev_True_values = IskanjeStolpcev_in_Enot(result[vrstica][PATH_COLUMN], "Rev")
                                Imm_True_values = IskanjeStolpcev_in_Enot(result[vrstica][PATH_COLUMN], "Imm")
                                Sensor_True_values = IskanjeStolpcev_in_Enot(result[vrstica][PATH_COLUMN], "Sensor")
                                try:

                                    if graf == 'forward':
                                        data1 =  graphDATA_file_read(    result[vrstica][PATH_COLUMN],  Freq_True_values[2]    , Fwd_True_values[2], 0)
                                    if graf == 'immunity':
                                        data1 =  graphDATA_file_read(    result[vrstica][PATH_COLUMN],  Freq_True_values[2]    , Imm_True_values[2], 0)
                                    if graf == 'net':
                                        data1 = graphDATA_file_read(result[vrstica][PATH_COLUMN], Freq_True_values[2], Fwd_True_values[2], Rev_True_values[2])
                                    if graf == 'sensor':
                                        data1 = graphDATA_file_read(    result[vrstica][PATH_COLUMN],  Freq_True_values[2]    , Sensor_True_values[2], 0)


                                    x_val = [x[0] for x in data1]
                                    y_val = [x[1] for x in data1]
                                except Exception as e:
                                    input(e  )
                                    print("\n Klikni enter za naprej")
                                    print(exc_info()[0])
                                    napake.writelines(str(e))
                                    napake.writelines("\n Klikni enter za naprej")
                                plt.semilogx(x_val, y_val, color=rainbow[indeks_barve], label=leto, linewidth=0.4)
                                # če je letos enako leto past 30, potem čekiraj ali imasta datoteki enak čas in isto frekvenco in hor ver, potem sta enaki

                                indeks_barve += 1
                                if indeks_barve >= 9:
                                    indeks_barve = 0
                                zadnji_indeks = vrstica
                indeks_barve = 0
                # konec in zacetek risanja grafa mora biti temeljito definiran
                x_start = int(re.compile(r'\d+(?:\.\d+)?').findall(result[zadnji_indeks][FREQUENCY_COLUMN])[0])
                x_stop = int(re.compile(r'\d+(?:\.\d+)?').findall(result[zadnji_indeks][FREQUENCY_COLUMN])[1])
                plt.xlim(xmin=x_start, xmax=x_stop)
                # TUKI DA BERE VSE IZ NET
                # IskanjeStolpcev_in_Enot(result[zadnji_indeks][PATH_COLUMN], "Freq")  # name, unit, index
                if x_start == Tabela_Frenkvenc_spremeni_po_potrebi_BCI[0]: #why not true
                    # plt.xticks(np.arange(zacetna_frekvenca, 250000, 100000))
                    plt.minorticks_off()
                    ax.set_xticks([x_start,100, 1000,10000,50000, x_stop])
                    #na logaritmske četrtine
                    plt.title("10k - 250 000 kHz")
                    plt.xlabel('Frequency in kHz')
                else:
                    plt.minorticks_on()
                    ax.set_xticks([x_start,300,350,x_stop])
                    plt.title(result[zadnji_indeks][FREQUENCY_COLUMN])
                    plt.xlabel('Frequency in MHz')


                # plt.ylim(ymin=10, ymax=80)
                # grid on
                plt.grid(visible=True, which='major', color='dimgray', linestyle='-')
                plt.grid(visible=True, which='minor', color='grey', linestyle='-', alpha=0.2)

                # zbrise 10^3
                plt.gca().xaxis.set_major_formatter(matplotlib.ticker.ScalarFormatter())
                plt.gca().xaxis.set_minor_formatter(matplotlib.ticker.ScalarFormatter())
                #da legendo vn
                lgd = plt.legend(bbox_to_anchor=(1.05, 1), loc='upper left')

                enote_list = EnotarAmper(result[zadnji_indeks][PATH_COLUMN])
                enota_forward = enote_list[4]
                enota_immunity = enote_list[2]
                enota_sensor = enote_list[3]
                enota_net = enote_list[4] #isto
                # TUKI DA BERE VSE IZ NET
                # IskanjeStolpcev_in_Enot(result[zadnji_indeks][PATH_COLUMN], "Freq")  # name, unit, indexn

                plt.xlabel('{} in {}'.format(Freq_True_values[0], Freq_True_values[1]))
                if graf == 'forward':
                    plt.ylabel('{} in {}'.format(Fwd_True_values[0], Fwd_True_values[1]))
                if graf == 'immunity':
                    plt.ylabel('{} in {}'.format(Imm_True_values[0], Imm_True_values[1]))
                if graf == 'net':
                    plt.ylabel('Net Power in {}'.format(Fwd_True_values[1] ))
                if graf == 'sensor':
                    plt.ylabel('{} in {}'.format(Sensor_True_values[0], Sensor_True_values[1]))

                i_old = ""
                grafic = ""
                for vrsta in result:
                    # prevrti tabelo do indeks katerega leta iščemo
                    year = vrsta[YEAR_COLUMN]
                    if year != i_old:
                        grafic += year + "..."
                    i_old = year
                # naslov narisat


                title_name = newpath +"\\"+ graf + "..." + grafic + result[zadnji_indeks][FREQUENCY_COLUMN] + "_" + result[zadnji_indeks][AXIS_COLUMN] + "_" + result[zadnji_indeks][IMMUNITY_COLUMN] + '.svg'
                if lgd is not None:
                    plt.savefig(title_name, bbox_extra_artists=(lgd,), bbox_inches='tight')


                plt.close()









napake.close()
print("Obdelava zaključena")
x_c = input("Potrdi za zaključek")
# https://coderslegacy.com/python/auto-py-to-exe-tutorial/


