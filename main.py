import xlrd
from prettytable import PrettyTable

def check_version(classeur):
    global version
    feuille = classeur.sheet_by_name("AMP1")
    f_name = feuille.cell_value(13,6)
    if f_name == "FLUENCE POMPE / DOMMAGE":
        version = 1
        return True
    if f_name == "FLUENCE POMP @ 45°":
        version = 2
        return True
    else:
        return False

def create_liste_amp(nombre_amp):
    liste_amp = []
    for i in range(1, nombre_amp + 1 + 1): #+1 pour récupérer l'énergie de sortie dans l'entrée de l'AMP N+1)
        liste_amp.append("AMP" + str(i))
    return liste_amp

def create_liste_amp_tableau(nombre_amp):
    liste_amp_tableau = ["Entrée"]
    for i in range(1, nombre_amp + 1):
        liste_amp_tableau.append("OUT AMP" + str(i))
    liste_amp_tableau.append("OUT Atténuateur")
    liste_amp_tableau.append("OUT Compresseur")
    liste_amp_tableau.append("Puissance système")
    return liste_amp_tableau

def create_liste_energie_tableau(liste_energie):
    global a_pp, a_pa, c_pp, c_pa
    liste_energie_tableau = []
    for elt in liste_energie:
        chaine_formatee = '{:,}'.format(elt).replace(',', ' ')
        liste_energie_tableau.append(chaine_formatee + " mJ")
    #attenuateur
    chaine_formatee = '{:,}'.format(round(liste_energie[len(liste_energie) - 1]*(1-a_pp)*(1-a_pa),2)).replace(',', ' ')
    liste_energie_tableau.append(chaine_formatee + " mJ")
    #compresseur
    chaine_formatee = '{:,}'.format(round(liste_energie[len(liste_energie) - 1]*(1-a_pp)*(1-a_pa)*(1-c_pp)*(1-c_pa),2)).replace(',', ' ')
    liste_energie_tableau.append(chaine_formatee + " mJ")
    #puissance
    p = round(liste_energie[len(liste_energie) - 1]*(1-a_pp)*(1-a_pa)*(1-c_pp)*(1-c_pa)/25,2)
    chaine_formatee = '{:,}'.format(p).replace(',', ' ')
    liste_energie_tableau.append(chaine_formatee + " TW")
    
    return liste_energie_tableau
    
def energy_scraping(classeur, liste_amp):
    E = []
    for elt in liste_amp:
        feuille = classeur.sheet_by_name(elt)
        E.append(round(feuille.cell_value(2,4),2)) #0A
    return E

def create_liste_power_tableau():
    return ["Objectif", "Puissance système", "Marge"]

def create_liste_power_val_tableau(liste_energie):
    global objectif
    p = round(liste_energie[len(liste_energie) - 1]*(1-a_pp)*(1-a_pa)*(1-c_pp)*(1-c_pa)/25,2)
    L = []

    chaine_formatee = '{:,}'.format(round(objectif,2)).replace(',', ' ')
    L.append(chaine_formatee + " TW")

    chaine_formatee = '{:,}'.format(round(p,2)).replace(',', ' ')
    L.append(chaine_formatee + " TW")

    chaine_formatee = '{:,}'.format(round((p-objectif)/objectif*100,2)).replace(',', ' ')
    L.append(chaine_formatee + " %")

    
    return L

def fluence_pompe_dommage(classeur, liste_amp):
    global debug, version

    if version == 1:
        for elt in liste_amp:
            feuille = classeur.sheet_by_name(elt)
            fluence_p_d = feuille.cell_value(13,9)
            if fluence_p_d > 1 and elt !="AMP1":
                debug = True
                fluence = feuille.cell_value(12,9)
                seuil = feuille.cell_value(21,4)
                print("\n[Debug - " + elt + "] Le rapport Fluence de pompe/Dommage est trop élevé (>100 %) !\n/!\ Fluence de pompe/Dommage : " + str(round(fluence_p_d,2)*100) + " % \n(Info. comp.) Fluence de pompe : " + str(round(fluence,2)) + " (conseillé [1.4-1.6])\n(Info. comp.) Seuil : " + str(seuil))
                
    if version == 2:
        for elt in liste_amp:
            feuille = classeur.sheet_by_name(elt)
            fluence = feuille.cell_value(12,9)
            seuil = feuille.cell_value(21,4)
            fluence_p_d = round(fluence/seuil,2)
            if fluence_p_d > 1 and elt !="AMP1":
                debug = True
                print("\n[Debug - " + elt + "] Le rapport Fluence de pompe/Dommage est trop élevé (>100 %) !\n/!\ Fluence de pompe/Dommage : " + str(round(fluence_p_d,2)*100) + " % \n(Info. comp.) Fluence de pompe : " + str(round(fluence,2)) + " (conseillé [1.4-1.6])\n(Info. comp.) Seuil : " + str(seuil))
   
def fluence_sortie(classeur, liste_amp):
    global debug
    for elt in liste_amp :
        feuille = classeur.sheet_by_name(elt)
        passage = int(feuille.cell_value(18,4))
        fluence_s = feuille.cell_value(9,13 + passage)
        if fluence_s < 1.1 and elt !="AMP1":
            debug = True
            print("\n[Debug - " + elt + "] La fluence de sortie est faible, l'amplificateur n'est peut-être pas assez exploité (<1.1) !\n/!\ Fluence de sortie : " + str(round(fluence_s,2)) + "\n(Conseil) Fluence de sortie [1.1 ; 1.5]")
        if fluence_s > 1.5 and elt !="AMP1" : 
            debug = True
            print("\n[Debug - " + elt + "] La fluence de sortie est trop élevé (>1.5) !\n/!\ Fluence de sortie : " + str(round(fluence_s,2)))

def eclairement(classeur, liste_amp):
    global debug
    for elt in liste_amp :
        feuille = classeur.sheet_by_name(elt)
        passage = int(feuille.cell_value(18,4))
        eclairement = feuille.cell_value(10,13 + passage)
        if eclairement > 5000 :
            debug = True
            print("\n[Debug - " + elt + "] L'éclairement est trop élevé (>5 GW/cm²) !\n/!\ Eclairement : " + str(round(eclairement,2)) + " MW/cm²\n(Conseil) Augmenter la durée d'étirement 7 ps -> 14 ps")

def largeur_spectrale(classeur, liste_energie):
    global debug
    feuille = classeur.sheet_by_name("AMP1")
    l_s = feuille.cell_value(11,4)
    p = round(liste_energie[len(liste_energie) - 1]*(1-a_pp)*(1-a_pa)*(1-c_pp)*(1-c_pa)/25,2)
    elt = "AMP1"
    if p > 1000:
        if l_s < 60 or l_s > 65:
            debug = True
            print("\n[Debug - " + elt + "] La largeur spectrale n'est pas bonne !\n/!\ Largeur spectrale : " + str(round(l_s,2)) + " nm\n(Conseil) La puissance du système dépasse 1 PW : [60 ; 65]")
    """
    if p < 1000:
        if l_s < 50 or l_s > 55:
            debug = True
            print("\n[Debug - " + elt + "] La largeur spectrale n'est pas bonne !\n/!\ Largeur spectrale : " + str(round(l_s,2)) + " nm\n(Conseil) La puissance du système est de l'ordre du TW : [50 ; 55]")
    """
#CONFIG
a_pp = 0.1 
a_pa = 0.01
c_pp = 1-0.701
c_pa = 0
objectif = 350

version = 0
liste_amp = []
liste_sortie = []
debug = False
find = True
cont = True

#C:\Users\T0280520\Desktop\Travaux d'introduction\Simulations Ampli + Bilan\Copie de Simu Amplification 200 TW.xls
#L:\TL_COMMUNS\40-Pilotage Offres\05-Repertoire_Offres\ITALY\QUARK 45TW_2019_ITALY_DSL_0778_0_1_INFN CATANE\60-Solution\Catane - AC - BILAN - Simulation 350 TW.xls
# Ouvrir le fichier Excel
path_to_excel = r"L:\TL_COMMUNS\40-Pilotage Offres\05-Repertoire_Offres\ITALY\QUARK 45TW_2019_ITALY_DSL_0778_0_1_INFN CATANE\60-Solution\Catane - AC - BILAN - Simulation 350 TW.xls"

try:
    classeur = xlrd.open_workbook(path_to_excel)
except:
    print("\n[Configuration] Fichier introuvable, vérifiez le chemin d'accès :\n"+ path_to_excel)
    find = False
    cont = False

if find:
    print("\n[Configuration] Fichier trouvé avec succès")
    if check_version(classeur):
        print("\n[Configuration] Version de la simulation reconnue (" + str(version) + ")")
    else:
        print("\n[Configuration] Version de la simulation non-reconnue, contactez loan.challeat@thalesgroup.com pour signaler cet éventuel problème.")
        cont = False

if cont:
    nombre_amp = int(input("\n[Question] Combien d'AMP sont effectifs ? [1,2,3,4,5]"))
    
    liste_amp = create_liste_amp(nombre_amp)
    liste_amp_tableau = create_liste_amp_tableau(nombre_amp)
    
    liste_energie = energy_scraping(classeur, liste_amp)
    liste_energie_tableau = create_liste_energie_tableau(liste_energie)
    
    liste_power_tableau = create_liste_power_tableau()
    
    liste_power_val_tableau = create_liste_power_val_tableau(liste_energie)
    
    table = [liste_amp_tableau, liste_energie_tableau]
    tab = PrettyTable(table[0])
    tab.add_rows(table[1:])
    print("\n[Informations] Energie d'entrée et de sortie des amplificateurs")
    print(tab)
    
    table = [liste_power_tableau, liste_power_val_tableau]
    tab = PrettyTable(table[0])
    tab.add_rows(table[1:])
    print("\n[Informations] Puissance et objectif du système complet")
    print(tab)
    
    
    print("\n[Debug] Recherche de problèmes")
    
    liste_amp.pop()
    
    fluence_pompe_dommage(classeur, liste_amp)
    fluence_sortie(classeur, liste_amp)
    eclairement(classeur, liste_amp)
    largeur_spectrale(classeur, liste_energie)
    
    if not debug:
        print("[Debug] Aucun problème détecté")
