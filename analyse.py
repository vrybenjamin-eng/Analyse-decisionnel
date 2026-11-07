import csv
import json
from openpyxl import load_workbook
 
BOUTIQUES = ["Marché Lagune", "Baobab Shop", "Comptoir Ivoire"]
 
# ----------------------------------------------------------------------
# Variables globales (remplies par l'option 1 du menu)
# ----------------------------------------------------------------------
dossier = "donnees"
catalogue = {}        # "P01" -> "Samsung Galaxy A15 128 Go"
offres = []           # liste de dictionnaires : une offre = une ligne d'une boutique
journal = []          # liste de textes : tout ce qu'on a nettoyé ou refusé
prix_valides = {}     # "P01" -> {"Baobab Shop": 104000, "Comptoir Ivoire": 92990}
donnees_chargees = False
 
 
# ----------------------------------------------------------------------
# ETAPE 2 : NORMALISATION
# ----------------------------------------------------------------------
def normaliser_nom(nom):
    """Met le nom en minuscules, enlève les accents, la ponctuation et les espaces en trop."""
    nom = str(nom).lower()
    for avec, sans in [("é", "e"), ("è", "e"), ("ê", "e"), ("à", "a"), ("ç", "c"), ("ô", "o"), ("î", "i")]:
        nom = nom.replace(avec, sans)
    resultat = ""
    for lettre in nom:
        if lettre.isalnum():          # lettre ou chiffre : on garde
            resultat = resultat + lettre
        else:                         # tout le reste (tiret, guillemet, point...) devient un espace
            resultat = resultat + " "
    return " ".join(resultat.split())   # split() + join() supprime les espaces en trop
 
 
def normaliser_prix(valeur):
    """Transforme '102 900 FCFA' en 102900. Retourne None si le prix est vide ou inutilisable."""
    if valeur is None:
        return None
    if isinstance(valeur, (int, float)):     # déjà un nombre
        return int(valeur)
    chiffres = ""
    for caractere in str(valeur):
        if caractere.isdigit():
            chiffres = chiffres + caractere
    if chiffres == "":                       # ex : "" ou "Prix sur demande"
        return None
    return int(chiffres)
 
 
# ----------------------------------------------------------------------
# ETAPE 1 : COLLECTE (une fonction par format de fichier)
# ----------------------------------------------------------------------
def lire_catalogue():
    with open(dossier + "/catalogue_reference.csv", encoding="utf-8-sig") as f:
        for ligne in csv.DictReader(f):
            catalogue[ligne["produit_id"]] = ligne["nom"]
 
 
def lire_lagune():
    deja_vus = []
    with open(dossier + "/marche_lagune.csv", encoding="utf-8-sig") as f:   # utf-8-sig enlève le BOM
        for ligne in csv.DictReader(f, delimiter=";"):
            ref = ligne["reference"]
            if ref in deja_vus:
                journal.append("Marché Lagune : doublon " + ref + " ignoré")
                continue
            deja_vus.append(ref)
 
            prix = normaliser_prix(ligne["prix"])
            if prix is None:
                journal.append("Marché Lagune : " + ref + " sans prix utilisable ('" + ligne["prix"] + "')")
            en_stock = ligne["disponibilite"] != "Rupture"
            # le prix de Lagune inclut déjà la livraison
            offres.append({"boutique": "Marché Lagune", "ref": ref, "nom": ligne["designation"],
                           "prix": prix, "en_stock": en_stock})
 
 
def lire_baobab():
    with open(dossier + "/baobab_shop.json", encoding="utf-8") as f:
        donnees = json.load(f)
    for article in donnees["items"]:
        prix = normaliser_prix(article["pricing"]["amount"])
        if prix is None:
            journal.append("Baobab Shop : " + article["sku"] + " sans prix (null)")
        else:
            prix = prix + article["pricing"]["delivery_fee"]     # le client paie prix + livraison
        offres.append({"boutique": "Baobab Shop", "ref": article["sku"], "nom": article["title"],
                       "prix": prix, "en_stock": article["stock"] > 0})
 
 
def lire_comptoir():
    classeur = load_workbook(dossier + "/comptoir_ivoire.xlsx", data_only=True)
    feuille = classeur.active
    # les données commencent à la ligne 4 (ligne 1 = titre, ligne 3 = en-têtes)
    for ligne in feuille.iter_rows(min_row=4, values_only=True):
        code, article, rayon, prix_catalogue, remise, quantite = ligne[0:6]
        if code is None:
            continue
        if isinstance(prix_catalogue, str):
            journal.append("Comptoir Ivoire : " + code + " prix écrit en texte, converti en nombre")
        prix = normaliser_prix(prix_catalogue)
        if prix is not None:
            prix = round(prix * (1 - remise / 100))      # le client paie après la remise
        offres.append({"boutique": "Comptoir Ivoire", "ref": code, "nom": article,
                       "prix": prix, "en_stock": quantite > 0})
 
 
# ----------------------------------------------------------------------
# ETAPE 3 : INDEXATION - retrouver à quel produit du catalogue correspond une offre
# Pour chaque produit, on donne des mots-clés. "a|b" veut dire : a OU b.
# Une offre correspond au produit si TOUS ses mots-clés sont dans son nom.
# ----------------------------------------------------------------------
MOTS_CLES = {
    "P01": ["a15"], "P02": ["a25"], "P03": ["a55"], "P04": ["spark"], "P05": ["camon"],
    "P06": ["hot"], "P07": ["infinix", "note"], "P08": ["itel"], "P09": ["iphone"],
    "P10": ["redmi"], "P11": ["250"], "P12": ["ideapad"], "P13": ["inspiron"],
    "P14": ["sandisk"], "P15": ["seagate"], "P16": ["m185"], "P17": ["deskjet"],
    "P18": ["archer"], "P19": ["samsung", "ecran|monitor"],
    "P20": ["hisense", "refrigerateur|fridge|frigo"], "P21": ["nasco"],
    "P22": ["samsung", "tv"], "P23": ["hisense", "tv"],
    "P24": ["lg", "climatiseur|clim|conditioner"], "P25": ["binatone"],
    "P26": ["moulinex", "mixeur|blender"], "P27": ["philips", "fer|iron"],
    "P28": ["samsung", "micro|microwave"], "P29": ["bouilloire|kettle"],
    "P30": ["hisense", "climatiseur|clim|conditioner"],
}
 
 
def trouver_produit(nom):
    """Retourne le produit_id (ex: 'P04') ou None si on ne trouve pas."""
    mots_du_nom = normaliser_nom(nom).split()
    for produit_id in MOTS_CLES:
        tout_trouve = True
        for exigence in MOTS_CLES[produit_id]:
            une_option_trouvee = False
            for option in exigence.split("|"):
                if option in mots_du_nom:
                    une_option_trouvee = True
            if not une_option_trouvee:
                tout_trouve = False
        if tout_trouve:
            return produit_id
    return None
 
 
# ----------------------------------------------------------------------
# Option 1 du menu : tout charger, nettoyer et préparer
# ----------------------------------------------------------------------
def charger_donnees():
    global donnees_chargees
    catalogue.clear(); offres.clear(); journal.clear(); prix_valides.clear()
 
    lire_catalogue()
    lire_lagune()
    lire_baobab()
    lire_comptoir()
    print("Offres lues :", len(offres))
 
    # on associe chaque offre à un produit et on garde seulement les offres utilisables
    for offre in offres:
        offre["produit_id"] = trouver_produit(offre["nom"])
        if offre["produit_id"] is None:
            journal.append(offre["boutique"] + " : '" + offre["nom"] + "' introuvable dans le catalogue")
            continue
        if offre["prix"] is None:
            continue                                  # déjà noté dans le journal
        if not offre["en_stock"]:
            journal.append(offre["boutique"] + " : " + offre["produit_id"] + " en rupture, exclu")
            continue
        pid = offre["produit_id"]
        if pid not in prix_valides:
            prix_valides[pid] = {}
        prix_valides[pid][offre["boutique"]] = offre["prix"]
 
    # prix aberrants : plus de 3 fois le prix le plus bas du même produit
    for pid in prix_valides:
        prix_min = min(prix_valides[pid].values())
        for boutique in list(prix_valides[pid]):
            if prix_valides[pid][boutique] > 3 * prix_min:
                journal.append(boutique + " : " + pid + " prix aberrant (" +
                               str(prix_valides[pid][boutique]) + " FCFA), exclu")
                del prix_valides[pid][boutique]
 
    donnees_chargees = True
    print("Nettoyage terminé :", len(journal), "remarques dans le journal.")
    voir = input("Afficher le journal ? (o/n) : ")
    if voir.lower() == "o":
        for ligne in journal:
            print(" -", ligne)
 
 
# ----------------------------------------------------------------------
# ETAPE 4 : COMPARAISON
# ----------------------------------------------------------------------
def calculer_comparatif():
    """Retourne une liste de lignes (une par produit présent dans au moins 2 boutiques)."""
    lignes = []
    for pid in sorted(prix_valides):
        offres_produit = prix_valides[pid]
        if len(offres_produit) < 2:           # une seule source : rien à comparer
            continue
        prix_min = min(offres_produit.values())
        prix_max = max(offres_produit.values())
        gagnants = []
        for boutique in offres_produit:
            if offres_produit[boutique] == prix_min:
                gagnants.append(boutique)
        lignes.append({
            "id": pid, "produit": catalogue[pid], "prix": offres_produit,
            "meilleure": " / ".join(gagnants), "prix_min": prix_min,
            "ecart": prix_max - prix_min,
            "ecart_pct": round((prix_max - prix_min) / prix_min * 100, 1),
        })
    return lignes
 
 
def afficher_comparatif():
    lignes = calculer_comparatif()
    print()
    print("Produit".ljust(38) + "Meilleure boutique".ljust(32) + "Prix min".rjust(10) + "Ecart".rjust(10) + "  %")
    print("-" * 95)
    for l in lignes:
        print(l["produit"][:36].ljust(38) + l["meilleure"].ljust(32) +
              str(l["prix_min"]).rjust(10) + str(l["ecart"]).rjust(10) + "  " + str(l["ecart_pct"]))
 
    # export CSV (ouvrable dans Excel)
    with open("comparatif_prix.csv", "w", encoding="utf-8-sig", newline="") as f:
        ecrivain = csv.writer(f, delimiter=";")
        ecrivain.writerow(["produit_id", "produit"] + BOUTIQUES +
                          ["meilleure_boutique", "prix_min", "ecart_fcfa", "ecart_pct"])
        for l in lignes:
            ecrivain.writerow([l["id"], l["produit"]] + [l["prix"].get(b, "") for b in BOUTIQUES] +
                              [l["meilleure"], l["prix_min"], l["ecart"], l["ecart_pct"]])
    print("\n" + str(len(lignes)) + " produits comparés. Fichier créé : comparatif_prix.csv")
 
 
# ----------------------------------------------------------------------
# ETAPE 5 : DECISION
# ----------------------------------------------------------------------
def afficher_synthese():
    lignes = calculer_comparatif()
    print("\n--- Combien de fois chaque boutique est la moins chère ---")
    for boutique in BOUTIQUES:
        victoires = 0
        for l in lignes:
            if boutique in l["meilleure"]:
                victoires = victoires + 1
        print(" ", boutique.ljust(18), victoires)
 
    print("\n--- Ecart moyen par rapport au prix le plus bas ---")
    for boutique in BOUTIQUES:
        total = 0
        nombre = 0
        for l in lignes:
            if boutique in l["prix"]:
                total = total + (l["prix"][boutique] - l["prix_min"]) / l["prix_min"] * 100
                nombre = nombre + 1
        if nombre > 0:
            print(" ", boutique.ljust(18), round(total / nombre, 1), "%")
 
    print("\n--- Les 3 produits où il y a le plus à gagner à comparer ---")
    classement = sorted(lignes, key=lambda l: l["ecart"], reverse=True)
    for l in classement[:3]:
        print(" ", l["produit"], ":", l["ecart"], "FCFA (", l["ecart_pct"], "%)")
 
 
def afficher_panier():
    panier = {"P04": 1, "P12": 1, "P14": 2, "P15": 1, "P28": 1}   # produit -> quantité
    print("\n--- Panier de rentrée ---")
    for boutique in BOUTIQUES:
        total = 0
        manquants = []
        for pid in panier:
            if boutique in prix_valides.get(pid, {}):
                total = total + prix_valides[pid][boutique] * panier[pid]
            else:
                manquants.append(pid)
        if manquants:
            print(" ", boutique.ljust(18), "incomplet (pas de prix valide pour", ", ".join(manquants) + ")")
        else:
            print(" ", boutique.ljust(18), total, "FCFA")
 
    print("\n  Chaque article là où il est le moins cher :")
    total = 0
    for pid in panier:
        offres_produit = prix_valides[pid]
        meilleure = min(offres_produit, key=offres_produit.get)   # boutique au prix le plus bas
        sous_total = offres_produit[meilleure] * panier[pid]
        total = total + sous_total
        print("   ", catalogue[pid][:30].ljust(32), meilleure.ljust(18), sous_total)
    print("  TOTAL optimisé :", total, "FCFA")
 
 
def afficher_historique():
    """Lit le gros fichier UNE SEULE FOIS, ligne par ligne, sans tout garder en mémoire."""
    somme_annee = {}      # (boutique, produit) -> total des prix
    nb_annee = {}         # (boutique, produit) -> nombre de jours
    somme_bf = {}         # idem mais seulement pour la semaine du Black Friday
    nb_bf = {}
    with open(dossier + "/historique_prix.csv", encoding="utf-8-sig") as f:
        lecteur = csv.reader(f)
        next(lecteur)                                  # on saute la ligne d'en-têtes
        for date, boutique, pid, prix in lecteur:
            prix = int(prix)
            cle = (boutique, pid)
            somme_annee[cle] = somme_annee.get(cle, 0) + prix
            nb_annee[cle] = nb_annee.get(cle, 0) + 1
            if "2025-11-24" <= date <= "2025-11-30":   # les dates AAAA-MM-JJ se comparent comme du texte
                somme_bf[cle] = somme_bf.get(cle, 0) + prix
                nb_bf[cle] = nb_bf.get(cle, 0) + 1
 
    print("\n--- Black Friday (24-30 nov. 2025) : prix de la semaine / prix moyen de l'année ---")
    for boutique in BOUTIQUES:
        total = 0
        nombre = 0
        for cle in somme_bf:
            if cle[0] == boutique:
                moyenne_annee = somme_annee[cle] / nb_annee[cle]
                moyenne_bf = somme_bf[cle] / nb_bf[cle]
                total = total + (moyenne_bf - moyenne_annee) / moyenne_annee * 100
                nombre = nombre + 1
        print(" ", boutique.ljust(18), round(total / nombre, 1), "%")
    print("\n  (un chiffre négatif = la semaine du Black Friday est moins chère que la moyenne)")
 
 
def chercher_produit():
    texte = normaliser_nom(input("Nom du produit à chercher (ex: iphone, tv, clim) : "))
    trouve = False
    for pid in sorted(catalogue):
        if texte in normaliser_nom(catalogue[pid]):
            trouve = True
            print("\n" + pid, catalogue[pid])
            offres_produit = prix_valides.get(pid, {})
            if not offres_produit:
                print("   aucun prix valide")
            for boutique in offres_produit:
                print("   ", boutique.ljust(18), offres_produit[boutique], "FCFA")
    if not trouve:
        print("Aucun produit trouvé.")
 
 
# ----------------------------------------------------------------------
# MENU PRINCIPAL
# ----------------------------------------------------------------------
def menu():
    global dossier
    reponse = input("Dossier des fichiers de données [donnees] : ")
    if reponse.strip() != "":
        dossier = reponse.strip()
 
    while True:
        print("\n===== ANALYSE COMPARATIVE DE MARCHE =====")
        print("1. Charger et nettoyer les données")
        print("2. Comparatif des prix (et export CSV)")
        print("3. Synthèse : quelle boutique est la meilleure ?")
        print("4. Panier de rentrée")
        print("5. Historique : Black Friday")
        print("6. Chercher un produit")
        print("0. Quitter")
        choix = input("Votre choix : ").strip()
 
        if choix == "0":
            print("Au revoir !")
            break
        elif choix == "1":
            charger_donnees()
        elif choix in ["2", "3", "4", "6"]:
            if not donnees_chargees:
                print("Faites d'abord l'option 1.")
            elif choix == "2":
                afficher_comparatif()
            elif choix == "3":
                afficher_synthese()
            elif choix == "4":
                afficher_panier()
            else:
                chercher_produit()
        elif choix == "5":
            afficher_historique()
        else:
            print("Choix invalide.")
 
 
menu()

