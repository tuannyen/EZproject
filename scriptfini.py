from selenium import webdriver
from selenium.webdriver.common.by import By
import pandas as pd
import time
import datetime
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import re

def Commandetofacture(numero_commande, ieme_commande):
    char = numero_commande.replace("-A", "")
    return f"FACV_{char}-{ieme_commande}"

def FacturetoCommande(reference):
    numero_complet = reference.split("_")[1]
    numero_final = "-".join(numero_complet.split("-")[:-1])
    return numero_final + "-A"

def extraire_dernier_numero(reference):
    return reference.split("-")[-1]

mois = input("Entrez le mois (01 à 12) : ").strip()
annee = input("Entrez l'année (YYYY) : ").strip()
EXCEL_FILE = f"ventes_{mois}_{annee}.xlsx"

try:
    df_old = pd.read_excel(EXCEL_FILE, sheet_name="matrice", skiprows=7)
    if "Numéro de facture" in df_old.columns:
        factures_existantes = df_old["Numéro de facture"].dropna().astype(str).tolist()
        derniere_facture = factures_existantes[-1] if factures_existantes else None
    else:
        factures_existantes = []
        derniere_facture = None
except Exception as e:
    print(f"Erreur lors de la lecture du fichier : {e}")
    factures_existantes = []
    derniere_facture = None

if derniere_facture is not None:
    derniere_commande_excel = FacturetoCommande(derniere_facture)
    dernier_num_commande = extraire_dernier_numero(derniere_facture)
else:
    derniere_commande_excel = None
    dernier_num_commande = input("Entrez le dernier numéro de commande : ")

options = webdriver.ChromeOptions()
options.add_argument(r"user-data-dir=CHEMIN/UTILISATEUR/ChromeProfile")
options.add_argument(r"profile-directory=Default")
options.add_argument("--disable-features=IsolateOrigins,site-per-process")
options.add_argument("--disable-popup-blocking")
options.add_argument("--disable-extensions")
driver = webdriver.Chrome(options=options)

time.sleep(2)

date_debut = datetime.datetime(int(annee), int(mois), 1, 0, 0, 0)
if int(mois) == 12:
    mois_suivant = 1
    annee_suivante = int(annee) + 1
else:
    mois_suivant = int(mois) + 1
    annee_suivante = int(annee)
date_fin = datetime.datetime(annee_suivante, mois_suivant, 1, 0, 0, 0)

startDate = int(time.mktime(date_debut.timetuple()) * 1000)
endDate = int(time.mktime(date_fin.timetuple()) * 1000)

URL_COMMANDES = f"https://exemple.com/orders?startDate={startDate}&endDate={endDate}"
driver.get(URL_COMMANDES)

time.sleep(4)

commande_element = driver.find_elements(By.XPATH, "(//td[contains(@class, '1btx70q') and contains(@id, 'order-list-date-created-id')])[1]")[0]
hauteur_element = driver.execute_script("return arguments[0].offsetHeight;", commande_element)
time.sleep(1)

nouvelles_factures = []
commandes_count = len(driver.find_elements(By.XPATH, "//div[contains(@class, 'ffbKvt')]"))
num_commande = str(int(int(dernier_num_commande) + commandes_count + 1))
scroll_value = 0

for i in range(commandes_count):
    commandes_elements = driver.find_elements(By.XPATH, "//div[contains(@class, 'ffbKvt')]")
    if i >= len(commandes_elements):
        break

    commande_element = commandes_elements[i]
    numero_commande = commande_element.text.strip()

    if derniere_commande_excel is not None and numero_commande == derniere_commande_excel:
        break

    commande_element.click()
    time.sleep(3)

    try:
        nom_client_element = driver.find_element(By.XPATH, "/html/body/div[1]/div/div[2]/div/div[2]/div/div/div/div[1]/div[2]/main/div/div[3]/div[2]/section/div[2]/div/p/span/span")
        nom_client = nom_client_element.text.strip()
    except:
        nom_client = "Nom introuvable"

    try:
        date_element = driver.find_element(By.XPATH, "/html/body/div[1]/div/div[2]/div/div[2]/div/div/div/div[1]/div[2]/div/div[1]/div/p")
        date_text = date_element.text.strip()
        match = re.search(r"\d{2}/\d{2}/\d{4}", date_text)
        date_commande = match.group(0) if match else "Date introuvable"
    except:
        date_commande = "Date introuvable"

    prix_ttc_element = driver.find_element(By.XPATH, "(//h3[contains(@class, 'iXUXQP') and contains(text(), '€')])[1]")
    prix_ttc_text = prix_ttc_element.text.strip().replace("€", "").replace(",", ".")
    prix_ttc = float(prix_ttc_text)

    prix_ht_elements = driver.find_elements(By.XPATH, "//div[contains(@class, 'eHcEem') and .//h3[contains(text(), 'SKU')]]//*[contains(text(), 'Prix produit total HT')]/following-sibling::*")
    tva_elements = driver.find_elements(By.XPATH, "//div[contains(@class, 'eHcEem') and .//h3[contains(text(), 'SKU')]]//div[contains(text(), 'Taxe')]")

    total_ht_20 = 0
    for i in range(len(prix_ht_elements)):
        prix_ht_text = prix_ht_elements[i].text.strip()
        tva_text = tva_elements[i].text.strip()
        try:
            prix_ht = float(prix_ht_text.split("x")[-1].replace("€", "").replace(",", ".").strip())
        except:
            continue
        if "20" in tva_text:
            total_ht_20 += prix_ht

    num_commande = str(int(num_commande) - 1)
    tempNumFacture = Commandetofacture(numero_commande, num_commande)
    nouvelles_factures.append({
        "Numéro de Facture": tempNumFacture,
        "Nom du client": nom_client,
        "Date de commande": date_commande,
        "TTC": prix_ttc,
        "Total HT 20%": total_ht_20
    })

    driver.back()
    time.sleep(2)

    scroll_value += hauteur_element
    driver.execute_script(f"window.scrollBy(0, {scroll_value});")
    time.sleep(2)

driver.quit()

df_nouvelles_factures = pd.DataFrame(nouvelles_factures)
print(df_nouvelles_factures.to_string(index=False))

if nouvelles_factures:
    wb = load_workbook(EXCEL_FILE)
    ws = wb["matrice"]
    col_num_facture = 2
    first_empty_row = None

    for row in range(9, ws.max_row + 1):
        cell = ws.cell(row=row, column=col_num_facture)
        cell_coord = f"{get_column_letter(col_num_facture)}{row}"
        is_merged = any(cell_coord in merged for merged in ws.merged_cells)
        if cell.value is None and not is_merged:
            first_empty_row = row
            break

    if first_empty_row:
        row_index = first_empty_row
        for facture in reversed(nouvelles_factures):
            ws.cell(row=row_index, column=2, value=facture["Numéro de Facture"])
            ws.cell(row=row_index, column=3, value=facture["Nom du client"])
            ws.cell(row=row_index, column=1, value=facture["Date de commande"])
            ws.cell(row=row_index, column=8, value=facture["TTC"])
            ws.cell(row=row_index, column=6, value=facture["Total HT 20%"])
            row_index += 1
        wb.save(EXCEL_FILE)
        wb.close()
    else:
        print("Aucune ligne vide et non fusionnée trouvée.")
else:
    print("Aucune nouvelle facture à ajouter.")
