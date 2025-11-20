import streamlit as st
import pdfplumber
import re
import os
import json
from openpyxl import load_workbook
from pathlib import Path
from PIL import Image
import zipfile
import io
from datetime import datetime

# --- CHEMINS RELATIFS ---
BASE_DIR = os.path.dirname(__file__)
CONFIG_FILE = os.path.join(BASE_DIR, 'config.json')
EAN_CORRECTIONS_FILE = os.path.join(BASE_DIR, 'corrections_ean.json')
EXCEL_TEMPLATE_FILE = os.path.join(BASE_DIR, 'EDI.xlsx')
LOGO_FILE = os.path.join(BASE_DIR, 'EDITHOR2.png')

# --- CONFIG STREAMLIT ---
st.set_page_config(page_title="EDITHORv3", layout="wide")
st.title("🧾 EDITHORv3 - PDF → Excel")

# --- AFFICHAGE LOGO ---
try:
    logo = Image.open(LOGO_FILE)
    st.image(logo, width=200)
except:
    st.warning("Logo non trouvé")

# --- GESTION CORRECTIONS EAN ---
def load_ean_corrections():
    try:
        with open(EAN_CORRECTIONS_FILE, 'r') as f:
            return json.load(f)
    except:
        with open(EAN_CORRECTIONS_FILE, 'w') as f:
            json.dump({}, f)
        return {}

def save_ean_corrections(corrections):
    with open(EAN_CORRECTIONS_FILE, 'w') as f:
        json.dump(corrections, f, indent=4)

ean_corrections = load_ean_corrections()

st.subheader("✍ Gestion Corrections EAN")
col1, col2 = st.columns(2)
with col1:
    old_ean = st.text_input("Ancien EAN")
with col2:
    new_ean = st.text_input("Nouveau EAN")

if st.button("Ajouter / Modifier EAN"):
    if old_ean and new_ean:
        ean_corrections[old_ean] = new_ean
        save_ean_corrections(ean_corrections)
        st.success(f"{old_ean} → {new_ean} enregistré")
    else:
        st.warning("Veuillez remplir les deux champs")

if st.button("Afficher toutes les corrections EAN"):
    st.json(ean_corrections)

# --- UPLOAD PDF(S) ---
uploaded_files = st.file_uploader("Déposez vos fichiers PDF", type="pdf", accept_multiple_files=True)

# --- NOM DOSSIER DE SORTIE INTERNE ---
timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
output_folder = os.path.join(BASE_DIR, f"EDITHOR_{timestamp}")
os.makedirs(output_folder, exist_ok=True)

# --- FONCTIONS DE TRAITEMENT PDF ---
def analyser_produit(line, ean_corrections):
    parts = re.split(r'\s+', line)
    if len(parts) >= 6:
        ean_brut = parts[2]
        ean_corrige = ean_corrections.get(ean_brut, ean_brut)
        return {
            "EAN": ean_corrige,
            "Description": " ".join(parts[3:-3]),
            "QuantiteCommandee": parts[-3],
            "PCB": parts[-2]
        }
    return {}

def parse_and_structure_text(text, commandes, current_commande, produits, inside_commande, ean_corrections):
    lines = text.split('\n')
    for line in lines:
        line = line.strip()
        if line.startswith("Commande n°"):
            inside_commande = True
            if current_commande:
                current_commande['Produits'] = produits
                commandes.append(current_commande)
                produits = []
            current_commande = {}

        if inside_commande:
            if line.startswith("Commande n°"):
                current_commande['Commande'] = line.split("Commande n°")[1].strip()
            elif line.startswith("Fournisseur"):
                current_commande['Fournisseur'] = line.split(":")[1].strip()
            elif line.startswith("Document"):
                current_commande['DateCommande'] = line.split(":")[1].strip()
            elif line.startswith("Livraison le"):
                current_commande['DateLivraison'] = line.split(":")[1].strip()
            elif "BAK FRANCE" in line:
                current_commande['NomClient'] = line.split("BAK FRANCE")[1].strip()
            elif line.startswith("Lieu dit"):
                current_commande['Adresse'] = line
            elif line.startswith("Poids total brut produits"):
                current_commande['PoidsTotal'] = line.split(":")[1].strip()
            elif line.startswith("Montant total ht commande"):
                current_commande['MontantTotal'] = line.split(":")[1].strip()
            elif re.match(r"^\d+ \d+", line):
                produit = analyser_produit(line, ean_corrections)
                if produit:
                    produits.append(produit)
            elif line.startswith("Récapitulatif"):
                inside_commande = False
                if current_commande:
                    current_commande['Produits'] = produits
                    commandes.append(current_commande)
                    produits = []
                    current_commande = None
    return commandes, current_commande, produits, inside_commande

def extraire_et_structurer_texte_pdf(pdf_file, ean_corrections):
    commandes, current_commande, produits, inside_commande = [], None, [], False
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                commandes, current_commande, produits, inside_commande = parse_and_structure_text(
                    text, commandes, current_commande, produits, inside_commande, ean_corrections
                )
    if current_commande and produits:
        current_commande['Produits'] = produits
        commandes.append(current_commande)
    return commandes

# --- GÉNÉRATION EXCEL ---
def creer_excel_a_partir_du_modele(modele_file, output_folder, commandes):
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
    for commande in commandes:
        if not commande.get('Produits'):
            continue
        wb = load_workbook(modele_file)
        ws = wb.active
        ws['E2'] = commande.get('DateCommande', "")[:10]
        ws['F2'] = commande.get('DateLivraison', "")[:10]
        nom_client = commande.get('NomClient', "").split("BAK")[0].strip()
        ws['I2'] = ws['K2'] = nom_client
        ws['L2'] = ws['M2'] = ws['N2'] = ""
        numero_commande = commande.get('Commande', "")
        nom_fichier = f"{nom_client}_{numero_commande}".replace(" ", "_").strip("_")
        ws['O2'] = nom_fichier
        for i, produit in enumerate(commande['Produits'], start=4):
            ws[f'C{i}'] = produit.get('EAN', "")
            ws[f'D{i}'] = 'PCE'
            ws[f'E{i}'] = produit.get('Description', "")
            ws[f'F{i}'] = produit.get('QuantiteCommandee', "")
            ws[f'G{i}'] = produit.get('PCB', "")
        wb.save(os.path.join(output_folder, f"{nom_fichier}.xlsx"))

# --- TRAITEMENT PDF ET GÉNÉRATION EXCEL ---
if uploaded_files:
    if st.button("🛠️ Traiter PDF(s) et générer Excel"):
        for pdf_file in uploaded_files:
            commandes = extraire_et_structurer_texte_pdf(pdf_file, ean_corrections)
            creer_excel_a_partir_du_modele(EXCEL_TEMPLATE_FILE, output_folder, commandes)
        st.success(f"✅ Excel généré dans le dossier temporaire '{output_folder}' !")

        # --- CRÉER ZIP POUR TÉLÉCHARGEMENT ---
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED) as zf:
            for f in os.listdir(output_folder):
                zf.write(os.path.join(output_folder, f), f)
        zip_buffer.seek(0)
        zip_name = f"EDITHOR_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip"
        st.download_button("📥 Télécharger tous les Excel", zip_buffer, file_name=zip_name)
