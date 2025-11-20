import streamlit as st
import pdfplumber
import pandas as pd
import json
import re
import os
from pathlib import Path
from datetime import datetime
from PIL import Image
import zipfile
import io
import requests
import openpyxl

# --- CONFIGURATION PAGE ---
st.set_page_config(page_title="EDITHOR", layout="wide")

# --- LOGO CENTRÉ ---
logo_col1, logo_col2, logo_col3 = st.columns([1,3,1])
with logo_col2:
    try:
        logo = Image.open("EDITHOR2.png")
        st.image(logo, width=200)
    except:
        st.warning("Logo non trouvé : EDITHOR2.png")

# --- DESCRIPTION COURTE ---
st.markdown("<p style='text-align:center; font-size:14px; color:#555;'>Application de traitement PDF → Excel pour vos commandes BAK France</p>", unsafe_allow_html=True)

# --- BOUTON AIDE ---
with st.expander("❓ Aide"):
    st.markdown("""
    **Comment utiliser l'application :**
    1. Sélectionnez un ou plusieurs fichiers PDF de commandes.
    2. Cliquez sur **Générer Excel(s)** pour créer les fichiers.
    3. Téléchargez vos fichiers via :
       - **Télécharger tout en ZIP** : regroupe tous les fichiers dans un seul ZIP.
       - **Télécharger tous les Excel** : télécharge tous les fichiers un par un directement.
       - **Boutons individuels** à côté de chaque fichier pour un téléchargement séparé.
    4. Gérez les corrections EAN dans la section prévue : ajouter, modifier ou supprimer.
    5. Vous pouvez tout supprimer et recommencer via le bouton prévu.
    """)

# --- CONFIG ET CHEMINS ---
CONFIG_FILE = 'config.json'
EAN_CORRECTIONS_FILE = 'corrections_ean.json'

# Modèle Excel depuis GitHub
GITHUB_MODEL_URL = "https://raw.githubusercontent.com/<ton_user>/<repo>/main/EDI.xlsx"
EXCEL_TEMPLATE_FILE = "EDI.xlsx"

# Télécharger le modèle Excel si non présent
if not os.path.exists(EXCEL_TEMPLATE_FILE):
    r = requests.get(GITHUB_MODEL_URL)
    if r.status_code == 200:
        with open(EXCEL_TEMPLATE_FILE, "wb") as f:
            f.write(r.content)
    else:
        st.error("Impossible de récupérer le modèle Excel depuis GitHub.")
        st.stop()

# Dossier de sortie temporaire
output_folder = Path.home() / "Downloads" / f"EDITHOR_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
os.makedirs(output_folder, exist_ok=True)

# Charger corrections EAN
if os.path.exists(EAN_CORRECTIONS_FILE):
    with open(EAN_CORRECTIONS_FILE, "r") as f:
        ean_corrections = json.load(f)
else:
    ean_corrections = {}

# --- UPLOAD PDF ---
st.subheader("1️⃣ Sélectionnez le(s) PDF(s) à traiter")
uploaded_files = st.file_uploader("Sélectionnez un ou plusieurs PDF", type=["pdf"], accept_multiple_files=True)

# --- FONCTIONS PDF ET EXCEL ---
def extract_and_process_pdf(pdf_bytes, corrections):
    commandes, current_commande, produits, inside_commande = [], None, [], False
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                commandes, current_commande, produits, inside_commande = parse_text(
                    text, commandes, current_commande, produits, inside_commande, corrections
                )
    if current_commande and produits:
        current_commande['Produits'] = produits
        commandes.append(current_commande)
    return commandes

def parse_text(text, commandes, current_commande, produits, inside_commande, corrections):
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
                produit = analyse_product(line, corrections)
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

def analyse_product(line, corrections):
    parts = re.split(r'\s+', line)
    if len(parts) >= 6:
        ean_brut = parts[2]
        ean_corrige = corrections.get(ean_brut, ean_brut)
        return {
            "EAN": ean_corrige,
            "Description": " ".join(parts[3:-3]),
            "QuantiteCommandee": parts[-3],
            "PCB": parts[-2]
        }
    return {}

def create_excel_from_template(modele_path, output_path, commandes):
    created_files = []
    for commande in commandes:
        if not commande.get('Produits'):
            continue
        wb = openpyxl.load_workbook(modele_path)
        ws = wb.active
        ws['E2'] = commande.get('DateCommande', '')[:10]
        ws['F2'] = commande.get('DateLivraison', '')[:10]
        nom_client = commande.get('NomClient', '').split("BAK")[0].strip()
        ws['I2'] = ws['K2'] = nom_client
        ws['L2'] = ws['M2'] = ws['N2'] = ''
        numero_commande = commande.get('Commande', '')
        nom_fichier = f"{nom_client}_{numero_commande}".replace(" ", "_").strip("_")
        ws['O2'] = nom_fichier
        for i, produit in enumerate(commande['Produits'], start=4):
            ws[f'C{i}'] = produit.get('EAN', "")
            ws[f'D{i}'] = 'PCE'
            ws[f'E{i}'] = produit.get('Description', "")
            ws[f'F{i}'] = produit.get('QuantiteCommandee', "")
            ws[f'G{i}'] = produit.get('PCB', "")
        save_path = os.path.join(output_path, f"{nom_fichier}.xlsx")
        wb.save(save_path)
        created_files.append(save_path)
    return created_files

# --- TRAITEMENT PDF ---
generated_files = []

if st.button("📂 Générer Excel(s)"):
    if uploaded_files:
        for pdf_file in uploaded_files:
            pdf_bytes = pdf_file.read()
            commandes = extract_and_process_pdf(pdf_bytes, ean_corrections)
            files = create_excel_from_template(EXCEL_TEMPLATE_FILE, output_folder, commandes)
            generated_files.extend(files)
        if generated_files:
            st.success(f"{len(generated_files)} fichiers Excel générés.")
        else:
            st.warning("Aucun fichier Excel généré.")
    else:
        st.warning("Veuillez sélectionner au moins un PDF.")

# --- TABLEAU DES FICHIERS GENERES ---
if generated_files:
    st.subheader("2️⃣ Fichiers Excel générés")

    # Créer fichiers en mémoire
    file_bytes_list = []
    for file_path in generated_files:
        with open(file_path, "rb") as f:
            file_bytes_list.append((os.path.basename(file_path), f.read()))

    col1, col2, col3 = st.columns(3)

    # 1️⃣ Télécharger tout en ZIP
    with col1:
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w") as zipf:
            for fname, fbytes in file_bytes_list:
                zipf.writestr(fname, fbytes)
        zip_buffer.seek(0)
        st.download_button("⬇️ Tout télécharger (ZIP)", data=zip_buffer, file_name="EDITHOR_All.zip")

    # 2️⃣ Télécharger tous les Excel directement
    with col2:
        for fname, fbytes in file_bytes_list:
            st.download_button(label=f"⬇️ {fname}", data=fbytes, file_name=fname, key=f"dl_{fname}")

    # 3️⃣ Tout supprimer
    with col3:
        if st.button("🗑️ Tout supprimer / Recommencer"):
            for file_path in generated_files:
                os.remove(file_path)
            generated_files.clear()
            st.experimental_rerun()

    # Affichage tableau
    df_files = pd.DataFrame({"Nom du fichier": [fname for fname, _ in file_bytes_list]})
    st.dataframe(df_files.style.set_properties(**{'background-color': '#f0f8ff', 'color': 'black'}), height=200)

# --- GESTION EAN INTERACTIVE ---
st.subheader("✍ Gestion des Corrections EAN")
with st.expander("Afficher / Modifier les corrections EAN"):
    # Ajouter EAN
    if st.button("+ Ajouter une correction"):
        ean_corrections["Nouvel_EAN"] = "Valeur"
        with open(EAN_CORRECTIONS_FILE, "w") as f:
            json.dump(ean_corrections, f, indent=4)
        st.experimental_rerun()

    # Tableau avec actions
    if ean_corrections:
        df_ean = pd.DataFrame(list(ean_corrections.items()), columns=["Ancien EAN", "Nouveau EAN"])
        for idx, row in df_ean.iterrows():
            col1, col2, col3 = st.columns([3,3,2])
            with col1:
                st.text_input("Ancien EAN", value=row["Ancien EAN"], key=f"old_{idx}")
            with col2:
                st.text_input("Nouveau EAN", value=row["Nouveau EAN"], key=f"new_{idx}")
            with col3:
                modif = st.button("Modifier", key=f"mod_{idx}")
                supp = st.button("Supprimer", key=f"supp_{idx}")
                if modif:
                    ean_corrections[st.session_state[f"old_{idx}"]] = st.session_state[f"new_{idx}"]
                    with open(EAN_CORRECTIONS_FILE, "w") as f:
                        json.dump(ean_corrections, f, indent=4)
                    st.experimental_rerun()
                if supp:
                    ean_corrections.pop(row["Ancien EAN"], None)
                    with open(EAN_CORRECTIONS_FILE, "w") as f:
                        json.dump(ean_corrections, f, indent=4)
                    st.experimental_rerun()

# --- FOOTER ---
st.markdown("---")
st.markdown("<p style='text-align:center; color:#ffaa00; font-size:20px;'>★★★★★</p>", unsafe_allow_html=True)
st.markdown("<p style='text-align:center; color:#8e8e93; font-size:12px; font-style:italic;'>Powered by IC - 2025</p>", unsafe_allow_html=True)
