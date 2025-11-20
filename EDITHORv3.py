import streamlit as st
import pdfplumber
import pandas as pd
import json
import re
from datetime import datetime
from pathlib import Path
import os

st.set_page_config(page_title="EDITHOR", layout="wide")

st.markdown("<h1 style='text-align:center; color:#007aff;'>EDITHOR</h1>", unsafe_allow_html=True)
st.markdown("---")

# --- CHEMINS ET CONFIGURATION ---
CONFIG_FILE = 'config.json'
EAN_CORRECTIONS_FILE = 'corrections_ean.json'
EXCEL_TEMPLATE_FILE = 'EDI.xlsx'

def load_config():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, 'r') as f:
            return json.load(f)
    return {}

def save_config(config):
    with open(CONFIG_FILE, 'w') as f:
        json.dump(config, f)

config = load_config()

# --- SIDEBAR pour Excel et dossier de sortie ---
st.sidebar.header("Paramètres")
uploaded_template = st.sidebar.file_uploader("Modèle Excel", type=["xlsx"], key="excel_template")
output_folder = st.sidebar.text_input("Dossier de sortie (ex: Downloads/EDITHOR)", value=str(Path.home() / "Downloads/EDITHOR"))

os.makedirs(output_folder, exist_ok=True)

# --- CORRECTIONS EAN ---
st.sidebar.markdown("### Gestion EAN")
if os.path.exists(EAN_CORRECTIONS_FILE):
    with open(EAN_CORRECTIONS_FILE, 'r') as f:
        ean_corrections = json.load(f)
else:
    ean_corrections = {}

ean_action = st.sidebar.selectbox("Action", ["Ajouter", "Modifier", "Supprimer"])
old_ean = st.sidebar.text_input("Ancien EAN")
new_ean = st.sidebar.text_input("Nouveau EAN")

if st.sidebar.button("Valider EAN"):
    if ean_action == "Ajouter":
        if old_ean and new_ean:
            ean_corrections[old_ean] = new_ean
            st.sidebar.success(f"EAN {old_ean} ajouté → {new_ean}")
    elif ean_action == "Modifier":
        if old_ean in ean_corrections:
            ean_corrections[old_ean] = new_ean
            st.sidebar.success(f"EAN {old_ean} modifié → {new_ean}")
        else:
            st.sidebar.warning("EAN non trouvé pour modifier")
    elif ean_action == "Supprimer":
        if old_ean in ean_corrections:
            del ean_corrections[old_ean]
            st.sidebar.success(f"EAN {old_ean} supprimé")
        else:
            st.sidebar.warning("EAN non trouvé pour supprimer")
    with open(EAN_CORRECTIONS_FILE, 'w') as f:
        json.dump(ean_corrections, f, indent=4)

# --- UPLOAD PDF ---
uploaded_files = st.file_uploader("Sélectionnez le(s) PDF(s)", type=["pdf"], accept_multiple_files=True)

# --- PROCESSUS DE GÉNÉRATION EXCEL ---
def extract_and_process_pdf(pdf_file, corrections):
    commandes, current_commande, produits, inside_commande = [], None, [], False
    with pdfplumber.open(pdf_file) as pdf:
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
    for commande in commandes:
        if not commande.get('Produits'):
            continue
        wb = pd.ExcelWriter(os.path.join(output_path, f"{commande.get('Commande', 'commande')}.xlsx"), engine='openpyxl')
        df = pd.DataFrame(commande['Produits'])
        df.to_excel(wb, index=False)
        wb.close()

# --- BOUTON GENERER ---
if st.button("📂 Générer Excel(s)"):
    if uploaded_files:
        for pdf_file in uploaded_files:
            pdf_bytes = pdf_file.read()
            with open("temp.pdf", "wb") as f:
                f.write(pdf_bytes)
            commandes = extract_and_process_pdf("temp.pdf", ean_corrections)
            create_excel_from_template(EXCEL_TEMPLATE_FILE, output_folder, commandes)
        st.success(f"Les fichiers Excel ont été créés dans : {output_folder}")

# --- FOOTER / SIGNATURE ---
st.markdown("---")
st.markdown("<p style='text-align:center; color:#ffaa00; font-size:20px;'>★★★★★</p>", unsafe_allow_html=True)
st.markdown("<p style='text-align:center; color:#8e8e93; font-size:12px; font-style:italic;'>Powered by IC - 2025</p>", unsafe_allow_html=True)
