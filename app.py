import os
import re
import zipfile
import shutil
import fitz   # PyMuPDF
import pandas as pd

from unidecode import unidecode
from rapidfuzz import process, fuzz

import streamlit as st

# -----------------------------
# Fonctions utilitaires
# -----------------------------
def smart_capitalize(s: str) -> str:
    def cap_token(tok: str) -> str:
        parts = tok.split("-")
        parts = [p[:1].upper() + p[1:].lower() if p else p for p in parts]
        return "-".join(parts)
    return "'".join(cap_token(p) for p in s.split("'"))

def norm_key(s: str) -> str:
    s = unidecode(s).lower()
    s = s.replace("-", " ")
    s = re.sub(r"\s+", " ", s).strip()
    return s

PATTERN = re.compile(r"MISSION\s*N[°º]?\s*\d+\s*DE\s*M(?:\.|ME)\s+(.+)$", flags=re.IGNORECASE)

def extract_prenom_and_tokens(text_page0: str):
    for line in (ln.strip() for ln in text_page0.splitlines() if ln.strip()):
        m = PATTERN.search(line)
        if m:
            tail = m.group(1).strip()
            tokens = tail.split()
            if not tokens:
                return None, []
            prenom_cap = smart_capitalize(tokens[0])
            return prenom_cap, tokens
    return None, []

def generate_candidates(tokens):
    if len(tokens) < 1:
        return []
    prenom_cap = smart_capitalize(tokens[0])
    cands = []
    for k in (1, 2, 3):
        if len(tokens) >= 1 + k:
            nom_tokens = tokens[-k:]
            nom_upper = " ".join(t.upper() for t in nom_tokens)
            cle_excel = f"{nom_upper} {prenom_cap}".strip()
            cle_affiche = f"{prenom_cap} {nom_upper}".strip()
            cands.append((cle_affiche, cle_excel))
    return cands

# -----------------------------
# Interface Streamlit
# -----------------------------
st.title("📂 Classement automatique des factures VIE")

zip_file = st.file_uploader("Uploader le fichier ZIP des factures VIE", type="zip")
excel_file = st.file_uploader("Uploader le fichier Excel Table de référence", type=["xls", "xlsx"])

if st.button("Lancer le traitement"):
    if not zip_file or not excel_file:
        st.error("⚠️ Merci d'uploader le ZIP et l'Excel")
    else:
        # chemins temporaires
        EXTRACT_DIR = "factures_extraites"
        FINAL_DIR = "factures_classees"
        RAPPORT_OUT = "rapport_classement.xlsx"
        ZIP_OUT = "resultat_complet.zip"

        with open("input.zip", "wb") as f: f.write(zip_file.read())
        with open("mapping.xlsx", "wb") as f: f.write(excel_file.read())

        # 1) Décompresser
        if os.path.exists(EXTRACT_DIR):
            shutil.rmtree(EXTRACT_DIR)
        os.makedirs(EXTRACT_DIR, exist_ok=True)
        with zipfile.ZipFile("input.zip", "r") as zf:
            zf.extractall(EXTRACT_DIR)

        # 2) Charger mapping Excel
        map_df = pd.read_excel("mapping.xlsx", header=1)
        map_df = map_df.rename(columns={col: col.strip() for col in map_df.columns})
        assert "NOM" in map_df.columns and "ENTITÉ" in map_df.columns, "Colonnes attendues: 'NOM' et 'ENTITÉ'."

        map_df["NOM"] = map_df["NOM"].astype(str).str.strip()
        map_df["ENTITÉ"] = map_df["ENTITÉ"].astype(str).str.strip()
        map_df["NORM"] = map_df["NOM"].apply(norm_key)

        norm_to_row = dict(zip(map_df["NORM"], map_df[["NOM","ENTITÉ"]].to_records(index=False)))
        all_norm_keys = list(norm_to_row.keys())

        # 3) Traiter les PDF
        if os.path.exists(FINAL_DIR):
            shutil.rmtree(FINAL_DIR)
        os.makedirs(FINAL_DIR, exist_ok=True)

        rows = []  # <= bien défini ici

        for fname in sorted(os.listdir(EXTRACT_DIR)):
            if not fname.lower().endswith(".pdf"):
                continue

            pdf_path = os.path.join(EXTRACT_DIR, fname)
            try:
                doc = fitz.open(pdf_path)
                text0 = doc[0].get_text() if doc.page_count > 0 else ""
            finally:
                try: doc.close()
                except Exception: pass

            prenom_cap, tokens = extract_prenom_and_tokens(text0)

            if not prenom_cap or not tokens:
                entite = "Sans entité"
                target_dir = os.path.join(FINAL_DIR, entite)
                os.makedirs(target_dir, exist_ok=True)
                shutil.move(pdf_path, os.path.join(target_dir, fname))
                rows.append({"Nom de la facture": fname, "Nom détecté": "", "Entité associée": entite,
                             "Dossier de classement": target_dir, "Méthode": "extraction_failed"})
                continue

            candidates = generate_candidates(tokens)
            match_info = None

            # Niveau 1 : exact
            for nom_affiche, cle_excel in candidates:
                key_norm = norm_key(cle_excel)
                if key_norm in norm_to_row:
                    nom_excel_orig, entite = norm_to_row[key_norm]
                    match_info = (nom_affiche, nom_excel_orig, entite, "exact_norm")
                    break

            # Niveau 2 : fuzzy
            if match_info is None:
                for nom_affiche, cle_excel in candidates:
                    key_norm = norm_key(cle_excel)
                    best = process.extractOne(
                        key_norm, all_norm_keys,
                        scorer=fuzz.token_sort_ratio
                    )
                    if best and best[1] >= 90:
                        matched_norm = best[0]
                        nom_excel_orig, entite = norm_to_row[matched_norm]
                        match_info = (nom_affiche, nom_excel_orig, entite, f"fuzzy_{best[1]}")
                        break

            if match_info:
                nom_affiche, nom_excel_orig, entite, method = match_info
                base, ext = os.path.splitext(fname)
                suffix = nom_excel_orig.replace(" ", "_")
                new_name = f"{base}_{suffix}{ext}"

                target_dir = os.path.join(FINAL_DIR, entite)
                os.makedirs(target_dir, exist_ok=True)
                shutil.move(pdf_path, os.path.join(target_dir, new_name))

                rows.append({
                    "Nom de la facture": fname,
                    "Nom détecté": nom_affiche,
                    "Entité associée": entite,
                    "Dossier de classement": target_dir,
                    "Méthode": method
                })
            else:
                entite = "Sans entité"
                target_dir = os.path.join(FINAL_DIR, entite)
                os.makedirs(target_dir, exist_ok=True)
                shutil.move(pdf_path, os.path.join(target_dir, fname))
                rows.append({
                    "Nom de la facture": fname,
                    "Nom détecté": f"{prenom_cap} " + " ".join(t.upper() for t in tokens[1:]),
                    "Entité associée": entite,
                    "Dossier de classement": target_dir,
                    "Méthode": "no_match"
                })

        # 4) Rapport + ZIP unique
        rapport_df = pd.DataFrame(rows, columns=[
            "Nom de la facture", "Nom détecté", "Entité associée", "Dossier de classement", "Méthode"
        ])
        rapport_path = os.path.join(FINAL_DIR, RAPPORT_OUT)
        rapport_df.to_excel(rapport_path, index=False)

        if os.path.exists(ZIP_OUT):
            os.remove(ZIP_OUT)
        shutil.make_archive("resultat_complet", "zip", FINAL_DIR)

        st.success("✅ Terminé !")

        with open(ZIP_OUT, "rb") as f:
            st.download_button("📦 Télécharger tout (factures + rapport)", f, file_name=ZIP_OUT)
