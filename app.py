import os
import re
import zipfile
import shutil
import tempfile
from pathlib import Path

import fitz  # PyMuPDF
import pandas as pd
from unidecode import unidecode
from rapidfuzz import process, fuzz

import streamlit as st


# -----------------------------
# Fonctions utilitaires
# -----------------------------
def smart_capitalize(s: str) -> str:
    """Capitalize prénom(s) avec gestion traits d'union/apostrophes."""
    def cap_token(tok: str) -> str:
        parts = tok.split("-")
        parts = [p[:1].upper() + p[1:].lower() if p else p for p in parts]
        return "-".join(parts)
    return "'".join(cap_token(p) for p in s.split("'"))


def norm_key(s: str) -> str:
    """Normalisation pour comparaison EXACTE/FUZZY."""
    s = unidecode(str(s)).lower()
    s = s.replace("-", " ")
    s = re.sub(r"\s+", " ", s).strip()
    return s


# -----------------------------
# Extraction nom depuis facture (ancien + nouveau format)
# -----------------------------
# Accepte :
# - "MISSION N°... DE M. DUPONT Jean"
# - "MISSION N°... DE Mme DUPONT Jean"
# - "MISSION N°... DE DUPONT Jean"
PATTERN = re.compile(
    r"MISSION\s*N[°º]?\s*\d+\s*DE\s+(.+?)\s*$",
    flags=re.IGNORECASE
)

# Civilités possibles (on les retire si présentes)
CIVILITY_PREFIX = re.compile(
    r"^(?:M\.?|MR\.?|MME|MADAME|MLLE|MADEMOISELLE|ME)\s+",
    flags=re.IGNORECASE
)


def extract_prenom_and_tokens(text_page0: str):
    """
    Retourne (prenom_cap, tokens_apres_civilite)
    tokens_apres_civilite: liste de tokens prénoms+noms
    """
    for line in (ln.strip() for ln in text_page0.splitlines() if ln.strip()):
        m = PATTERN.search(line)
        if not m:
            continue

        tail = m.group(1).strip()
        tail = CIVILITY_PREFIX.sub("", tail).strip()

        tokens = tail.split()
        if not tokens:
            return None, []

        prenom_cap = smart_capitalize(tokens[0])
        return prenom_cap, tokens

    return None, []


def generate_candidates(tokens):
    """
    Génère des hypothèses "NOM Prénom" en testant NOM = 1,2,3 derniers tokens.
    Retourne liste de tuples (cle_affiche, cle_excel)
    """
    if len(tokens) < 1:
        return []

    prenom_cap = smart_capitalize(tokens[0])
    cands = []
    for k in (1, 2, 3):
        if len(tokens) >= 1 + k:
            nom_tokens = tokens[-k:]
            nom_upper = " ".join(t.upper() for t in nom_tokens)
            cle_excel = f"{nom_upper} {prenom_cap}".strip()     # "NOM Prénom"
            cle_affiche = f"{prenom_cap} {nom_upper}".strip()   # "Prénom NOM"
            cands.append((cle_affiche, cle_excel))
    return cands


# -----------------------------
# Interface Streamlit
# -----------------------------
st.title("📂 Classement automatique des factures VIE")

zip_file = st.file_uploader("Uploader le fichier ZIP des factures VIE", type="zip")
excel_file = st.file_uploader("Uploader le fichier Excel Table de référence", type=["xls", "xlsx"])

# Même base / mêmes étapes : bouton de lancement
if st.button("Lancer le traitement"):
    if not zip_file or not excel_file:
        st.error("⚠️ Merci d'uploader le ZIP et l'Excel.")
    else:
        # -----------------------------
        # Paramètres (comme avant) + dossiers temporaires isolés
        # -----------------------------
        # IMPORTANT : on isole chaque exécution dans un dossier unique (évite mélange + évite permissions OneDrive)
        run_dir = Path(tempfile.mkdtemp(prefix="vie_factures_run_"))

        EXTRACT_DIR = run_dir / "factures_extraites"
        FINAL_DIR = run_dir / "factures_classees"
        RAPPORT_OUT = run_dir / "rapport_classement.xlsx"
        ZIP_OUT = run_dir / "resultat_complet.zip"

        input_zip_path = run_dir / "input.zip"
        mapping_xlsx_path = run_dir / "mapping.xlsx"

        # -----------------------------
        # Nettoyage complet avant exécution (nouvelle évolution)
        # -----------------------------
        # Ici, comme run_dir est unique, c'est déjà "propre".
        # Mais on garde la logique : si relance dans la même session, on supprime si ça existe.
        for p in [EXTRACT_DIR, FINAL_DIR]:
            if p.exists():
                shutil.rmtree(p, ignore_errors=True)
        for p in [RAPPORT_OUT, ZIP_OUT]:
            if p.exists():
                try:
                    p.unlink()
                except Exception:
                    pass

        st.info("🧹 Nettoyage terminé. Démarrage du traitement...")

        # -----------------------------
        # 1) Sauvegarder localement les fichiers uploadés
        # -----------------------------
        input_zip_path.write_bytes(zip_file.read())
        mapping_xlsx_path.write_bytes(excel_file.read())

        # -----------------------------
        # 2) Décompresser
        # -----------------------------
        EXTRACT_DIR.mkdir(parents=True, exist_ok=True)
        with zipfile.ZipFile(str(input_zip_path), "r") as zf:
            zf.extractall(str(EXTRACT_DIR))

        # -----------------------------
        # 3) Charger mapping Excel
        # -----------------------------
        try:
            map_df = pd.read_excel(str(mapping_xlsx_path), header=1)
        except PermissionError:
            st.error(
                "❌ Impossible d'ouvrir l'Excel (PermissionError). "
                "Assure-toi qu'il n'est pas ouvert dans Excel et qu'il est bien disponible localement."
            )
            st.stop()

        map_df = map_df.rename(columns={col: col.strip() for col in map_df.columns})
        if "NOM" not in map_df.columns or "ENTITÉ" not in map_df.columns:
            st.error("❌ Colonnes attendues dans l'Excel : 'NOM' et 'ENTITÉ'.")
            st.stop()

        map_df["NOM"] = map_df["NOM"].astype(str).str.strip()
        map_df["ENTITÉ"] = map_df["ENTITÉ"].astype(str).str.strip()

        map_df["NORM"] = map_df["NOM"].apply(norm_key)
        norm_to_row = dict(zip(map_df["NORM"], map_df[["NOM", "ENTITÉ"]].to_records(index=False)))
        all_norm_keys = list(norm_to_row.keys())

        # -----------------------------
        # 4) Traiter les PDF (classement)
        # -----------------------------
        FINAL_DIR.mkdir(parents=True, exist_ok=True)
        rows = []

        pdf_files = [f for f in sorted(EXTRACT_DIR.iterdir()) if f.is_file() and f.suffix.lower() == ".pdf"]
        total = len(pdf_files)

        progress = st.progress(0)
        status = st.empty()

        for i, pdf_path in enumerate(pdf_files, start=1):
            status.write(f"📄 Traitement : {pdf_path.name} ({i}/{total})")

            # Lire la 1ère page
            try:
                doc = fitz.open(str(pdf_path))
                text0 = doc[0].get_text() if doc.page_count > 0 else ""
            finally:
                try:
                    doc.close()
                except Exception:
                    pass

            prenom_cap, tokens = extract_prenom_and_tokens(text0)

            if not prenom_cap or not tokens:
                entite = "Sans entité"
                target_dir = FINAL_DIR / entite
                target_dir.mkdir(parents=True, exist_ok=True)

                shutil.move(str(pdf_path), str(target_dir / pdf_path.name))

                rows.append({
                    "Nom de la facture": pdf_path.name,
                    "Nom détecté": "",
                    "Entité associée": entite,
                    "Dossier de classement": str(target_dir),
                    "Méthode": "extraction_failed"
                })
                progress.progress(int(i / max(total, 1) * 100))
                continue

            candidates = generate_candidates(tokens)

            match_info = None  # (nom_affiche, nom_excel_original, entite, method)

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

                base = pdf_path.stem
                ext = pdf_path.suffix
                suffix = str(nom_excel_orig).replace(" ", "_")
                new_name = f"{base}_{suffix}{ext}"

                target_dir = FINAL_DIR / entite
                target_dir.mkdir(parents=True, exist_ok=True)
                shutil.move(str(pdf_path), str(target_dir / new_name))

                rows.append({
                    "Nom de la facture": pdf_path.name,
                    "Nom détecté": nom_affiche,
                    "Entité associée": entite,
                    "Dossier de classement": str(target_dir),
                    "Méthode": method
                })
            else:
                entite = "Sans entité"
                target_dir = FINAL_DIR / entite
                target_dir.mkdir(parents=True, exist_ok=True)
                shutil.move(str(pdf_path), str(target_dir / pdf_path.name))

                rows.append({
                    "Nom de la facture": pdf_path.name,
                    "Nom détecté": f"{prenom_cap} " + " ".join(t.upper() for t in tokens[1:]),
                    "Entité associée": entite,
                    "Dossier de classement": str(target_dir),
                    "Méthode": "no_match"
                })

            progress.progress(int(i / max(total, 1) * 100))

        status.write("🧾 Génération du rapport...")

        # -----------------------------
        # 5) Rapport + ZIP unique (comme avant, mais propre)
        # -----------------------------
        rapport_df = pd.DataFrame(rows, columns=[
            "Nom de la facture", "Nom détecté", "Entité associée", "Dossier de classement", "Méthode"
        ])
        rapport_df.to_excel(str(RAPPORT_OUT), index=False)

        # Copier le rapport DANS le dossier final (comme ton ancien script)
        shutil.copy2(str(RAPPORT_OUT), str(FINAL_DIR / "rapport_classement.xlsx"))

        # Créer un ZIP global contenant le dossier final (factures + rapport)
        base_zip = str(run_dir / "resultat_complet")  # make_archive ajoute .zip
        if os.path.exists(base_zip + ".zip"):
            os.remove(base_zip + ".zip")
        shutil.make_archive(base_zip, "zip", str(FINAL_DIR))

        st.success("✅ Terminé !")

        # -----------------------------
        # 6) Téléchargement (ZIP final)
        # -----------------------------
        with open(base_zip + ".zip", "rb") as f:
            st.download_button(
                "📦 Télécharger tout (factures + rapport)",
                f,
                file_name="resultat_complet.zip"
            )

        st.caption(f"Dossier de travail temporaire : {run_dir}")
