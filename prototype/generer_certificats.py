from pathlib import Path
from datetime import date
import re
import shutil
import subprocess
import sys
import unicodedata
from openpyxl import load_workbook
from docx import Document
from docx.shared import Cm, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

ROOT = Path(__file__).resolve().parent
INPUT = Path(sys.argv[1]) if len(sys.argv) > 1 else ROOT / "donnees-participants.xlsx"
OUT = Path(sys.argv[2]) if len(sys.argv) > 2 else ROOT / "certificats"

# ---------------------------------------------------------------------------
# Les trois lignes ci-dessous sont les seules a changer pour un nouveau client.
# ---------------------------------------------------------------------------
ORGANISME = "Forma Exemple SAS"          # nom de l'organisme de formation
SIGNATAIRE = "Élodie Laurent"            # personne qui signe les certificats
FONCTION_SIGNATAIRE = "Responsable de l'organisme"
VILLE_SIGNATURE = "Lyon"                 # ville indiquee au dessus de la signature
PIED_DE_PAGE = "Forma Exemple SAS - Démonstration fictive - SIREN 000 000 000"

def shade(cell, fill):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), fill)
    tc_pr.append(shd)

def set_cell_margins(cell, top=120, start=140, bottom=120, end=140):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcMar = tcPr.first_child_found_in("w:tcMar")
    if tcMar is None:
        tcMar = OxmlElement("w:tcMar")
        tcPr.append(tcMar)
    for m, v in (("top", top), ("start", start), ("bottom", bottom), ("end", end)):
        node = OxmlElement(f"w:{m}")
        node.set(qn("w:w"), str(v)); node.set(qn("w:type"), "dxa")
        tcMar.append(node)

def remove_paragraph_borders(paragraph):
    ppr = paragraph._p.get_or_add_pPr()
    border = ppr.find(qn("w:pBdr"))
    if border is not None:
        ppr.remove(border)

def format_date(value):
    return value.strftime("%d/%m/%Y") if hasattr(value, "strftime") else str(value)

def add_field(table, label, value, alt=False):
    cells = table.add_row().cells
    cells[0].text = label
    cells[1].text = str(value)
    for c in cells:
        c.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        set_cell_margins(c)
        if alt: shade(c, "F3F6F8")
    cells[0].paragraphs[0].runs[0].bold = True
    cells[0].paragraphs[0].runs[0].font.color.rgb = RGBColor(23, 50, 77)

def build_certificate(r):
    doc = Document()
    sec = doc.sections[0]
    sec.page_width = Cm(21.0); sec.page_height = Cm(29.7)
    sec.top_margin = Cm(1.8); sec.bottom_margin = Cm(1.8)
    sec.left_margin = Cm(2.2); sec.right_margin = Cm(2.2)
    styles = doc.styles
    styles["Normal"].font.name = "Arial"; styles["Normal"].font.size = Pt(10.5)
    title_ppr = styles["Title"].element.get_or_add_pPr()
    title_border = title_ppr.find(qn("w:pBdr"))
    if title_border is not None: title_ppr.remove(title_border)
    title = doc.add_paragraph(style="Title")
    remove_paragraph_borders(title)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title.add_run("Certificat de réalisation")
    run.font.name = "Arial"; run.font.size = Pt(22); run.font.bold = True; run.font.color.rgb = RGBColor(0,0,0)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    rr = p.add_run("Action de formation professionnelle")
    rr.font.name = "Arial"; rr.font.size = Pt(11); rr.font.color.rgb = RGBColor(75,85,99)
    doc.add_paragraph("")
    p = doc.add_paragraph()
    p.add_run(f"Je soussignée, {SIGNATAIRE}, représentante légale de l'organisme fictif ").bold = False
    p.add_run(ORGANISME).bold = True
    p.add_run(", atteste que la personne désignée ci-dessous a suivi l'action de formation indiquée.")
    table = doc.add_table(rows=0, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    table.columns[0].width = Cm(5.2); table.columns[1].width = Cm(10.5)
    fields = [
      ("Bénéficiaire", f"{r['Civilité']} {r['Prénom']} {r['Nom']}"),
      ("Entreprise", r["Entreprise"]),
      ("Action", r["Intitulé formation"]),
      ("Période", f"du {format_date(r['Date début'])} au {format_date(r['Date fin'])}"),
      ("Durée réalisée", f"{r['Heures réalisées']:g} heures"),
      ("Lieu", r["Lieu"]),
      ("Nature", "Action de formation"),
    ]
    for i, (lab, val) in enumerate(fields): add_field(table, lab, val, i % 2 == 1)
    doc.add_paragraph("")
    p = doc.add_paragraph("Le présent certificat est établi sur la base des éléments de suivi et justificatifs conservés par l'organisme de formation.")
    p.paragraph_format.space_after = Pt(18)
    p = doc.add_paragraph(f"Fait à {VILLE_SIGNATURE}, le {date.today().strftime('%d/%m/%Y')}")
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p = doc.add_paragraph(f"{SIGNATAIRE}\n{FONCTION_SIGNATAIRE}\nSignature")
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p.paragraph_format.space_before = Pt(14)
    footer = sec.footer.paragraphs[0]
    footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    fr = footer.add_run(PIED_DE_PAGE)
    fr.font.name = "Arial"; fr.font.size = Pt(8); fr.font.color.rgb = RGBColor(107,114,128)
    return doc

COLONNES_REQUISES = ["Identifiant","Civilité","Prénom","Nom","Entreprise","Intitulé formation",
                     "Date début","Date fin","Heures prévues","Heures réalisées","Lieu","Statut"]

def nom_de_fichier(texte, minuscules=True):
    """Transforme un nom en morceau de nom de fichier sûr: pas d'accent, pas d'espace."""
    texte = unicodedata.normalize("NFKD", str(texte))
    texte = "".join(c for c in texte if not unicodedata.combining(c))
    texte = re.sub(r"[^A-Za-z0-9]+", "-", texte).strip("-")
    if minuscules: texte = texte.lower()
    return texte or "sans-nom"

def nombre(valeur):
    """Renvoie la valeur en nombre, ou None si ce n'est pas un nombre."""
    if isinstance(valeur, bool) or valeur is None: return None
    if isinstance(valeur, (int, float)): return float(valeur)
    try: return float(str(valeur).replace(",", ".").strip())
    except ValueError: return None

def probleme_sur_la_ligne(row):
    """Renvoie la raison d'écarter la ligne, ou None si la ligne est bonne."""
    for col in ("Identifiant","Prénom","Nom","Intitulé formation","Date début","Date fin"):
        if row.get(col) in (None, ""):
            return f"la colonne « {col} » est vide"
    prevues = nombre(row.get("Heures prévues"))
    realisees = nombre(row.get("Heures réalisées"))
    if realisees is None: return "la durée réalisée n'est pas un nombre"
    if realisees <= 0: return "la durée réalisée est nulle ou négative"
    if prevues is None: return "la durée prévue n'est pas un nombre"
    if realisees > prevues:
        return f"la durée réalisée ({realisees:g} h) dépasse la durée prévue ({prevues:g} h)"
    return None

def chercher_libreoffice():
    """Cherche LibreOffice sur Mac, Windows et Linux. Renvoie le chemin ou None."""
    for nom in ("soffice", "libreoffice", "soffice.exe"):
        trouve = shutil.which(nom)
        if trouve: return trouve
    pistes = [
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    ]
    for piste in pistes:
        if Path(piste).exists(): return piste
    return None

def convertir_en_pdf(fichiers_word, dossier):
    """Fabrique un PDF à côté de chaque Word. Renvoie (liste des PDF, message)."""
    if not fichiers_word: return [], "aucun Word à convertir"
    soffice = chercher_libreoffice()
    if soffice:
        profil = dossier / "_libreoffice_temporaire"
        commande = [soffice, f"-env:UserInstallation=file://{profil.resolve()}",
                    "--headless", "--convert-to", "pdf", "--outdir", str(dossier)]
        commande += [str(f) for f in fichiers_word]
        try:
            subprocess.run(commande, check=True, capture_output=True, timeout=600)
        except subprocess.TimeoutExpired:
            return [], "LibreOffice n'a pas répondu dans le temps imparti"
        except subprocess.CalledProcessError as erreur:
            return [], f"LibreOffice a refusé la conversion ({erreur.returncode})"
        finally:
            shutil.rmtree(profil, ignore_errors=True)
        pdfs = [f.with_suffix(".pdf") for f in fichiers_word if f.with_suffix(".pdf").exists()]
        return pdfs, f"conversion par LibreOffice, {len(pdfs)} PDF sur {len(fichiers_word)}"
    try:
        from docx2pdf import convert
    except ImportError:
        return [], ("ni LibreOffice ni Word n'ont été trouvés sur cet ordinateur, "
                    "les Word sont prêts mais les PDF n'ont pas pu être fabriqués")
    try:
        convert(str(dossier))
    except Exception as erreur:
        return [], f"Word a refusé la conversion ({erreur})"
    pdfs = [f.with_suffix(".pdf") for f in fichiers_word if f.with_suffix(".pdf").exists()]
    return pdfs, f"conversion par Microsoft Word, {len(pdfs)} PDF sur {len(fichiers_word)}"

def ecrire_journal(dossier, produits, non_termines, anomalies, message_pdf):
    lignes = [
        "Journal de génération des certificats",
        f"Date : {date.today().strftime('%d/%m/%Y')}",
        f"Fichier lu : {INPUT}",
        "",
        f"Certificats produits : {len(produits)}",
    ]
    for nom in produits: lignes.append(f"  {nom}")
    lignes += ["", f"PDF : {message_pdf}", "",
               f"Participants non terminés, ignorés volontairement : {len(non_termines)}"]
    for ligne, identifiant, raison in non_termines:
        lignes.append(f"  ligne {ligne} ({identifiant}) : {raison}")
    lignes += ["", f"Lignes à corriger : {len(anomalies)}"]
    for ligne, identifiant, raison in anomalies:
        lignes.append(f"  ligne {ligne} ({identifiant}) : {raison}")
    if anomalies:
        lignes += ["", "Corrigez ces lignes dans le tableur puis relancez le programme. "
                       "Les certificats déjà produits seront simplement réécrits."]
    else:
        lignes += ["", "Aucune anomalie. Tous les participants terminés ont leur certificat."]
    (dossier / "journal.txt").write_text("\n".join(lignes), encoding="utf-8")

def main():
    if not INPUT.exists():
        print(f"Fichier introuvable : {INPUT}")
        return 1
    OUT.mkdir(parents=True, exist_ok=True)
    wb = load_workbook(INPUT, data_only=True)
    if "Participants" not in wb.sheetnames:
        print("Le classeur ne contient pas d'onglet « Participants ».")
        return 1
    ws = wb["Participants"]
    headers = [c.value for c in ws[1]]
    manquantes = [c for c in COLONNES_REQUISES if c not in headers]
    if manquantes:
        print("Le tableur ne contient pas les colonnes suivantes, indispensables au certificat :")
        for c in manquantes: print(f"  {c}")
        print("Ajoutez ces colonnes dans l'onglet « Participants », sans changer l'orthographe.")
        return 1

    produits, non_termines, anomalies = [], [], []
    for numero, values in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        row = dict(zip(headers, values))
        if all(v in (None, "") for v in values): continue
        identifiant = row.get("Identifiant") or "sans identifiant"
        statut = (row.get("Statut") or "").strip()
        if statut != "Terminé":
            non_termines.append((numero, identifiant, f"statut « {statut or 'vide'} »"))
            continue
        raison = probleme_sur_la_ligne(row)
        if raison:
            anomalies.append((numero, identifiant, raison))
            continue
        row["Heures réalisées"] = nombre(row["Heures réalisées"])
        chemin = OUT / (f"certificat-{nom_de_fichier(identifiant, minuscules=False)}"
                        f"-{nom_de_fichier(row['Nom'])}.docx")
        try:
            build_certificate(row).save(chemin)
        except Exception as erreur:
            anomalies.append((numero, identifiant, f"le document n'a pas pu être écrit ({erreur})"))
            continue
        produits.append(chemin)

    pdfs, message_pdf = convertir_en_pdf(produits, OUT)
    ecrire_journal(OUT, [p.name for p in produits], non_termines, anomalies, message_pdf)

    print(f"{len(produits)} certificats Word et {len(pdfs)} PDF dans {OUT}")
    if anomalies:
        print(f"{len(anomalies)} ligne(s) à corriger, le détail est dans journal.txt")
    if produits and not pdfs:
        print(f"Aucun PDF : {message_pdf}")
    return 0

if __name__ == "__main__": sys.exit(main())
