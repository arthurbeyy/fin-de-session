from pathlib import Path
import sys
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
    p.add_run("Je soussignée, Élodie Laurent, représentante légale de l'organisme fictif ").bold = False
    p.add_run("Forma Exemple SAS").bold = True
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
    p = doc.add_paragraph(f"Fait à Lyon, le 12/09/2026")
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p = doc.add_paragraph("Élodie Laurent\nResponsable de l'organisme\nSignature")
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p.paragraph_format.space_before = Pt(14)
    footer = sec.footer.paragraphs[0]
    footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    fr = footer.add_run("Forma Exemple SAS - Démonstration fictive - SIREN 000 000 000")
    fr.font.name = "Arial"; fr.font.size = Pt(8); fr.font.color.rgb = RGBColor(107,114,128)
    return doc

def main():
    OUT.mkdir(parents=True, exist_ok=True)
    wb = load_workbook(INPUT, data_only=True)
    ws = wb["Participants"]
    headers = [c.value for c in ws[1]]
    required = {"Identifiant","Civilité","Prénom","Nom","Entreprise","Intitulé formation","Date début","Date fin","Heures prévues","Heures réalisées","Lieu","Statut"}
    missing = required - set(headers)
    if missing: raise ValueError(f"Colonnes manquantes: {sorted(missing)}")
    generated = []
    for values in ws.iter_rows(min_row=2, values_only=True):
        row = dict(zip(headers, values))
        if row["Statut"] != "Terminé": continue
        if row["Heures réalisées"] is None or row["Heures réalisées"] <= 0: raise ValueError(f"Durée invalide pour {row['Identifiant']}")
        if row["Heures réalisées"] > row["Heures prévues"]: raise ValueError(f"Durée réalisée supérieure à la durée prévue pour {row['Identifiant']}")
        filename = f"certificat-{row['Identifiant']}-{row['Nom'].lower()}"
        path = OUT / f"{filename}.docx"
        build_certificate(row).save(path)
        generated.append(path)
    print(f"{len(generated)} certificats DOCX générés")
    for p in generated: print(p)

if __name__ == "__main__": main()
