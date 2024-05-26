import openpyxl
import docx
from docx.shared import Inches, Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT


# Charger le fichier Excel
workbook = openpyxl.load_workbook('C:/Users/nelson/Desktop/projects/python/studentAttestations/ASG_RESULTS_TEST.xlsx')
sheet = workbook.active

# Créer le document Word
document = docx.Document()


# Access the header of the first section
header = document.sections[0].header



# Create a table in the header
htable = header.add_table(1, 3, Inches(6))
htable.style.space_before = Pt(0)  # Set the space before the table to 0

htab_cells = htable.rows[0].cells

# Add the left-aligned structured text to the first cell
ht0 = htab_cells[0].add_paragraph()
ht0.paragraph_format.space_before = Pt(0)  # Remove extra space before paragraph
ht0.text = "REPUBLIQUE DU CAMEROUN"
ht0.style.font.size = Pt(8)
ht0.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
ht0.paragraph_format.space_after = Pt(0)  # Remove extra space after paragraph

ht0 = htab_cells[0].add_paragraph()
ht0.text = "Paix – Travail – Patrie"
ht0.style.font.size = Pt(8)
ht0.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
ht0.paragraph_format.space_after = Pt(0)  # Remove extra space after paragraph

ht0 = htab_cells[0].add_paragraph()
ht0.text = "-----"
ht0.style.font.size = Pt(8)
ht0.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
ht0.paragraph_format.space_after = Pt(0)  # Remove extra space after paragraph

ht0 = htab_cells[0].add_paragraph()
ht0.text = "MINISTERE DE LA SANTE PUBLIQUE"
ht0.style.font.size = Pt(8)
ht0.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
ht0.paragraph_format.space_after = Pt(0)  # Remove extra space after paragraph

ht0 = htab_cells[0].add_paragraph()
ht0.text = "-----"
ht0.style.font.size = Pt(8)
ht0.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
ht0.paragraph_format.space_after = Pt(0)  # Remove extra space after paragraph

ht0 = htab_cells[0].add_paragraph()
ht0.text = "DIRECTION DES RESSOURCES HUMAINES"
ht0.style.font.size = Pt(8)
ht0.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
ht0.paragraph_format.space_after = Pt(0)  # Remove extra space after paragraph

ht0 = htab_cells[0].add_paragraph()
ht0.text = "-----"
ht0.style.font.size = Pt(8)
ht0.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
ht0.paragraph_format.space_after = Pt(0)  # Remove extra space after paragraph

ht0 = htab_cells[0].add_paragraph()
ht0.text = "ÉCOLE DES SCIENCES MEDICALES ET D'APPLICATION MARIE ZAMBO"
ht0.style.font.size = Pt(8)
ht0.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
ht0.paragraph_format.space_after = Pt(0)  # Remove extra space after paragraph

ht0 = htab_cells[0].add_paragraph()
ht0.text = "BP :"
ht0.style.font.size = Pt(8)
ht0.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
ht0.paragraph_format.space_after = Pt(0)  # Remove extra space after paragraph

ht0 = htab_cells[0].add_paragraph()
ht0.text = "Tel : +237 655 666 140"
ht0.style.font.size = Pt(8)
ht0.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
ht0.paragraph_format.space_after = Pt(0)  # Remove extra space after paragraph


# Add the image to the second cell
kh = htab_cells[1].add_paragraph()
kh.add_run().add_picture('C:/Users/nelson/Desktop/projects/python/studentAttestations/esmaLogo.png', width=Inches(1))
kh.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

# Add the RIGHT-aligned structured text to the third cell
ht1 = htab_cells[2].add_paragraph()
ht1.paragraph_format.space_before = Pt(0)  # Remove extra space before paragraph
ht1.text = "REPUBLIC OF CAMEROON"
ht1.style.font.size = Pt(8)
ht1.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
ht1.paragraph_format.space_after = Pt(0)  # Remove extra space after paragraph

ht1 = htab_cells[2].add_paragraph()
ht1.text = "Peace - Work - Fatherland"
ht1.style.font.size = Pt(8)
ht1.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
ht1.paragraph_format.space_after = Pt(0)  # Remove extra space after paragraph

ht1 = htab_cells[2].add_paragraph()
ht1.text = "-----"
ht1.style.font.size = Pt(8)
ht1.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
ht1.paragraph_format.space_after = Pt(0)  # Remove extra space after paragraph

ht1 = htab_cells[2].add_paragraph()
ht1.text = "MINISTRY OF PUBLIC HEALTH"
ht1.style.font.size = Pt(8)
ht1.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
ht1.paragraph_format.space_after = Pt(0)  # Remove extra space after paragraph

ht1 = htab_cells[2].add_paragraph()
ht1.text = "-----"
ht1.style.font.size = Pt(8)
ht1.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
ht1.paragraph_format.space_after = Pt(0)  # Remove extra space after paragraph

ht1 = htab_cells[2].add_paragraph()
ht1.text = "HUMAN RESOURCES DEPARTMENT"
ht1.style.font.size = Pt(8)
ht1.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
ht1.paragraph_format.space_after = Pt(0)  # Remove extra space after paragraph

ht1 = htab_cells[2].add_paragraph()
ht1.text = "-----"
ht1.style.font.size = Pt(8)
ht1.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
ht1.paragraph_format.space_after = Pt(0)  # Remove extra space after paragraph

ht1 = htab_cells[2].add_paragraph()
ht1.text = "SCHOOL OF MEDICAL AND APPLIED SCIENCES MARIE ZAMBO"
ht1.style.font.size = Pt(8)
ht1.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
ht1.paragraph_format.space_after = Pt(0)  # Remove extra space after paragraph

ht1 = htab_cells[2].add_paragraph()
ht1.text = "PO Box:"
ht1.style.font.size = Pt(8)
ht1.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
ht1.paragraph_format.space_after = Pt(0)  # Remove extra space after paragraph

ht1 = htab_cells[2].add_paragraph()
ht1.text = "Tel: +237 655 666 140"
ht1.style.font.size = Pt(8)
ht1.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
ht1.paragraph_format.space_after = Pt(0)  # Remove extra space after paragraph



# Load the Word document
# Find the last used row
last_used_row = 0
for row in range(1, sheet.max_row + 1):
    if any(sheet.cell(row=row, column=col).value for col in range(1, sheet.max_column + 1)):
        last_used_row = row

# Parcourir les lignes du fichier Excel
for row in range(2, last_used_row + 1):
    # Récupérer les informations de l'étudiant
    nom_prenom = sheet.cell(row=row, column=1).value
    matricule = sheet.cell(row=row, column=2).value
    ecole = sheet.cell(row=row, column=3).value
    sexe = sheet.cell(row=row, column=4).value
    date_naissance = sheet.cell(row=row, column=5).value
    lieu_naissance = sheet.cell(row=row, column=6).value
    nbre_modules_valides = sheet.cell(row=row, column=7).value
    pct_validation = sheet.cell(row=row, column=8).value
    total_points = sheet.cell(row=row, column=9).value
    moyenne = sheet.cell(row=row, column=10).value
    rang = sheet.cell(row=row, column=11).value
    observations = sheet.cell(row=row, column=8).value

    # Ajouter une nouvelle page au document Word
    document.add_heading(f"Résultats de {nom_prenom}", 0)

    # Ajouter les informations de l'étudiant au document Word
    document.add_paragraph(f"Nom et Prénom: {nom_prenom}")
    document.add_paragraph(f"Matricule: {matricule}")
    document.add_paragraph(f"École: {ecole}")
    document.add_paragraph(f"Sexe: {sexe}")
    document.add_paragraph(f"Date de Naissance: {date_naissance}")
    document.add_paragraph(f"Lieu de Naissance: {lieu_naissance}")
    document.add_paragraph(f"Nombre de Modules Validés: {nbre_modules_valides}")
    document.add_paragraph(f"Pourcentage de Validation: {pct_validation}%")
    document.add_paragraph(f"Total des Points: {total_points}")
    document.add_paragraph(f"Moyenne: {moyenne}/20")
    document.add_paragraph(f"Rang: {rang}")
    document.add_paragraph(f"Observations: {observations}")

    # Ajouter une nouvelle page
    document.add_page_break()

# Enregistrer le document Word
document.save('C:/Users/nelson/Desktop/projects/python/studentAttestations/results/resultats_etudiants.docx')