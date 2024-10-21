import logging
import json
from datetime import datetime
import pandas as pd
from docxtpl import DocxTemplate
from app.core.config import settings
import os
import unicodedata
import math
from docx.shared import Pt
from docx.oxml import OxmlElement

# Configure logging
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# Fonction pour lire la configuration des ECTS depuis un fichier JSON
def read_ects_config():
    with open(settings.ECTS_JSON_PATH, 'r') as file:
        data = json.load(file)
    return data

# Fonction pour normaliser une chaîne de caractères
def normalize_string(s):
    if not isinstance(s, str):
        s = str(s)
    return ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn').lower()

# Modifiez la fonction extract_grades_and_coefficients
def extract_grades_and_coefficients(grade_str):
    grades_coefficients = []
    special_case = None

    if "Validé ( - ASE)" in grade_str:
        special_case = "Validé"
    elif "Non Validé ( - ASE)" in grade_str:
        special_case = "Non Validé"
    elif "(CCHM)" in grade_str:
        special_case = grade_str.replace("(CCHM)", "").strip()
    elif not grade_str.strip() or "Validé" in grade_str:
        return grades_coefficients, None

    if special_case:
        return grades_coefficients, special_case

    parts = grade_str.split(" - ")
    for part in parts:
        if "Absent au devoir" in part:
            continue
        try:
            if "(" in part:
                grade_part, coefficient_part = part.rsplit("(", 1)
                coefficient_part = coefficient_part.rstrip(")")
            else:
                grade_part = part
                coefficient_part = "1.0"
            grade = grade_part.replace(",", ".").strip()
            coefficient = coefficient_part.replace(",", ".").strip()
            
            if grade.lower() == 'cchm':
                grade = '1'
            elif not grade or float(grade) == 0:
                continue
            grades_coefficients.append((float(grade), float(coefficient)))
        except ValueError:
            continue

    return grades_coefficients, None

# Modifiez la fonction calculate_weighted_average pour gérer les cas spéciaux
def calculate_weighted_average(notes, ects):
    if not notes or not ects:
        return 0.0

    # Filtrer les notes et les ects où ects est zéro
    filtered_notes = [note for note, ect in zip(notes, ects) if ect != 0]
    filtered_ects = [ect for ect in ects if ect != 0]

    # Si aucune note valide ne reste après filtrage, retourner 0.0
    if not filtered_notes or not filtered_ects:
        return 0.0

    total_grade = sum(note * ect for note, ect in zip(filtered_notes, filtered_ects))
    total_ects = sum(filtered_ects)
    
    return total_grade / total_ects if total_ects != 0 else 0.0

# Fonction pour générer les placeholders pour le document Word
def generate_placeholders(titles_row, case_config, student_data, current_date, ects_data):
    logger.debug(f"Received ECTS data: {ects_data}")
    placeholders = {
        "nomApprenant": student_data["Nom"],
        "etendugroupe": student_data["Étendu Groupe"],
        "dateNaissance": student_data["Date de Naissance"],
        "codegroupe": student_data["Code Groupe"],
        "groupe": student_data["Nom Groupe"],
        "campus": student_data["Nom Site"],
        "justifiee": student_data["ABS justifiées"],
        "injustifiee": student_data["ABS injustifiées"],
        "retard": student_data["Retards"],
        "datedujour": current_date,
        "appreciations": student_data["Appreciations"],
        "CodeApprenant": student_data["CodeApprenant"]
    }

    # Mise à jour des placeholders en fonction de la clé du cas M1, M2
    if case_config["key"] == "M1_S1":
        placeholders.update({
            "UE1_Title": titles_row[0],
            "matiere1": titles_row[1],
            "matiere2": titles_row[2],
            "matiere3": titles_row[3],
            "UE2_Title": titles_row[4],
            "matiere4": titles_row[5],
            "UE3_Title": titles_row[6],
            "matiere5": titles_row[7],
            "matiere6": titles_row[8],
            "UE4_Title": titles_row[9],
            "matiere7": titles_row[10],
            "matiere8": titles_row[11],
            "matiere9": titles_row[12],
            "matiere10": titles_row[13],
            "matiere11": titles_row[14],
            "matiere12": titles_row[15],
            "UESPE_Title": titles_row[16],
            "matiere13": titles_row[17],
            "matiere14": titles_row[18],
            "matiere15": titles_row[19],
        })
    elif case_config["key"] == "M1_S2":
        placeholders.update({
            "UE1_Title": titles_row[0],
            "matiere1": titles_row[1],
            "matiere2": titles_row[2],
            "matiere3": titles_row[3],
            "UE2_Title": titles_row[4],
            "matiere4": titles_row[5],
            "matiere5": titles_row[6],
            "UE3_Title": titles_row[7],
            "matiere6": titles_row[8],
            "matiere7": titles_row[9],
            "matiere8": titles_row[10],
            "matiere9": titles_row[11],
            "matiere10": titles_row[12],
            "matiere11": titles_row[13],
            "matiere12": titles_row[14],
            "UESPE_Title": titles_row[15],
            "matiere13": titles_row[16],
            "matiere14": titles_row[17],
            "matiere15": titles_row[18],
            "matiere16": titles_row[19],
        })
    elif case_config["key"] == "M2_S3_MAGI":
        placeholders.update({
            "UE1_Title": titles_row[0],
            "matiere1": titles_row[1],
            "matiere2": titles_row[2],
            "UE2_Title": titles_row[3],
            "matiere3": titles_row[4],
            "UE3_Title": titles_row[5],
            "matiere4": titles_row[6],
            "matiere5": titles_row[7],
            "matiere6": titles_row[8],
            "matiere7": titles_row[9],
            "matiere8": titles_row[10],
            "matiere9": titles_row[11],
            "UESPE_Title": titles_row[12],
            "matiere10": titles_row[13],
            "matiere11": titles_row[14],
            "matiere12": titles_row[15],
            "matiere13": titles_row[16],
        })
    elif case_config["key"] == "M2_S3_MEFIM":
        placeholders.update({
            "UE1_Title": titles_row[0],
            "matiere1": titles_row[1],
            "matiere2": titles_row[2],
            "UE2_Title": titles_row[3],
            "matiere3": titles_row[4],
            "UE3_Title": titles_row[5],
            "matiere4": titles_row[6],
            "matiere5": titles_row[7],
            "matiere6": titles_row[8],
            "matiere7": titles_row[9],
            "matiere8": titles_row[10],
            "matiere9": titles_row[11],
            "UESPE_Title": titles_row[12],
            "matiere10": titles_row[13],
            "matiere11": titles_row[14],
            "matiere12": titles_row[15],
            "matiere13": titles_row[16],
        })
    elif case_config["key"] == "M2_S3_MAPI":
        placeholders.update({
            "UE1_Title": titles_row[0],
            "matiere1": titles_row[1],
            "matiere2": titles_row[2],
            "UE2_Title": titles_row[3],
            "matiere3": titles_row[4],
            "UE3_Title": titles_row[5],
            "matiere4": titles_row[6],
            "matiere5": titles_row[7],
            "matiere6": titles_row[8],
            "matiere7": titles_row[9],
            "matiere8": titles_row[10],
            "matiere9": titles_row[11],
            "UESPE_Title": titles_row[12],
            "matiere10": titles_row[13],
            "matiere11": titles_row[14],
            "matiere12": titles_row[15],
            "matiere13": titles_row[16],
            "matiere14": titles_row[17],
        })
    elif case_config["key"] == "M2_S4":
        placeholders.update({
            "UE1_Title": titles_row[0],
            "matiere1": titles_row[1],
            "UE2_Title": titles_row[2],
            "matiere2": titles_row[3],
            "matiere3": titles_row[4],
            "UE3_Title": titles_row[5],
            "matiere4": titles_row[6],
            "matiere5": titles_row[7],
            "matiere6": titles_row[8],
            "matiere7": titles_row[9],
            "matiere8": titles_row[10],
            "UESPE_Title": titles_row[11],
            "matiere9": titles_row[12],
            "matiere10": titles_row[13],
            "matiere11": titles_row[14],
        })

    # Mise à jour des placeholders en fonction de la clé du cas BG ALT 1 2 3 4 5 6
    elif case_config["key"] == "BG_ALT_1": 
        placeholders.update({
            "UE1_Title": titles_row[0],
            "matiere1": titles_row[1],
            "matiere2": titles_row[2],
            "matiere3": titles_row[3],
            "UE2_Title": titles_row[4],
            "matiere4": titles_row[5],
            "matiere5": titles_row[6],
            "UE3_Title": titles_row[7],
            "matiere6": titles_row[8],
            "UE4_Title": titles_row[9],
            "matiere7": titles_row[10],
            "matiere8": titles_row[11],
            "matiere9": titles_row[12],
            "matiere10": titles_row[13],
            "matiere11": titles_row[14],
            "matiere12": titles_row[15],
            "matiere13": titles_row[16],
            "matiere14": titles_row[17],
        })
    elif case_config["key"] == "BG_ALT_2": 
        placeholders.update({
            "UE1_Title": titles_row[0],
            "matiere1": titles_row[1],
            "matiere2": titles_row[2],
            "matiere3": titles_row[3],
            "matiere4": titles_row[4],
            "UE2_Title": titles_row[5],
            "matiere5": titles_row[6],
            "matiere6": titles_row[7],
            "matiere7": titles_row[8],
            "UE3_Title": titles_row[9],
            "matiere8": titles_row[10],
            "UE4_Title": titles_row[11],
            "matiere9": titles_row[12],
            "matiere10": titles_row[13],
            "matiere11": titles_row[14],
            "matiere12": titles_row[15],
            "matiere13": titles_row[16],
            "matiere14": titles_row[17],
            "matiere15": titles_row[18],
        })
    elif case_config["key"] == "BG_ALT_3": 
        placeholders.update({
            "UE1_Title": titles_row[0],
            "matiere1": titles_row[1],
            "matiere2": titles_row[2],
            "matiere3": titles_row[3],
            "matiere4": titles_row[4],
            "matiere5": titles_row[5],
            "UE2_Title": titles_row[6],
            "matiere6": titles_row[7],
            "matiere7": titles_row[8],
            "UE3_Title": titles_row[9],
            "matiere8": titles_row[10],
            "UE4_Title": titles_row[11],
            "matiere9": titles_row[12],
            "matiere10": titles_row[13],
            "matiere11": titles_row[14],
            "matiere12": titles_row[15],
            "matiere13": titles_row[16],
        })
    elif case_config["key"] == "BG_ALT_4": 
        placeholders.update({
            "UE1_Title": titles_row[0],
            "matiere1": titles_row[1],
            "matiere2": titles_row[2],
            "matiere3": titles_row[3],
            "UE2_Title": titles_row[4],
            "matiere4": titles_row[5],
            "matiere5": titles_row[6],
            "matiere6": titles_row[7],
            "matiere7": titles_row[8],
            "UE3_Title": titles_row[9],
            "matiere8": titles_row[10],
            "UE4_Title": titles_row[11],
            "matiere9": titles_row[12],
            "matiere10": titles_row[13],
            "matiere11": titles_row[14],
            "matiere12": titles_row[15],
            "matiere13": titles_row[16],
        })
    elif case_config["key"] == "BG_ALT_5": 
        placeholders.update({
            "UE1_Title": titles_row[0],
            "matiere1": titles_row[1],
            "matiere2": titles_row[2],
            "matiere3": titles_row[3],
            "UE2_Title": titles_row[4],
            "matiere4": titles_row[5],
            "matiere5": titles_row[6],
            "matiere6": titles_row[7],
            "UE3_Title": titles_row[8],
            "matiere7": titles_row[9],
            "UE4_Title": titles_row[10],
            "matiere8": titles_row[11],
            "matiere9": titles_row[12],
            "matiere10": titles_row[13],
            "matiere11": titles_row[14],
            "matiere12": titles_row[15],
            "matiere13": titles_row[16],
            "matiere14": titles_row[17],
        })
    elif case_config["key"] == "BG_ALT_6": 
        placeholders.update({
            "UE1_Title": titles_row[0],
            "matiere1": titles_row[1],
            "matiere2": titles_row[2],
            "UE2_Title": titles_row[3],
            "matiere3": titles_row[4],
            "matiere4": titles_row[5],
            "UE3_Title": titles_row[6],
            "matiere5": titles_row[7],
            "matiere6": titles_row[8],
            "UE4_Title": titles_row[9],
            "matiere7": titles_row[10],
            "matiere8": titles_row[11],
            "matiere9": titles_row[12],
            "matiere10": titles_row[13],
            "matiere11": titles_row[14],
            "matiere12": titles_row[15],
        })
    
    # Mise à jour des placeholders en fonction de la clé du cas BG TP 1 2 3 4 5 6
    elif case_config["key"] == "BG_TP_1": 
        placeholders.update({
            "UE1_Title": titles_row[0],
            "matiere1": titles_row[1],
            "matiere2": titles_row[2],
            "matiere3": titles_row[3],
            "matiere4": titles_row[5],
            "matiere5": titles_row[6],
            "matiere6": titles_row[8],
            "matiere7": titles_row[7],
            "UE2_Title": titles_row[8],
            "matiere8": titles_row[9],
            "matiere9": titles_row[10],
            "matiere10": titles_row[11],
            "matiere11": titles_row[12],
            "matiere12": titles_row[13],
            "UE3_Title": titles_row[14],
            "matiere13": titles_row[15],
            "matiere14": titles_row[16],
            "UE4_Title": titles_row[17],
            "matiere15": titles_row[18],
            "matiere16": titles_row[19],
            "matiere17": titles_row[20],
            "matiere18": titles_row[21],
            "matiere19": titles_row[22],
            "matiere20": titles_row[23],
            "matiere21": titles_row[24],
            "matiere22": titles_row[25],
            
        })
    elif case_config["key"] == "BG_TP_2": 
        placeholders.update({
            "UE1_Title": titles_row[0],
            "matiere1": titles_row[1],
            "matiere2": titles_row[2],
        })
    elif case_config["key"] == "BG_TP_3": 
        placeholders.update({
            "UE1_Title": titles_row[0],
            "matiere1": titles_row[1],
            "matiere2": titles_row[2],
            "matiere3": titles_row[3],
            "matiere4": titles_row[4],
            "UE2_Title": titles_row[5],
            "matiere5": titles_row[6],
            "matiere6": titles_row[7],
            "matiere7": titles_row[8],
            "matiere8": titles_row[9],
            "matiere9": titles_row[10],
            "UE3_Title": titles_row[11],
            "matiere10": titles_row[12],
            "matiere11": titles_row[13],
            "UE4_Title": titles_row[14],
            "matiere12": titles_row[15],
            "matiere13": titles_row[16],
            "matiere14": titles_row[17],
            "matiere15": titles_row[18],
        })
    elif case_config["key"] == "BG_TP_4": 
        placeholders.update({
            "UE1_Title": titles_row[0],
            "matiere1": titles_row[1],
        })
    elif case_config["key"] == "BG_TP_5": 
        placeholders.update({
            "UE1_Title": titles_row[0],
            "matiere1": titles_row[1],
            "matiere2": titles_row[2],
            "matiere3": titles_row[3],
            "matiere4": titles_row[4],
            "matiere5": titles_row[5],
            "UE2_Title": titles_row[6],
            "matiere6": titles_row[7],
            "matiere7": titles_row[8],
            "matiere8": titles_row[9],
            "matiere9": titles_row[10],
            "matiere10": titles_row[11],
            "UE3_Title": titles_row[12],
            "matiere11": titles_row[13],
            "matiere12": titles_row[14],
            "matiere13": titles_row[15],
            "UE4_Title": titles_row[16],
            "matiere14": titles_row[17],
            "matiere15": titles_row[18],
            "matiere16": titles_row[19],
            "matiere17": titles_row[20],
            "matiere18": titles_row[21],
            "matiere19": titles_row[22],
        })
    elif case_config["key"] == "BG_TP_6": 
        placeholders.update({
            "UE1_Title": titles_row[0],
            "matiere1": titles_row[1],
            "matiere2": titles_row[2],
            "matiere3": titles_row[3],
        })
    
    
    # Ajouter les valeurs ECTS aux placeholders, en masquant celles spécifiées
    for i in range(1, 17):
        if i not in case_config["hidden_ects"]:
            placeholders[f"ECTS{i}"] = ects_data.get(f"ECTS{i}", 0)

    return placeholders


def calculate_ue_state(notes):
    notes_between_8_and_10 = sum(8 <= note < 10 for note in notes)
    notes_below_8 = sum(note < 8 for note in notes)

    if all(note >= 10 for note in notes):
        return "VA", ["" for _ in notes]
    elif notes_between_8_and_10 == 1 and notes_below_8 == 0:
        return "VA", ["C" if 8 <= note < 10 else "" for note in notes]
    else:
        states = []
        for note in notes:
            if note < 8:
                states.append("R")
            elif 8 <= note < 10:
                states.append("R" if notes_below_8 > 0 or notes_between_8_and_10 > 1 else "C")
            else:
                states.append("")
        return "NV", states

# Modify the logic where "R" is assigned
def process_ue_notes(placeholders, ue_name, note_indices, grade_column_indices, student_data, case_config):
    ue_notes = []
    ue_ects = []
    valid_indices = []
    
    for i in note_indices:
        grade_str = str(student_data.iloc[grade_column_indices[i-1]]).strip() if pd.notna(student_data.iloc[grade_column_indices[i-1]]) else ""
        ects_value = placeholders.get(f"ECTS{i}", "")

        placeholders[f"etat{i}"] = ""
        placeholders[f"note{i}"] = ""

        if grade_str and grade_str != 'Note' and ects_value and i not in case_config["hidden_ects"]:
            grades_coefficients, special_case = extract_grades_and_coefficients(grade_str)
            if special_case:
                placeholders[f"note{i}"] = special_case
                placeholders[f"ECTS{i}"] = ""  # Ne pas attribuer d'ECTS pour les cas spéciaux
            elif grades_coefficients:  # Vérifier si grades_coefficients n'est pas None ou vide
                individual_average = calculate_weighted_average([g[0] for g in grades_coefficients], [g[1] for g in grades_coefficients])
                if individual_average is not None:
                    ue_notes.append(individual_average)
                    ue_ects.append(float(ects_value))
                    placeholders[f"note{i}"] = f"{individual_average:.2f}"
                    valid_indices.append(i)
                    logging.debug(f"Valid note for index {i}: {individual_average:.2f}")
            else:
                logging.debug(f"No valid grades or coefficients for index {i}")

    # Calculate UE average
    if ue_notes and ue_ects:
        ue_average = calculate_weighted_average(ue_notes, ue_ects)
        placeholders[f"moy{ue_name}"] = f"{ue_average:.2f}"
        logging.debug(f"UE average: {ue_average:.2f}")
    else:
        placeholders[f"moy{ue_name}"] = ""
        logging.debug("No valid notes for UE average calculation")

    # Determine UE state
    if all(note >= 10 for note in ue_notes):
        placeholders[f"etat{ue_name}"] = "VA"
        logging.debug("UE state: VA (all notes >= 10)")
    elif ue_average >= 10:
        placeholders[f"etat{ue_name}"] = "VA"
        logging.debug("UE state: VA (average >= 10)")
    else:
        placeholders[f"etat{ue_name}"] = "NV"
        logging.debug("UE state: NV")

    # Assign individual states and adjust ECTS for display
    for i, note in zip(valid_indices, ue_notes):
        if note < 8:
            placeholders[f"etat{i}"] = "R"
            placeholders[f"ECTS{i}"] = 0  # Set ECTS to 0 for display purposes
            logging.debug(f"Rattrapage for index {i}: note={note:.2f}, ECTS set to 0 for display")

def process_ue4(placeholders, note_indices, grade_column_indices, student_data, case_config):
    ue_notes = []
    ue_ects = []
    for i in note_indices:
        grade_str = str(student_data.iloc[grade_column_indices[i-1]]).strip() if pd.notna(student_data.iloc[grade_column_indices[i-1]]) else ""
        ects_value = placeholders.get(f"ECTS{i}", "")
        
        placeholders[f"note{i}"] = ""
        placeholders[f"etat{i}"] = ""

        if grade_str and grade_str != 'Note' and ects_value and i not in case_config["hidden_ects"]:
            grades_coefficients, special_case = extract_grades_and_coefficients(grade_str)
            if special_case:
                placeholders[f"note{i}"] = special_case
                placeholders[f"ECTS{i}"] = ""  # Ne pas attribuer d'ECTS pour les cas spéciaux
            elif grades_coefficients:
                individual_average = calculate_weighted_average([g[0] for g in grades_coefficients], [g[1] for g in grades_coefficients])
                if individual_average is not None:
                    ue_notes.append(individual_average)
                    placeholders[f"note{i}"] = f"{individual_average:.2f}"
                    if individual_average < 8:
                        placeholders[f"etat{i}"] = "R"
                        placeholders[f"ECTS{i}"] = 0  # Set ECTS to 0 when state is "R"
                    elif 8 <= individual_average < 10:
                        placeholders[f"etat{i}"] = "C"
                        ue_ects.append(float(ects_value))
                    else:
                        ue_ects.append(float(ects_value))
            else:
                logging.debug(f"No valid grades or coefficients for index {i}")

    if ue_notes and ue_ects:
        ue_average = calculate_weighted_average(ue_notes, ue_ects)
        if ue_average is not None:
            placeholders["moyUE4"] = f"{ue_average:.2f}"
            placeholders["etatUE4"] = "VA" if ue_average >= 10 else "NV"
        else:
            placeholders["moyUE4"] = ""
            placeholders["etatUE4"] = "NV"
    else:
        placeholders["moyUE4"] = ""
        placeholders["etatUE4"] = "NV"  # If no valid notes, consider UE as not validated

    # Final check to ensure no state is assigned to empty notes
    for i in note_indices:
        if not placeholders[f"note{i}"]:
            placeholders[f"etat{i}"] = ""
            placeholders[f"ECTS{i}"] = ""  # Ensure ECTS is empty for empty notes
        elif placeholders[f"etat{i}"] == "R":
            placeholders[f"ECTS{i}"] = 0  # Set ECTS to 0 when state is "R"

    return placeholders

def process_and_evaluate_ue(placeholders, ue_name, note_indices, grade_column_indices, student_data, case_config):
    ue_notes = []
    ue_ects = []
    valid_indices = []
    
    for i in note_indices:
        grade_str = str(student_data.iloc[grade_column_indices[i-1]]).strip() if pd.notna(student_data.iloc[grade_column_indices[i-1]]) else ""
        ects_value = placeholders.get(f"ECTS{i}", "")
        placeholders[f"etat{i}"] = ""
        placeholders[f"note{i}"] = ""
        
        if grade_str and grade_str != 'Note' and ects_value and i not in case_config["hidden_ects"]:
            grades_coefficients, special_case = extract_grades_and_coefficients(grade_str)
            if special_case:
                placeholders[f"note{i}"] = special_case
                placeholders[f"ECTS{i}"] = ""  # Ne pas attribuer d'ECTS pour les cas spéciaux
            elif grades_coefficients:
                individual_average = calculate_weighted_average([g[0] for g in grades_coefficients], [g[1] for g in grades_coefficients])
                if individual_average is not None:
                    ue_notes.append(individual_average)
                    ue_ects.append(float(ects_value))
                    placeholders[f"note{i}"] = f"{individual_average:.2f}"
                    valid_indices.append(i)
                    logging.debug(f"Note valide pour l'index {i}: {individual_average:.2f}")
                else:
                    logging.debug(f"Aucune note ou coefficient valide pour l'index {i}")
    
    # Calculer la moyenne de l'UE
    if ue_notes and ue_ects:
        ue_average = calculate_weighted_average(ue_notes, ue_ects)
        placeholders[f"moy{ue_name}"] = f"{ue_average:.2f}"
        logging.debug(f"UE average: {ue_average:.2f}")
    else:
        placeholders[f"moy{ue_name}"] = ""
        logging.debug("Aucune note valide pour le calcul de la moyenne UE")
    
    # Déterminer l'état de l'UE
    if ue_notes:
        notes_below_8 = sum(1 for note in ue_notes if note < 8)
        notes_between_8_and_10 = sum(1 for note in ue_notes if 8 <= note < 10)
        
        if all(note >= 10 for note in ue_notes):
            placeholders[f"etat{ue_name}"] = "VA"
            logging.debug("État de l'UE : VA (toutes les notes >= 10)")
        elif notes_below_8 == 0 and notes_between_8_and_10 <= 1:
            placeholders[f"etat{ue_name}"] = "VA"
            logging.debug("État de l'UE : VA (au plus une note entre 8 et 10, aucune en dessous de 8)")
        else:
            placeholders[f"etat{ue_name}"] = "NV"
            logging.debug("État de l'UE : NV")
        
        # Attribuer des états individuels et ajuster ECTS pour l'affichage
        for i, note in zip(valid_indices, ue_notes):
            if note < 8:
                placeholders[f"etat{i}"] = "R"
                placeholders[f"ECTS{i}"] = 0
                logging.debug(f"Rattrapage pour l'index {i}: note={note:.2f}, ECTS défini sur 0 pour l'affichage")
            elif 8 <= note < 10:
                if notes_below_8 > 0 or notes_between_8_and_10 > 1:
                    placeholders[f"etat{i}"] = "R"
                    placeholders[f"ECTS{i}"] = 0
                else:
                    placeholders[f"etat{i}"] = "C"
            else:
                placeholders[f"etat{i}"] = ""
    else:
        placeholders[f"etat{ue_name}"] = ""
        for i in note_indices:
            placeholders[f"etat{i}"] = ""
    
    return placeholders

def process_UE4_and_evaluate_ue(placeholders, ue_name, note_indices, grade_column_indices, student_data, case_config):
    ue_notes = []
    ue_ects = []
    
    logging.debug(f"note_indices for UE4: {note_indices}")
    logging.debug(f"Length of grade_column_indices: {len(grade_column_indices)}")
    logging.debug(f"grade_column_indices: {grade_column_indices}")
    
    valid_note_indices = [i for i in note_indices if i-1 < len(grade_column_indices)]
    
    for i in valid_note_indices:
        if i-1 < len(grade_column_indices):
            grade_str = str(student_data.iloc[grade_column_indices[i-1]]).strip() if pd.notna(student_data.iloc[grade_column_indices[i-1]]) else ""
            ects_value = placeholders.get(f"ECTS{i}", "")
            
            placeholders[f"note{i}"] = ""
            placeholders[f"etat{i}"] = ""
            
            if grade_str and grade_str != 'Note' and ects_value and i not in case_config["hidden_ects"]:
                grades_coefficients, special_case = extract_grades_and_coefficients(grade_str)
                if special_case:
                    placeholders[f"note{i}"] = special_case
                    placeholders[f"ECTS{i}"] = ""  # Ne pas attribuer d'ECTS pour les cas spéciaux
                elif grades_coefficients:
                    individual_average = calculate_weighted_average([g[0] for g in grades_coefficients], [g[1] for g in grades_coefficients])
                    if individual_average is not None:
                        ue_notes.append(individual_average)
                        placeholders[f"note{i}"] = f"{individual_average:.2f}"
                        if individual_average < 8:
                            placeholders[f"etat{i}"] = "R"
                            placeholders[f"ECTS{i}"] = 0  # Set ECTS to 0 when state is "R"
                        elif 8 <= individual_average < 10:
                            placeholders[f"etat{i}"] = "C"
                            ue_ects.append(float(ects_value))
                        else:
                            ue_ects.append(float(ects_value))
                else:
                    logging.debug(f"No valid grades or coefficients for index {i}")
        else:
            logging.warning(f"Index {i-1} is out of range for grade_column_indices. Skipping this index.")
    
    # If no valid notes, consider UE as not validated
    if ue_notes and ue_ects:
        ue_average = calculate_weighted_average(ue_notes, ue_ects)
        if ue_average is not None:
            placeholders[f"moy{ue_name}"] = f"{ue_average:.2f}"
            placeholders[f"etat{ue_name}"] = "VA" if ue_average >= 10 else "NV"
        else:
            placeholders[f"moy{ue_name}"] = ""
            placeholders[f"etat{ue_name}"] = "NV"
    else:
        placeholders[f"moy{ue_name}"] = ""
        placeholders[f"etat{ue_name}"] = "NV"
    
    # Final check to ensure no state is assigned to empty notes
    for i in note_indices:
        if f"note{i}" not in placeholders:
            placeholders[f"note{i}"] = ""
        if f"etat{i}" not in placeholders:
            placeholders[f"etat{i}"] = ""
        if f"ECTS{i}" not in placeholders:
            placeholders[f"ECTS{i}"] = ""
        
        if not placeholders[f"note{i}"]:
            placeholders[f"etat{i}"] = ""
            placeholders[f"ECTS{i}"] = ""  # Ensure ECTS is empty for empty notes
        elif placeholders[f"etat{i}"] == "R":
            placeholders[f"ECTS{i}"] = 0  # Set ECTS to 0 when state is "R"
    
    return placeholders

def set_hidden_text(paragraph):
    """Set all runs in the paragraph as hidden."""
    for run in paragraph.runs:
        run.font.size = Pt(1)  # Optionnel : taille minimale pour encore plus de discrétion
        # Ajouter un élément caché à la propriété run
        rPr = run._element.get_or_add_rPr()
        vanishing = OxmlElement('w:vanish')
        rPr.append(vanishing)

def generate_word_document(student_data, case_config, template_path, output_dir):
    ects_config = read_ects_config()
    current_date = datetime.now().strftime("%d/%m/%Y")
    group_name = student_data["Nom Groupe"]
    is_relevant_group = group_name in settings.RELEVANT_GROUPS
    logger.debug("Processing document for group: %s", group_name)

    # Corriger la clé du cas si nécessaire
    corrected_key = case_config["key"].replace("_", "-")

    ects_data_key = corrected_key
    if corrected_key == "M2_S3_MAGI_MEFIM":
        if "MAGI" in student_data["Nom Groupe"]:
            ects_data_key = "M2-S3-MAGI"
        elif "MEFIM" in student_data["Nom Groupe"]:
            ects_data_key = "M2-S3-MEFIM"

    ects_data = ects_config.get(ects_data_key, [{}])[0]
    logger.debug(f"ECTS data for {corrected_key}: {ects_data}")

    placeholders = generate_placeholders(case_config["titles_row"], case_config, student_data, current_date, ects_data)

    # New logic for M1-S1
    if case_config["key"] == "M1_S1":
        process_ue_notes(placeholders, "UE1", [1, 2, 3], case_config["grade_column_indices"], student_data, case_config)
        process_ue_notes(placeholders, "UE2", [4], case_config["grade_column_indices"], student_data, case_config)
        process_ue_notes(placeholders, "UE3", [5, 6], case_config["grade_column_indices"], student_data, case_config)
        process_ue4(placeholders, [7, 8, 9, 10, 11, 12], case_config["grade_column_indices"], student_data, case_config)
        process_ue_notes(placeholders, "UESPE", [13, 14, 15], case_config["grade_column_indices"], student_data, case_config)

        # Get UE1 notes, treating empty strings as None
        ue1_notes = [float(placeholders[f"note{i}"]) if placeholders[f"note{i}"] and placeholders[f"note{i}"] != "" and i not in case_config["hidden_ects"] and placeholders.get(f"ECTS{i}", "") != "" else None for i in range(1, 4)]

        # Initialize all states to empty string
        placeholders["etatUE1"] = ""
        for i in range(1, 4):
            placeholders[f"etat{i}"] = ""

        # Only process if there are any non-None values
        if any(note is not None for note in ue1_notes):
            # Count notes in different ranges, ignoring None values
            notes_between_8_and_10 = sum(8 <= note < 10 for note in ue1_notes if note is not None)
            notes_below_8 = sum(note < 8 for note in ue1_notes if note is not None)

            # Determine UE1 state and individual states
            if all(note >= 10 for note in ue1_notes if note is not None):
                placeholders["etatUE1"] = "VA"
            elif notes_between_8_and_10 == 1 and notes_below_8 == 0:
                placeholders["etatUE1"] = "VA"
                for i, note in enumerate(ue1_notes, start=1):
                    if note is not None and 8 <= note < 10 and i not in case_config["hidden_ects"] and placeholders.get(f"ECTS{i}", "") != "":
                        placeholders[f"etat{i}"] = "C"
            else:
                placeholders["etatUE1"] = "NV"
                for i, note in enumerate(ue1_notes, start=1):
                    if note is not None and i not in case_config["hidden_ects"] and placeholders.get(f"ECTS{i}", "") != "":
                        if note < 8:
                            placeholders[f"etat{i}"] = "R"
                        elif 8 <= note < 10:
                            placeholders[f"etat{i}"] = "R" if notes_below_8 > 0 or notes_between_8_and_10 > 1 else "C"
                        else:
                            placeholders[f"etat{i}"] = ""
                    else:
                        placeholders[f"etat{i}"] = ""
        else:
            # If all notes are None or empty, set all states to empty string
            placeholders["etatUE1"] = ""
            for i in range(1, 4):
                placeholders[f"etat{i}"] = ""
    elif case_config["key"] == "M1_S2":
        process_ue_notes(placeholders, "UE1", [1, 2, 3], case_config["grade_column_indices"], student_data, case_config)
        process_ue_notes(placeholders, "UE2", [4, 5], case_config["grade_column_indices"], student_data, case_config)
        process_ue_notes(placeholders, "UE3", [6, 7, 8, 9, 10, 11, 12], case_config["grade_column_indices"], student_data, case_config)
        process_ue_notes(placeholders, "UESPE", [13, 14, 15, 16], case_config["grade_column_indices"], student_data, case_config)

        # Traitement spécifique pour UE1, similaire à M1-S1
        ue1_notes = [float(placeholders[f"note{i}"]) if placeholders[f"note{i}"] and placeholders[f"note{i}"] != "" and i not in case_config["hidden_ects"] and placeholders.get(f"ECTS{i}", "") != "" else None for i in range(1, 4)]

        placeholders["etatUE1"] = ""
        for i in range(1, 4):
            placeholders[f"etat{i}"] = ""

        if any(note is not None for note in ue1_notes):
            notes_between_8_and_10 = sum(8 <= note < 10 for note in ue1_notes if note is not None)
            notes_below_8 = sum(note < 8 for note in ue1_notes if note is not None)

            if all(note >= 10 for note in ue1_notes if note is not None):
                placeholders["etatUE1"] = "VA"
            elif notes_between_8_and_10 == 1 and notes_below_8 == 0:
                placeholders["etatUE1"] = "VA"
                for i, note in enumerate(ue1_notes, start=1):
                    if note is not None and 8 <= note < 10 and i not in case_config["hidden_ects"] and placeholders.get(f"ECTS{i}", "") != "":
                        placeholders[f"etat{i}"] = "C"
            else:
                placeholders["etatUE1"] = "NV"
                for i, note in enumerate(ue1_notes, start=1):
                    if note is not None and i not in case_config["hidden_ects"] and placeholders.get(f"ECTS{i}", "") != "":
                        if note < 8:
                            placeholders[f"etat{i}"] = "R"
                        elif 8 <= note < 10:
                            placeholders[f"etat{i}"] = "R" if notes_below_8 > 0 or notes_between_8_and_10 > 1 else "C"
                        else:
                            placeholders[f"etat{i}"] = ""
                    else:
                        placeholders[f"etat{i}"] = ""
        else:
            placeholders["etatUE1"] = ""
            for i in range(1, 4):
                placeholders[f"etat{i}"] = ""
    elif case_config["key"] == "M2_S3_MAGI":
        process_ue_notes(placeholders, "UE1", [1, 2, 3], case_config["grade_column_indices"], student_data, case_config)
        process_ue_notes(placeholders, "UE2", [4], case_config["grade_column_indices"], student_data, case_config)
        process_ue_notes(placeholders, "UE3", [5, 6, 7, 8, 9], case_config["grade_column_indices"], student_data, case_config)
        process_ue_notes(placeholders, "UESPE", [10, 11, 12, 13], case_config["grade_column_indices"], student_data, case_config)

        # Gestion des notes et des états pour UE1
        ue1_notes = [float(placeholders[f"note{i}"]) if placeholders[f"note{i}"] and placeholders[f"note{i}"] != "" and i not in case_config["hidden_ects"] and placeholders.get(f"ECTS{i}", "") != "" else None for i in range(1, 4)]
        placeholders["etatUE1"] = ""
        for i in range(1, 4):
            placeholders[f"etat{i}"] = ""

        if any(note is not None for note in ue1_notes):
            notes_between_8_and_10 = sum(8 <= note < 10 for note in ue1_notes if note is not None)
            notes_below_8 = sum(note < 8 for note in ue1_notes if note is not None)

            if all(note >= 10 for note in ue1_notes if note is not None):
                placeholders["etatUE1"] = "VA"
            elif notes_between_8_and_10 == 1 and notes_below_8 == 0:
                placeholders["etatUE1"] = "VA"
                for i, note in enumerate(ue1_notes, start=1):
                    if note is not None and 8 <= note < 10 and i not in case_config["hidden_ects"] and placeholders.get(f"ECTS{i}", "") != "":
                        placeholders[f"etat{i}"] = "C"
            else:
                placeholders["etatUE1"] = "NV"
                for i, note in enumerate(ue1_notes, start=1):
                    if note is not None and i not in case_config["hidden_ects"] and placeholders.get(f"ECTS{i}", "") != "":
                        if note < 8:
                            placeholders[f"etat{i}"] = "R"
                        elif 8 <= note < 10:
                            placeholders[f"etat{i}"] = "R" if notes_below_8 > 0 or notes_between_8_and_10 > 1 else "C"
                        else:
                            placeholders[f"etat{i}"] = ""
                    else:
                        placeholders[f"etat{i}"] = ""
        else:
            placeholders["etatUE1"] = ""
            for i in range(1, 4):
                placeholders[f"etat{i}"] = ""
    elif case_config["key"] == "M2_S3_MEFIM":
        process_ue_notes(placeholders, "UE1", [1, 2, 3], case_config["grade_column_indices"], student_data, case_config)
        process_ue_notes(placeholders, "UE2", [4], case_config["grade_column_indices"], student_data, case_config)
        process_ue_notes(placeholders, "UE3", [5, 6, 7, 8, 9], case_config["grade_column_indices"], student_data, case_config)
        process_ue_notes(placeholders, "UESPE", [10, 11, 12, 13], case_config["grade_column_indices"], student_data, case_config)
        
        # Gestion des notes et des états pour UE1 (similaire à M2_S3_MAGI)
        ue1_notes = [float(placeholders[f"note{i}"]) if placeholders[f"note{i}"] and placeholders[f"note{i}"] != "" and i not in case_config["hidden_ects"] and placeholders.get(f"ECTS{i}", "") != "" else None for i in range(1, 4)]
        placeholders["etatUE1"] = ""
        for i in range(1, 4):
            placeholders[f"etat{i}"] = ""

        if any(note is not None for note in ue1_notes):
            notes_between_8_and_10 = sum(8 <= note < 10 for note in ue1_notes if note is not None)
            notes_below_8 = sum(note < 8 for note in ue1_notes if note is not None)

            if all(note >= 10 for note in ue1_notes if note is not None):
                placeholders["etatUE1"] = "VA"
            elif notes_between_8_and_10 == 1 and notes_below_8 == 0:
                placeholders["etatUE1"] = "VA"
                for i, note in enumerate(ue1_notes, start=1):
                    if note is not None and 8 <= note < 10 and i not in case_config["hidden_ects"] and placeholders.get(f"ECTS{i}", "") != "":
                        placeholders[f"etat{i}"] = "C"
            else:
                placeholders["etatUE1"] = "NV"
                for i, note in enumerate(ue1_notes, start=1):
                    if note is not None and i not in case_config["hidden_ects"] and placeholders.get(f"ECTS{i}", "") != "":
                        if note < 8:
                            placeholders[f"etat{i}"] = "R"
                        elif 8 <= note < 10:
                            placeholders[f"etat{i}"] = "R" if notes_below_8 > 0 or notes_between_8_and_10 > 1 else "C"
                        else:
                            placeholders[f"etat{i}"] = ""
                    else:
                        placeholders[f"etat{i}"] = ""
        else:
            placeholders["etatUE1"] = ""
            for i in range(1, 4):
                placeholders[f"etat{i}"] = ""
    elif case_config["key"] == "M2_S3_MAPI":
        process_ue_notes(placeholders, "UE1", [1, 2, 3], case_config["grade_column_indices"], student_data, case_config)
        process_ue_notes(placeholders, "UE2", [4], case_config["grade_column_indices"], student_data, case_config)
        process_ue_notes(placeholders, "UE3", [5, 6, 7, 8, 9], case_config["grade_column_indices"], student_data, case_config)
        process_ue_notes(placeholders, "UESPE", [10, 11, 12, 13, 14], case_config["grade_column_indices"], student_data, case_config)

        # Gestion des notes et des états pour UE1 (similaire à M2_S3_MAGI)
        ue1_notes = [float(placeholders[f"note{i}"]) if placeholders[f"note{i}"] and placeholders[f"note{i}"] != "" and i not in case_config["hidden_ects"] and placeholders.get(f"ECTS{i}", "") != "" else None for i in range(1, 4)]
        placeholders["etatUE1"] = ""
        for i in range(1, 4):
            placeholders[f"etat{i}"] = ""

        if any(note is not None for note in ue1_notes):
            notes_between_8_and_10 = sum(8 <= note < 10 for note in ue1_notes if note is not None)
            notes_below_8 = sum(note < 8 for note in ue1_notes if note is not None)

            if all(note >= 10 for note in ue1_notes if note is not None):
                placeholders["etatUE1"] = "VA"
            elif notes_between_8_and_10 == 1 and notes_below_8 == 0:
                placeholders["etatUE1"] = "VA"
                for i, note in enumerate(ue1_notes, start=1):
                    if note is not None and 8 <= note < 10 and i not in case_config["hidden_ects"] and placeholders.get(f"ECTS{i}", "") != "":
                        placeholders[f"etat{i}"] = "C"
            else:
                placeholders["etatUE1"] = "NV"
                for i, note in enumerate(ue1_notes, start=1):
                    if note is not None and i not in case_config["hidden_ects"] and placeholders.get(f"ECTS{i}", "") != "":
                        if note < 8:
                            placeholders[f"etat{i}"] = "R"
                        elif 8 <= note < 10:
                            placeholders[f"etat{i}"] = "R" if notes_below_8 > 0 or notes_between_8_and_10 > 1 else "C"
                        else:
                            placeholders[f"etat{i}"] = ""
                    else:
                        placeholders[f"etat{i}"] = ""
        else:
            placeholders["etatUE1"] = ""
            for i in range(1, 4):
                placeholders[f"etat{i}"] = ""
    elif case_config["key"] == "M2_S4":
        process_ue_notes(placeholders, "UE1", [1], case_config["grade_column_indices"], student_data, case_config)
        process_ue_notes(placeholders, "UE2", [2, 3], case_config["grade_column_indices"], student_data, case_config)
        process_ue_notes(placeholders, "UE3", [4, 5, 6, 7, 8], case_config["grade_column_indices"], student_data, case_config)
        process_ue_notes(placeholders, "UESPE", [9, 10, 11], case_config["grade_column_indices"], student_data, case_config)

        # Gestion des notes et des états pour UE1 (cas M2_S4 avec un seul indice pour UE1)
        ue1_notes = [float(placeholders[f"note{i}"]) if placeholders[f"note{i}"] and placeholders[f"note{i}"] != "" and i not in case_config["hidden_ects"] and placeholders.get(f"ECTS{i}", "") != "" else None for i in range(1, 2)]
        placeholders["etatUE1"] = ""
        for i in range(1, 2):
            placeholders[f"etat{i}"] = ""

        if any(note is not None for note in ue1_notes):
            notes_between_8_and_10 = sum(8 <= note < 10 for note in ue1_notes if note is not None)
            notes_below_8 = sum(note < 8 for note in ue1_notes if note is not None)

            if all(note >= 10 for note in ue1_notes if note is not None):
                placeholders["etatUE1"] = "VA"
            elif notes_between_8_and_10 == 1 and notes_below_8 == 0:
                placeholders["etatUE1"] = "VA"
                for i, note in enumerate(ue1_notes, start=1):
                    if note is not None and 8 <= note < 10 and i not in case_config["hidden_ects"] and placeholders.get(f"ECTS{i}", "") != "":
                        placeholders[f"etat{i}"] = "C"
            else:
                placeholders["etatUE1"] = "NV"
                for i, note in enumerate(ue1_notes, start=1):
                    if note is not None and i not in case_config["hidden_ects"] and placeholders.get(f"ECTS{i}", "") != "":
                        if note < 8:
                            placeholders[f"etat{i}"] = "R"
                        elif 8 <= note < 10:
                            placeholders[f"etat{i}"] = "R" if notes_below_8 > 0 or notes_between_8_and_10 > 1 else "C"
                        else:
                            placeholders[f"etat{i}"] = ""
                    else:
                        placeholders[f"etat{i}"] = ""
        else:
            placeholders["etatUE1"] = ""
            for i in range(1, 2):
                placeholders[f"etat{i}"] = ""
    # Cas BG_ALT_1
    elif case_config["key"] == "BG_ALT_1":
        process_and_evaluate_ue(placeholders, "UE1", [1, 2, 3], case_config["grade_column_indices"], student_data, case_config)
        process_and_evaluate_ue(placeholders, "UE2", [4, 5], case_config["grade_column_indices"], student_data, case_config)
        process_and_evaluate_ue(placeholders, "UE3", [6], case_config["grade_column_indices"], student_data, case_config)
        process_UE4_and_evaluate_ue(placeholders, "UE4", [7, 8, 9, 10, 11, 12, 13, 14], case_config["grade_column_indices"], student_data, case_config)

        # Gestion des notes et des états pour UE1
        ue1_notes = [float(placeholders[f"note{i}"]) if placeholders[f"note{i}"] and placeholders[f"note{i}"] != "" and i not in case_config["hidden_ects"] and placeholders.get(f"ECTS{i}", "") != "" else None for i in range(1, 4)]
        placeholders["etatUE1"] = ""
        for i in range(1, 4):
            placeholders[f"etat{i}"] = ""

        if any(note is not None for note in ue1_notes):
            notes_between_8_and_10 = sum(8 <= note < 10 for note in ue1_notes if note is not None)
            notes_below_8 = sum(note < 8 for note in ue1_notes if note is not None)

            if all(note >= 10 for note in ue1_notes if note is not None):
                placeholders["etatUE1"] = "VA"
            elif notes_between_8_and_10 == 1 and notes_below_8 == 0:
                placeholders["etatUE1"] = "VA"
                for i, note in enumerate(ue1_notes, start=1):
                    if note is not None and 8 <= note < 10 and i not in case_config["hidden_ects"] and placeholders.get(f"ECTS{i}", "") != "":
                        placeholders[f"etat{i}"] = "C"
            else:
                placeholders["etatUE1"] = "NV"
                for i, note in enumerate(ue1_notes, start=1):
                    if note is not None and i not in case_config["hidden_ects"] and placeholders.get(f"ECTS{i}", "") != "":
                        if note < 8:
                            placeholders[f"etat{i}"] = "R"
                        elif 8 <= note < 10:
                            placeholders[f"etat{i}"] = "R" if notes_below_8 > 0 or notes_between_8_and_10 > 1 else "C"
                        else:
                            placeholders[f"etat{i}"] = ""
                    else:
                        placeholders[f"etat{i}"] = ""
        else:
            placeholders["etatUE1"] = ""
            for i in range(1, 4):
                placeholders[f"etat{i}"] = ""
    # Cas BG_ALT_2
    elif case_config["key"] == "BG_ALT_2":
        process_and_evaluate_ue(placeholders, "UE1", [1, 2, 3, 4], case_config["grade_column_indices"], student_data, case_config)
        process_and_evaluate_ue(placeholders, "UE2", [5, 6, 7], case_config["grade_column_indices"], student_data, case_config)
        process_and_evaluate_ue(placeholders, "UE3", [8], case_config["grade_column_indices"], student_data, case_config)
        process_UE4_and_evaluate_ue(placeholders, "UE4", [9, 10, 11, 12, 13, 14, 15], case_config["grade_column_indices"], student_data, case_config)

        # Gestion des notes et des états pour UE1
        ue1_notes = [float(placeholders[f"note{i}"]) if placeholders[f"note{i}"] and placeholders[f"note{i}"] != "" and i not in case_config["hidden_ects"] and placeholders.get(f"ECTS{i}", "") != "" else None for i in range(1, 5)]
        placeholders["etatUE1"] = ""
        for i in range(1, 5):
            placeholders[f"etat{i}"] = ""

        if any(note is not None for note in ue1_notes):
            notes_between_8_and_10 = sum(8 <= note < 10 for note in ue1_notes if note is not None)
            notes_below_8 = sum(note < 8 for note in ue1_notes if note is not None)

            if all(note >= 10 for note in ue1_notes if note is not None):
                placeholders["etatUE1"] = "VA"
            elif notes_between_8_and_10 == 1 and notes_below_8 == 0:
                placeholders["etatUE1"] = "VA"
                for i, note in enumerate(ue1_notes, start=1):
                    if note is not None and 8 <= note < 10 and i not in case_config["hidden_ects"] and placeholders.get(f"ECTS{i}", "") != "":
                        placeholders[f"etat{i}"] = "C"
            else:
                placeholders["etatUE1"] = "NV"
                for i, note in enumerate(ue1_notes, start=1):
                    if note is not None and i not in case_config["hidden_ects"] and placeholders.get(f"ECTS{i}", "") != "":
                        if note < 8:
                            placeholders[f"etat{i}"] = "R"
                        elif 8 <= note < 10:
                            placeholders[f"etat{i}"] = "R" if notes_below_8 > 0 or notes_between_8_and_10 > 1 else "C"
                        else:
                            placeholders[f"etat{i}"] = ""
                    else:
                        placeholders[f"etat{i}"] = ""
        else:
            placeholders["etatUE1"] = ""
            for i in range(1, 5):
                placeholders[f"etat{i}"] = ""
    # Cas BG_ALT_3
    elif case_config["key"] == "BG_ALT_3":
        process_and_evaluate_ue(placeholders, "UE1", [1, 2, 3, 4, 5], case_config["grade_column_indices"], student_data, case_config)
        process_and_evaluate_ue(placeholders, "UE2", [6, 7, 8], case_config["grade_column_indices"], student_data, case_config)
        process_and_evaluate_ue(placeholders, "UE3", [9], case_config["grade_column_indices"], student_data, case_config)
        process_UE4_and_evaluate_ue(placeholders, "UE4", [10, 11, 12, 13], case_config["grade_column_indices"], student_data, case_config)

        # Gestion des notes et des états pour UE1
        ue1_notes = [float(placeholders[f"note{i}"]) if placeholders[f"note{i}"] and placeholders[f"note{i}"] != "" and i not in case_config["hidden_ects"] and placeholders.get(f"ECTS{i}", "") != "" else None for i in range(1, 6)]
        placeholders["etatUE1"] = ""
        for i in range(1, 6):
            placeholders[f"etat{i}"] = ""

        if any(note is not None for note in ue1_notes):
            notes_between_8_and_10 = sum(8 <= note < 10 for note in ue1_notes if note is not None)
            notes_below_8 = sum(note < 8 for note in ue1_notes if note is not None)

            if all(note >= 10 for note in ue1_notes if note is not None):
                placeholders["etatUE1"] = "VA"
            elif notes_between_8_and_10 == 1 and notes_below_8 == 0:
                placeholders["etatUE1"] = "VA"
                for i, note in enumerate(ue1_notes, start=1):
                    if note is not None and 8 <= note < 10 and i not in case_config["hidden_ects"] and placeholders.get(f"ECTS{i}", "") != "":
                        placeholders[f"etat{i}"] = "C"
            else:
                placeholders["etatUE1"] = "NV"
                for i, note in enumerate(ue1_notes, start=1):
                    if note is not None and i not in case_config["hidden_ects"] and placeholders.get(f"ECTS{i}", "") != "":
                        if note < 8:
                            placeholders[f"etat{i}"] = "R"
                        elif 8 <= note < 10:
                            placeholders[f"etat{i}"] = "R" if notes_below_8 > 0 or notes_between_8_and_10 > 1 else "C"
                        else:
                            placeholders[f"etat{i}"] = ""
                    else:
                        placeholders[f"etat{i}"] = ""
        else:
            placeholders["etatUE1"] = ""
            for i in range(1, 6):
                placeholders[f"etat{i}"] = ""
        # Cas BG_ALT_4
    elif case_config["key"] == "BG_ALT_4":
        process_and_evaluate_ue(placeholders, "UE1", [1, 2, 3], case_config["grade_column_indices"], student_data, case_config)
        process_and_evaluate_ue(placeholders, "UE2", [4, 5, 6], case_config["grade_column_indices"], student_data, case_config)
        process_and_evaluate_ue(placeholders, "UE3", [7, 8, 9, 10, 11, 12], case_config["grade_column_indices"], student_data, case_config)


        # Gestion des notes et des états pour UE1 (similaire à BG_ALT_1)
        ue1_notes = [float(placeholders[f"note{i}"]) if placeholders[f"note{i}"] and placeholders[f"note{i}"] != "" and i not in case_config["hidden_ects"] and placeholders.get(f"ECTS{i}", "") != "" else None for i in range(1, 4)]
        placeholders["etatUE1"] = ""
        for i in range(1, 4):
            placeholders[f"etat{i}"] = ""

        if any(note is not None for note in ue1_notes):
            notes_between_8_and_10 = sum(8 <= note < 10 for note in ue1_notes if note is not None)
            notes_below_8 = sum(note < 8 for note in ue1_notes if note is not None)

            if all(note >= 10 for note in ue1_notes if note is not None):
                placeholders["etatUE1"] = "VA"
            elif notes_between_8_and_10 == 1 and notes_below_8 == 0:
                placeholders["etatUE1"] = "VA"
                for i, note in enumerate(ue1_notes, start=1):
                    if note is not None and 8 <= note < 10 and i not in case_config["hidden_ects"] and placeholders.get(f"ECTS{i}", "") != "":
                        placeholders[f"etat{i}"] = "C"
            else:
                placeholders["etatUE1"] = "NV"
                for i, note in enumerate(ue1_notes, start=1):
                    if note is not None and i not in case_config["hidden_ects"] and placeholders.get(f"ECTS{i}", "") != "":
                        if note < 8:
                            placeholders[f"etat{i}"] = "R"
                        elif 8 <= note < 10:
                            placeholders[f"etat{i}"] = "R" if notes_below_8 > 0 or notes_between_8_and_10 > 1 else "C"
                        else:
                            placeholders[f"etat{i}"] = ""
                    else:
                        placeholders[f"etat{i}"] = ""
        else:
            placeholders["etatUE1"] = ""
            for i in range(1, 4):
                placeholders[f"etat{i}"] = ""
    # Cas BG_ALT_5
    elif case_config["key"] == "BG_ALT_5":
        process_and_evaluate_ue(placeholders, "UE1", [1, 2, 3], case_config["grade_column_indices"], student_data, case_config)
        process_and_evaluate_ue(placeholders, "UE2", [4, 5, 6], case_config["grade_column_indices"], student_data, case_config)
        process_and_evaluate_ue(placeholders, "UE3", [7, 8], case_config["grade_column_indices"], student_data, case_config)
        process_UE4_and_evaluate_ue(placeholders, "UE4", [9, 10, 11, 12, 13, 14], case_config["grade_column_indices"], student_data, case_config)

        # Gestion des notes et des états pour UE1 (similaire à BG_ALT_1)
        ue1_notes = [float(placeholders[f"note{i}"]) if placeholders[f"note{i}"] and placeholders[f"note{i}"] != "" and i not in case_config["hidden_ects"] and placeholders.get(f"ECTS{i}", "") != "" else None for i in range(1, 4)]
        placeholders["etatUE1"] = ""
        for i in range(1, 4):
            placeholders[f"etat{i}"] = ""

        if any(note is not None for note in ue1_notes):
            notes_between_8_and_10 = sum(8 <= note < 10 for note in ue1_notes if note is not None)
            notes_below_8 = sum(note < 8 for note in ue1_notes if note is not None)

            if all(note >= 10 for note in ue1_notes if note is not None):
                placeholders["etatUE1"] = "VA"
            elif notes_between_8_and_10 == 1 and notes_below_8 == 0:
                placeholders["etatUE1"] = "VA"
                for i, note in enumerate(ue1_notes, start=1):
                    if note is not None and 8 <= note < 10 and i not in case_config["hidden_ects"] and placeholders.get(f"ECTS{i}", "") != "":
                        placeholders[f"etat{i}"] = "C"
            else:
                placeholders["etatUE1"] = "NV"
                for i, note in enumerate(ue1_notes, start=1):
                    if note is not None and i not in case_config["hidden_ects"] and placeholders.get(f"ECTS{i}", "") != "":
                        if note < 8:
                            placeholders[f"etat{i}"] = "R"
                        elif 8 <= note < 10:
                            placeholders[f"etat{i}"] = "R" if notes_below_8 > 0 or notes_between_8_and_10 > 1 else "C"
                        else:
                            placeholders[f"etat{i}"] = ""
                    else:
                        placeholders[f"etat{i}"] = ""
        else:
            placeholders["etatUE1"] = ""
            for i in range(1, 4):
                placeholders[f"etat{i}"] = ""
    # Cas BG_ALT_6
    elif case_config["key"] == "BG_ALT_6":
        process_and_evaluate_ue(placeholders, "UE1", [1, 2, 3], case_config["grade_column_indices"], student_data, case_config)
        process_and_evaluate_ue(placeholders, "UE2", [4, 5], case_config["grade_column_indices"], student_data, case_config)
        process_and_evaluate_ue(placeholders, "UE3", [6], case_config["grade_column_indices"], student_data, case_config)
        process_UE4_and_evaluate_ue(placeholders, "UE4", [7, 8, 9, 10, 11, 12], case_config["grade_column_indices"], student_data, case_config)

        # Gestion des notes et des états pour UE1 (similaire à BG_ALT_1)
        ue1_notes = [float(placeholders[f"note{i}"]) if placeholders[f"note{i}"] and placeholders[f"note{i}"] != "" and i not in case_config["hidden_ects"] and placeholders.get(f"ECTS{i}", "") != "" else None for i in range(1, 4)]
        placeholders["etatUE1"] = ""
        for i in range(1, 4):
            placeholders[f"etat{i}"] = ""

        if any(note is not None for note in ue1_notes):
            notes_between_8_and_10 = sum(8 <= note < 10 for note in ue1_notes if note is not None)
            notes_below_8 = sum(note < 8 for note in ue1_notes if note is not None)

            if all(note >= 10 for note in ue1_notes if note is not None):
                placeholders["etatUE1"] = "VA"
            elif notes_between_8_and_10 == 1 and notes_below_8 == 0:
                placeholders["etatUE1"] = "VA"
                for i, note in enumerate(ue1_notes, start=1):
                    if note is not None and 8 <= note < 10 and i not in case_config["hidden_ects"] and placeholders.get(f"ECTS{i}", "") != "":
                        placeholders[f"etat{i}"] = "C"
            else:
                placeholders["etatUE1"] = "NV"
                for i, note in enumerate(ue1_notes, start=1):
                    if note is not None and i not in case_config["hidden_ects"] and placeholders.get(f"ECTS{i}", "") != "":
                        if note < 8:
                            placeholders[f"etat{i}"] = "R"
                        elif 8 <= note < 10:
                            placeholders[f"etat{i}"] = "R" if notes_below_8 > 0 or notes_between_8_and_10 > 1 else "C"
                        else:
                            placeholders[f"etat{i}"] = ""
                    else:
                        placeholders[f"etat{i}"] = ""
        else:
            placeholders["etatUE1"] = ""
            for i in range(1, 4):
                placeholders[f"etat{i}"] = ""
    # Cas BG_TP_1
    elif case_config["key"] == "BG_TP_1":
        process_and_evaluate_ue(placeholders, "UE1", [1, 2, 3, 4, 5, 6, 7], case_config["grade_column_indices"], student_data, case_config)
        process_and_evaluate_ue(placeholders, "UE2", [8, 9, 10, 11, 12], case_config["grade_column_indices"], student_data, case_config)
        process_and_evaluate_ue(placeholders, "UE3", [13, 14], case_config["grade_column_indices"], student_data, case_config)
        process_UE4_and_evaluate_ue(placeholders, "UE4", [15, 16, 17, 18, 19, 20, 21, 22], case_config["grade_column_indices"], student_data, case_config)

        # Gestion des notes et des états pour UE1
        ue1_notes = [float(placeholders[f"note{i}"]) if placeholders[f"note{i}"] and placeholders[f"note{i}"] != "" and i not in case_config["hidden_ects"] and placeholders.get(f"ECTS{i}", "") != "" else None for i in range(1, 8)]
        placeholders["etatUE1"] = ""
        for i in range(1, 8):
            placeholders[f"etat{i}"] = ""

        if any(note is not None for note in ue1_notes):
            notes_between_8_and_10 = sum(8 <= note < 10 for note in ue1_notes if note is not None)
            notes_below_8 = sum(note < 8 for note in ue1_notes if note is not None)

            if all(note >= 10 for note in ue1_notes if note is not None):
                placeholders["etatUE1"] = "VA"
            elif notes_between_8_and_10 == 1 and notes_below_8 == 0:
                placeholders["etatUE1"] = "VA"
                for i, note in enumerate(ue1_notes, start=1):
                    if note is not None and 8 <= note < 10 and i not in case_config["hidden_ects"] and placeholders.get(f"ECTS{i}", "") != "":
                        placeholders[f"etat{i}"] = "C"
            else:
                placeholders["etatUE1"] = "NV"
                for i, note in enumerate(ue1_notes, start=1):
                    if note is not None and i not in case_config["hidden_ects"] and placeholders.get(f"ECTS{i}", "") != "":
                        if note < 8:
                            placeholders[f"etat{i}"] = "R"
                        elif 8 <= note < 10:
                            placeholders[f"etat{i}"] = "R" if notes_below_8 > 0 or notes_between_8_and_10 > 1 else "C"
                        else:
                            placeholders[f"etat{i}"] = ""
                    else:
                        placeholders[f"etat{i}"] = ""
        else:
            placeholders["etatUE1"] = ""
            for i in range(1, 8):
                placeholders[f"etat{i}"] = ""
    # Cas BG_TP_2
    elif case_config["key"] == "BG_TP_2":
        process_and_evaluate_ue(placeholders, "UE1", [1, 2, 3], case_config["grade_column_indices"], student_data, case_config)

        # Gestion des notes et des états pour UE1
        ue1_notes = [float(placeholders[f"note{i}"]) if placeholders[f"note{i}"] and placeholders[f"note{i}"] != "" and i not in case_config["hidden_ects"] and placeholders.get(f"ECTS{i}", "") != "" else None for i in range(1, 4)]
        placeholders["etatUE1"] = ""
        for i in range(1, 4):
            placeholders[f"etat{i}"] = ""

        if any(note is not None for note in ue1_notes):
            notes_between_8_and_10 = sum(8 <= note < 10 for note in ue1_notes if note is not None)
            notes_below_8 = sum(note < 8 for note in ue1_notes if note is not None)

            if all(note >= 10 for note in ue1_notes if note is not None):
                placeholders["etatUE1"] = "VA"
            elif notes_between_8_and_10 == 1 and notes_below_8 == 0:
                placeholders["etatUE1"] = "VA"
                for i, note in enumerate(ue1_notes, start=1):
                    if note is not None and 8 <= note < 10 and i not in case_config["hidden_ects"] and placeholders.get(f"ECTS{i}", "") != "":
                        placeholders[f"etat{i}"] = "C"
            else:
                placeholders["etatUE1"] = "NV"
                for i, note in enumerate(ue1_notes, start=1):
                    if note is not None and i not in case_config["hidden_ects"] and placeholders.get(f"ECTS{i}", "") != "":
                        if note < 8:
                            placeholders[f"etat{i}"] = "R"
                        elif 8 <= note < 10:
                            placeholders[f"etat{i}"] = "R" if notes_below_8 > 0 or notes_between_8_and_10 > 1 else "C"
                        else:
                            placeholders[f"etat{i}"] = ""
                    else:
                        placeholders[f"etat{i}"] = ""
        else:
            placeholders["etatUE1"] = ""
            for i in range(1, 4):
                placeholders[f"etat{i}"] = ""
    # Cas BG_TP_3
    elif case_config["key"] == "BG_TP_3":
        process_and_evaluate_ue(placeholders, "UE1", [1, 2, 3, 4, 5], case_config["grade_column_indices"], student_data, case_config)
        process_and_evaluate_ue(placeholders, "UE2", [6, 7, 8, 9, 10], case_config["grade_column_indices"], student_data, case_config)
        process_and_evaluate_ue(placeholders, "UE3", [11, 12, 13], case_config["grade_column_indices"], student_data, case_config)
        process_UE4_and_evaluate_ue(placeholders, "UE4", [14, 15], case_config["grade_column_indices"], student_data, case_config)

        # Gestion des notes et des états pour UE1
        ue1_notes = [float(placeholders[f"note{i}"]) if placeholders[f"note{i}"] and placeholders[f"note{i}"] != "" and i not in case_config["hidden_ects"] and placeholders.get(f"ECTS{i}", "") != "" else None for i in range(1, 6)]
        placeholders["etatUE1"] = ""
        for i in range(1, 6):
            placeholders[f"etat{i}"] = ""

        if any(note is not None for note in ue1_notes):
            notes_between_8_and_10 = sum(8 <= note < 10 for note in ue1_notes if note is not None)
            notes_below_8 = sum(note < 8 for note in ue1_notes if note is not None)

            if all(note >= 10 for note in ue1_notes if note is not None):
                placeholders["etatUE1"] = "VA"
            elif notes_between_8_and_10 == 1 and notes_below_8 == 0:
                placeholders["etatUE1"] = "VA"
                for i, note in enumerate(ue1_notes, start=1):
                    if note is not None and 8 <= note < 10 and i not in case_config["hidden_ects"] and placeholders.get(f"ECTS{i}", "") != "":
                        placeholders[f"etat{i}"] = "C"
            else:
                placeholders["etatUE1"] = "NV"
                for i, note in enumerate(ue1_notes, start=1):
                    if note is not None and i not in case_config["hidden_ects"] and placeholders.get(f"ECTS{i}", "") != "":
                        if note < 8:
                            placeholders[f"etat{i}"] = "R"
                        elif 8 <= note < 10:
                            placeholders[f"etat{i}"] = "R" if notes_below_8 > 0 or notes_between_8_and_10 > 1 else "C"
                        else:
                            placeholders[f"etat{i}"] = ""
                    else:
                        placeholders[f"etat{i}"] = ""
        else:
            placeholders["etatUE1"] = ""
            for i in range(1, 6):
                placeholders[f"etat{i}"] = ""       
    
    # Cas BG_TP_4
    elif case_config["key"] == "BG_TP_4":
        process_and_evaluate_ue(placeholders, "UE1", [1], case_config["grade_column_indices"], student_data, case_config)

        # Gestion des notes et des états pour UE1
        ue1_notes = [float(placeholders[f"note{i}"]) if placeholders[f"note{i}"] and placeholders[f"note{i}"] != "" and i not in case_config["hidden_ects"] and placeholders.get(f"ECTS{i}", "") != "" else None for i in range(1, 2)]
        placeholders["etatUE1"] = ""
        for i in range(1, 2):
            placeholders[f"etat{i}"] = ""

        if any(note is not None for note in ue1_notes):
            notes_between_8_and_10 = sum(8 <= note < 10 for note in ue1_notes if note is not None)
            notes_below_8 = sum(note < 8 for note in ue1_notes if note is not None)

            if all(note >= 10 for note in ue1_notes if note is not None):
                placeholders["etatUE1"] = "VA"
            elif notes_between_8_and_10 == 1 and notes_below_8 == 0:
                placeholders["etatUE1"] = "VA"
                for i, note in enumerate(ue1_notes, start=1):
                    if note is not None and 8 <= note < 10 and i not in case_config["hidden_ects"] and placeholders.get(f"ECTS{i}", "") != "":
                        placeholders[f"etat{i}"] = "C"
            else:
                placeholders["etatUE1"] = "NV"
                for i, note in enumerate(ue1_notes, start=1):
                    if note is not None and i not in case_config["hidden_ects"] and placeholders.get(f"ECTS{i}", "") != "":
                        if note < 8:
                            placeholders[f"etat{i}"] = "R"
                        elif 8 <= note < 10:
                            placeholders[f"etat{i}"] = "R" if notes_below_8 > 0 or notes_between_8_and_10 > 1 else "C"
                        else:
                            placeholders[f"etat{i}"] = ""
                    else:
                        placeholders[f"etat{i}"] = ""
        else:
            placeholders["etatUE1"] = ""
            for i in range(1, 2):
                placeholders[f"etat{i}"] = ""
    # Cas BG_TP_5
    elif case_config["key"] == "BG_TP_5":
        process_and_evaluate_ue(placeholders, "UE1", [1, 2, 3, 4, 5], case_config["grade_column_indices"], student_data, case_config)
        process_and_evaluate_ue(placeholders, "UE2", [6, 7, 8, 9, 10], case_config["grade_column_indices"], student_data, case_config)
        process_and_evaluate_ue(placeholders, "UE3", [11, 12, 13], case_config["grade_column_indices"], student_data, case_config)
        process_UE4_and_evaluate_ue(placeholders, "UE4", [14, 15, 16, 17, 18, 19], case_config["grade_column_indices"], student_data, case_config)

        # Gestion des notes et des états pour UE1
        ue1_notes = [float(placeholders[f"note{i}"]) if placeholders[f"note{i}"] and placeholders[f"note{i}"] != "" and i not in case_config["hidden_ects"] and placeholders.get(f"ECTS{i}", "") != "" else None for i in range(1, 6)]
        placeholders["etatUE1"] = ""
        for i in range(1, 6):
            placeholders[f"etat{i}"] = ""

        if any(note is not None for note in ue1_notes):
            notes_between_8_and_10 = sum(8 <= note < 10 for note in ue1_notes if note is not None)
            notes_below_8 = sum(note < 8 for note in ue1_notes if note is not None)

            if all(note >= 10 for note in ue1_notes if note is not None):
                placeholders["etatUE1"] = "VA"
            elif notes_between_8_and_10 == 1 and notes_below_8 == 0:
                placeholders["etatUE1"] = "VA"
                for i, note in enumerate(ue1_notes, start=1):
                    if note is not None and 8 <= note < 10 and i not in case_config["hidden_ects"] and placeholders.get(f"ECTS{i}", "") != "":
                        placeholders[f"etat{i}"] = "C"
            else:
                placeholders["etatUE1"] = "NV"
                for i, note in enumerate(ue1_notes, start=1):
                    if note is not None and i not in case_config["hidden_ects"] and placeholders.get(f"ECTS{i}", "") != "":
                        if note < 8:
                            placeholders[f"etat{i}"] = "R"
                        elif 8 <= note < 10:
                            placeholders[f"etat{i}"] = "R" if notes_below_8 > 0 or notes_between_8_and_10 > 1 else "C"
                        else:
                            placeholders[f"etat{i}"] = ""
                    else:
                        placeholders[f"etat{i}"] = ""
        else:
            placeholders["etatUE1"] = ""
            for i in range(1, 6):
                placeholders[f"etat{i}"] = ""
    # Cas BG_TP_6
    elif case_config["key"] == "BG_TP_6":
        process_and_evaluate_ue(placeholders, "UE1", [1, 2, 3], case_config["grade_column_indices"], student_data, case_config)

        # Gestion des notes et des états pour UE1
        ue1_notes = [float(placeholders[f"note{i}"]) if placeholders[f"note{i}"] and placeholders[f"note{i}"] != "" and i not in case_config["hidden_ects"] and placeholders.get(f"ECTS{i}", "") != "" else None for i in range(1, 4)]
        placeholders["etatUE1"] = ""
        for i in range(1, 4):
            placeholders[f"etat{i}"] = ""

        if any(note is not None for note in ue1_notes):
            notes_between_8_and_10 = sum(8 <= note < 10 for note in ue1_notes if note is not None)
            notes_below_8 = sum(note < 8 for note in ue1_notes if note is not None)

            if all(note >= 10 for note in ue1_notes if note is not None):
                placeholders["etatUE1"] = "VA"
            elif notes_between_8_and_10 == 1 and notes_below_8 == 0:
                placeholders["etatUE1"] = "VA"
                for i, note in enumerate(ue1_notes, start=1):
                    if note is not None and 8 <= note < 10 and i not in case_config["hidden_ects"] and placeholders.get(f"ECTS{i}", "") != "":
                        placeholders[f"etat{i}"] = "C"
            else:
                placeholders["etatUE1"] = "NV"
                for i, note in enumerate(ue1_notes, start=1):
                    if note is not None and i not in case_config["hidden_ects"] and placeholders.get(f"ECTS{i}", "") != "":
                        if note < 8:
                            placeholders[f"etat{i}"] = "R"
                        elif 8 <= note < 10:
                            placeholders[f"etat{i}"] = "R" if notes_below_8 > 0 or notes_between_8_and_10 > 1 else "C"
                        else:
                            placeholders[f"etat{i}"] = ""
                    else:
                        placeholders[f"etat{i}"] = ""
        else:
            placeholders["etatUE1"] = ""
            for i in range(1, 4):
                placeholders[f"etat{i}"] = ""
    
    total_ects = 0  # Initialiser le total des ECTS

    # Dans la fonction generate_word_document, modifiez la partie qui traite les notes
    for i, col_index in enumerate(case_config["grade_column_indices"], start=1):
        grade_str = str(student_data.iloc[col_index]).strip() if pd.notna(student_data.iloc[col_index]) else ""
        if grade_str and grade_str != 'Note':
            grades_coefficients, special_case = extract_grades_and_coefficients(grade_str)
            if special_case:
                placeholders[f"note{i}"] = special_case
                placeholders[f"ECTS{i}"] = ""  # Ne pas attribuer d'ECTS pour les cas spéciaux
            else:
                logger.warning(f"Invalid column index: {col_index} for student {student_data['Nom']}")
                individual_average = calculate_weighted_average([g[0] for g in grades_coefficients], [g[1] for g in grades_coefficients])
                placeholders[f"note{i}"] = f"{individual_average:.2f}" if individual_average else ""
                if individual_average > 8 and i not in case_config["hidden_ects"]:
                    ects_value = int(ects_data.get(f"ECTS{i}", 1))
                    placeholders[f"ECTS{i}"] = ects_value
                elif individual_average > 0:
                    placeholders[f"ECTS{i}"] = 0
                else:
                    placeholders[f"ECTS{i}"] = ""
        else:
            placeholders[f"note{i}"] = ""
            placeholders[f"ECTS{i}"] = ""

    # Calcul correct des ECTS pour chaque UE
    for ue, indices in case_config["ects_sum_indices"].items():
        ue_sum = 0
        ue_ects = 0
        valid_notes_count = 0  # Initialiser valid_notes_count ici

        for index in indices:
            note_str = placeholders[f"note{index}"]
            if note_str not in ["Validé ( - ASE)", "Non Validé ( - ASE)"] and not note_str.endswith("(CCHM)"):
                try:
                    note = float(note_str) if note_str not in ["", None] else 0
                    ects = int(placeholders[f"ECTS{index}"]) if placeholders[f"ECTS{index}"] not in ["", None] else 0
                    if ects != 0:
                        ue_sum += note * ects
                        ue_ects += ects
                        valid_notes_count += 1
                except ValueError:
                    continue

        average_ue = math.ceil(ue_sum / ue_ects * 100) / 100 if ue_ects > 0 else 0
        placeholders[f"moy{ue}"] = f"{average_ue:.2f}" if average_ue and valid_notes_count > 0 else ""
        placeholders[f"ECTS{ue}"] = ue_ects if ue_ects else ""

    # Calcul correct du total des ECTS
    total_ects = sum(int(placeholders[f"ECTS{ue}"]) for ue in case_config["ects_sum_indices"].keys() if placeholders[f"ECTS{ue}"] not in ["", None])
    placeholders["moyenneECTS"] = total_ects

    placeholders["moyenneECTS"] = total_ects
    
    # Après le traitement de toutes les notes et ECTS, ajoutez cette nouvelle boucle :
    for i in range(1, len(case_config["grade_column_indices"]) + 1):
        if placeholders.get(f"etat{i}") == "R":
            placeholders[f"ECTS{i}"] = 0
            
    for ue, indices in case_config["ects_sum_indices"].items():
        ue_ects = sum(int(placeholders[f"ECTS{index}"]) for index in indices if placeholders[f"ECTS{index}"] not in ["", None])
        placeholders[f"ECTS{ue}"] = ue_ects

    # Recalcul du total des ECTS après application de la règle
    total_ects = sum(int(placeholders[f"ECTS{ue}"]) for ue in case_config["ects_sum_indices"].keys() if placeholders[f"ECTS{ue}"] not in ["", None])
    placeholders["moyenneECTS"] = total_ects
    
    all_ue_states = [placeholders.get(f"etat{ue}") for ue in case_config["ects_sum_indices"].keys()]
    
    if all(state == "VA" for state in all_ue_states if state):
        placeholders["totaletat"] = "VA"
    elif any(state == "NV" for state in all_ue_states if state):
        placeholders["totaletat"] = "NV"
    else:
        placeholders["totaletat"] = ""  # Au cas où il n'y aurait pas d'états définis
    
    # Calcul de la moyenne générale en fonction des moyennes des UE
    total_ue_notes = sum(
        float(placeholders[f"moy{ue}"]) * int(placeholders[f"ECTS{ue}"])
        for ue in case_config["ects_sum_indices"].keys()
        if placeholders[f"moy{ue}"] not in ["", None] and placeholders[f"ECTS{ue}"] not in ["", 0, None]
    )
    total_ue_ects = sum(
        int(placeholders[f"ECTS{ue}"])
        for ue in case_config["ects_sum_indices"].keys()
        if placeholders[f"ECTS{ue}"] not in ["", 0, None]
    )

    # Calcul de la moyenne générale arrondie au centième près
    placeholders["moyenne"] = f"{math.ceil(total_ue_notes / total_ue_ects * 100) / 100:.2f}" if total_ue_ects else 0
    

    # Supprimer les placeholders pour les ECTS masqués du document final
    for hidden_ects in case_config["hidden_ects"]:
        placeholders.pop(f"ECTS{hidden_ects}", None)

    logger.debug(f"Placeholders: {placeholders}")  # Log des placeholders pour vérifier leurs valeurs

    doc = DocxTemplate(template_path)
    doc.render(placeholders)
    # Hide the 'CodeApprenant' identifier in the document
    for paragraph in doc.paragraphs:
        if 'Identifiant' in paragraph.text:
            set_hidden_text(paragraph)  # Hide the text for 'Identifiant'

    output_filename = f"{normalize_string(student_data['Nom'])}_bulletin.docx"
    output_filepath = os.path.join(output_dir, output_filename)
    doc.save(output_filepath)
    return output_filepath