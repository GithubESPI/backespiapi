from pydantic_settings import BaseSettings
import os

class Settings(BaseSettings):
    PROJECT_NAME: str = "Upload de Bulletins"
    BASE_DIR: str = os.getenv('BASE_DIR', '/code')

    # DOCUMENTS_DIR: str = os.path.join(os.getenv('USERPROFILE', os.getenv('HOME')), 'Documents')
    DOCUMENTS_DIR: str = os.path.join(BASE_DIR, 'documents')
    OUTPUT_DIR: str = os.path.join(DOCUMENTS_DIR, "outputs")
    TEMPLATE_FILE: str = os.path.join(BASE_DIR, "template", "modeleM1S1.docx")

    #DOWNLOAD_DIR: str = os.path.join(os.getenv('USERPROFILE', os.getenv('HOME')), 'Downloads')
    DOWNLOAD_DIR: str = os.path.join(BASE_DIR, 'downloads')
    
    # M1-S1 excel empty
    M1_S1_MAPI_TEMPLATE: str = os.path.join(BASE_DIR, "excel", "M1-S1", "M1-S1-MAPI.xlsx")
    M1_S1_MAGI_TEMPLATE: str = os.path.join(BASE_DIR, "excel", "M1-S1", "M1-S1-MAGI.xlsx")
    M1_S1_MEFIM_TEMPLATE: str = os.path.join(BASE_DIR, "excel", "M1-S1", "M1-S1-MEFIM.xlsx")
    # M1-S1 excel not empty
    M1_S1_MAPI_TEMPLATE_NOT_EMPTY: str = os.path.join(BASE_DIR, "template", "S1", "M1-S1-MAPI.xlsx")
    M1_S1_MAGI_TEMPLATE_NOT_EMPTY: str = os.path.join(BASE_DIR, "template", "S1", "M1-S1-MAGI.xlsx")
    M1_S1_MEFIM_TEMPLATE_NOT_EMPTY: str = os.path.join(BASE_DIR, "template", "S1", "M1-S1-MEFIM.xlsx")
    # M1-S1 excel bulletin
    M1_S1_MAPI_TEMPLATE_WORD: str = os.path.join(BASE_DIR, "template", "modeleM1S1.docx")
    M1_S1_MAGI_TEMPLATE_WORD: str = os.path.join(BASE_DIR, "template", "modeleM1S1.docx")
    M1_S1_MEFIM_TEMPLATE_WORD: str = os.path.join(BASE_DIR, "template", "modeleM1S1.docx")

    # M1-S2
    M1_S2_MAPI_TEMPLATE: str = os.path.join(BASE_DIR, "excel", "M1-S2", "M1-S2-MAPI.xlsx")
    M1_S2_MAGI_TEMPLATE: str = os.path.join(BASE_DIR, "excel", "M1-S2", "M1-S2-MAGI.xlsx")
    M1_S2_MEFIM_TEMPLATE: str = os.path.join(BASE_DIR, "excel", "M1-S2", "M1-S2-MEFIM.xlsx")
    M1_S2_MAPI_TEMPLATE_NOT_EMPTY: str = os.path.join(BASE_DIR, "template", "S1", "M1-S2-MAPI.xlsx")
    M1_S2_MAGI_TEMPLATE_NOT_EMPTY: str = os.path.join(BASE_DIR, "template", "S1", "M1-S2-MAGI.xlsx")
    M1_S2_MEFIM_TEMPLATE_NOT_EMPTY: str = os.path.join(BASE_DIR, "template", "S1", "M1-S2-MEFIM.xlsx")
    M1_S2_MAPI_TEMPLATE_WORD: str = os.path.join(BASE_DIR, "template", "modeleM1S2.docx")
    M1_S2_MAGI_TEMPLATE_WORD: str = os.path.join(BASE_DIR, "template", "modeleM1S2.docx")
    M1_S2_MEFIM_TEMPLATE_WORD: str = os.path.join(BASE_DIR, "template", "modeleM1S2.docx")

    # M2-S3
    M2_S3_MAPI_TEMPLATE: str = os.path.join(BASE_DIR, "excel", "M2-S3", "M2-S3-MAPI.xlsx")
    M2_S3_MAGI_TEMPLATE: str = os.path.join(BASE_DIR, "excel", "M2-S3", "M2-S3-MAGI.xlsx")
    M2_S3_MEFIM_TEMPLATE: str = os.path.join(BASE_DIR, "excel", "M2-S3", "M2-S3-MEFIM.xlsx")
    M2_S3_MAPI_TEMPLATE_NOT_EMPTY: str = os.path.join(BASE_DIR, "template", "S1", "M2-S3-MAPI.xlsx")
    M2_S3_MAGI_TEMPLATE_NOT_EMPTY: str = os.path.join(BASE_DIR, "template", "S1", "M2-S3-MAGI.xlsx")
    M2_S3_MEFIM_TEMPLATE_NOT_EMPTY: str = os.path.join(BASE_DIR, "template", "S1", "M2-S3-MEFIM.xlsx")
    M2_S3_MAPI_TEMPLATE_WORD: str = os.path.join(BASE_DIR, "template", "modeleM2S3MAPI.docx")
    M2_S3_MAGI_TEMPLATE_WORD: str = os.path.join(BASE_DIR, "template", "modeleM2S3.docx")
    M2_S3_MEFIM_TEMPLATE_WORD: str = os.path.join(BASE_DIR, "template", "modeleM2S3.docx")

    # M2-S4
    M2_S4_MAPI_TEMPLATE: str = os.path.join(BASE_DIR, "excel", "M2-S4", "M2-S4-MAPI.xlsx")
    M2_S4_MAGI_TEMPLATE: str = os.path.join(BASE_DIR, "excel", "M2-S4", "M2-S4-MAGI.xlsx")
    M2_S4_MEFIM_TEMPLATE: str = os.path.join(BASE_DIR, "excel", "M2-S4", "M2-S4-MEFIM.xlsx")
    M2_S4_MAPI_TEMPLATE_NOT_EMPTY: str = os.path.join(BASE_DIR, "template", "S1", "M2-S4-MAPI.xlsx")
    M2_S4_MAGI_TEMPLATE_NOT_EMPTY: str = os.path.join(BASE_DIR, "template", "S1", "M2-S4-MAGI.xlsx")
    M2_S4_MEFIM_TEMPLATE_NOT_EMPTY: str = os.path.join(BASE_DIR, "template", "S1", "M2-S4-MEFIM.xlsx")
    M2_S4_MAPI_TEMPLATE_WORD: str = os.path.join(BASE_DIR, "template", "modeleM2S4.docx")
    M2_S4_MAGI_TEMPLATE_WORD: str = os.path.join(BASE_DIR, "template", "modeleM2S4.docx")
    M2_S4_MEFIM_TEMPLATE_WORD: str = os.path.join(BASE_DIR, "template", "modeleM2S4.docx")
    
    # BG ALT excel empty
    BG_ALT_1_TEMPLATE: str = os.path.join(BASE_DIR, "excel", "BG-ALT-S1", "BG-ALT-S1.xlsx")
    BG_ALT_2_TEMPLATE: str = os.path.join(BASE_DIR, "excel", "BG-ALT-S2", "BG-ALT-S2.xlsx")
    BG_ALT_3_TEMPLATE: str = os.path.join(BASE_DIR, "excel", "BG-ALT-S3", "BG-ALT-S3.xlsx")
    BG_ALT_4_TEMPLATE: str = os.path.join(BASE_DIR, "excel", "BG-ALT-S4", "BG-ALT-S4.xlsx")
    BG_ALT_5_TEMPLATE: str = os.path.join(BASE_DIR, "excel", "BG-ALT-S5", "BG-ALT-S5.xlsx")
    BG_ALT_6_TEMPLATE: str = os.path.join(BASE_DIR, "excel", "BG-ALT-S6", "BG-ALT-S6.xlsx")
    
    # BG ALT  excel not empty
    BG_ALT_1_TEMPLATE_NOT_EMPTY: str = os.path.join(BASE_DIR, "template", "S1", "BG-ALT-S1.xlsx")
    BG_ALT_2_TEMPLATE_NOT_EMPTY: str = os.path.join(BASE_DIR, "template", "S1", "BG-ALT-S2.xlsx")
    BG_ALT_3_TEMPLATE_NOT_EMPTY: str = os.path.join(BASE_DIR, "template", "S1", "BG-ALT-S3.xlsx")
    BG_ALT_4_TEMPLATE_NOT_EMPTY: str = os.path.join(BASE_DIR, "template", "S1", "BG-ALT-S4.xlsx")
    BG_ALT_5_TEMPLATE_NOT_EMPTY: str = os.path.join(BASE_DIR, "template", "S1", "BG-ALT-S5.xlsx")
    BG_ALT_6_TEMPLATE_NOT_EMPTY: str = os.path.join(BASE_DIR, "template", "S1", "BG-ALT-S6.xlsx")
    
    # BG ALT  Word bulletin
    BG_ALT_1_TEMPLATE_WORD: str = os.path.join(BASE_DIR, "template", "modeleBGALT1.docx")
    BG_ALT_2_TEMPLATE_WORD: str = os.path.join(BASE_DIR, "template", "modeleBGALT2.docx")
    BG_ALT_3_TEMPLATE_WORD: str = os.path.join(BASE_DIR, "template", "modeleBGALT3.docx")
    BG_ALT_4_TEMPLATE_WORD: str = os.path.join(BASE_DIR, "template", "modeleBGALT4.docx")
    BG_ALT_5_TEMPLATE_WORD: str = os.path.join(BASE_DIR, "template", "modeleBGALT5.docx")
    BG_ALT_6_TEMPLATE_WORD: str = os.path.join(BASE_DIR, "template", "modeleBGALT6.docx")
    
    # BG TP excel empty
    BG_TP_1_TEMPLATE: str = os.path.join(BASE_DIR, "excel", "BG-TP-S1", "BG-TP-S1.xlsx")
    BG_TP_2_TEMPLATE: str = os.path.join(BASE_DIR, "excel", "BG-TP-S2", "BG-TP-S2.xlsx")
    BG_TP_3_TEMPLATE: str = os.path.join(BASE_DIR, "excel", "BG-TP-S3", "BG-TP-S3.xlsx")
    BG_TP_4_TEMPLATE: str = os.path.join(BASE_DIR, "excel", "BG-TP-S4", "BG-TP-S4.xlsx")
    BG_TP_5_TEMPLATE: str = os.path.join(BASE_DIR, "excel", "BG-TP-S5", "BG-TP-S5.xlsx")
    BG_TP_6_TEMPLATE: str = os.path.join(BASE_DIR, "excel", "BG-TP-S6", "BG-TP-S6.xlsx")
    
    # BG TP  excel not empty
    BG_TP_1_TEMPLATE_NOT_EMPTY: str = os.path.join(BASE_DIR, "template", "S1", "BG-TP-S1.xlsx")
    BG_TP_2_TEMPLATE_NOT_EMPTY: str = os.path.join(BASE_DIR, "template", "S1", "BG-TP-S2.xlsx")
    BG_TP_3_TEMPLATE_NOT_EMPTY: str = os.path.join(BASE_DIR, "template", "S1", "BG-TP-S3.xlsx")
    BG_TP_4_TEMPLATE_NOT_EMPTY: str = os.path.join(BASE_DIR, "template", "S1", "BG-TP-S4.xlsx")
    BG_TP_5_TEMPLATE_NOT_EMPTY: str = os.path.join(BASE_DIR, "template", "S1", "BG-TP-S5.xlsx")
    BG_TP_6_TEMPLATE_NOT_EMPTY: str = os.path.join(BASE_DIR, "template", "S1", "BG-TP-S6.xlsx")
    
    # BG ALT  Word bulletin
    BG_TP_1_TEMPLATE_WORD: str = os.path.join(BASE_DIR, "template", "modeleBGTP1.docx")
    BG_TP_2_TEMPLATE_WORD: str = os.path.join(BASE_DIR, "template", "modeleBGTP2.docx")
    BG_TP_3_TEMPLATE_WORD: str = os.path.join(BASE_DIR, "template", "modeleBGTP3.docx")
    BG_TP_4_TEMPLATE_WORD: str = os.path.join(BASE_DIR, "template", "modeleBGTP4.docx")
    BG_TP_5_TEMPLATE_WORD: str = os.path.join(BASE_DIR, "template", "modeleBGTP5.docx")
    BG_TP_6_TEMPLATE_WORD: str = os.path.join(BASE_DIR, "template", "modeleBGTP6.docx")
    

    # ECTS
    ECTS_JSON_PATH: str = os.path.join(BASE_DIR, "json", "ects.json")

    RELEVANT_GROUPS: list = [
        "N-M1 MAPI ALT 1", "P-M1 MAPI ALT 2", "L-M1 MAPI ALT 2", "MP-M1 MAPI ALT",
        "P-M1 MAPI ALT 5", "L-M1 MAPI ALT 1", "P-M1 MAPI ALT 1", "P-M1 MAPI ALT 3",
        "B-M1 MAPI ALT 1", "M-M1 MAPI ALT 1", "LI-M1 MAPI ALT", "N-M1 MAPI ALT 2",
        "M-M1 MAPI ALT 2", "P-M1 MAPI ALT 4", "B-M1 MAPI ALT 2", "MP-M1 MAPI ALT",
        "L-M1 MAPI ALT 3", "P-M1 MAGI ALT 1", "N-M1 MAGI ALT", "M-M1 MAGI ALT",
        "LI-M1 MAGI ALT", "B-M1 MAGI ALT", "MP-M1 MAGI ALT", "L-M1 MAGI ALT",
        "P-M1 MAGI ALT 2", "LI-M1 MAGI ALT", "P-M1 MAGI ALT 2", "M-M1 MIFIM ALT",
        "N-M1 MIFIM ALT", "P-M1 MIFIM ALT 1", "P-M1 MIFIM ALT 2", "P-M1 MIFIM ALT 3",
        "LI-M1 MIFIM ALT", "B-M1 MIFIM ALT", "MP-M1 MIFIM ALT", "L-M1 MIFIM ALT"
    ]
    RELEVANT_GROUPS_M2: list = [
        "L-M2 MAPI ALT 1", "N-M2 MAGI ALT", "P-M2 MAPI ALT 3", "B-M2 MAGI ALT",
        "P-M2 MAPI ALT 5", "N-M2 MAPI ALT 2", "B-M2 MAPI ALT 1", "P-M2 MAGI ALT 2",
        "M-M2 MAPI ALT 2", "P-M2 MAPI ALT 1", "M-M2 MAPI ALT 1", "L-M2 MAPI ALT 2",
        "P-M2 MAPI ALT 2", "P-M2 MAPI ALT 4", "N-M2 MAPI ALT 1", "L-M2 MAGI ALT",
        "P-M2 MAGI ALT 1", "M-M2 MAGI ALT", "LI-M2 MAPI ALT", "M-M2 MAPI ALT 3",
        "M-M2 2ESI ALT", "N-M2 2ESI ALT", "N-M2 MIFIM ALT", "P-M2 2ESI ALT",
        "P-M2 MIFIM ALT 1", "P-M2 MIFIM ALT 2", "P-M2 MIFIM ALT 3", "M-M2 MIFIM ALT",
        "MP-M2 MAGI ALT", "MP-M2 MAPI ALT 1", "MP-M2 MAPI ALT 2", "B-M2 2ESI ALT",
        "B-M2 MIFIM ALT", "B-M2 MAPI ALT 2", "L-M2 MIFIM ALT", "L-M2 2ESI ALT",
        "P-M2 MAPI RP", "P-M2 MIFIM RP", "P-M2 MAGI RP", "CA-M2 MIFIM TP", "CA-M2 MAPI TP", "N-M2 MAGI ALT 1"
    ]
    RELEVANT_GROUPS_TP: list = ["B-BG1 TP", "L-BG1 TP", "M-BG1 TP", "N-BG1 TP", "P-BG1 TP 1", "P-BG1 TP 2 Rentrée décalée", ]
    RELEVANT_GROUPS_TP_2: list = ["P-BG2 TP"]
    RELEVANT_GROUPS_TP_3: list = ["P-BG3 TP 1", "P-BG3 TP 2 Rentrée décalée", "P-BG3 TP Section Internationale"]
    RELEVANT_GROUPS_ALT: list = ["GE-BG1 ALT 1 TEST","L-BG1 ALT 1", "L-BG1 ALT 2", "LI-BG1 ALT", "M-BG1 ALT", "N-BG1 ALT",   "P-BG1 ALT 1", "P-BG1 ALT 2", "P-BG1 ALT 3"]
    RELEVANT_GROUPS_ALT_2: list = ["B-BG2 ALT", "L-BG2 ALT", "M-BG2 ALT", "MP-BG2 ALT", "N-BG2 ALT","P-BG2 ALT 1", "P-BG2 ALT 2", "P-BG2 ALT 3", "P-BG2 ALT 4"]
    RELEVANT_GROUPS_ALT_3: list = [ "B-BG3 ALT 1", "B-BG3 ALT 2", "L-BG3 ALT 1", "L-BG3 ALT 2",  "LI-BG3 ALT",  
    "M-BG3 ALT 1", "M-BG3 ALT 2", "M-BG3 ALT 3", "M-BG3 ALT 3", "MP-BG1 ALT",  "MP-BG3 ALT", "N-BG3 ALT 1", "N-BG3 ALT 2", "N-BG3 ALT 3",  "P-BG3 ALT 1",
        "P-BG3 ALT 2", "P-BG3 ALT 3", "P-BG3 ALT 4", "P-BG3 ALT 5", "P-BG3 ALT 6", "P-BG3 ALT 7", "P-BG3 ALT 8"]
    
    # Paramètres d'API externe
    YPAERO_BASE_URL: str
    YPAERO_API_TOKEN: str
    BASE_DIR: str

    class Config:
        # Chargez les variables d'environnement à partir d'un fichier .env situé à la racine du projet.
        env_file = ".env"

# Instanciez les paramètres pour qu'ils soient importés et utilisés dans d'autres fichiers
settings = Settings()

