import spacy
import pandas as pd
import numpy as np
import string
import time
import re
import en_core_web_lg
from difflib import SequenceMatcher, get_close_matches
from textblob import TextBlob
import tkinter as tk
from tkinter import filedialog


def open_file_dialog():
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xls;*.xlsx")])
    return file_path


class MaudeAnalysis:
    def __init__(self, input_file, output_file, manufacturers, product_codes, causes_dict):
        self.input_file = input_file
        self.output_file = output_file
        self.manufacturers = manufacturers
        self.product_codes = product_codes
        self.causes_dict = causes_dict
        self.df = None
        self.filtered_df = None
        self.nlp = spacy.load('en_core_web_lg')

    def read_data(self):
        excel_file = self.input_file
        self.df = pd.read_excel(excel_file, usecols=['Manufacturer', 'Product Code', ' Brand Name', 'Event Text'])
        self.df['tokens'] = self.df['Event Text'].apply(lambda x: [token.text for token in self.nlp(str(x))])

    def find_instrument(self, manufacturer, instruments, product_codes, df):
        indices = []
        for i, row in df.iterrows():
            if any(keyword.lower() in ' '.join(row['tokens']).lower() for cause, keywords in causes_dict.items() for keyword in keywords):
                if manufacturer.lower() in str(row['Manufacturer']).lower() and str(row['Product Code']).lower() in [pc.lower() for pc in product_codes]:
                    brand_name_processed = re.sub(r'[^\w\s]', '', str(row[' Brand Name']).lower()).split()
                    for instrument in instruments:
                        instrument_processed = re.sub(r'[^\w\s]', '', instrument.lower()).split()
                        matcher = SequenceMatcher(None, instrument_processed, brand_name_processed)
                        score = matcher.ratio()
                        if score >= 0.5:
                            indices.append(i)
                            df.at[i, 'Score'] = score
                            break
        return indices

    def process_data(self):
        filtered_df = pd.DataFrame()
        for manufacturer, instruments in manufacturers.items():
            indices = self.find_instrument(manufacturer, instruments, product_codes, self.df)
            manufacturer_df = self.df.loc[indices].copy()
            manufacturer_df = manufacturer_df[manufacturer_df['Product Code'].isin(product_codes)]
            manufacturer_df['Manufacturer'] = manufacturer
            filtered_df = pd.concat([filtered_df, manufacturer_df])

        self.filtered_df = filtered_df

    def compare_tokens_to_dict(self, tokens):
        lower_tokens = [token.lower() for token in tokens]
        for values in causes_dict.values():
            lower_values = [value.lower() for value in values]
            for i in range(len(lower_tokens)):
                for j in range(i, len(lower_tokens)):
                    contiguous_tokens = lower_tokens[i:j+1]
                    contiguous_string = ' '.join(contiguous_tokens)
                    if contiguous_string in lower_values:
                        return True
        return False

    def find_analytes_listed(self, tokens):
        causes = []
        for i, token in enumerate(tokens):
            if token.lower() == "na" and i > 0 and (tokens[i+1].lower() == "measurements" or tokens[i+1].lower() == "measurement"):
                if "Na".lower() not in causes:
                    causes.append("Na".lower())
            elif token.upper() == "NA" and i > 0 and tokens[i+1].lower() == ".":
                if "Not Applicable".lower() not in causes:
                    pass
                
            elif token.lower() == "injury":
                if i > 0 and tokens[i - 1].lower() == "of":
                    continue
                if "no reports of serious injury or death".lower() in " ".join(tokens).lower() or "no reports of death or serious injury".lower() in " ".join(tokens).lower():
                    continue

                surrounding_text = ' '.join(tokens[max(0, i - 5):min(len(tokens), i + 6)])
                sentiment = TextBlob(surrounding_text).sentiment

                if sentiment.polarity <= 0:
                    if "injury".lower() not in causes:
                        causes.append("injury".lower())
        
            elif token.lower() == "death" or token.lower() == "deceased":
                if "no reports of serious injury or death".lower() not in " ".join(tokens).lower() and "the death was not related".lower() not in " ".join(tokens).lower() and "no reports of death or serious injury".lower() not in " ".join(tokens).lower():
                    pass
            elif token.lower() == "deceased":
                if "the death was not related".lower() not in " ".join(tokens).lower():
                    pass
            else:
                for cause, keywords in causes_dict.items():
                    if token.lower() in [keyword.lower() for keyword in keywords]:
                        lower_token = token.lower()
                        if lower_token not in causes:
                            causes.append(lower_token)
                    elif i < len(tokens)-1 and f"{token.lower()} {tokens[i+1].lower()}" in [keyword.lower() for keyword in keywords]:
                        lower_token = f"{token.lower()} {tokens[i+1].lower()}"
                        if lower_token not in causes:
                            causes.append(lower_token)
        return causes

    def find_unknown_causes(self, text):
        unknown_causes = ['unknown', 'no further investigation', 'pending investigation', 'no patient was harmed', 'there was no patient harm', 'there was no harm', 'investigation has been finalized', 'udi', 'unique identifier', 'investigation is underway', 'no information', 'no further issue', 'Na.', 'investigation is ongoing', 'further analysis', 'no further evaluation', 'additional information to be provided', 'no reports of death or serious injury', 'no reports of serious injury or death', 'further details will be provided', 'limited information available',  'could not be determined']
        lowercase_text = text.lower()
        for cause in unknown_causes:
            if cause.lower() in lowercase_text:
                return "unknown"
        return ''

    def find_known_causes(self, text):
        known_causes = ['power supply', 'interferent detected', 'leak', 'human error', 'maintenance', 'clot', 'equipment is burnt', 'hot to touch', 'irenat (perchlorate)']
        lowercase_text = text.lower()
        for cause in known_causes:
            if cause.lower() in lowercase_text:
                return cause
        return ''

    def analyze_data(self):
        self.filtered_df['comparison_result'] = self.filtered_df['tokens'].apply(self.compare_tokens_to_dict)
        self.filtered_df['Analytes listed in Event'] = self.filtered_df['tokens'].apply(self.find_analytes_listed)
        self.filtered_df['root cause'] = self.filtered_df['Event Text'].apply(lambda x: self.find_unknown_causes(x) if self.find_unknown_causes(x) else self.find_known_causes(x))

    def save_data(self):
        self.filtered_df.sort_index()
        self.filtered_df.to_excel(self.output_file, index=False)

    def run(self):
        self.read_data()
        self.process_data()
        self.analyze_data()
        self.save_data()

               
if __name__ == "__main__":
    print("Make sure the sheet that has the MAUDE data is the first sheet in your Excel file.")
    input_file = open_file_dialog()
    output_file_name = input("Please enter the desired name for the output file (including the file extension, e.g., .xlsx): ")
    output_file = output_file_name.strip()

    manufacturers = {
        'SIEMENS HEALTHCARE DIAGNOSTICS INC.': ['RapidPoint 500 BLOOD GAS SYSTEM', 'RAPIDPOINT 500 BLOOD GAS ANALYZER', 'RapidPoint 500E BLOOD GAS SYSTEM', 'Advia 1800', 'Advia Centaur XPT'],
        'ABBOTT Point of Care': ['Alinity', 'I-STAT', 'I-STAT1 ANALYZER, IMMUNO READY, WIRELESS'],
        'EPOCAL INC.': ['epoc reader', 'epoc reader & power supply'],
        'NOVA BIOMEDICAL Corp.': ['Nova Statsensor CREATININE HOSPITAL METER', 'Nova Stat profile prime plus ANALYZER SYSTEM', 'Nova Stat PROFILE PHOX ULTRA ANALYZER SYSTEM'],
        'radiometer medical aps': ['ABL90 FLEX', 'ABL90 FLEX PLUS Analyzer', 'ABL800 FLEX', 'ABL800 FLEX Analyzer'],
        'ROCHE DIAGNOSTICS': ['Roche Omni S', 'OMNI', 'Cobas Integra 400 PLUS', 'COBAS C111 sd', 'Cobas h232', 'ACCUTREND ® TEST STRIPS'],
        'DRUMMOND SCIENTIFIC COMPANY': ['AQUA-CAP'],
        'SENDX MEDICAL, INC.': ['ABL80 FLEX CO-OX'],
        'INSTRUMENTATION LABORATORY COMPANY': ['GEM PREMIER 4000']
    }

    product_codes = ['CDQ', 'CDS', 'CEM', 'CGA', 'CGL', 'CGZ', 'CHL', 'CIG', 'GGZ', 'GHS', 'GIO', 'GKA', 'GKF', 'GKR', 'GLY', 'JFL', 'JFP', 'JFY', 'JGS', 'JJE', 'JJS', 'JPI', 'KHG', 'KHP', 'KHS', 'KQI', 'MQM']

    causes_dict = {
        'analyte': ['Na+', 'Na', 'K', 'K+', 'Cl-', 'Cl', 'iCa', 'ionized calcium', 'egfr', 'Ca', 'Ca++', 'BUN', 'Urea', 'Crea', 'creatinine', 'hematocrit', 'Hct', 'conductivity', 'lactate', 'pCO2', 'CO2', 'carbon dioxide', 'pH', 'pO2', 'O2', 'oxygen', 'TCO2', 'MeasuredTCO2', 'calcium', 'sodium', 'potassium', 'chloride', 'calcium', 'tHb', 'hemoglobin', 'glucose', 'bilirubin'],
        'non-analyte': ['interferent', 'use error', 'cut', 'injury', 'fall', 'shock', 'death', 'deceased', 'fire', 'smoke', 'sparks', 'burn', 'smoking', 'heat', 'hit', 'hot', 'barcode', 'uncalibration', 'calibration failures']
    }

    maude_analysis = MaudeAnalysis(input_file, output_file, manufacturers, product_codes, causes_dict)
    maude_analysis.run()

