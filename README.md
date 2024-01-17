# Maude DB Analysis Project

## Overview
The MaudeAnalysis project is a Python-based tool designed to analyze and process medical device-related data, particularly focusing on the Manufacturer and User Facility Device Experience (MAUDE) data from the Food and Drug Administration (FDA). This script automates the analysis of medical data, identifying potential causes of incidents and extracting key information to aid in healthcare decision-making.



## Features
- **Data Import**: Opens and reads data from specified Excel files, focusing on key columns like Manufacturer, Product Code, Brand Name, and Event Text.
- **Natural Language Processing**: Utilizes `spaCy` for tokenization of event texts, facilitating the analysis of medical data.
- **Data Matching**: Employs `difflib` for sequence matching, identifying relevant instruments based on manufacturer and product codes.
- **Sentiment Analysis**: Uses `TextBlob` to evaluate the sentiment of text surrounding specific terms, aiding in the interpretation of the data.
- **Data Processing and Analysis**: The script processes the data to identify listed analytes and potential causes, filtering and categorizing entries based on predefined criteria.
- **Output Generation**: Exports the analyzed and processed data into a user-defined Excel file.

## Usage
1. Run the script to initiate the file dialog for selecting MAUDE data in Excel format.
2. Input the desired name for the output file when prompted.
3. The script processes the data and saves it to the specified output file.

## Requirements
- Python
- Libraries: `spacy`, `pandas`, `numpy`, `textblob`, `tkinter`, `re`
- Pre-downloaded `en_core_web_lg` model for `spaCy`.

## Note
This project is designed to simplify and automate the analysis of medical data, contributing to more efficient healthcare data management and decision-making processes.
