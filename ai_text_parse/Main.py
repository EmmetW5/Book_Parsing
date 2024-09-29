
import os
from progress.bar import Bar

from AIParser import AIParser
from ExcelParser import ExcelParser

# Prompt for the model, this is the instructions for the model to follow, it is rather large and inflates the cost since 
# it is charged per token. We may want to consider a more efficient way to do this, but for now it works.
my_prompt_us = f""" 
    Classify and structure the given text according to the following categories:

    <Company Name> the name of the company, will be in all caps
    <Executive Office> the address of the executive office, the address will follow "EXECUTIVE OFFICE: " text
    <Regional Sales> the addresses of regional sales offices of the company. They will follow "REGIONAL SALES: " text, and can have many entries
    <Sales Office> the address of the sales office, will follow "SALES OFFICE: " text
    <Fiber> the fibers the company produces. They will follow after the last regional sales office, and can have many entries. 
    <Plant Location> After each fiber, there will be a list of locations where they are produced. There can be more than one location for each fiber.
    <Additional Fibers> For a given fiber, there is usually a list of additional fiber names following the plant location.
    <Notes> Any additional text that is not covered by the above categories.

    If you do not see one of the above for a given company, do not "make up" a category. Simply skip it.
    If you get an input that has multiple all caps lines, the first one is the company name, and for the following add:
    <Division> the division name, will be in all caps
    If a company does not list any fibers, do not include the fiber category. However, if there is a PLANT location, include that.
    If you encounter an input that is not a company, return it as is.
    As an example, the following company will be parsed into the following categories:
    Input:
        ALBANY INTERNATIONAL - MONOFILAMENT PLANT•
        EXECUTIVE OFFICE: 156 S. Main Street, Homer, NY 13077
        Fluoropolymers Homer, NY Halar, Ky nm; Teflon, Tejzel (monofil)
        Nylon 6, 66, & 612 Homer (monofil)
        Polyester Homer (monofil - regular & hydrolysis-resistant)

    Output:
        <Company Name> ALBANY INTERNATIONAL
        <Executive Office> 156 S. Main Street, Homer, NY 13077
        <Fiber> Fluoropolymers
        <Plant Location> Homer, NY
        <Additional Fibers> Halar, Ky nm; Teflon, Tejzel
        <Fiber> Nylon 6, 66, & 612
        <Plant Location> Homer
        <Fiber> Polyester
        <Plant Location> Homer
        <Additional Fibers> regular & hydrolysis-resistant
    """


# Prompt for the model, this is the instructions for the model to follow. This prompt is for the global data.
my_prompt_world = f"""
    Classify and structure the given text according to the following categories:

    <Country> the country the company is located in
    <Company Name> the name of the company
    <Location> the location of the company
    <Notes> any additional text that is not covered by the above categories

    If you do not see one of the above for a given company, do not "make up" a category. Simply skip it.
    If you get an input that does not have a country, company name, or location, return it as is.
    Separate each company with a newline character. Repeat the categories for each company.
    Correct small typos or misspellings, including OCR errors (e.g., "0" instead of "O").

    As an example:
    Input: 
    A R G E N T I N A
    •Ducilo" S.A. Productora de Ray6n
    Beraza tegui
    Reysol, S.A. at Zarate
    Rhodiaseta Argentina, S.A.
    Quilmes
    A U S T R A L I A
    V S
    V S
    A
    •Caurtaulds Ltd. (Great Britain) is considering
    building a rayon plant at a
    site ,still to be selected. Final decision
    has not yet been made as to the
    process to be used or the products to be
    produced. Probability is that viscose
    yarn and staple will be manufactured.
    A U S T R I A
    Erste (,esterre1chische Glanzstoff-.
    Fabrik, A.G.
    . St. Pol ten (Soviet Zone) V
    Lenzinger Zellwolle und Papierfabrik,
    A.G.
    Lenzing (U. S. Zone) •v s

    Output:
    <Country> ARGENTINA
    <Company Name> Ducilo" S.A. Productora de Ray6n
    <Location> Beraza tegui

    <Country> ARGENTINA
    <Company Name> Reysol, S.A.
    <Location> Zarate

    <Country> ARGENTINA
    <Company Name> Rhodiaseta Argentina, S.A.
    <Location> Quilmes

    <Country> AUSTRALIA
    <Company Name> Caurtaulds Ltd. (Great Britain)
    <Notes> building a rayon plant at a site ,still to be selected. Final decision has not yet been made as to the process to be used or the products to be produced. Probability is that viscose yarn and staple will be manufactured.

    <Country> AUSTRIA
    <Company Name> Erste Oesterreichische Glanzstoff-Fabrik, A.G.
    <Location> St. Polten (Soviet Zone)

    <Country> AUSTRIA
    <Company Name> Lenzinger Zellwolle und Papierfabrik, A.G.
    <Location> Lenzing (U. S. Zone)
"""


# Given a folder of text files, process each file and write the output to an Excel file
def process_input_folder(folder_name, prompt):
    parserAI = AIParser("gpt-4o-mini", prompt)

    # Loop through each file in the folder
    for input_file in os.listdir(folder_name):
        if input_file.endswith(".txt"):
            file_path = os.path.join(folder_name, input_file)

            # skip a file if it has "unfinished" in the name
            if "unfinished" in file_path:
                continue

            # Process the text file
            output = parserAI.process_text(file_path)


# Process all files in the input folder that have not been processed yet
def process_input_folder_unprocessed(folder_name, prompt):
    parserAI = AIParser("gpt-4o-mini", prompt)

    # Loop through each file in the folder, checking if it has a corresponding file in the model_output folder
    for input_file in os.listdir(folder_name):
        if input_file.endswith(".txt"):
            file_path = os.path.join(folder_name, input_file)
            output_file = os.path.join("model_output", input_file.replace(".txt", "_gpt-4o-mini.txt"))

            # skip a file if it has "unfinished" in the name
            if "unfinished" in file_path:
                continue

            # Process the text file if it has not been processed yet
            if not os.path.exists(output_file):
                output = parserAI.process_text(file_path)

    
def write_output_to_excel(input_folder, excel_sheet_name):
    parserExcel = ExcelParser()
    
    # Loop through each file in the folder
    total_iterations = len(os.listdir(input_folder))
    bar = Bar(f"Converting files in {input_folder} to excel: ", max=total_iterations)
    for input_file in os.listdir(input_folder):
        if input_file.endswith(".txt"):
            file_path = os.path.join(input_folder, input_file)

            # Process the text file
            parserExcel.text_to_excel(file_path, excel_sheet_name)
            bar.next()
    bar.finish()

    print(f"Text conversion to excel sheet complete - check {excel_sheet_name} for the results")


input_folder = "text_files_world"
model_output_folder = "model_output_world"
output_excel = "data.xlsx"

# Process the input folder, writing the output to the model_output folder

# --- Two options for processing the input folder: ---

# 1. Process all files in the input folder
# process_input_folder(input_folder)

# 2. Process only files that have not been processed yet
    # No new files, this should not do anything

process_input_folder_unprocessed(input_folder, my_prompt_world)


# Write the output to an Excel sheet
#write_output_to_excel(model_output_folder, output_excel)
            