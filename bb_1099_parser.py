# bb_1099_parser.py

import csv
import os
import io
import time
import logging

import pdfplumber


logging.debug('bb_1099_parser')

logger = logging.getLogger(__name__)

def remove_extra_spaces(entry):
    return ' '.join(entry.split())

def get_name(page):
    ignored_words = ["Request"]
    #'x0': Decimal('67.367'), 'x1': Decimal('155.323'), 'top': Decimal('276.063'), 'bottom': Decimal('287.063')
    #page = page.crop([67.367, 95.487, 387.710, 106.487])
    page = page.crop([67.367, 50.487, 387.710, 106.487])

    name = page.extract_words(
    x_tolerance=5, 
    y_tolerance=1, 
    keep_blank_chars=True, 
    use_text_flow=False, 
    horizontal_ltr=True, 
    vertical_ttb=False, 
    extra_attrs=["size"])

    full_name = ""
    if len(name) != 0:
        for entry in name:
            if entry["size"] == 11.0:
                full_name += entry["text"] + " "
        full_name = remove_extra_spaces(full_name)
        return full_name.upper()
    else:
        # Log here
        return ""

def get_address(page):
    page = page.crop([67.367, 251.550, 240.000, 262.550])

    address = page.extract_words(
    x_tolerance=5, 
    y_tolerance=1, 
    keep_blank_chars=True, 
    use_text_flow=False, 
    horizontal_ltr=True, 
    vertical_ttb=False, 
    extra_attrs=["size"])

    if len(address) != 0:
        return address[0]['text'].upper()
    else:
        return ""

def get_city_state_zip(page):
    #'x0': Decimal('67.367'), 'x1': Decimal('155.323'), 'top': Decimal('276.063'), 'bottom': Decimal('287.063')
    page = page.crop([67.367, 276.063, 240.000, 287.063])

    address = page.extract_words(
    x_tolerance=5, 
    y_tolerance=1, 
    keep_blank_chars=True, 
    use_text_flow=False, 
    horizontal_ltr=True, 
    vertical_ttb=False, 
    extra_attrs=[])

    if len(address) != 0:
        return address[0]['text'].upper()
    else:
        return ""

def get_tax_id(page):
    # x0, top, x1, bottom
    #page = page.crop([421.167, 334.784, 483.449, 359.784]) # Starting point
    page = page.crop([420.167, 334.784, 580.449, 359.784])
    #'x0': Decimal('421.167'), 'x1': Decimal('483.449'), 'top': Decimal('348.784')332.542, 'bottom': Decimal('359.784')

    tax_id = page.extract_words(
    x_tolerance=1,  # was 5
    y_tolerance=1, 
    keep_blank_chars=True, 
    use_text_flow=False, 
    horizontal_ltr=True, 
    vertical_ttb=False, 
    extra_attrs=[])
    
    #print(tax_id)
    final_id = ""
    for id in tax_id:
        new_id = ""
        new_id = ''.join(filter(str.isdigit, id['text']))
        if len(new_id) != 9:
            continue
        else:
            final_id = new_id[0:3] + "-" + new_id[3:5] + "-" + new_id[5:9]
            break

    return final_id

def log_errors(output):
    final_error = ""
    missing_keys = ""
    if len(output["name"].split(" ")) < 2:
        final_error += "[Name Too Short] "

    for key in output:
        if output[key] == "":
            missing_keys += key + " "

    if len(missing_keys) != 0:
        final_error += "[Missing: ( " + missing_keys + " )]" 
    
    return final_error

def get_pdf_text(pdf_path):
    pdf = pdfplumber.open(pdf_path)
    page = pdf.pages[0]
    output = {}

    address = ""
    city_state_zip = ""
    output["fileName"] = pdf_path.split("\\")[-1]
    output["name"] = get_name(page)
    output["taxId"] = get_tax_id(page)
    address = get_address(page)
    city_state_zip = get_city_state_zip(page)
    output["fullAddress"] = address + " " + city_state_zip

    pdf.close()
    output["notes"] = log_errors(output)
    return output

def export_as_csv(output, csv_path):
    fields = output[list(output.keys())[0]].keys()
    #print(fields)
    with open(csv_path, 'w', newline='') as csvfile:
        writer = csv.DictWriter(csvfile, fieldnames=fields)
        writer.writeheader()
        for key in output:
            writer.writerow(output[key])

def parse_w9s():
    csv_path = "output\\" + "w9_output.csv"
    output = {}
    for subdir, dirs, files in os.walk('pdfs'):
        for file in files:
            filepath = subdir + os.sep + file
            if filepath.endswith(".pdf"):
                file_name = filepath.split("\\")[-1]
                if "audit_trail" not in file_name:
                    output[file_name] = get_pdf_text(filepath)
                    #print(output[file_name])
    export_as_csv(output, csv_path)

def remove_audit_trail_files():
    for subdir, dirs, files in os.walk('pdfs'):
        for file in files:
            filepath = subdir + os.sep + file
            if filepath.endswith(".pdf"):
                file_name = filepath.split("\\")[-1]
                if "audit_trail" in file_name:
                    os.remove(filepath)

def find_duplicate_tax_ids(file_path):
    tax_ids = []
    duplicates = []
    with open(file_path) as csvfile:
        reader = csv.DictReader(csvfile)
        for row in reader:
            if row["taxId"] != "":
                if row["taxId"] in tax_ids:
                    duplicates.append(row["taxId"])
                tax_ids.append(row["taxId"])
    return duplicates

if __name__ == '__main__':
    parse_w9s()
    #remove_audit_trail_files()
    #print(len(find_duplicate_tax_ids("output/w9_output_second_pass.csv")))