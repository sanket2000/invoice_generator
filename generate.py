import csv
import os
import shutil
from pprint import pprint as pp #Pretty print.


def read_placeholder_data(path):
    """Reads values to put into invoice."""
    with open(path) as csv_file:
        csv_data = csv.DictReader(csv_file) # itterable returns dict like rows
        data_sheet = [data for data in csv_data] # list of dict with row data
        # pp(data_sheet[0])
        return data_sheet

def read_docfile_template_content(path):
    """Reads content.xml from template odt file."""
    with open(path, "rt") as f:
        template_file_content = f.read()
    return template_file_content

def modify_docfile_content(row, template_file_content):
    """Modifys content.xml from template odt file."""
    for key in row.keys():
        template_file_content = template_file_content.replace(key, row[key])
    return template_file_content

def write_docfile_content(data, path):
    """Writes modified data to content.xml from template odt file."""
    with open(path, "wt") as f:
        f.seek(0)
        f.write(data)

def export_doc(file_dir, file_name, export_dir):
    """Exports unpacked and modified odt folder to valid odt file."""
    try:
        shutil.make_archive(export_dir+file_name, "zip", file_dir)
        os.rename(export_dir+file_name+".zip", export_dir+file_name+".odt")
        print(f"Generated: {file_name}")
    except FileExistsError:
        print(f"File already exixts: {file_name}")

def export_pdf(odt_export_dir, pdf_export_dir):
    """Uses soffice cmd args to export odt to pdf."""
    os.system(f'"C:/Program Files/LibreOffice/program/soffice" \
                --headless \
                --convert-to pdf \
                --outdir {pdf_export_dir} \
                {odt_export_dir}/*.odt')

def make_temp_dir(temp_dir, template_path):
    """Creates temp directory and generates necessary data."""
    ##########
    if not os.path.isdir(temp_dir):
        os.mkdir(temp_dir)
    ##########    
    shutil.copy(template_path, temp_dir)
    ##########
    name = os.path.splitext(os.path.basename(template_path))[0]
    try:
        os.rename(temp_dir+name+".odt", temp_dir+"template"+".zip")
    except FileExistsError:
        pass
    ##########
    shutil.unpack_archive(f"{temp_dir}/template.zip", f"{temp_dir}/template")
    shutil.copy(f"{temp_dir}/template/content.xml", f"{temp_dir}")

def get_imp_paths(data_dir):
    """Returns CSV and ODT file path from data dir."""
    from glob import glob as g
    try:
        datasheet_path = g(f"{data_dir}/*.csv")[0]
        template_path = g(f"{data_dir}/*.odt")[0]
    except IndexError:
        datasheet_path = input("Enter CSV path:")
        template_path = input("Enter template ODT path:")
    
    return datasheet_path, template_path

def generate_doc(data_sheet, template_file_content, content_dir, odt_export_dir):
    """Generates valid ODT files from CSV data and template."""
    for row in data_sheet:
        modified_data = modify_docfile_content(row, template_file_content)
        write_docfile_content(modified_data, content_dir+"/content.xml")
        export_doc(content_dir, row['file_name'], odt_export_dir)

def clean_up_dirs(*dirs):
    """Deletes given directory recursively."""
    for dir in dirs:
        print(f"Deleting: {dir}")
        shutil.rmtree(dir)

def main():

    # read only user input paths
    temp_dir = "temp/"
    export_dir = "export/"
    data_dir = "data/"
    datasheet_path, template_path = get_imp_paths(data_dir)
    
    # for internal use
    odt_export_dir = export_dir+"odt/"
    pdf_export_dir = export_dir
    content_dir = f"{temp_dir}/template/"
    content_path = f"{temp_dir}/content.xml"
    
    # basically find, replace, export
    make_temp_dir(temp_dir, template_path)
    data_sheet = read_placeholder_data(datasheet_path)
    template_file_content = read_docfile_template_content(content_path)
    generate_doc(data_sheet, template_file_content, content_dir, odt_export_dir)
    export_pdf(odt_export_dir, pdf_export_dir)

    # delete unwanted dirs
    clean_up_dirs(temp_dir, odt_export_dir)
            

if __name__ == "__main__":
    print("""
Hello,

Note: This program assumes you have python3.6 or later

With this program you can generate pdf files from .csv data and .odt template file.

Step1: Create a formated .odt file as template.
Step2: Put place-holder values in template file to be replaced.
Step3: Create a .csv file with place-holder vales as headers.
Step4: Put replacement value under header rows.
Step5: Place .csv and .odt file into 'data' folder

    """)
    main()
