import csv
from msilib.schema import File
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

def create_doc(file_dir, file_name, export_dir):
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
                {odt_export_dir}*.odt')

def create_temp_dir(temp_dir, template_path):
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

def get_imp_paths():
    """Returns CSV and ODT file path from data dir."""
    from glob import glob as g
    try:
        datasheet_path = g("data/*.csv")[0]
        template_path = g("data/*.odt")[0]
    except IndexError:
        datasheet_path = input("Enter CSV path:")
        template_path = input("Enter ODT path:")
    
    return datasheet_path, template_path

def generate_doc(data_sheet, template_file_content, content_dir, odt_export_dir):
    for row in data_sheet:
        modified_data = modify_docfile_content(row, template_file_content)
        write_docfile_content(modified_data, content_dir+"/content.xml")
        create_doc(content_dir, row['file_name'], odt_export_dir)

def clean_up_dirs(*dirs):
    """Deletes given directory recursively."""
    for dir in dirs:
        shutil.rmtree(dir)

def main():
    # read only user input paths
    temp_dir = "temp/"
    export_dir = "export/"
    datasheet_path, template_path = get_imp_paths()
    
    # for internal use
    odt_export_dir = export_dir+"odt/"
    pdf_export_dir = export_dir
    content_dir = f"{temp_dir}/template/"
    content_path = f"{temp_dir}/content.xml"
    
    create_temp_dir(temp_dir, template_path)
    data_sheet = read_placeholder_data(datasheet_path)
    template_file_content = read_docfile_template_content(content_path)
    generate_doc(data_sheet, template_file_content, content_dir, odt_export_dir)
    export_pdf(odt_export_dir, pdf_export_dir)

    clean_up_dirs(temp_dir, odt_export_dir)
            

if __name__ == "__main__":
    main()
