import os
import zipfile
import shutil
import tempfile
from lxml import etree

def remove_on_behalf_of_anywhere(docx_path):
    with zipfile.ZipFile(docx_path, 'r') as docx_zip:
        temp_dir = tempfile.mkdtemp()
        docx_zip.extractall(temp_dir)

    found = False
    for root_dir, _, files in os.walk(temp_dir):
        for file in files:
            if file.endswith('.xml'):
                xml_path = os.path.join(root_dir, file)
                try:
                    tree = etree.parse(xml_path)
                    modified = False

                    for elem in tree.xpath(f"//*[contains(text(), '{property}')]"):
                        print(f"Found in: {xml_path} — tag: {elem.tag}")
                    
                    for elem in tree.iter():
                        if elem.text and property in elem.tag:
                            print(f"Clearing tag text: {elem.tag}")
                            elem.text = ''
                            modified = True
                            found = True
                        elif elem.get('name') == property:
                            print(f"Clearing named element value in: {xml_path}")
                            if len(elem):
                                elem[0].text = ''
                                modified = True
                                found = True

                    if modified:
                        tree.write(xml_path, xml_declaration=True, encoding='UTF-8', standalone=True)

                except Exception as e:
                    # Not all XML files may be valid XML (some binary content)
                    continue

    if found:
        # Repackage the DOCX file
        temp_docx_path = docx_path + '.tmp'
        with zipfile.ZipFile(temp_docx_path, 'w', zipfile.ZIP_DEFLATED) as new_zip:
            for foldername, subfolders, filenames in os.walk(temp_dir):
                for filename in filenames:
                    file_path = os.path.join(foldername, filename)
                    arcname = os.path.relpath(file_path, temp_dir)
                    new_zip.write(file_path, arcname)

        shutil.move(temp_docx_path, docx_path)

    shutil.rmtree(temp_dir)

def batch_process_folder(folder_path):
    for filename in os.listdir(folder_path):
        if filename.endswith('.docx'):
            full_path = os.path.join(folder_path, filename)
            print(f"\nProcessing: {filename}")
            remove_on_behalf_of_anywhere(full_path)

# Update this path
folder_with_docs = r'./'
property = 'OnBehalfOf'
batch_process_folder(folder_with_docs)
