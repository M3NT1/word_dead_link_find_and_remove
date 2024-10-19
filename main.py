import os
import zipfile
import tkinter as tk
from tkinter import filedialog
import logging
from lxml import etree
from openpyxl import Workbook
from datetime import datetime

def choose_file():
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(
        title="Válassza ki a DOCX fájlt",
        filetypes=[("Word files", "*.docx")]
    )
    root.destroy()
    return file_path

def choose_directory():
    root = tk.Tk()
    root.withdraw()
    directory_path = filedialog.askdirectory(
        title="Válassza ki a mentési könyvtárat"
    )
    root.destroy()
    return directory_path

def extract_docx(docx_path, extract_to):
    with zipfile.ZipFile(docx_path, 'r') as zip_ref:
        zip_ref.extractall(extract_to)

def save_docx(extract_from, save_to):
    with zipfile.ZipFile(save_to, 'w') as docx:
        for foldername, subfolders, filenames in os.walk(extract_from):
            for filename in filenames:
                filepath = os.path.join(foldername, filename)
                arcname = os.path.relpath(filepath, extract_from)
                docx.write(filepath, arcname)

def find_and_remove_ghost_links(doc_path):
    with open(doc_path, 'rb') as file:
        tree = etree.parse(file)

    namespaces = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

    bookmarks = set()
    for bm in tree.findall('.//w:bookmarkStart', namespaces):
        name = bm.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}name')
        bookmarks.add(name)

    ghost_links = []
    for hl in tree.findall('.//w:hyperlink', namespaces):
        anchor = hl.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}anchor')
        if anchor not in bookmarks:
            ghost_links.append(hl)
            for elem in hl:
                hl.addprevious(elem)
                # Remove hyperlink formatting
                rPr = elem.find('.//w:rPr', namespaces)
                if rPr is not None:
                    color = rPr.find('.//w:color', namespaces)
                    if color is not None and color.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val') == '0000FF':
                        rPr.remove(color)
                    underline = rPr.find('.//w:u', namespaces)
                    if underline is not None:
                        rPr.remove(underline)
            hl.getparent().remove(hl)

    orphan_bookmarks = []
    for bm in tree.findall('.//w:bookmarkStart', namespaces):
        name = bm.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}name')
        if name not in [hl.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}anchor') for hl in tree.findall('.//w:hyperlink', namespaces)]:
            orphan_bookmarks.append(bm)
            bm_end = tree.find(f'.//w:bookmarkEnd[@w:id="{bm.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id")}"]', namespaces)
            if bm_end is not None:
                bm_end.getparent().remove(bm_end)
            bm.getparent().remove(bm)

    tree.write(doc_path, xml_declaration=True, encoding='UTF-8')

    return ghost_links, orphan_bookmarks, namespaces

def main():
    logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')

    docx_path = choose_file()
    if not docx_path:
        logging.error("Nem választott ki DOCX fájlt.")
        print("Nem választott ki DOCX fájlt. A folyamat leáll.")
        return

    save_directory = choose_directory()
    if not save_directory:
        logging.error("Nem választott ki mentési könyvtárat.")
        print("Nem választott ki mentési könyvtárat. A folyamat leáll.")
        return

    extract_to = os.path.join(save_directory, "extracted")
    os.makedirs(extract_to, exist_ok=True)

    extract_docx(docx_path, extract_to)
    document_xml_path = os.path.join(extract_to, "word", "document.xml")

    ghost_links, orphan_bookmarks, namespaces = find_and_remove_ghost_links(document_xml_path)

    log_filename = os.path.join(save_directory, "process_log.txt")
    logging.basicConfig(filename=log_filename, level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Hivatkozások"
    sheet.append(["Hivatkozás/könyvjelző szövege", "Tétel típusa", "Cél"])

    for hl in ghost_links:
        anchor = hl.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}anchor')
        text_elem = hl.find('.//w:t', namespaces)
        text = text_elem.text if text_elem is not None else "N/A"
        logging.info("Szellemhivatkozás eltávolítva: %s", anchor)
        sheet.append([text, "Szellemhivatkozás", anchor])

    for bm in orphan_bookmarks:
        name = bm.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}name')
        parent = bm.getparent()
        grandparent = parent.getparent() if parent is not None else None
        text_elem = grandparent.find('.//w:t', namespaces) if grandparent is not None else None
        text = text_elem.text if text_elem is not None else "N/A"
        logging.info("Árva könyvjelző eltávolítva: %s", name)
        sheet.append([text, "Árva könyvjelző", name])

    excel_filename = os.path.join(save_directory, "hivatkozasok.xlsx")
    workbook.save(excel_filename)

    # Save the modified document with a new name
    base_name = os.path.splitext(os.path.basename(docx_path))[0]
    date_postfix = datetime.now().strftime("%Y%m%d")
    modified_docx_path = os.path.join(save_directory, f"{base_name}_MOD_{date_postfix}.docx")
    save_docx(extract_to, modified_docx_path)

    print(f"Az eredmények elmentve a következő helyre: {save_directory}")

if __name__ == "__main__":
    main()
