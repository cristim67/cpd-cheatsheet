import os
import re

from docx import Document

FILE_PAIRS = [
    ('valori/n1.txt', 'sample/n1.docx', '/tmp/sample_rez_n1.docx'),
    ('valori/n2.txt', 'sample/n2.docx', '/tmp/sample_rez_n2.docx'),
    ('valori/n3.txt', 'sample/n3.docx', '/tmp/sample_rez_n3.docx')
]

def proceseaza_fisier_text(nume_fisier):
    chei_valori = {}
    with open(nume_fisier, encoding='utf-8') as f:
        for line in f:
            match = re.match(r'([A-ZpH][\w\[\]|]+)\s*=\s*([0-9.eE\-]*)', line.strip())
            if match:
                chei_valori[match.group(1)] = match.group(2)
    return chei_valori

def inlocuire(match, chei_valori):
    cheie = match.group(1)
    spatii1 = match.group(2)
    spatii2 = match.group(3)
    valoare_veche = match.group(4)
    if cheie in chei_valori:
        valoare_noua = chei_valori[cheie]
        if valoare_veche.strip() != valoare_noua:
            print(f"Inlocuire: {cheie} = {valoare_veche.strip()} -> {cheie} = {valoare_noua}")
        return f"{cheie}{spatii1}={spatii2}{valoare_noua}"
    return match.group(0)

for txt_file, docx_file, output_file in FILE_PAIRS:
    print(f"Procesez perechea: {txt_file} -> {docx_file}")
    
    chei_valori = proceseaza_fisier_text(txt_file)
    print(f"Procesat {txt_file}: {len(chei_valori)} chei")
    
    doc = Document(docx_file)
    
    for para in doc.paragraphs:
        text_init = para.text
        text_nou = re.sub(r'([A-ZpH][\w\[\]|]+)(\s*)=(\s*)([^\n\r]*)', 
                         lambda m: inlocuire(m, chei_valori), text_init)
        if text_nou != text_init:
            for run in para.runs:
                run.text = text_nou
    
    doc.save(output_file)
    print(f"Fișierul rezultat a fost salvat ca {output_file}")

from merge_docx import (
    merge_docx,  # pip install git+https://github.com/ryanpierson/merge_docx.git
)

merge_docx("/tmp/sample_rez_n1.docx", "/tmp/sample_rez_n2.docx", "/tmp/rez.docx")
merge_docx("/tmp/rez.docx", "/tmp/sample_rez_n3.docx", "rez_final.docx")

for _, _, temp_file in FILE_PAIRS:
    if os.path.exists(temp_file):
        os.remove(temp_file)
if os.path.exists("/tmp/rez.docx"):
    os.remove("/tmp/rez.docx")
