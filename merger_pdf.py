import os
from PyPDF2 import PdfMerger

folder_hasil = "Hasil Merger"

if not os.path.exists(folder_hasil):
    os.makedirs(folder_hasil)

semua_file = os.listdir()
daftar_base = []

for file in semua_file:
    if file.endswith(" INV.pdf"):
        base_name = file.replace(" INV.pdf", "")
        if base_name not in daftar_base:
            daftar_base.append(base_name)
    elif file.endswith(" APV.pdf"):
        base_name = file.replace(" APV.pdf", "")
        if base_name not in daftar_base:
            daftar_base.append(base_name)

for base in daftar_base:
    file_inv = base + " INV.pdf"
    file_apv = base + " APV.pdf"

    if file_inv in semua_file and file_apv in semua_file:
        merger = PdfMerger()
        merger.append(file_inv)
        merger.append(file_apv)
        
        output_path = os.path.join(folder_hasil, base + ".pdf")
        merger.write(output_path)
        merger.close()
        print("--> Berhasil menggabungkan: " + base + ".pdf")
    else:
        if file_inv not in semua_file:
            print("--> File pasangan tidak ditemukan: " + file_inv)
        if file_apv not in semua_file:
            print("--> File pasangan tidak ditemukan: " + file_apv)

input("--> Tekan enter untuk menutup")
