import os
from PyPDF2 import PdfReader
from PyPDF2.generic import IndirectObject

def list_annotation_types(pdf_path):
    try:
        reader = PdfReader(pdf_path)
        all_types = {}

        for i, page in enumerate(reader.pages):
            annots = page.get("/Annots")
            if annots:
                if isinstance(annots, IndirectObject):
                    annots = annots.get_object()
                for annot_ref in annots:
                    annot = annot_ref.get_object()
                    subtype = annot.get("/Subtype", "BRAK")
                    contents = annot.get("/Contents", "")
                    all_types.setdefault(subtype, 0)
                    all_types[subtype] += 1
                    print(f"Strona {i+1} | Typ: {subtype} | Tekst: {str(contents)[:80]}")

        print("\n📋 Zliczone typy adnotacji:")
        for t, c in all_types.items():
            print(f"• {t}: {c} szt.")

    except Exception as e:
        print(f"❌ Błąd: {e}")

if __name__ == "__main__":
    pdf_path = input("📄 Podaj ścieżkę do pliku PDF: ").strip('"')
    if os.path.isfile(pdf_path) and pdf_path.lower().endswith(".pdf"):
        list_annotation_types(pdf_path)
    else:
        print("❌ Podany plik nie istnieje lub nie jest plikiem PDF.")
