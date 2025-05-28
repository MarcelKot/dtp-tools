import os
from pypdf import PdfReader

def count_pages(pdf_path):
    """Zlicza liczbę stron w pliku PDF i rozróżnia rozkładówki."""
    try:
        reader = PdfReader(pdf_path)
        single_pages = 0
        spreads = 0

        for page in reader.pages:
            width = page.mediabox.width
            height = page.mediabox.height

            # Rozkładówka → jeśli szerokość jest 2 razy większa niż wysokość
            if width / height > 1.5:  
                spreads += 1
            else:
                single_pages += 1

        total_pages = single_pages + (spreads * 2)  # Liczymy rozkładówki jako 2 strony
        return single_pages, spreads, total_pages

    except Exception as e:
        print(f"Błąd podczas przetwarzania pliku {pdf_path}: {e}")
        return 0, 0, 0

def process_pdfs_in_folder(folder_path):
    """Zlicza liczbę stron i rozkładówek we wszystkich plikach .pdf w folderze."""
    total_single_pages = 0
    total_spreads = 0
    total_pages = 0
    pdf_files = [f for f in os.listdir(folder_path) if f.lower().endswith('.pdf')]

    if not pdf_files:
        print("Brak plików .pdf w podanym folderze.")
        return 0, 0, 0

    for pdf_file in pdf_files:
        file_path = os.path.join(folder_path, pdf_file)
        single_pages, spreads, pages = count_pages(file_path)
        print(f"{pdf_file}: {single_pages} pojedynczych stron, {spreads} rozkładówek (→ {pages} stron rzeczywistych)")

        total_single_pages += single_pages
        total_spreads += spreads
        total_pages += pages

    return total_single_pages, total_spreads, total_pages

if __name__ == "__main__":
    folder_path = input("Podaj ścieżkę do folderu z plikami .pdf: ").strip()

    if not os.path.exists(folder_path):
        print("Podana ścieżka nie istnieje. Sprawdź poprawność.")
    else:
        single, spreads, total = process_pdfs_in_folder(folder_path)
        print(f"\n📊 Podsumowanie dla folderu:")
        print(f"🔹 Pojedyncze strony: {single}")
        print(f"🔹 Rozkładówki: {spreads} (liczone jako {spreads * 2} stron)")
        print(f"📌 Łączna liczba stron (uwzględniając rozkładówki): {total}")
