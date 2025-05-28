import os
from pptx import Presentation


def count_visible_slides(pptx_path):
    """Zlicza liczbę nieukrytych slajdów w prezentacji PowerPoint."""
    try:
        presentation = Presentation(pptx_path)
        visible_slides = sum(
            1 for slide in presentation.slides if not slide._element.get("show") == "0")
        return visible_slides
    except Exception as e:
        print(f"Błąd podczas przetwarzania pliku {pptx_path}: {e}")
        return 0


def process_pptx_in_folder(folder_path):
    """Zlicza liczbę nieukrytych slajdów we wszystkich plikach .pptx w folderze."""
    total_slides = 0
    pptx_files = [f for f in os.listdir(
        folder_path) if f.lower().endswith('.pptx')]

    if not pptx_files:
        print("Brak plików .pptx w podanym folderze.")
        return 0

    for pptx_file in pptx_files:
        file_path = os.path.join(folder_path, pptx_file)
        slide_count = count_visible_slides(file_path)
        print(f"{pptx_file}: {slide_count} slajdów")
        total_slides += slide_count

    return total_slides


if __name__ == "__main__":
    folder_path = input("Podaj ścieżkę do folderu z plikami .pptx: ").strip()

    if not os.path.exists(folder_path):
        print("Podana ścieżka nie istnieje. Sprawdź poprawność.")
    else:
        total = process_pptx_in_folder(folder_path)
        print(
            f"\nŁączna liczba nieukrytych slajdów we wszystkich plikach: {total}")
