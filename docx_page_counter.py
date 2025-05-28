import os
import comtypes.client


def count_pages_word(word_path):
    """Zlicza liczbę stron w pliku .docx i .doc za pomocą Word COM API."""
    try:
        word = comtypes.client.CreateObject("Word.Application")
        word.Visible = False  # Ukrywamy Worda
        doc = word.Documents.Open(word_path)
        page_count = doc.ComputeStatistics(2)  # Wartość 2 oznacza liczbę stron
        doc.Close(False)
        word.Quit()
        return page_count
    except Exception as e:
        print(f"Błąd podczas przetwarzania pliku {word_path}: {e}")
        return 0


def process_docs_in_folder(folder_path):
    """Rekurencyjnie przeszukuje folder i podfoldery, zliczając liczbę stron w plikach .doc i .docx."""
    total_pages = 0
    file_count = 0

    # Przechodzenie przez foldery i podfoldery
    for root, _, files in os.walk(folder_path):
        for file in files:
            if file.lower().endswith(('.doc', '.docx')):
                file_path = os.path.join(root, file)
                page_count = count_pages_word(file_path)

                print(f"{file_path}: {page_count} stron")
                total_pages += page_count
                file_count += 1

    if file_count == 0:
        print("Brak plików .doc lub .docx w podanej lokalizacji.")

    return total_pages


if __name__ == "__main__":
    folder_path = input(
        "Podaj ścieżkę do folderu z plikami .doc i .docx: ").strip()

    if not os.path.exists(folder_path):
        print("Podana ścieżka nie istnieje. Sprawdź poprawność.")
    else:
        total = process_docs_in_folder(folder_path)
        print(f"\nŁączna liczba stron we wszystkich plikach: {total}")
