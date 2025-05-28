import os
import docx
import pytesseract
from PIL import Image
from io import BytesIO

# Ścieżka do Tesseract OCR (zmień, jeśli masz inną)
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"


def extract_images_from_docx(docx_path):
    """Ekstrahuje obrazy z pliku Word."""
    try:
        doc = docx.Document(docx_path)
        images = []
        for rel in doc.part.rels:
            if "image" in doc.part.rels[rel].target_ref:
                image_data = doc.part.rels[rel].target_part.blob
                img = Image.open(BytesIO(image_data))
                images.append(img)
        return images
    except Exception as e:
        print(f"Błąd w {docx_path}: {e}")
        return []


def count_images_with_text(docx_path):
    """Zlicza obrazy zawierające tekst."""
    images = extract_images_from_docx(docx_path)
    text_images_count = 0

    for img in images:
        extracted_text = pytesseract.image_to_string(
            img, lang="eng")  # OCR na obrazach
        if extracted_text.strip():  # Sprawdza, czy na obrazie jest jakikolwiek tekst
            text_images_count += 1

    return text_images_count


def process_word_files(folder_path):
    """Przeszukuje folder i sprawdza pliki Word pod kątem nieedytowalnych grafik z tekstem."""
    word_files = [f for f in os.listdir(
        folder_path) if f.lower().endswith(('.docx', '.doc'))]

    if not word_files:
        print("Brak plików Word w podanym folderze.")
        return

    for word_file in word_files:
        file_path = os.path.join(folder_path, word_file)
        text_image_count = count_images_with_text(file_path)

        if text_image_count > 0:
            print(
                f"{word_file}: {text_image_count} grafik zawiera tekst (nieedytowalny).")


if __name__ == "__main__":
    folder_path = input("Podaj ścieżkę do folderu z plikami Word: ").strip()

    if not os.path.exists(folder_path):
        print("Podana ścieżka nie istnieje. Sprawdź poprawność.")
    else:
        process_word_files(folder_path)
