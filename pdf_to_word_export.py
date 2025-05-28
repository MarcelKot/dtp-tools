import os
import glob
from adobe.pdfservices.client.auth.credentials import Credentials
from adobe.pdfservices.client.execution_context import ExecutionContext
from adobe.pdfservices.client.pdfops.export_pdf_operation import ExportPDFOperation
from adobe.pdfservices.client.pdfops.options.export_pdf_options import ExportPDFOptions, ExportPDFTargetFormat
from adobe.pdfservices.client.io.file_ref import FileRef
from adobe.pdfservices.client.errors import ServiceApiException, ServiceUsageException, SdkException

def convert_pdfs_to_docx(folder_path):
    # Utwórz poświadczenia z pliku JSON
    credentials = Credentials.service_account_credentials_builder() \
        .from_file("pdfservices-api-credentials.json") \
        .build()

    execution_context = ExecutionContext.create(credentials)

    # Znajdź wszystkie pliki PDF w folderze
    pdf_files = glob.glob(os.path.join(folder_path, "*.pdf"))

    if not pdf_files:
        print("⚠️ Nie znaleziono żadnych plików PDF.")
        return

    for pdf_file in pdf_files:
        try:
            print(f"🔄 Przetwarzanie: {os.path.basename(pdf_file)}")

            # Przygotuj operację eksportu do DOCX
            operation = ExportPDFOperation.builder() \
                .with_input(FileRef.create_from_local_file(pdf_file)) \
                .with_options(ExportPDFOptions.builder()
                              .with_target_format(ExportPDFTargetFormat.DOCX)
                              .build()) \
                .build()

            # Wykonaj konwersję
            result = operation.execute(execution_context)

            # Zapisz wynik
            output_path = os.path.splitext(pdf_file)[0] + ".docx"
            result.save_as(output_path)

            print(f"✅ Zapisano: {output_path}")

        except (ServiceApiException, ServiceUsageException, SdkException) as e:
            print(f"❌ Błąd przy {os.path.basename(pdf_file)}: {e}")

if __name__ == "__main__":
    folder = input("📂 Podaj ścieżkę do folderu z PDF-ami: ").strip()
    if os.path.isdir(folder):
        convert_pdfs_to_docx(folder)
    else:
        print("❌ Niepoprawna ścieżka folderu.")
