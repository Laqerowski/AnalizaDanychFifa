import json
import re
import pandas as pd
import os

# Sekcja 1: Oczyszczanie i przygotowanie tekstu
def clean_and_format_text(text):
    """
    Oczyszcza tekst, usuwając numery indeksów, dodając cudzysłowy do kluczy i formatując go do postaci JSON.
    """
    # Usuwanie numerów indeksów (np. 0:, 1:, 2:) i usuwanie zbędnych spacji
    cleaned_text = re.sub(r"\d+\s*:\s*", "", text)  # Usuwamy numery indeksów na początku każdej linii

    # Dodanie odpowiednich cudzysłowów do kluczy
    cleaned_text = cleaned_text.replace("x:", '"x":').replace("y:", '"y":').replace("shop:", '"shop":').replace("name:", '"name":')

    # Dodanie przecinków między obiektami i nawiasów klamrowych na początku i końcu
    cleaned_text = re.sub(r'(?<=\})(?=\s*{)', ',', cleaned_text)  # Dodanie przecinków po każdym obiekcie
    cleaned_text = f"[{cleaned_text}]"

    # Logowanie oczyszczonego tekstu (tylko początkowa część)
    print(f"Cleaned Text:\n{cleaned_text[:500]}...")  # Logowanie tylko pierwszych 500 znaków

    return cleaned_text

# Sekcja 2: Konwersja oczyszczonego tekstu na JSON
def convert_text_to_json(cleaned_text):
    """
    Konwertuje oczyszczony tekst na dane w formacie JSON.
    """
    try:
        # Konwersja tekstu na dane JSON
        data = json.loads(cleaned_text)
        return data
    except json.JSONDecodeError as e:
        print(f"Error decoding JSON: {e}")
        print(f"Problematic Text:\n{cleaned_text[:500]}...")  # Logowanie tylko pierwszych 500 znaków
        return None

# Sekcja 3: Zapis do pliku Excel
def save_json_to_excel(json_data, excel_path):
    """
    Zapisuje dane JSON do pliku Excel.
    """
    # Tworzenie DataFrame z danych JSON
    df = pd.DataFrame(json_data)
    
    # Zapisanie do pliku Excel
    df.to_excel(excel_path, index=False)
    print(f"Plik Excel zapisany jako: {excel_path}")

# Funkcja do przetwarzania pliku i zapisania go do Excela
def process_file(file_path, excel_folder, excel_filename=None):
    """
    Oczyszcza, konwertuje na JSON i zapisuje dane do pliku Excel.
    """
    try:
        with open(file_path, 'r') as f:
            raw_data = f.read()

        # Sekcja 1: Oczyszczanie i formatowanie tekstu
        cleaned_text = clean_and_format_text(raw_data)

        # Sekcja 2: Konwersja na JSON
        json_data = convert_text_to_json(cleaned_text)

        if json_data is not None:
            print("Dane zostały pomyślnie przetworzone na JSON!")

            # Ustalamy nazwę pliku Excel, jeśli nie została podana
            if not excel_filename:
                # Domyślna nazwa pliku (możesz ją dostosować)
                excel_filename = 'plik.xlsx'

            # Sekcja 3: Zapisanie danych do pliku Excel w folderze 'dane'
            excel_path = os.path.join(excel_folder, excel_filename)  # Tworzenie pełnej ścieżki do pliku

            # Zapisanie danych do pliku Excel
            save_json_to_excel(json_data, excel_path)
        else:
            print("Błąd podczas przetwarzania danych.")
    
    except FileNotFoundError:
        print(f"Błąd: Plik '{file_path}' nie został znaleziony.")
