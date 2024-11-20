import pandas as pd

# Methode zum Einlesen der CSV-Datei
def load_employee_absence_data(file_path):
    """
    Lädt die Abwesenheitsdaten aus einer CSV-Datei und gibt sie als DataFrame zurück.
    
    :param file_path: Der Pfad zur CSV-Datei
    :return: Pandas DataFrame mit den geladenen Abwesenheitsdaten
    """
    try:
        # CSV-Datei einlesen
        data = pd.read_csv(file_path)
        print("Daten erfolgreich geladen.")
        return data
    except FileNotFoundError:
        print(f"Die Datei {file_path} wurde nicht gefunden.")
        return None
    except Exception as e:
        print(f"Ein Fehler ist aufgetreten: {e}")
        return None

# Hauptmethode
def main():
    # Pfad zur CSV-Datei
    file_path = 'employee_absence.csv'
    
    # Daten einlesen
    absence_data = load_employee_absence_data(file_path)
    
    # Überprüfen, ob die Daten erfolgreich geladen wurden
    if absence_data is not None:
        # Hier kannst du mit den geladenen Daten weiterarbeiten
        print(absence_data.head())  # Ausgabe der ersten 5 Zeilen der geladenen Daten
    
if __name__ == "__main__":
    main()
