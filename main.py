import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import os
import sys

def load_employee_absence_data(file_path):
    """
    Lädt die Abwesenheitsdaten aus einer CSV-Datei und gibt sie als DataFrame zurück.
    
    :param file_path: Der Pfad zur CSV-Datei
    :return: Pandas DataFrame mit den geladenen Abwesenheitsdaten
    """
    try:
        data = pd.read_csv(file_path)
        print(f"Daten erfolgreich aus '{file_path}' geladen.")
        return data
    except FileNotFoundError:
        print(f"Die Datei '{file_path}' wurde nicht gefunden.")
        sys.exit(1)
    except pd.errors.EmptyDataError:
        print(f"Die Datei '{file_path}' ist leer.")
        sys.exit(1)
    except pd.errors.ParserError:
        print(f"Fehler beim Parsen der Datei '{file_path}'. Bitte überprüfen Sie das Dateiformat.")
        sys.exit(1)
    except Exception as e:
        print(f"Ein unerwarteter Fehler ist aufgetreten: {e}")
        sys.exit(1)

def calculate_absence_duration(df):
    """
    Berechnet die Dauer der Abwesenheit in Tagen und fügt sie als neue Spalte hinzu.
    
    :param df: Pandas DataFrame mit den Abwesenheitsdaten
    :return: DataFrame mit der neuen Spalte 'absence_duration'
    """
    try:
        df['absence_start'] = pd.to_datetime(df['absence_start'])
        df['absence_end'] = pd.to_datetime(df['absence_end'])
        df['absence_duration'] = (df['absence_end'] - df['absence_start']).dt.days + 1
        print("Abwesenheitsdauer erfolgreich berechnet.")
        return df
    except KeyError as e:
        print(f"Die erwartete Spalte fehlt in den Daten: {e}")
        sys.exit(1)
    except Exception as e:
        print(f"Fehler bei der Berechnung der Abwesenheitsdauer: {e}")
        sys.exit(1)

def sanitize_filename(name):
    """
    Sanitizes den Dateinamen, indem unerwünschte Zeichen entfernt oder ersetzt werden.
    
    :param name: Ursprünglicher Name
    :return: Sanitierter Name
    """
    valid_chars = "-_.() abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
    sanitized = ''.join(c if c in valid_chars else '_' for c in name)
    return sanitized.replace(' ', '_')

def plot_absence_distribution(df, absence_reason, output_dir):
    """
    Erstellt ein Histogramm zur Verteilung der Abwesenheitsdauer für einen bestimmten Abwesenheitsgrund und speichert es.
    
    :param df: Gefilterter Pandas DataFrame
    :param absence_reason: Der Abwesenheitsgrund
    :param output_dir: Verzeichnis zum Speichern der Bilder
    """
    try:
        plt.figure(figsize=(10, 6))
        sns.histplot(data=df, x='absence_duration', bins=range(1, df['absence_duration'].max()+2), kde=False, color='skyblue')
        plt.title(f"Verteilung der Abwesenheitsdauer für '{absence_reason}'")
        plt.xlabel('Abwesenheitsdauer (Tage)')
        plt.ylabel('Anzahl der Abwesenheiten')
        plt.xticks(range(1, df['absence_duration'].max()+1))
        plt.grid(axis='y', linestyle='--', alpha=0.7)
        plt.tight_layout()
        
        # Dateinamen sanitieren
        sanitized_reason = sanitize_filename(absence_reason)
        histogram_path = os.path.join(output_dir, f"histogram_{sanitized_reason}.png")
        plt.savefig(histogram_path)
        plt.close()
        print(f"Histogramm für '{absence_reason}' gespeichert als '{histogram_path}'.")
    except Exception as e:
        print(f"Fehler beim Erstellen des Histogramms für '{absence_reason}': {e}")

def plot_absence_boxplot(df, absence_reason, output_dir):
    """
    Erstellt einen Boxplot zur Verteilung der Abwesenheitsdauer für einen bestimmten Abwesenheitsgrund und speichert ihn.
    
    :param df: Gefilterter Pandas DataFrame
    :param absence_reason: Der Abwesenheitsgrund
    :param output_dir: Verzeichnis zum Speichern der Bilder
    """
    try:
        plt.figure(figsize=(10, 6))
        sns.boxplot(data=df, x='absence_reason', y='absence_duration', palette='Set2')
        plt.title(f"Boxplot der Abwesenheitsdauer für '{absence_reason}'")
        plt.xlabel('Abwesenheitsgrund')
        plt.ylabel('Abwesenheitsdauer (Tage)')
        plt.grid(axis='y', linestyle='--', alpha=0.7)
        plt.tight_layout()
        
        # Dateinamen sanitieren
        sanitized_reason = sanitize_filename(absence_reason)
        boxplot_path = os.path.join(output_dir, f"boxplot_{sanitized_reason}.png")
        plt.savefig(boxplot_path)
        plt.close()
        print(f"Boxplot für '{absence_reason}' gespeichert als '{boxplot_path}'.")
    except Exception as e:
        print(f"Fehler beim Erstellen des Boxplots für '{absence_reason}': {e}")

def main():
    # Pfad zur CSV-Datei
    csv_file = 'employee_absence.csv'  # Passen Sie den Pfad nach Bedarf an
    
    # Verzeichnis zum Speichern der Diagramme
    output_dir = 'plots'
    os.makedirs(output_dir, exist_ok=True)
    
    # Laden der Daten
    data = load_employee_absence_data(csv_file)
    
    # Berechnung der Abwesenheitsdauer
    data = calculate_absence_duration(data)
    
    # Ermittlung der einzigartigen Abwesenheitsgründe
    absence_reasons = data['absence_reason'].unique()
    
    # Iteration über jeden Abwesenheitsgrund
    for reason in absence_reasons:
        # Filtern der Daten für den aktuellen Abwesenheitsgrund
        filtered_data = data[data['absence_reason'] == reason]
        
        if filtered_data.empty:
            print(f"Keine Daten für den Abwesenheitsgrund '{reason}' gefunden.")
            continue
        
        # Erstellung und Speicherung der Diagramme
        plot_absence_distribution(filtered_data, reason, output_dir)
        plot_absence_boxplot(filtered_data, reason, output_dir)
    
    print("Alle Diagramme wurden erfolgreich erstellt und gespeichert.")

if __name__ == "__main__":
    main()
