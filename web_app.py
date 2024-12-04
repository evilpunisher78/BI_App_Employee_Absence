import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import streamlit as st

# Methode zum Einlesen der CSV-Datei
@st.cache_data
def load_employee_absence_data(file_path):
    """
    Lädt die Abwesenheitsdaten aus einer CSV-Datei und gibt sie als DataFrame zurück.
    
    :param file_path: Der Pfad zur CSV-Datei
    :return: Pandas DataFrame mit den geladenen Abwesenheitsdaten
    """
    try:
        # CSV-Datei einlesen
        data = pd.read_csv(file_path)
        return data
    except FileNotFoundError:
        st.error(f"Die Datei {file_path} wurde nicht gefunden.")
        return None
    except Exception as e:
        st.error(f"Ein Fehler ist aufgetreten: {e}")
        return None

# Methode zur Berechnung der Abwesenheitsdauer
def calculate_absence_duration(df):
    """
    Berechnet die Dauer der Abwesenheit in Tagen und fügt sie als neue Spalte hinzu.
    
    :param df: Pandas DataFrame mit den Abwesenheitsdaten
    :return: DataFrame mit der neuen Spalte 'absence_duration'
    """
    df['absence_start'] = pd.to_datetime(df['absence_start'])
    df['absence_end'] = pd.to_datetime(df['absence_end'])
    df['absence_duration'] = (df['absence_end'] - df['absence_start']).dt.days + 1
    return df

# Methode zur Erstellung des Histogramms
def plot_absence_distribution(df):
    """
    Erstellt ein Histogramm zur Verteilung der Abwesenheitsdauer.
    
    :param df: Pandas DataFrame mit den Abwesenheitsdaten
    :return: matplotlib Figure Objekt
    """
    fig, ax = plt.subplots(figsize=(10, 6))
    sns.histplot(data=df, x='absence_duration', bins=range(1, df['absence_duration'].max()+2), kde=False, color='skyblue', ax=ax)
    ax.set_title('Verteilung der Abwesenheitsdauer')
    ax.set_xlabel('Abwesenheitsdauer (Tage)')
    ax.set_ylabel('Anzahl der Abwesenheiten')
    ax.set_xticks(range(1, df['absence_duration'].max()+1))
    ax.grid(axis='y', linestyle='--', alpha=0.7)
    plt.tight_layout()
    return fig

# Methode zur Erstellung des Boxplots
def plot_absence_boxplot(df):
    """
    Erstellt einen Boxplot zur Verteilung der Abwesenheitsdauer nach Abwesenheitsgrund.
    
    :param df: Pandas DataFrame mit den Abwesenheitsdaten
    :return: matplotlib Figure Objekt
    """
    fig, ax = plt.subplots(figsize=(10, 6))
    sns.boxplot(data=df, x='absence_reason', y='absence_duration', palette='Set2', ax=ax)
    ax.set_title('Boxplot der Abwesenheitsdauer nach Abwesenheitsgrund')
    ax.set_xlabel('Abwesenheitsgrund')
    ax.set_ylabel('Abwesenheitsdauer (Tage)')
    ax.grid(axis='y', linestyle='--', alpha=0.7)
    plt.tight_layout()
    return fig

def main():
    st.title("Mitarbeiter Abwesenheitsanalyse")
    
    # Datei-Upload
    uploaded_file = st.file_uploader("CSV-Datei hochladen", type=["csv"])
    if uploaded_file is not None:
        data = load_employee_absence_data(uploaded_file)
        if data is not None:
            data = calculate_absence_duration(data)
            
            # Abwesenheitsgründe Auswahl
            absence_reasons = data['absence_reason'].unique().tolist()
            selected_reasons = st.multiselect("Abwesenheitsgründe auswählen", absence_reasons, default=absence_reasons)
            
            # Daten filtern nach Abwesenheitsgründen
            if selected_reasons:
                filtered_data = data[data['absence_reason'].isin(selected_reasons)]
            else:
                filtered_data = data.copy()
            
            if filtered_data.empty:
                st.warning("Keine Daten für die ausgewählten Abwesenheitsgründe verfügbar.")
            else:
                # Visualisierung der Abwesenheitsdauerverteilung
                st.subheader("Verteilung der Abwesenheitsdauer")
                fig_hist = plot_absence_distribution(filtered_data)
                st.pyplot(fig_hist)
                
                # Visualisierung der Abwesenheitsdauer nach Grund
                st.subheader("Boxplot der Abwesenheitsdauer nach Abwesenheitsgrund")
                fig_box = plot_absence_boxplot(filtered_data)
                st.pyplot(fig_box)
                
                # Optional: Download der gefilterten Daten
                st.download_button(
                    label="Gefilterte Daten herunterladen",
                    data=filtered_data.to_csv(index=False).encode('utf-8'),
                    file_name='filtered_employee_absence.csv',
                    mime='text/csv',
                )

if __name__ == "__main__":
    main()
