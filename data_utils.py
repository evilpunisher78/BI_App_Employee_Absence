import pandas as pd
from datetime import date
from constants import CSV_DATEI, WOCHENTAG_MAP, MONAT_MAP


def load_data():
    """Lädt die Daten aus der CSV-Datei."""
    try:
        df = pd.read_csv(CSV_DATEI, sep=";", parse_dates=["Startdatum", "Enddatum"])
        df["Startdatum"] = df["Startdatum"].dt.normalize()
        df["Enddatum"] = df["Enddatum"].dt.normalize()
        if not df.empty:
            df["Fehltage"] = (df["Enddatum"] - df["Startdatum"]).dt.days + 1
        return df
    except FileNotFoundError:
        return pd.DataFrame(columns=["Mitarbeiter-ID", "Name", "Startdatum", "Enddatum", "Grund"])


def expand_abwesenheiten(df: pd.DataFrame) -> pd.DataFrame:
    """Expandiert das DataFrame für die Visualisierung."""
    if df.empty:
        return pd.DataFrame()

    expanded = pd.DataFrame([
        {
            "Mitarbeiter-ID": row["Mitarbeiter-ID"],
            "Name": row["Name"],
            "Datum": current,
            "Grund": row["Grund"],
        }
        for _, row in df.iterrows()
        for current in pd.date_range(row["Startdatum"], row["Enddatum"])
        if not pd.isna(row["Startdatum"]) and not pd.isna(row["Enddatum"])
    ])

    if not expanded.empty:
        expanded["Wochentag"] = expanded["Datum"].dt.weekday.map(WOCHENTAG_MAP)
        expanded["Monat"] = expanded["Datum"].dt.month.map(MONAT_MAP)

    return expanded


def create_krank_uebersicht(df: pd.DataFrame) -> pd.DataFrame:
    """Erstellt die Krank-Übersicht mit Smileys."""
    if df.empty:
        return pd.DataFrame(columns=["Mitarbeiter-ID", "Name", "Summe Krank-Fehltage", "Smiley"])

    krank_df = df[df["Grund"] == "Krank"]
    if krank_df.empty:
        return pd.DataFrame(columns=["Mitarbeiter-ID", "Name", "Summe Krank-Fehltage", "Smiley"])

    uebersicht = (
        krank_df
        .groupby(["Mitarbeiter-ID", "Name"])["Fehltage"]
        .sum()
        .reset_index()
        .rename(columns={"Fehltage": "Summe Krank-Fehltage"})
    )

    uebersicht["Smiley"] = uebersicht["Summe Krank-Fehltage"].apply(
        lambda x: "😄" if x <= 10 else "😐" if x <= 20 else "😕" if x <= 30 else "😢"
    )
    return uebersicht


def filter_date_range(df: pd.DataFrame, start_date, end_date):
    """Filtert das DataFrame nach Datumsbereich."""
    if not all([start_date, end_date]):
        return None, "Bitte wählen Sie ein Start- und Enddatum aus."

    start_dt = pd.to_datetime(start_date)
    end_dt = pd.to_datetime(end_date)

    if start_dt > end_dt:
        return None, "Das Startdatum darf nicht nach dem Enddatum liegen!"

    filtered_df = df[(df["Startdatum"] >= start_dt) & (df["Enddatum"] <= end_dt)]

    if filtered_df.empty:
        return None, "Keine Daten im ausgewählten Zeitraum gefunden!"

    return filtered_df, ""
