# -*- coding: utf-8 -*-
"""
Analyse-Anwendung für Mitarbeiter-Abwesenheiten
Erstellt am 4. Januar 2025
Letzte Aktualisierung: 2025-01-22 17:03:07 UTC
@author: Helena Baranowsky, Katja Eppendorfer
"""

import os
import dash
from dash import dcc, html, Input, Output, State, dash_table
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import date
import uuid
import webbrowser

# ----------------------------------------------------
# (A) Konstanten und Konfiguration
# ----------------------------------------------------
CSV_DATEI = "abwesenheitsaufzeichnungen.csv"

WOCHENTAGE = ["Montag", "Dienstag", "Mittwoch", "Donnerstag", "Freitag", "Samstag", "Sonntag"]
MONATE = ["Januar", "Februar", "März", "April", "Mai", "Juni",
          "Juli", "August", "September", "Oktober", "November", "Dezember"]
ABWESENHEITSGRUENDE = ["Krank", "Urlaub", "Persönliche Gründe", "Fortbildung"]

# Mapping-Dictionaries für schnelleren Zugriff
WOCHENTAG_MAP = dict(enumerate(WOCHENTAGE))
MONAT_MAP = dict(enumerate(MONATE, 1))

# Zentrale Style-Definitionen
STYLES = {
    "container": {
        "backgroundColor": "#ffffff",
        "border": "1px solid #ddd",
        "borderRadius": "8px",
        "boxShadow": "0 2px 4px rgba(0, 0, 0, 0.1)",
        "padding": "20px",
        "marginBottom": "20px",
    },
    "button": {
        "backgroundColor": "#0056b3",
        "color": "#fff",
        "border": "none",
        "borderRadius": "4px",
        "padding": "10px 15px",
        "cursor": "pointer",
    },
    "heading": {"color": "#0056b3"},
    "flex_container": {
        "display": "flex",
        "alignItems": "center",
        "gap": "20px"
    }
}

# ----------------------------------------------------
# (B) Hilfsfunktionen
# ----------------------------------------------------
def load_data():
    """Lädt die Daten aus der CSV-Datei"""
    try:
        df = pd.read_csv(CSV_DATEI, sep=";", parse_dates=["Startdatum", "Enddatum"])
        df["Startdatum"] = df["Startdatum"].dt.normalize()
        df["Enddatum"] = df["Enddatum"].dt.normalize()
        if not df.empty:
            df["Fehltage"] = (df["Enddatum"] - df["Startdatum"]).dt.days + 1
        return df
    except FileNotFoundError:
        return pd.DataFrame(columns=["Mitarbeiter-ID", "Name", "Startdatum", "Enddatum", "Grund"])

def expand_abwesenheiten(df):
    """Expandiert das DataFrame für die Visualisierung"""
    if df.empty:
        return pd.DataFrame()

    expanded = pd.DataFrame([
        {
            "Mitarbeiter-ID": row["Mitarbeiter-ID"],
            "Name": row["Name"],
            "Datum": date,
            "Grund": row["Grund"]
        }
        for _, row in df.iterrows()
        for date in pd.date_range(row["Startdatum"], row["Enddatum"])
        if not pd.isna(row["Startdatum"]) and not pd.isna(row["Enddatum"])
    ])

    if not expanded.empty:
        expanded["Wochentag"] = expanded["Datum"].dt.weekday.map(WOCHENTAG_MAP)
        expanded["Monat"] = expanded["Datum"].dt.month.map(MONAT_MAP)
    
    return expanded

def create_krank_uebersicht(df):
    """Erstellt die Krank-Übersicht mit Smileys"""
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

def generate_figures(expanded_df):
    """Generiert alle Visualisierungen"""
    if expanded_df.empty:
        dummy = px.bar(title="Keine Daten verfügbar")
        return dummy, dummy, dummy, dummy

    # Grund-Trends
    grund_trends = (
        expanded_df.groupby("Grund")["Datum"]
        .count()
        .reset_index(name="Tage")
    )
    grund_figure = px.bar(
        grund_trends,
        x="Grund",
        y="Tage",
        color="Grund",
        title="Abwesenheitstrends nach Grund (Tage)"
    )
    grund_figure.update_layout(legend_title_text="Abwesenheitsgrund")

    # Wochentag-Trends
    wochentag_trends = (
        expanded_df.groupby(["Wochentag", "Grund"])["Datum"]
        .count()
        .reset_index(name="Tage")
    )
    wochentag_trends["sort_index"] = wochentag_trends["Wochentag"].map(
        lambda x: WOCHENTAGE.index(x)
    )
    wochentag_trends = wochentag_trends.sort_values(["sort_index", "Grund"])
    
    wochentag_figure = px.bar(
        wochentag_trends,
        x="Wochentag",
        y="Tage",
        color="Grund",
        barmode="group",
        title="Abwesenheitstrends nach Wochentag und Grund"
    )
    wochentag_figure.update_layout(legend_title_text="Abwesenheitsgrund")

    # Monats-Trends
    monat_trends = (
        expanded_df.groupby(["Monat", "Grund"])["Datum"]
        .count()
        .reset_index(name="Tage")
    )
    monat_trends["sort_index"] = monat_trends["Monat"].map(lambda x: MONATE.index(x))
    monat_trends = monat_trends.sort_values(["sort_index", "Grund"])

    monat_figure = create_monthly_figure(monat_trends)
    statistik_figure = create_statistics_figure(expanded_df)

    return grund_figure, wochentag_figure, monat_figure, statistik_figure

def create_monthly_figure(monat_trends):
    """Erstellt das monatliche Trend-Diagramm"""
    fig = go.Figure()
    
    for monat in MONATE:
        monat_data = monat_trends[monat_trends["Monat"] == monat]
        if not monat_data.empty:
            for grund in monat_data["Grund"].unique():
                wert = monat_data[monat_data["Grund"] == grund]["Tage"].values[0]
                fig.add_trace(
                    go.Bar(
                        name=f"{monat} - {grund}",
                        x=[monat],
                        y=[wert],
                        legendgroup=monat,
                        showlegend=True
                    )
                )

    fig.update_layout(
        title="Abwesenheitstrends nach Monat und Grund",
        barmode="group",
        xaxis_title="Monat",
        yaxis_title="Tage",
        showlegend=True,
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="right",
            x=1,
            groupclick="toggleitem"
        )
    )
    return fig

def create_statistics_figure(expanded_df):
    """Erstellt das statistische Analyse-Diagramm"""
    if expanded_df.empty:
        return px.line(title="Keine Daten verfügbar")

    # Erstelle vollständigen Datumsbereich
    min_date = expanded_df["Datum"].min()
    max_date = expanded_df["Datum"].max()
    all_dates = pd.date_range(start=min_date, end=max_date, freq='D')
    
    # Erstelle DataFrame mit allen Tagen
    all_days_df = pd.DataFrame({'Datum': all_dates})
    all_days_df['Monat'] = all_days_df['Datum'].dt.month.map(MONAT_MAP)
    
    # Zähle Abwesenheiten pro Tag
    daily_absences = (
        expanded_df.groupby(['Datum'])
        .size()
        .reset_index(name='Anzahl_Abwesenheiten')
    )
    
    # Füge Abwesenheiten hinzu
    all_days_df = all_days_df.merge(daily_absences, on='Datum', how='left')
    all_days_df['Anzahl_Abwesenheiten'] = all_days_df['Anzahl_Abwesenheiten'].fillna(0)
    
    # Berechne Statistiken
    stats_df = (
        all_days_df.groupby('Monat')
        .agg({
            'Anzahl_Abwesenheiten': [
                ('Durchschnitt', 'mean'),
                ('Std', 'std'),
                ('Max', 'max'),
                ('Min', 'min'),
                ('Tage_mit_Abwesenheit', lambda x: (x > 0).sum()),
                ('Tage_gesamt', 'count')
            ]
        })
    )
    
    stats_df.columns = stats_df.columns.droplevel(0)
    stats_df = stats_df.reset_index()
    stats_df['Std'] = stats_df['Std'].fillna(0)
    stats_df['Abwesenheitsquote'] = (
        stats_df['Tage_mit_Abwesenheit'] / stats_df['Tage_gesamt'] * 100
    ).round(1)
    
    # Sortiere nach Monaten
    stats_df['Monat_Sort'] = stats_df['Monat'].map(lambda x: MONATE.index(x))
    stats_df = stats_df.sort_values('Monat_Sort')
    
    return create_statistics_plot(stats_df)

def create_statistics_plot(stats_df):
    """Erstellt das Plot-Objekt für die statistische Analyse"""
    fig = go.Figure()
    
    # Hauptlinie (Durchschnitt)
    fig.add_trace(
        go.Scatter(
            name='Durchschnittliche Abwesenheiten pro Tag',
            x=stats_df['Monat'],
            y=stats_df['Durchschnitt'],
            line=dict(color='rgb(31, 119, 180)', width=2),
            mode='lines+markers'
        )
    )
    
    # Maximum und Minimum
    fig.add_trace(
        go.Scatter(
            name='Maximum pro Tag',
            x=stats_df['Monat'],
            y=stats_df['Max'],
            line=dict(color='rgba(255, 0, 0, 0.5)', dash='dash'),
            mode='lines'
        )
    )
    
    fig.add_trace(
        go.Scatter(
            name='Minimum pro Tag',
            x=stats_df['Monat'],
            y=stats_df['Min'],
            line=dict(color='rgba(0, 255, 0, 0.5)', dash='dash'),
            mode='lines'
        )
    )
    
    # Konfidenzbereich
    fig.add_trace(
        go.Scatter(
            name='±1 Standardabweichung',
            x=stats_df['Monat'].tolist() + stats_df['Monat'].tolist()[::-1],
            y=(stats_df['Durchschnitt'] + stats_df['Std']).tolist() + 
              (stats_df['Durchschnitt'] - stats_df['Std']).tolist()[::-1],
            fill='toself',
            fillcolor='rgba(31, 119, 180, 0.2)',
            line=dict(color='rgba(255,255,255,0)'),
            hoverinfo='skip',
            showlegend=True
        )
    )
    
    # Annotations
    annotations = [
        dict(
            x=row['Monat'],
            y=row['Durchschnitt'],
            text=(f"Ø: {row['Durchschnitt']:.2f}/Tag<br>"
                  f"σ: {row['Std']:.2f}<br>"
                  f"Tage mit Abw.: {row['Tage_mit_Abwesenheit']}/{row['Tage_gesamt']}<br>"
                  f"Quote: {row['Abwesenheitsquote']}%"),
            showarrow=True,
            arrowhead=7,
            ax=0,
            ay=-40
        )
        for _, row in stats_df.iterrows()
    ]
    
    fig.update_layout(
        title='Statistische Analyse der Abwesenheiten pro Tag und Monat',
        xaxis_title='Monat',
        yaxis_title='Anzahl Abwesenheiten pro Tag',
        hovermode='x unified',
        showlegend=True,
        annotations=annotations,
        legend=dict(
            orientation='h',
            yanchor='bottom',
            y=1.02,
            xanchor='right',
            x=1
        )
    )
    
    return fig

def filter_date_range(df, start_date, end_date):
    """Filtert das DataFrame nach Datumsbereich"""
    if not all([start_date, end_date]):
        return None, "Bitte wählen Sie ein Start- und Enddatum aus."
        
    start_dt = pd.to_datetime(start_date)
    end_dt = pd.to_datetime(end_date)
    
    if start_dt > end_dt:
        return None, "Das Startdatum darf nicht nach dem Enddatum liegen!"

    filtered_df = df[
        (df["Startdatum"] >= start_dt) &
        (df["Enddatum"] <= end_dt)
    ]

    if filtered_df.empty:
        return None, "Keine Daten im ausgewählten Zeitraum gefunden!"

    return filtered_df, ""

# ----------------------------------------------------
# (C) Dash-App Initialisierung
# ----------------------------------------------------
app = dash.Dash(__name__)
app.title = "Mitarbeiter-Abwesenheitsmanagement (Deutsch)"

# Lade initiale Daten
abwesenheiten = load_data()
expanded_initial = expand_abwesenheiten(abwesenheiten)
grund_fig_init, wochentag_fig_init, monat_fig_init, statistik_fig_init = generate_figures(expanded_initial)
initial_krank_uebersicht_df = create_krank_uebersicht(abwesenheiten)

# ----------------------------------------------------
# (D) Layout
# ----------------------------------------------------
app.layout = html.Div(
    style={"backgroundColor": "#f4f7fb", "padding": "20px", "maxWidth": "1200px", "margin": "auto"},
    children=[
        # Header
        html.H1("Mitarbeiter-Abwesenheitsmanagement", 
                style={"textAlign": "center", "color": "#0056b3"}),
        html.H4("Dieses Dashboard gehört zum Projekt FHD 2025 Modul Wirtschaftsinformatik, "
                "erstellt von Helena Baranowsky und Katja Eppendorfer",
                style={"textAlign": "center", "color": "#0056b3"}),

        # Eingabeformular
        html.Div(
            style=STYLES["container"],
            children=[
                html.H3("Abwesenheit hinzufügen", style=STYLES["heading"]),
                html.Div(
                    style=STYLES["flex_container"],
                    children=[
                        html.Div(
                            style={"flex": "1"},
                            children=[
                                html.Label("Name", style={"fontWeight": "bold"}),
                                dcc.Input(
                                    id="mitarbeiter_name",
                                    type="text",
                                    placeholder="Name des Mitarbeiters",
                                    style={"width": "100%"}
                                ),
                            ],
                        ),
                        html.Div(
                            style={"flex": "1"},
                            children=[
                                html.Label("Startdatum", style={"fontWeight": "bold"}),
                                dcc.DatePickerSingle(
                                    id="start_datum",
                                    date=date.today(),
                                    style={"width": "100%"}
                                ),
                            ],
                        ),
                        html.Div(
                            style={"flex": "1"},
                            children=[
                                html.Label("Enddatum", style={"fontWeight": "bold"}),
                                dcc.DatePickerSingle(
                                    id="end_datum",
                                    date=date.today(),
                                    style={"width": "100%"}
                                ),
                            ],
                        ),
                        html.Div(
                            style={"flex": "1.5"},
                            children=[
                                html.Label("Grund", style={"fontWeight": "bold"}),
                                dcc.Dropdown(
                                    id="grund_dropdown",
                                    options=[{"label": g, "value": g} for g in ABWESENHEITSGRUENDE]
                                    + [{"label": "Andere", "value": "Andere"}],
                                    placeholder="Grund auswählen",
                                    style={"width": "100%"}
                                ),
                                dcc.Input(
                                    id="anderer_grund",
                                    type="text",
                                    placeholder="Anderen Grund angeben",
                                    style={"display": "none", "width": "100%"}
                                ),
                            ],
                        ),
                    ],
                ),
                html.Button(
                    "Abwesenheit hinzufügen",
                    id="abwesenheit_hinzufuegen",
                    n_clicks=0,
                    style={**STYLES["button"], "marginTop": "20px"}
                ),
                html.Div(id="abwesenheit_rueckmeldung", 
                        style={"color": "green", "marginTop": "10px"}),
            ],
        ),

        # Abwesenheitstabelle und Export
        html.Div(
            style=STYLES["container"],
            children=[
                html.H3("Abwesenheitsaufzeichnungen", style=STYLES["heading"]),
                dash_table.DataTable(
                    id="abwesenheit_tabelle",
                    columns=[{"name": c, "id": c} for c in abwesenheiten.columns],
                    style_table={"overflowX": "auto"},
                    data=abwesenheiten.to_dict("records"),
                ),
                # Export-Bereich
                html.Div(
                    style={"marginTop": "20px"},
                    children=[
                        html.H4("Zeitraum für Export auswählen:", 
                               style={**STYLES["heading"], "marginBottom": "10px"}),
                        html.Div(
                            style=STYLES["flex_container"],
                            children=[
                                html.Div([
                                    html.Label("Von:", style={"fontWeight": "bold"}),
                                    dcc.DatePickerSingle(
                                        id="export_start_datum",
                                        date=date.today(),
                                        style={"width": "100%"}
                                    ),
                                ]),
                                html.Div([
                                    html.Label("Bis:", style={"fontWeight": "bold"}),
                                    dcc.DatePickerSingle(
                                        id="export_end_datum",
                                        date=date.today(),
                                        style={"width": "100%"}
                                    ),
                                ]),
                            ]
                        ),
                        html.Div(
                            style={"marginTop": "20px", "display": "flex", "gap": "20px"},
                            children=[
                                html.Button(
                                    "CSV herunterladen",
                                    id="download_csv",
                                    style=STYLES["button"]
                                ),
                                html.Button(
                                    "Excel herunterladen",
                                    id="download_excel",
                                    style=STYLES["button"]
                                )
                            ]
                        ),
                        html.Div(id="export_error_message", 
                                style={"color": "red", "marginTop": "10px"}),
                    ]
                ),
                dcc.Download(id="csv_download"),
                dcc.Download(id="excel_download"),
            ],
        ),

        # Krank-Übersicht
        html.Div(
            style=STYLES["container"],
            children=[
                html.H3("Übersicht: Summe Krank-Fehltage pro Mitarbeiter (mit Smiley)", 
                        style=STYLES["heading"]),
                dash_table.DataTable(
                    id="ma_uebersicht_krank_tabelle",
                    columns=[
                        {"name": "Mitarbeiter-ID", "id": "Mitarbeiter-ID"},
                        {"name": "Name", "id": "Name"},
                        {"name": "Summe Krank-Fehltage", "id": "Summe Krank-Fehltage"},
                        {"name": "Smiley", "id": "Smiley"},
                    ],
                    style_table={"overflowX": "auto"},
                    data=initial_krank_uebersicht_df.to_dict("records") 
                         if not initial_krank_uebersicht_df.empty else []
                ),
            ],
        ),

        # Diagramme
        html.Div(
            style=STYLES["container"],
            children=[
                html.H3("Abwesenheitstrends", style=STYLES["heading"]),
                dcc.Graph(id="abwesenheit_trends", figure=grund_fig_init),
                dcc.Graph(id="wochentag_trends", figure=wochentag_fig_init),
                dcc.Graph(id="monat_trends", figure=monat_fig_init),
            ],
        ),

        # Statistik-Diagramm
        html.Div(
            style=STYLES["container"],
            children=[
                html.H3("Statistische Analyse", style=STYLES["heading"]),
                dcc.Graph(id="statistik_trends", figure=statistik_fig_init),
            ],
        ),
    ],
)

# ----------------------------------------------------
# (E) Callbacks
# ----------------------------------------------------
@app.callback(
    Output("anderer_grund", "style"),
    Input("grund_dropdown", "value"),
    prevent_initial_call=True
)
def toggle_anderen_grund_feld(grund):
    """Zeigt/Versteckt das Feld für andere Gründe"""
    return {"display": "block", "width": "100%"} if grund == "Andere" else {"display": "none"}

@app.callback(
    [
        Output("abwesenheit_rueckmeldung", "children"),
        Output("abwesenheit_tabelle", "data"),
        Output("ma_uebersicht_krank_tabelle", "data"),
        Output("abwesenheit_trends", "figure"),
        Output("wochentag_trends", "figure"),
        Output("monat_trends", "figure"),
        Output("statistik_trends", "figure"),
    ],
    Input("abwesenheit_hinzufuegen", "n_clicks"),
    [
        State("mitarbeiter_name", "value"),
        State("start_datum", "date"),
        State("end_datum", "date"),
        State("grund_dropdown", "value"),
        State("anderer_grund", "value"),
    ],
    prevent_initial_call=True
)
def abwesenheit_hinzufuegen(n_clicks, name, start_datum, end_datum, grund, anderer_grund):
    """Fügt eine neue Abwesenheit hinzu und aktualisiert alle Ansichten"""
    global abwesenheiten

    if not all([name, start_datum, end_datum, grund]):
        return "Alle Felder müssen ausgefüllt werden!", abwesenheiten.to_dict("records"), [], px.bar(title="Keine Daten verfügbar"), px.bar(title="Keine Daten verfügbar"), px.bar(title="Keine Daten verfügbar"), px.bar(title="Keine Daten verfügbar")

    start_dt = pd.to_datetime(start_datum).normalize()
    end_dt = pd.to_datetime(end_datum).normalize()
    
    if start_dt > end_dt:
        return "Das Startdatum darf nicht nach dem Enddatum liegen!", abwesenheiten.to_dict("records"), [], px.bar(title="Keine Daten verfügbar"), px.bar(title="Keine Daten verfügbar"), px.bar(title="Keine Daten verfügbar"), px.bar(title="Keine Daten verfügbar")

    # Verwende anderen Grund falls ausgewählt
    if grund == "Andere" and anderer_grund:
        grund = anderer_grund

    # Finde existierende oder erstelle neue Mitarbeiter-ID
    existing_row = abwesenheiten[abwesenheiten["Name"] == name].head(1)
    mitarbeiter_id = existing_row["Mitarbeiter-ID"].iloc[0] if not existing_row.empty else f"EMP-{uuid.uuid4().hex[:8]}"

    # Füge neuen Eintrag hinzu
    neuer_eintrag = {
        "Mitarbeiter-ID": mitarbeiter_id,
        "Name": name,
        "Startdatum": start_dt,
        "Enddatum": end_dt,
        "Grund": grund
    }

    abwesenheiten = pd.concat([abwesenheiten, pd.DataFrame([neuer_eintrag])], ignore_index=True)
    abwesenheiten["Fehltage"] = (abwesenheiten["Enddatum"] - abwesenheiten["Startdatum"]).dt.days + 1

    # Speichere CSV
    abwesenheiten.to_csv(CSV_DATEI, sep=";", index=False)

    # Aktualisiere alle Ansichten
    expanded_df = expand_abwesenheiten(abwesenheiten)
    updated_krank_uebersicht_df = create_krank_uebersicht(abwesenheiten)
    grund_fig, wochentag_fig, monat_fig, statistik_fig = generate_figures(expanded_df)

    return (
        "Abwesenheit erfolgreich hinzugefügt!",
        abwesenheiten.to_dict("records"),
        updated_krank_uebersicht_df.to_dict("records"),
        grund_fig,
        wochentag_fig,
        monat_fig,
        statistik_fig
    )

@app.callback(
    [Output("csv_download", "data"),
     Output("export_error_message", "children")],
    [Input("download_csv", "n_clicks")],
    [State("export_start_datum", "date"),
     State("export_end_datum", "date")],
    prevent_initial_call=True
)
def download_csv(n_clicks, start_datum, end_datum):
    """Handled den CSV-Download mit Datumfilter"""
    if n_clicks is None:
        raise dash.exceptions.PreventUpdate

    filtered_df, error_message = filter_date_range(abwesenheiten, start_datum, end_datum)
    if error_message:
        return None, error_message

    return dcc.send_data_frame(filtered_df.to_csv, "abwesenheitsaufzeichnungen.csv", 
                              index=False, sep=";"), ""

@app.callback(
    [Output("excel_download", "data"),
     Output("export_error_message", "children", allow_duplicate=True)],
    [Input("download_excel", "n_clicks")],
    [State("export_start_datum", "date"),
     State("export_end_datum", "date")],
    prevent_initial_call=True
)
def download_excel(n_clicks, start_datum, end_datum):
    """Handled den Excel-Download mit Datumfilter"""
    if n_clicks is None:
        raise dash.exceptions.PreventUpdate

    filtered_df, error_message = filter_date_range(abwesenheiten, start_datum, end_datum)
    if error_message:
        return None, error_message

    return dcc.send_data_frame(
        filtered_df.to_excel,
        "abwesenheitsaufzeichnungen.xlsx",
        sheet_name="Abwesenheiten",
        index=False,
        engine='openpyxl'
    ), ""

# ----------------------------------------------------
# (F) Start der Anwendung
# ----------------------------------------------------
if __name__ == "__main__":
    print("Starte Mitarbeiter-Abwesenheitsmanagement...")
    print(f"Letzte Aktualisierung: 2025-01-22 17:06:28")
    print(f"Benutzer: evilpunisher78")
    webbrowser.open("http://127.0.0.1:8050/")
    app.run_server(debug=True)