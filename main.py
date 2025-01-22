# -*- coding: utf-8 -*-
"""
Analyse-Anwendung für Mitarbeiter-Abwesenheiten
Erstellt am 4. Januar 2025
Letzte Aktualisierung: 2025-01-21 18:40:46 UTC
@author: Helena, Katja
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
import io
import base64

CSV_DATEI = "abwesenheitsaufzeichnungen.csv"

# Wir definieren die Wochentagsnamen und Monatsnamen auf Deutsch
WOCHENTAGE = ["Montag", "Dienstag", "Mittwoch", "Donnerstag", "Freitag", "Samstag", "Sonntag"]
MONATE = [
    "Januar", "Februar", "März", "April", "Mai", "Juni",
    "Juli", "August", "September", "Oktober", "November", "Dezember"
]

wochentag_map = {0: "Montag", 1: "Dienstag", 2: "Mittwoch", 3: "Donnerstag", 4: "Freitag", 5: "Samstag", 6: "Sonntag"}
monat_map = {
    1: "Januar", 2: "Februar", 3: "März", 4: "April", 5: "Mai", 6: "Juni",
    7: "Juli", 8: "August", 9: "September", 10: "Oktober", 11: "November", 12: "Dezember"
}

# ----------------------------------------------------
# (A) CSV einlesen, falls vorhanden
# ----------------------------------------------------
if os.path.exists(CSV_DATEI):
    abwesenheiten = pd.read_csv(CSV_DATEI, sep=";", parse_dates=["Startdatum", "Enddatum"])
    abwesenheiten["Startdatum"] = abwesenheiten["Startdatum"].dt.normalize()
    abwesenheiten["Enddatum"]   = abwesenheiten["Enddatum"].dt.normalize()
else:
    abwesenheiten = pd.DataFrame(columns=["Mitarbeiter-ID", "Name", "Startdatum", "Enddatum", "Grund"])

# ----------------------------------------------------
# (B) Hilfsfunktionen
# ----------------------------------------------------
def expand_abwesenheiten(df: pd.DataFrame) -> pd.DataFrame:
    """
    Erzeugt ein "expandiertes" DataFrame mit einer Zeile pro Tag
    (wichtig für die Diagramme).
    """
    all_rows = []
    for _, row in df.iterrows():
        start = row["Startdatum"]
        end   = row["Enddatum"]
        if pd.isna(start) or pd.isna(end):
            continue

        date_range = pd.date_range(start=start, end=end, freq="D")
        for single_date in date_range:
            all_rows.append({
                "Mitarbeiter-ID": row["Mitarbeiter-ID"],
                "Name": row["Name"],
                "Datum": single_date,
                "Grund": row["Grund"]
            })

    expanded = pd.DataFrame(all_rows)
    if not expanded.empty:
        expanded["Wochentag"] = expanded["Datum"].dt.weekday.map(wochentag_map)
        expanded["Monat"]     = expanded["Datum"].dt.month.map(monat_map)
    return expanded

def generate_figures_from_expanded(expanded_df: pd.DataFrame):
    """
    Erzeugt 4 Plotly-Figuren (Grund-, Wochentag-, Monatstrends und Statistik-Liniendiagramm)
    aus dem "expandierten" DataFrame in deutscher Sprache.
    """
    if expanded_df.empty:
        dummy = px.bar(title="Keine Daten verfügbar")
        return dummy, dummy, dummy, dummy

    # Grundtrends
    grund_trends = expanded_df.groupby("Grund")["Datum"].count().reset_index(name="Tage")
    grund_figure = px.bar(
        grund_trends, x="Grund", y="Tage", color="Grund",
        title="Abwesenheitstrends nach Grund (Tage)"
    )
    grund_figure.update_layout(legend_title_text="Abwesenheitsgrund")

    # Wochentagtrends
    wochentag_trends = expanded_df.groupby(["Wochentag", "Grund"])["Datum"].count().reset_index(name="Tage")
    wochentag_trends["sort_index"] = wochentag_trends["Wochentag"].apply(lambda x: WOCHENTAGE.index(x))
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

    # Monatstrends (modifiziert)
    monat_trends = expanded_df.groupby(["Monat", "Grund"])["Datum"].count().reset_index(name="Tage")
    monat_trends["sort_index"] = monat_trends["Monat"].apply(lambda m: MONATE.index(m))
    monat_trends = monat_trends.sort_values(["sort_index", "Grund"])
    
    # Erstelle separate Balken für jeden Monat
    monat_figure = go.Figure()
    
    # Füge für jeden Monat einen eigenen Trace hinzu
    for monat in MONATE:
        monat_data = monat_trends[monat_trends["Monat"] == monat]
        if not monat_data.empty:
            for grund in monat_data["Grund"].unique():
                wert = monat_data[monat_data["Grund"] == grund]["Tage"].values[0]
                monat_figure.add_trace(
                    go.Bar(
                        name=f"{monat} - {grund}",
                        x=[monat],
                        y=[wert],
                        legendgroup=monat,
                        showlegend=True
                    )
                )

    # Konfiguriere das Layout für das Monatsdiagramm
    monat_figure.update_layout(
        title="Abwesenheitstrends nach Monat und Grund",
        barmode="group",
        xaxis_title="Monat",
        yaxis_title="Tage",
        showlegend=True,
        legend=dict(
            orientation="h",     # horizontale Legende
            yanchor="bottom",
            y=1.02,             # Position über dem Diagramm
            xanchor="right",
            x=1,
            groupclick="toggleitem"  # Ermöglicht Einzelauswahl
        )
    )

   # Verbesserte statistische Berechnung
    # Zuerst erstellen wir einen vollständigen Datumsbereich für das gesamte Jahr
    if not expanded_df.empty:
        min_date = expanded_df["Datum"].min()
        max_date = expanded_df["Datum"].max()
        all_dates = pd.date_range(start=min_date, end=max_date, freq='D')
        
        # Erstelle ein DataFrame mit allen Tagen
        all_days_df = pd.DataFrame({'Datum': all_dates})
        all_days_df['Monat'] = all_days_df['Datum'].dt.month.map(monat_map)
        
        # Zähle Abwesenheiten pro Tag
        daily_absences = (
            expanded_df.groupby(['Datum'])
            .size()
            .reset_index(name='Anzahl_Abwesenheiten')
        )
        
        # Füge die Abwesenheiten dem all_days_df hinzu
        all_days_df = all_days_df.merge(
            daily_absences, 
            on='Datum', 
            how='left'
        )
        all_days_df['Anzahl_Abwesenheiten'] = all_days_df['Anzahl_Abwesenheiten'].fillna(0)
        
        # Berechne die Statistiken
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
        
        # Behandle NaN-Werte in der Standardabweichung
        stats_df['Std'] = stats_df['Std'].fillna(0)
        
        # Füge zusätzliche Informationen hinzu
        stats_df['Abwesenheitsquote'] = (stats_df['Tage_mit_Abwesenheit'] / stats_df['Tage_gesamt'] * 100).round(1)
        
        # Sortiere die Monate in der richtigen Reihenfolge
        stats_df['Monat_Sort'] = stats_df['Monat'].map(lambda x: MONATE.index(x))
        stats_df = stats_df.sort_values('Monat_Sort')
        
        # Erstelle das Liniendiagramm
        statistik_figure = go.Figure()
        
        # Füge die Hauptlinie (Durchschnitt) hinzu
        statistik_figure.add_trace(
            go.Scatter(
                name='Durchschnittliche Abwesenheiten pro Tag',
                x=stats_df['Monat'],
                y=stats_df['Durchschnitt'],
                line=dict(color='rgb(31, 119, 180)', width=2),
                mode='lines+markers'
            )
        )
        
        # Füge Minimum und Maximum als gestrichelte Linien hinzu
        statistik_figure.add_trace(
            go.Scatter(
                name='Maximum pro Tag',
                x=stats_df['Monat'],
                y=stats_df['Max'],
                line=dict(color='rgba(255, 0, 0, 0.5)', dash='dash'),
                mode='lines'
            )
        )
        
        statistik_figure.add_trace(
            go.Scatter(
                name='Minimum pro Tag',
                x=stats_df['Monat'],
                y=stats_df['Min'],
                line=dict(color='rgba(0, 255, 0, 0.5)', dash='dash'),
                mode='lines'
            )
        )
        
        # Füge den Konfidenzbereich hinzu (±1 Standardabweichung)
        statistik_figure.add_trace(
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
        
        # Füge detaillierte Annotations für die Werte hinzu
        annotations = []
        for idx, row in stats_df.iterrows():
            annotations.append(
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
            )
        
        # Update Layout
        statistik_figure.update_layout(
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
    else:
        statistik_figure = px.line(title="Keine Daten verfügbar")

    return grund_figure, wochentag_figure, monat_figure, statistik_figure

def create_krank_uebersicht_df(df: pd.DataFrame):
    """
    Erzeugt ein DataFrame mit aufsummierten Krank-Fehltagen pro Mitarbeiter
    und hängt eine "Smiley"-Spalte an.
    """
    krank_df = df[df["Grund"] == "Krank"]
    if krank_df.empty:
        return pd.DataFrame(columns=["Mitarbeiter-ID", "Name", "Summe Krank-Fehltage", "Smiley"])

    ma_uebersicht_krank = (
        krank_df
        .groupby(["Mitarbeiter-ID", "Name"])["Fehltage"]
        .sum()
        .reset_index()
        .rename(columns={"Fehltage": "Summe Krank-Fehltage"})
    )

    def get_smiley(tage):
        if tage <= 10:
            return "😄"
        elif tage <= 20:
            return "😐"
        elif tage <= 30:
            return "😕"
        else:
            return "😢"

    ma_uebersicht_krank["Smiley"] = ma_uebersicht_krank["Summe Krank-Fehltage"].apply(get_smiley)
    return ma_uebersicht_krank

# ----------------------------------------------------
# (C) Vorab Fehltage berechnen & initiale Krank-Übersicht
# ----------------------------------------------------
if not abwesenheiten.empty:
    abwesenheiten["Fehltage"] = (abwesenheiten["Enddatum"] - abwesenheiten["Startdatum"]).dt.days + 1

initial_krank_uebersicht_df = create_krank_uebersicht_df(abwesenheiten)

# Initiale Diagramme erstellen
expanded_initial = expand_abwesenheiten(abwesenheiten)
grund_fig_init, wochentag_fig_init, monat_fig_init, statistik_fig_init = generate_figures_from_expanded(expanded_initial)

# ----------------------------------------------------
# (D) Dash-App
# ----------------------------------------------------
app = dash.Dash(__name__)
app.title = "Mitarbeiter-Abwesenheitsmanagement (Deutsch)"

global_style = {
    "fontFamily": "Arial, sans-serif",
    "backgroundColor": "#f4f7fb",
    "color": "#333",
    "margin": "0",
    "padding": "0",
}

abwesenheitsgruende = ["Krank", "Urlaub", "Persönliche Gründe", "Fortbildung"]

app.layout = html.Div(
    style={"backgroundColor": global_style["backgroundColor"], "padding": "20px", "maxWidth": "1200px", "margin": "auto"},
    children=[
        # Titel
        html.H1(
            "Mitarbeiter-Abwesenheitsmanagement",
            style={"textAlign": "center", "color": "#0056b3", "fontFamily": global_style["fontFamily"]},
        ),
        html.H4(
            "Dieses Dashboard gehört zum Projekt FHD 2025 Modul Wirtschaftsinformatik, erstellt von Helena Baranowsky und Katja Eppendorfer",
            style={"textAlign": "center", "color": "#0056b3"},
        ),

        # Abschnitt: Abwesenheit hinzufügen
        html.Div(
            style={
                "backgroundColor": "#ffffff",
                "border": "1px solid #ddd",
                "borderRadius": "8px",
                "boxShadow": "0 2px 4px rgba(0, 0, 0, 0.1)",
                "padding": "20px",
                "marginBottom": "20px",
            },
            children=[
                html.H3("Abwesenheit hinzufügen", style={"color": "#0056b3"}),
                html.Div(
                    style={"display": "flex", "alignItems": "center", "gap": "20px"},
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
                                    options=[{"label": g, "value": g} for g in abwesenheitsgruende]
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
                    style={
                        "backgroundColor": "#0056b3",
                        "color": "#fff",
                        "border": "none",
                        "borderRadius": "4px",
                        "padding": "10px 15px",
                        "cursor": "pointer",
                        "marginTop": "20px",
                    },
                ),
                html.Div(id="abwesenheit_rueckmeldung", style={"color": "green", "marginTop": "10px"}),
            ],
        ),

        # Erste Tabelle (ein Eintrag pro Abwesenheit)
        html.Div(
            style={
                "backgroundColor": "#ffffff",
                "border": "1px solid #ddd",
                "borderRadius": "8px",
                "boxShadow": "0 2px 4px rgba(0, 0, 0, 0.1)",
                "padding": "20px",
                "marginBottom": "20px",
            },
            children=[
                html.H3("Abwesenheitsaufzeichnungen", style={"color": "#0056b3"}),
                dash_table.DataTable(
                    id="abwesenheit_tabelle",
                    columns=[{"name": c, "id": c} for c in abwesenheiten.columns],
                    style_table={"overflowX": "auto"},
                    data=abwesenheiten.to_dict("records"),
                ),
                html.Div(
                    style={"marginTop": "20px", "display": "flex", "gap": "20px"},
                    children=[
                        html.Div(
                            style={"marginTop": "20px", "marginBottom": "20px"},
                            children=[
                                html.H4("Zeitraum für Export auswählen:", style={"color": "#0056b3", "marginBottom": "10px"}),
                                html.Div(
                                    style={"display": "flex", "gap": "20px", "alignItems": "center"},
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
                                            style={
                                                "backgroundColor": "#0056b3",
                                                "color": "#fff",
                                                "border": "none",
                                                "borderRadius": "4px",
                                                "padding": "10px 15px"
                                            }
                                        ),
                                        html.Button(
                                            "Excel herunterladen",
                                            id="download_excel",
                                            style={
                                                "backgroundColor": "#0056b3",
                                                "color": "#fff",
                                                "border": "none",
                                                "borderRadius": "4px",
                                                "padding": "10px 15px"
                                            }
                                        )
                                    ]
                                ),
                                html.Div(id="export_error_message", style={"color": "red", "marginTop": "10px"}),
                            ]
                        ),                        
                    ]
                ),
                dcc.Download(id="csv_download"),
                dcc.Download(id="excel_download"),
            ],
        ),

        # Krank-Übersicht
        html.Div(
            style={
                "backgroundColor": "#ffffff",
                "border": "1px solid #ddd",
                "borderRadius": "8px",
                "boxShadow": "0 2px 4px rgba(0, 0, 0, 0.1)",
                "padding": "20px",
                "marginBottom": "20px",
            },
            children=[
                html.H3("Übersicht: Summe Krank-Fehltage pro Mitarbeiter (mit Smiley)", style={"color": "#0056b3"}),
                dash_table.DataTable(
                    id="ma_uebersicht_krank_tabelle",
                    columns=[
                        {"name": "Mitarbeiter-ID",          "id": "Mitarbeiter-ID"},
                        {"name": "Name",                    "id": "Name"},
                        {"name": "Summe Krank-Fehltage",    "id": "Summe Krank-Fehltage"},
                        {"name": "Smiley",                  "id": "Smiley"},
                    ],
                    style_table={"overflowX": "auto"},
                    data=initial_krank_uebersicht_df.to_dict("records") if not initial_krank_uebersicht_df.empty else []
                ),
            ],
        ),

        # Diagramme
        html.Div(
            style={
                "backgroundColor": "#ffffff",
                "border": "1px solid #ddd",
                "borderRadius": "8px",
                "boxShadow": "0 2px 4px rgba(0, 0, 0, 0.1)",
                "padding": "20px",
                "marginBottom": "20px",
            },
            children=[
                html.H3("Abwesenheitstrends", style={"color": "#0056b3"}),
                dcc.Graph(id="abwesenheit_trends", figure=grund_fig_init),
                dcc.Graph(id="wochentag_trends", figure=wochentag_fig_init),
                dcc.Graph(id="monat_trends", figure=monat_fig_init),
            ],
        ),

        # Statistik-Diagramm
        html.Div(
            style={
                "backgroundColor": "#ffffff",
                "border": "1px solid #ddd",
                "borderRadius": "8px",
                "boxShadow": "0 2px 4px rgba(0, 0, 0, 0.1)",
                "padding": "20px",
                "marginTop": "20px",
            },
            children=[
                html.H3("Statistische Analyse", style={"color": "#0056b3"}),
                dcc.Graph(id="statistik_trends", figure=statistik_fig_init),
            ],
        ),
    ],
)

# Callback: "Andere Gründe" -> Feld anzeigen
@app.callback(
    Output("anderer_grund", "style"),
    Input("grund_dropdown", "value"),
    prevent_initial_call=True
)
def toggle_anderen_grund_feld(grund):
    if grund == "Andere":
        return {"display": "block", "width": "100%"}
    return {"display": "none"}

# Callback: Neue Abwesenheit hinzufügen & aktualisieren
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
    global abwesenheiten

    if not name or not start_datum or not end_datum or not grund:
        return (
            "Alle Felder müssen ausgefüllt werden!",
            abwesenheiten.to_dict("records"),
            [],
            px.bar(title="Keine Daten verfügbar"),
            px.bar(title="Keine Daten verfügbar"),
            px.bar(title="Keine Daten verfügbar"),
            px.bar(title="Keine Daten verfügbar"),
        )

    start_dt = pd.to_datetime(start_datum).normalize()
    end_dt   = pd.to_datetime(end_datum).normalize()
    if start_dt > end_dt:
        return (
            "Das Startdatum darf nicht nach dem Enddatum liegen!",
            abwesenheiten.to_dict("records"),
            [],
            px.bar(title="Keine Daten verfügbar"),
            px.bar(title="Keine Daten verfügbar"),
            px.bar(title="Keine Daten verfügbar"),
            px.bar(title="Keine Daten verfügbar"),
        )

    if grund == "Andere":
        grund = anderer_grund

    existing_row = abwesenheiten[abwesenheiten["Name"] == name].head(1)
    if not existing_row.empty:
        mitarbeiter_id = existing_row["Mitarbeiter-ID"].iloc[0]
    else:
        mitarbeiter_id = f"EMP-{uuid.uuid4().hex[:8]}"

    neuer_eintrag = {
        "Mitarbeiter-ID": mitarbeiter_id,
        "Name": name,
        "Startdatum": start_dt,
        "Enddatum": end_dt,
        "Grund": grund
    }

    abwesenheiten = pd.concat([abwesenheiten, pd.DataFrame([neuer_eintrag])], ignore_index=True)
    abwesenheiten["Fehltage"] = (abwesenheiten["Enddatum"] - abwesenheiten["Startdatum"]).dt.days + 1

    # CSV abspeichern
    abwesenheiten.to_csv(CSV_DATEI, sep=";", index=False)

    # Neue Krank-Übersicht erzeugen
    updated_krank_uebersicht_df = create_krank_uebersicht_df(abwesenheiten)

    # Diagramme aktualisieren
    expanded_df = expand_abwesenheiten(abwesenheiten)
    grund_fig, wochentag_fig, monat_fig, statistik_fig = generate_figures_from_expanded(expanded_df)

    return (
        "Abwesenheit erfolgreich hinzugefügt!",
        abwesenheiten.to_dict("records"),
        updated_krank_uebersicht_df.to_dict("records"),
        grund_fig,
        wochentag_fig,
        monat_fig,
        statistik_fig
    )

# Aktualisierte Callback-Funktionen für den Download:
@app.callback(
    [Output("csv_download", "data"),
     Output("export_error_message", "children")],
    [Input("download_csv", "n_clicks")],
    [State("export_start_datum", "date"),
     State("export_end_datum", "date")],
    prevent_initial_call=True
)
def download_csv(n_clicks, start_datum, end_datum):
    if n_clicks is None:
        raise dash.exceptions.PreventUpdate

    if not start_datum or not end_datum:
        return None, "Bitte wählen Sie ein Start- und Enddatum aus."

    start_dt = pd.to_datetime(start_datum)
    end_dt = pd.to_datetime(end_datum)

    if start_dt > end_dt:
        return None, "Das Startdatum darf nicht nach dem Enddatum liegen!"

    # Filtere die Daten nach dem gewählten Zeitraum
    filtered_df = abwesenheiten[
        (abwesenheiten["Startdatum"] >= start_dt) &
        (abwesenheiten["Enddatum"] <= end_dt)
    ]

    if filtered_df.empty:
        return None, "Keine Daten im ausgewählten Zeitraum gefunden!"

    return dcc.send_data_frame(
        filtered_df.to_csv,
        "abwesenheitsaufzeichnungen.csv",
        index=False,
        sep=";"
    ), ""

@app.callback(
    [Output("excel_download", "data"),
     Output("export_error_message", "children", allow_duplicate=True)],
    [Input("download_excel", "n_clicks")],
    [State("export_start_datum", "date"),
     State("export_end_datum", "date")],
    prevent_initial_call=True
)
def download_excel(n_clicks, start_datum, end_datum):
    if n_clicks is None:
        raise dash.exceptions.PreventUpdate

    if not start_datum or not end_datum:
        return None, "Bitte wählen Sie ein Start- und Enddatum aus."

    start_dt = pd.to_datetime(start_datum)
    end_dt = pd.to_datetime(end_datum)

    if start_dt > end_dt:
        return None, "Das Startdatum darf nicht nach dem Enddatum liegen!"

    # Filtere die Daten nach dem gewählten Zeitraum
    filtered_df = abwesenheiten[
        (abwesenheiten["Startdatum"] >= start_dt) &
        (abwesenheiten["Enddatum"] <= end_dt)
    ]

    if filtered_df.empty:
        return None, "Keine Daten im ausgewählten Zeitraum gefunden!"

    return dcc.send_data_frame(
        filtered_df.to_excel,
        "abwesenheitsaufzeichnungen.xlsx",
        sheet_name="Abwesenheiten",
        index=False,
        engine='openpyxl'
    ), ""

if __name__ == "__main__":
    webbrowser.open("http://127.0.0.1:8050/")
    app.run_server(debug=True)