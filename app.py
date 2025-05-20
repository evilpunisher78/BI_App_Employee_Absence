import dash
from dash import dcc, html, Input, Output, State, dash_table
import pandas as pd
import plotly.express as px
import uuid
from datetime import date

from constants import STYLES, ABWESENHEITSGRUENDE, CSV_DATEI
from data_utils import (
    load_data,
    expand_abwesenheiten,
    create_krank_uebersicht,
    filter_date_range,
)
from figures import generate_figures

app = dash.Dash(__name__)
app.title = "Mitarbeiter-Abwesenheitsmanagement (Deutsch)"

abwesenheiten = load_data()
expanded_initial = expand_abwesenheiten(abwesenheiten)
(
    grund_fig_init,
    wochentag_fig_init,
    monat_fig_init,
    statistik_fig_init,
) = generate_figures(expanded_initial)
initial_krank_uebersicht_df = create_krank_uebersicht(abwesenheiten)

app.layout = html.Div(
    style={"backgroundColor": "#f4f7fb", "padding": "20px", "maxWidth": "1200px", "margin": "auto"},
    children=[
        html.H1(
            "Mitarbeiter-Abwesenheitsmanagement",
            style={"textAlign": "center", "color": "#0056b3"},
        ),
        html.H4(
            "Dieses Dashboard gehört zum Projekt FHD 2025 Modul Wirtschaftsinformatik, erstellt von Helena Baranowsky und Katja Eppendorfer",
            style={"textAlign": "center", "color": "#0056b3"},
        ),
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
                                    style={"width": "100%"},
                                ),
                            ],
                        ),
                        html.Div(
                            style={"flex": "1"},
                            children=[
                                html.Label("Startdatum", style={"fontWeight": "bold"}),
                                dcc.DatePickerSingle(id="start_datum", date=date.today(), style={"width": "100%"}),
                            ],
                        ),
                        html.Div(
                            style={"flex": "1"},
                            children=[
                                html.Label("Enddatum", style={"fontWeight": "bold"}),
                                dcc.DatePickerSingle(id="end_datum", date=date.today(), style={"width": "100%"}),
                            ],
                        ),
                        html.Div(
                            style={"flex": "1.5"},
                            children=[
                                html.Label("Grund", style={"fontWeight": "bold"}),
                                dcc.Dropdown(
                                    id="grund_dropdown",
                                    options=[{"label": g, "value": g} for g in ABWESENHEITSGRUENDE] + [{"label": "Andere", "value": "Andere"}],
                                    placeholder="Grund auswählen",
                                    style={"width": "100%"},
                                ),
                                dcc.Input(
                                    id="anderer_grund",
                                    type="text",
                                    placeholder="Anderen Grund angeben",
                                    style={"display": "none", "width": "100%"},
                                ),
                            ],
                        ),
                    ],
                ),
                html.Button(
                    "Abwesenheit hinzufügen",
                    id="abwesenheit_hinzufuegen",
                    n_clicks=0,
                    style={**STYLES["button"], "marginTop": "20px"},
                ),
                html.Div(id="abwesenheit_rueckmeldung", style={"color": "green", "marginTop": "10px"}),
            ],
        ),
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
                html.Div(
                    style={"marginTop": "20px"},
                    children=[
                        html.H4(
                            "Zeitraum für Export auswählen:",
                            style={**STYLES["heading"], "marginBottom": "10px"},
                        ),
                        html.Div(
                            style=STYLES["flex_container"],
                            children=[
                                html.Div([
                                    html.Label("Von:", style={"fontWeight": "bold"}),
                                    dcc.DatePickerSingle(id="export_start_datum", date=date.today(), style={"width": "100%"}),
                                ]),
                                html.Div([
                                    html.Label("Bis:", style={"fontWeight": "bold"}),
                                    dcc.DatePickerSingle(id="export_end_datum", date=date.today(), style={"width": "100%"}),
                                ]),
                            ],
                        ),
                        html.Div(
                            style={"marginTop": "20px", "display": "flex", "gap": "20px"},
                            children=[
                                html.Button("CSV herunterladen", id="download_csv", style=STYLES["button"]),
                                html.Button("Excel herunterladen", id="download_excel", style=STYLES["button"]),
                            ],
                        ),
                        html.Div(id="export_error_message", style={"color": "red", "marginTop": "10px"}),
                    ],
                ),
                dcc.Download(id="csv_download"),
                dcc.Download(id="excel_download"),
            ],
        ),
        html.Div(
            style=STYLES["container"],
            children=[
                html.H3("Übersicht: Summe Krank-Fehltage pro Mitarbeiter (mit Smiley)", style=STYLES["heading"]),
                dash_table.DataTable(
                    id="ma_uebersicht_krank_tabelle",
                    columns=[
                        {"name": "Mitarbeiter-ID", "id": "Mitarbeiter-ID"},
                        {"name": "Name", "id": "Name"},
                        {"name": "Summe Krank-Fehltage", "id": "Summe Krank-Fehltage"},
                        {"name": "Smiley", "id": "Smiley"},
                    ],
                    style_table={"overflowX": "auto"},
                    data=initial_krank_uebersicht_df.to_dict("records") if not initial_krank_uebersicht_df.empty else [],
                ),
            ],
        ),
        html.Div(
            style=STYLES["container"],
            children=[
                html.H3("Abwesenheitstrends", style=STYLES["heading"]),
                dcc.Graph(id="abwesenheit_trends", figure=grund_fig_init),
                dcc.Graph(id="wochentag_trends", figure=wochentag_fig_init),
                dcc.Graph(id="monat_trends", figure=monat_fig_init),
            ],
        ),
        html.Div(
            style=STYLES["container"],
            children=[
                html.H3("Statistische Analyse", style=STYLES["heading"]),
                dcc.Graph(id="statistik_trends", figure=statistik_fig_init),
            ],
        ),
    ],
)

@app.callback(Output("anderer_grund", "style"), Input("grund_dropdown", "value"), prevent_initial_call=True)
def toggle_anderen_grund_feld(grund):
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
    prevent_initial_call=True,
)
def abwesenheit_hinzufuegen(n_clicks, name, start_datum, end_datum, grund, anderer_grund):
    global abwesenheiten
    if not all([name, start_datum, end_datum, grund]):
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
    end_dt = pd.to_datetime(end_datum).normalize()
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
    if grund == "Andere" and anderer_grund:
        grund = anderer_grund
    existing_row = abwesenheiten[abwesenheiten["Name"] == name].head(1)
    mitarbeiter_id = existing_row["Mitarbeiter-ID"].iloc[0] if not existing_row.empty else f"EMP-{uuid.uuid4().hex[:8]}"
    neuer_eintrag = {
        "Mitarbeiter-ID": mitarbeiter_id,
        "Name": name,
        "Startdatum": start_dt,
        "Enddatum": end_dt,
        "Grund": grund,
    }
    abwesenheiten = pd.concat([abwesenheiten, pd.DataFrame([neuer_eintrag])], ignore_index=True)
    abwesenheiten["Fehltage"] = (abwesenheiten["Enddatum"] - abwesenheiten["Startdatum"]).dt.days + 1
    abwesenheiten.to_csv(CSV_DATEI, sep=";", index=False)
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
        statistik_fig,
    )

@app.callback(
    [Output("csv_download", "data"), Output("export_error_message", "children")],
    [Input("download_csv", "n_clicks")],
    [State("export_start_datum", "date"), State("export_end_datum", "date")],
    prevent_initial_call=True,
)
def download_csv(n_clicks, start_datum, end_datum):
    if n_clicks is None:
        raise dash.exceptions.PreventUpdate
    filtered_df, error_message = filter_date_range(abwesenheiten, start_datum, end_datum)
    if error_message:
        return None, error_message
    return dcc.send_data_frame(filtered_df.to_csv, "abwesenheitsaufzeichnungen.csv", index=False, sep=";"), ""


@app.callback(
    [Output("excel_download", "data"), Output("export_error_message", "children", allow_duplicate=True)],
    [Input("download_excel", "n_clicks")],
    [State("export_start_datum", "date"), State("export_end_datum", "date")],
    prevent_initial_call=True,
)
def download_excel(n_clicks, start_datum, end_datum):
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
        engine="openpyxl",
    ), ""
