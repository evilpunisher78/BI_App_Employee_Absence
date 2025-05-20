import plotly.express as px
import plotly.graph_objects as go
import pandas as pd
from constants import MONATE, MONAT_MAP, WOCHENTAGE


def generate_figures(expanded_df: pd.DataFrame):
    """Generiert alle Visualisierungen."""
    if expanded_df.empty:
        dummy = px.bar(title="Keine Daten verfügbar")
        return dummy, dummy, dummy, dummy

    grund_trends = (
        expanded_df.groupby("Grund")["Datum"].count().reset_index(name="Tage")
    )
    grund_figure = px.bar(
        grund_trends,
        x="Grund",
        y="Tage",
        color="Grund",
        title="Abwesenheitstrends nach Grund (Tage)",
    )
    grund_figure.update_layout(legend_title_text="Abwesenheitsgrund")

    wochentag_trends = (
        expanded_df.groupby(["Wochentag", "Grund"])["Datum"].count().reset_index(name="Tage")
    )
    wochentag_trends["sort_index"] = wochentag_trends["Wochentag"].map(lambda x: WOCHENTAGE.index(x))
    wochentag_trends = wochentag_trends.sort_values(["sort_index", "Grund"])
    wochentag_figure = px.bar(
        wochentag_trends,
        x="Wochentag",
        y="Tage",
        color="Grund",
        barmode="group",
        title="Abwesenheitstrends nach Wochentag und Grund",
    )
    wochentag_figure.update_layout(legend_title_text="Abwesenheitsgrund")

    monat_trends = (
        expanded_df.groupby(["Monat", "Grund"])["Datum"].count().reset_index(name="Tage")
    )
    monat_trends["sort_index"] = monat_trends["Monat"].map(lambda x: MONATE.index(x))
    monat_trends = monat_trends.sort_values(["sort_index", "Grund"])
    monat_figure = create_monthly_figure(monat_trends)
    statistik_figure = create_statistics_figure(expanded_df)

    return grund_figure, wochentag_figure, monat_figure, statistik_figure


def create_monthly_figure(monat_trends: pd.DataFrame):
    """Erstellt das monatliche Trend-Diagramm."""
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
                        showlegend=True,
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
            groupclick="toggleitem",
        ),
    )
    return fig


def create_statistics_figure(expanded_df: pd.DataFrame):
    """Erstellt das statistische Analyse-Diagramm."""
    if expanded_df.empty:
        return px.line(title="Keine Daten verfügbar")

    min_date = expanded_df["Datum"].min()
    max_date = expanded_df["Datum"].max()
    all_dates = pd.date_range(start=min_date, end=max_date, freq="D")
    all_days_df = pd.DataFrame({"Datum": all_dates})
    all_days_df["Monat"] = all_days_df["Datum"].dt.month.map(MONAT_MAP)

    daily_absences = (
        expanded_df.groupby(["Datum"]).size().reset_index(name="Anzahl_Abwesenheiten")
    )
    all_days_df = all_days_df.merge(daily_absences, on="Datum", how="left")
    all_days_df["Anzahl_Abwesenheiten"] = all_days_df["Anzahl_Abwesenheiten"].fillna(0)

    stats_df = (
        all_days_df.groupby("Monat").agg({
            "Anzahl_Abwesenheiten": [
                ("Durchschnitt", "mean"),
                ("Std", "std"),
                ("Max", "max"),
                ("Min", "min"),
                ("Tage_mit_Abwesenheit", lambda x: (x > 0).sum()),
                ("Tage_gesamt", "count"),
            ]
        })
    )

    stats_df.columns = stats_df.columns.droplevel(0)
    stats_df = stats_df.reset_index()
    stats_df["Std"] = stats_df["Std"].fillna(0)
    stats_df["Abwesenheitsquote"] = (
        stats_df["Tage_mit_Abwesenheit"] / stats_df["Tage_gesamt"] * 100
    ).round(1)
    stats_df["Monat_Sort"] = stats_df["Monat"].map(lambda x: MONATE.index(x))
    stats_df = stats_df.sort_values("Monat_Sort")

    return create_statistics_plot(stats_df)


def create_statistics_plot(stats_df: pd.DataFrame):
    """Erstellt das Plot-Objekt für die statistische Analyse."""
    fig = go.Figure()
    fig.add_trace(
        go.Scatter(
            name="Durchschnittliche Abwesenheiten pro Tag",
            x=stats_df["Monat"],
            y=stats_df["Durchschnitt"],
            line=dict(color="rgb(31, 119, 180)", width=2),
            mode="lines+markers",
        )
    )
    fig.add_trace(
        go.Scatter(
            name="Maximum pro Tag",
            x=stats_df["Monat"],
            y=stats_df["Max"],
            line=dict(color="rgba(255, 0, 0, 0.5)", dash="dash"),
            mode="lines",
        )
    )
    fig.add_trace(
        go.Scatter(
            name="Minimum pro Tag",
            x=stats_df["Monat"],
            y=stats_df["Min"],
            line=dict(color="rgba(0, 255, 0, 0.5)", dash="dash"),
            mode="lines",
        )
    )
    fig.add_trace(
        go.Scatter(
            name="±1 Standardabweichung",
            x=stats_df["Monat"].tolist() + stats_df["Monat"].tolist()[::-1],
            y=(stats_df["Durchschnitt"] + stats_df["Std"]).tolist() +
              (stats_df["Durchschnitt"] - stats_df["Std"]).tolist()[::-1],
            fill="toself",
            fillcolor="rgba(31, 119, 180, 0.2)",
            line=dict(color="rgba(255,255,255,0)"),
            hoverinfo="skip",
            showlegend=True,
        )
    )
    annotations = [
        dict(
            x=row["Monat"],
            y=row["Durchschnitt"],
            text=(
                f"Ø: {row['Durchschnitt']:.2f}/Tag<br>"
                f"σ: {row['Std']:.2f}<br>"
                f"Tage mit Abw.: {row['Tage_mit_Abwesenheit']}/{row['Tage_gesamt']}<br>"
                f"Quote: {row['Abwesenheitsquote']}%"
            ),
            showarrow=True,
            arrowhead=7,
            ax=0,
            ay=-40,
        )
        for _, row in stats_df.iterrows()
    ]
    fig.update_layout(
        title="Statistische Analyse der Abwesenheiten pro Tag und Monat",
        xaxis_title="Monat",
        yaxis_title="Anzahl Abwesenheiten pro Tag",
        hovermode="x unified",
        showlegend=True,
        annotations=annotations,
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="right",
            x=1,
        ),
    )
    return fig
