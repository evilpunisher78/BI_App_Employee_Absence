CSV_DATEI = "abwesenheitsaufzeichnungen.csv"

WOCHENTAGE = ["Montag", "Dienstag", "Mittwoch", "Donnerstag", "Freitag", "Samstag", "Sonntag"]
MONATE = ["Januar", "Februar", "März", "April", "Mai", "Juni",
          "Juli", "August", "September", "Oktober", "November", "Dezember"]
ABWESENHEITSGRUENDE = ["Krank", "Urlaub", "Persönliche Gründe", "Fortbildung"]

WOCHENTAG_MAP = dict(enumerate(WOCHENTAGE))
MONAT_MAP = dict(enumerate(MONATE, 1))

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
