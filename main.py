from app import app
import webbrowser

if __name__ == "__main__":
    print("Starte Mitarbeiter-Abwesenheitsmanagement...")
    webbrowser.open("http://127.0.0.1:8050/")
    app.run(debug=True)
