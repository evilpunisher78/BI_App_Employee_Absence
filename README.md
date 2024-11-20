# BI_App_Employee_Absence
Mitarbeiter Abwesenheiten visuallisieren und filtern

# Flowchart Übersicht:

**Start:** 
  
  Der Benutzer startet die App.

**Datenimport:** 
    
  Der Benutzer lädt Abwesenheitsdaten (CSV, Excel oder Datenbank) in die App.

**Datenbereinigung:** 

  Die App prüft die Daten auf Fehler oder fehlende Werte und bereinigt diese gegebenenfalls. (optional)

**Datenanalyse:**

  Berechnet die Abwesenheitstage der Mitarbeiter.
  Berechnet abteilungs- oder mitarbeiterspezifische Auswertungen.

**Visualisierung:**
    
  Darstellung der Analyseergebnisse in verschiedenen Diagrammen (z.B. Balken-, Liniendiagramme).
    
**Ergebnisse anzeigen:** 

  Zeigt den Endbericht oder die Ergebnisse der Analyse an (Abwesenheiten pro Monat, Abteilung, **was uns noch so einfällt...**). (Nur Zahlen)

**Benutzereingaben:** 
  Ermöglicht dem Benutzer, nach bestimmten Kriterien zu filtern (z.B. nach Abteilung, Zeitraum oder Abwesenheitsgrund, **was uns noch so einfällt...**).

**Berichterstellung:** 
  
  Möglichkeit, die Ergebnisse zu exportieren (z.B. als PDF oder CSV).

**Ende:** 

  Der Benutzer beendet die Anwendung.


    +--------------------------+
    |         Start             |
    +--------------------------+
               |
               v
    +--------------------------+
    |     Datenimport           |  <--- Benutzer lädt CSV/Excel/Datenbank
    +--------------------------+
               |
               v
    +--------------------------+
    |     Datenbereinigung      |  <--- Überprüfung auf fehlende Werte und Formatfehler
    +--------------------------+
               |
               v
    +--------------------------+
    |     Datenanalyse          |  <--- Berechnung der Abwesenheitstage, etc.
    +--------------------------+
               |
               v
    +--------------------------+
    |     Visualisierung        |  <--- Erstellung von Diagrammen
    +--------------------------+
               |
               v
    +--------------------------+
    |     Ergebnisse anzeigen   |  <--- Anzeigen der Analyseergebnisse
    +--------------------------+
               |
               v
    +--------------------------+
    |     Benutzereingaben      |  <--- Benutzer filtert nach Abteilung, Zeitraum, Abwesenheitsdauer etc.
    +--------------------------+
               |
               v
    +--------------------------+
    |     Berichterstellung     |  <--- Export der Ergebnisse (z.B. PDF/CSV)
    +--------------------------+
               |
               v
    +--------------------------+
    |         Ende              |
    +--------------------------+


