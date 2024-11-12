import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import random

# Parameter
num_employees = 50
total_entries = 200
start_date = datetime(2023, 1, 1)
end_date = datetime(2023, 12, 31)

# Generiere Mitarbeiter-IDs
employee_ids = [f"E{str(i).zfill(3)}" for i in range(1, num_employees + 1)]

# Abwesenheitsgründe und -typen
absence_reasons = ['Urlaub', 'Krankheit', 'unentschuldigtes Fehlen']
absence_types = ['geplante Abwesenheit', 'unerwartete Abwesenheit']

# Schritt 1: Weisen Sie jedem Mitarbeiter eine Anzahl von Urlaubstagen zwischen 28 und 33 zu
urlaub_tage_pro_mitarbeiter = {emp_id: random.randint(28, 33) for emp_id in employee_ids}

# Schritt 2: Teilen Sie die Urlaubstage in mehrere Urlaubseinträge auf (z.B. 1 bis 3 Urlaubsperioden)
urlaub_entries_per_mitarbeiter = {}
for emp_id, urlaub_tage in urlaub_tage_pro_mitarbeiter.items():
    num_periods = random.randint(1, 3)  # Jeder hat zwischen 1 und 3 Urlaubsperioden
    tage_pro_period = [urlaub_tage // num_periods] * num_periods
    for i in range(urlaub_tage % num_periods):
        tage_pro_period[i] += 1  # Verteile die restlichen Tage
    urlaub_entries_per_mitarbeiter[emp_id] = tage_pro_period

# Schritt 3: Generiere Urlaubseinträge
data = []
for emp_id, tage_liste in urlaub_entries_per_mitarbeiter.items():
    for tage in tage_liste:
        # Zufälliges Startdatum
        random_days = random.randint(0, (end_date - start_date).days - tage)
        absence_start = start_date + timedelta(days=random_days)
        
        # Enddatum basierend auf der Dauer
        absence_end = absence_start + timedelta(days=tage - 1)
        if absence_end > end_date:
            absence_end = end_date
        
        data.append({
            'employee_id': emp_id,
            'absence_start': absence_start.strftime('%Y-%m-%d'),
            'absence_end': absence_end.strftime('%Y-%m-%d'),
            'absence_reason': 'Urlaub',
            'absence_type': 'geplante Abwesenheit'
        })

# Schritt 4: Berechne die verbleibenden Einträge für weitere Abwesenheiten
# Zunächst zählen wir die bereits erstellten Urlaubseinträge
current_entries = len(data)
remaining_entries = total_entries - current_entries

# Wenn remaining_entries negativ ist, müssen wir die Gesamtanzahl erhöhen
if remaining_entries < 0:
    print(f"Warnung: Die Urlaubseinträge überschreiten die gewünschte Gesamtanzahl von {total_entries}.")
    remaining_entries = 0

# Schritt 5: Verteile die verbleibenden Abwesenheiten auf die Mitarbeiter
# Berechne durchschnittliche zusätzliche Abwesenheiten pro Mitarbeiter
avg_additional = remaining_entries / num_employees if num_employees else 0

# Generiere zusätzliche Abwesenheiten basierend auf normaler Verteilung
mean_additional = avg_additional
std_dev_additional = max(1, mean_additional / 2)  # Mindestens 1, um Variabilität zu gewährleisten

additional_absences_per_employee = np.random.normal(mean_additional, std_dev_additional, num_employees).astype(int)
additional_absences_per_employee = np.clip(additional_absences_per_employee, 0, None)  # Keine negativen Werte

# Korrigiere die Gesamtzahl der verbleibenden Einträge
current_total_additional = additional_absences_per_employee.sum()
difference = remaining_entries - current_total_additional

while difference != 0:
    for i in range(num_employees):
        if difference == 0:
            break
        if difference > 0:
            additional_absences_per_employee[i] += 1
            difference -= 1
        elif additional_absences_per_employee[i] > 0:
            additional_absences_per_employee[i] -= 1
            difference += 1

# Schritt 6: Generiere zusätzliche Abwesenheitseinträge (Krankheit, Sonderurlaub)
for emp_index, emp_id in enumerate(employee_ids):
    count = additional_absences_per_employee[emp_index]
    for _ in range(count):
        # Zufälliges Startdatum
        random_days = random.randint(0, (end_date - start_date).days)
        absence_start = start_date + timedelta(days=random_days)
        
        # Dauer der Abwesenheit (1 bis 10 Tage, normalverteilt)
        duration = max(1, int(np.random.normal(5, 2)))
        absence_end = absence_start + timedelta(days=duration)
        if absence_end > end_date:
            absence_end = end_date
        
        # Zufälliger Grund und Typ
        absence_reason = random.choice(absence_reasons[1:])
        absence_type = 'unerwartete Abwesenheit' #if absence_reason == 'Krankheit' else random.choice(absence_types)
        
        data.append({
            'employee_id': emp_id,
            'absence_start': absence_start.strftime('%Y-%m-%d'),
            'absence_end': absence_end.strftime('%Y-%m-%d'),
            'absence_reason': absence_reason,
            'absence_type': absence_type
        })

# Schritt 7: Optional - Mischen der Daten, um Urlaubseinträge nicht alle zuerst zu haben
random.shuffle(data)

# Schritt 8: Erstelle DataFrame und exportiere als CSV
df = pd.DataFrame(data)
df.to_csv('abwesenheiten.csv', index=False, encoding='utf-8-sig')

print("CSV-Datei 'abwesenheiten.csv' wurde erfolgreich erstellt.")
print(f"Gesamtanzahl der Einträge: {len(df)}")
