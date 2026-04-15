import pandas as pd

TOLERANCE = 0.2

# Excel einlesen
df = pd.read_excel("BAC 0,1 mM Werte.xlsx")

# Daten in Gruppen extrahieren (jede Paar-Spalte enthält Zeit und Wert)
groups = []
for i in range(0, df.shape[1], 2):
    times = df.iloc[:, i]
    values = df.iloc[:, i + 1]
    group = []
    for t, v in zip(times, values):
        if pd.notna(t) and pd.notna(v):
            group.append((float(t), float(v)))
    groups.append(group)

# Vergleichen mit bis zu 100 Zeilen oberhalb/unterhalb innerhalb des Fehlerbereichs
results = []
seen_keys = set()

for ref_group_idx, ref_group in enumerate(groups):
    for ref_idx, (ref_time, ref_value) in enumerate(ref_group):
        match_indices = [(ref_group_idx, ref_idx)]
        matched_times = [ref_time]
        matched_values = [ref_value]
        all_matched = True

        for other_group_idx, other_group in enumerate(groups):
            if other_group_idx == ref_group_idx:
                continue

            candidates = []
            for other_idx, (other_time, other_value) in enumerate(other_group):
                if abs(other_time - ref_time) <= TOLERANCE:
                    candidates.append((abs(other_time - ref_time), other_idx, other_time, other_value))

            if not candidates:
                all_matched = False
                break

            best_match = min(candidates, key=lambda item: item[0])
            _, matched_idx, matched_time, matched_value = best_match
            matched_times.append(matched_time)
            matched_values.append(matched_value)
            match_indices.append((other_group_idx, matched_idx))

        if not all_matched:
            continue

        key = tuple(sorted(match_indices))
        if key in seen_keys:
            continue

        seen_keys.add(key)
        mean_time = sum(matched_times) / len(matched_times)
        mean_value = sum(matched_values) / len(matched_values)
        results.append((mean_time, mean_value))

# Schritt 3: Ergebnis in DataFrame
result_df = pd.DataFrame(results, columns=["Zeit", "Mittelwert"])
result_df = result_df.sort_values(by="Zeit")

# Schritt 4: Neue Excel speichern
output_file = "BAC 0,1mM Durchschnitt.xlsx"
result_df.to_excel(output_file, index=False)

print(f"Fertig! Datei gespeichert als: {output_file}")