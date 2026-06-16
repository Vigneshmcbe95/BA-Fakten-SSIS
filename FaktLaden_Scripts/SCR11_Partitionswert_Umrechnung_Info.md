# Fakten-Laden – Partitionswert-Umrechnung (SCR11)

---

## 🇩🇪 Deutsch

**Betreff: Fakten-Laden – Korrektur der Partitionswert-Umrechnung (SCR11)**

Hallo zusammen,

die fehlgeschlagenen Partitionen in `SCR13` kamen daher, dass `SCR11` die
Oracle-Grenzwerte (`HIGH_VALUE`) falsch in echte Datenwerte umgerechnet hat.
Der Vergleich `HIGH_VALUE` ↔ `MON_ID` zeigt **zwei Formate**, die
unterschiedlich behandelt werden müssen.

**6-stellig (YYYYMM, z. B. `mon_id`):**
- `HIGH_VALUE` = Datenwert (`200705` → `200705`).
- Jahresende-Marker `YYYY13` → Januar Folgejahr (`200713` → `200801`).
- Letzter Wert (z. B. `202606`) ist leer → wird entfernt.

**8-stellig (YYYYMM + 2 Ziffern Anhang, z. B. `mow_id`):**
- Letzte 2 Ziffern = Anhangnummer, **kein Tag**.
- Datenwert = **`HIGH_VALUE − 1`** (`…07` → `…06`).
- Erster Wert mit Anhang `00` (`19990100`) ist leer → wird übersprungen.
- Letzter Wert (`20251107` → `20251106`) hat Daten → bleibt erhalten.

Das Format entscheidet automatisch über die Methode. Leere Partitionen werden
jetzt mit Hinweis übersprungen statt als Fehler abgebrochen.

Viele Grüße

---

## 🇬🇧 English

**Subject: Fact Loading – Fix for partition value conversion (SCR11)**

Hi all,

The failing partitions in `SCR13` were caused by `SCR11` converting the Oracle
boundary values (`HIGH_VALUE`) into the wrong data values. Comparing
`HIGH_VALUE` ↔ `MON_ID` showed **two formats** that must be handled differently.

**6-digit (YYYYMM, e.g. `mon_id`):**
- `HIGH_VALUE` = data value (`200705` → `200705`).
- Year-end marker `YYYY13` → next-year January (`200713` → `200801`).
- Last value (e.g. `202606`) is empty → removed.

**8-digit (YYYYMM + 2-digit append, e.g. `mow_id`):**
- Last 2 digits = append number, **not a day**.
- Data value = **`HIGH_VALUE − 1`** (`…07` → `…06`).
- First value with append `00` (`19990100`) is empty → skipped.
- Last value (`20251107` → `20251106`) has data → kept.

The format decides the method automatically. Empty partitions are now skipped
with a note instead of failing the run.

Best regards

---

## Quick reference

| | First row | Last row |
|---|---|---|
| **6-digit** (`mon_id`) | `200705` → has data | `202606` → empty → dropped |
| **8-digit** (`mow_id`) | `19990100` → empty → skipped | `20251107` → has data (`20251106`) → kept |
