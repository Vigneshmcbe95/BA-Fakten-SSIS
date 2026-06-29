# Steuerliste – Unterstützte Filter-/Partitions-Schreibweisen (Fakten Laden)

Verarbeitet durch **SCR03** (Tabellennamen extrahieren), **SCR04** (Filter → Wert/Regex) und **SCR11** (Partitionsauswahl gegen die realen Oracle-Werte).

**Referenzdatum** = 1. des laufenden Monats (Beispiele unten mit Referenzmonat **Juni 2026** → `202606`, Referenzjahr `2026`).

**Zwei Verhalten:**
- **Cut-Off** = der berechnete/feste Wert ist eine **Obergrenze**; es werden **alle** Partitionen **≤ Wert** geladen (vorhandene werden neu geladen = Full Load bis Cut-Off).
- **Mengen-Treffer** = es werden **genau die Partitionen** geladen, die auf das Muster passen.

> In **beiden** Fällen werden nur Partitionen geladen, die in Oracle **tatsächlich Daten** enthalten – leere Grenz-/Ankerpartitionen werden automatisch übersprungen.

| Schreibweise (Beispiel) | Typ | Vom Skript erzeugt | Was geladen wird |
|---|---|---|---|
| `tabelle` | AUTOMATIC | – | Alle in Oracle vorhandenen Partitionen (Delta gegen MSSQL) |
| `tabelle_*` | Wildcard (alle) | Regex `^.{0,128}$` | **Alle** Partitionen |
| `tabelle_202*` | Wildcard-Präfix | Regex `^202.{0,128}$` | Alle Partitionen, die mit `202` beginnen (2020er) |
| `tabelle_2026*` | Wildcard-Präfix | Regex `^2026.{0,128}$` | Alle Partitionen des Jahres 2026 |
| `tabelle_20260[1-6]` | Bereich (Klammer) | Regex `^20260[1-6]$` | `202601`–`202606` (Jan–Jun 2026) |
| `tabelle_2026[01][0-9]` | Klammerklasse | Regex `^2026[01][0-9]$` | Passende Monate 2026 |
| `tabelle_2026?` | `?`-Quantor | Regex `^2026?$` | `202` oder `2026` (selten genutzt) |
| `tabelle_202604` | Fester Wert (6-stellig) | `partition_wert = 202604` | **Cut-Off:** alle Partitionen ≤ `202604` |
| `tabelle_20260401` | Fester Wert (8-stellig) | `partition_wert = 20260401` | **Cut-Off:** alle ≤ `20260401` |
| `tabelle:202604` | Fester Wert nach `:` | `partition_wert = 202604` | **Cut-Off:** alle ≤ `202604` |
| `tabelle:MONID(-1)` | Relativer Monat | `202605` | **Cut-Off:** alle ≤ `202605` |
| `tabelle:MONID6(0)` | Relativer Monat | `202606` | **Cut-Off:** alle ≤ `202606` |
| `tabelle_:MONID4(-1)` | Rel. Monat, 2-stell. Jahr | `2605` (JJMM) | **Cut-Off** (Suffix-Form) |
| `tabelle:YEAR(-1)` | Relatives Jahr | `2025` | **Cut-Off:** alle ≤ `2025` |
| `tabelle:YYYYMM(202601,202604)` | IN-Liste Monate | Regex `^(202601|202604)` | Genau diese Monate |
| `tabelle:YYYY(2025,2026)` | IN-Liste Jahre | Regex `^(2025|2026)` | Alle Partitionen aus 2025 **und** 2026 |
| `tabelle:LAST_MM(3)` | Letzte n Monate | Regex `^(202604|202605|202606)` | Letzte 3 Monate (inkl. aktuellem) |
| `tabelle:LAST_YYYYMM(6)` | Letzte n Monate | Regex `^(202601|…|202606)` | Letzte 6 Monate (inkl. aktuellem) |
| `tabelle:LAST_YY(1)` | Letzte n Jahre | Regex `^(2025|2026)` | Letztes + aktuelles Jahr |
| `tabelle:LAST_YYYY(2)` | Letzte n Jahre | Regex `^(2024|2025|2026)` | Letzte 2 Jahre (inkl. aktuellem) |

## Hinweise

- **Cut-Off** greift nur bei **einzelnen** festen Werten und Datums-Token (`:MONID`, `:MONID6`, `:MONID4`, `:YEAR`). Ein einzelner Wert in einer IN-Liste (z. B. `:YYYYMM(202604)`) gilt als **Mengen-Treffer** und lädt **nur** diesen Monat.
- **Wildcards:** `*` = beliebige Zeichenfolge (intern `.{0,128}`), `?` und `+` werden direkt als Regex-Quantoren übernommen.
- **Klammern:** `[12]`, `[0-9]`, `[5-9]` (Bereich) werden direkt als Regex-Zeichenklasse übernommen.
- **Trennung Tabellenname / Partitionsteil:** entweder nach `:` (Token/Wert) oder als Suffix nach `_` (z. B. `tf_bb_bewa_zkt_202*`).
- Das `*` im Tabellen-/Objektnamen selbst (Oracle-Objektsuche per `star2regex`) ist hier **nicht** abgebildet – diese Tabelle betrifft nur die **Partitionsauswahl**.
