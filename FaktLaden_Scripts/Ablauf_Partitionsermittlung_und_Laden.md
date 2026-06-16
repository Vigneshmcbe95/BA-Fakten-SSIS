# Faktenladen – Ablauf Partitionsermittlung → Laden (für Sebastian)

## 🇩🇪 E-Mail

**Betreff: Faktenladen – frühere 0-Zeilen-Meldung erledigt + kurzer Ablauf**

Hallo Sebastian,

die frühere Fehlermeldung („… lieferte 0 Zeilen …“) ist **nicht mehr gültig** –
sie kam von leeren Grenzpartitionen. Das wird jetzt sauber abgefangen. Kurz
zum Ablauf, wie ermittelt und geladen wird:

**1. Was der Anwender vorgibt (Steuerliste):**
Eine Zeile pro Tabelle, z. B.
- `tf_bvd_sgb2` → laden
- `tf_bvd_sgb2_*` → alle Partitionen (automatisch)
- `tf_bvd_sgb2:202701` → genau diese Partition

**2. Wie geprüft wird (SCR11 – Partitionsgrenzen):**
- Liest aus den **Oracle-Partitions-Metadaten**, welche Monate es gibt
  (kein Scan über Milliarden Zeilen).
- Liest, welche Monate in **SQL Server** schon vorhanden sind.
- **Differenz** = nur die fehlenden Monate werden geladen.
- **Leere Grenzpartitionen** (Monat ohne Daten, z. B. die erste/letzte Grenze)
  werden erkannt und übersprungen.
- Ist nichts Neues da → Meldung **„Alle Daten bereits geladen“**.

**3. Wie geladen wird (SCR13 – Daten laden):**
- Für jeden ermittelten Monat: Daten aus Oracle (`ext.<tabelle>`) je Partition
  nach SQL Server laden, dann per Partition-Switch in die Faktentabelle.
- Kommt hier wider Erwarten ein Monat **ohne Daten** an, **bricht** der Lauf
  mit klarer Meldung ab (statt still zu überspringen).

**Ein Beispiel (`tf_bvd_sgb2`):**
- Oracle hat Monate **200707–202701**, SQL Server hat schon **200707–202703**.
- Ergebnis: **nichts Neues** → „Alle Daten bereits geladen
  (Oracle 200707-202701 | MSSQL 200707-202703)“.
- Käme in Oracle ein neuer Monat **202702** dazu, würde **nur dieser eine**
  Monat geladen und eingeschoben.

Viele Grüße

---

## 🇬🇧 Flow (reference)

**User gives** (control list): `table`, `table_*` (all), or `table:value` (one).
**SCR11 checks**: Oracle months (metadata, no scan) vs SQL months → load only the
difference; skip empty boundary partitions; if none → "all already loaded".
**SCR13 loads**: per missing month, Oracle → SQL per partition → partition switch.
A truly empty month arriving at SCR13 now **fails loudly** (clear message).

**Example** `tf_bvd_sgb2`: Oracle 200707–202701, SQL 200707–202703 → nothing new →
"all data already loaded". New Oracle month 202702 → only that one is loaded.
