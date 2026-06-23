# Partitionstausch – Ursache & Lösung

**Betreff:** Partitionstausch – Ursache & Lösung

Der Partitionstausch ist nicht am Daten**typ** selbst gescheitert, sondern an einer **Abweichung bei Präzision/Skala** (z. B. Quelle `decimal(38,0)` vs. Ziel `decimal(17,10)`), was SQL Server beim `SWITCH` ablehnt.

Es wird jetzt jede Spalte auf die **in unserer SQL-Server-Template-/Faktentabelle definierte Präzision/Skala konvertiert**, sodass Staging und Ziel exakt übereinstimmen und der Tausch erfolgreich durchläuft.

---

**Subject:** Partition Switch – Cause & Fix

The partition switch was not failing on the data **type** itself but on a **precision/scale mismatch** (e.g. source `decimal(38,0)` vs. target `decimal(17,10)`), which SQL Server rejects on `SWITCH`.

It now converts each column to the **precision/scale defined in our SQL Server template/fact table**, so staging and target match exactly and the switch succeeds.
