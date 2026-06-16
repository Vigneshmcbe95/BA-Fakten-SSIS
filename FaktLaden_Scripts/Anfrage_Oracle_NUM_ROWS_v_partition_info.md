# Anfrage an Oracle/DBA: Spalte NUM_ROWS in der Partitionssicht ergänzen

## Hintergrund (kurz)
Der SSIS-Faktenladeprozess liest die zu ladenden Partitionswerte aus der
Oracle-View `V_PARTITION_INFO` (per PolyBase als `ext.v_partition_info`).
Aktuell liefert die View u. a. `HIGH_VALUE`, aber **keine Zeilenanzahl**.
Dadurch lassen sich **leere Grenzpartitionen** (z. B. die unterste Partition
ohne Daten) nicht rein über Metadaten erkennen. Mit einer Spalte `NUM_ROWS`
könnten wir leere Partitionen direkt und ohne Tabellen-Scan herausfiltern.

---

## 🇩🇪 E-Mail

**Betreff: Bitte um Ergänzung der View V_PARTITION_INFO um Spalte NUM_ROWS**

Hallo zusammen,

für den Faktenladeprozess lesen wir die Partitionsinformationen über die View
`V_PARTITION_INFO` (Schema `BI_DM_EXPORT`) per PolyBase aus.

Könntet ihr die View bitte um die Spalte **`NUM_ROWS`** aus
`DBA_TAB_PARTITIONS` (bzw. `ALL_TAB_PARTITIONS`) erweitern? Damit können wir
**leere Grenzpartitionen** (Partitionen ohne Daten) bereits anhand der
Metadaten erkennen und überspringen – ohne einen teuren Tabellen-Scan.

Beispiel:
```sql
CREATE OR REPLACE VIEW BI_DM_EXPORT.V_PARTITION_INFO AS
SELECT
    p.TABLE_OWNER          AS OWNER,
    p.TABLE_NAME           AS TABLE_NAME,
    t.PARTITIONING_TYPE    AS PARTITIONING_TYPE,
    t.SUBPARTITIONING_TYPE AS SUBPARTITIONING_TYPE,
    pk.PARTITION_KEY       AS PARTITION_KEY,
    p.PARTITION_POSITION   AS PARTITION_POSITION,
    p.PARTITION_NAME       AS PARTITION_NAME,
    p.HIGH_VALUE           AS HIGH_VALUE,
    p.NUM_ROWS             AS NUM_ROWS      -- << neu
FROM   DBA_TAB_PARTITIONS p
       /* ... bestehende Joins für PARTITIONING_TYPE / PARTITION_KEY ... */;
```

Hinweise:
- `NUM_ROWS` stammt aus der Optimizer-Statistik. Für eine verlässliche Anzeige
  sollte die Statistik der betroffenen Tabellen aktuell sein (`DBMS_STATS`).
- Es genügt eine Näherung (> 0 / = 0); wir nutzen den Wert nur, um leere
  Grenzpartitionen auszublenden.

Vielen Dank und viele Grüße

---

## 🇬🇧 Summary (for reference)
Please add a **`NUM_ROWS`** column (from `DBA_TAB_PARTITIONS`) to the
`V_PARTITION_INFO` view so the ETL can filter out **empty boundary partitions**
from metadata alone (no table scan). `NUM_ROWS` comes from optimizer stats, so
the affected tables' stats should be reasonably fresh; we only need it as a
> 0 / = 0 indicator.

---

## Nach der Oracle-Änderung (unsererseits)
1. In `SCR_02_Vorbereitungen` die externe Tabelle `ext.v_partition_info` um die
   Spalte `NUM_ROWS FLOAT NULL` ergänzen (Spaltenanzahl muss zur View passen).
2. In `SCR11_Partitionsgrenzen` (`OracleAlleWerteLaden`) den Filter
   `AND ISNULL(NUM_ROWS,0) > 0` ergänzen.
3. Danach ist die Einzel-Partition-Prüfung im FULL-LOAD-Zweig (Option A)
   nicht mehr nötig und kann entfallen.
