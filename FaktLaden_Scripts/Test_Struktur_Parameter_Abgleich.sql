-- =====================================================================
-- Abgleich Parameter FaktenClusteredIndex <-> echte Faktentabelle
-- In der Fakten-Datenbank ausfuehren. Vorher anpassen:
--   <ParameterDB>        z.B. Datenbank mit der Parametertabelle
--   <Parametertabelle>   z.B. tm_<db>_param
-- Ergebnis 'FEHLER' = SCR10 bricht dieses Verfahren ab
-- (Tabelle loeschen, wird neu angelegt, oder Parameter anpassen).
-- =====================================================================
SELECT  p.Verfahren,
        p.Wert                                   AS Faktentabelle,
        UPPER(pi.Wert)                           AS Parameter_Index,
        CASE UPPER(pi.Wert) WHEN 'CCI' THEN 'CCI'
                            WHEN 'CI'  THEN 'CI'
                            WHEN 'TRUE' THEN 'CI'
                            ELSE 'Heap' END      AS Soll,
        CASE x.typ WHEN 5 THEN 'CCI' WHEN 1 THEN 'CI' ELSE 'Heap' END AS Ist,
        CASE WHEN x.partitioniert = 1 THEN 'ja' ELSE 'nein' END       AS Partitioniert,
        CASE WHEN x.zeilen > 0 THEN 'ja' ELSE 'nein' END              AS Hat_Daten,
        CASE
            WHEN CASE UPPER(pi.Wert) WHEN 'CCI' THEN 5 WHEN 'CI' THEN 1 WHEN 'TRUE' THEN 1 ELSE 0 END = x.typ
             AND x.partitioniert = 1                       THEN 'OK'
            WHEN x.zeilen = 0                              THEN 'LEER (wird neu angelegt)'
            ELSE 'FEHLER'
        END                                      AS Ergebnis
FROM    <ParameterDB>.dbo.<Parametertabelle> p
JOIN    <ParameterDB>.dbo.<Parametertabelle> pi
        ON pi.Verfahren = p.Verfahren AND pi.Parameter = 'FaktenClusteredIndex'
CROSS APPLY (
        SELECT  ISNULL(MAX(i.type), 0) AS typ,
                MAX(CASE WHEN ps.data_space_id IS NOT NULL THEN 1 ELSE 0 END) AS partitioniert,
                (SELECT SUM(pa.rows) FROM sys.partitions pa
                  WHERE pa.object_id = OBJECT_ID('dbo.' + QUOTENAME(p.Wert)) AND pa.index_id <= 1) AS zeilen
        FROM    sys.indexes i
        LEFT JOIN sys.partition_schemes ps ON ps.data_space_id = i.data_space_id
        WHERE   i.object_id = OBJECT_ID('dbo.' + QUOTENAME(p.Wert)) AND i.index_id <= 1
) x
WHERE   p.Parameter = 'Faktentabelle'
  AND   OBJECT_ID('dbo.' + QUOTENAME(p.Wert), 'U') IS NOT NULL
ORDER BY Ergebnis DESC, p.Verfahren;
