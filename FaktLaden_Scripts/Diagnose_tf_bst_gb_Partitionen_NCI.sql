/* =============================================================================
   Diagnose: vf_bst_gb / tf_bst_gb - Partitionsgrenzen und NCI-Abdeckung
   Zweck   : Root Cause fuer SCR19-Fehler
             "no identical index ... 'nci_tf_bst_gb' ..." bestaetigen, bevor
             Code geaendert wird. Prueft:
             1) Wie viele Partitionsgrenzen die Faktentabelle JETZT hat
                (Erwartung laut SCR11-Log: 595).
             2) Ob/wie viele _in_ / _out_ Staging-Tabellen ueberhaupt existieren.
             3) Von den existierenden Staging-Tabellen: wie viele haben den
                Index nci_tf_bst_gb bereits (SCR16 gelaufen) und wie viele
                nicht (SCR16 hat sie nicht erreicht).
             4) Ob die Faktentabelle noch eine grosse "Resttabelle" (Katalog-
                Partition mit sehr vielen Zeilen) hat, die auf einen
                unvollstaendigen SCR11-Lauf hindeutet.
   Ausfuehren gegen: msi_dm_bst_v3
   ============================================================================= */

SET NOCOUNT ON;

DECLARE @Faktentabelle sysname = N'tf_bst_gb';
DECLARE @NciName       sysname = N'nci_tf_bst_gb';

-- -----------------------------------------------------------------------
-- 1) Aktuelle Partitionsgrenzen der Faktentabelle: Anzahl + Zeilen je
--    Partition. Viele Partitionen mit sehr hoher Zeilenzahl im Vergleich
--    zu anderen = Hinweis auf noch nicht durchgefuehrte Splits (SCR11
--    unvollstaendig).
-- -----------------------------------------------------------------------
SELECT
    p.partition_number,
    p.rows,
    CONVERT(int, r.value)              AS Partitionsgrenze_bis_exklusive,
    CASE WHEN p.rows > 5000000 THEN '>>> AUFFAELLIG GROSS' ELSE '' END AS Hinweis
FROM sys.partitions p
JOIN sys.indexes i        ON i.object_id = p.object_id AND i.index_id = p.index_id AND i.index_id < 2
JOIN sys.tables t         ON t.object_id = p.object_id
JOIN sys.data_spaces d    ON i.data_space_id = d.data_space_id
LEFT JOIN sys.partition_schemes s    ON d.name = s.name
LEFT JOIN sys.partition_functions f  ON s.function_id = f.function_id
LEFT JOIN sys.partition_range_values r
       ON r.function_id = f.function_id
      AND r.boundary_id + f.boundary_value_on_right = p.partition_number
WHERE t.name = @Faktentabelle AND t.schema_id = SCHEMA_ID('dbo')
ORDER BY p.partition_number;

-- -----------------------------------------------------------------------
-- 2) Gesamtzahl Partitionsgrenzen (sollte laut SCR11-Log 595 sein)
-- -----------------------------------------------------------------------
SELECT
    COUNT(*) AS AnzahlPartitionsgrenzen_Aktuell,
    595      AS Erwartet_Laut_SCR11_Log
FROM sys.partition_range_values prv
JOIN sys.partition_functions pfn ON prv.function_id = pfn.function_id
JOIN sys.partition_schemes ps    ON ps.function_id = pfn.function_id
JOIN sys.data_spaces d           ON d.data_space_id = ps.data_space_id
JOIN sys.indexes i               ON i.data_space_id = d.data_space_id
JOIN sys.tables t                ON t.object_id = i.object_id
WHERE t.name = @Faktentabelle AND t.schema_id = SCHEMA_ID('dbo');

-- -----------------------------------------------------------------------
-- 3) Existierende _in_ / _out_ Staging-Tabellen fuer dieses Verfahren:
--    wie viele gibt es UEBERHAUPT (unabhaengig vom Index)?
-- -----------------------------------------------------------------------
SELECT
    COUNT(*) AS AnzahlStagingTabellen_Vorhanden,
    SUM(CASE WHEN name LIKE @Faktentabelle + '\_in\_%'  ESCAPE '\' THEN 1 ELSE 0 END) AS Davon_in,
    SUM(CASE WHEN name LIKE @Faktentabelle + '\_out\_%' ESCAPE '\' THEN 1 ELSE 0 END) AS Davon_out,
    595 * 2 AS Erwartet_Wenn_Alle_Partitionen_Geladen
FROM sys.tables
WHERE schema_id = SCHEMA_ID('dbo')
  AND (name LIKE @Faktentabelle + '\_in\_%'  ESCAPE '\'
       OR name LIKE @Faktentabelle + '\_out\_%' ESCAPE '\');

-- -----------------------------------------------------------------------
-- 4) Von den existierenden Staging-Tabellen: hat @NciName den Index oder
--    nicht? Das ist die direkte Bestaetigung/Widerlegung der SCR16-Luecke.
-- -----------------------------------------------------------------------
SELECT
    t.name AS Tabelle,
    CASE WHEN t.name LIKE @Faktentabelle + '\_in\_%'  ESCAPE '\' THEN 'in'
         WHEN t.name LIKE @Faktentabelle + '\_out\_%' ESCAPE '\' THEN 'out'
         ELSE '?' END AS Typ,
    CASE WHEN EXISTS (
        SELECT 1 FROM sys.indexes ix
        WHERE ix.object_id = t.object_id AND ix.name = @NciName
    ) THEN 'JA' ELSE 'FEHLT' END AS Hat_NCI
INTO #StagingCheck
FROM sys.tables t
WHERE t.schema_id = SCHEMA_ID('dbo')
  AND (t.name LIKE @Faktentabelle + '\_in\_%'  ESCAPE '\'
       OR t.name LIKE @Faktentabelle + '\_out\_%' ESCAPE '\');

SELECT
    Typ,
    Hat_NCI,
    COUNT(*) AS Anzahl
FROM #StagingCheck
GROUP BY Typ, Hat_NCI
ORDER BY Typ, Hat_NCI;

-- Details der betroffenen (fehlenden) Tabellen, falls weiter untersucht werden soll
SELECT * FROM #StagingCheck WHERE Hat_NCI = 'FEHLT' ORDER BY Typ, Tabelle;

DROP TABLE #StagingCheck;

-- -----------------------------------------------------------------------
-- 5) Alle NONCLUSTERED (Rowstore) Indizes der Faktentabelle - zur
--    Bestaetigung, dass nci_tf_bst_gb tatsaechlich (weiterhin) so existiert
--    wie im Fehlerbericht beschrieben, und ob es noch weitere gibt, die
--    ebenfalls repliziert werden muessten.
-- -----------------------------------------------------------------------
SELECT
    i.name AS IndexName,
    i.type_desc,
    i.is_unique
FROM sys.indexes i
WHERE i.object_id = OBJECT_ID('dbo.' + @Faktentabelle)
  AND i.index_id > 1
  AND i.type_desc = 'NONCLUSTERED';
