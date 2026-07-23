/* =============================================================================
   Diagnose: tf_bst_gb - Aktuelle Indizes + ungefaehrer Erstellungs-/
             Aenderungszeitpunkt.
   Zweck   : Pruefen, ob nci_tf_bst_gb (oder ein anderer Index) WAEHREND des
             SCR16-Laufs (2026-07-21, ca. 23:05 - 23:07 Uhr fuer die ersten
             8 Partitionen, danach keine NCI-Replikation mehr geloggt) neu
             angelegt oder geaendert wurde. SQL Server speichert kein
             natives "Index erstellt am" - STATS_DATE() aktualisiert sich
             aber bei CREATE INDEX / REBUILD und ist damit ein guter Proxy.
   Ausfuehren gegen: msi_dm_bst_v3
   ============================================================================= */

SET NOCOUNT ON;

DECLARE @Faktentabelle sysname = N'tf_bst_gb';

-- -----------------------------------------------------------------------
-- 1) Alle aktuell vorhandenen Indizes auf der Faktentabelle + wann sie
--    laut Statistik zuletzt erstellt/rebuilt wurden (Proxy fuer "wann
--    angelegt/geaendert").
-- -----------------------------------------------------------------------
SELECT
    i.name                                  AS IndexName,
    i.type_desc,
    i.is_unique,
    i.index_id,
    STATS_DATE(i.object_id, i.index_id)     AS Statistik_Zuletzt_Aktualisiert,
    p.rows                                  AS Zeilen_Partition1
FROM sys.indexes i
JOIN sys.tables t   ON t.object_id = i.object_id
LEFT JOIN sys.partitions p ON p.object_id = i.object_id AND p.index_id = i.index_id AND p.partition_number = 1
WHERE t.name = @Faktentabelle AND t.schema_id = SCHEMA_ID('dbo')
ORDER BY i.index_id;

-- -----------------------------------------------------------------------
-- 2) Objekt selbst: wann wurde die Tabelle zuletzt strukturell geaendert
--    (create_date / modify_date auf sys.objects - modify_date aktualisiert
--    sich bei manchen DDL-Operationen, ist aber kein 100% verlaesslicher
--    Index-Timestamp; nur als Zusatzindiz).
-- -----------------------------------------------------------------------
SELECT
    name AS Tabelle,
    create_date,
    modify_date
FROM sys.objects
WHERE name = @Faktentabelle AND schema_id = SCHEMA_ID('dbo');

-- -----------------------------------------------------------------------
-- 3) Falls Query Store aktiv ist: zeigt DDL-relevante Ereignisse (CREATE/
--    ALTER INDEX) im relevanten Zeitfenster des Laufs. Nur Datensaetze
--    liefern, wenn Query Store fuer diese Datenbank aktiviert ist -
--    sonst leere Ergebnismenge (kein Fehler).
-- -----------------------------------------------------------------------
IF EXISTS (SELECT 1 FROM sys.database_query_store_options WHERE actual_state <> 0)
BEGIN
    SELECT
        qsq.query_id,
        qsqt.query_sql_text,
        qsrs.first_execution_time,
        qsrs.last_execution_time
    FROM sys.query_store_query qsq
    JOIN sys.query_store_query_text qsqt ON qsq.query_text_id = qsqt.query_text_id
    JOIN sys.query_store_runtime_stats qsrs ON qsrs.plan_id IN (
        SELECT plan_id FROM sys.query_store_plan WHERE query_id = qsq.query_id
    )
    WHERE qsqt.query_sql_text LIKE '%' + @Faktentabelle + '%'
      AND (qsqt.query_sql_text LIKE '%CREATE%INDEX%' OR qsqt.query_sql_text LIKE '%DROP%INDEX%')
      AND qsrs.last_execution_time >= '2026-07-21 22:00:00'
      AND qsrs.last_execution_time <= '2026-07-22 20:00:00'
    ORDER BY qsrs.last_execution_time;
END
ELSE
BEGIN
    SELECT 'Query Store nicht aktiv fuer diese Datenbank - Punkt 3 liefert keine Daten.' AS Hinweis;
END

-- -----------------------------------------------------------------------
-- 4) Aktuelle Anzahl _in_/_out_ Staging-Tabellen mit/ohne nci_tf_bst_gb -
--    Momentaufnahme (identisch zum vorherigen Diagnoseskript, hier erneut
--    zur bequemen gemeinsamen Ausfuehrung mit obigen Punkten).
-- -----------------------------------------------------------------------
DECLARE @NciName sysname = N'nci_tf_bst_gb';
SELECT
    CASE WHEN t.name LIKE @Faktentabelle + '\_in\_%'  ESCAPE '\' THEN 'in'
         WHEN t.name LIKE @Faktentabelle + '\_out\_%' ESCAPE '\' THEN 'out'
         ELSE '?' END AS Typ,
    CASE WHEN EXISTS (
        SELECT 1 FROM sys.indexes ix WHERE ix.object_id = t.object_id AND ix.name = @NciName
    ) THEN 'JA' ELSE 'FEHLT' END AS Hat_NCI,
    COUNT(*) AS Anzahl
FROM sys.tables t
WHERE t.schema_id = SCHEMA_ID('dbo')
  AND (t.name LIKE @Faktentabelle + '\_in\_%'  ESCAPE '\'
       OR t.name LIKE @Faktentabelle + '\_out\_%' ESCAPE '\')
GROUP BY
    CASE WHEN t.name LIKE @Faktentabelle + '\_in\_%'  ESCAPE '\' THEN 'in'
         WHEN t.name LIKE @Faktentabelle + '\_out\_%' ESCAPE '\' THEN 'out'
         ELSE '?' END,
    CASE WHEN EXISTS (
        SELECT 1 FROM sys.indexes ix WHERE ix.object_id = t.object_id AND ix.name = @NciName
    ) THEN 'JA' ELSE 'FEHLT' END
ORDER BY Typ, Hat_NCI;
