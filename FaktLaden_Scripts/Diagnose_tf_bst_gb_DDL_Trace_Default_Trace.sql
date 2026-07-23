/* =============================================================================
   Diagnose: tf_bst_gb - Wer/wann wurde nci_tf_bst_gb angelegt/geaendert?
   Zweck   : Ueber den SQL Server Default Trace (falls aktiv und Datei noch
             nicht ueberschrieben/rotiert) nach CREATE/ALTER/DROP INDEX
             Ereignissen auf tf_bst_gb im relevanten Zeitfenster suchen -
             liefert im Erfolgsfall Login, Hostname, Anwendung und exakten
             Zeitstempel. Zusaetzlich ein Extended-Events-Setup-Skript
             (auskommentiert) fuer den Fall, dass der Default Trace nichts
             mehr enthaelt (Standard-Aufbewahrung ist begrenzt/rotierend).
   Ausfuehren gegen: msi_dm_bst_v3 (Teil 1+2) bzw. Server-Ebene (Teil 3, XE)
   ============================================================================= */

SET NOCOUNT ON;

DECLARE @Faktentabelle sysname = N'tf_bst_gb';
DECLARE @VonZeit datetime = '2026-07-21 20:00:00';
DECLARE @BisZeit datetime = '2026-07-22 02:00:00';

-- -----------------------------------------------------------------------
-- 1) Ist der Default Trace ueberhaupt aktiv?
-- -----------------------------------------------------------------------
SELECT
    name,
    CAST(value AS int)     AS Konfigurierter_Wert,
    CAST(value_in_use AS int) AS Aktuell_Aktiv
FROM sys.configurations
WHERE name = 'default trace enabled';

-- -----------------------------------------------------------------------
-- 2) Default-Trace-Dateien einlesen und nach DDL-Ereignissen auf
--    tf_bst_gb im relevanten Zeitfenster filtern. Object-Ereignisklassen:
--    46 = Object:Created, 47 = Object:Deleted, 164 = Object:Altered
--    (deckt CREATE INDEX / DROP INDEX / ALTER INDEX ab).
--    fn_trace_gettable liest automatisch alle rotierten Trace-Dateien mit.
-- -----------------------------------------------------------------------
DECLARE @TracePfad nvarchar(260) =
    (SELECT CAST(value AS nvarchar(260)) FROM sys.fn_trace_getinfo(0) WHERE property = 2);

IF @TracePfad IS NOT NULL
BEGIN
    SELECT
        te.name                         AS EreignisTyp,
        t.StartTime,
        t.LoginName,
        t.HostName,
        t.ApplicationName,
        t.ObjectName,
        t.DatabaseName,
        t.TextData
    FROM sys.fn_trace_gettable(@TracePfad, DEFAULT) t
    JOIN sys.trace_events te ON te.trace_event_id = t.EventClass
    WHERE t.DatabaseName = DB_NAME()
      AND (t.ObjectName = @Faktentabelle OR t.ObjectName LIKE @Faktentabelle + '\_%' ESCAPE '\')
      AND t.EventClass IN (46, 47, 164)
      AND t.StartTime BETWEEN @VonZeit AND @BisZeit
    ORDER BY t.StartTime;
END
ELSE
BEGIN
    SELECT 'Default Trace nicht auffindbar/deaktiviert - Punkt 2 liefert keine Daten. Siehe Punkt 1.' AS Hinweis;
END

-- -----------------------------------------------------------------------
-- 3) FALLS Default Trace nichts (mehr) enthaelt: Extended-Events-Session
--    einrichten, damit ein KUENFTIGES Auftreten sofort mit Login/Zeitpunkt
--    erfasst wird. Einmalig auf Server-Ebene ausfuehren (sysadmin/
--    ALTER ANY EVENT SESSION erforderlich). Session bleibt inaktiv bis
--    manuell gestartet und muss NICHT dauerhaft laufen - nur waehrend des
--    naechsten Laufs fuer msi_dm_bst_v3 aktivieren.
-- -----------------------------------------------------------------------
/*
IF EXISTS (SELECT 1 FROM sys.server_event_sessions WHERE name = 'Diagnose_tf_bst_gb_DDL')
    DROP EVENT SESSION [Diagnose_tf_bst_gb_DDL] ON SERVER;

CREATE EVENT SESSION [Diagnose_tf_bst_gb_DDL] ON SERVER
ADD EVENT sqlserver.object_altered (
    WHERE (sqlserver.database_name = N'msi_dm_bst_v3'))
ADD EVENT sqlserver.object_created (
    WHERE (sqlserver.database_name = N'msi_dm_bst_v3'))
ADD EVENT sqlserver.object_deleted (
    WHERE (sqlserver.database_name = N'msi_dm_bst_v3'))
ADD TARGET package0.event_file (
    SET filename = N'Diagnose_tf_bst_gb_DDL.xel', max_file_size = 50)
WITH (MAX_MEMORY = 4096 KB, EVENT_RETENTION_MODE = ALLOW_SINGLE_EVENT_LOSS,
      STARTUP_STATE = OFF);

-- Vor dem naechsten Lauf manuell starten:
-- ALTER EVENT SESSION [Diagnose_tf_bst_gb_DDL] ON SERVER STATE = START;

-- Nach dem Lauf auslesen:
-- SELECT
--     event_data.value('(event/@name)[1]','varchar(100)')            AS EreignisTyp,
--     event_data.value('(event/@timestamp)[1]','datetime2')          AS Zeitpunkt,
--     event_data.value('(event/data[@name="object_name"]/value)[1]','varchar(200)') AS ObjektName,
--     event_data.value('(event/action[@name="sql_text"]/value)[1]','nvarchar(max)') AS SqlText,
--     event_data.value('(event/action[@name="username"]/value)[1]','varchar(200)')  AS Benutzer
-- FROM (
--     SELECT CAST(event_data AS xml) AS event_data
--     FROM sys.fn_xe_file_target_read_file('Diagnose_tf_bst_gb_DDL*.xel', NULL, NULL, NULL)
-- ) AS x
-- WHERE event_data.value('(event/data[@name="object_name"]/value)[1]','varchar(200)') LIKE 'tf_bst_gb%'
-- ORDER BY Zeitpunkt;
*/
