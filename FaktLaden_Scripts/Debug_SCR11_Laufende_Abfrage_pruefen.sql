-- =============================================================================
-- Debug_SCR11_Laufende_Abfrage_pruefen.sql
-- Zeigt, welche Abfrage SQL Server gerade fuer ein Verfahren ausfuehrt.
-- Zweck: pruefen, ob SCR11 noch arbeitet oder wirklich haengt.
--
-- Auswertung:
--   text = SELECT TOP 1 1 ... WHERE [mon_id] = <Wert>
--       -> SCR11 prueft Partition fuer Partition. Nach 1 Minute erneut
--          ausfuehren: hat sich der Wert geaendert, laeuft es noch.
--   Gleicher Wert bleibt lange stehen -> einzelne Pruefung sehr langsam
--          (Full Scan der Oracle-View).
--   Keine Zeilen -> nichts laeuft gegen diese Tabelle; Paket/Job pruefen.
-- =============================================================================
DECLARE @verfahren nvarchar(128) = N'vf_fst_km_bd';   -- anpassen

SELECT r.session_id,
       r.status,
       r.wait_type,
       r.total_elapsed_time / 1000 AS sek,
       t.text
FROM sys.dm_exec_requests r
CROSS APPLY sys.dm_exec_sql_text(r.sql_handle) t
WHERE t.text LIKE N'%ext.%' + @verfahren + N'%'
  AND r.session_id <> @@SPID;

-- =============================================================================
-- 2) Alle laufenden Abfragen auf dem Server (wenn 1) keine Zeilen liefert)
--    Liefert auch das keine Zeilen: fehlende Berechtigung VIEW SERVER STATE
--    oder falscher Server. Paket-Sitzungen erkennt man an program_name
--    (SSIS, DTEXEC, .Net SqlClient Data Provider).
-- =============================================================================
SELECT r.session_id, s.program_name, s.login_name, r.status, r.wait_type,
       r.total_elapsed_time / 1000 AS sek, DB_NAME(r.database_id) AS db, t.text
FROM sys.dm_exec_requests r
JOIN sys.dm_exec_sessions s ON s.session_id = r.session_id
CROSS APPLY sys.dm_exec_sql_text(r.sql_handle) t
WHERE r.session_id <> @@SPID
ORDER BY r.total_elapsed_time DESC;

-- =============================================================================
-- 3) Wo ist der Lauf stehen geblieben? Status je Verfahren in der Arbeitsliste
--    PARTITIONSGRENZEN_ERSTELLT -> SCR11 fertig, Lauf ist weitergegangen
--    FEHLER                     -> Fehlermeldung steht in derselben Zeile
--    frueherer Status           -> Paket in SCR11 gestoppt oder abgestuerzt
-- =============================================================================
SELECT TOP 20 *
FROM dbo.ETL_Fakt_Arbeitsliste
ORDER BY 1 DESC;
