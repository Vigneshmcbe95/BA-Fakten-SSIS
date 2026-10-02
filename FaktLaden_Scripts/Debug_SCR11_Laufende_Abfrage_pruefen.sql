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
