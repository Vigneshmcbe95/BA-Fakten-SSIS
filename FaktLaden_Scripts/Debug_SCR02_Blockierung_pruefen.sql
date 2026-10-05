-- =============================================================================
-- Debug_SCR02_Blockierung_pruefen.sql
-- Zweck: pruefen, warum SCR02 Schritt 4 (Credential neu anlegen) haengt.
-- DROP EXTERNAL TABLE wartet, solange eine alte Sitzung die ext-Tabelle nutzt.
-- =============================================================================

-- 1) Wer wartet, und auf wen?
SELECT r.session_id, r.blocking_session_id, r.status, r.wait_type,
       r.wait_time / 1000 AS wartet_sek, r.command, t.text
FROM sys.dm_exec_requests r
CROSS APPLY sys.dm_exec_sql_text(r.sql_handle) t
WHERE r.session_id <> @@SPID
ORDER BY r.wait_time DESC;

-- 2) Details zur blockierenden Sitzung (Nummer aus blocking_session_id einsetzen)
DECLARE @blocker int = 0;  -- anpassen
SELECT s.session_id, s.program_name, s.login_name, s.login_time, s.status, t.text
FROM sys.dm_exec_sessions s
JOIN sys.dm_exec_connections c ON c.session_id = s.session_id
CROSS APPLY sys.dm_exec_sql_text(c.most_recent_sql_handle) t
WHERE s.session_id = @blocker;
