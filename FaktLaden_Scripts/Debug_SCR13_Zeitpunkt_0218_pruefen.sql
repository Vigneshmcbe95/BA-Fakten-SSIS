-- =============================================================================
-- Debug_SCR13_Zeitpunkt_0218_pruefen.sql
-- Zweck: herausfinden, was am 06.10.2026 gegen 02:18 auf Q0003713 passiert ist
--        (SCR13: "PRIMARY filegroup is full" in msi_dm_fst).
-- Ergaenzt Debug_SCR13_Speicherplatz_pruefen.sql, falls der Default Trace nicht laeuft.
-- =============================================================================

-- 0) Laeuft der Default Trace ueberhaupt? Keine Zeile mit is_default = 1 -> nein.
SELECT id, status, is_default, path, start_time, last_event_time
FROM sys.traces;

-- 1) SQL-Server-Fehlerprotokoll im Zeitfenster (alle Eintraege, nicht nur msi_dm_fst).
--    Suchen nach: "not enough space on the disk", OS-Fehler 112,
--    "Autogrow ... cancelled or timed out".
--    Wurde das Protokoll inzwischen gewechselt: erste 0 durch 1 ersetzen.
EXEC sys.xp_readerrorlog 0, 1, NULL, NULL, '2026-10-06 01:30', '2026-10-06 04:00', 'asc';

-- 2) Fehler aus der system_health-Session (laeuft standardmaessig).
--    Zeiten sind UTC: 02:18 Ortszeit = 00:18 UTC.
SELECT x.value('(@timestamp)[1]', 'datetime2') AS zeit_utc,
       x.value('(data[@name="error_number"]/value)[1]', 'int') AS fehler,
       x.value('(data[@name="message"]/value)[1]', 'nvarchar(max)') AS meldung
FROM (SELECT CAST(event_data AS xml) AS d
      FROM sys.fn_xe_file_target_read_file('system_health*.xel', NULL, NULL, NULL)
      WHERE object_name = 'error_reported') e
CROSS APPLY e.d.nodes('/event') n(x)
WHERE x.value('(@timestamp)[1]', 'datetime2') BETWEEN '2026-10-05 23:00' AND '2026-10-06 03:00'
ORDER BY zeit_utc;

-- 3a) Alle Datenbankdateien auf E: (wer belegt den Platz noch?)
SELECT DB_NAME(database_id) AS db, name, physical_name, size / 128 AS size_mb
FROM sys.master_files
WHERE physical_name LIKE 'E:%'
ORDER BY size DESC;

-- 3b) Backups in der Nacht. Backup-Dateien auf E:, die spaeter geloescht wurden,
--     wuerden erklaeren, warum E: um 02:18 voll war und jetzt ca. 174 GB frei hat.
SELECT bs.database_name, bs.backup_start_date, bs.backup_finish_date,
       bs.backup_size / 1048576 AS mb, bmf.physical_device_name
FROM msdb.dbo.backupset bs
JOIN msdb.dbo.backupmediafamily bmf ON bmf.media_set_id = bs.media_set_id
WHERE bs.backup_start_date BETWEEN '2026-10-05 22:00' AND '2026-10-06 06:00'
ORDER BY bs.backup_start_date;

-- 4) Ausserhalb von SQL Server (per RDP auf Q0003713):
--    Ereignisanzeige -> System: Ereignis 2013 "E: disk is at or near capacity" um 02:18.
--    Ereignisanzeige -> Anwendung: MSSQL-Eintraege im selben Zeitfenster.
-- 5) Plattenueberwachung des Infrastruktur-Teams: Verlauf des freien Platzes auf E:.
