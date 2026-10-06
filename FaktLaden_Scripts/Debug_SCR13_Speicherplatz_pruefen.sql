-- =============================================================================
-- Debug_SCR13_Speicherplatz_pruefen.sql
-- Zweck: pruefen, warum SCR13 mit "PRIMARY filegroup is full" abbricht.
-- Auf Q0003713 in der Datenbank msi_dm_fst ausfuehren.
-- =============================================================================
USE msi_dm_fst;

-- 1) Groesse, Belegung und Maximalgroesse je Datei
--    used_mb nahe max_mb -> Datei hat ihr Limit erreicht
--    max_mb = unlimited  -> vermutlich ist das Laufwerk voll
SELECT f.name, fg.name AS filegroup, f.physical_name,
       f.size / 128 AS size_mb,
       FILEPROPERTY(f.name, 'SpaceUsed') / 128 AS used_mb,
       CASE f.max_size WHEN -1 THEN 'unlimited' ELSE CAST(f.max_size / 128 AS varchar) END AS max_mb,
       f.growth, f.is_percent_growth
FROM sys.database_files f
LEFT JOIN sys.filegroups fg ON fg.data_space_id = f.data_space_id;

-- 2) Freier Platz auf den Laufwerken der Datenbankdateien
SELECT DISTINCT vs.volume_mount_point,
       vs.total_bytes / 1048576 AS gesamt_mb,
       vs.available_bytes / 1048576 AS frei_mb
FROM sys.database_files f
CROSS APPLY sys.dm_os_volume_stats(DB_ID(), f.file_id) vs;

-- 3) Uebrig gebliebene _LOADING-Tabellen aus abgebrochenen Ladevorgaengen
SELECT t.name, SUM(a.total_pages) / 128 AS mb
FROM sys.tables t
JOIN sys.partitions p ON p.object_id = t.object_id
JOIN sys.allocation_units a ON a.container_id = p.partition_id
WHERE t.name LIKE '%[_]LOADING'
GROUP BY t.name
ORDER BY mb DESC;

-- 4) Autogrow-Ereignisse aus dem Default Trace (warum ist das Wachstum um 02:18 gescheitert?)
SELECT t.StartTime, te.name, t.FileName, t.Duration / 1000 AS ms
FROM sys.traces st
CROSS APPLY sys.fn_trace_gettable(st.path, DEFAULT) t
JOIN sys.trace_events te ON te.trace_event_id = t.EventClass
WHERE st.is_default = 1
  AND t.DatabaseName = 'msi_dm_fst'
  AND t.StartTime >= '2026-10-06 01:00'
ORDER BY t.StartTime;

-- 5) Eintraege im SQL-Server-Fehlerprotokoll zur Datenbank
EXEC sys.xp_readerrorlog 0, 1, N'msi_dm_fst';

-- 6) NUR DURCH DEN DBA: Datendatei vor dem Lauf vergroessern (Beispiel +100 GB).
--    Vorher pruefen, dass auf E: genug Platz frei ist.
-- ALTER DATABASE msi_dm_fst
-- MODIFY FILE (NAME = msi_dm_fst, SIZE = 809000MB);
