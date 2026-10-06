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
