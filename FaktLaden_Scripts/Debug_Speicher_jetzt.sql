-- =============================================================================
-- Debug_Speicher_jetzt.sql
-- Zweck: aktueller Speicherplatz auf Q0003713 - Laufwerke und Datenbankdateien
--        aller Datenbanken dieser Instanz. Nur lesend.
-- =============================================================================

-- 1) Freier Platz je Laufwerk, auf dem Datenbankdateien liegen
SELECT DISTINCT vs.volume_mount_point AS laufwerk,
       vs.total_bytes / 1073741824 AS gesamt_gb,
       vs.available_bytes / 1073741824 AS frei_gb,
       CAST(100.0 * vs.available_bytes / vs.total_bytes AS decimal(5, 1)) AS frei_prozent
FROM sys.master_files f
CROSS APPLY sys.dm_os_volume_stats(f.database_id, f.file_id) vs
ORDER BY laufwerk;

-- 2) Alle Datenbankdateien dieser Instanz: Groesse, Wachstum, Maximalgroesse
SELECT DB_NAME(f.database_id) AS db, f.name, f.type_desc, f.physical_name,
       CAST(f.size / 128 AS bigint) AS groesse_mb,
       CASE WHEN f.is_percent_growth = 1 THEN CAST(f.growth AS varchar(10)) + ' %'
            ELSE CAST(f.growth / 128 AS varchar(10)) + ' MB' END AS wachstum,
       CASE f.max_size WHEN -1 THEN 'unbegrenzt' WHEN 0 THEN 'kein Wachstum'
            ELSE CAST(CAST(f.max_size AS bigint) / 128 AS varchar(20)) + ' MB' END AS max_groesse
FROM sys.master_files f
ORDER BY f.size DESC;

-- 3) msi_dm_fst: belegter und freier Platz in den Datendateien
USE msi_dm_fst;
SELECT f.name, fg.name AS dateigruppe,
       f.size / 128 AS groesse_mb,
       FILEPROPERTY(f.name, 'SpaceUsed') / 128 AS belegt_mb,
       (f.size - FILEPROPERTY(f.name, 'SpaceUsed')) / 128 AS frei_in_datei_mb
FROM sys.database_files f
LEFT JOIN sys.filegroups fg ON fg.data_space_id = f.data_space_id;
