-- =============================================================================
-- Debug_Speicher_Datamarts_einfach.sql
-- Zweck: Groesse, belegt und frei ("Space Available") je Datamart sowie freier
--        Platz je Laufwerk. Ohne Temp-Tabellen, nur lesend.
-- =============================================================================

-- 1) Groesse, belegt und frei je Datamart (nur Datendateien), groesster freier Platz zuerst
DECLARE @teile nvarchar(max) = N'';
SELECT @teile = @teile + CASE WHEN @teile = N'' THEN N'' ELSE N' UNION ALL ' END
     + N'SELECT N''' + REPLACE(name, '''', '''''') + N''' AS db, SUM(au.total_pages) / 128 AS belegt_mb FROM '
     + QUOTENAME(name) + N'.sys.allocation_units au'
FROM sys.databases
WHERE database_id > 4 AND state_desc = 'ONLINE';

DECLARE @sql nvarchar(max) = N'
SELECT m.db, m.groesse_mb, b.belegt_mb,
       m.groesse_mb - b.belegt_mb AS frei_mb,
       CAST(100.0 * (m.groesse_mb - b.belegt_mb) / NULLIF(m.groesse_mb, 0) AS decimal(5, 1)) AS frei_prozent
FROM (SELECT DB_NAME(database_id) AS db, SUM(CAST(size AS bigint)) / 128 AS groesse_mb
      FROM sys.master_files WHERE type_desc = ''ROWS'' GROUP BY database_id) m
JOIN (' + @teile + N') b ON b.db = m.db
ORDER BY frei_mb DESC;';

EXEC sys.sp_executesql @sql;

-- 2) Freier Platz je Laufwerk
SELECT DISTINCT vs.volume_mount_point AS laufwerk,
       vs.total_bytes / 1073741824 AS gesamt_gb,
       vs.available_bytes / 1073741824 AS frei_gb
FROM sys.master_files f
CROSS APPLY sys.dm_os_volume_stats(f.database_id, f.file_id) vs;
