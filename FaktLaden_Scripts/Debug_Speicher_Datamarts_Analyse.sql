-- =============================================================================
-- Debug_Speicher_Datamarts_Analyse.sql
-- Zweck: Speicheranalyse fuer ALLE Datamarts dieser Instanz.
--   1) Laufwerke: gesamt / frei
--   2) Je Datenbank und Datei: Groesse, belegt, frei ("Space Available")
--   3) Je Datenbank: die 10 groessten Tabellen
--   4) Je Datenbank: Reste von Ladelaeufen (_in_, _out_, _LOADING)
--   5) OPTIONAL: freier Platz am Dateiende fuer EINE Datenbank
--      (nur dieser Teil ist mit TRUNCATEONLY sofort freigebbar)
-- Nur lesend - aendert nichts. Teile 1 bis 4 laufen schnell.
-- =============================================================================
SET NOCOUNT ON;

-- -----------------------------------------------------------------------------
-- 1) Laufwerke
-- -----------------------------------------------------------------------------
SELECT DISTINCT vs.volume_mount_point AS laufwerk,
       vs.total_bytes / 1073741824 AS gesamt_gb,
       vs.available_bytes / 1073741824 AS frei_gb,
       CAST(100.0 * vs.available_bytes / vs.total_bytes AS decimal(5, 1)) AS frei_prozent
FROM sys.master_files f
CROSS APPLY sys.dm_os_volume_stats(f.database_id, f.file_id) vs
ORDER BY laufwerk;

-- Sammeltabellen fuer die Teile 2 bis 4
IF OBJECT_ID('tempdb..#Dateien') IS NOT NULL DROP TABLE #Dateien;
IF OBJECT_ID('tempdb..#Tabellen') IS NOT NULL DROP TABLE #Tabellen;
IF OBJECT_ID('tempdb..#Reste') IS NOT NULL DROP TABLE #Reste;

CREATE TABLE #Dateien (db sysname, database_id int, file_id int, datei sysname, typ nvarchar(60),
                       dateigruppe sysname NULL, pfad nvarchar(520),
                       groesse_mb bigint, belegt_mb bigint, frei_mb bigint,
                       wachstum nvarchar(30), max_groesse nvarchar(30));
CREATE TABLE #Tabellen (db sysname, tabelle nvarchar(300), reserviert_mb bigint, zeilen bigint, rang int);
CREATE TABLE #Reste (db sysname, tabelle nvarchar(300), art varchar(10), reserviert_mb bigint, zeilen bigint);

-- Je Benutzerdatenbank (online) die Teile 2 bis 4 einsammeln
DECLARE @db sysname, @sql nvarchar(max);
DECLARE c CURSOR LOCAL FAST_FORWARD FOR
    SELECT name FROM sys.databases
    WHERE database_id > 4 AND state_desc = 'ONLINE' AND HAS_DBACCESS(name) = 1
    ORDER BY name;
OPEN c;
FETCH NEXT FROM c INTO @db;
WHILE @@FETCH_STATUS = 0
BEGIN
    SET @sql = N'USE ' + QUOTENAME(@db) + N';

    INSERT #Dateien
    SELECT DB_NAME(), DB_ID(), f.file_id, f.name, f.type_desc, fg.name, f.physical_name,
           CAST(f.size AS bigint) / 128,
           CAST(FILEPROPERTY(f.name, ''SpaceUsed'') AS bigint) / 128,
           CAST(f.size - FILEPROPERTY(f.name, ''SpaceUsed'') AS bigint) / 128,
           CASE WHEN f.is_percent_growth = 1 THEN CAST(f.growth AS varchar(10)) + '' %''
                ELSE CAST(f.growth / 128 AS varchar(10)) + '' MB'' END,
           CASE f.max_size WHEN -1 THEN ''unbegrenzt'' WHEN 0 THEN ''kein Wachstum''
                ELSE CAST(CAST(f.max_size AS bigint) / 128 AS varchar(20)) + '' MB'' END
    FROM sys.database_files f
    LEFT JOIN sys.filegroups fg ON fg.data_space_id = f.data_space_id;

    WITH t AS (
        SELECT s.name + ''.'' + o.name AS tabelle,
               SUM(ps.reserved_page_count) / 128 AS reserviert_mb,
               SUM(CASE WHEN ps.index_id IN (0, 1) THEN ps.row_count ELSE 0 END) AS zeilen
        FROM sys.dm_db_partition_stats ps
        JOIN sys.objects o ON o.object_id = ps.object_id AND o.type = ''U''
        JOIN sys.schemas s ON s.schema_id = o.schema_id
        GROUP BY s.name, o.name)
    INSERT #Tabellen
    SELECT DB_NAME(), tabelle, reserviert_mb, zeilen, rang
    FROM (SELECT *, ROW_NUMBER() OVER (ORDER BY reserviert_mb DESC) AS rang FROM t) x
    WHERE rang <= 10;

    INSERT #Reste
    SELECT DB_NAME(), s.name + ''.'' + o.name,
           CASE WHEN o.name LIKE ''%[_]LOADING'' THEN ''_LOADING''
                WHEN o.name LIKE ''%[_]out[_]%'' THEN ''_out_''
                ELSE ''_in_'' END,
           SUM(ps.reserved_page_count) / 128,
           SUM(CASE WHEN ps.index_id IN (0, 1) THEN ps.row_count ELSE 0 END)
    FROM sys.dm_db_partition_stats ps
    JOIN sys.objects o ON o.object_id = ps.object_id AND o.type = ''U''
    JOIN sys.schemas s ON s.schema_id = o.schema_id
    WHERE o.name LIKE ''%[_]in[_]%'' OR o.name LIKE ''%[_]out[_]%'' OR o.name LIKE ''%[_]LOADING''
    GROUP BY s.name, o.name;';

    BEGIN TRY
        EXEC sys.sp_executesql @sql;
    END TRY
    BEGIN CATCH
        PRINT N'Uebersprungen: ' + @db + N' - ' + ERROR_MESSAGE();
    END CATCH;

    FETCH NEXT FROM c INTO @db;
END;
CLOSE c;
DEALLOCATE c;

-- -----------------------------------------------------------------------------
-- 2) Je Datenbank und Datei: Groesse, belegt, frei (groesster freier Platz zuerst)
-- -----------------------------------------------------------------------------
SELECT d.db, d.datei, d.typ, d.dateigruppe, vs.volume_mount_point AS laufwerk,
       d.groesse_mb, d.belegt_mb, d.frei_mb,
       CAST(100.0 * d.frei_mb / NULLIF(d.groesse_mb, 0) AS decimal(5, 1)) AS frei_prozent,
       d.wachstum, d.max_groesse, d.pfad
FROM #Dateien d
CROSS APPLY sys.dm_os_volume_stats(d.database_id, d.file_id) vs
ORDER BY d.frei_mb DESC;

-- Zusammenfassung je Datenbank (nur Datendateien)
SELECT db,
       SUM(groesse_mb) AS groesse_mb,
       SUM(belegt_mb) AS belegt_mb,
       SUM(frei_mb) AS frei_mb,
       CAST(100.0 * SUM(frei_mb) / NULLIF(SUM(groesse_mb), 0) AS decimal(5, 1)) AS frei_prozent
FROM #Dateien
WHERE typ = 'ROWS'
GROUP BY db
ORDER BY frei_mb DESC;

-- -----------------------------------------------------------------------------
-- 3) Je Datenbank: die 10 groessten Tabellen
-- -----------------------------------------------------------------------------
SELECT db, rang, tabelle, reserviert_mb, zeilen
FROM #Tabellen
ORDER BY db, rang;

-- -----------------------------------------------------------------------------
-- 4) Reste von Ladelaeufen. ACHTUNG: eine _out_-Tabelle MIT Zeilen kann die
--    einzige Kopie alter Daten sein (fehlgeschlagener Rueckschalt-Versuch in
--    SCR19) - vor jedem Loeschen pruefen!
-- -----------------------------------------------------------------------------
SELECT db, art, COUNT(*) AS anzahl, SUM(reserviert_mb) AS reserviert_mb, SUM(zeilen) AS zeilen
FROM #Reste
GROUP BY db, art
ORDER BY reserviert_mb DESC;

SELECT db, tabelle, art, reserviert_mb, zeilen
FROM #Reste
ORDER BY reserviert_mb DESC;

-- -----------------------------------------------------------------------------
-- 5) OPTIONAL: freier Platz am Dateiende fuer EINE Datenbank.
--    Nur dieser Teil wird mit DBCC SHRINKFILE (..., TRUNCATEONLY) sofort und
--    ohne Datenverschiebung freigegeben. Der Rest liegt "in der Mitte".
--    Liest die Allokationsseiten (sys.dm_db_database_page_allocations, nicht
--    dokumentiert). Bei grossen Datenbanken mehrere Minuten - ausserhalb der
--    Ladezeit ausfuehren. Zum Ausfuehren den Block markieren und starten.
-- -----------------------------------------------------------------------------
/*
USE msi_dm_fst;   -- Datenbank anpassen
SELECT f.name AS datei,
       CAST(f.size AS bigint) / 128 AS groesse_mb,
       CAST(FILEPROPERTY(f.name, 'SpaceUsed') AS bigint) / 128 AS belegt_mb,
       (CAST(f.size AS bigint) - ISNULL(a.hoechste_seite, 0) - 1) / 128 AS frei_am_ende_mb,
       (CAST(f.size AS bigint) - CAST(FILEPROPERTY(f.name, 'SpaceUsed') AS bigint)
          - (CAST(f.size AS bigint) - ISNULL(a.hoechste_seite, 0) - 1)) / 128 AS frei_in_der_mitte_mb
FROM sys.database_files f
OUTER APPLY (SELECT MAX(pa.allocated_page_page_id) AS hoechste_seite
             FROM sys.dm_db_database_page_allocations(DB_ID(), NULL, NULL, NULL, 'LIMITED') pa
             WHERE pa.allocated_page_file_id = f.file_id AND pa.is_allocated = 1) a
WHERE f.type_desc = 'ROWS';
*/
