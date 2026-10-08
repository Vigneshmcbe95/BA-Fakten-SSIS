-- =============================================================================
-- Speicher_Ueberwachung_Nacht.sql
-- Zweck: protokolliert alle 10 Minuten bis 08:00 am naechsten Morgen
--        - freien Platz je Laufwerk mit Datenbankdateien,
--        - Groesse aller Datenbankdateien auf E:,
--        - belegten und freien Platz in den Datendateien von msi_dm_fst.
-- Ziel: tempdb.dbo.Speicher_Nacht (kein neues ETL-Objekt; tempdb wird beim
--       naechsten Neustart des Dienstes automatisch geleert).
-- Aufruf: in SSMS auf Q0003713 ausfuehren und das Fenster offen lassen.
--         RDP-Sitzung trennen ist ok, Abmelden beendet die Abfrage.
-- Liest nur - aendert nichts an den Datenbanken.
-- =============================================================================
USE msi_dm_fst;
SET NOCOUNT ON;

IF OBJECT_ID('tempdb.dbo.Speicher_Nacht') IS NULL
    CREATE TABLE tempdb.dbo.Speicher_Nacht
    (
        Zeit       datetime2(0)  NOT NULL,
        Art        varchar(20)   NOT NULL,   -- Laufwerk | Datei | fst_Datei
        Name       nvarchar(400) NOT NULL,
        Groesse_MB bigint        NULL,
        Belegt_MB  bigint        NULL,
        Frei_MB    bigint        NULL
    );

DECLARE @Ende datetime2(0) = DATEADD(HOUR, 32, CAST(CAST(GETDATE() AS date) AS datetime2(0)));  -- morgen 08:00
DECLARE @Zeit datetime2(0);
DECLARE @Meldung nvarchar(200);

WHILE SYSDATETIME() < @Ende
BEGIN
    SET @Zeit = SYSDATETIME();

    -- 1) Freier Platz je Laufwerk
    INSERT tempdb.dbo.Speicher_Nacht (Zeit, Art, Name, Groesse_MB, Frei_MB)
    SELECT DISTINCT @Zeit, 'Laufwerk', vs.volume_mount_point,
           vs.total_bytes / 1048576, vs.available_bytes / 1048576
    FROM sys.master_files f
    CROSS APPLY sys.dm_os_volume_stats(f.database_id, f.file_id) vs;

    -- 2) Groesse aller Datenbankdateien auf E: (wer waechst?)
    INSERT tempdb.dbo.Speicher_Nacht (Zeit, Art, Name, Groesse_MB)
    SELECT @Zeit, 'Datei', DB_NAME(f.database_id) + ' | ' + f.physical_name,
           CAST(f.size AS bigint) / 128
    FROM sys.master_files f
    WHERE f.physical_name LIKE 'E:%';

    -- 3) msi_dm_fst: belegt und frei in den Datendateien
    INSERT tempdb.dbo.Speicher_Nacht (Zeit, Art, Name, Groesse_MB, Belegt_MB, Frei_MB)
    SELECT @Zeit, 'fst_Datei', f.name,
           CAST(f.size AS bigint) / 128,
           CAST(FILEPROPERTY(f.name, 'SpaceUsed') AS bigint) / 128,
           CAST(f.size - FILEPROPERTY(f.name, 'SpaceUsed') AS bigint) / 128
    FROM sys.database_files f
    WHERE f.type_desc = 'ROWS';

    SET @Meldung = CONVERT(nvarchar(20), @Zeit, 120) + N' protokolliert';
    RAISERROR(@Meldung, 0, 1) WITH NOWAIT;

    WAITFOR DELAY '00:10:00';
END;

-- =============================================================================
-- Auswertung am Morgen (in einem zweiten Fenster ausfuehren)
-- =============================================================================
-- Freier Platz je Laufwerk ueber die Nacht:
-- SELECT Zeit, Name AS laufwerk, Frei_MB / 1024 AS frei_gb
-- FROM tempdb.dbo.Speicher_Nacht WHERE Art = 'Laufwerk' ORDER BY Name, Zeit;
--
-- Welche Datei auf E: ist ueber Nacht gewachsen?
-- SELECT Name, MIN(Groesse_MB) AS start_mb, MAX(Groesse_MB) AS ende_mb,
--        MAX(Groesse_MB) - MIN(Groesse_MB) AS gewachsen_mb
-- FROM tempdb.dbo.Speicher_Nacht WHERE Art = 'Datei'
-- GROUP BY Name ORDER BY gewachsen_mb DESC;
--
-- Danach aufraeumen:
-- DROP TABLE tempdb.dbo.Speicher_Nacht;
