-- =====================================================================
-- Test Master Key / Credential: 3 Szenarien OHNE Loeschen von Data Source
-- und ext-Tabellen
-- ---------------------------------------------------------------------
-- Voraussetzungen:
--   - SQL Server 2019/2022 mit PolyBase (installiert und aktiviert:
--     EXEC sp_configure 'polybase enabled', 1; RECONFIGURE;)
--   - SQL-Server-Authentifizierung aktiv (gemischter Modus)
--   - sysadmin-Rechte, Testinstanz (NICHT Produktion)
-- Kein Oracle noetig: PolyBase zeigt auf dieselbe SQL-Server-Instanz.
--
-- Ausfuehrung: Abschnitte 0 bis 4 nacheinander ausfuehren (jeweils markieren + F5).
-- Ergebnis: Tabelle dbo.ergebnis (Spalte "passt" = Ergebnis wie erwartet).
-- Bricht eine Verbindung ab ("kill state"), neu verbinden und den Abschnitt
-- erneut ausfuehren.
-- =====================================================================


-- =====================================================================
-- Abschnitt 0: Testdatenbank, Login, Quelltabelle, Hilfsprozeduren
-- =====================================================================
USE master;
GO
IF DB_ID('PB_MK_Test') IS NOT NULL
BEGIN
    ALTER DATABASE PB_MK_Test SET SINGLE_USER WITH ROLLBACK IMMEDIATE;
    DROP DATABASE PB_MK_Test;
END;
CREATE DATABASE PB_MK_Test;
IF SUSER_ID('pb_mk_test') IS NULL
    CREATE LOGIN pb_mk_test WITH PASSWORD = 'PbTest#2026!x', CHECK_POLICY = OFF;
GO
USE PB_MK_Test;
GO
CREATE USER pb_mk_test FOR LOGIN pb_mk_test;
CREATE TABLE dbo.quelle (id int);
INSERT dbo.quelle VALUES (1), (2), (3);
GRANT SELECT ON dbo.quelle TO pb_mk_test;

CREATE TABLE dbo.ergebnis (
    nr       int IDENTITY,
    szenario nvarchar(10),
    schritt  nvarchar(200),
    erwartet nvarchar(10),
    ist      nvarchar(10),
    passt    AS CASE WHEN erwartet = ist OR erwartet = '?' THEN 'ja' ELSE 'NEIN' END,
    info     nvarchar(4000)
);
GO

-- Schritt ausfuehren und protokollieren (erwartet: 'ok', 'fehler' oder '?')
CREATE OR ALTER PROCEDURE dbo.pruefe
    @sz nvarchar(10), @schritt nvarchar(200), @erwartet nvarchar(10), @sql nvarchar(max)
AS
BEGIN
    BEGIN TRY
        EXEC sp_executesql @sql;
        INSERT dbo.ergebnis (szenario, schritt, erwartet, ist, info) VALUES (@sz, @schritt, @erwartet, 'ok', NULL);
    END TRY
    BEGIN CATCH
        INSERT dbo.ergebnis (szenario, schritt, erwartet, ist, info) VALUES (@sz, @schritt, @erwartet, 'fehler', ERROR_MESSAGE());
    END CATCH;
END;
GO

-- Objekte zaehlen (bleiben Data Source und ext-Tabelle erhalten?)
CREATE OR ALTER PROCEDURE dbo.objekte @sz nvarchar(10), @schritt nvarchar(200)
AS
BEGIN
    INSERT dbo.ergebnis (szenario, schritt, erwartet, ist, info)
    SELECT @sz, @schritt, '?', 'info',
           'ext-Tabellen=' + CAST((SELECT COUNT(*) FROM sys.external_tables) AS varchar(5)) +
           ' DataSources=' + CAST((SELECT COUNT(*) FROM sys.external_data_sources) AS varchar(5)) +
           ' Credentials=' + CAST((SELECT COUNT(*) FROM sys.database_scoped_credentials) AS varchar(5)) +
           ' MK_automatisch=' + CAST((SELECT is_master_key_encrypted_by_server FROM sys.databases WHERE name = DB_NAME()) AS varchar(1));
END;
GO

-- Kompletter Neuaufbau: Master Key (Kennwort), Credential, Data Source, ext-Tabelle.
-- @smk = 1: Master Key automatisch oeffenbar; 0: Kopplung an Service Master Key entfernen.
CREATE OR ALTER PROCEDURE dbo.aufbau @kennwort nvarchar(128), @smk bit
AS
BEGIN
    DECLARE @s nvarchar(max);
    IF OBJECT_ID('dbo.ext_quelle') IS NOT NULL EXEC (N'DROP EXTERNAL TABLE dbo.ext_quelle;');
    IF EXISTS (SELECT 1 FROM sys.external_data_sources WHERE name = 'pb_src') EXEC (N'DROP EXTERNAL DATA SOURCE pb_src;');
    IF EXISTS (SELECT 1 FROM sys.database_scoped_credentials WHERE name = 'pb_cred') EXEC (N'DROP DATABASE SCOPED CREDENTIAL pb_cred;');
    IF EXISTS (SELECT 1 FROM sys.symmetric_keys WHERE name = '##MS_DatabaseMasterKey##') EXEC (N'DROP MASTER KEY;');

    SET @s = N'CREATE MASTER KEY ENCRYPTION BY PASSWORD = ''' + @kennwort + N''';';
    EXEC (@s);
    EXEC (N'CREATE DATABASE SCOPED CREDENTIAL pb_cred WITH IDENTITY = ''pb_mk_test'', SECRET = ''PbTest#2026!x'';');
    -- Bei Zertifikatsfehlern CONNECTION_OPTIONS anpassen (z.B. Encrypt=Optional)
    EXEC (N'CREATE EXTERNAL DATA SOURCE pb_src WITH (LOCATION = ''sqlserver://localhost'', CREDENTIAL = pb_cred,
            CONNECTION_OPTIONS = ''TrustServerCertificate=Yes'');');
    EXEC (N'CREATE EXTERNAL TABLE dbo.ext_quelle (id int) WITH (LOCATION = ''PB_MK_Test.dbo.quelle'', DATA_SOURCE = pb_src);');

    IF @smk = 0
        EXEC (N'ALTER MASTER KEY DROP ENCRYPTION BY SERVICE MASTER KEY;');
    BEGIN TRY EXEC (N'CLOSE MASTER KEY;'); END TRY BEGIN CATCH END CATCH;
END;
GO


-- =====================================================================
-- Abschnitt 1: Szenario 1 - Master Key automatisch oeffenbar,
--              Kennwort in der Konfiguration ist ein anderes
-- Erwartung: alles funktioniert ohne Loeschen.
-- =====================================================================
USE PB_MK_Test;
GO
EXEC dbo.aufbau @kennwort = N'Alt#Kennwort1', @smk = 1;
EXEC dbo.objekte 'S1', 'nach Aufbau';
EXEC dbo.pruefe 'S1', 'SELECT ext-Tabelle', 'ok', N'SELECT COUNT(*) FROM dbo.ext_quelle;';
EXEC dbo.pruefe 'S1', 'OPEN MASTER KEY mit falschem Kennwort (wie DIM SCR_02)', 'fehler',
     N'OPEN MASTER KEY DECRYPTION BY PASSWORD = ''Neu#Kennwort2'';';
EXEC dbo.pruefe 'S1', 'Neues Kennwort zusaetzlich hinzufuegen (ohne Drop)', 'ok',
     N'ALTER MASTER KEY ADD ENCRYPTION BY PASSWORD = ''Neu#Kennwort2'';';
EXEC dbo.pruefe 'S1', 'SELECT ext-Tabelle danach', 'ok', N'SELECT COUNT(*) FROM dbo.ext_quelle;';
EXEC dbo.objekte 'S1', 'am Ende';
SELECT * FROM dbo.ergebnis WHERE szenario = 'S1' ORDER BY nr;
GO


-- =====================================================================
-- Abschnitt 2: Szenario 2 - Master Key NICHT automatisch oeffenbar,
--              altes Kennwort bekannt
-- Erwartung: mit altem Kennwort oeffnen + koppeln, ohne Loeschen.
-- =====================================================================
USE PB_MK_Test;
GO
EXEC dbo.aufbau @kennwort = N'Alt#Kennwort1', @smk = 0;
EXEC dbo.objekte 'S2', 'nach Aufbau (nicht gekoppelt)';
EXEC dbo.pruefe 'S2', 'SELECT ext-Tabelle ohne geoeffneten Key', 'fehler', N'SELECT COUNT(*) FROM dbo.ext_quelle;';
EXEC dbo.pruefe 'S2', 'Mit altem Kennwort oeffnen + an Service Master Key koppeln', 'ok',
     N'OPEN MASTER KEY DECRYPTION BY PASSWORD = ''Alt#Kennwort1'';
       ALTER MASTER KEY ADD ENCRYPTION BY SERVICE MASTER KEY;
       CLOSE MASTER KEY;';
EXEC dbo.pruefe 'S2', 'SELECT ext-Tabelle danach', 'ok', N'SELECT COUNT(*) FROM dbo.ext_quelle;';
EXEC dbo.objekte 'S2', 'am Ende';
SELECT * FROM dbo.ergebnis WHERE szenario = 'S2' ORDER BY nr;
GO


-- =====================================================================
-- Abschnitt 3: Szenario 3 - Master Key NICHT oeffenbar, Kennwort unbekannt
-- Frage: Reicht FORCE REGENERATE + ALTER CREDENTIAL, ohne Data Source und
--        ext-Tabelle zu loeschen?  -> Ergebnis "?" = genau das wird getestet
-- =====================================================================
USE PB_MK_Test;
GO
EXEC dbo.aufbau @kennwort = N'Alt#Kennwort1', @smk = 0;
EXEC dbo.objekte 'S3', 'nach Aufbau (nicht gekoppelt)';
EXEC dbo.pruefe 'S3', 'OPEN MASTER KEY mit neuem (falschem) Kennwort', 'fehler',
     N'OPEN MASTER KEY DECRYPTION BY PASSWORD = ''Neu#Kennwort2'';';
EXEC dbo.pruefe 'S3', 'FORCE REGENERATE mit neuem Kennwort + koppeln', '?',
     N'ALTER MASTER KEY FORCE REGENERATE WITH ENCRYPTION BY PASSWORD = ''Neu#Kennwort2'';
       ALTER MASTER KEY ADD ENCRYPTION BY SERVICE MASTER KEY;';
EXEC dbo.objekte 'S3', 'nach FORCE REGENERATE';
EXEC dbo.pruefe 'S3', 'SELECT ext-Tabelle vor ALTER CREDENTIAL', '?', N'SELECT COUNT(*) FROM dbo.ext_quelle;';
EXEC dbo.pruefe 'S3', 'ALTER CREDENTIAL (Kennwort neu setzen)', '?',
     N'ALTER DATABASE SCOPED CREDENTIAL pb_cred WITH IDENTITY = ''pb_mk_test'', SECRET = ''PbTest#2026!x'';';
EXEC dbo.pruefe 'S3', 'SELECT ext-Tabelle nach ALTER CREDENTIAL (KERNFRAGE)', '?', N'SELECT COUNT(*) FROM dbo.ext_quelle;';
EXEC dbo.objekte 'S3', 'am Ende';
SELECT * FROM dbo.ergebnis WHERE szenario = 'S3' ORDER BY nr;
GO


-- =====================================================================
-- Abschnitt 4: Gesamtergebnis
-- =====================================================================
USE PB_MK_Test;
GO
SELECT szenario, nr, schritt, erwartet, ist, passt, info FROM dbo.ergebnis ORDER BY nr;
GO


-- =====================================================================
-- Aufraeumen (nach dem Test)
-- =====================================================================
-- USE master;
-- ALTER DATABASE PB_MK_Test SET SINGLE_USER WITH ROLLBACK IMMEDIATE;
-- DROP DATABASE PB_MK_Test;
-- DROP LOGIN pb_mk_test;
