-- =====================================================================
-- Einmalige Bereinigung der alten gemeinsamen PolyBase-Objekte
-- (vor der Trennung in _FAKT / _DIM)
--
-- Nur ausfuehren, wenn KEIN DIM- und KEIN Fakten-Lauf auf der Datenbank
-- aktiv ist. Die neuen Objekte (Oracle-<Umgebung>-FAKT / -DIM) legen die
-- Pakete beim naechsten Lauf selbst an.
--
-- Ablauf: Werte anpassen, mit @ausfuehren = 0 pruefen, dann mit 1 ausfuehren.
-- =====================================================================
USE [msi_dm_zwg];   -- Datenbank anpassen
GO
SET NOCOUNT ON;

DECLARE @alteDataSource sysname = N'Oracle-istat';            -- alter Name (ohne -FAKT/-DIM)
DECLARE @altesCredential sysname = N'Q0003713_BI_DM_EXPORT';  -- alter Name (<Server>_<Benutzer>)
DECLARE @ausfuehren bit = 0;                                  -- 0 = nur anzeigen, 1 = ausfuehren

-- Was geloescht wird
SELECT 'EXTERNAL TABLE' AS typ, s.name + '.' + t.name AS objekt
FROM sys.external_tables t
JOIN sys.schemas s ON s.schema_id = t.schema_id
JOIN sys.external_data_sources ds ON ds.data_source_id = t.data_source_id
WHERE ds.name = @alteDataSource
UNION ALL SELECT 'EXTERNAL DATA SOURCE', name FROM sys.external_data_sources WHERE name = @alteDataSource
UNION ALL SELECT 'CREDENTIAL', name FROM sys.database_scoped_credentials WHERE name = @altesCredential
UNION ALL SELECT 'TABLE (alte dbo-Kopie)', 'dbo.vm_ddl_sql_server' WHERE OBJECT_ID('dbo.vm_ddl_sql_server', 'U') IS NOT NULL;

-- Weitere Data Sources, die das alte Credential nutzen (werden NICHT geloescht)
SELECT 'nutzt altes Credential - bleibt' AS hinweis, ds.name
FROM sys.external_data_sources ds
JOIN sys.database_scoped_credentials c ON c.credential_id = ds.credential_id
WHERE c.name = @altesCredential AND ds.name <> @alteDataSource;

IF @ausfuehren = 0
BEGIN
    PRINT 'Nur Anzeige. Zum Ausfuehren @ausfuehren = 1 setzen.';
    RETURN;
END;

DECLARE @sql nvarchar(max) = N'';
SELECT @sql += N'DROP EXTERNAL TABLE ' + QUOTENAME(s.name) + N'.' + QUOTENAME(t.name) + N';'
FROM sys.external_tables t
JOIN sys.schemas s ON s.schema_id = t.schema_id
JOIN sys.external_data_sources ds ON ds.data_source_id = t.data_source_id
WHERE ds.name = @alteDataSource;
EXEC sp_executesql @sql;

IF EXISTS (SELECT 1 FROM sys.external_data_sources WHERE name = @alteDataSource)
BEGIN
    SET @sql = N'DROP EXTERNAL DATA SOURCE ' + QUOTENAME(@alteDataSource) + N';';
    EXEC sp_executesql @sql;
END;

IF EXISTS (SELECT 1 FROM sys.database_scoped_credentials WHERE name = @altesCredential)
   AND NOT EXISTS (SELECT 1 FROM sys.external_data_sources ds
                   JOIN sys.database_scoped_credentials c ON c.credential_id = ds.credential_id
                   WHERE c.name = @altesCredential)
BEGIN
    SET @sql = N'DROP DATABASE SCOPED CREDENTIAL ' + QUOTENAME(@altesCredential) + N';';
    EXEC sp_executesql @sql;
END;

IF OBJECT_ID('dbo.vm_ddl_sql_server', 'U') IS NOT NULL
    DROP TABLE dbo.vm_ddl_sql_server;

PRINT 'Alte PolyBase-Objekte bereinigt.';
GO
