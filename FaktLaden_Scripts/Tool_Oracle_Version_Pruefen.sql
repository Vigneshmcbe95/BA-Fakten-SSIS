-- =============================================================================
--  Tool   : Tool_Oracle_Version_Pruefen
--  Paket  : Fakten Laden (SSIS)
--  Zweck  : Ermittelt die Oracle-Datenbankversion ueber PolyBase, indem eine
--           externe Tabelle auf die Oracle-View SYS.PRODUCT_COMPONENT_VERSION
--           gelegt und abgefragt wird.
--  Hinweis: PRODUCT_COMPONENT_VERSION ist fuer normale Benutzer lesbar
--           (kein V$-Recht noetig).
-- =============================================================================

-- -----------------------------------------------------------------------------
-- Schritt 1: Vorhandene PolyBase-Datenquelle(n) und LOCATION-Format ermitteln
-- -----------------------------------------------------------------------------
SELECT name, location
FROM   sys.external_data_sources;

-- Eine funktionierende externe Tabelle als Muster fuer das LOCATION-Format
-- (Ihre Faktentabellen nutzen <ENV>.<SCHEMA>.<TABELLE> in GROSSBUCHSTABEN):
SELECT t.name AS ext_table, ds.name AS data_source, t.location
FROM   sys.external_tables t
JOIN   sys.external_data_sources ds ON ds.data_source_id = t.data_source_id;

-- -----------------------------------------------------------------------------
-- Schritt 2: Externe Tabelle auf die Oracle-Versions-View anlegen
--  - <YOUR_DATA_SOURCE> durch den Namen aus Schritt 1 ersetzen.
--  - LOCATION-Praefix an Ihr Muster anpassen (<ENV>.SYS.PRODUCT_COMPONENT_VERSION).
-- -----------------------------------------------------------------------------
IF OBJECT_ID('ext.oracle_version', 'U') IS NOT NULL
    DROP EXTERNAL TABLE ext.oracle_version;
GO

CREATE EXTERNAL TABLE ext.oracle_version
(
    PRODUCT  NVARCHAR(80) COLLATE Latin1_General_100_CI_AS_SC_UTF8,
    VERSION  NVARCHAR(60) COLLATE Latin1_General_100_CI_AS_SC_UTF8,
    STATUS   NVARCHAR(60) COLLATE Latin1_General_100_CI_AS_SC_UTF8
)
WITH (
    DATA_SOURCE = [<YOUR_DATA_SOURCE>],
    LOCATION    = '<ENV>.SYS.PRODUCT_COMPONENT_VERSION'
);
GO

-- -----------------------------------------------------------------------------
-- Schritt 3: Version abfragen
--  Beispielergebnis:
--    PRODUCT                                  VERSION       STATUS
--    Oracle Database 19c Enterprise Edition   19.0.0.0.0    Production
-- -----------------------------------------------------------------------------
SELECT * FROM ext.oracle_version;
GO

-- -----------------------------------------------------------------------------
-- Optional (nur Oracle 12.2+): Patch-Level ueber VERSION_FULL ermitteln.
-- Spalte VERSION_FULL der Spaltenliste oben hinzufuegen, sonst Spaltenanzahl-Fehler.
--   VERSION_FULL NVARCHAR(60) COLLATE Latin1_General_100_CI_AS_SC_UTF8
-- -----------------------------------------------------------------------------

-- -----------------------------------------------------------------------------
-- Aufraeumen
-- -----------------------------------------------------------------------------
-- DROP EXTERNAL TABLE ext.oracle_version;
-- GO
