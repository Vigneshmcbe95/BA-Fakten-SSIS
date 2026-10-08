-- =============================================================================
-- Debug_LegacyCE_pruefen.sql
-- Zweck: prueft die drei Bedingungen aus der PolyBase-Fehlermeldung
--   "Queries that reference external tables are not supported by the legacy
--    cardinality estimation framework".
-- In der betroffenen Datenbank ausfuehren (USE anpassen). Nur lesend.
-- =============================================================================
USE msi_dm_fst;   -- betroffene Datenbank eintragen

-- 1) Legacy Cardinality Estimator per Database Scoped Configuration
--    value = 1 -> aktiv -> PolyBase-Abfragen schlagen fehl
SELECT DB_NAME() AS db, name, value, value_for_secondary
FROM sys.database_scoped_configurations
WHERE name = 'LEGACY_CARDINALITY_ESTIMATION';

-- 2) Kompatibilitaetsgrad: muss mindestens 120 sein
SELECT name AS db, compatibility_level
FROM sys.databases
WHERE name = DB_NAME();

-- 3) Trace Flag 9481 (erzwingt den alten Estimator): Status 1 = aktiv
DBCC TRACESTATUS (9481, -1);

-- =============================================================================
-- Nur falls 1) value = 1 zeigt - und nur nach Absprache mit dem Datamart-
-- Verantwortlichen (aendert die Abfrageplaene der Berichte):
-- ALTER DATABASE SCOPED CONFIGURATION SET LEGACY_CARDINALITY_ESTIMATION = OFF;
-- =============================================================================
