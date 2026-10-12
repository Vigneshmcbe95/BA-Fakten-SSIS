-- =============================================================================
-- Debug_Lauf_Stand_pruefen.sql
-- Zweck: Stand des letzten Faktenlaufs - wie viele Tabellen fertig sind.
-- Nur lesend.
-- =============================================================================
USE msi_dm_fst;

-- 1) Anzahl Tabellen je Status
SELECT Status, COUNT(*) AS Tabellen
FROM dbo.ETL_Fakt_Arbeitsliste
GROUP BY Status
ORDER BY Status;

-- 2) Jede Tabelle mit Status und letztem Schritt
SELECT Verfahren, Status, LetzterSchritt, Versuche, AktualisiertAm, Fehlermeldung
FROM dbo.ETL_Fakt_Arbeitsliste
ORDER BY Status, Verfahren;

-- 3) Was gerade in der Datenbank liegt:
--    _in_ = Monat fertig geladen (wartet auf Partitionstausch)
--    _LOADING = Monat abgebrochen / laeuft noch
SELECT LEFT(name, CHARINDEX('_in_', name) - 1) AS Tabelle,
       SUM(CASE WHEN name LIKE '%[_]LOADING' THEN 0 ELSE 1 END) AS Monate_fertig,
       SUM(CASE WHEN name LIKE '%[_]LOADING' THEN 1 ELSE 0 END) AS Monate_LOADING
FROM sys.tables
WHERE name LIKE 'vf[_]%[_]in[_]%'
GROUP BY LEFT(name, CHARINDEX('_in_', name) - 1)
ORDER BY Tabelle;
