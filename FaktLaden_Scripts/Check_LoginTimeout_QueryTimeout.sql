-- Check: aktuell konfigurierte CONNECTION_OPTIONS (LoginTimeout/QueryTimeout)
-- fuer die externen Oracle-Datenquellen.
--
-- Fehlt LoginTimeout/QueryTimeout im connection_options-Wert, laufen sie
-- auf den Oracle-ODBC-Treiber-Standardwerten:
--   LoginTimeout (LT) = 15 Sekunden
--   QueryTimeout (QT) = 0 (unbegrenzt)
--
-- Quelle: https://learn.microsoft.com/en-us/sql/t-sql/statements/create-external-data-source-connection-options

-- Einzelne Datenquelle pruefen:
SELECT
    name,
    type_desc,
    location,
    connection_options,
    credential_id
FROM sys.external_data_sources
WHERE name = 'Oracle-istat';  -- Namen ggf. anpassen

-- Alle externen Datenquellen auf einmal:
SELECT
    name,
    location,
    connection_options
FROM sys.external_data_sources;
