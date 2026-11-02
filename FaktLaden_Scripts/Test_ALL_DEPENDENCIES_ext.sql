/* ============================================================
   Test: Oracle ALL_DEPENDENCIES per PolyBase lesen
   Ziel: Basistabellen einer View und deren Partitionen finden
   In der Faktendatenbank ausführen (Schema ext vorhanden).
   ============================================================ */

-- 1) Datenquelle und Location von ext.v_partition_info ermitteln
SELECT t.name        AS ext_tabelle,
       t.location    AS location,
       d.name        AS data_source
FROM   sys.external_tables t
JOIN   sys.external_data_sources d ON d.data_source_id = t.data_source_id
WHERE  t.name = 'v_partition_info';

-- Werte aus 1) hier eintragen (Umgebung = erster Teil der Location)
-- DATA_SOURCE: [Oracle-ESTAT]   LOCATION: 'ESTAT.SYS.ALL_DEPENDENCIES'


-- 2) Externe Tabelle anlegen
IF OBJECT_ID('ext.all_dependencies') IS NOT NULL
    DROP EXTERNAL TABLE ext.all_dependencies;

CREATE EXTERNAL TABLE ext.all_dependencies
(
    OWNER                VARCHAR(512) COLLATE Latin1_General_100_CS_AS_SC_UTF8 NULL,
    NAME                 VARCHAR(512) COLLATE Latin1_General_100_CS_AS_SC_UTF8 NULL,
    TYPE                 VARCHAR(80)  COLLATE Latin1_General_100_CS_AS_SC_UTF8 NULL,
    REFERENCED_OWNER     VARCHAR(512) COLLATE Latin1_General_100_CS_AS_SC_UTF8 NULL,
    REFERENCED_NAME      VARCHAR(512) COLLATE Latin1_General_100_CS_AS_SC_UTF8 NULL,
    REFERENCED_TYPE      VARCHAR(80)  COLLATE Latin1_General_100_CS_AS_SC_UTF8 NULL,
    REFERENCED_LINK_NAME VARCHAR(512) COLLATE Latin1_General_100_CS_AS_SC_UTF8 NULL,
    DEPENDENCY_TYPE      VARCHAR(16)  COLLATE Latin1_General_100_CS_AS_SC_UTF8 NULL
)
WITH (
    DATA_SOURCE = [Oracle-ESTAT],
    LOCATION    = N'ESTAT.SYS.ALL_DEPENDENCIES'
);


-- 3) Basistabellen/-views einer View
SELECT OWNER, NAME, TYPE, REFERENCED_OWNER, REFERENCED_NAME, REFERENCED_TYPE
FROM   ext.all_dependencies
WHERE  NAME = 'VF_LSTP_BG_BS'
  AND  REFERENCED_TYPE IN ('TABLE', 'VIEW');


-- 4) Partitionen der Basistabellen (aus V_PARTITION_INFO)
SELECT d.NAME            AS view_name,
       d.REFERENCED_NAME AS basistabelle,
       p.PARTITION_NAME,
       p.HIGH_VALUE
FROM   ext.all_dependencies d
JOIN   ext.v_partition_info p
       ON  p.OWNER      = d.REFERENCED_OWNER
       AND p.TABLE_NAME = d.REFERENCED_NAME
WHERE  d.NAME = 'VF_LSTP_BG_BS'
  AND  d.REFERENCED_TYPE = 'TABLE'
ORDER BY d.REFERENCED_NAME, p.PARTITION_NAME;


-- 5) Aufräumen nach dem Test
-- DROP EXTERNAL TABLE ext.all_dependencies;
