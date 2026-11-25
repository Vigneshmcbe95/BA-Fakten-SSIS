-- =============================================================================
-- Prüfung: Hat jeder Monat von TF_STEA eine eigene Oracle-Partition?
-- Hintergrund: SCR11 findet 242 Partitionswerte, Daten haben 324 Monate.
-- Ausführen in msi_dm_ast.
-- =============================================================================

-- 1. Owner/Tabelle prüfen (Schreibweise, Collation ist case-sensitiv)
SELECT DISTINCT OWNER, TABLE_NAME
FROM ext.[v_partition_info]
WHERE TABLE_NAME LIKE '%STEA%';

-- 2. Alle Partitionen der Reihe nach
SELECT PARTITION_POSITION, PARTITION_NAME, HIGH_VALUE, PARTITIONING_TYPE, PARTITION_KEY
FROM ext.[v_partition_info]
WHERE OWNER = 'XRO_DM_STAT_AST_AV_STEA' AND TABLE_NAME = 'TF_STEA'
ORDER BY PARTITION_POSITION;

-- 3. Jeder Monat 199912-202612: Partition vorhanden oder FEHLT
WITH hv AS (
    SELECT DISTINCT
           CASE WHEN TRY_CAST(HIGH_VALUE AS INT) % 100 = 13
                THEN (TRY_CAST(HIGH_VALUE AS INT) / 100 + 1) * 100 + 1
                ELSE TRY_CAST(HIGH_VALUE AS INT) END AS wert
    FROM ext.[v_partition_info]
    WHERE OWNER = 'XRO_DM_STAT_AST_AV_STEA' AND TABLE_NAME = 'TF_STEA'
),
kal AS (
    SELECT TOP (325) CAST(FORMAT(DATEADD(MONTH, ROW_NUMBER() OVER (ORDER BY (SELECT 1)) - 1, '19991201'), 'yyyyMM') AS INT) AS monat
    FROM sys.all_objects
)
SELECT k.monat,
       CASE WHEN h.wert IS NULL THEN 'FEHLT' ELSE 'OK' END AS oracle_partition
FROM kal k
LEFT JOIN hv h ON h.wert = k.monat
ORDER BY k.monat;

-- 4. Zusammenfassung
WITH hv AS (
    SELECT DISTINCT
           CASE WHEN TRY_CAST(HIGH_VALUE AS INT) % 100 = 13
                THEN (TRY_CAST(HIGH_VALUE AS INT) / 100 + 1) * 100 + 1
                ELSE TRY_CAST(HIGH_VALUE AS INT) END AS wert
    FROM ext.[v_partition_info]
    WHERE OWNER = 'XRO_DM_STAT_AST_AV_STEA' AND TABLE_NAME = 'TF_STEA'
),
kal AS (
    SELECT TOP (325) CAST(FORMAT(DATEADD(MONTH, ROW_NUMBER() OVER (ORDER BY (SELECT 1)) - 1, '19991201'), 'yyyyMM') AS INT) AS monat
    FROM sys.all_objects
)
SELECT SUM(CASE WHEN h.wert IS NULL THEN 0 ELSE 1 END) AS vorhanden,
       SUM(CASE WHEN h.wert IS NULL THEN 1 ELSE 0 END) AS fehlt,
       MIN(CASE WHEN h.wert IS NULL THEN k.monat END) AS erster_fehlender,
       MAX(CASE WHEN h.wert IS NULL THEN k.monat END) AS letzter_fehlender
FROM kal k
LEFT JOIN hv h ON h.wert = k.monat;
