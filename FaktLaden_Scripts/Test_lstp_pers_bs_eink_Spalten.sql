-- =============================================================================
-- Prüfung: Spaltennamen tf_lstp_pers_bs_eink (SQL Server vs. Oracle-DDL)
-- Hintergrund: SCR13 SELECT INTO -> Invalid column name
--              mdi_id_zverwerb_mm, mdi_id_zewt_mm, mdi_id_zset_mm
-- Oracle-Spalten aus dbo.ddl_staging_fakt (Kopie der Oracle-DDL aus SCR07),
-- keine ext-Tabelle nötig.
-- Ausführen in msi_dm_lst_sgbii.
-- =============================================================================

-- 1. Oracle-Namen der Kandidaten
SELECT THMNAME, TABNAME, COLNO, COLNAME, TYPNAME
FROM dbo.ddl_staging_fakt
WHERE TABNAME LIKE '%LSTP_PERS_BS_EINK%'
  AND (COLNAME LIKE '%VERWERB%' OR COLNAME LIKE '%ZEWT%' OR COLNAME LIKE '%ZSET%' OR COLNAME LIKE 'MDI_ID_Z%')
ORDER BY TABNAME, COLNO;

-- 2. Spalten der Faktentabelle, die in Oracle fehlen
SELECT f.name AS fehlt_in_oracle
FROM sys.columns f
WHERE f.object_id = OBJECT_ID('dbo.tf_lstp_pers_bs_eink')
  AND f.is_computed = 0
  AND NOT EXISTS (SELECT 1 FROM dbo.ddl_staging_fakt d
                  WHERE d.TABNAME LIKE '%LSTP_PERS_BS_EINK%'
                    AND UPPER(d.COLNAME) COLLATE DATABASE_DEFAULT = UPPER(f.name) COLLATE DATABASE_DEFAULT)
ORDER BY f.column_id;

-- 3. Oracle-Spalten, die in der Faktentabelle fehlen
SELECT d.TABNAME, d.COLNO, d.COLNAME
FROM dbo.ddl_staging_fakt d
WHERE d.TABNAME LIKE '%LSTP_PERS_BS_EINK%'
  AND NOT EXISTS (SELECT 1 FROM sys.columns f
                  WHERE f.object_id = OBJECT_ID('dbo.tf_lstp_pers_bs_eink')
                    AND UPPER(f.name) COLLATE DATABASE_DEFAULT = UPPER(d.COLNAME) COLLATE DATABASE_DEFAULT)
ORDER BY d.TABNAME, d.COLNO;

-- 4. Falls ddl_staging_fakt leer/nicht vorhanden: gleiche Prüfung über tm_polybase_struktur
SELECT tabname, colno, colname
FROM dbo.tm_polybase_struktur
WHERE tabname LIKE '%lstp_pers_bs_eink%'
  AND (colname LIKE '%verwerb%' OR colname LIKE '%zewt%' OR colname LIKE '%zset%' OR colname LIKE 'mdi_id_z%')
ORDER BY tabname, colno;
