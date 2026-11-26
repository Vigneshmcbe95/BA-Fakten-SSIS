-- =============================================================================
-- Prüfung: Spaltennamen tf_lstp_pers_bs_eink (SQL Server vs. Oracle/ext)
-- Hintergrund: SCR13 SELECT INTO -> Invalid column name
--              mdi_id_zverwerb_mm, mdi_id_zewt_mm, mdi_id_zset_mm
-- Ausführen in msi_dm_lst_sgbii.
-- =============================================================================

-- 1. Kandidaten in der ext-Tabelle (Oracle-Namen)
SELECT t.name AS ext_tabelle, c.name AS spalte, ty.name AS typ
FROM sys.external_tables t
JOIN sys.columns c ON c.object_id = t.object_id
JOIN sys.types ty ON ty.user_type_id = c.user_type_id
WHERE t.name LIKE '%lstp_pers_bs_eink%'
  AND (c.name LIKE '%verwerb%' OR c.name LIKE '%zewt%' OR c.name LIKE '%zset%' OR c.name LIKE 'mdi_id_z%')
ORDER BY t.name, c.name;

-- 2. Spalten der Faktentabelle, die in der ext-Tabelle fehlen
SELECT f.name AS fehlt_in_ext
FROM sys.columns f
WHERE f.object_id = OBJECT_ID('dbo.tf_lstp_pers_bs_eink')
  AND f.is_computed = 0
  AND NOT EXISTS (SELECT 1
                  FROM sys.external_tables t
                  JOIN sys.columns c ON c.object_id = t.object_id
                  WHERE t.name LIKE '%lstp_pers_bs_eink%'
                    AND c.name = f.name)
ORDER BY f.column_id;

-- 3. Spalten der ext-Tabelle, die in der Faktentabelle fehlen
SELECT t.name AS ext_tabelle, c.name AS fehlt_in_faktentabelle
FROM sys.external_tables t
JOIN sys.columns c ON c.object_id = t.object_id
WHERE t.name LIKE '%lstp_pers_bs_eink%'
  AND NOT EXISTS (SELECT 1 FROM sys.columns f
                  WHERE f.object_id = OBJECT_ID('dbo.tf_lstp_pers_bs_eink')
                    AND f.name = c.name)
ORDER BY t.name, c.column_id;
