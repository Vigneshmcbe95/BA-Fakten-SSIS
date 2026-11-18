-- =============================================================================
-- Prüfung: Aufruf usp_Faktenladen_StagetabellenDML_* (SCR14)
-- Hintergrund: glf_id_xsoz in tf_lstp_pers_bs nicht auf 0 gesetzt,
--              in _eink und _bst_verb schon.
-- Ausführen in der Datamart-Datenbank (z.B. msi_dm_lst_sgbii).
-- =============================================================================

-- 1. Prozeduren vorhanden? Schema, Erstell- und Änderungsdatum
SELECT SCHEMA_NAME(schema_id) AS schema_name, name, create_date, modify_date
FROM sys.procedures
WHERE name IN ('usp_Faktenladen_StagetabellenDML_tf_lstp_pers_bs',
               'usp_Faktenladen_StagetabellenDML_tf_lstp_pers_bs_eink',
               'usp_Faktenladen_StagetabellenDML_tf_lstp_pers_bs_bst_verb')
ORDER BY name;

-- 2. Code der Prozedur: wird glf_id_xsoz gesetzt?
SELECT OBJECT_DEFINITION(OBJECT_ID('dbo.usp_Faktenladen_StagetabellenDML_tf_lstp_pers_bs')) AS code_pers_bs;

-- 3. Wird die Prozedur von SCR14 gefunden? (gleiche Prüfung wie im Skript, 1 = gefunden)
SELECT COUNT(*) AS gefunden
FROM sys.objects
WHERE object_id = OBJECT_ID('usp_Faktenladen_StagetabellenDML_tf_lstp_pers_bs')
  AND type IN ('P', 'PC');

-- 4. Standardschema des angemeldeten Benutzers
SELECT SUSER_SNAME() AS login_name, USER_NAME() AS user_name, SCHEMA_NAME() AS default_schema;

-- 5. Daten: Anteil glf_id_xsoz = 0 je Tabelle (Beispielpartition)
SELECT 'tf_lstp_pers_bs' AS tabelle, COUNT(*) AS gesamt, SUM(CASE WHEN glf_id_xsoz = 0 THEN 1 ELSE 0 END) AS xsoz_0
FROM dbo.tf_lstp_pers_bs WHERE mow_id = 20250503
UNION ALL
SELECT 'tf_lstp_pers_bs_eink', COUNT(*), SUM(CASE WHEN glf_id_xsoz = 0 THEN 1 ELSE 0 END)
FROM dbo.tf_lstp_pers_bs_eink WHERE mow_id = 20250503
UNION ALL
SELECT 'tf_lstp_pers_bs_bst_verb', COUNT(*), SUM(CASE WHEN glf_id_xsoz = 0 THEN 1 ELSE 0 END)
FROM dbo.tf_lstp_pers_bs_bst_verb WHERE mow_id = 20250503;

-- 6. Letzter Status in der Arbeitsliste (Status / LetzterSchritt)
SELECT TOP 20 a.RunID, a.Verfahren, a.Status, a.LetzterSchritt
FROM dbo.ETL_Fakt_Arbeitsliste a
WHERE a.Verfahren LIKE '%lstp_pers_bs%'
ORDER BY a.RunID DESC, a.Verfahren;
