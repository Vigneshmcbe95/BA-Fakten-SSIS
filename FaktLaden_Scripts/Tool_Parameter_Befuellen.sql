/* =============================================================================
   Tool: Parameter befuellen
   -----------------------------------------------------------------------------
   1) Fuellt die 13 noch leeren vf_bst_* Verfahren mit komplettem Parametersatz
      (Vorlage: tf_bst_aufenthalt), Partitionsspalte = mow_id, Datentyp = BIGINT.
   2) Korrigiert tf_bst_hr / tf_bst_hr_25: Faktenpartitionsspalte mon_id_v -> mow_id.
   3) Stellt sicher, dass jedes Verfahren Faktenpartitionsdatentyp = BIGINT hat.
   Lauf: auf msi_dm_bst_v3 ausfuehren.
   ============================================================================= */

SET NOCOUNT ON;

DECLARE @Datamart sysname = N'msi_dm_bst_v3_fakten';

/* ---- (1) Liste der zu befuellenden Verfahren ---- */
DECLARE @verf TABLE (v sysname);
INSERT INTO @verf (v) VALUES
    ('vf_bst_bew_bm_denorm'),('vf_bst_gb'),('vf_bst_sv_vm'),('vf_bst_sv'),
    ('vf_bst_sonst'),('vf_bst_kfb_vm'),('vf_bst_bew_bv'),('vf_bst_bew_gb'),
    ('vf_bst_geb_vm'),('vf_bst_bv_sonst'),('vf_bst_bv'),('vf_bst_bew_sv'),
    ('vf_bst_btg_v2_vm');

/* Parametervorlage (Wert '' = wird je Verfahren auf den Verfahrensnamen gesetzt) */
;WITH vorlage(Parameter, Wert, Beschreibung) AS (
    SELECT * FROM (VALUES
        ('FaktentabelleTemplate',    '',      'Spaltenschema für die Erstellung der Partitionstabelle'),
        ('Faktentabelle',            '',      'Name der Partitionstabelle'),
        ('Faktenpartitionsspalte',   'mow_id','mon_id oder mow_id'),
        ('Faktenkomprimierung',      'NONE',  'PAGE oder NONE'),
        ('FaktenClusteredIndex',     'CCI',   'Clustered Index auf Partitionsspalte erstellen TRUE oder FALSE'),
        ('ZeilenTrennzeichen',       '2',     'Bulk Insert - Zeilentrenner 1=LF, 2=Spaltentrenner+LF, 3=CR+LF, 4=Spaltentrenner+CR+LF'),
        ('SpaltenTrennzeichen',      '{|}',   'Bulk Insert - Spaltentrenner'),
        ('DateFormat',               'dmy',   'Bulk Insert - DateFormat z.B.: dmy'),
        ('Faktendatei',              '',      'Faktenquelldatei'),
        ('EndungUnload',             '.unl',  'Endung der Unload-Dateien (z.B.: .unl) oder leer'),
        ('Anzahl_ParallelTasks',     '0',     'Überschreibt Package.MaxConcurrentExecutables'),
        ('FaktenNccIndex',           'FALSE', 'Nonclustered Columnstore Index TRUE oder FALSE'),
        ('Faktenpartitionsdatentyp', 'BIGINT','INT oder BIGINT')
    ) t(Parameter, Wert, Beschreibung)
)
INSERT INTO dbo.tm_msi_dm_bst_v3_param (Datamart, Verfahren, Parameter, Wert, Beschreibung)
SELECT @Datamart, v.v, x.Parameter,
       CASE x.Parameter
           WHEN 'FaktentabelleTemplate' THEN v.v + '_template'
           WHEN 'Faktentabelle'         THEN v.v
           WHEN 'Faktendatei'           THEN v.v
           ELSE x.Wert
       END,
       x.Beschreibung
FROM   @verf v
CROSS JOIN vorlage x
WHERE  NOT EXISTS (SELECT 1 FROM dbo.tm_msi_dm_bst_v3_param p
                   WHERE p.Verfahren = v.v AND p.Parameter = x.Parameter);

PRINT 'Parameter fuer leere vf_-Verfahren ergaenzt.';

/* ---- (2) tf_bst_hr / tf_bst_hr_25: mon_id_v -> mow_id ---- */
UPDATE dbo.tm_msi_dm_bst_v3_param
SET    Wert = 'mow_id'
WHERE  Parameter = 'Faktenpartitionsspalte'
  AND  Verfahren IN ('tf_bst_hr','tf_bst_hr_25')
  AND  LOWER(LTRIM(RTRIM(Wert))) = 'mon_id_v';

PRINT 'Partitionsspalte fuer tf_bst_hr / tf_bst_hr_25 auf mow_id gesetzt.';

/* ---- (3) Faktenpartitionsdatentyp = BIGINT sicherstellen ---- */
-- vorhandene auf BIGINT setzen
UPDATE dbo.tm_msi_dm_bst_v3_param
SET    Wert = 'BIGINT'
WHERE  Parameter = 'Faktenpartitionsdatentyp'
  AND  UPPER(LTRIM(RTRIM(ISNULL(Wert,'')))) <> 'BIGINT';

-- fehlende anlegen (fuer Verfahren, die den Parameter noch nicht haben)
INSERT INTO dbo.tm_msi_dm_bst_v3_param (Datamart, Verfahren, Parameter, Wert, Beschreibung)
SELECT DISTINCT @Datamart, p.Verfahren, 'Faktenpartitionsdatentyp', 'BIGINT', 'INT oder BIGINT'
FROM   dbo.tm_msi_dm_bst_v3_param p
WHERE  NOT EXISTS (SELECT 1 FROM dbo.tm_msi_dm_bst_v3_param q
                   WHERE q.Verfahren = p.Verfahren AND q.Parameter = 'Faktenpartitionsdatentyp');

PRINT 'Faktenpartitionsdatentyp ueberall auf BIGINT gesetzt/ergaenzt.';

/* ---- Kontrolle ---- */
SELECT Verfahren,
       MAX(CASE WHEN Parameter='Faktenpartitionsspalte'   THEN Wert END) AS Partitionsspalte,
       MAX(CASE WHEN Parameter='Faktenpartitionsdatentyp' THEN Wert END) AS Datentyp,
       COUNT(*) AS anzahl_parameter
FROM   dbo.tm_msi_dm_bst_v3_param
GROUP BY Verfahren
ORDER BY Verfahren;
