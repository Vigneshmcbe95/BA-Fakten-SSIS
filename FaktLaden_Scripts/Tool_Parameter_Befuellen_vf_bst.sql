/* =============================================================================
   Tool: Parameter befuellen - vf_bst_* Verfahren (aktuelle STL-Datei)
   -----------------------------------------------------------------------------
   Legt fuer die 8 Verfahren aus v.istat.thm_dm_stat_bst.111283.csv alle
   Pflichtparameter (SCR05) plus die ueblichen Zusatzparameter an.
   Idempotent: bereits vorhandene (Verfahren, Parameter)-Zeilen werden
   uebersprungen - der Lauf kann gefahrlos wiederholt werden.

   Pflichtparameter (SCR05): FaktentabelleTemplate, Faktentabelle,
   Faktenpartitionsspalte, Faktenkomprimierung, FaktenClusteredIndex,
   DateFormat, Anzahl_ParallelTasks

   WICHTIG vor dem Lauf: Partitionsspalte = mow_id wird angenommen
   (Wert 20230906 aus der STL-Datei entspricht dem mow_id-Muster).
   Bei Zweifel in Oracle pruefen, z. B.:
     SELECT DISTINCT MOW_ID FROM ext.[vf_bst_bv] ORDER BY 1;

   Lauf: auf msi_dm_bst_v3 ausfuehren.
   ============================================================================= */

USE [msi_dm_bst_v3];
SET NOCOUNT ON;

DECLARE @Datamart sysname = N'msi_dm_bst_v3_fakten';

/* ---- Die 8 Verfahren aus der aktuellen STL-Datei ---- */
DECLARE @verf TABLE (v sysname);
INSERT INTO @verf (v) VALUES
    ('vf_bst_bv'),
    ('vf_bst_bv_sonst'),
    ('vf_bst_sonst'),
    ('vf_bst_gb'),
    ('vf_bst_sv'),
    ('vf_bst_bew_bv'),
    ('vf_bst_bew_sv'),
    ('vf_bst_bew_gb');

/* ---- Parametervorlage (Pflicht + uebliche Zusatzparameter) ---- */
DECLARE @vorlage TABLE (Parameter sysname, Wert nvarchar(400), Beschreibung nvarchar(1000));
INSERT INTO @vorlage (Parameter, Wert, Beschreibung) VALUES
    ('FaktentabelleTemplate',    '',       'Spaltenschema fuer die Erstellung der Partitionstabelle'),
    ('Faktentabelle',            '',       'Name der Partitionstabelle'),
    ('Faktenpartitionsspalte',   'mow_id', 'mon_id oder mow_id'),
    ('Faktenkomprimierung',      'NONE',   'PAGE oder NONE'),
    ('FaktenClusteredIndex',     'CCI',    'Clustered Index auf Partitionsspalte erstellen TRUE oder FALSE'),
    ('DateFormat',               'dmy',    'Bulk Insert - DateFormat z.B.: dmy'),
    ('Anzahl_ParallelTasks',     '0',      'Ueberschreibt Package.MaxConcurrentExecutables'),
    ('FaktenNccIndex',           'FALSE',  'Nonclustered Columnstore Index TRUE oder FALSE'),
    ('Faktenpartitionsdatentyp', 'BIGINT', 'INT oder BIGINT');

/* ---- Fehlende (Verfahren, Parameter)-Kombinationen einfuegen ---- */
DECLARE @maxId INT = (SELECT ISNULL(MAX(Id),0) FROM dbo.tm_msi_dm_bst_v3_param);

INSERT INTO dbo.tm_msi_dm_bst_v3_param (Id, Datamart, Verfahren, Parameter, Wert, Beschreibung)
SELECT @maxId + ROW_NUMBER() OVER (ORDER BY v.v, x.Parameter),
       @Datamart,
       v.v,
       x.Parameter,
       CASE x.Parameter
           WHEN 'FaktentabelleTemplate' THEN v.v + '_template'
           WHEN 'Faktentabelle'         THEN v.v
           ELSE x.Wert
       END,
       x.Beschreibung
FROM   @verf v
CROSS JOIN @vorlage x
WHERE  NOT EXISTS (SELECT 1 FROM dbo.tm_msi_dm_bst_v3_param p
                   WHERE LOWER(LTRIM(RTRIM(p.Verfahren))) = v.v
                     AND LOWER(LTRIM(RTRIM(p.Parameter))) = LOWER(x.Parameter));

PRINT CAST(@@ROWCOUNT AS varchar(10)) + ' Parameterzeilen eingefuegt.';

/* ---- Kontrolle: alle 8 Verfahren muessen 9 Parameter haben,
       Pflichtparameter ohne leere Werte ---- */
SELECT p.Verfahren,
       COUNT(*)                                                          AS anzahl_parameter,
       MAX(CASE WHEN p.Parameter='Faktenpartitionsspalte' THEN p.Wert END) AS Partitionsspalte,
       MAX(CASE WHEN p.Parameter='Faktentabelle'          THEN p.Wert END) AS Faktentabelle,
       SUM(CASE WHEN p.Wert IS NULL OR LTRIM(RTRIM(p.Wert)) = '' THEN 1 ELSE 0 END) AS leere_werte
FROM   dbo.tm_msi_dm_bst_v3_param p
WHERE  LOWER(LTRIM(RTRIM(p.Verfahren))) IN (SELECT v FROM @verf)
GROUP BY p.Verfahren
ORDER BY p.Verfahren;
