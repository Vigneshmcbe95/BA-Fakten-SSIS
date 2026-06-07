/* =============================================================================
   Tool: Parameter-Pruefung je Verfahren
   -----------------------------------------------------------------------------
   Prueft fuer alle Verfahren der Arbeitsliste:
     1) Sind alle Parameter-Werte befuellt? (JA/NEIN + Anzahl leerer Werte)
     2) Ist die Faktenpartitionsspalte = mow_id?
   Lauf: auf der Datamart-DB (msi_dm_bst_v3).
   ============================================================================= */

SET NOCOUNT ON;

/* ---- Pflichtparameter (anpassen falls noetig) ---- */
DECLARE @Pflicht TABLE (Parameter sysname);
INSERT INTO @Pflicht (Parameter) VALUES
    ('Faktentabelle'),('FaktentabelleTemplate'),('Faktenpartitionsspalte'),
    ('Faktenkomprimierung'),('FaktenClusteredIndex'),('Anzahl_ParallelTasks');
DECLARE @PflichtAnzahl INT = (SELECT COUNT(*) FROM @Pflicht);
/* -------------------------------------------------- */

;WITH prm AS (
    SELECT p.Verfahren,
           LTRIM(RTRIM(p.Parameter)) AS Parameter,
           p.Wert
    FROM   [msi_dm_bst_v3].[dbo].[tm_msi_dm_bst_v3_param] p
    WHERE  EXISTS (SELECT 1 FROM dbo.ETL_Fkt_Arbeitsliste a
                   WHERE LOWER(LTRIM(RTRIM(a.Verfahren))) = LOWER(LTRIM(RTRIM(p.Verfahren))))
)
SELECT
    prm.Verfahren,
    MAX(CASE WHEN prm.Parameter = 'Faktenpartitionsspalte'   THEN prm.Wert END) AS Partitionsspalte,
    CASE WHEN LOWER(LTRIM(RTRIM(
              MAX(CASE WHEN prm.Parameter='Faktenpartitionsspalte' THEN prm.Wert END)))) = 'mow_id'
         THEN 'JA' ELSE 'NEIN' END                                              AS ist_mow_id,
    MAX(CASE WHEN prm.Parameter = 'Faktenpartitionsdatentyp' THEN prm.Wert END) AS Partitionsdatentyp,
    COUNT(*)                                                                    AS anzahl_parameter,
    SUM(CASE WHEN prm.Wert IS NULL OR LTRIM(RTRIM(prm.Wert)) = '' THEN 1 ELSE 0 END) AS leere_werte,
    -- alle vorhandenen Werte befuellt?
    CASE WHEN SUM(CASE WHEN prm.Wert IS NULL OR LTRIM(RTRIM(prm.Wert)) = '' THEN 1 ELSE 0 END) = 0
         THEN 'JA' ELSE 'NEIN' END                                              AS alle_werte_befuellt,
    -- alle Pflichtparameter vorhanden UND befuellt?
    COUNT(DISTINCT CASE WHEN prm.Parameter IN (SELECT Parameter FROM @Pflicht)
                         AND prm.Wert IS NOT NULL AND LTRIM(RTRIM(prm.Wert)) <> ''
                        THEN prm.Parameter END)                                 AS pflicht_ok,
    CASE WHEN COUNT(DISTINCT CASE WHEN prm.Parameter IN (SELECT Parameter FROM @Pflicht)
                         AND prm.Wert IS NOT NULL AND LTRIM(RTRIM(prm.Wert)) <> ''
                        THEN prm.Parameter END) = @PflichtAnzahl
         THEN 'JA' ELSE 'NEIN' END                                             AS pflicht_vollstaendig
FROM   prm
GROUP BY prm.Verfahren
ORDER BY prm.Verfahren;


/* ---- Detailliste: welche Werte sind leer? (zum Nachbessern) ---- */
SELECT p.Verfahren, LTRIM(RTRIM(p.Parameter)) AS Parameter, p.Wert
FROM   [msi_dm_bst_v3].[dbo].[tm_msi_dm_bst_v3_param] p
WHERE  (p.Wert IS NULL OR LTRIM(RTRIM(p.Wert)) = '')
  AND  EXISTS (SELECT 1 FROM dbo.ETL_Fkt_Arbeitsliste a
               WHERE LOWER(LTRIM(RTRIM(a.Verfahren))) = LOWER(LTRIM(RTRIM(p.Verfahren))))
ORDER BY p.Verfahren, Parameter;
