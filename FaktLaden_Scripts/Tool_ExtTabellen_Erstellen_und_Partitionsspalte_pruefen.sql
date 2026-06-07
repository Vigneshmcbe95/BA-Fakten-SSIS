/* =============================================================================
   Tool: Externe Tabellen fuer alle Steuerlisten-Verfahren erstellen
         und Partitionsspalte (mow_id / mon_id / mon_id_v) pruefen
   -----------------------------------------------------------------------------
   Zweck : Erstellt je Verfahren der Steuerliste die externe Tabelle
           (ext.<tabname>) aus den columns_ext-Metadaten in
           dbo.tm_polybase_struktur (gleiche Logik wie SCR09) und listet
           anschliessend, welche Tabelle eine Partitionsspalte besitzt.
   Lauf  : Auf der Datamart-DB ausfuehren (z. B. msi_dm_bst_v3).
   Hinweis: Quelle der Spalten ist dbo.tm_polybase_struktur (von SCR07 gefuellt).
            Verfahren ohne Schemadaten dort werden uebersprungen (PRINT).
   ============================================================================= */

SET NOCOUNT ON;

/* ---- Anpassen falls noetig (Werte aus dem Lauf-Log) ---- */
DECLARE @DataSource sysname = N'Oracle-istat';   -- PARAM_EXT_SOURCE_NAME
DECLARE @Env        sysname = N'ISTAT';          -- 1. Teil von ExtTableLocation (istat.STATRT...)
DECLARE @ExtSchema  sysname = N'ext';
/* ------------------------------------------------------- */

DECLARE @tab sysname, @thema sysname, @cols nvarchar(max), @loc nvarchar(400), @sql nvarchar(max);

DECLARE cur CURSOR LOCAL FAST_FORWARD FOR
    SELECT DISTINCT ps.tabname, ps.themengebiet
    FROM   dbo.tm_polybase_struktur ps
    WHERE  EXISTS (SELECT 1 FROM dbo.tm_steuerlistenfile_Fakten s
                   WHERE LOWER(LTRIM(RTRIM(s.tabelle))) = LOWER(ps.tabname));

OPEN cur;
FETCH NEXT FROM cur INTO @tab, @thema;
WHILE @@FETCH_STATUS = 0
BEGIN
    SELECT @cols = STRING_AGG(CAST(columns_ext AS nvarchar(max)), ',' + CHAR(13)+CHAR(10))
                   WITHIN GROUP (ORDER BY colno)
    FROM   dbo.tm_polybase_struktur
    WHERE  tabname = @tab AND themengebiet = @thema;

    IF @cols IS NULL
        PRINT 'KEINE Schemadaten: ' + @tab;
    ELSE
    BEGIN
        SET @loc = @Env + N'.' + UPPER(@thema) + N'.' + UPPER(@tab);

        IF EXISTS (SELECT 1 FROM sys.external_tables
                   WHERE schema_id = SCHEMA_ID(@ExtSchema) AND name = @tab)
        BEGIN
            SET @sql = N'DROP EXTERNAL TABLE ' + QUOTENAME(@ExtSchema) + N'.' + QUOTENAME(@tab) + N';';
            EXEC sp_executesql @sql;
        END

        SET @sql = N'CREATE EXTERNAL TABLE ' + QUOTENAME(@ExtSchema) + N'.' + QUOTENAME(@tab) + N' ('
                 + CHAR(13)+CHAR(10) + @cols + CHAR(13)+CHAR(10)
                 + N') WITH (DATA_SOURCE=' + QUOTENAME(@DataSource)
                 + N', LOCATION=''' + @loc + N''');';
        BEGIN TRY
            EXEC sp_executesql @sql;
            PRINT 'OK: ' + @ExtSchema + '.' + @tab + '  ->  ' + @loc;
        END TRY
        BEGIN CATCH
            PRINT 'FEHLER ' + @tab + ': ' + ERROR_MESSAGE();
        END CATCH
    END

    FETCH NEXT FROM cur INTO @tab, @thema;
END
CLOSE cur; DEALLOCATE cur;

/* ===== Pruefung: welche ext-Tabelle hat mow_id / mon_id / mon_id_v ===== */
SELECT t.name AS ext_tabelle,
       MAX(CASE WHEN LOWER(c.name)='mow_id'   THEN 'mow_id'   END) AS hat_mow_id,
       MAX(CASE WHEN LOWER(c.name)='mon_id'   THEN 'mon_id'   END) AS hat_mon_id,
       MAX(CASE WHEN LOWER(c.name)='mon_id_v' THEN 'mon_id_v' END) AS hat_mon_id_v,
       COALESCE(
         MAX(CASE WHEN LOWER(c.name)='mow_id'   THEN 'mow_id'   END),
         MAX(CASE WHEN LOWER(c.name)='mon_id_v' THEN 'mon_id_v' END),
         MAX(CASE WHEN LOWER(c.name)='mon_id'   THEN 'mon_id'   END)
       ) AS partitionsspalte
FROM   sys.external_tables t
JOIN   sys.columns c ON c.object_id = t.object_id
WHERE  t.schema_id = SCHEMA_ID(@ExtSchema)
GROUP BY t.name
ORDER BY t.name;
