-- =====================================================================
-- Test SCR11-Datenprüfung: count(*) gegen TOP 1 / rownum = 1
-- =====================================================================

-- ---------------------------------------------------------------------
-- Teil 1: Oracle (SQL Developer)
-- Schema, Tabelle, Partitionsspalte und Wert anpassen; jede Abfrage 2x ausführen, Zeiten notieren.
-- ---------------------------------------------------------------------

select count(*) from MSI_DM_ZWG.TF_AGZU_VORGANG partition for (20260900);
select 1 from MSI_DM_ZWG.TF_AGZU_VORGANG partition for (20260900) where rownum = 1;
select 1 from MSI_DM_ZWG.TF_AGZU_VORGANG where MON_ID = 20260900 and rownum = 1;


-- ---------------------------------------------------------------------
-- Teil 2: SQL Server (SSMS) - drei Varianten über PolyBase mit Zeitmessung
-- Variante 1: COUNT(*) (liest ganze Partition, Vergleichswert)
-- Variante 2: TOP 1 1 (so prüft SCR11 heute)
-- Variante 3: TOP 1 1 mit FORCE EXTERNALPUSHDOWN (Abfrage wird an Oracle übergeben)
-- Ausgabe: je Variante und Lauf Ergebnis und Dauer in ms / s
-- ---------------------------------------------------------------------

SET NOCOUNT ON;

-- Hier anpassen
DECLARE @db      sysname = N'msi_dm_bvd';
DECLARE @schema  sysname = N'ext';
DECLARE @tabelle sysname = N'tf_bvd_aa';
DECLARE @partcol sysname = N'MON_ID';
DECLARE @wert    int     = 202610;
DECLARE @laeufe  int     = 2;      -- jede Variante 2x (1. Lauf = realistisch, 2. Lauf = Cache)

DECLARE @obj   nvarchar(400) = QUOTENAME(@db) + N'.' + QUOTENAME(@schema) + N'.' + QUOTENAME(@tabelle);
DECLARE @where nvarchar(400) = N' WHERE ' + QUOTENAME(@partcol) + N' = @w';

DECLARE @varianten TABLE (nr int, name nvarchar(100), sqltext nvarchar(max));
INSERT @varianten VALUES
 (1, N'COUNT(*)',                         N'SELECT @r = COUNT_BIG(*) FROM ' + @obj + @where),
 (2, N'TOP 1 1 (aktuell SCR11)',          N'SELECT TOP 1 @r = 1 FROM '     + @obj + @where),
 (3, N'TOP 1 1 + FORCE EXTERNALPUSHDOWN', N'SELECT TOP 1 @r = 1 FROM '     + @obj + @where + N' OPTION (FORCE EXTERNALPUSHDOWN)');

DECLARE @ergebnis TABLE (nr int, variante nvarchar(100), lauf int, ergebnis bigint, dauer_ms int, fehler nvarchar(4000));

DECLARE @nr int, @name nvarchar(100), @sql nvarchar(max), @lauf int, @r bigint, @start datetime2;

DECLARE c CURSOR LOCAL FAST_FORWARD FOR SELECT nr, name, sqltext FROM @varianten ORDER BY nr;
OPEN c;
FETCH NEXT FROM c INTO @nr, @name, @sql;
WHILE @@FETCH_STATUS = 0
BEGIN
    SET @lauf = 1;
    WHILE @lauf <= @laeufe
    BEGIN
        SET @r = NULL;
        SET @start = SYSDATETIME();
        BEGIN TRY
            EXEC sp_executesql @sql, N'@w int, @r bigint OUTPUT', @w = @wert, @r = @r OUTPUT;
            INSERT @ergebnis VALUES (@nr, @name, @lauf, @r, DATEDIFF(ms, @start, SYSDATETIME()), NULL);
        END TRY
        BEGIN CATCH
            INSERT @ergebnis VALUES (@nr, @name, @lauf, NULL, DATEDIFF(ms, @start, SYSDATETIME()), ERROR_MESSAGE());
        END CATCH;
        SET @lauf += 1;
    END;
    FETCH NEXT FROM c INTO @nr, @name, @sql;
END;
CLOSE c; DEALLOCATE c;

SELECT @obj AS tabelle, @partcol + N' = ' + CAST(@wert AS nvarchar(20)) AS partition,
       nr, variante, lauf, ergebnis, dauer_ms, CAST(dauer_ms / 1000.0 AS decimal(10,1)) AS dauer_s, fehler
FROM @ergebnis
ORDER BY nr, lauf;
