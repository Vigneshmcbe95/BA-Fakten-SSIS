-- =============================================================================
--  Tool   : Partitionswerte in Oracle pruefen (tf_bst_hr_25)
--  Zweck  : Untersuchen, welche mon_id_v-Werte in Oracle tatsaechlich
--           existieren, bevor ein Wert in die STL-Datei geschrieben wird.
--           Hintergrund: Lauf mit mon_id_v = 20250206 lud 0 Zeilen -
--           die echten Werte folgen vermutlich dem Muster YYYYMM00.
--  Nutzung: Im SSMS gegen die Datamart-DB ausfuehren, Schritt fuer Schritt.
-- =============================================================================

USE [msi_dm_bst_v3];
GO

-- -----------------------------------------------------------------------------
-- Schritt 1: Pruefen, ob die externe Tabelle aus dem Lauf noch existiert
--            (SCR09 legt ext.tf_bst_hr_25 an; bei erfolgreichem Lauf bleibt sie)
-- -----------------------------------------------------------------------------
SELECT name, create_date
FROM   sys.external_tables
WHERE  name = 'tf_bst_hr_25';
GO

-- -----------------------------------------------------------------------------
-- Schritt 2 (nur falls Schritt 1 nichts liefert): externe Tabelle anlegen.
--            Struktur = columns_ext (alle Spalten NULL, Namen UPPERCASE).
-- -----------------------------------------------------------------------------
IF NOT EXISTS (SELECT 1 FROM sys.external_tables WHERE name = 'tf_bst_hr_25')
BEGIN
    EXEC ('
    CREATE EXTERNAL TABLE ext.[tf_bst_hr_25]
    (
        MON_ID_V      int           NULL,
        MOW_ID        int           NULL,
        MON_ID        int           NULL,
        KNG_ID        smallint      NULL,
        BXE_ID        int           NULL,
        BZP_ID        smallint      NULL,
        XPH_ID        smallint      NULL,
        BLD_ID        int           NULL,
        C25_ID        smallint      NULL,
        SEX_ID        smallint      NULL,
        ABZ_ID        smallint      NULL,
        BFR_ID        smallint      NULL,
        DIM_25_ANTEIL decimal(11,8) NULL,
        HR_25_ANZAHL  int           NULL
    )
    WITH (
        DATA_SOURCE = [Oracle-istat],
        LOCATION    = N''ISTAT.XRO_DM_STAT_BST.TF_BST_HR_25''
    );');
END
GO

-- -----------------------------------------------------------------------------
-- Schritt 3: Welche mon_id_v-Werte gibt es wirklich in Oracle?
--            (Das ist die Liste der gueltigen Werte fuer die STL-Datei.)
-- -----------------------------------------------------------------------------
SELECT   MON_ID_V, COUNT(*) AS Zeilen
FROM     ext.[tf_bst_hr_25]
GROUP BY MON_ID_V
ORDER BY MON_ID_V;
GO

-- -----------------------------------------------------------------------------
-- Schritt 4: Den konkreten Wert aus dem letzten Lauf gegenpruefen
--            (erwartet: 0 -> deshalb wurde nichts geladen)
-- -----------------------------------------------------------------------------
SELECT COUNT(*) AS Zeilen_20250206
FROM   ext.[tf_bst_hr_25]
WHERE  MON_ID_V = 20250206;
GO

-- -----------------------------------------------------------------------------
-- Schritt 5 (Vergleich): Was steht bereits in der MSSQL-Faktentabelle?
-- -----------------------------------------------------------------------------
SELECT   mon_id_v, COUNT(*) AS Zeilen
FROM     dbo.[tf_bst_hr_25]
GROUP BY mon_id_v
ORDER BY mon_id_v;
GO
