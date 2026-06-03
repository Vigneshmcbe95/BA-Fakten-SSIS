-- ============================================================
-- ETL_Fakt_Reset_Override
-- Operator inserts a row here BEFORE retry to force a specific
-- verfahren back to a specific step.
-- The row is consumed (Verarbeitet=1) after the package reads it.
-- ============================================================

IF NOT EXISTS (
    SELECT 1 FROM sys.tables
    WHERE name = 'ETL_Fakt_Reset_Override'
    AND schema_id = SCHEMA_ID('dbo')
)
BEGIN
    CREATE TABLE dbo.ETL_Fakt_Reset_Override
    (
        ID            INT            IDENTITY(1,1) PRIMARY KEY,
        Verfahren     VARCHAR(200)   NOT NULL,
        ResetZu       VARCHAR(100)   NOT NULL,
        Grund         NVARCHAR(1000) NULL,
        ErstelltAm    DATETIME       NOT NULL DEFAULT GETDATE(),
        Verarbeitet   BIT            NOT NULL DEFAULT 0,
        VerarbeitetAm DATETIME       NULL,
        RunID         INT            NULL
    );

    PRINT 'ETL_Fakt_Reset_Override erstellt.';
END
ELSE
    PRINT 'ETL_Fakt_Reset_Override bereits vorhanden.';

-- ============================================================
-- Valid values for ResetZu (the status chain):
-- ============================================================
-- AUSSTEHEND                (restart from beginning)
-- SCHEMADATEN_KOPIERT       (redo template + everything after)
-- TEMPLATE_ERSTELLT         (redo ext table + everything after)
-- EXT_TABELLE_ERSTELLT      (redo fact table + everything after)
-- FAKTENTABELLE_ERSTELLT    (redo partitions + everything after)
-- PARTITIONSGRENZEN_ERSTELLT(redo staging + everything after)
-- STAGING_ERSTELLT          (redo DATEN_LADEN + everything after)  ← most common deep reset
-- DATEN_GELADEN             (redo STAGE_DML + everything after)
-- STAGE_DML_ERFOLG          (redo index + everything after)
-- INDEX_IN_OUT_ERSTELLT     (redo compression + everything after)
-- KOMPRIMIERUNG_ERSTELLT    (redo NCCI + everything after)
-- NCCI_OUT_ERSTELLT         (redo partition switch only)

-- ============================================================
-- EXAMPLES — how operator uses this:
-- ============================================================

-- Example 1: Column mismatch in PARTITIONSTAUSCH → need new _in tables
-- INSERT INTO dbo.ETL_Fakt_Reset_Override (Verfahren, ResetZu, Grund)
-- VALUES ('tf_meine_tabelle', 'STAGING_ERSTELLT', 'Spaltenfehler SWITCH - neue _in Tabellen benötigt');

-- Example 2: Wrong data loaded → full reload
-- INSERT INTO dbo.ETL_Fakt_Reset_Override (Verfahren, ResetZu, Grund)
-- VALUES ('tf_meine_tabelle', 'AUSSTEHEND', 'Falsche Daten geladen - kompletter Neustart');

-- Example 3: STAGE_DML proc had a bug → just redo from after load
-- No override needed — smart reset handles this automatically.

-- Example 4: Reset multiple verfahren at once
-- INSERT INTO dbo.ETL_Fakt_Reset_Override (Verfahren, ResetZu, Grund)
-- SELECT Verfahren, 'STAGING_ERSTELLT', 'Schema update - alle neu laden'
-- FROM dbo.ETL_Fkt_Arbeitsliste
-- WHERE Status = 'FEHLER';

-- ============================================================
-- ETL_Fakt_Laden_Abschluss
-- Written ONLY when a single _in partition fully completes loading.
-- On retry: if entry exists → skip reload (fully done)
--           if no entry     → truncate _in + reload (partial or not started)
-- ============================================================

IF NOT EXISTS (
    SELECT 1 FROM sys.tables
    WHERE name = 'ETL_Fakt_Laden_Abschluss'
    AND schema_id = SCHEMA_ID('dbo')
)
BEGIN
    CREATE TABLE dbo.ETL_Fakt_Laden_Abschluss
    (
        ID            INT            IDENTITY(1,1) PRIMARY KEY,
        RunID         INT            NOT NULL,
        Verfahren     VARCHAR(200)   NOT NULL,
        InTabelle     VARCHAR(300)   NOT NULL,   -- e.g. tf_xxx_in_202601
        Zeilen        INT            NOT NULL,
        AbgeschlossenAm DATETIME     NOT NULL DEFAULT GETDATE()
    );
    CREATE INDEX IX_Laden_Abschluss_InTabelle
        ON dbo.ETL_Fakt_Laden_Abschluss (InTabelle);
    PRINT 'ETL_Fakt_Laden_Abschluss erstellt.';
END
ELSE
    PRINT 'ETL_Fakt_Laden_Abschluss bereits vorhanden.';

-- ============================================================
-- CHECK: see pending overrides
-- ============================================================
-- SELECT * FROM dbo.ETL_Fakt_Reset_Override WHERE Verarbeitet = 0;

-- ============================================================
-- CHECK: see override history
-- ============================================================
-- SELECT * FROM dbo.ETL_Fakt_Reset_Override ORDER BY ErstelltAm DESC;
