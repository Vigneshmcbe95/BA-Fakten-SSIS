/* ================================================================================
   Debug_Partitionswerte_pruefen.sql
   Paket : Fakten Laden (SSIS)
   Zweck : Vergleicht die von SCR11 berechneten Partitionswerte (aus
           ext.v_partition_info HIGH_VALUE) mit den TATSAECHLICH in der
           Oracle-Quelle vorhandenen Partitionswerten (DISTINCT mow_id) -
           fuer EIN Verfahren.

   Hintergrund:
     SCR13 bricht ab, wenn eine Partition 0 Zeilen liefert, obwohl SCR11 sie
     als "mit Daten" eingeplant hat. Ursache ist i.d.R., dass SCR11 fuer das
     8-stellige Schema (mow_id) Werte aus den HIGH_VALUE-Metadaten ableitet
     (Datenwert = HIGH_VALUE - 1, Anhang 00 = leerer Anker), die in der
     Quelle keine Daten haben (leere Grenz-/Anker-/Zukunftspartition).

   Anwendung:
     1. Die 4 Variablen unten anpassen.
     2. In der Ziel-Datamart-DB ausfuehren (dort liegen dbo.tm_polybase_struktur
        und die PolyBase-Datenquelle [Oracle-<ENV>]).
     3. Ergebnis (A) DISTINCT mow_id  = was real existiert.
        Ergebnis (B) DISTINCT HIGH_VALUE = was SCR11 als Basis liest.
        Tauchen die fehlgeschlagenen Werte (z.B. 20251206 / 20260303 /
        20260402 / 20260600) in (B) auf, aber NICHT in (A), bestaetigt das:
        SCR11 gibt leere Grenz-/Ankerpartitionen aus -> Fix gehoert in SCR11.
     4. Aufraeumen: DROP EXTERNAL TABLE ext.dbg_<verfahren>;

   Hinweis: Die externe Debug-Tabelle bekommt einen eigenen Namen (dbg_<verf>),
            damit sie NICHT mit der Pipeline-Tabelle ext.<verf> kollidiert.
   ================================================================================ */


/* ===== Anpassen ============================================================== */
DECLARE @thema nvarchar(128) = N'bst';          -- Themengebiet (lower, wie in tm_polybase_struktur)
DECLARE @verf  nvarchar(128) = N'vf_bst_bv';    -- Verfahren / Oracle-Objekt (lower)
DECLARE @src   sysname       = N'Oracle-ESTAT'; -- = BA::ExtSourceName
DECLARE @env   sysname       = N'ESTAT';        -- Umgebung: ESTAT / ISTAT / PSTAT
/* ============================================================================= */

DECLARE @extName sysname = N'dbg_' + @verf;     -- eigener Name -> kein Konflikt mit ext.<verf>
DECLARE @cols nvarchar(max), @loc nvarchar(300), @sql nvarchar(max);

-- Vollstaendige externe Struktur aus denselben Metadaten wie SCR09
SELECT @cols = STRING_AGG(CAST(columns_ext AS nvarchar(max)), CONCAT(N',', CHAR(13), CHAR(10)))
               WITHIN GROUP (ORDER BY colno)
FROM   dbo.tm_polybase_struktur
WHERE  themengebiet = @thema AND tabname = @verf;

IF @cols IS NULL
    THROW 50001, 'Keine columns_ext Metadaten fuer dieses Verfahren/Themengebiet', 1;

SET @loc = UPPER(@env) + N'.' + UPPER(@thema) + N'.' + UPPER(@verf);

IF OBJECT_ID(N'ext.' + QUOTENAME(@extName), 'U') IS NOT NULL
    EXEC(N'DROP EXTERNAL TABLE ext.' + QUOTENAME(@extName) + N';');

SET @sql = N'CREATE EXTERNAL TABLE ext.' + QUOTENAME(@extName) + N' (' + CHAR(13) + CHAR(10) +
           @cols + CHAR(13) + CHAR(10) +
           N') WITH (DATA_SOURCE=[' + @src + N'], LOCATION=''' + @loc + N''');';
EXEC sp_executesql @sql;

-- (A) Tatsaechlich in Oracle vorhandene Partitionswerte
EXEC(N'SELECT DISTINCT mow_id AS Ist_mow_id FROM ext.' + QUOTENAME(@extName) + N' ORDER BY mow_id;');

-- (B) Was SCR11 als Basis liest (Rohgrenzen). Datenwert = HIGH_VALUE - 1 (8-stellig),
--     Anhang 00 = leerer Anker. OWNER-Filter = Themengebiet (wie in SCR11).
SELECT DISTINCT HIGH_VALUE
FROM   ext.[v_partition_info]
WHERE  TABLE_NAME = UPPER(@verf)
  AND  OWNER      = UPPER(@thema)
ORDER  BY HIGH_VALUE;

-- Aufraeumen (bei Bedarf einkommentieren):
-- EXEC(N'DROP EXTERNAL TABLE ext.' + QUOTENAME(@extName) + N';');


/* --------------------------------------------------------------------------------
   OPTIONAL: ext.v_partition_info anlegen, falls sie noch nicht existiert.
   Spaltentypen ggf. an die tatsaechliche Oracle-View anpassen. Die View liegt
   unter dem Partitionsschema BI_DM_EXPORT (= PARAM_PARTITION_SCHEMA).
   --------------------------------------------------------------------------------

IF OBJECT_ID('ext.v_partition_info','U') IS NOT NULL
    DROP EXTERNAL TABLE ext.[v_partition_info];

CREATE EXTERNAL TABLE ext.[v_partition_info]
(
    OWNER          NVARCHAR(128),
    TABLE_NAME     NVARCHAR(128),
    PARTITION_NAME NVARCHAR(128),
    HIGH_VALUE     NVARCHAR(4000)        -- SCR11 fuehrt Integer.TryParse darauf aus
)
WITH (
    DATA_SOURCE = [Oracle-ESTAT],                            -- = BA::ExtSourceName
    LOCATION    = 'ESTAT.BI_DM_EXPORT.V_PARTITION_INFO'      -- <ENV>.<partition_schema>.V_PARTITION_INFO
);
-------------------------------------------------------------------------------- */
