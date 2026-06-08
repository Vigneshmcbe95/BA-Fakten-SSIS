/* =============================================================================
   Tool: Oracle-Partitionswerte schnell ueber PolyBase lesen
   -----------------------------------------------------------------------------
   Ziel : Statt "SELECT DISTINCT mow_id FROM ext.<fakt>" (scannt Milliarden Zeilen)
          die Partitionswerte aus dem Oracle Data Dictionary lesen.

   ERKENNTNIS (aus Diskussion):
     - Der PARTITIONSNAME enthaelt den Wert, z.B.  PARTITION_19990101
       -> mow_id = numerischer Teil des Partitionsnamens (hier 19990101).
     - PARTITION_NAME ist ein lesbarer VARCHAR2 -> KEIN LONG-Problem.
     - HIGH_VALUE (z.B. TO_DATE('1999-01-02'...)) ist die EXKLUSIVE Obergrenze
       (alles "less than" landet in der Partition) und ist Typ LONG.
       => HIGH_VALUE wird NICHT gelesen, wir nehmen nur den Partitionsnamen.
   ============================================================================= */


/* =============================================================================
   TEIL 1  -  AUF ORACLE ausfuehren
   -----------------------------------------------------------------------------
   View OHNE HIGH_VALUE (nur lesbare Spalten). Owner ggf. anpassen.
   =============================================================================

CREATE OR REPLACE VIEW STATRT.VM_TAB_PARTITIONS AS
SELECT table_owner,
       table_name,
       partition_name,
       partition_position
FROM   dba_tab_partitions
WHERE  table_owner = 'XRO_DM_STAT_BST';     -- <-- Oracle-Owner der Faktentabellen
-- (falls keine DBA-Rechte: all_tab_partitions statt dba_tab_partitions)

-- Test in Oracle:
-- SELECT * FROM STATRT.VM_TAB_PARTITIONS WHERE table_name='TF_BST_AUFENTHALT';

   ============================================================================= */


/* =============================================================================
   TEIL 2  -  AUF SQL SERVER (msi_dm_bst_v3) ausfuehren
   -----------------------------------------------------------------------------
   External Table auf die Oracle-View. KEINE HIGH_VALUE-Spalte (kein LONG).
   DATA_SOURCE = vorhandene Quelle eurer Fakt-ext-Tabellen.
   ============================================================================= */

IF EXISTS (SELECT 1 FROM sys.external_tables
           WHERE schema_id = SCHEMA_ID('ext') AND name = 'vm_tab_partitions')
    DROP EXTERNAL TABLE ext.[vm_tab_partitions];

CREATE EXTERNAL TABLE ext.[vm_tab_partitions]
(
    TABLE_OWNER         NVARCHAR(128) NULL,
    TABLE_NAME          NVARCHAR(128) NULL,
    PARTITION_NAME      NVARCHAR(128) NULL,
    PARTITION_POSITION  INT           NULL
)
WITH (
    DATA_SOURCE = [Oracle-istat],
    LOCATION    = 'ISTAT.STATRT.VM_TAB_PARTITIONS'   -- <-- ggf. anpassen
);


/* =============================================================================
   TEIL 3  -  Partitionswerte (mow_id) je Faktentabelle - schnell
   -----------------------------------------------------------------------------
   mow_id = Ziffern aus dem Partitionsnamen (z.B. PARTITION_19990101 -> 19990101).
   TRY_CONVERT liefert NULL fuer nicht-numerische Namen (z.B. SYS_P123) -> gefiltert.
   ============================================================================= */

SELECT DISTINCT
       TRY_CONVERT(BIGINT,
            -- nur die Ziffern aus dem Partitionsnamen ziehen:
            STUFF(PARTITION_NAME, 1, PATINDEX('%[0-9]%', PARTITION_NAME + '0') - 1, '')
       ) AS mow_id
FROM   ext.[vm_tab_partitions]
WHERE  TABLE_NAME = 'TF_BST_AUFENTHALT'        -- Oracle-Name (UPPER)
  AND  TRY_CONVERT(BIGINT,
            STUFF(PARTITION_NAME, 1, PATINDEX('%[0-9]%', PARTITION_NAME + '0') - 1, '')
       ) IS NOT NULL
ORDER  BY mow_id;

/* Einfachere Variante, falls der Praefix immer 'PARTITION_' ist: */
-- SELECT DISTINCT TRY_CONVERT(BIGINT, REPLACE(PARTITION_NAME,'PARTITION_','')) AS mow_id
-- FROM   ext.[vm_tab_partitions]
-- WHERE  TABLE_NAME = 'TF_BST_AUFENTHALT'
--   AND  TRY_CONVERT(BIGINT, REPLACE(PARTITION_NAME,'PARTITION_','')) IS NOT NULL
-- ORDER  BY mow_id;
