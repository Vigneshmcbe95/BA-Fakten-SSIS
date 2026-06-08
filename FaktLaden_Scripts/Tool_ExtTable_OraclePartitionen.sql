/* =============================================================================
   Tool: Oracle-Partitionswerte schnell ueber PolyBase lesen
   -----------------------------------------------------------------------------
   Ziel : Statt "SELECT DISTINCT mow_id FROM ext.<fakt>" (scannt Milliarden Zeilen)
          die Partitions-Grenzwerte aus dem Oracle Data Dictionary lesen
          (ALL_TAB_PARTITIONS) -> sofort verfuegbar.

   WICHTIG: ALL_TAB_PARTITIONS.HIGH_VALUE ist in Oracle vom Typ LONG.
            PolyBase kann LONG NICHT lesen. Deshalb:
            (1) Oracle-View anlegen, die HIGH_VALUE als ZAHL liefert
            (2) PolyBase External Table auf DIESE View
            (3) Werte je Tabelle abfragen
   ============================================================================= */


/* =============================================================================
   TEIL 1  -  AUF ORACLE ausfuehren (vom DBA / mit passenden Rechten)
   -----------------------------------------------------------------------------
   Legt eine View an, die je Tabelle/Partition den Grenzwert (HIGH_VALUE)
   als NUMBER bereitstellt. Der DBMS_XMLGEN-Trick liest den LONG als Text.
   Owner/Schema (XRO_DM_STAT_BST) ggf. anpassen.
   =============================================================================

CREATE OR REPLACE VIEW STATRT.VM_TAB_PARTITIONS AS
SELECT tp.table_owner,
       tp.table_name,
       tp.partition_name,
       tp.partition_position,
       TO_NUMBER(
         EXTRACTVALUE(
           DBMS_XMLGEN.GETXMLTYPE(
                'SELECT high_value FROM all_tab_partitions'
             || ' WHERE table_owner='''||tp.table_owner||''''
             || ' AND table_name='''  ||tp.table_name ||''''
             || ' AND partition_name='''||tp.partition_name||''''
           ), '//text()'
         )
       ) AS high_value_num
FROM   all_tab_partitions tp
WHERE  tp.table_owner = 'XRO_DM_STAT_BST';     -- <-- Oracle-Owner der Faktentabellen

-- Test in Oracle:
-- SELECT * FROM STATRT.VM_TAB_PARTITIONS WHERE table_name='TF_BST_AUFENTHALT';

   ============================================================================= */


/* =============================================================================
   TEIL 2  -  AUF SQL SERVER (msi_dm_bst_v3) ausfuehren
   -----------------------------------------------------------------------------
   External Table auf die Oracle-View. DATA_SOURCE = vorhandene Quelle eurer
   Fakt-ext-Tabellen (Oracle-istat). LOCATION = Owner.Schema.View (UPPER).
   ============================================================================= */

IF EXISTS (SELECT 1 FROM sys.external_tables
           WHERE schema_id = SCHEMA_ID('ext') AND name = 'vm_tab_partitions')
    DROP EXTERNAL TABLE ext.[vm_tab_partitions];

CREATE EXTERNAL TABLE ext.[vm_tab_partitions]
(
    TABLE_OWNER         NVARCHAR(128) NULL,
    TABLE_NAME          NVARCHAR(128) NULL,
    PARTITION_NAME      NVARCHAR(128) NULL,
    PARTITION_POSITION  INT           NULL,
    HIGH_VALUE_NUM      BIGINT        NULL
)
WITH (
    DATA_SOURCE = [Oracle-istat],
    LOCATION    = 'ISTAT.STATRT.VM_TAB_PARTITIONS'   -- <-- ggf. anpassen
);


/* =============================================================================
   TEIL 3  -  Partitionswerte je Faktentabelle abfragen (schnell)
   -----------------------------------------------------------------------------
   HINWEIS zur Bedeutung von HIGH_VALUE_NUM:
     - RANGE-Partition ("VALUES LESS THAN (X)"): Grenze ist X (exklusiv).
       Der enthaltene mow_id-Wert ist dann i.d.R. X-1  -> ggf. -1 rechnen.
     - LIST-Partition  ("VALUES (X)"): Grenze IST der Wert X.
   Prueft an einer bekannten Tabelle, welcher Fall vorliegt, und passt
   die Spalte (HIGH_VALUE_NUM bzw. HIGH_VALUE_NUM-1) entsprechend an.
   ============================================================================= */

-- Alle Partitionswerte einer Faktentabelle:
SELECT PARTITION_POSITION, HIGH_VALUE_NUM
FROM   ext.[vm_tab_partitions]
WHERE  TABLE_NAME = 'TF_BST_AUFENTHALT'      -- Oracle-Name (UPPER)
ORDER  BY PARTITION_POSITION;

-- Distinct-Liste der mow_id-Werte (Beispiel LIST-Partition: Wert = HIGH_VALUE_NUM):
-- SELECT DISTINCT HIGH_VALUE_NUM AS mow_id
-- FROM   ext.[vm_tab_partitions]
-- WHERE  TABLE_NAME = 'TF_BST_AUFENTHALT'
--   AND  HIGH_VALUE_NUM IS NOT NULL
-- ORDER  BY mow_id;
