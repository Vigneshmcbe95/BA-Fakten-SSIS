-- =============================================================================
--  Tool   : Partitionierung einer dbo-Tabelle anzeigen
--  Zweck  : Zeigt fuer (partitionierte) Tabellen das Partitionsschema, die
--           Partitionsfunktion, die tatsaechliche Partitionsspalte und die
--           Erstellungsdaten - z. B. um zu pruefen, ob die physische
--           Partitionierung zur Faktenpartitionsspalte in der
--           Parametertabelle passt (Namenskonvention aus SCR10:
--           PF_<spalte>_<tabelle> / PS_<spalte>_<tabelle>).
--  Nutzung: Im SSMS gegen die Datamart-DB ausfuehren.
-- =============================================================================

USE [msi_dm_bst_v3];
GO

-- -----------------------------------------------------------------------------
-- Abfrage 1: Schema / Funktion / Partitionsspalte / Erstellungsdatum
--            (WHERE-Filter weglassen, um ALLE partitionierten Tabellen zu sehen)
-- -----------------------------------------------------------------------------
SELECT
    t.name                AS Tabelle,
    ps.name               AS PartitionsSchema,
    pf.name               AS PartitionsFunktion,
    c.name                AS Partitionsspalte,
    pf.create_date        AS PF_erstellt_am,
    t.create_date         AS Tabelle_erstellt_am,
    pf.fanout             AS AnzahlPartitionen
FROM sys.tables t
JOIN sys.indexes i
      ON i.object_id = t.object_id AND i.index_id IN (0,1)   -- Heap oder Clustered/CCI
JOIN sys.partition_schemes ps
      ON ps.data_space_id = i.data_space_id
JOIN sys.partition_functions pf
      ON pf.function_id = ps.function_id
JOIN sys.index_columns ic
      ON ic.object_id = i.object_id AND ic.index_id = i.index_id AND ic.partition_ordinal = 1
JOIN sys.columns c
      ON c.object_id = ic.object_id AND c.column_id = ic.column_id
WHERE t.name = 'tf_bst_hr_25'      -- <- Tabellenname anpassen oder Zeile entfernen
ORDER BY t.name;
GO

-- -----------------------------------------------------------------------------
-- Abfrage 2: Grenzwerte der Partitionsfunktion (welche Partition = welcher Bereich)
-- -----------------------------------------------------------------------------
SELECT pf.name AS PartitionsFunktion,
       prv.boundary_id,
       prv.value AS Grenzwert
FROM sys.partition_functions pf
JOIN sys.partition_range_values prv ON prv.function_id = pf.function_id
WHERE pf.name LIKE 'PF%tf_bst_hr_25'   -- <- anpassen
ORDER BY prv.boundary_id;
GO

-- -----------------------------------------------------------------------------
-- Abfrage 3: Zeilen je Partition (zeigt auch leere Partitionen)
-- -----------------------------------------------------------------------------
SELECT p.partition_number,
       p.rows,
       prv.value AS Obergrenze
FROM sys.partitions p
LEFT JOIN sys.indexes i
       ON i.object_id = p.object_id AND i.index_id = p.index_id
LEFT JOIN sys.partition_schemes ps
       ON ps.data_space_id = i.data_space_id
LEFT JOIN sys.partition_range_values prv
       ON prv.function_id = ps.function_id AND prv.boundary_id = p.partition_number
WHERE p.object_id = OBJECT_ID('dbo.tf_bst_hr_25')   -- <- anpassen
  AND p.index_id IN (0,1)
ORDER BY p.partition_number;
GO
