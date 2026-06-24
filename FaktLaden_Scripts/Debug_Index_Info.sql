/* ================================================================================
   Debug_Index_Info.sql
   Paket : Fakten Laden (SSIS)
   Zweck : Diagnose-Abfragen zu Indizes einer Faktentabelle. Dient zur Klaerung,
           warum ein Partitionstausch (SWITCH) mit
             "There is no identical index in source table ... for the index '<idx>'
              in target table ..."
           fehlschlaegt. Ursache ist i.d.R. ein NONCLUSTERED (Rowstore) Index auf
           der Faktentabelle, der NICHT von den Ladeskripten erzeugt wird
           (Konvention der Pipeline: CCI_ / NCCI_ / CI_) und daher auf den
           Staging-Tabellen (_in_/_out_/_tmp_) fehlt.

   Hintergrund:
     - SWITCH verlangt, dass Quelle und Ziel IDENTISCHE Indizes haben
       (Clustered + ALLE Nonclustered) und partition-aligned sind.
     - Die Pipeline repliziert auf Staging nur den CLUSTERED Index
       (Typ aus Parameter FaktenClusteredIndex) sowie optional NCCI
       (Parameter FaktenNccIndex). Ein manuell angelegter Rowstore-NCI
       (z.B. nci_bak_zsk_mow) ist in keinem Parameter hinterlegt.

   Anwendung:
     - @fakt und @idx anpassen, in der Ziel-Datamart-DB ausfuehren.
     - HINWEIS: SQL Server speichert KEIN Erstellungsdatum / keinen Ersteller
       je Index (sys.indexes hat keine create_date-Spalte). "Wer/wann" laesst
       sich daher nur ueber Default Trace / Extended Events / DDL-Audit klaeren,
       sofern vorhanden. Naeherung: create_date/modify_date der Tabelle (Query 3).
   ================================================================================ */

DECLARE @fakt sysname = N'tf_bb_bewa_zkt';     -- Faktentabelle anpassen
DECLARE @idx  sysname = N'nci_bak_zsk_mow';    -- betroffener Index anpassen


-- 1) ALLE Indizes der Tabelle + ob partition-aligned -----------------------------
--    storage_type = PARTITION_SCHEME -> aligned (SWITCH moeglich, wenn auf Staging repliziert)
--    storage_type = ROWS_FILEGROUP   -> NICHT aligned (SWITCH unmoeglich, Index muss
--                                       auf das Partitionsschema neu gebaut werden)
SELECT i.index_id,
       i.name              AS index_name,
       i.type_desc,
       i.is_unique,
       i.is_primary_key,
       i.filter_definition,
       ds.name             AS storage,
       ds.type_desc        AS storage_type
FROM   sys.indexes i
JOIN   sys.data_spaces ds ON ds.data_space_id = i.data_space_id
WHERE  i.object_id = OBJECT_ID('dbo.' + @fakt)
ORDER  BY i.index_id;


-- 2) Spalten des betroffenen Index (Schluessel + INCLUDE) -------------------------
--    Liefert die Definition, um den Index 1:1 auf den Staging-Tabellen nachzubauen.
SELECT ic.key_ordinal,
       c.name              AS column_name,
       ic.is_included_column,
       ic.is_descending_key
FROM   sys.index_columns ic
JOIN   sys.columns c ON c.object_id = ic.object_id AND c.column_id = ic.column_id
WHERE  ic.object_id = OBJECT_ID('dbo.' + @fakt)
  AND  ic.index_id  = (SELECT index_id FROM sys.indexes
                       WHERE object_id = OBJECT_ID('dbo.' + @fakt) AND name = @idx)
ORDER  BY ic.is_included_column, ic.key_ordinal;


-- 3) Tabellen-Kontext (Erstell-/Aenderungsdatum) ----------------------------------
SELECT name, create_date, modify_date
FROM   sys.tables
WHERE  object_id = OBJECT_ID('dbo.' + @fakt);


-- 4) Fertige CREATE-Anweisung des betroffenen Index erzeugen ----------------------
--    (zum Nachbauen auf Staging bzw. zum Dokumentieren). Beruecksichtigt
--    Schluesselspalten, INCLUDE-Spalten, UNIQUE und Filter.
SELECT
    'CREATE ' + CASE WHEN i.is_unique = 1 THEN 'UNIQUE ' ELSE '' END
  + i.type_desc COLLATE DATABASE_DEFAULT + ' INDEX [' + i.name + '] ON dbo.[' + @fakt + '] ('
  + STUFF((SELECT ', [' + c.name + ']' + CASE WHEN ic.is_descending_key = 1 THEN ' DESC' ELSE '' END
           FROM sys.index_columns ic
           JOIN sys.columns c ON c.object_id = ic.object_id AND c.column_id = ic.column_id
           WHERE ic.object_id = i.object_id AND ic.index_id = i.index_id AND ic.is_included_column = 0
           ORDER BY ic.key_ordinal
           FOR XML PATH(''), TYPE).value('.', 'nvarchar(max)'), 1, 2, '')
  + ')'
  + ISNULL(' INCLUDE ('
      + STUFF((SELECT ', [' + c.name + ']'
               FROM sys.index_columns ic
               JOIN sys.columns c ON c.object_id = ic.object_id AND c.column_id = ic.column_id
               WHERE ic.object_id = i.object_id AND ic.index_id = i.index_id AND ic.is_included_column = 1
               ORDER BY ic.index_column_id
               FOR XML PATH(''), TYPE).value('.', 'nvarchar(max)'), 1, 2, '')
      + ')', '')
  + ISNULL(' WHERE ' + i.filter_definition, '')
  + ' ON ' + ds.name + ';' AS create_index_stmt
FROM   sys.indexes i
JOIN   sys.data_spaces ds ON ds.data_space_id = i.data_space_id
WHERE  i.object_id = OBJECT_ID('dbo.' + @fakt) AND i.name = @idx;
