CREATE OR REPLACE PACKAGE                          "PKG_EX_WORK_NEWAPI" AUTHID CURRENT_USER AS
-- $Id: create_PKG_EX_WORK_NEWAPI.sql 1628 2024-09-18 09:29:23Z seideld006 $
-- -----------------------------------------------------------------------------
--@START DOKU
-- -----------------------------------------------------------------------------
-- MODUL
-- Revision     $LastChangedRevision: 1628 $
-- Last Revised $LastChangedDate: 2024-09-18 11:29:23 +0200 (Mi, 18 Sep 2024) $
-- Author       $LastChangedBy: seideld006 $
-- -----------------------------------------------------------------------------
-- DATEI
--  File         $HeadURL: https://svn.sdst.sbaintern.de/put/trunk/Anwendungen/DMExport/src/stat_exp_datamart_mcfg/sql/packages/create_PKG_EX_WORK_NEWAPI.sql $
--
-- SYNTAX
--
-- BENOETIGT
-- Tabelle WRK_EX_OBJ
--         WRK_EX_JOB
--
--
-- BESCHREIBUNG
--  Bereitet Daten in der Tabelle WRK_EX_OBJ auf, erzeugt aus den Auftraegen
--  konkrete Export-Jobs in WRK_EX_JOB. Es werden reguläre Ausdrücke und Views
--  aufgelöst.
--  Exportiert die Jobs in die entsprechenden Verrzeichnisse.
--  Die exportierten Verzeichnisse werden gezippt.
--
--  04.03.2015 CebucL - Initiale Version
--  15.06.2015 CebucL - Erweiterungen, Houskeeping regulaere Ausdruecke
--  20.11.2015 CebucL - Erweiterungen, IAB-Export
--  25.01.2016 CebucL - Erweiterungen, Houskeeping u. Bugfixing
--  27.01.2016 CebucL - Kein Datumssuffix bei IAB-Export
--  17.06.2016 CebucL - Bugfixing Errorhandling u. PLSQL-Cursor (LIMIT 100)
--  30.06.2016 CebucL - Änderung bei der Interpretation von SteurlisteneintrÃ¤gen
--                      Bei Einträgen der Form "objektname_YYYYMM" erst als
--                      ganzes Objekt suchen, dann erst die Aufteilung in Objekt
--                      und Partition
--  01.08.2016 CebucL - Create Statement bei Partitionierten Tabellen mit MON_ID
--                      geändert
--  30.09.2016 CebucL - Regulären Ausdruck bei Verarbeitung von '*' geändert
--  20.12.2016 CebucL - Einführung STL_KZ = 'v' neues Kennzeichen. Bewirkt bei
--                      Views das MONID's wie Partitionen behandelt werden.
--                      ansonsten ist das Verhalten wie bei STL_KZ = 'f'
--  03.04.2017 CebucL - Bugfixing bei Patitonierungsschlüssel vom Typ Date
--  28.07.2017 CebucL - Variabeln Ergänzung für ISTAT
--  04.10.2017 CebucL - Änderung STL_KZ = 'v' Exporte der Views mit Prefix 've_'
--                      werden umbenannt in 'te_'
--                    - Änderung Statuslisste:
--                         Name wurde mit Zeitstempel u. Schemaname ergänzt
--                         Inhalt wurde mit MONID bei part. Objekten ergänzt
--                    - Housekeeping ans Ende der Ausführung
--                    - Logging reduziert
--  22.12.2017 CebucL - Erweiterung IAB-Export * bei Schemata erlauben
--                    - Perfonmance-Bug bei Jobs erzeugen
--  04.05.2018 CebucL/SeidelD006
--                    - Behandlung von Partitionen mit NULL-Values hinzugefügt
--  12.07.2018 SeidelD006
--                    - Erkennung des Partitionssuffixes verfeinert:
--                    - Fehlerhafte Erkennung bei vf_fst_kern_5p_all eliminiert
--  23.07.2018 CebucL - In dumpjobddl Else-Zweig hinzugefügt für IAB-Export von
--                      Views auf Fakten
--  12.09.2018 SeidelD006
--                    - Ausführung von Unix-Kommandos jetzt über den Oracle-Scheduler
--                      Hierzu gab es einige Anpassungen im Code
--  17.10.2018 SeidelD006
--                    - In InsPViewJob wird jetzt auch die MOW_ID berücksichtigt
--                      Falls in einem View MON_ID und MOW_ID verwendet werden,
--                      so hat die MON_ID Priorität
--  30.10.2018 SeidelD006
--                    - Das Suchmuster _* am Ende einer Zeile führte dazu,
--                      das der Rest vor dem _* als partitionierte Tabelle
--                      gewertet wurde. Das führte insbesondere beim IAB-Export
--                      zur Nichtbeachtung von nichtpartitionierten Tabellen.
--                      Die Auflösung des Suchmusters erfolgt jetzt zweiteilig,
--                      einmal als partitionierte Tabelle und zum anderen als
--                      normale Tabelle.
-- 31.10.2018 SeidelD006
--                    - Aus Performancegründen wird zur Erzeugung der Semaphoren
--                      ein einziger Unix-Aufruf in Form eines Scriptes verwendet.
--                      Bisher wurden per Semaphore bis zu drei Unix-Calls verwendet.
-- 12.06.2019 SeidelD006
--                    - Platzhalter MONID und Referenzdatum-Parameter eingeführt.
-- 22.07.2019 SeidelD006
--                    - Partitionserkennung funktionierte erst ab 2000*
-- 15.10.2019 SeidelD006
--                    - Platzhalter YEAR eingeführt.
-- 23.10.2019 SeidelD006
--                    - In exportJobs Cursor-For-Loop gegen
--                      Einzelsatzabfrage ausgetauscht, weil es häufiger
--                      bei Langläufern zu ORA-01555 Snapshot too old
--                      Fehlern gekommen ist.
-- 21.08.2020 SeidelD006
--                    - Beim Materialisieren von partitionierten Tabellen wurde
--                      kein PartitionPruning eingesetzt. Das wird mit dieser
--                      Version korrigiert.
-- 09.10.2020 SeidelD006
--                    - Die Semaphorendatei *.unl.ok werden zeitnah nach dem
--                      Entladen eines Jobs erzeugt. Hierdurch kann durch
--                      OPS22 der Transfer gestartet werden, bevor der komplette
--                      Unload abgeschlossen ist.
--  19.01.2021 SeidelD006
--                    - Referenzen auf PKG_EX_WORK auf PKG_EX_WORK_NEWAPI geändert
--                    - Neue Funktionalität fopen mit Batchanlage von leeren Dateien
--                      damit die Penalty beim FOPEN durch den Scheduler-Einsatz
--                      nahezu eliminiert wird.
--  04.03.2021 SeidelD006
--                    - Fehlerbehebung ORA-6502 bei der Verwendung von Filtern
--                      BIDW-408
--  10.03.2021 SeidelD006
--                    - Optimierung gemäß BIDW-409 bei vollqualifizierten
--                      Zeitscheiben
--                    - Fehlerbehebung im Housekeeping (BIDW-403)
--                      Löschen der Altdaten aus den Tabellen gemäß
--                      der Foreign-Key-Abhängigkeiten
--  28.05.2021 SeidelD006
--                    - Korrektur Verarbeitung IAB-Ausschlussliste (BIDW-415)
--  11.06.2021 SeidelD006
--                    - zusätzlicher Filter auf DDL-Generierung, um
--                      Parallel-Clause zu eliminieren (BIDW-418)
--  18.08.2021 SeidelD006
--                    - Entfernung überflüssiger vJOB_PART-Bedingung
--                      wenn Klausel gefüllt ist (BIDW-424)
--  22.10.2021 SeidelD006
--                    - Maximale Parallelität auf 96 gedeckelt (BIDW-436)
--                    - Umstellung auf direktes Schreiben in Datei ohne Puffer
--                      in dumpJob2CVSPipe (BIDW-436)
--  02.11.2021 SeidelD006
--                    - Fehlerbehebung beim Reload von PERM-Steuerlisten (BIDW-440)
--  15.11.2021 SeidelD006
--                    - Anpassungen an 128Zeichen lange Bezeichner (BIDW-441)
--  02.12.2021 SeidelD006
--                    - Austausch DBMS_LOCK.SLEEP gegen DBMS_SESSION.SLEEP
--  05.04.2022 SeidelD006
--                    - BIDW-455; Fehlermeldung, falls alle Einträge der Steuerliste fehlerhaft sind
--                    - BIDW-448; Redundanter äußerer Query bei Partitionsermittlung von partitioned Views
--                      entfernt
--                    - Neue Prozedur checkUnixAPI, damit DM-Export nicht ohne funktionierendes Unix-API
--                      startet
--
--  11.05.2022 SeidelD006
--                    - BIDW-457 Steuerlistentyp "f", Tabellen-DDL an Stelle von View-DDL erzeugen
--  31.03.2023 SeidelD006
--                    - BIDW-471 Verarbeitung von IAB-Steuerlisten deaktiviert
--  21.06.2024 SeidelD006
--                    - BIDW-488 Impl. Schalter für Entladen in UTF8
--  18.09.2023 SeidelD006
--                    - BIDW-495 Materialisierung als GTT

--@ENDE DOKU
-- -----------------------------------------------------------------------------

/*------------------------------------------------------------------------------
Name:         startWork

Parameter: p_STL_NAME  Name der Steuerliste die exportiert werden soll.
                       (ohne Pfad)
           p_RELOAD    Steuert den Ablauf, Normal- oder Wiederanlauf
                       (Default FALSE also KEIN Wiederanlauf)
           p_PARALLEL  Steuert die Paralleliaet mit welcher der DM-Export
                       ausgeführt wird. (Default NULL bedeutet dass die
                       Parallelität automatisch ermittelt wird)
           p_DELIM     Trennzeichen zwischen Spalten (Default "|" erlaubt sind:
                       chr(xxx), ;, |, ...)
           p_PIPE      Steuert, ob der Export über Pipes läuft oder über Files
                       (Default ist über Pipes)
           p_VIEWOBJ   Steuert ob Views, bei Dimensionen, Materialisiert
                       exportiert werden oder ob die Views in Einzelobjejte
                       aufgelöst werden.(Default TRUE also Dimensions-Views
                       werden materialisiert)
           p_EXP       Steuert ob exportiert werden soll oder ob nach dem
                       aufarbeiten der Steuerliste aufgehört werden soll.
                       Wird nur für Test- und Wartungszweck benötigt.
                       (Default TRUE also vollständig exportiert)
           p_REFDATE   Referenzdatum für MONID-Ermittlung,Default ist SYSDATE

           p_DO_UTF8_UNLOOAD
                       Default:false, True: Unload erfolgt in UTF8

Beschreibung: Prozedur startet die Verarbeitung.
              - Erzeugt die benötigten Verzeichnisse Falls nicht vorhanden
              - Räumt alle Dateien weg, die älter als 90 Tage sind
              - Lädt die Steuerliste
              - Erzeugt die Export-Jobs
              - Erzeugt die Verzeichnisse für die Jobs falls nicht vorhanden
              - Exportiert die Jobs
              - Statusliste Semaphore und Protokoll werden erzeugt.
------------------------------------------------------------------------------*/
procedure startWork( p_STL_NAME VARCHAR2,
                     p_RELOAD   BOOLEAN DEFAULT FALSE,
                     p_PARALLEL INTEGER DEFAULT NULL,
                     p_DELIM    VARCHAR2 DEFAULT 'chr(124)',
                     p_PIPE     BOOLEAN DEFAULT TRUE,
                     p_VIEW_MAT BOOLEAN DEFAULT TRUE,
                     p_EXP      BOOLEAN DEFAULT TRUE,
                     p_REFDATE  DATE DEFAULT SYSDATE,
                     p_DO_UTF8_Unload BOOLEAN DEFAULT FALSE);


procedure testSemaphore(pSession in number);

END PKG_EX_WORK_NEWAPI;
/


CREATE OR REPLACE PACKAGE BODY                          "PKG_EX_WORK_NEWAPI"
IS
-- $Id: create_PKG_EX_WORK_NEWAPI.sql 1628 2024-09-18 09:29:23Z seideld006 $
-- $HeadURL: https://svn.sdst.sbaintern.de/put/trunk/Anwendungen/DMExport/src/stat_exp_datamart_mcfg/sql/packages/create_PKG_EX_WORK_NEWAPI.sql $


vVerbose boolean := false;

vEX_SESSION  NUMBER                  := STAT_EXP_DATAMART_MCFG.SEQ_EX_SESSION.nextval;
vCurrentJobId NUMBER                 := NULL;
vDataSchema  constant VARCHAR2(128)   := 'STAT_EXP_DATAMART_MWRK';

vPathESTAT   constant VARCHAR2(2000) := '/batches/stat_e/work/DM_EXPORT';
vPathFSTAT   constant VARCHAR2(2000) := '/batches/stat_f/work/DM_EXPORT';
vPathPSTAT   constant VARCHAR2(2000) := '/batches/stat_p/work/DM_EXPORT';
vPathLSTAT   constant VARCHAR2(2000) := '/batches/stat_l/work/DM_EXPORT';
vPathASTAT   constant VARCHAR2(2000) := '/batches/stat_a/work/DM_EXPORT';
vPathISTAT   constant VARCHAR2(2000) := '/batches/stat_i/work/DM_EXPORT';

-- Parallelität mit welcher exportiert werden soll
vParallel    INTEGER                 := NULL;

-- steuert ob der Export ueber Pipes läuft oder über Files
vPIPE        BOOLEAN                 := TRUE;

-- steuert ob bei Dimensionen Views aufgelöst werden oder nicht (TRUE auflösen).
vVIEW_MAT    BOOLEAN                 := TRUE;

vDelimiter   VARCHAR2(8)             := 'chr(124)';

vDoUTF8Unload BOOLEAN                := FALSE;
vUTF8CharSet constant varchar2(100)  := 'AL32UTF8';
vDefaultCharSet constant varchar2(100):='WE8ISO8859P1';

vStarReplacement VARCHAR2(10) := '.{0,128}';
-- nach 99 Tagen werden alle erzeugte Dateien und Arbeitsdaten gelöscht
vAnzTageHK   constant INTEGER        := 99;

-- nach 40 Tagen werden slle erzeugte Dateien und Verzeichnisse des IAB-Exports gelöscht
vAnzTageHK_IAB   constant INTEGER    := 40;

-- Rechte generierte Datenfiles
vMODE        VARCHAR2(3)             := '666';
-- Rechte generierte Skripte
vMODE_EXE    VARCHAR2(3)             := '774';
-- Rechte Verzeichnisse
vMODE_DIR    VARCHAR2(3)             := '777';

vRefDate     date                   := sysdate;

-- Forward-Deklaration
procedure createSemaphoresJob(pJobId in number);

/*------------------------------------------------------------------------------
Name:         start2regex

Parameter:    String

Return:       Eingabestring ohne *, aber mit regulärem Ausdruck

Beschreibung: Ersetzt * durch regulären Ausdruck
------------------------------------------------------------------------------*/
  FUNCTION star2regex(pIn in varchar2) return varchar2
  as
  begin
    return replace(pIn,'*',vStarReplacement);
  end;



/*------------------------------------------------------------------------------
Name:         getMainPath

Parameter:    keine .

Return:       Hauptpfad

Beschreibung: Ermittelt das "root" des Pfades für die Dateien aus dem
              Datenbanknamen
------------------------------------------------------------------------------*/
  FUNCTION getMainPath
    RETURN VARCHAR2
  AS
vPath  VARCHAR2(2000);
BEGIN

  SELECT
    CASE
      WHEN upper(NAME) LIKE 'ESTAT%'
      THEN vPathESTAT
      WHEN upper(NAME) LIKE 'FSTAT%'
      THEN vPathFSTAT
      WHEN upper(NAME) LIKE 'PSTAT%'
      THEN vPathPSTAT
      WHEN upper(NAME) LIKE 'LSTAT%'
      THEN vPathLSTAT
      WHEN upper(NAME) LIKE 'ASTAT%'
      THEN vPathASTAT
      WHEN upper(NAME) LIKE 'ISTAT%'
      THEN vPathISTAT
      ELSE NULL
    END
  INTO vPath
  FROM (select UPPER(sys_context('USERENV','DB_NAME')) as NAME from DUAL);

  RETURN vPath;
EXCEPTION
  WHEN OTHERS THEN
    pkg_lib_ex_basic_newapi.WRITELOG(P_EVENTTYPE      => pkg_lib_ex_basic_newapi.EventT_SysError,
                              P_SESSION        => vEX_SESSION,
                              P_IDENTIFICATION => $$PLSQL_UNIT,
                              P_TXT            => 'Systemfehler');
END getMainPath;




/*------------------------------------------------------------------------------
Name:         getMaxCPU

Parameter:    keine

Return:       Anzahl CPUs

Beschreibung: Ermittels die Anzahl der CPUs die der DB zur Verfügung stehen.
------------------------------------------------------------------------------*/
  FUNCTION getMaxCPU
    RETURN INTEGER
  AS
vMax  INTEGER;
BEGIN
   -- CPU-Anzahl auf 96 limitieren
   SELECT least(VALUE,96) INTO vMax FROM V$OSSTAT WHERE STAT_NAME = 'NUM_CPUS';

  RETURN vMax;
EXCEPTION
  WHEN OTHERS THEN
    pkg_lib_ex_basic_newapi.WRITELOG(P_EVENTTYPE      => pkg_lib_ex_basic_newapi.EventT_SysError,
                              P_SESSION        => vEX_SESSION,
                              P_IDENTIFICATION => $$PLSQL_UNIT,
                              P_TXT            => 'Systemfehler');
END getMaxCPU;


/*------------------------------------------------------------------------------
Name:         logInfo

Parameter:    beliebiger Text

Beschreibung: Der übergebene Text wird noch ergänzt und es wird ein Log-Eintrag
              daraus erzeugt..
------------------------------------------------------------------------------*/
procedure logInfo (pText VARCHAR2) IS
BEGIN
  pkg_lib_ex_basic_newapi.WRITELOG(P_EVENTTYPE      => pkg_lib_ex_basic_newapi.EventT_Information,
                            P_SESSION        => vEX_SESSION,
                            P_JOB_ID         => vCurrentJobId,
                            P_IDENTIFICATION => $$PLSQL_UNIT,
                            P_TXT            => pText);
EXCEPTION
  WHEN OTHERS THEN
    pkg_lib_ex_basic_newapi.WRITELOG(P_EVENTTYPE      => pkg_lib_ex_basic_newapi.EventT_SysError,
                              P_SESSION        => vEX_SESSION,
                              P_JOB_ID         => vCurrentJobId,
                              P_IDENTIFICATION => $$PLSQL_UNIT,
                              P_TXT            => 'Systemfehler');
END logInfo;


/*------------------------------------------------------------------------------
Name:         logVerbose

Parameter:    beliebiger Text

Beschreibung: Der übergebene Text wird noch ergänzt und es wird ein Log-Eintrag
              daraus erzeugt..
------------------------------------------------------------------------------*/
procedure logVerbose (pText VARCHAR2) IS
BEGIN

    if vVerbose then
        pkg_lib_ex_basic_newapi.WRITELOG(P_EVENTTYPE      => pkg_lib_ex_basic_newapi.EventT_Information,
                            P_SESSION        => vEX_SESSION,
                            P_JOB_ID         => vCurrentJobId,
                            P_IDENTIFICATION => $$PLSQL_UNIT,
                            P_TXT            => pText);
    end if;

EXCEPTION
  WHEN OTHERS THEN
    pkg_lib_ex_basic_newapi.WRITELOG(P_EVENTTYPE      => pkg_lib_ex_basic_newapi.EventT_SysError,
                              P_SESSION        => vEX_SESSION,
                              P_JOB_ID         => vCurrentJobId,
                              P_IDENTIFICATION => $$PLSQL_UNIT,
                              P_TXT            => 'Systemfehler');
END logVerbose;




/*------------------------------------------------------------------------------
Name:         delFilesFromDir

Parameter:    pDirName   Name eines Directory-Objekts

Beschreibung: Löscht alle Dateien aus dem Pfad des Directory-Objekts die
              älter als 90 Tage sind.
              ACHTUNG löscht auch in Unterverzeichnissen.
------------------------------------------------------------------------------*/
procedure delFilesFromDir (pDirName VARCHAR2, pAnzTageHK INTEGER) IS

vPath           VARCHAR2(2000);
vSQL            VARCHAR2(4000);
vUnixCom        VARCHAR2(4000);
vPathExists     INTEGER;
BEGIN

  vPath       := pkg_lib_ex_basic_newapi.getPath(pDirName);
  vPathExists := dbms_lob.fileexists(bfilename(pDirName, '.'));

  IF vPathExists <> 0 THEN

    IF pAnzTageHK <> 0 THEN
     logInfo('Löschen aller Dateien älter als ' || pAnzTageHK || ' aus dem Pfad: ' || vPath );
     vUnixCom := 'find '|| vPath || ' -type f  -a ! -iname *_perm.csv -a ! -iname login.sql  -mtime +' || pAnzTageHK || ' -delete';
    ELSE
     logInfo('Löschen aller Dateien aus dem Pfad: ' || vPath );
     vUnixCom := 'find '|| vPath || ' -type f -iname * -delete';
     logInfo('Löschen aller Dateien aus dem Pfad: ' || vPath );
     vUnixCom := 'find '|| vPath || ' -type p -iname * -delete';
    END IF;

    pkg_lib_ex_basic_newapi.CALLUNIX(vUnixCom, vEX_SESSION);

  ELSE
     logInfo('Der Pfad: ' || vPath  || 'existiert nicht' );
  END IF;

EXCEPTION
  WHEN OTHERS THEN
    pkg_lib_ex_basic_newapi.WRITELOG(P_EVENTTYPE      => pkg_lib_ex_basic_newapi.EventT_SysError,
                              P_SESSION        => vEX_SESSION,
                              P_JOB_ID         => vCurrentJobId,
                              P_IDENTIFICATION => $$PLSQL_UNIT,
                              P_TXT            => 'Systemfehler');
END delFilesFromDir;

/*
FUNCTION fopen(location in varchar2,filename in varchar2) return UTL_FILE.FILE_TYPE
AS
l_file      UTL_FILE.FILE_TYPE;
lPath       VARCHAR2(4000);
lUnixCom    VARCHAR2(4000);
BEGIN
  lPath := pkg_lib_ex_basic_newapi.getPath(location);
  lUnixCom :=     'rm '|| lPath || '/' || Filename||' 2>/dev/null'||chr(10)
                ||'touch ' || lPath || '/' || Filename||chr(10)
                ||'chmod ' || vMODE || ' ' || lPath || '/' || Filename;
  pkg_lib_ex_basic_newapi.CALLUNIX(lUnixCom, vEX_SESSION);

  l_file := UTL_FILE.FOPEN(
            location => location,
            filename => Filename,
            open_mode => 'a',
            max_linesize => 32767);
  return l_file;
END;
*/
FUNCTION fopen(location in varchar2,filename in varchar2) return UTL_FILE.FILE_TYPE
AS
BEGIN
  return pkg_lib_ex_basic_newapi.fopen(pSession=>vEX_SESSION,pTempDir=>'EX_TEMP',pMode=>vMODE,pLocation=>location,pFileName=>filename);
END;

/*------------------------------------------------------------------------------
Name:         doHousekeeping

Parameter:    keine

Beschreibung: Prozedur löscht Sequences und Tabellen die von fehlerhaften
              Läufen stammen.
              Löscht aus dem Dateisysten Dateien älter als 100 Tage.
------------------------------------------------------------------------------*/
procedure doHousekeeping IS

vDelDate        DATE;
vSQL            VARCHAR2(4000);
vUnixCom        VARCHAR2(4000);
vPathExists     INTEGER;

BEGIN
    --logInfo('Start Housekeeping');
    commit;
    -- Datum ermitteln ab dem gelöscht werden soll
    vSQL := 'SELECT SYSDATE - INTERVAL ' || chr(39) || vAnzTageHK || chr(39) || ' DAY FROM dual';
    execute immediate vSQL INTO vDelDate;

    -- Nicht mehr benöigte PreCreateFiles löschen
    pkg_lib_ex_basic_newapi.prc_RemovePreCreateFiles(vEX_SESSION,'EX_TEMP');


    logInfo('Löschen von Sequences, die wegen fehlerhaften Läufen nicht gelöscht wurden');
    -- fuer alle nicht gelöschten Sequences älter als vAnzTageHK Tage
    FOR cSeq IN
    (SELECT SEQUENCE_NAME
       FROM dba_sequences, dba_objects
      WHERE SEQUENCE_OWNER = USER
        AND SEQUENCE_NAME LIKE 'SEQ_EXP%'
        AND SEQUENCE_OWNER = OWNER
        AND SEQUENCE_NAME = OBJECT_NAME
        AND CREATED < vDelDate
    )
    LOOP
      vSQL :='DROP sequence ' || cSeq.SEQUENCE_NAME;
      execute immediate vSQL;
    END LOOP;

    logInfo('Löschen von Tabellen, die wegen fehlerhafter Läufe nicht gelöscht wurden');
    -- fuer alle nich geloeschten Tabellen aelter als vAnzTageHK Tage
    FOR cTab IN
    (SELECT t.OWNER, t.TABLE_NAME
       FROM dba_tables t, dba_objects o
      WHERE t.OWNER = o.OWNER
        AND t.TABLE_NAME = o.OBJECT_NAME
        AND t.OWNER = vDataSchema
        AND CREATED < vDelDate
    )
    LOOP
      vSQL :='DROP TABLE ' || cTab.OWNER || '.' || cTab.TABLE_NAME;
      execute immediate vSQL;
    END LOOP;

    logInfo('Löschen von Dateien älter als ' || vAnzTageHK || ' Tage aus dem Dateisystem');
    -- Alle Dateien älter als vAnzTageHK Tage aus dem Dateisystem löschen
    delFilesFromDir ('EX_EXP', vAnzTageHK);

    logInfo('Löschen von Daten älter als ' || vAnzTageHK || ' Tage aus den Arbeitstabellen');
    -- Alle Daten älter als vAnzTageHK Tage aus den Arbeitstabellen löschen
    -- dabei die Abhängigkeiten beachten . . .
    DELETE FROM STAT_EXP_DATAMART_MCFG.WRK_EX_JOB
    WHERE JOB_SESSION IN (SELECT DISTINCT STL_SESSION FROM STAT_EXP_DATAMART_MCFG.WRK_EX_OBJ WHERE stl_zbdat < vDelDate);
    COMMIT;
    DELETE  FROM STAT_EXP_DATAMART_MCFG.ERR_EX_PROT
    WHERE EX_SESSION IN (SELECT DISTINCT STL_SESSION FROM STAT_EXP_DATAMART_MCFG.WRK_EX_OBJ WHERE stl_zbdat <vDelDate);
    COMMIT;
    DELETE FROM STAT_EXP_DATAMART_MCFG.WRK_EX_OBJ
    WHERE STL_SESSION IN (SELECT DISTINCT STL_SESSION FROM STAT_EXP_DATAMART_MCFG.WRK_EX_OBJ WHERE stl_zbdat < vDelDate);
    COMMIT;

/*  rausgenommen weil auch Datumssuffix bei IAB-Export rausgenommen wurde
    logInfo('Loeschen von IAB-Verzeichnissen(mit Inhalt) ' || vAnzTageHK_IAB || ' Tage aus den Dateisystem');
    -- Alle IAB-Verzeichnissen mit Dateien aelter als vAnzTageHK_IAB Tage aus dem Dateisystem loeschen
    vSQL := 'SELECT SYSDATE - INTERVAL ' || chr(39) || vAnzTageHK_IAB || chr(39) || ' DAY FROM dual';
    execute immediate vSQL INTO vDelDate;
    FOR cDir IN
    (SELECT *
       FROM DBA_DIRECTORIES
      WHERE DIRECTORY_NAME LIKE 'EX_IAB_%'
        AND to_date(substr(DIRECTORY_NAME, 8, 13 ) , 'YYMMDD') < vDelDate
    )
    LOOP
      -- alle Dateien des gefundenen Verrzeichisses loeschen, egal welches datum
      delFilesFromDir (cDir.DIRECTORY_NAME, 0);

      vPathExists := dbms_lob.fileexists(bfilename(cDir.DIRECTORY_NAME, '.'));

      IF vPathExists <> 0 THEN
        vUnixCom := 'rmdir '|| cDir.DIRECTORY_PATH;
        pkg_lib_ex_basic_newapi.CALLUNIX(vUnixCom,vEX_SESSION);
      END IF;

      vSQL :='DROP directory ' || cDir.DIRECTORY_NAME;
      execute immediate vSQL;
    END LOOP;
*/
    logInfo('Löschen von Export-Verzeichnissen, die länger als ' || vAnzTageHK || ' Tage nicht genutzt wurden');
    -- fuer alle Export-Verzeichnisse die nicht mehr in WRK_EX_OBJ vorkommen
    FOR cEXDir IN
    (SELECT *
        FROM dba_directories
        WHERE DIRECTORY_PATH LIKE '%/DM_EXPORT/%'
        AND SUBSTR(DIRECTORY_NAME, 3, LENGTH(DIRECTORY_NAME)) NOT IN
          (SELECT DISTINCT REPLACE(upper(STL_SCHEMA),'_', '') AS DIR_Name
          FROM STAT_EXP_DATAMART_MCFG.WRK_EX_OBJ
          )
        AND (DIRECTORY_NAME LIKE 'XE%'
        OR DIRECTORY_NAME LIKE 'XF%'
        OR DIRECTORY_NAME LIKE 'XD%')
    )
    LOOP
      vSQL :='DROP DIRECTORY ' || cEXDir.DIRECTORY_NAME;
      execute immediate vSQL;
    END LOOP;


    logInfo('Ende Housekeeping');
EXCEPTION
  WHEN OTHERS THEN
    pkg_lib_ex_basic_newapi.WRITELOG(P_EVENTTYPE      => pkg_lib_ex_basic_newapi.EventT_SysError,
                              P_SESSION        => vEX_SESSION,
                              P_IDENTIFICATION => $$PLSQL_UNIT,
                              P_TXT            => 'Systemfehler');
END doHousekeeping;

/*------------------------------------------------------------------------------
Name:         setDelimiter

Parameter:    zu setzender Delimiter (erlaubt sind: chr(xxx), ;, |, ...)

Beschreibung: Prozedur setzt und maskiert den Delimiter.
------------------------------------------------------------------------------*/
procedure setDelimiter (p_DELIM    VARCHAR2) IS
BEGIN
  IF upper(substr(p_DELIM,1,3))= 'CHR' THEN
     vDelimiter := p_DELIM;
  ELSE
     vDelimiter :=  chr(39) || p_DELIM || chr(39);
  END IF;
EXCEPTION
  WHEN OTHERS THEN
    pkg_lib_ex_basic_newapi.WRITELOG(P_EVENTTYPE      => pkg_lib_ex_basic_newapi.EventT_SysError,
                              P_SESSION        => vEX_SESSION,
                              P_IDENTIFICATION => $$PLSQL_UNIT,
                              P_TXT            => 'Systemfehler');
END setDelimiter;

/*------------------------------------------------------------------------------
Name:         setStatusOBJ

Parameter:    pStlID      ID eines Steuerlisten-Objekts
              pStlStatus  Status Steuerlisten-Objekts

Beschreibung: Prozedur setzt den gewünschten Status eines Steuerlisten-Objekts.
------------------------------------------------------------------------------*/
procedure setStatusOBJ (pStlID STAT_EXP_DATAMART_MCFG.WRK_EX_OBJ.STL_ID%TYPE,
                       pStlStatus STAT_EXP_DATAMART_MCFG.WRK_EX_OBJ.STL_STATUS%TYPE,
                       pForce BOOLEAN DEFAULT FALSE) IS
vJOB_ID NUMBER(20,0);
vSQL VARCHAR2(4000);

BEGIN
  IF pForce THEN
    UPDATE STAT_EXP_DATAMART_MCFG.WRK_EX_OBJ
       SET STL_STATUS = pStlStatus
     WHERE STL_ID = pStlID
       AND STL_SESSION = vEX_SESSION;
  ELSE
    UPDATE STAT_EXP_DATAMART_MCFG.WRK_EX_OBJ
       SET STL_STATUS = pStlStatus
     WHERE STL_ID = pStlID
       AND STL_STATUS <> 'F'
       AND STL_SESSION = vEX_SESSION;
  END IF;
EXCEPTION
  WHEN OTHERS THEN
    pkg_lib_ex_basic_newapi.WRITELOG(P_EVENTTYPE      => pkg_lib_ex_basic_newapi.EventT_SysError,
                              P_SESSION        => vEX_SESSION,
                              P_IDENTIFICATION => $$PLSQL_UNIT,
                              P_TXT            => 'Systemfehler');
END setStatusOBJ;

/*------------------------------------------------------------------------------
Name:         getJobError

Parameter:    pJobID   id eines generierten Jobs

Return:       Text Fehlermeldung

Beschreibung: Funktion ermittelt für eine JOB_ID eine Fehlermeldung, falls
              vorhanden.
------------------------------------------------------------------------------*/
FUNCTION getJobError(pJobID STAT_EXP_DATAMART_MCFG.WRK_EX_JOB.JOB_ID%TYPE)
    RETURN STAT_EXP_DATAMART_MCFG.ERR_EX_PROT.TEXT%TYPE
AS
--vIdent STAT_EXP_DATAMART_MCFG.ERR_EX_PROT.IDENTIFICATION%TYPE;
vError VARCHAR2(32000);
BEGIN

 -- vIdent := $$PLSQL_UNIT || '_JOB_ID_' || pJobID;
  --dbms_output.put_line(vIdent);
  FOR cError IN ( SELECT TEXT
                    FROM STAT_EXP_DATAMART_MCFG.ERR_EX_PROT
                   WHERE JOB_ID = pJobId
                     AND EVENT_TYPE IN  ( 'E', 'S')
                 )
  LOOP
    vError := vError || chr(10) || cError.TEXT;
  END LOOP;
  RETURN vError;
EXCEPTION
  WHEN NO_DATA_FOUND THEN
    NULL;
  WHEN OTHERS THEN
    pkg_lib_ex_basic_newapi.WRITELOG(P_EVENTTYPE      => pkg_lib_ex_basic_newapi.EventT_SysError,
                             P_SESSION        => vEX_SESSION,
                              P_IDENTIFICATION => $$PLSQL_UNIT,
                              P_TXT            => 'Systemfehler');
  END getJobError;

/*------------------------------------------------------------------------------
Name:         getPosZeichen

Parameter:    pLine      Name des Directory-Objekts in welchem die STL sind.

Return:       Postition des gefundenen Musters z.B '*'

Beschreibung: Funktion prüft ob der Eingabestring einem gewissen Muster
              entspricht
------------------------------------------------------------------------------*/
  FUNCTION getPosZeichen(
      pLine        IN VARCHAR2,
      pRegZeichen  IN VARCHAR2)
    RETURN INTEGER
  AS
    vPos         INTEGER;
  BEGIN

    -- finden '*'
    SELECT INSTR(pLine, pRegZeichen)
    INTO vPos
    FROM DUAL;

    RETURN vPos;
  EXCEPTION
  WHEN no_data_found THEN
     RETURN vPos;
  WHEN OTHERS THEN
    pkg_lib_ex_basic_newapi.WRITELOG(P_EVENTTYPE      => pkg_lib_ex_basic_newapi.EventT_SysError,
                              P_SESSION        => vEX_SESSION,
                              P_IDENTIFICATION => $$PLSQL_UNIT,
                              P_TXT            => 'Systemfehler');
  END getPosZeichen;

FUNCTION MonId (p_expr IN VARCHAR2,p_start IN NUMBER default 1) RETURN VARCHAR2
AS
  l_expr VARCHAR2(1000);
  l_split NUMBER;
  l_window NUMBER;
  l_tmp VARCHAR2(1000);
  l_part VARCHAR2(1000);
  l_pos NUMBER;
BEGIN
  logVerbose('MonId('||p_expr ||','||p_start||')');

  l_expr := p_expr;
  l_pos := instr(l_expr,':MONID',p_start);
  if nvl(l_pos,0) = 0
  then
    return l_expr;
  end if;

    -- ${MONID} und ${MONID6} => ${MONID-0}
    --l_expr := regexp_replace(l_expr,'\:\{MONID[6]{0,1})(\})','${MONID-0}');

    -- ${MONID2} => ${MONID4-0}
    --l_expr := regexp_replace(l_expr,'(\$\{MONID4)(\})','${MONID4-0}');

    -- MONID und MONID6 => 4stelliges Jahr
    l_tmp := regexp_substr(l_expr,'(\:MONID[6]{0,1})(\()([-]{0,1}[0-9]+)(\))');

    IF l_tmp IS NOT NULL
    THEN
      l_window:=TO_NUMBER(REGEXP_REPLACE(l_tmp,'(\:MONID[6]{0,1})(\()([-]{0,1}[0-9]+)(\))','\3'));
    ELSE
      l_window := null;
    END IF;

    IF l_window IS NOT NULL
    THEN
      l_part := TO_CHAR(ADD_MONTHS(TRUNC(vRefDate),l_window),'YYYYMM');
      l_expr := REPLACE(l_expr,l_tmp,l_part);
    END IF;

    l_tmp := regexp_substr(l_expr,'(\:MONID4)(\()([-]{0,1}[0-9]+)(\))');

    -- MONID4 => 2stelliges Jahr
    IF l_tmp IS NOT NULL
    THEN
      l_window:=TO_NUMBER(REGEXP_REPLACE(l_tmp,'(\:MONID4)(\()([-]{0,1}[0-9]+)(\))','\3'));
    ELSE
      l_window := null;
    END IF;

    IF l_window IS NOT NULL
    THEN
      l_part := TO_CHAR(ADD_MONTHS(TRUNC(vRefDate),l_window),'RRMM');
      l_expr := REPLACE(l_expr,l_tmp,l_part);
    END IF;


  return MonId(l_expr,l_pos+1);
END MonId;

FUNCTION YearId (p_expr IN VARCHAR2,p_start IN NUMBER default 1) RETURN VARCHAR2
AS
  l_expr VARCHAR2(1000);
  l_split NUMBER;
  l_window NUMBER;
  l_tmp VARCHAR2(1000);
  l_part VARCHAR2(1000);
  l_pos NUMBER;
BEGIN
  logVerbose('YearId('||p_expr ||','||p_start||')');

  l_expr := p_expr;
  l_pos := instr(l_expr,':YEAR',p_start);
  if nvl(l_pos,0) = 0
  then
    return l_expr;
  end if;

    -- ${MONID} und ${MONID6} => ${MONID-0}
    --l_expr := regexp_replace(l_expr,'\:\{MONID[6]{0,1})(\})','${MONID-0}');

    -- ${MONID2} => ${MONID4-0}
    --l_expr := regexp_replace(l_expr,'(\$\{MONID4)(\})','${MONID4-0}');

    -- MONID und MONID6 => 4stelliges Jahr
    l_tmp := regexp_substr(l_expr,'(\:YEAR)(\()([-]{0,1}[0-9]+)(\))');

    IF l_tmp IS NOT NULL
    THEN
      l_window:=TO_NUMBER(REGEXP_REPLACE(l_tmp,'(\:YEAR)(\()([-]{0,1}[0-9]+)(\))','\3'));
    ELSE
      l_window := null;
    END IF;

    IF l_window IS NOT NULL
    THEN
      l_part := TO_CHAR(ADD_MONTHS(TRUNC(vRefDate),12*l_window),'YYYY');
      l_expr := REPLACE(l_expr,l_tmp,l_part);
    END IF;

  return YearId(l_expr,l_pos+1);
END YearId;


/*------------------------------------------------------------------------------
Name:      getKlausel

Parameter:    pZusatz      String-Teil aus Steuerliste, die geparst werden soll.

Return:       gibt eine Kausel zum Vergleich mit einem Feld zurück
Beschreibung: Funktion erzeugt eine Kausel zum Vergleich mit einem Feld anhand
              von gewissen Muster des Strings pZusatz
------------------------------------------------------------------------------*/
  FUNCTION getKlausel( pZusatz        IN VARCHAR2,
                       pVergleich     IN VARCHAR2)
    RETURN VARCHAR2
  AS
    vKlausel     VARCHAR2(2000);
    vKommando    VARCHAR2(60);
    vBedinngung  VARCHAR2(2000);
    vPos         INTEGER;
    vAnz         INTEGER;
    lZusatz      VARCHAR2(4000);
    -- LAST_MM, LAST_YYYYMM LAST_YY, LAST_YYMM, LAST_YYYY
    vRegLast     VARCHAR2(60) := ':LAST_[YM]{2,6}\({1}[1-9]{1,2}\){1}$';
    vRegFix      VARCHAR2(60) := ':[YM]{2,6}\({1}.{1,255}\){1}$';
  BEGIN

    --lZusatz := MonId(pZusatz);
    lZusatz := pZusatz;

    vPos := getPosZeichen(lZusatz, '(');

    IF vPos <> 0 THEN
      vKommando    := upper(SUBSTR(lZusatz, 1, vPos - 1));
      vBedinngung  := SUBSTR(lZusatz, vPos, LENGTH(lZusatz));

      CASE
        WHEN vKommando IN ('LAST_YY', 'LAST_YYYY') THEN                         -- Jahr
          SELECT TRIM(CHR(40)  FROM vBedinngung) into vBedinngung FROM dual;    -- (
          SELECT TRIM(CHR(41)  FROM vBedinngung) into vBedinngung FROM dual;    -- )
          vAnz :=  vBedinngung * 12;                                -- n Jahre
          vKlausel := ' WHERE substr(' || pVergleich || ',1,6) > ' || to_char(add_months(vRefDate, - vAnz), 'YYYYMM');
        WHEN vKommando IN ('LAST_MM', 'LAST_YYYYMM', 'LAST_YYMM') THEN          -- Monat
          SELECT TRIM(CHR(40)  FROM vBedinngung) into vBedinngung FROM dual;    -- (
          SELECT TRIM(CHR(41)  FROM vBedinngung) into vBedinngung FROM dual;    -- )
          vAnz :=  vBedinngung * 1;                                 -- n Monate
          vKlausel := ' WHERE substr(' || pVergleich || ',1,6) > ' || to_char(add_months(vRefDate, - vAnz), 'YYYYMM');
        WHEN vKommando = 'YYYY' THEN                                            -- konkretes Jahr
          vKlausel := ' WHERE substr(' || pVergleich || ',1,4) IN ' || vBedinngung ;
        WHEN vKommando = 'YYYYMM' THEN                                          -- konkrete MON_ID
          vKlausel := ' WHERE substr(' || pVergleich || ',1,6) IN ' || vBedinngung ;
        ELSE
          vKlausel := NULL;
        END CASE;

    ELSE
      vKlausel := NULL;
    END IF;

    RETURN vKlausel;
  EXCEPTION
  WHEN OTHERS THEN
    pkg_lib_ex_basic_newapi.WRITELOG(P_EVENTTYPE      => pkg_lib_ex_basic_newapi.EventT_SysError,
                              P_SESSION        => vEX_SESSION,
                              P_IDENTIFICATION => $$PLSQL_UNIT,
                              P_TXT            => 'Systemfehler');
  END getKlausel;

/*------------------------------------------------------------------------------
Name:      getPosZusatz

Parameter:    pLine      Name des Directory-Objekts in welchem die STL sind.

Return:       Postition des gefundenen Musters z.B 'LAST...'

Beschreibung: Funktion, prüft ob der Eingabestring einem gewissen Muster
              entspricht und gibt die Position zurück
------------------------------------------------------------------------------*/
  FUNCTION getPosZusatz(
      pLine        IN VARCHAR2)
    RETURN INTEGER
  AS
    vPos         INTEGER;
    -- LAST_MM, LAST_YYYYMM LAST_YY, LAST_YYMM, LAST_YYYY
    vRegLast     VARCHAR2(60) := ':LAST_[YM]{2,6}\({1}[1-9]{1,2}\){1}$';
    vRegFix      VARCHAR2(60) := ':[YM]{2,6}\({1}.{1,255}\){1}$';
  BEGIN

    SELECT REGEXP_INSTR(upper(pLine), vRegLast)
    INTO vPos
    FROM DUAL;

    IF vPos = 0 THEN
      SELECT REGEXP_INSTR(upper(pLine), vRegFix)
      INTO vPos
      FROM DUAL;
    END IF;

    RETURN vPos;
  EXCEPTION
  WHEN no_data_found THEN
     RETURN vPos;
  WHEN OTHERS THEN
    pkg_lib_ex_basic_newapi.WRITELOG(P_EVENTTYPE      => pkg_lib_ex_basic_newapi.EventT_SysError,
                              P_SESSION        => vEX_SESSION,
                              P_IDENTIFICATION => $$PLSQL_UNIT,
                              P_TXT            => 'Systemfehler');
  END getPosZusatz;

/*------------------------------------------------------------------------------
Name:      getPosReg

Parameter:    pLine      Name des Directory-Objekts in welchem die STL sind.

Return:       Postition des gefundenen Musters

Beschreibung: Funktion, prüft ob der Eingabestring einem gewissen Muster
              entspricht
------------------------------------------------------------------------------*/
  FUNCTION getPosReg(
      pLine        IN VARCHAR2)
    RETURN INTEGER
  AS
    vPos         INTEGER;
    vPosKlammerL INTEGER;
    vPosKlammerR INTEGER;
    vPosStern    INTEGER;
    vPosStrich   INTEGER;
    vPosFageZ    INTEGER;
    vPosPlus     INTEGER;
    vRegFix4     VARCHAR2(60) := '_[0-9]{2}[0-1]{1}[0-9]{1}$';                     -- z.B 1502     YYMM
    vRegFix6     VARCHAR2(60) := '_(19|20)[0-9]{2}[0-1]{1}[0-9]{1}$';                   -- z.B 201502   YYYYMM
    vRegFix8     VARCHAR2(60) := '_(19|20)[0-9]{2}[0-1]{1}[0-9]{1}[0-3]{1}[0-9]{1}$';   -- z.B 20150212 YYYYMMDD
    vRegSond     VARCHAR2(60) := '_[0-9?+*]{1,16}$';                               -- z.B 1502, 201502 oder 20150212 mit ? und +
    vReg         VARCHAR2(60) := '\[{0,1}[0-9]{0,8}\-{0,1}[0-9]{0,1}\]{0,1}[^A-Z_]+';      -- z.B [0-9] , [12] oder 12 ... kann auch leer sein
    vReg1        VARCHAR2(60) := '_\[{0,1}[0-9]{1,8}\-{0,1}[0-9]{0,1}\]{0,1}[^A-Z_]+';     -- z.B _[0-9] , _[12] oder _12 ... und nicht leer
    vReg8        VARCHAR2(2000);
    vPART        VARCHAR2(2000);

  BEGIN


    -- finden 1502
    SELECT REGEXP_INSTR(pLine, vRegFix4)
    INTO vPos
    FROM DUAL;

    -- finden 201502
    IF vPos = 0 THEN
      SELECT REGEXP_INSTR(pLine, vRegFix6)
      INTO vPos
      FROM DUAL;
    END IF;

    -- finden 20150212
    IF vPos = 0 THEN
      SELECT REGEXP_INSTR(pLine, vRegFix8)
      INTO vPos
      FROM DUAL;
    END IF;

    -- wenn obere nicht gefunden dann  suchen mit "?" , "+"  und "*"in der Partition
    IF vPos = 0 THEN

      SELECT INSTR(pLine, '?') INTO vPosFageZ  FROM DUAL;
      SELECT INSTR(pLine, '+') INTO vPosPlus   FROM DUAL;
      SELECT INSTR(pLine, '*') INTO vPosStern  FROM DUAL;

      -- das ganze nur wenn Sonderzeichen drin sind
      IF vPosFageZ > 0 OR vPosPlus > 0 OR vPosStern > 0 THEN
        SELECT REGEXP_INSTR(pLine, vRegSond)
        INTO vPos
        FROM DUAL;
      END IF;

      -- Länge des Partitionsmusters ohne Sonderzeichen darf nicht größer als 8 sein
      vPART  := SUBSTR(pLine, vPos + 1, LENGTH(pLine));
      vPART  := REPLACE(vPART,'?','');
      vPART  := REPLACE(vPART,'+','');
      vPART  := REPLACE(vPART,'*','');

      IF LENGTH(vPART) > 8 THEN
        vPos := 0;
      END IF;

    END IF;

    -- wenn obere nicht gefunden dann  _[12] ..... finden.
    IF vPos = 0 THEN

      SELECT INSTR(pLine, '[') INTO vPosKlammerL FROM DUAL;
      SELECT INSTR(pLine, ']') INTO vPosKlammerR FROM DUAL;

      -- das ganze nur wenn Klammern drin sind
      IF vPosKlammerL > 0 AND vPosKlammerR > 0 THEN
        vReg8 := vReg1 || vReg || vReg || vReg || vReg || vReg || vReg || vReg ||'$';
        SELECT REGEXP_INSTR(pLine, vReg8)
        INTO vPos
        FROM DUAL;
      END IF;

    END IF;

    RETURN vPos;
  EXCEPTION
  WHEN no_data_found THEN
     RETURN vPos;
  WHEN OTHERS THEN
    pkg_lib_ex_basic_newapi.WRITELOG(P_EVENTTYPE      => pkg_lib_ex_basic_newapi.EventT_SysError,
                              P_SESSION        => vEX_SESSION,
                              P_IDENTIFICATION => $$PLSQL_UNIT,
                              P_TXT            => 'Systemfehler');
  END getPosReg;


/*------------------------------------------------------------------------------
Name:      getPartColType

Parameter:    pOwner      Schema des Objekts
              pObjName    Name   des Objekts
              pColName    Spalte deren Wert umformatiert werden soll

Return:       Datentyp

Beschreibung: Funktion, gibt den Datentyp der Partitionierungspalte zurück, wenn
              dieser erlaubt ist
------------------------------------------------------------------------------*/
  FUNCTION getPartColType(
      pOwner       IN VARCHAR2,
      pObjName     IN VARCHAR2,
      pColName     IN VARCHAR2)
    RETURN VARCHAR2
  AS
    vColType  VARCHAR2(128) := NULL;
BEGIN
  SELECT DATA_TYPE  INTO vColType FROM DBA_TAB_COLUMNS
   WHERE upper(COLUMN_NAME) = upper(pColName)
     AND upper(OWNER) = upper(pOwner)
     AND upper(TABLE_NAME) = upper(pObjName)
     AND upper(DATA_TYPE) IN ('VARCHAR2','DATE','NUMBER');

  RETURN vColType;
EXCEPTION
  WHEN no_data_found THEN
     RETURN vColType;
  WHEN OTHERS THEN
    pkg_lib_ex_basic_newapi.WRITELOG(P_EVENTTYPE      => pkg_lib_ex_basic_newapi.EventT_SysError,
                              P_SESSION        => vEX_SESSION,
                              P_IDENTIFICATION => $$PLSQL_UNIT,
                              P_TXT            => 'Systemfehler');
  END getPartColType;


/*------------------------------------------------------------------------------
Name:      getObjType

Parameter:    pLine      Name des Directory-Objekts in welchem die STL sind.

Return:       Postition des gefundenen Musters

Beschreibung: Funktion, prüft ob der Eingabestring einem gewissen Muster
              entspricht
------------------------------------------------------------------------------*/
  FUNCTION getObjType(
      pOwner       IN VARCHAR2,
      pObjName     IN VARCHAR2)
    RETURN CHAR
  AS
    vObjTypeName  VARCHAR2(30);
    vObjType   CHAR(1);
BEGIN
  SELECT MIN(
    CASE
      WHEN OBJECT_TYPE = 'TABLE'
      THEN 't'
      WHEN OBJECT_TYPE = 'TABLE PARTITION'
      THEN 'p'
      WHEN OBJECT_TYPE = 'TABLE SUBPARTITION'
      THEN 'p'
      WHEN OBJECT_TYPE = 'VIEW'
      THEN 'v'
      WHEN OBJECT_TYPE = 'MATERIALIZED VIEW'
      THEN 't'
      WHEN OBJECT_TYPE = 'SYNONYM'
      THEN 's'
      ELSE '0' -- anders Objekt in DB
    END ) AS objType
    INTO vObjType
  FROM dba_objects
  WHERE OWNER     = upper(pOwner)
  AND OBJECT_NAME = upper(pObjName);
  RETURN nvl(vObjType, '0');
EXCEPTION
  WHEN no_data_found THEN
     vObjType := '0'; -- nicht gefunden in DB
     RETURN vObjType;
  WHEN OTHERS THEN
    pkg_lib_ex_basic_newapi.WRITELOG(P_EVENTTYPE      => pkg_lib_ex_basic_newapi.EventT_SysError,
                              P_SESSION        => vEX_SESSION,
                              P_IDENTIFICATION => $$PLSQL_UNIT,
                              P_TXT            => 'Systemfehler');
  END getObjType;


/*------------------------------------------------------------------------------
Name:      getFormatCol

Parameter:    pOwner      Schema des Objekts
              pObjName    Name   des Objekts
              pColName    Spalte deren Wert umformatiert werden soll

Return:       Formatierter String

Beschreibung: Funktion, formatiert die Partitionsspalte in einen String
------------------------------------------------------------------------------*/
  FUNCTION getFormatCol(
      pOwner       IN VARCHAR2,
      pObjName     IN VARCHAR2,
      pColName     IN VARCHAR2)
    RETURN VARCHAR2
  AS
    vColName VARCHAR2(128):= NULL;
    vColType  VARCHAR2(30);
BEGIN
  vColType:= getPartColType(pOwner, pObjName, pColName);
  CASE
  WHEN vColType = 'VARCHAR2' THEN
    vColName  := pColName;
  WHEN vColType = 'NUMBER' THEN
    vColName  := 'TO_CHAR(' || pColName || ')';
  WHEN vColType = 'DATE' THEN
    vColName  := 'TO_CHAR(' || pColName || ', ' || chr(39) || 'YYYYMMDD' || chr(39) || ')';
  ELSE
    NULL;
  END CASE;
  RETURN vColName;
EXCEPTION
  WHEN OTHERS THEN
    pkg_lib_ex_basic_newapi.WRITELOG(P_EVENTTYPE      => pkg_lib_ex_basic_newapi.EventT_SysError,
                              P_SESSION        => vEX_SESSION,
                              P_IDENTIFICATION => $$PLSQL_UNIT,
                              P_TXT            => 'Systemfehler');
  END getFormatCol;

/*------------------------------------------------------------------------------
Name:      getFormatCol

Parameter:    pOwner      Schema des Objekts
              pObjName    Name   des Objekts
              pColName    Spalte deren Wert umformatiert werden soll

Return:       Formatierter String

Beschreibung: Funktion, formatiert die Partitionsspalte in den Tabellentyp,
              damit Partitionpruning funktioniert
------------------------------------------------------------------------------*/
  FUNCTION getFormatColMat(
      pOwner       IN VARCHAR2,
      pObjName     IN VARCHAR2,
      pColName     IN VARCHAR2)
    RETURN VARCHAR2
  AS
    vColName VARCHAR2(128):= NULL;
    vColType  VARCHAR2(30);
BEGIN
  vColType:= getPartColType(pOwner, pObjName, pColName);
  CASE
  WHEN vColType = 'VARCHAR2' THEN
    vColName  := pColName;
  WHEN vColType = 'NUMBER' THEN
    vColName  :=  pColName;
  WHEN vColType = 'DATE' THEN
    vColName  := 'TO_DATE(' || pColName || ', ' || chr(39) || 'YYYYMMDD' || chr(39) || ')';
  ELSE
    NULL;
  END CASE;
  RETURN vColName;
EXCEPTION
  WHEN OTHERS THEN
    pkg_lib_ex_basic_newapi.WRITELOG(P_EVENTTYPE      => pkg_lib_ex_basic_newapi.EventT_SysError,
                              P_SESSION        => vEX_SESSION,
                              P_IDENTIFICATION => $$PLSQL_UNIT,
                              P_TXT            => 'Systemfehler');
END getFormatColMat;




/*------------------------------------------------------------------------------
Name:         insertTableJob

Parameter:    pViewObj Record mit den Werten für das Insert

Beschreibung: Fügt einen Job vom Typ Tabelle eine.
------------------------------------------------------------------------------*/
procedure insertTableJob(pTableObj STAT_EXP_DATAMART_MCFG.WRK_EX_OBJ%ROWTYPE) IS
job_rec STAT_EXP_DATAMART_MCFG.WRK_EX_JOB%ROWTYPE;
BEGIN
      job_rec.JOB_ID        := STAT_EXP_DATAMART_MCFG.SEQ_JOB_IDS.NEXTVAL;
      job_rec.JOB_OBJ_NAME  := lower(pTableObj.STL_OBJ_NAME);
      job_rec.JOB_OBJ_TYPE  := 't';
      job_rec.JOB_PART_COL :=  'NO';
      job_rec.JOB_PART      := 'NO';
      job_rec.JOB_KZ        := pTableObj.STL_KZ;
      job_rec.JOB_SCHEMA    := pTableObj.STL_SCHEMA;
      job_rec.JOB_SESSION   := pTableObj.STL_SESSION;
      job_rec.JOB_STATUS    := 'C';
      job_rec.JOB_ZBDAT     := sysdate;
      job_rec.JOB_STL_ID    := pTableObj.STL_ID;

      INSERT INTO STAT_EXP_DATAMART_MCFG.WRK_EX_JOB VALUES job_rec;

EXCEPTION
  WHEN OTHERS THEN
    pkg_lib_ex_basic_newapi.WRITELOG(P_EVENTTYPE      => pkg_lib_ex_basic_newapi.EventT_SysError,
                              P_SESSION        => vEX_SESSION,
                              P_IDENTIFICATION => $$PLSQL_UNIT,
                              P_TXT            => 'Systemfehler');
END insertTableJob;

/*------------------------------------------------------------------------------
Name:         insertPTableJob

Parameter:    pViewObj Record mit den Werten für das Insert

Beschreibung: Fügt einen Job vom Typ partitionierte Tabelle eine.
------------------------------------------------------------------------------*/
procedure insertPTableJob(pPTableObj STAT_EXP_DATAMART_MCFG.WRK_EX_OBJ%ROWTYPE,
                          pJOB_PART  STAT_EXP_DATAMART_MCFG.WRK_EX_JOB.JOB_PART%TYPE,
                          pZusatz    VARCHAR2) IS
job_rec       STAT_EXP_DATAMART_MCFG.WRK_EX_JOB%ROWTYPE;
TYPE          cPart_Typ IS REF CURSOR;
cPart         cPart_Typ;
vPartCol      DBA_PART_KEY_COLUMNS.column_name%TYPE;
vPart         STAT_EXP_DATAMART_MCFG.WRK_EX_JOB.JOB_PART%TYPE;
vJOB_PART     VARCHAR2(128);
vSTL_OBJ_NAME VARCHAR2(128);
vKlausel      VARCHAR2(2000) := ' IN (201101, 201102, 201103)';
vColType      VARCHAR2(60);
vObjType      CHAR;
v_SQL         CLOB := NULL;
vWhere        VARCHAR2(4000);
BEGIN
   IF  pJOB_PART IS NULL THEN
     vJOB_PART  := '^*$';
   ELSE
     vJOB_PART  := '^' || pJOB_PART || '$';
   END IF;


   vSTL_OBJ_NAME :=  '^' ||star2regex(upper(pPTableObj.STL_OBJ_NAME))||  '$';

    logVerbose('vSTL_OBJ_NAME >'||vSTL_OBJ_NAME ||'<');
    logVerbose('vJOB_PART     >'||vJOB_PART ||'<');
    logVerbose('pZusatz       >'||pZusatz ||'<');




    FOR cPartCol IN
    (SELECT t.table_name, c.column_name
      FROM DBA_PART_TABLES t
     INNER JOIN DBA_PART_KEY_COLUMNS c
        ON t.table_name           = c.name
       AND t.owner                = c.owner
     WHERE REGEXP_LIKE (UPPER(t.table_name), vSTL_OBJ_NAME)
       AND UPPER(t.owner)         = UPPER(pPTableObj.STL_SCHEMA)
    )
    LOOP

        vColType:= NULL;
        vColType:= getPartColType(pPTableObj.STL_SCHEMA, pPTableObj.STL_OBJ_NAME, cPartCol.column_name);



        IF vColType IS NULL THEN
          vKlausel := NULL;
          vObjType := 't';
          vPartCol :=  cPartCol.column_name;
        ELSE
          vKlausel := getKlausel(pZusatz, cPartCol.column_name);
          vObjType := 'p';
          vPartCol := getFormatCol(pPTableObj.STL_SCHEMA, pPTableObj.STL_OBJ_NAME, cPartCol.column_name);
        END IF;




        logVerbose('vKlausel      >'||vKlausel||'<');
        vJOB_PART := star2regex(vJOB_PART);
           logInfo('vPartCol:'||vPartCol);
           logInfo('vJobPart:'||vJob_Part);
           logInfo('column_name:'||cPartCol.column_name);

        if vKlausel is null and vJOB_PART is not null
        and rtrim(translate(vJOB_PART,'^$0123456789',' '))is null
        and (cPartCol.column_name like '%MOW_ID' or cPartCol.column_name like '%MON_ID')
        then
           logInfo('Kandidat für Optimierung:');
           logInfo('vPartCol:'||vPartCol);
           logInfo('vJobPart:'||vJob_Part);
           vWhere := ' AND '||cPartCol.column_name||'='||translate(vJOB_PART,'^$',' ')||' AND ROWNUM=1 ';
        elsif vKlausel is not null and vJob_Part = '^'||vStarReplacement ||'$' -- *
        then
           vWhere := replace(vKlausel,'WHERE','AND') || ' GROUP BY ' || cPartCol.column_name;
        else
           vWhere := replace(vKlausel,'WHERE','AND') ||' AND REGEXP_LIKE ('||cPartCol.column_name||', '''||vJOB_PART||''') '|| ' GROUP BY ' || cPartCol.column_name;
        end if;
        logInfo('vWhere:'||vWhere);

        v_SQL :=  'SELECT ' || vPartCol || ' FROM ' ||
          '( ' ||
          'SELECT /*+ parallel */ /* '||vEX_Session||' */' || cPartCol.column_name || ' FROM ' || UPPER(pPTableObj.STL_SCHEMA) ||
          '.' || UPPER(cPartCol.table_name) ||' WHERE 1=1 '|| vWhere ||
          ' )';
          --dbms_output.put_line('SQL:>'||v_sql||'<');
        logVerbose(v_SQL);
        OPEN cPart FOR
          v_SQL;
        LOOP
          FETCH cPart INTO vPart;
          EXIT WHEN cPart%NOTFOUND;

          -- * ersetzen fuer REGEXP
          vJOB_PART := star2regex(vJOB_PART);
          --dbms_output.put_line('vJOB_PART:>'||vJOB_PART||'<');
          IF REGEXP_LIKE (nvl(vPart,'null'), vJOB_PART) THEN
            job_rec.JOB_ID        := STAT_EXP_DATAMART_MCFG.SEQ_JOB_IDS.NEXTVAL;
            job_rec.JOB_OBJ_NAME  := lower(cPartCol.table_name);
            job_rec.JOB_OBJ_TYPE  := vObjType;
            job_rec.JOB_PART_COL  := lower(cPartCol.column_name);
            --dbms_output.put_line('vPART: >'||vPART||'<');
            job_rec.JOB_PART      := nvl(vPart,'null');
            job_rec.JOB_KZ        := pPTableObj.STL_KZ;
            job_rec.JOB_SCHEMA    := pPTableObj.STL_SCHEMA;
            job_rec.JOB_SESSION   := pPTableObj.STL_SESSION;
            job_rec.JOB_STATUS    := 'C';
            job_rec.JOB_ZBDAT     := sysdate;
            job_rec.JOB_STL_ID    := pPTableObj.STL_ID;

            INSERT INTO STAT_EXP_DATAMART_MCFG.WRK_EX_JOB VALUES job_rec;
          END IF;

        END LOOP; -- cPart
        CLOSE cPart;

      END LOOP; -- cPartCol
EXCEPTION
  WHEN OTHERS THEN
    pkg_lib_ex_basic_newapi.WRITELOG(P_EVENTTYPE      => pkg_lib_ex_basic_newapi.EventT_SysError,
                              P_SESSION        => vEX_SESSION,
                              P_IDENTIFICATION => $$PLSQL_UNIT,
                              P_TXT            => 'Systemfehler');
END insertPTableJob;



/*------------------------------------------------------------------------------
Name:         insertPViewJob

Parameter:    pViewObj Record mit den Werten für das Insert

Beschreibung: Fügt einen Job vom Typ partitionierte Tabelle eine.
------------------------------------------------------------------------------*/
procedure insertPViewJob(pPTableObj STAT_EXP_DATAMART_MCFG.WRK_EX_OBJ%ROWTYPE,
                          pJOB_PART  STAT_EXP_DATAMART_MCFG.WRK_EX_JOB.JOB_PART%TYPE,
                          pZusatz    VARCHAR2) IS
job_rec       STAT_EXP_DATAMART_MCFG.WRK_EX_JOB%ROWTYPE;
TYPE          cPart_Typ IS REF CURSOR;
cPart         cPart_Typ;
vPartCol      DBA_PART_KEY_COLUMNS.column_name%TYPE;
vPart         STAT_EXP_DATAMART_MCFG.WRK_EX_JOB.JOB_PART%TYPE;
vJOB_PART     VARCHAR2(128);
vSTL_OBJ_NAME VARCHAR2(128);
vKlausel      VARCHAR2(128) := ' IN (201101, 201102, 201103)';
vSQL          VARCHAR2(4000);
vWhere        VARCHAR2(4000);
BEGIN
   logVerbose('Beginn PViewJob');
   logVerbose('STL_OBJ_NAME >'||pPTableObj.STL_OBJ_NAME ||'<');
   logVerbose('pJOB_PART    >'||pJOB_PART ||'<');
   logVerbose('pZusatz      >'||pZusatz ||'<');


   IF  pJOB_PART IS NULL THEN
     vJOB_PART  := '^*$';
   ELSE
     vJOB_PART  := '^' || pJOB_PART || '$';
   END IF;

   vSTL_OBJ_NAME :=  '^' || star2regex(upper(pPTableObj.STL_OBJ_NAME)) ||  '$';
    -- MOW_ID ist bei einem partitioned View jetzt auch zulässig.
    -- Aus Gründen der Abwärtskompatibilität wird die MON_ID bevorzugt,
    -- falls beide Spalten in einem View vorhanden sind
    FOR cPartCol IN
--    (SELECT view_name, 'MON_ID' AS column_name
--      FROM DBA_VIEWS
--     WHERE REGEXP_LIKE (UPPER(view_name), vSTL_OBJ_NAME)
--       AND UPPER(owner)         = UPPER(pPTableObj.STL_SCHEMA)
--    )
    (
--    WITH v_view as
--    (SELECT a.view_name,b.column_name AS column_name, count(*) over (partition by a.view_name order by a.view_name,b.column_name ) colorder
--      FROM DBA_VIEWS a
--      INNER JOIN dba_tab_columns b on (a.owner=b.owner and a.view_name=b.table_name and b.column_name in ('MON_ID','MOW_ID'))
--     WHERE REGEXP_LIKE (UPPER(view_name), vSTL_OBJ_NAME)
--       AND UPPER(a.owner)         = UPPER(pPTableObj.STL_SCHEMA)
--    union SELECT a.table_name,b.column_name AS column_name, count(*) over (partition by a.table_name order by a.table_name,b.column_name ) colorder
--      FROM DBA_TABLES a
--      INNER JOIN dba_tab_columns b on (a.owner=b.owner and a.table_name=b.table_name and b.column_name in ('MON_ID','MOW_ID'))
--     WHERE REGEXP_LIKE (UPPER(a.table_name), vSTL_OBJ_NAME)
--       AND UPPER(a.owner)         = UPPER(pPTableObj.STL_SCHEMA)
--     )
--     select view_name,column_name
--     from v_view
--     where colorder=1
with
-- View-Dependencies
v_deps as
(select owner,name,type,referenced_owner ref_owner,referenced_name ref_name,referenced_type ref_type
from dba_dependencies
where type='VIEW'
and referenced_type in ('VIEW','TABLE')
--and referenced_owner not in  ('SYS','SYSTEM','CTXSYS','DBSNMP')
)
-- Bildung der transitiven Hülle
, v_rek (owner,name,type,ebene,path) as (
select distinct owner,name,type,1,owner||'.'||name from v_deps where 1=1
and REGEXP_LIKE (UPPER(name), vSTL_OBJ_NAME)
       AND UPPER(owner)         = UPPER(pPTableObj.STL_SCHEMA)
union all select t2.ref_owner,t2.ref_name,t2.ref_type,t1.ebene+1,path||'=>'||t2.ref_owner||'.'||t2.ref_name from v_deps t2 inner join v_rek t1 on t1.owner=t2.owner and t2.name=t1.name and t1.type=t2.type
),
-- Alle Tabellen aus der transitiven Hülle selektieren
v_tabs as (select owner,name table_name from v_rek where type='TABLE')
-- Ermittlung der Partitionsspalte
,v_part_col as (select column_name from (select column_name from dba_part_key_columns
where (owner,name) in (select owner,table_name from v_tabs) order by column_name desc ) -- MOW_ID vor MON_ID
where rownum=1)
-- Viewname und Partitionskriterium
,v_views_part_table as (select 1 sortorder,view_name,column_name
from dba_views v
inner join dba_tab_columns c on (v.view_name=c.table_name and v.owner=c.owner)
where 1=1
and REGEXP_LIKE (UPPER(v.view_name), vSTL_OBJ_NAME)
       AND UPPER(v.owner)         = UPPER(pPTableObj.STL_SCHEMA)
and c.column_name in (select column_name from v_part_col))
-- Fallback Ermittlung des Partitionskriteriums an Hand der View-Spalten
,v_view as
    (SELECT a.view_name,b.column_name AS column_name, count(*) over (partition by a.view_name order by a.view_name,b.column_name desc) colorder
      FROM DBA_VIEWS a
      INNER JOIN dba_tab_columns b on (a.owner=b.owner and a.view_name=b.table_name and b.column_name in ('MON_ID','MOW_ID'))
     WHERE REGEXP_LIKE (UPPER(view_name), vSTL_OBJ_NAME)
       AND UPPER(a.owner)         = UPPER(pPTableObj.STL_SCHEMA))
-- Partitionskriterium aus View
,v_views_part_view as (     select 2 sortorder,view_name,column_name
     from v_view
     where colorder=1)
-- Beide Ergebnisse zusammenschmeißen
,v_sel as (select * from v_views_part_table
union select * from v_views_part_view
order by 1
)
-- Partitionskriterium der partitionierten Tabelle hat Vorrang
select view_name,column_name from v_sel
where rownum=1
    )
    LOOP

        vPartCol :=  cPartCol.column_name;

        vKlausel := getKlausel(pZusatz, vPartCol);

        logVerbose('========================================');
        logVerbose('vPartCol    >'||vPartCol ||'<');
        logVerbose('vKlausel    >'||vKlausel ||'<');

        vJOB_PART := star2regex(vJOB_PART);

        if vKlausel is null and vJOB_PART is not null
        and rtrim(translate(vJOB_PART,'^$0123456789',' '))is null
        and (cPartCol.column_name like '%MOW_ID' or cPartCol.column_name like '%MON_ID')
        then
           logInfo('Kandidat für Optimierung:');
           logInfo('vPartCol:'||vPartCol);
           logInfo('vJobPart:'||vJob_Part);
           vWhere := ' AND '||cPartCol.column_name||'='||translate(vJOB_PART,'^$',' ')||' AND ROWNUM=1 ';
        elsif vKlausel is not null and vJob_Part = '^'||vStarReplacement||'$' -- *
        then
           vWhere := replace(vKlausel,'WHERE','AND') || ' GROUP BY ' || cPartCol.column_name;
        else
           vWhere := replace(vKlausel,'WHERE','AND') ||' AND REGEXP_LIKE ('||cPartCol.column_name||', '''||vJOB_PART||''') '|| ' GROUP BY ' || cPartCol.column_name;
        end if;
        logInfo('vWhere:'||vWhere);



        vSQL :=  --'SELECT ' || vPartCol || ' FROM ' ||
--          '( ' ||
          'SELECT /*+ parallel */ /* '||vEX_Session||' */' || vPartCol || ' FROM /*'||vEX_SESSION||' */ ' || UPPER(pPTableObj.STL_SCHEMA) ||
--          '.' || UPPER(cPartCol.view_name) ||' WHERE 1=1 '|| replace(vKlausel,'WHERE','AND') ||' AND REGEXP_LIKE ('||vPartCol||', '''||vJOB_PART||''') '|| ' GROUP BY ' || vPartCol ||
          '.' || UPPER(cPartCol.view_name) ||' WHERE 1=1 '|| vWhere
--          ||' )'
          ;
          --dbms_output.put_line('SQL:>'||v_sql||'<');
          logVerbose(vSQL);
--        OPEN cPart FOR
--          'SELECT DISTINCT ' || vPartCol || ' FROM ' || UPPER(pPTableObj.STL_SCHEMA) ||
--          '.' || UPPER(cPartCol.view_name) || vKlausel ;
        OPEN cPart FOR
          vSQL;

        LOOP
          FETCH cPart INTO vPart;
          EXIT WHEN cPart%NOTFOUND;
        logVerbose('----------------------------------------');
        logVerbose('vPart    >'||vPart ||'<');


          -- * ersetzen fuer REGEXP
          vJOB_PART := star2regex(vJOB_PART);


          IF REGEXP_LIKE (vPart, vJOB_PART) THEN
            logVerbose('vPart    >'||vPart ||'< inserted.');
            job_rec.JOB_ID        := STAT_EXP_DATAMART_MCFG.SEQ_JOB_IDS.NEXTVAL;
            job_rec.JOB_OBJ_NAME  := lower(cPartCol.view_name);
            job_rec.JOB_OBJ_TYPE  := 'p';
            job_rec.JOB_PART_COL  := vPartCol;
            job_rec.JOB_PART      := vPart;
            job_rec.JOB_KZ        := pPTableObj.STL_KZ;
            job_rec.JOB_SCHEMA    := pPTableObj.STL_SCHEMA;
            job_rec.JOB_SESSION   := pPTableObj.STL_SESSION;
            job_rec.JOB_STATUS    := 'C';
            job_rec.JOB_ZBDAT     := sysdate;
            job_rec.JOB_STL_ID    := pPTableObj.STL_ID;

            INSERT INTO STAT_EXP_DATAMART_MCFG.WRK_EX_JOB VALUES job_rec;
          END IF;

        END LOOP; -- cPart
        CLOSE cPart;

      END LOOP; -- cPartCol
--   logInfo('Ende PViewJob');

EXCEPTION
  WHEN OTHERS THEN
    pkg_lib_ex_basic_newapi.WRITELOG(P_EVENTTYPE      => pkg_lib_ex_basic_newapi.EventT_SysError,
                              P_SESSION        => vEX_SESSION,
                              P_IDENTIFICATION => $$PLSQL_UNIT,
                              P_TXT            => 'Systemfehler');
END insertPViewJob;


/*------------------------------------------------------------------------------
Name:         insertSynonymeJob

Parameter:    pViewObj Record mit den Werten für das Insert

Beschreibung: Fügt einen Job vom Typ Synonym ein.
------------------------------------------------------------------------------*/
procedure insertSynonymeJob(pSynonymeObj STAT_EXP_DATAMART_MCFG.WRK_EX_OBJ%ROWTYPE) IS
job_rec STAT_EXP_DATAMART_MCFG.WRK_EX_JOB%ROWTYPE;
BEGIN
  job_rec.JOB_ID        := STAT_EXP_DATAMART_MCFG.SEQ_JOB_IDS.NEXTVAL;
  job_rec.JOB_OBJ_NAME  := lower(pSynonymeObj.STL_OBJ_NAME);
  job_rec.JOB_OBJ_TYPE  := 's';
  job_rec.JOB_PART_COL  := 'NO';
  job_rec.JOB_PART      := 'NO';
  job_rec.JOB_KZ        := pSynonymeObj.STL_KZ;
  job_rec.JOB_SCHEMA    := pSynonymeObj.STL_SCHEMA;
  job_rec.JOB_SESSION   := pSynonymeObj.STL_SESSION;
  job_rec.JOB_STATUS    := 'C';
  job_rec.JOB_ZBDAT     := sysdate;
  job_rec.JOB_STL_ID    := pSynonymeObj.STL_ID;

  INSERT INTO STAT_EXP_DATAMART_MCFG.WRK_EX_JOB VALUES job_rec;
EXCEPTION
  WHEN OTHERS THEN
    pkg_lib_ex_basic_newapi.WRITELOG(P_EVENTTYPE      => pkg_lib_ex_basic_newapi.EventT_SysError,
                              P_SESSION        => vEX_SESSION,
                              P_IDENTIFICATION => $$PLSQL_UNIT,
                              P_TXT            => 'Systemfehler');
END insertSynonymeJob;

/*------------------------------------------------------------------------------
Name:         insertViewJob

Parameter:    pViewObj Record mit den Werten für das Insert

Beschreibung: Fügt einen Job vom Typ View ein.
------------------------------------------------------------------------------*/
procedure insertViewJob(pViewObj STAT_EXP_DATAMART_MCFG.WRK_EX_OBJ%ROWTYPE) IS
job_rec STAT_EXP_DATAMART_MCFG.WRK_EX_JOB%ROWTYPE;
BEGIN
  job_rec.JOB_ID        := STAT_EXP_DATAMART_MCFG.SEQ_JOB_IDS.NEXTVAL;
  job_rec.JOB_OBJ_NAME  := lower(pViewObj.STL_OBJ_NAME);
  job_rec.JOB_OBJ_TYPE  := 'v';
  job_rec.JOB_PART_COL  := 'NO';
  job_rec.JOB_PART      := 'NO';
  job_rec.JOB_KZ        := pViewObj.STL_KZ;
  job_rec.JOB_SCHEMA    := pViewObj.STL_SCHEMA;
  job_rec.JOB_SESSION   := pViewObj.STL_SESSION;
  job_rec.JOB_STATUS    := 'C';
  job_rec.JOB_ZBDAT     := sysdate;
  job_rec.JOB_STL_ID    := pViewObj.STL_ID;

  INSERT INTO STAT_EXP_DATAMART_MCFG.WRK_EX_JOB VALUES job_rec;

EXCEPTION
  WHEN OTHERS THEN
    pkg_lib_ex_basic_newapi.WRITELOG(P_EVENTTYPE      => pkg_lib_ex_basic_newapi.EventT_SysError,
                              P_SESSION        => vEX_SESSION,
                              P_IDENTIFICATION => $$PLSQL_UNIT,
                              P_TXT            => 'Systemfehler');
END insertViewJob;

/*------------------------------------------------------------------------------
Name:         insertDimViewJob

Parameter:    pViewObj Record mit den Werten fuer das Insert

Beschreibung: Fügt einen Job vom Typ Dimensions-View ein. Eventuell mit
              mehreren Unter-Jobs.
------------------------------------------------------------------------------*/
procedure insertDimViewJob(pViewObj STAT_EXP_DATAMART_MCFG.WRK_EX_OBJ%ROWTYPE) IS
job_rec STAT_EXP_DATAMART_MCFG.WRK_EX_JOB%ROWTYPE;
newViewObj STAT_EXP_DATAMART_MCFG.WRK_EX_OBJ%ROWTYPE;
BEGIN
  job_rec.JOB_ID        := STAT_EXP_DATAMART_MCFG.SEQ_JOB_IDS.NEXTVAL;
  job_rec.JOB_OBJ_NAME  := lower(pViewObj.STL_OBJ_NAME);
  job_rec.JOB_OBJ_TYPE  := 'v';
  job_rec.JOB_PART_COL  := 'NO';
  job_rec.JOB_PART      := 'NO';
  job_rec.JOB_KZ        := pViewObj.STL_KZ;
  job_rec.JOB_SCHEMA    := pViewObj.STL_SCHEMA;
  job_rec.JOB_SESSION   := pViewObj.STL_SESSION;
  job_rec.JOB_STATUS    := 'C';
  job_rec.JOB_ZBDAT     := sysdate;
  job_rec.JOB_STL_ID    := pViewObj.STL_ID;

  INSERT INTO STAT_EXP_DATAMART_MCFG.WRK_EX_JOB VALUES job_rec;

  newViewObj := pViewObj;
  FOR cDepend IN --alle selektierte Objekte in der View
  (SELECT REFERENCED_OWNER, REFERENCED_NAME, REFERENCED_TYPE
     FROM dba_dependencies
    WHERE NAME = UPPER(pViewObj.STL_OBJ_NAME)
      AND OWNER  = UPPER(pViewObj.STL_SCHEMA)
  )
  LOOP
  IF cDepend.REFERENCED_TYPE = 'VIEW' THEN
    newViewObj := pViewObj;
    newViewObj.STL_OBJ_NAME := lower(cDepend.REFERENCED_NAME);
    newViewObj.STL_SCHEMA := lower(cDepend.REFERENCED_OWNER);
    insertDimViewJob(newViewObj);
  ELSE
    job_rec.JOB_ID        := STAT_EXP_DATAMART_MCFG.SEQ_JOB_IDS.NEXTVAL;
    job_rec.JOB_OBJ_NAME  := lower(cDepend.REFERENCED_NAME);
    job_rec.JOB_PART_COL  := 'NO';
    job_rec.JOB_PART      := 'NO';
    job_rec.JOB_KZ        := pViewObj.STL_KZ;
    job_rec.JOB_SCHEMA    := lower(cDepend.REFERENCED_OWNER);
    job_rec.JOB_SESSION   := pViewObj.STL_SESSION;
    job_rec.JOB_STATUS    := 'C';
    job_rec.JOB_ZBDAT     := sysdate;
    job_rec.JOB_STL_ID    := pViewObj.STL_ID;

     IF cDepend.REFERENCED_TYPE = 'TABLE' THEN
        job_rec.JOB_OBJ_TYPE  := 't';
     ELSIF cDepend.REFERENCED_TYPE = 'SYNONYM' THEN
     job_rec.JOB_OBJ_TYPE  := 's';
     END IF;

    INSERT INTO STAT_EXP_DATAMART_MCFG.WRK_EX_JOB VALUES job_rec;
  END IF;
  END LOOP;

EXCEPTION
  WHEN OTHERS THEN
    pkg_lib_ex_basic_newapi.WRITELOG(P_EVENTTYPE      => pkg_lib_ex_basic_newapi.EventT_SysError,
                              P_SESSION        => vEX_SESSION,
                              P_IDENTIFICATION => $$PLSQL_UNIT,
                              P_TXT            => 'Systemfehler');
END insertDimViewJob;

/*------------------------------------------------------------------------------
Name:         createJobBranch

Parameter:    pObjRec               Aktueller Objekt-Rekord
              pScanForPartition     Default:True, Auf Partitions-Suffix prüfen

Beschreibung: Erzeugt aus den Steuerlistenobjekten die Export-Jobs
------------------------------------------------------------------------------*/
procedure createJobBranch ( pObjRec in out nocopy STAT_EXP_DATAMART_MCFG.WRK_EX_OBJ%ROWTYPE ,
                            pScanForPartition in Boolean default true)
IS
obj_rec         STAT_EXP_DATAMART_MCFG.WRK_EX_OBJ%ROWTYPE;
job_rec         STAT_EXP_DATAMART_MCFG.WRK_EX_JOB%ROWTYPE;
vSTL_SCHEMA     VARCHAR2(128);
vSTL_OBJ_NAME   VARCHAR2(255);
vSTL_OBJ_NAMEexp   VARCHAR2(255);
vSTL_OBJ_NAME_P VARCHAR2(2000);
vOBJ_TYPE       VARCHAR2(2000);
vIAB_OBJ_NAME   VARCHAR2(2000);
--vJOB_PART       STAT_EXP_DATAMART_MCFG.WRK_EX_JOB.JOB_PART%TYPE;
vJOB_PART       varchar2(1000);
vPosLast        INTEGER;
vPos            INTEGER;
objExists       INTEGER;
vZusatz         VARCHAR2(2000);
BEGIN
     logInfo('Start Erzeugen JobBranch ('||case when pScanforPartition then 'Partition' else 'Table' end ||')');

    -- für alle Steuerlisten-Objekte
    vSTL_OBJ_NAMEexp := MonId(upper(pObjRec.STL_OBJ_NAME));
    vSTL_OBJ_NAMEexp := YearId(upper(vSTL_OBJ_NAMEexp));

    vPosLast   := getPosZusatz(vSTL_OBJ_NAMEexp);
     logVerbose('vPosLast:'||vPosLast);
      IF vPosLast <> 0 THEN
        vSTL_OBJ_NAME_P :=  upper(SUBSTR(vSTL_OBJ_NAMEexp, 1, vPosLast - 1));
        vZusatz := SUBSTR(vSTL_OBJ_NAMEexp, vPosLast + 1, LENGTH(vSTL_OBJ_NAMEexp)); -- ohne ':'
      ELSE
        vSTL_OBJ_NAME_P :=  upper(vSTL_OBJ_NAMEexp);
        vZusatz := NULL;
      END IF;
     logVerbose('vZusatz:'||vZusatz);

     If pScanForPartition
     then
        vPos       := getPosReg(vSTL_OBJ_NAME_P);
     else
        vPos := 0;
     end if;

      IF vPos <> 0 THEN
        vSTL_OBJ_NAME :=  '^' || star2regex(upper(SUBSTR(vSTL_OBJ_NAME_P, 1, vPos - 1))) ||  '$';
        vJOB_PART     := SUBSTR(vSTL_OBJ_NAME_P, vPos + 1, LENGTH(vSTL_OBJ_NAME_P)); -- ohne '_'
      ELSE
        vSTL_OBJ_NAME :=  '^' || star2regex(upper(vSTL_OBJ_NAME_P)) || '$';
        vJOB_PART     := NULL;
      END IF;

     logVerbose('vSTL_OBJ_NAME:'||vSTL_OBJ_NAME);
     logVerbose('vJOB_PART:'||vJOB_PART);


      -- Pruefen ob es ein Objekt mit dem Origalnamen aus Steuerliste gibt
      -- z.B.  VF_AST_BEW_PLUS_20160600
      SELECT count(*) INTO objExists
        FROM dba_objects
       WHERE OBJECT_NAME = upper(vSTL_OBJ_NAMEexp)
         AND OWNER = upper(pObjRec.STL_SCHEMA);
     logVerbose('objExists:'||objExists);

      -- Wenn es ein Objekt mit dem Origalnamen gibt dann wir dieser genommen.
      -- Es wird nicht versucht den Zusatz "_20160600" als Partition zu interpretieren.
      IF objExists > 0  THEN
        vSTL_OBJ_NAME :=  '^' || upper(vSTL_OBJ_NAMEexp) || '$';
        vZusatz := NULL;
      END IF;
     logVerbose('vSTL_OBJ_NAME:'||vSTL_OBJ_NAME);
     logVerbose('vZusatz:'||vZusatz);



      -- * in Schemaname zulassen bei IAB
      vSTL_SCHEMA  :=  '^' || star2regex(upper(pObjRec.STL_SCHEMA)) || '$';

      FOR cObj IN
      (SELECT DISTINCT OBJECT_NAME,OWNER
        FROM dba_objects
       WHERE REGEXP_LIKE (upper(OBJECT_NAME), vSTL_OBJ_NAME)
         AND REGEXP_LIKE (upper(OWNER), vSTL_SCHEMA)
      )
      LOOP
        obj_rec   := pObjRec;
        vOBJ_TYPE := getObjType(cObj.OWNER, cObj.OBJECT_NAME);
        obj_rec.STL_OBJ_NAME := cObj.OBJECT_NAME;
        obj_rec.STL_SCHEMA := cObj.OWNER;
        logVerbose('vOBJ_TYPE:'||vOBJ_TYPE);
        logVerbose('STL_OBJ_NAME:'||obj_rec.STL_OBJ_NAME);
        logVerbose('STL_SCHEMA:'||obj_rec.STL_SCHEMA);

        CASE
        WHEN vOBJ_TYPE = 'p' AND not pScanForPartition THEN
           -- Eintrag ignorieren
            logVerbose('Partitioniertes Objekt im Tabellenscan ignoriert.');
           continue;
           --insertPViewJob(obj_rec, vJOB_PART, vZusatz);
           --setStatusOBJ (obj_rec.STL_ID, 'J');  --Status Jobs generiert setzen
        WHEN vOBJ_TYPE = 't' and not pScanForPartition THEN
           logVerbose('TableJob');
           insertTableJob(obj_rec);
           setStatusOBJ (obj_rec.STL_ID, 'J');  --Status Jobs generiert setzen
        WHEN vOBJ_TYPE = 'p' THEN
           logVerbose('PTableJob');
           insertPTableJob(obj_rec, vJOB_PART, vZusatz);
           setStatusOBJ (obj_rec.STL_ID, 'J');  --Status Jobs generiert setzen
        WHEN vOBJ_TYPE = 'v' AND obj_rec.STL_KZ = 'd' AND NOT vVIEW_MAT THEN
           logVerbose('DimViewJob');
           insertDimViewJob(obj_rec);
           setStatusOBJ (obj_rec.STL_ID, 'J');  --Status Jobs generiert setzen
        WHEN vOBJ_TYPE = 'v' AND obj_rec.STL_KZ = 'd' AND vVIEW_MAT THEN
           logVerbose('ViewJob');
           insertViewJob(obj_rec);
           setStatusOBJ (obj_rec.STL_ID, 'J');  --Status Jobs generiert setzen
        WHEN vOBJ_TYPE = 'v' AND obj_rec.STL_KZ = 'v' THEN
           logVerbose('PViewJob');
           insertPViewJob(obj_rec, vJOB_PART, vZusatz);
           setStatusOBJ (obj_rec.STL_ID, 'J');  --Status Jobs generiert setzen
        WHEN vOBJ_TYPE = 'v' AND obj_rec.STL_KZ NOT IN ( 'd', 'v') THEN
           logVerbose('ViewJob');
           insertViewJob(obj_rec);
           setStatusOBJ (obj_rec.STL_ID, 'J');  --Status Jobs generiert setzen
        WHEN vOBJ_TYPE = 's' THEN
           logVerbose('SynonymeJob');
           insertSynonymeJob(obj_rec);
           setStatusOBJ (obj_rec.STL_ID, 'J');  --Status Jobs generiert setzen
        ELSE
          logVerbose('Nix');
          NULL;
        END CASE;
        commit;
      END LOOP; -- cObj
    logInfo('Ende erzeugen JobBranch');
EXCEPTION
  WHEN OTHERS THEN
    pkg_lib_ex_basic_newapi.WRITELOG(P_EVENTTYPE      => pkg_lib_ex_basic_newapi.EventT_SysError,
                              P_SESSION        => vEX_SESSION,
                              P_IDENTIFICATION => $$PLSQL_UNIT,
                              P_TXT            => 'Systemfehler');
    raise;
END createJobBranch;


/*------------------------------------------------------------------------------
Name:         createJobs

Parameter:    keine

Beschreibung: Erzeugt aus den Steuerlistenobjekten die Export-Jobs
------------------------------------------------------------------------------*/
procedure createJobs IS
--obj_rec         STAT_EXP_DATAMART_MCFG.WRK_EX_OBJ%ROWTYPE;
--job_rec         STAT_EXP_DATAMART_MCFG.WRK_EX_JOB%ROWTYPE;
--vSTL_SCHEMA     VARCHAR2(60);
--vSTL_OBJ_NAME   VARCHAR2(60);
--vSTL_OBJ_NAME_P VARCHAR2(60);
--vOBJ_TYPE       VARCHAR2(60);
--vIAB_OBJ_NAME   VARCHAR2(60);
--vJOB_PART       STAT_EXP_DATAMART_MCFG.WRK_EX_JOB.JOB_PART%TYPE;
--vPosLast        INTEGER;
--vPos            INTEGER;
--objExists       INTEGER;
--vZusatz         VARCHAR2(60);
  vCnt NUMBER;
BEGIN
     logInfo('Start erzeugen der Export-Aufträge');



    -- fuer alle Steuerlisten-Objekte
    FOR cType IN
    (SELECT *
       FROM STAT_EXP_DATAMART_MCFG.WRK_EX_OBJ
      WHERE STL_STATUS = 'G'
        AND STL_SESSION = vEX_SESSION
      ORDER BY STL_ID
    )
    LOOP
      createJobBranch(pObjRec=>CType,pScanForPartition=>true);
      createJobBranch(pObjRec=>CType,pScanForPartition=>false);
    END LOOP;   -- cType

    COMMIT;

    -- Prüfen, ob die Steuerliste überhaupt aufgelöst werden konnte
    SELECT COUNT(*)
    INTO vCNT
    FROM STAT_EXP_DATAMART_MCFG.WRK_EX_JOB
    WHERE JOB_SESSION=vEX_SESSION;

    IF vCNT=0
    THEN
      raise_application_error(-20001,'Steuerliste konnte nicht aufgelöst werden, Bitte Schemazuordnung prüfen!');
    END IF;

    logInfo('Ende erzeugen der Export-Aufträge');
EXCEPTION
  WHEN OTHERS THEN
    pkg_lib_ex_basic_newapi.WRITELOG(P_EVENTTYPE      => pkg_lib_ex_basic_newapi.EventT_SysError,
                              P_SESSION        => vEX_SESSION,
                              P_IDENTIFICATION => $$PLSQL_UNIT,
                              P_TXT            => 'Systemfehler');
    RAISE;
END createJobs;


/*------------------------------------------------------------------------------
Name:         createDirWhenEmpty

Parameter:    pDirName   Name eines Directory-Objekts
              pPath      Physikalischer Pfad im Dateisystem

Beschreibung: Erzeugt ein physikalisches Verzeichnis im Dateisystem  und das
              dazugehörige Directory-Objekt, falls nicht vorhanden.
------------------------------------------------------------------------------*/
procedure createDirWhenEmpty (pDirName VARCHAR2, pPath VARCHAR2) IS

vSQL            VARCHAR2(4000);
vUnixCom        VARCHAR2(4000);
vDirExists      INTEGER;
vPathExists     INTEGER;
BEGIN

  SELECT count(*) INTO vDirExists from dba_DIRECTORIES where DIRECTORY_NAME = pDirName;

  IF vDirExists = 0 THEN
     logInfo('Verzeichnis-Objekt ' || pDirName || ' wird angelegt');
     vSQL :='CREATE OR REPLACE directory ' || pDirName || ' AS ' || chr(39) || pPath || chr(39);
     execute immediate vSQL;
  END IF;

  vPathExists := dbms_lob.fileexists(bfilename(pDirName, '.'));

  IF vPathExists = 0 THEN
     logInfo('Pfad im Dateisystem ' || pPath || ' wird angelegt');
     vUnixCom :=    'mkdir ' || pPath ||chr(10)||
                    'chmod ' || vMODE_DIR || ' ' || pPath;
    pkg_lib_ex_basic_newapi.CALLUNIX(vUnixCom, vEX_SESSION);


  END IF;

EXCEPTION
  WHEN OTHERS THEN
    pkg_lib_ex_basic_newapi.WRITELOG(P_EVENTTYPE      => pkg_lib_ex_basic_newapi.EventT_SysError,
                              P_SESSION        => vEX_SESSION,
                              P_IDENTIFICATION => $$PLSQL_UNIT,
                              P_TXT            => 'Systemfehler');
END createDirWhenEmpty;


/*------------------------------------------------------------------------------
Name:         createDirDefault

Parameter:    keine

Beschreibung: Erzeugt allgemeine Verzeichnisse im Dateisystem und die
              dazugehörigen Directory-Objekte.
------------------------------------------------------------------------------*/
procedure createDirDefault IS
vDirName        VARCHAR2(200);
vPath           VARCHAR2(2000);
vMainPath       VARCHAR2(2000);
BEGIN

    vMainPath := getMainPath;

    -- Hauptverzeichnis anlegen
    vDirName  := 'EX_EXP';
    vPath     := vMainPath;
    createDirWhenEmpty (vDirName , vPath );

    -- TEMP-Verzeichnis
    vDirName  := 'EX_TEMP';
    vPath     := vMainPath ||'/tmp';
    createDirWhenEmpty (vDirName , vPath );

     -- Verzeichnis fuer Steuerlisten
    vDirName  := 'EX_STL';
    vPath     := vMainPath || '/STL';
    createDirWhenEmpty (vDirName , vPath );

     -- Verzeichnis fuer Steuerlisten todo
    vDirName  := 'EX_STL_TODO';
    vPath     := vMainPath || '/STL/todo';
    createDirWhenEmpty (vDirName , vPath );

     -- Verzeichnis fuer Steuerlisten done
    vDirName  := 'EX_STL_DONE';
    vPath     := vMainPath || '/STL/done';
    createDirWhenEmpty (vDirName , vPath );

     -- Verzeichnis fuer Steuerlisten IAB
    vDirName  := 'EX_STL_IAB';
    vPath     := vMainPath || '/STL/iab';
    createDirWhenEmpty (vDirName , vPath );


EXCEPTION
  WHEN OTHERS THEN
    pkg_lib_ex_basic_newapi.WRITELOG(P_EVENTTYPE      => pkg_lib_ex_basic_newapi.EventT_SysError,
                              P_SESSION        => vEX_SESSION,
                              P_IDENTIFICATION => $$PLSQL_UNIT,
                              P_TXT            => 'Systemfehler');
END createDirDefault;

/*------------------------------------------------------------------------------
Name:         getTargetDir

Parameter:    pSTL_KZ      KZ des des Objejts
              pSTL_SCHEMA  Schemaname des Objejts

Return        vDirName    Diretory Name

Beschreibung: Funktion ermittelt Directory Name in welches geschrieben werden
              soll anhand von STL_KZ.
------------------------------------------------------------------------------*/
function  getTargetDir (pSTL_KZ     STAT_EXP_DATAMART_MCFG.WRK_EX_OBJ.STL_KZ%TYPE,
                        pSTL_SCHEMA STAT_EXP_DATAMART_MCFG.WRK_EX_OBJ.STL_SCHEMA%TYPE)
RETURN VARCHAR2
IS

vDirName  VARCHAR2(200);
BEGIN
    IF    pSTL_KZ = 'i' THEN
      vDirName   := 'EX_IAB';
    ELSIF pSTL_KZ = 'd' THEN
        vDirName := 'XD' || replace(upper(pSTL_SCHEMA),'_','');
    ELSIF pSTL_KZ = 'b' THEN
      vDirName   := 'XF' || replace(upper(pSTL_SCHEMA),'_','');
    ELSIF pSTL_KZ = 'f' THEN
      vDirName   := 'XF' || replace(upper(pSTL_SCHEMA),'_','');
    ELSIF pSTL_KZ = 'v' THEN
      vDirName   := 'XF' || replace(upper(pSTL_SCHEMA),'_','');
    ELSE
        vDirName := 'EX_TMP';
    END IF;

 RETURN vDirName;

EXCEPTION
  WHEN OTHERS THEN
    pkg_lib_ex_basic_newapi.WRITELOG(P_EVENTTYPE      => pkg_lib_ex_basic_newapi.EventT_SysError,
                              P_SESSION        => vEX_SESSION,
                              P_IDENTIFICATION => $$PLSQL_UNIT,
                              P_TXT            => 'Systemfehler');
END getTargetDir;

/*------------------------------------------------------------------------------
Name:         createDirJobs

Parameter:    keine

Beschreibung: Erzeugt jobspezifische Verzeichisse im Dateisystem und die
              dazugehörigen Directory-Objekte.
------------------------------------------------------------------------------*/
procedure createDirJobs ( p_STL_NAME VARCHAR2) IS
vDirName        VARCHAR2(200);
vPath           VARCHAR2(2000);
vMainPath       VARCHAR2(2000);
vExistsIAB      INTEGER;

BEGIN
    logInfo('Start erzeugen Verzeichnis-Objekte für Export, Falls nicht vorhanden');
    vMainPath := getMainPath;

    -- fuer alle normalen Tabellen
    FOR cDir IN
    (SELECT DISTINCT JOB_SCHEMA
       FROM STAT_EXP_DATAMART_MCFG.WRK_EX_JOB
      WHERE JOB_SESSION = vEX_SESSION
    )
    LOOP
      -- Hauptverzeichnis anlegen
      vDirName  := 'XE' || replace(upper(cDir.JOB_SCHEMA),'_','');
      vPath     := vMainPath || '/' || lower(cDir.JOB_SCHEMA);
      createDirWhenEmpty (vDirName , vPath );

      -- Verzeichnis für Dimensionen
      vDirName  := getTargetDir('d', cDir.JOB_SCHEMA);
      vPath     := vMainPath || '/' || lower(cDir.JOB_SCHEMA) || '/DIMENSIONEN';
      createDirWhenEmpty (vDirName , vPath );

       -- Verzeichnis fuer Fakten
      vDirName  := getTargetDir('f', cDir.JOB_SCHEMA);
      vPath     := vMainPath || '/' || lower(cDir.JOB_SCHEMA)|| '/FAKTEN';
      createDirWhenEmpty (vDirName , vPath );

      -- Verzeichnis fuer IAB
      vExistsIAB  := dbms_lob.fileexists(bfilename('EX_STL_IAB',  p_STL_NAME));
      IF vExistsIAB <> 0 THEN
        vDirName  := 'EX_IAB';
        vPath     := vMainPath || '/EXPORT_IAB';
        createDirWhenEmpty (vDirName , vPath );
      END IF;

    END LOOP;
    logInfo('Ende erzeugen Verzeichnis-Objekte (Falls nicht vorhanden) für Export');
EXCEPTION
  WHEN OTHERS THEN
    pkg_lib_ex_basic_newapi.WRITELOG(P_EVENTTYPE      => pkg_lib_ex_basic_newapi.EventT_SysError,
                              P_SESSION        => vEX_SESSION,
                              P_IDENTIFICATION => $$PLSQL_UNIT,
                              P_TXT            => 'Systemfehler');
END createDirJobs;

/*------------------------------------------------------------------------------
Name:         getParallel

Parameter:    pJob Record vom Typ Job.

Return:       anzal parallel

Beschreibung: Ermittelt Parallelität mit welcher ein Job exportiert werden soll.
              1/2 MaxCPU von DB wenn Mehr als 100 Mio Datensätze
              2/6 MaxCPU von DB wenn Mehr als 10  Mio Datensätze
              1/6 MaxCPU von DB wenn Mehr als 1   Mio Datensätze
------------------------------------------------------------------------------*/
FUNCTION getParallel (pJob STAT_EXP_DATAMART_MCFG.WRK_EX_JOB%ROWTYPE)
  RETURN INTEGER
AS
  pPar INTEGER;
  vSQL      VARCHAR2(200);
  vAnz      STAT_EXP_DATAMART_MCFG.WRK_EX_JOB.JOB_ANZ_MAT%TYPE;
BEGIN

  IF pJob.JOB_OBJ_TYPE IN ('t', 'm') THEN
    vSQL := 'SELECT count(*) FROM ' || pJob.JOB_SCHEMA || '.' || pJob.JOB_OBJ_NAME;
    execute immediate vSQL INTO vAnz;
  ELSE
    SELECT JOB_ANZ_MAT INTO vAnz FROM STAT_EXP_DATAMART_MCFG.WRK_EX_JOB WHERE  JOB_ID = pJob.JOB_ID;
  END IF;

  pPar := round(getMaxCPU / 2);


  IF vAnz >= 100000000 THEN                         -- ueber 100 Mio
    RETURN pPar;
  ELSIF vAnz >= 10000000 AND vAnz < 100000000 THEN  -- 10 Mio bis 100 Mio
    pPar := round(pPar * 2 / 3);
    RETURN pPar;
  ELSIF vAnz >= 1000000 AND vAnz < 10000000  THEN   -- 1 Mio bis 10 Mio
    pPar := round(pPar / 3);
    RETURN pPar;
  ELSE
    RETURN 1;
  END IF ;
EXCEPTION
WHEN OTHERS THEN
      pkg_lib_ex_basic_newapi.WRITELOG(P_EVENTTYPE      => pkg_lib_ex_basic_newapi.EventT_SysError,
                              P_SESSION        => vEX_SESSION,
                              P_IDENTIFICATION => $$PLSQL_UNIT,
                              P_TXT            => 'Systemfehler');
END getParallel;

/*------------------------------------------------------------------------------
Name:         setStatusJOB

Parameter:    pJobID      ID eines Jobs
              pJobStatus  gewünschter Status.
              pForce      zwingend

Beschreibung: Setzt gewünschten Status fuer einen Job, falls dieser nicht schon
              den Status "Fehlerhaft" hat.
              Dieses verhalten kann auch mit pForce übersteuert werden.
------------------------------------------------------------------------------*/
procedure setStatusJOB (pJobID STAT_EXP_DATAMART_MCFG.WRK_EX_JOB.JOB_ID%TYPE,
                       pJobStatus STAT_EXP_DATAMART_MCFG.WRK_EX_JOB.JOB_STATUS%TYPE,
                       pForce BOOLEAN DEFAULT FALSE) IS

vJOB_ID STAT_EXP_DATAMART_MCFG.WRK_EX_JOB.JOB_ID%TYPE;

BEGIN
  vJOB_ID := pJobID;
  IF pForce THEN
     UPDATE STAT_EXP_DATAMART_MCFG.WRK_EX_JOB SET JOB_STATUS = pJobStatus WHERE JOB_ID = pJobID;
  ELSE
     UPDATE STAT_EXP_DATAMART_MCFG.WRK_EX_JOB SET JOB_STATUS = pJobStatus WHERE JOB_ID = pJobID AND JOB_STATUS NOT IN ( 'I', 'F' );
  END IF;

EXCEPTION
  WHEN OTHERS THEN
    pkg_lib_ex_basic_newapi.WRITELOG(P_EVENTTYPE      => pkg_lib_ex_basic_newapi.EventT_SysError,
                              P_SESSION        => vEX_SESSION,
                              P_JOB_ID         => vCurrentJobId,
                              P_IDENTIFICATION => $$PLSQL_UNIT || '_JOB_ID_' || vJOB_ID,
                              P_TXT            => 'Systemfehler');
END setStatusJOB;

/*------------------------------------------------------------------------------
Name:         matJobObject

Parameter:    pJob Record vom Typ Job.

Beschreibung: Materialisiert die zu exportierenden Objekte.
------------------------------------------------------------------------------*/
procedure matJobObject(pJob STAT_EXP_DATAMART_MCFG.WRK_EX_JOB%ROWTYPE) IS

vJOB_ID    NUMBER(20,0);
vSQL       VARCHAR2(4000);
vTableName VARCHAR2(128);
pPar  INTEGER := round(getMaxCPU / 2);
vAnzMat    STAT_EXP_DATAMART_MCFG.WRK_EX_JOB.JOB_ANZ_MAT%TYPE;
vAnzMatTM  STAT_EXP_DATAMART_MCFG.WRK_EX_JOB.JOB_ANZ_MAT%TYPE;
vJOB_PART_COL VARCHAR2(128);
vPART_CLAUSE VARCHAR2(1000);
BEGIN
  -- BIDW-495 - SeidelD006 - 20240916
  -- Umstellung auf GTTs

  vJOB_ID :=pJob.JOB_ID;
  vTableName := 'EXP' || pJob.JOB_ID;
  IF pJob.JOB_OBJ_TYPE IN ('t', 'm') THEN
    vSQL :='CREATE GLOBAL TEMPORARY TABLE ' || vDataSchema || '.EXP'
                           || pJob.JOB_ID || ' ON COMMIT PRESERVE ROWS AS SELECT * FROM '
                           || pJob.JOB_SCHEMA || '.'
                           || pJob.JOB_OBJ_NAME
                           || ' WHERE ROWNUM = 1 ';
  ELSIF pJob.JOB_OBJ_TYPE = 'p' AND pJob.JOB_PART != 'NO' THEN
    vJOB_PART_COL := getFormatColMat(pJob.JOB_SCHEMA, pJob.JOB_OBJ_NAME, pJob.JOB_PART_COL);
    if pJob.JOB_PART ='null'
    then
        vPART_CLAUSE :=  vJOB_PART_COL || ' IS NULL';
    else
         vPART_CLAUSE :=  vJOB_PART_COL || ' = ' || pJob.JOB_PART;
    end if;
    vSQL :='CREATE GLOBAL TEMPORARY TABLE ' || vDataSchema || '.EXP'
                           || pJob.JOB_ID || ' ON COMMIT PRESERVE ROWS parallel ' || pPar || ' AS SELECT * FROM '
                           || pJob.JOB_SCHEMA || '.'
                           || pJob.JOB_OBJ_NAME || ' WHERE '
                           || vPART_CLAUSE;
  ELSIF upper(pJob.JOB_OBJ_NAME) = 'DUAL' THEN
    vSQL :='CREATE GLOBAL TEMPORARY TABLE ' || vDataSchema || '.EXP'
                           || pJob.JOB_ID || ' ON COMMIT PRESERVE ROWS AS SELECT * FROM '
                           || pJob.JOB_OBJ_NAME;
  ELSE
    vSQL :='CREATE GLOBAL TEMPORARY TABLE ' || vDataSchema || '.EXP'
                           || pJob.JOB_ID || ' ON COMMIT PRESERVE ROWS parallel ' || pPar || ' AS SELECT * FROM '
                           || pJob.JOB_SCHEMA || '.'
                           || pJob.JOB_OBJ_NAME;
  END IF;
  logInfo('DDL:'||vSQL);
  execute immediate vSQL;

  --DBMS_STATS.GATHER_TABLE_STATS (vDataSchema, vTableName );
  
  -- NUM_ROWS wird bei GTTs nicht gepflegt
  vSQL := 'SELECT /*+ PARALLEL */ COUNT(*) FROM '||vDataSchema||'.'||vTableName;
  execute immediate vSQL into vAnzMat;
  
  -- SELECT NUM_ROWS INTO vAnzMat FROM dba_tables WHERE TABLE_NAME = vTableName AND OWNER = vDataSchema;
  IF pJob.JOB_OBJ_TYPE IN ('t', 'm') THEN
    vAnzMat := 0;
  END IF;
  UPDATE STAT_EXP_DATAMART_MCFG.WRK_EX_JOB SET JOB_ANZ_MAT = vAnzMat WHERE JOB_ID = vJOB_ID;

EXCEPTION
  WHEN OTHERS THEN
    setStatusJOB(vJOB_ID, 'F');
    pkg_lib_ex_basic_newapi.WRITELOG(P_EVENTTYPE      => pkg_lib_ex_basic_newapi.EventT_SysError,
                              P_SESSION        => vEX_SESSION,
                              P_JOB_ID         => vCurrentJobId,
                              P_IDENTIFICATION => $$PLSQL_UNIT || '_JOB_ID_' || vJOB_ID,
                              P_TXT            => 'Systemfehler');
END matJobObject;


/*------------------------------------------------------------------------------
Name:      replaceName_b

Parameter:    pName String

Return:       Dateinamen

Beschreibung: Funktion, ersetzt Prefix eines Objejtnamens für Biodata-Export
------------------------------------------------------------------------------*/
  FUNCTION replaceName_b(
      pName VARCHAR2)
    RETURN VARCHAR2
  AS
    vName  VARCHAR2(4000);
  BEGIN
        vName := pName;
        vName := replace(vName, 'vf_' , 'tf_');
        vName := replace(vName, 'vh_' , 'th_');
        vName := replace(vName, 'vd_' , 'td_');
        vName := replace(vName, 'vr_' , 'tr_');
        vName := replace(vName, 'vt_' , 'tt_');
      RETURN vName;
EXCEPTION
  WHEN OTHERS THEN
    pkg_lib_ex_basic_newapi.WRITELOG(P_EVENTTYPE      => pkg_lib_ex_basic_newapi.EventT_SysError,
                              P_SESSION        => vEX_SESSION,
                              P_JOB_ID         => vCurrentJobId,
                              P_IDENTIFICATION => $$PLSQL_UNIT || '_PName_' || pName,
                              P_TXT            => 'Systemfehler');
END replaceName_b;


/*------------------------------------------------------------------------------
Name:      replaceName_v

Parameter:    pName String

Return:       Dateinamen

Beschreibung: Function, ersetzt Prefix eines Objejtnamens für View-Export
------------------------------------------------------------------------------*/
  FUNCTION replaceName_v(
      pName VARCHAR2)
    RETURN VARCHAR2
  AS
    vName  VARCHAR2(4000);
  BEGIN
        vName := pName;
        vName := replace(vName, 've_' , 'te_');
        vName := replace(vName, 'vf_' , 'tf_');
        vName := replace(vName, 'vh_' , 'th_');
        vName := replace(vName, 'vd_' , 'td_');
        vName := replace(vName, 'vr_' , 'tr_');
        vName := replace(vName, 'vt_' , 'tt_');
      RETURN vName;
EXCEPTION
  WHEN OTHERS THEN
    pkg_lib_ex_basic_newapi.WRITELOG(P_EVENTTYPE      => pkg_lib_ex_basic_newapi.EventT_SysError,
                              P_SESSION        => vEX_SESSION,
                              P_JOB_ID         => vCurrentJobId,
                              P_IDENTIFICATION => $$PLSQL_UNIT || '_PName_' || pName,
                              P_TXT            => 'Systemfehler');
END replaceName_v;



/*------------------------------------------------------------------------------
Name:         dumpJobDDL

Parameter:    pJob Record vom Typ Job.

Beschreibung: Erzeugt das DDL für das zu exportierende DB-Objekt
------------------------------------------------------------------------------*/
procedure dumpJobDDL (pJob STAT_EXP_DATAMART_MCFG.WRK_EX_JOB%ROWTYPE) IS
l_file        UTL_FILE.FILE_TYPE;
vSQL          VARCHAR2(32767);
vSQLStmt      VARCHAR2(32767);
vTabName      VARCHAR2(128);
vFilename     VARCHAR2(200);
vObj          VARCHAR2(128);
vObj_Tmp      VARCHAR2(128);
vView         VARCHAR2(128);
vRepl1        VARCHAR2(128);
vRepl2        VARCHAR2(128);
vBis          INTEGER;
vJOB_ID       NUMBER(20,0);
vSchema       VARCHAR2(128);
vDirName      VARCHAR2(200);
vPath         VARCHAR2(2000);
vUnixCom      VARCHAR2(4000);
vTransHandle  PLS_INTEGER;

BEGIN
  vJOB_ID   := pJob.JOB_ID;

  vTransHandle := DBMS_METADATA.SESSION_TRANSFORM;

  DBMS_METADATA.SET_TRANSFORM_PARAM(vTransHandle,'DEFAULT',TRUE);
  DBMS_METADATA.SET_TRANSFORM_PARAM(vTransHandle,'PRETTY',TRUE);
  DBMS_METADATA.SET_TRANSFORM_PARAM(vTransHandle,'SEGMENT_ATTRIBUTES',FALSE);
  DBMS_METADATA.SET_TRANSFORM_PARAM(vTransHandle,'CONSTRAINTS',TRUE);

  IF pJob.JOB_OBJ_TYPE = 't' THEN
    vSchema   := vDataSchema;
    vFilename := pJob.JOB_OBJ_NAME || '.schema';
    vObj      := 'EXP' || pJob.JOB_ID;
    vSQL      := DBMS_METADATA.GET_DDL ('TABLE', vObj , vSchema);
  ELSIF pJob.JOB_OBJ_TYPE = 'p' THEN
    vSchema   := vDataSchema;
    vFilename := pJob.JOB_OBJ_NAME || '_' || pJob.JOB_PART || '.schema';
    vObj      := 'EXP' || pJob.JOB_ID;
    vSQL      := DBMS_METADATA.GET_DDL ('TABLE', vObj, vSchema);
    IF  pJob.JOB_KZ = 'v' THEN
        -- Objektprefix ändern  (vf_ >> ts)
        vFilename := replaceName_v(vFilename);
    END IF;
  ELSIF  pJob.JOB_OBJ_TYPE = 'v' AND pJob.JOB_KZ = 'v' AND vVIEW_MAT THEN
    vSchema   := UPPER(pJob.JOB_SCHEMA);
    vFilename := pJob.JOB_OBJ_NAME || '.schema';
    vObj_Tmp  :=  'EXP' || pJob.JOB_ID;
    vObj      :=  UPPER(pJob.JOB_OBJ_NAME);
    vSQL      := DBMS_METADATA.GET_DDL ('TABLE' , vObj_Tmp, vDataSchema );
    vSQL      := REPLACE(vSQL,vObj_Tmp, vObj);
    vSQL      := REPLACE(vSQL,vDataSchema, vSchema);
-- --8<-- BIDW-457 Steuerlistentyp "f", Tabellen-DDL an Stelle von View-DDL erzeugen
--  ELSIF  pJob.JOB_OBJ_TYPE = 'v' AND pJob.JOB_KZ = 'd' AND vVIEW_MAT THEN
    ELSIF  pJob.JOB_OBJ_TYPE = 'v' AND pJob.JOB_KZ in ('d','f') AND vVIEW_MAT THEN
-- -->8-- BIDW-457 Steuerlistentyp "f", Tabellen-DDL an Stelle von View-DDL erzeugen
    vSchema   := UPPER(pJob.JOB_SCHEMA);
    vFilename := pJob.JOB_OBJ_NAME || '.schema';
    vObj_Tmp  :=  'EXP' || pJob.JOB_ID;
    vObj      :=  UPPER(pJob.JOB_OBJ_NAME);
    vSQL      := DBMS_METADATA.GET_DDL ('TABLE' , vObj_Tmp, vDataSchema );
    vSQL      := REPLACE(vSQL,vObj_Tmp, vObj);
    vSQL      := REPLACE(vSQL,vDataSchema, vSchema);
-- --8<-- BIDW-457 Steuerlistentyp "f", Tabellen-DDL an Stelle von View-DDL erzeugen
--  ELSIF  pJob.JOB_OBJ_TYPE = 'v' AND pJob.JOB_KZ = 'd' AND NOT vVIEW_MAT THEN
  ELSIF  pJob.JOB_OBJ_TYPE = 'v' AND pJob.JOB_KZ in ('d','f') AND NOT vVIEW_MAT THEN
-- -->8-- BIDW-457 Steuerlistentyp "f", Tabellen-DDL an Stelle von View-DDL erzeugen
    vSchema   := UPPER(pJob.JOB_SCHEMA);
    vFilename := pJob.JOB_OBJ_NAME || '.schema';
    vObj      :=  UPPER(pJob.JOB_OBJ_NAME);
    vSQL      := DBMS_METADATA.GET_DDL ('VIEW' , vObj, vSchema );
-- --8<-- BIDW-457 Steuerlistentyp "f", Tabellen-DDL an Stelle von View-DDL erzeugen
--  ELSIF  pJob.JOB_OBJ_TYPE = 'v' AND pJob.JOB_KZ = 'f' THEN
--    vSchema   := UPPER(pJob.JOB_SCHEMA);
--    vFilename := pJob.JOB_OBJ_NAME || '.schema';
--    vObj      :=  UPPER(pJob.JOB_OBJ_NAME);
--    vSQL      := DBMS_METADATA.GET_DDL ('VIEW' , vObj, vSchema );
-- -->8-- BIDW-457 Steuerlistentyp "f", Tabellen-DDL an Stelle von View-DDL erzeugen
  ELSIF pJob.JOB_OBJ_TYPE = 's' THEN
    vSchema   := vDataSchema;
    vFilename := pJob.JOB_OBJ_NAME || '.schema';
    vObj      := 'EXP' || pJob.JOB_ID;
    vSQL      := DBMS_METADATA.GET_DDL ('TABLE', vObj, vSchema);
  ELSE
    vSchema   := vDataSchema;
    vFilename := pJob.JOB_OBJ_NAME || '.schema';
    vObj      := 'EXP' || pJob.JOB_ID;
    vSQL      := DBMS_METADATA.GET_DDL ('TABLE', vObj , vSchema);
  END IF;


  -- Schemanamen entfernen
  vRepl1    := chr(34) || vSchema || chr(34) || '.' || chr(34) || vObj || chr(34);
  vRepl2    := chr(34) || UPPER(pJob.JOB_OBJ_NAME)  || chr(34);
  vSQL      := REPLACE(vSQL,vRepl1, vRepl2) || ';';


    -- bei Export von Partitionen den Partitionsschlüssel mit den Objektnamen nehmen
  IF pJob.JOB_OBJ_TYPE = 'p' THEN
    vTabName  := (pJob.JOB_OBJ_NAME || '_' || pJob.JOB_PART);
    -- Umbennenen der View in Fakt
    IF  pJob.JOB_KZ = 'v' THEN
        vTabName := replaceName_v(vTabName);
    END IF;
    vSQL      := REPLACE(vSQL,upper(pJob.JOB_OBJ_NAME), upper(vTabName));
  END IF;

  -- BIDW-495 - SeidelD006 - 20240916
  -- Umstellung auf GTTs
  -- GTT wieder aus CREATE -Table herausschneiden
  vSQL:=replace(replace(vSQL,'GLOBAL TEMPORARY ',''),'ON COMMIT PRESERVE ROWS ','');

  -- PARALLEL nn am Ende herausschneiden
  vSQL:=regexp_replace(vSQL,'\)\s*PARALLEL\s*[0-9]*\s*;',');',1,0,'im');

  l_file := FOPEN(
            location => 'EX_TEMP',
            filename => vFilename
            /*,open_mode => 'w',
            max_linesize => 32767*/);

  UTL_FILE.PUT_LINE(l_file, vSQL);
  UTL_FILE.FCLOSE(l_file);

  vDirName  := getTargetDir (pJob.JOB_KZ, pJob.JOB_SCHEMA);
  UTL_FILE.FRENAME('EX_TEMP', vFilename, vDirName, vFilename, TRUE);

--  vPath := pkg_lib_ex_basic_newapi.getPath(vDirName);
--  vUnixCom := 'chmod ' || vMODE || ' ' || vPath || '/' || vFilename;
--  pkg_lib_ex_basic_newapi.CALLUNIX(vUnixCom, vEX_SESSION);

EXCEPTION
  WHEN UTL_FILE.INVALID_PATH THEN
    pkg_lib_ex_basic_newapi.WRITELOG(P_EVENTTYPE      => pkg_lib_ex_basic_newapi.EventT_FuncError,
                              P_SESSION        => vEX_SESSION,
                              P_JOB_ID         => vCurrentJobId,
                              P_IDENTIFICATION => $$PLSQL_UNIT || '_JOB_ID_' || vJOB_ID,
                              P_TXT            => 'Fehler, Pfad für DDL-Datei ist nicht valide');
  WHEN UTL_FILE.RENAME_FAILED THEN
    pkg_lib_ex_basic_newapi.WRITELOG(P_EVENTTYPE      => pkg_lib_ex_basic_newapi.EventT_FuncError,
                              P_SESSION        => vEX_SESSION,
                              P_JOB_ID         => vCurrentJobId,
                              P_IDENTIFICATION => $$PLSQL_UNIT || '_JOB_ID_' || vJOB_ID,
                              P_TXT            => 'Fehler, DDL-Datei kann nicht umbenannt werden');
  WHEN UTL_FILE.INVALID_OPERATION THEN
    pkg_lib_ex_basic_newapi.WRITELOG(P_EVENTTYPE      => pkg_lib_ex_basic_newapi.EventT_FuncError,
                              P_SESSION        => vEX_SESSION,
                              P_JOB_ID         => vCurrentJobId,
                              P_IDENTIFICATION => $$PLSQL_UNIT || '_JOB_ID_' || vJOB_ID,
                              P_TXT            => 'Fehler, DDL-Datei kann nicht gelesen werden');
  WHEN OTHERS THEN
    UTL_FILE.FCLOSE(l_file);
    setStatusJOB(vJOB_ID, 'F');
    pkg_lib_ex_basic_newapi.WRITELOG(P_EVENTTYPE      => pkg_lib_ex_basic_newapi.EventT_SysError,
                              P_SESSION        => vEX_SESSION,
                              P_JOB_ID         => vCurrentJobId,
                              P_IDENTIFICATION => $$PLSQL_UNIT || '_JOB_ID_' || vJOB_ID,
                              P_TXT            => 'Systemfehler');
END dumpJobDDL;

/*------------------------------------------------------------------------------
Name:      getFilenameZIP

Parameter:    pJob Record vom Typ Job.

Return:       GZIP-Dateinamen

Beschreibung: Funktion, erstellt den Namen der GZIP-Datei
------------------------------------------------------------------------------*/
  FUNCTION getFilenameZIP(
      pJob STAT_EXP_DATAMART_MCFG.WRK_EX_JOB%ROWTYPE, pCount INTEGER)
    RETURN VARCHAR2
  AS
    vFilenameZIP  VARCHAR2(200);
  BEGIN

    IF pJob.JOB_PART = 'NO' THEN
      vFilenameZIP := pJob.JOB_OBJ_NAME || '.' || pCount || '.unl.gz';
    ELSE
      vFilenameZIP := pJob.JOB_OBJ_NAME || '_' || pJob.JOB_PART || '.' || pCount || '.unl.gz';
    END IF;

    IF pJob.JOB_KZ = 'b' THEN
        vFilenameZIP := replaceName_b(vFilenameZIP);
    ELSIF pJob.JOB_KZ = 'v' THEN
        vFilenameZIP := replaceName_v(vFilenameZIP);
    END IF;

    RETURN vFilenameZIP;
  EXCEPTION
  WHEN OTHERS THEN
    pkg_lib_ex_basic_newapi.WRITELOG(P_EVENTTYPE      => pkg_lib_ex_basic_newapi.EventT_SysError,
                              P_SESSION        => vEX_SESSION,
                              P_JOB_ID         => vCurrentJobId,
                              P_IDENTIFICATION => $$PLSQL_UNIT,
                              P_TXT            => 'Systemfehler');
  END getFilenameZIP;

/*------------------------------------------------------------------------------
Name:         alterSessionDME

Parameter:    keine

Beschreibung: Funktion, setzt diverse Parameter zur Session z.B Datumsformate.
              Variante DME
------------------------------------------------------------------------------*/
PROCEDURE alterSessionDME
AS
pragma autonomous_transaction;
  l_alter_sess_stmt VARCHAR2(500);
  vVersion INTEGER;
BEGIN
  l_alter_sess_stmt := 'alter session set "parallel_force_local" = true';

  EXECUTE IMMEDIATE l_alter_sess_stmt;

  l_alter_sess_stmt := 'alter session set NLS_DATE_FORMAT = ''dd.mm.yyyy''';
  --l_alter_sess_stmt := 'alter session set NLS_DATE_FORMAT = ''yyyymmdd''';

  EXECUTE IMMEDIATE l_alter_sess_stmt;

  l_alter_sess_stmt := 'alter session set NLS_NUMERIC_CHARACTERS = ''.,'''; -- Amerikanisches Format für IFX

  EXECUTE IMMEDIATE l_alter_sess_stmt;

  l_alter_sess_stmt := 'alter session set NLS_TIMESTAMP_FORMAT = ''yyyy-mm-dd HH24:MI:SS''';

  EXECUTE IMMEDIATE l_alter_sess_stmt;

  SELECT SUBSTR( version,1,2) INTO vVersion FROM V$INSTANCE;

  -- Ab Datenbankversion 12.... für BUG in PLSQL-Cursor-Verarbeitung (LIMIT 100)
  IF vVersion > 11 THEN
    l_alter_sess_stmt := 'alter session set "_rdbms_internal_fplib_enabled" = TRUE';
    EXECUTE IMMEDIATE l_alter_sess_stmt;
  END IF;

  EXCEPTION
  WHEN OTHERS THEN
    pkg_lib_ex_basic_newapi.WRITELOG(P_EVENTTYPE      => pkg_lib_ex_basic_newapi.EventT_SysError,
                              P_SESSION        => vEX_SESSION,
                              P_JOB_ID         => vCurrentJobId,
                              P_IDENTIFICATION => $$PLSQL_UNIT,
                              P_TXT            => 'Systemfehler');
  END alterSessionDME;

/*------------------------------------------------------------------------------
Name:         alterSessionIAB

Parameter:    keine

Beschreibung: Function, setzt diverse Parameter zur Session z.B Datumsformate.
              Variante DME
------------------------------------------------------------------------------*/
PROCEDURE alterSessionIAB
AS
pragma autonomous_transaction;
  l_alter_sess_stmt VARCHAR2(500);
  vVersion INTEGER;
BEGIN
  l_alter_sess_stmt := 'alter session set "parallel_force_local" = true';

  EXECUTE IMMEDIATE l_alter_sess_stmt;

  l_alter_sess_stmt := 'alter session set NLS_DATE_FORMAT = ''yyyy-mm-dd''';
  --l_alter_sess_stmt := 'alter session set NLS_DATE_FORMAT = ''yyyymmdd''';

  EXECUTE IMMEDIATE l_alter_sess_stmt;

  l_alter_sess_stmt := 'alter session set NLS_NUMERIC_CHARACTERS = ''.,'''; -- Amerikanisches Format für IFX

  EXECUTE IMMEDIATE l_alter_sess_stmt;

  l_alter_sess_stmt := 'alter session set NLS_TIMESTAMP_FORMAT = ''yyyy-mm-dd HH24:MI:SS''';

  EXECUTE IMMEDIATE l_alter_sess_stmt;

  SELECT SUBSTR( version,1,2) INTO vVersion FROM V$INSTANCE;

  -- Ab Datenbankversion 12.... für BUG in PLSQL-Cursor-Verarbeitung (LIMIT 100)
  IF vVersion > 11 THEN
    l_alter_sess_stmt := 'alter session set "_rdbms_internal_fplib_enabled" = TRUE';
    EXECUTE IMMEDIATE l_alter_sess_stmt;
  END IF;

  EXCEPTION
  WHEN OTHERS THEN
    pkg_lib_ex_basic_newapi.WRITELOG(P_EVENTTYPE      => pkg_lib_ex_basic_newapi.EventT_SysError,
                              P_SESSION        => vEX_SESSION,
                              P_JOB_ID         => vCurrentJobId,
                              P_IDENTIFICATION => $$PLSQL_UNIT,
                              P_TXT            => 'Systemfehler');
  END alterSessionIAB;

/*------------------------------------------------------------------------------
Name:         dumpJob2CVS

Parameter:    pJob Record vom Typ Job.

Beschreibung: Exportiert dasw materialisierte DB-Objekt als Datei in Filesystem.
              Die Datei wird gezippt.
------------------------------------------------------------------------------*/
procedure dumpJob2CVS (pJob STAT_EXP_DATAMART_MCFG.WRK_EX_JOB%ROWTYPE) IS
vFilenameZIP  VARCHAR2(200);
vFilenameSH   VARCHAR2(200);
vFilename     VARCHAR2(200);
vDirName      VARCHAR2(200);
vTabname      VARCHAR2(128);
vTabnamePipe  VARCHAR2(200);
vUnixCom      VARCHAR2(4000);
vPathTemp     VARCHAR2(2000);
vPath         VARCHAR2(2000);
anz           STAT_EXP_DATAMART_MCFG.WRK_EX_JOB.JOB_ANZ_DUMP%TYPE;
vJOB_ID       NUMBER(20,0);
pPar          INTEGER;
l_file        UTL_FILE.FILE_TYPE;

BEGIN
  vJOB_ID := pJob.JOB_ID;
  vTabname     := 'EXP' || pJob.JOB_ID;

  -- wenn Parallelitaet nicht vorgegeben dann automatisch ermitteln
  IF vParallel IS NULL THEN
    pPar := getParallel(pJob);
  ELSE
    pPar := vParallel;
  END IF;

-- Dateien vorher mit den richtigen Rechten anlegen
--  vFilenameSH := 'EXP' || pJob.JOB_ID|| '.sh';
--  l_file := FOPEN(
--            location => 'EX_TEMP',
--            filename => vFilenameSH
--            /*,open_mode => 'w',
--            max_linesize => 32767*/);

--  UTL_FILE.PUT_LINE(l_file, '#!/bin/ksh');
--  UTL_FILE.PUT_LINE(l_file, 'PATH=/bin:/usr/bin:.');

  vUnixCom:=null;

  -- Dateien vorbereiten und anlegen
  FOR i IN 1..pPar
  LOOP
    vPathTemp    := pkg_lib_ex_basic_newapi.getPath('EX_TEMP');
    vTabnamePipe := 'EXP' || pJob.JOB_ID || '.' || i || '.pipe';
    vFilenameZIP := getFilenameZIP(pJob, i);

    -- named Pipe anlegen.
--    UTL_FILE.PUT_LINE(l_file, 'rm '  || vPathTemp || '/' || vTabnamePipe||' 2>/dev/null');
--    UTL_FILE.PUT_LINE(l_file, 'touch '  || vPathTemp || '/' || vTabnamePipe);
--    UTL_FILE.PUT_LINE(l_file, 'chmod ' || vMODE || ' ' || vPathTemp || '/' || vTabnamePipe);

      vUnixCom:=vUnixcom||chr(10)||'rm '  || vPathTemp || '/' || vTabnamePipe||' 2>/dev/null';
      vUnixCom:=vUnixcom||chr(10)||'touch '  || vPathTemp || '/' || vTabnamePipe;
      vUnixCom:=vUnixcom||chr(10)||'chmod ' || vMODE || ' ' || vPathTemp || '/' || vTabnamePipe;


  END LOOP;

  -- SH-Skript ausfuerbar
--  vUnixCom := 'chmod ' || vMODE_EXE || ' ' || vPathTemp || '/' || vFilenameSH;
--  pkg_lib_ex_basic_newapi.CALLUNIX(vUnixCom, vEX_SESSION);

  -- SH-Skript ausfuehren
--  vUnixCom := 'sh ' || chr(34) || vPathTemp || '/' || vFilenameSH || chr(34);
  pkg_lib_ex_basic_newapi.CALLUNIX(vUnixCom, vEX_SESSION);




  IF pJob.JOB_OBJ_TYPE IN ('t', 'm') THEN
     anz := PKG_EX_DUMP_CSV.tab2csv_parallel(p_directory => 'EX_TEMP',
                                             p_filename  => vTabname,
                                             p_schema    => upper(pJob.JOB_SCHEMA),
                                             p_tabname   => upper(pJob.JOB_OBJ_NAME),
                                             p_recordend => chr(10),
                                             p_charset => case when vDoUTF8UNLOAD then vUTF8Charset else vDefaultCharset end,
                                             p_parallel  => pPar,
                                             p_delimiter => vDelimiter);
  ELSE
     anz := PKG_EX_DUMP_CSV.tab2csv_parallel(p_directory => 'EX_TEMP',
                                             p_filename  => vTabname,
                                             p_schema    => vDataSchema,
                                             p_tabname   => vTabname,
                                             p_recordend => chr(10),
                                             p_charset => case when vDoUTF8UNLOAD then vUTF8Charset else vDefaultCharset end,
                                             p_parallel  => pPar,
                                             p_delimiter => vDelimiter);
  END IF;

  UPDATE STAT_EXP_DATAMART_MCFG.WRK_EX_JOB SET JOB_ANZ_DUMP = anz WHERE JOB_ID = vJOB_ID;

--  vFilenameSH := 'EXP' || pJob.JOB_ID|| '.sh';
--  l_file := FOPEN(
--            location => 'EX_TEMP',
--            filename => vFilenameSH
--            /*,open_mode => 'w',
--            max_linesize => 32767*/);

--  UTL_FILE.PUT_LINE(l_file, '#!/bin/ksh');
--  UTL_FILE.PUT_LINE(l_file, 'PATH=/bin:/usr/bin:.');

  vUnixCom:=null;

  -- entladene Dateien umbenennen  Und ZIP-Aufreg erstellen
  FOR i IN 1..pPar
  LOOP
    vPathTemp    := pkg_lib_ex_basic_newapi.getPath('EX_TEMP');
    vTabnamePipe := 'EXP' || pJob.JOB_ID || '.' || i || '.pipe';

    IF pJob.JOB_PART = 'NO' THEN
      vFilename    := pJob.JOB_OBJ_NAME || '.' || i || '.unl';
    ELSE
      vFilename    := pJob.JOB_OBJ_NAME || '_' || pJob.JOB_PART || '.' || i || '.unl';
    END IF;

    IF pJob.JOB_KZ = 'b' THEN
        vFilename := replaceName_b(vFilename);
    ELSIF pJob.JOB_KZ = 'v' THEN
        vFilename := replaceName_v(vFilename);
    END IF;

    -- umbenennen
    UTL_FILE.FRENAME('EX_TEMP', vTabnamePipe, 'EX_TEMP', vFilename, TRUE);

    -- ZIP-Auftrag in SH
--    UTL_FILE.PUT_LINE(l_file, 'gzip ' || vPathTemp || '/' || vFilename || ' &');
    vUnixCOm:=vUnixCom||chr(10)||'gzip ' || vPathTemp || '/' || vFilename || ' &';

  END LOOP;

--  UTL_FILE.FCLOSE(l_file);

  -- SH-Skript ausfuerbar
--  vUnixCom := 'chmod ' || vMODE_EXE || ' ' || vPathTemp || '/' || vFilenameSH;
--  pkg_lib_ex_basic_newapi.CALLUNIX(vUnixCom, vEX_SESSION);

  -- SH-Skript ausfuehren
--  vUnixCom := 'sh ' || chr(34) || vPathTemp || '/' || vFilenameSH || chr(34);
  pkg_lib_ex_basic_newapi.CALLUNIX(vUnixCom, vEX_SESSION);

  -- SH-Skript loeschen
  --vUnixCom := 'rm ' || vPathTemp || '/' || vFilenameSH;
  --pkg_lib_ex_basic_newapi.CALLUNIX(vUnixCom, vEX_SESSION);
  UTL_FILE.FREMOVE('EX_TEMP','vFilenameSH');

  vDirName  := getTargetDir (pJob.JOB_KZ, pJob.JOB_SCHEMA);
  vPath := pkg_lib_ex_basic_newapi.getPath(vDirName);

  -- loeschen alte Gezippte Dateien und evtl. vorhandene Semaphoren
  vUnixCom := 'find '|| vPath || ' -type f -iname ' || pJob.JOB_OBJ_NAME || '.*.unl.*' || ' -delete';
  pkg_lib_ex_basic_newapi.CALLUNIX(vUnixCom, vEX_SESSION);


  -- Gezippte Dateien verschieben und Rechte einstellen
  FOR i IN 1..pPar
  LOOP
    IF pJob.JOB_PART = 'NO' THEN
      vFilenameZIP := pJob.JOB_OBJ_NAME || '.' || i || '.unl.gz';
    ELSE
      vFilenameZIP := pJob.JOB_OBJ_NAME || '_' || pJob.JOB_PART || '.' || i || '.unl.gz';
    END IF;

    vFilenameZIP := getFilenameZIP(pJob, i);

    -- verschieben in fachlichen Ordner
    UTL_FILE.FRENAME('EX_TEMP', vFilenameZIP, vDirName, vFilenameZIP, TRUE);

--    vUnixCom := 'chmod ' || vMODE || ' ' || vPath || '/' || vFilenameZIP;
--    pkg_lib_ex_basic_newapi.CALLUNIX(vUnixCom, vEX_SESSION);
  END LOOP;

  IF pJob.JOB_OBJ_TYPE ='p' THEN
    logInfo('Das Objekt ' || pJob.JOB_OBJ_NAME || '_' || pJob.JOB_PART || ' wurde ' || pPar || '-fach parallel exportiert');
  ELSE
    logInfo('Das Objekt ' || pJob.JOB_OBJ_NAME || ' wurde ' || pPar || '-fach parallel exportiert');
  END IF;

EXCEPTION
  WHEN UTL_FILE.INVALID_PATH THEN
    setStatusJOB(vJOB_ID, 'F');
    if UTL_FILE.IS_OPEN(L_FILE) THEN UTL_FILE.FCLOSE(l_file); END IF;
    pkg_lib_ex_basic_newapi.WRITELOG(P_EVENTTYPE      => pkg_lib_ex_basic_newapi.EventT_FuncError,
                              P_SESSION        => vEX_SESSION,
                              P_JOB_ID         => vCurrentJobId,
                              P_IDENTIFICATION => $$PLSQL_UNIT || '_JOB_ID_' || vJOB_ID,
                              P_TXT            => 'Fehler, Pfad fuer GZ-Datei ist nicht valide');
  WHEN UTL_FILE.RENAME_FAILED THEN
    setStatusJOB(vJOB_ID, 'F');
    if UTL_FILE.IS_OPEN(L_FILE) THEN UTL_FILE.FCLOSE(l_file); END IF;
    pkg_lib_ex_basic_newapi.WRITELOG(P_EVENTTYPE      => pkg_lib_ex_basic_newapi.EventT_FuncError,
                              P_SESSION        => vEX_SESSION,
                              P_JOB_ID         => vCurrentJobId,
                              P_IDENTIFICATION => $$PLSQL_UNIT || '_JOB_ID_' || vJOB_ID,
                              P_TXT            => 'Fehler, GZ-Datei kann nicht umbenannt werden');
  WHEN UTL_FILE.INVALID_OPERATION THEN
    setStatusJOB(vJOB_ID, 'F');
    if UTL_FILE.IS_OPEN(L_FILE) THEN UTL_FILE.FCLOSE(l_file); END IF;
    pkg_lib_ex_basic_newapi.WRITELOG(P_EVENTTYPE      => pkg_lib_ex_basic_newapi.EventT_FuncError,
                              P_SESSION        => vEX_SESSION,
                              P_JOB_ID         => vCurrentJobId,
                              P_IDENTIFICATION => $$PLSQL_UNIT || '_JOB_ID_' || vJOB_ID,
                              P_TXT            => 'Fehler, GZ-Datei kann nicht gelesen werden');
  WHEN OTHERS THEN
    setStatusJOB(vJOB_ID, 'F');
    if UTL_FILE.IS_OPEN(L_FILE) THEN UTL_FILE.FCLOSE(l_file); END IF;
    pkg_lib_ex_basic_newapi.WRITELOG(P_EVENTTYPE      => pkg_lib_ex_basic_newapi.EventT_SysError,
                              P_SESSION        => vEX_SESSION,
                              P_JOB_ID         => vCurrentJobId,
                              P_IDENTIFICATION => $$PLSQL_UNIT || '_JOB_ID_' || vJOB_ID,
                              P_TXT            => 'Systemfehler');
END dumpJob2CVS;

/*------------------------------------------------------------------------------
Name:         dumpJob2CVSPipe

Parameter:    pJob Record vom Typ Job.

Beschreibung: Exportiert das materialisierte DB-Objekt als Datei in Filesystem.
              Die Datei wird on the Fly gezippt.
------------------------------------------------------------------------------*/
procedure dumpJob2CVSPipe (pJob STAT_EXP_DATAMART_MCFG.WRK_EX_JOB%ROWTYPE) IS
vFilenameZIP  VARCHAR2(200);
vDirName      VARCHAR2(200);
vTabname      VARCHAR2(128);
vTabnamePipe  VARCHAR2(200);
vFilename     VARCHAR2(200);
vUnixCom      VARCHAR2(32767);
vPathTemp     VARCHAR2(2000);
vPath         VARCHAR2(2000);
vSqlStmt      VARCHAR2(2000);
vParLocalStmt VARCHAR2(200);
vParLocalAlt  VARCHAR2(10);
anz           STAT_EXP_DATAMART_MCFG.WRK_EX_JOB.JOB_ANZ_DUMP%TYPE;
vJOB_ID       NUMBER(20,0);
pPar          INTEGER;
l_file        UTL_FILE.FILE_TYPE;
v_plsql_block VARCHAR2(32767);
v_line number := $$PLSQL_LINE;

BEGIN
  -- Auf einem Knoten bleiben und wert von Parameter merken
  v_line := $$PLSQL_LINE;
  SELECT value INTO vParLocalAlt from V$PARAMETER WHERE name = 'parallel_force_local';




  v_line := $$PLSQL_LINE;
  IF pJob.JOB_KZ = 'i' THEN
     alterSessionIAB;
  ELSE
     alterSessionDME;
  END IF;

  v_line := $$PLSQL_LINE;
  vJOB_ID := pJob.JOB_ID;
  vTabname  := 'EXP' || pJob.JOB_ID;
  vFilename := 'EXP' || pJob.JOB_ID|| '.sh';

  -- wenn Parallelitaet nicht vorgegeben dann automatisch ermitteln
  v_line := $$PLSQL_LINE;
  IF vParallel IS NULL THEN
    pPar := getParallel(pJob);
  ELSE
    pPar := vParallel;
  END IF;



  vPathTemp    := pkg_lib_ex_basic_newapi.getPath('EX_TEMP');
  vUnixCom     := NULL;


  -- SeidelD006 22.10.2021 -- Direkt in File schreiben, ohne Pufferung in vUnixCom
  l_file := FOPEN(
                    location => 'EX_TEMP',
                    filename => vFilename
                    /*,open_mode => 'w',
                    max_linesize => 32767*/
                 );
      logVerbose('Datei '||vFileName||' geöffnet');
      UTL_FILE.PUT_LINE(l_file, '#!/bin/ksh');
      UTL_FILE.PUT_LINE(l_file, 'PATH=/bin:/usr/bin:.');



  -- named Pipe vorbereiten und anlegen
  v_line := $$PLSQL_LINE;
  FOR i IN 1..pPar
  LOOP
  v_line := $$PLSQL_LINE;

    vTabnamePipe := 'EXP' || pJob.JOB_ID || '.'|| i || '.pipe';
    vFilenameZIP := getFilenameZIP(pJob, i);
    logVerbose('Pipe:'||vTabnamePipe);
    -- named Pipe anlegen.
    UTL_FILE.PUT_LINE(l_file, 'mkfifo -m ' || vMODE || ' ' || vPathTemp || '/' ||vTabnamePipe);
    UTL_FILE.PUT_LINE(l_file, 'rm '  || vPathTemp || '/' || vFilenameZIP||' 2>/dev/null');
    UTL_FILE.PUT_LINE(l_file, 'touch '  || vPathTemp || '/' || vFilenameZIP);
    UTL_FILE.PUT_LINE(l_file, 'chmod ' || vMODE || ' ' || vPathTemp || '/' || vFilenameZIP);
    UTL_FILE.PUT_LINE(l_file, 'gzip < ' || vPathTemp || '/' || vTabnamePipe ||' >> ' || vPathTemp || '/' || vFilenameZIP || ' &');
--    vUnixCom:=vUnixCom||chr(10)||'mkfifo -m ' || vMODE || ' ' || vPathTemp || '/' ||vTabnamePipe;
--    vUnixCom:=vUnixCom||chr(10)||'rm '  || vPathTemp || '/' || vFilenameZIP||' 2>/dev/null';
--    vUnixCom:=vUnixCom||chr(10)||'touch '  || vPathTemp || '/' || vFilenameZIP;
--    vUnixCom:=vUnixCom||chr(10)||'chmod ' || vMODE || ' ' || vPathTemp || '/' || vFilenameZIP;
--    vUnixCom:=vUnixCom||chr(10)||'gzip < ' || vPathTemp || '/' || vTabnamePipe ||' >> ' || vPathTemp || '/' || vFilenameZIP || ' &';
  END LOOP;
  v_line := $$PLSQL_LINE;




  --vUnixCom:=vUnixCom||chr(10)||'touch ' || vPathTemp || '/' || vFilename || '.done'   ;
    UTL_FILE.PUT_LINE(l_file,'touch ' || vPathTemp || '/' || vFilename || '.done');


  --vUnixCom:=vUnixCom||chr(10)||'wait';
    UTL_FILE.PUT_LINE(l_file, 'wait');


    UTL_FILE.PUT(l_file,vUnixCom);
    UTL_FILE.FCLOSE(l_file);


    --SH-Skript ausfuehren
    vUnixCom := 'chmod ' || vMODE_EXE || ' ' || vPathTemp || '/' || vFilename||chr(10)||
                  'sh ' || chr(34) || vPathTemp || '/' || vFilename || chr(34);
      v_line := $$PLSQL_LINE;
      pkg_lib_ex_basic_newapi.CALLUNIXASYNC(vUnixCom, vEX_SESSION);

  -- Warten, bis alle Pipes angelegt wurden und dann SH-Skript loeschen
  v_line := $$PLSQL_LINE;
  vUnixCom := 'while [[ ! -f ' || vPathTemp || '/' || vFilename || '.done  ]] ; do sleep 1 ; done '||chr(10)||
  'rm ' || vPathTemp || '/' || vFilename ||' '||vPathTemp || '/' || vFilename || '.done 2>/dev/null'||
  --||chr(10)||'date >> '|| vPathTemp || '/ps.log'
  --||chr(10)||'ps -ef | grep gzip >> '|| vPathTemp || '/ps.log'
  --||chr(10)||'uname -a >> '|| vPathTemp || '/ps.log'
  chr(10)||  'true'
  --||chr(10)||'date>>'|| vPathTemp || '/cleanup.log'
  ;

  v_line := $$PLSQL_LINE;

  logVerbose('Cleanup:'||vUnixCom);

  pkg_lib_ex_basic_newapi.CALLUNIX(vUnixCom, vEX_SESSION);

  logVerbose('Vor parallelem Unload.');
  v_line := $$PLSQL_LINE;
  IF pJob.JOB_OBJ_TYPE IN ('t', 'm') THEN
     v_line := $$PLSQL_LINE;
     anz := PKG_EX_DUMP_CSV.tab2csv_parallel(p_directory => 'EX_TEMP',
                                             p_filename  => vTabname,
                                             p_schema    => upper(pJob.JOB_SCHEMA),
                                             p_tabname   => upper(pJob.JOB_OBJ_NAME),
                                             p_recordend => chr(10),
                                             p_charset => case when vDoUTF8UNLOAD then vUTF8Charset else vDefaultCharset end,
                                             p_parallel  => pPar,
                                             p_delimiter => vDelimiter);
  ELSE

     v_line := $$PLSQL_LINE;
     anz := PKG_EX_DUMP_CSV.tab2csv_parallel(p_directory => 'EX_TEMP',
                                             p_filename  => vTabname,
                                             p_schema    => vDataSchema,
                                             p_tabname   => vTabname,
                                             p_recordend => chr(10),
                                             p_charset => case when vDoUTF8UNLOAD then vUTF8Charset else vDefaultCharset end,
                                             p_parallel  => pPar,
                                             p_delimiter => vDelimiter);
  END IF;

    v_line := $$PLSQL_LINE;
    UPDATE STAT_EXP_DATAMART_MCFG.WRK_EX_JOB SET JOB_ANZ_DUMP = anz WHERE JOB_ID = vJOB_ID;

  v_line := $$PLSQL_LINE;
  vDirName  := getTargetDir (pJob.JOB_KZ, pJob.JOB_SCHEMA);
  vPath := pkg_lib_ex_basic_newapi.getPath(vDirName);

  -- loeschen alte Gezippte Dateien, evtl. vorhandene Semaphoren und PIPES
--  v_line := $$PLSQL_LINE;
--  vUnixCom :=   'find '|| vPath || ' -type f -iname ' || pJob.JOB_OBJ_NAME || '.*.unl.*' || ' -delete'||chr(10)||
--                'rm ' || vPathTemp || '/' || vTabname || '.*.pipe';
--  v_line := $$PLSQL_LINE;
--  pkg_lib_ex_basic_newapi.CALLUNIX(vUnixCom, vEX_SESSION);

  -- entladene Dateien  verschieben  und PIPES loeschen
  v_line := $$PLSQL_LINE;
  FOR i IN 1..pPar
  LOOP
  v_line := $$PLSQL_LINE;
    vFilenameZIP := getFilenameZIP(pJob, i);

    -- verschieben in fachlichen Ordner
    v_line := $$PLSQL_LINE;
    UTL_FILE.FRENAME('EX_TEMP', vFilenameZIP, vDirName, vFilenameZIP, TRUE);
    --UTL_FILE.FREMOVE('EX_TEMP',vTabname || '.'||i||'.pipe');
--    vUnixCom := 'chmod ' || vMODE || ' ' || vPath || '/' || vFilenameZIP;
--    pkg_lib_ex_basic_newapi.CALLUNIX(vUnixCom, vEX_SESSION);
  END LOOP;

  -- Parameter auf alten Wert zur?etzen
  v_line := $$PLSQL_LINE;
  vParLocalStmt := 'alter session set "parallel_force_local" = ' || vParLocalAlt;
  v_line := $$PLSQL_LINE;
  EXECUTE IMMEDIATE vParLocalStmt;

  IF pJob.JOB_OBJ_TYPE ='p' THEN
    v_line := $$PLSQL_LINE;
    logInfo('Das Objekt ' || pJob.JOB_OBJ_NAME || '_' || pJob.JOB_PART || ' wurde ' || pPar || '-fach parallel exportiert');
  ELSE
    v_line := $$PLSQL_LINE;
    logInfo('Das Objekt ' || pJob.JOB_OBJ_NAME || ' wurde ' || pPar || '-fach parallel exportiert');
  END IF;

EXCEPTION
  WHEN UTL_FILE.INVALID_PATH THEN
    -- Parameter auf alten Wert zur?etzen
    vParLocalStmt := 'alter session set "parallel_force_local" = ' || vParLocalAlt;
    EXECUTE IMMEDIATE vParLocalStmt;
    setStatusJOB(vJOB_ID, 'F');
    if UTL_FILE.IS_OPEN(L_FILE) THEN UTL_FILE.FCLOSE(l_file); END IF;
    pkg_lib_ex_basic_newapi.WRITELOG(P_EVENTTYPE      => pkg_lib_ex_basic_newapi.EventT_FuncError,
                              P_SESSION        => vEX_SESSION,
                              P_JOB_ID         => vCurrentJobId,
                              P_IDENTIFICATION => $$PLSQL_UNIT || '_JOB_ID_' || vJOB_ID,
                              P_TXT            => 'Fehler, Pfad für GZ-Datei ist nicht valide');
  WHEN UTL_FILE.RENAME_FAILED THEN
    -- Parameter auf alten Wert zur?etzen
    vParLocalStmt := 'alter session set "parallel_force_local" = ' || vParLocalAlt;
    EXECUTE IMMEDIATE vParLocalStmt;
    setStatusJOB(vJOB_ID, 'F');
    if UTL_FILE.IS_OPEN(L_FILE) THEN UTL_FILE.FCLOSE(l_file); END IF;
    pkg_lib_ex_basic_newapi.WRITELOG(P_EVENTTYPE      => pkg_lib_ex_basic_newapi.EventT_FuncError,
                              P_SESSION        => vEX_SESSION,
                              P_JOB_ID         => vCurrentJobId,
                              P_IDENTIFICATION => $$PLSQL_UNIT || '_JOB_ID_' || vJOB_ID,
                              P_TXT            => 'Fehler, GZ-Datei kann nicht umbenannt werden');
  WHEN UTL_FILE.INVALID_OPERATION THEN
    -- Parameter auf alten Wert zur?etzen
    vParLocalStmt := 'alter session set "parallel_force_local" = ' || vParLocalAlt;
    EXECUTE IMMEDIATE vParLocalStmt;
    setStatusJOB(vJOB_ID, 'F');
    if UTL_FILE.IS_OPEN(L_FILE) THEN UTL_FILE.FCLOSE(l_file); END IF;
    pkg_lib_ex_basic_newapi.WRITELOG(P_EVENTTYPE      => pkg_lib_ex_basic_newapi.EventT_FuncError,
                              P_SESSION        => vEX_SESSION,
                              P_JOB_ID         => vCurrentJobId,
                              P_IDENTIFICATION => $$PLSQL_UNIT || '_JOB_ID_' || vJOB_ID,
                              P_TXT            => 'Fehler, GZ-Datei kann nicht gelesen werden');
  WHEN OTHERS THEN
    setStatusJOB(vJOB_ID, 'F');
    if UTL_FILE.IS_OPEN(L_FILE) THEN UTL_FILE.FCLOSE(l_file); END IF;
    -- Parameter auf alten Wert zur?etzen
    vParLocalStmt := 'alter session set "parallel_force_local" = ' || vParLocalAlt;
    EXECUTE IMMEDIATE vParLocalStmt;
    pkg_lib_ex_basic_newapi.WRITELOG(P_EVENTTYPE      => pkg_lib_ex_basic_newapi.EventT_SysError,
                              P_SESSION        => vEX_SESSION,
                              P_JOB_ID         => vCurrentJobId,
                              P_IDENTIFICATION => $$PLSQL_UNIT || '_JOB_ID_' || vJOB_ID||'('||v_line||')',
                              P_TXT            => 'Systemfehler');
END dumpJob2CVSPipe;


/*------------------------------------------------------------------------------
Name:         dropJobMat

Parameter:    pJobID ID eines Jobs

Beschreibung: Materialiiertes Objekt wird gelöscht.
------------------------------------------------------------------------------*/
procedure dropJobMat (pJobID STAT_EXP_DATAMART_MCFG.WRK_EX_JOB.JOB_ID%TYPE) IS
vJOB_ID STAT_EXP_DATAMART_MCFG.WRK_EX_JOB.JOB_ID%TYPE;
vSQL    VARCHAR2(4000);

BEGIN
  vJOB_ID :=pJobID;
  -- BIDW-495 - SeidelD006 - 20240916
  -- Umstellung auf GTTs
  
  -- GTT vor dem Drop leeren
  vSQL :='TRUNCATE TABLE ' || vDataSchema || '.' || 'EXP' || vJOB_ID;
  execute immediate vSQL;

  -- GTT droppen
  vSQL :='DROP TABLE ' || vDataSchema || '.' || 'EXP' || vJOB_ID;
  execute immediate vSQL;

EXCEPTION
  WHEN OTHERS THEN
    setStatusJOB(vJOB_ID, 'F');
    pkg_lib_ex_basic_newapi.WRITELOG(P_EVENTTYPE      => pkg_lib_ex_basic_newapi.EventT_SysError,
                              P_SESSION        => vEX_SESSION,
                              P_JOB_ID         => vCurrentJobId,
                              P_IDENTIFICATION => $$PLSQL_UNIT || '_JOB_ID_' || vJOB_ID,
                              P_TXT            => 'Systemfehler');
END dropJobMat;


/*------------------------------------------------------------------------------
Name:         doExpStop

Parameter:    pJobID ID eines Jobs

Beschreibung: Stoppt den DM-Export solange die Semaphore 'DM_EXPORT.STOP'
              im Pfad des Directory-Objekts 'EX_EXP' gesetzt ist.
              Es wird jede Minute geprüft, ob Semaphore noch gesetzt ist.
------------------------------------------------------------------------------*/
procedure doExpStop (pJobID STAT_EXP_DATAMART_MCFG.WRK_EX_JOB.JOB_ID%TYPE) IS

vJOB_ID           STAT_EXP_DATAMART_MCFG.WRK_EX_JOB.JOB_ID%TYPE;
vSemaporeExists   INTEGER;
vPath             VARCHAR2(2000);
vSleep            BOOLEAN := FALSE;

BEGIN
  vJOB_ID :=pJobID;
  vSemaporeExists := dbms_lob.fileexists(bfilename('EX_EXP', 'DM_EXPORT.STOP'));
  IF vSemaporeExists = 1 THEN
     vPath  := pkg_lib_ex_basic_newapi.getPath('EX_EXP');
     logInfo('Semaphore "DM_EXPORT.STOP" in ' || vPath || ' ist gesetzt, Verarbeitung wird angehalten');
     vSleep := TRUE;
  END IF;

  -- solange die Semaphore 'DM_EXPORT.STOP' gesetzt ist.
  WHILE vSemaporeExists = 1
  LOOP
    DBMS_SESSION.SLEEP (60);
    vSemaporeExists := dbms_lob.fileexists(bfilename('EX_EXP', 'DM_EXPORT.STOP'));
  END LOOP;

  IF vSleep THEN
     logInfo('Semaphore "DM_EXPORT.STOP" in ' || vPath || ' wurde aufgehoben, Verarbeitung geht weiter');
  END IF;
EXCEPTION
  WHEN OTHERS THEN
    setStatusJOB(vJOB_ID, 'F');
    pkg_lib_ex_basic_newapi.WRITELOG(P_EVENTTYPE      => pkg_lib_ex_basic_newapi.EventT_SysError,
                              P_SESSION        => vEX_SESSION,
                              P_JOB_ID         => vCurrentJobId,
                              P_IDENTIFICATION => $$PLSQL_UNIT || '_JOB_ID_' || vJOB_ID,
                              P_TXT            => 'Systemfehler');
END doExpStop;



/*------------------------------------------------------------------------------
Name:         exportJob

Parameter:    pJob Record vom Typ Job.

Beschreibung: Erledigt alle Arbeiten eines Export-Jobs Materialisieren des
              Objekts, erzeugen des DDL's, der Export selbst, aufräumen diverse
              Stati setzen.
              Das ganze abhängig von JOB_KZ  und JOB_OBJ_TYPE.
------------------------------------------------------------------------------*/
procedure exportJob(pJob STAT_EXP_DATAMART_MCFG.WRK_EX_JOB%ROWTYPE) IS
BEGIN
    -- View bei Dimensionen nicht materialisieren nur seine Unterobjekte
    IF pJob.JOB_KZ = 'd' AND pJob.JOB_OBJ_TYPE = 'v' AND NOT vVIEW_MAT THEN
      setStatusJOB(pJob.JOB_ID, 'M');
      dumpJobDDL(pJob);
      --logInfo('DDL fuer Job:' || pJob.JOB_ID || ' wurde generiert');
      setStatusJOB(pJob.JOB_ID, 'A');
    -- View bei Biodata materialisieren kein DDL und keine Unterobjekte
    ELSIF pJob.JOB_KZ = 'b' AND pJob.JOB_OBJ_TYPE = 'v' THEN
      matJobObject(pJob);
      --logInfo('Objekt fuer Job:' || pJob.JOB_ID || ' wurde materialisiert');
      setStatusJOB(pJob.JOB_ID, 'M');
      IF vPIPE THEN
        dumpJob2CVSPipe(pJob);
      ELSE
        dumpJob2CVS(pJob);
      END IF;
      --logInfo('Objekt für Job:' || pJob.JOB_ID || ' wurde ins File-System geschrieben');
      dropJobMat(pJob.JOB_ID);
      --logInfo('Materialisiertes Objekt fuer Job:' || pJob.JOB_ID || ' wurde geloescht');
      setStatusJOB(pJob.JOB_ID, 'A');
    ELSE
      matJobObject(pJob);
      --logInfo('Objekt fuer Job:' || pJob.JOB_ID || ' wurde materialisiert');
      setStatusJOB(pJob.JOB_ID, 'M');
      dumpJobDDL(pJob);
      --logInfo('DDL fuer Job:' || pJob.JOB_ID || ' wurde generiert');
      IF vPIPE THEN
        dumpJob2CVSPipe(pJob);
      ELSE
        dumpJob2CVS(pJob);
      END IF;
      --logInfo('Objekt fuer Job:' || pJob.JOB_ID || ' wurde ins File-System geschrieben');
      dropJobMat(pJob.JOB_ID);
      --logInfo('Materialisiertes Objekt fuer Job:' || pJob.JOB_ID || ' wurde geloescht');
      setStatusJOB(pJob.JOB_ID, 'A');
    END IF;
EXCEPTION
  WHEN OTHERS THEN
    pkg_lib_ex_basic_newapi.WRITELOG(P_EVENTTYPE      => pkg_lib_ex_basic_newapi.EventT_SysError,
                              P_SESSION        => vEX_SESSION,
                              P_JOB_ID         => vCurrentJobId,
                              P_IDENTIFICATION => $$PLSQL_UNIT,
                              P_TXT            => 'Systemfehler');
END exportJob;


/*------------------------------------------------------------------------------
Name:         exportJobs

Parameter:    keine

Beschreibung: Ermittelt alle zu exportierenden Jobs und exportiert sie.
------------------------------------------------------------------------------*/
procedure exportJobs IS
cJob STAT_EXP_DATAMART_MCFG.WRK_EX_JOB%ROWTYPE;
cursor curs(pSession in number) is (
                 SELECT *
     FROM STAT_EXP_DATAMART_MCFG.WRK_EX_JOB
    WHERE JOB_STATUS ='C'
      AND JOB_SESSION = pSESSION
--    ORDER BY 1
               ) ORDER BY JOB_STL_ID,JOB_ID;
BEGIN
  logInfo('Start Export aller erzeugten Jobs');
  -- Kein FOR-Cursor mehr, da bei Langläufern
  -- häufig "ORA-1555 Snapshot too old" Fehler

--  FOR cJob IN  -- alle zu exportierende Objekte
--  (SELECT *
--     FROM STAT_EXP_DATAMART_MCFG.WRK_EX_JOB
--    WHERE JOB_STATUS ='C'
--      AND JOB_SESSION = vEX_SESSION
--  )
  LOOP
    OPEN curs(vEX_SESSION);
    FETCH curs into cJob;
    EXIT WHEN curs%NOTFOUND;
    CLOSE curs;

    -- Job exportieren, wenn Fehler dann weitermachen
    BEGIN
      vCurrentJobId := cJob.JOB_ID;
      doExpStop(cJob.JOB_ID);  --stoppen wenn Semaphore gesetzt
      exportJob(cJob);
      -- Prüfen, ob aktueller Job auf die Bretter gegangen ist
      declare
       vCnt number;
      begin
       select count(*)
       into vCnt
       from STAT_EXP_DATAMART_MCFG.ERR_EX_PROT
       where ex_session=vEX_SESSION
       and event_type in ('S','E')
       and job_id=cJob.Job_ID;

       if vCnt>0
       then
         -- Job auf Fehler setzen
         setStatusJOB(cJob.JOB_ID, 'F');
       end if;
      end;
      createSemaphoresJob(vCurrentJobId);

      vCurrentJobId := NULL;
      --logInfo('Export-Auftrag fuer ' || cJob.JOB_OBJ_NAME || ' abgeschlossen');
    EXCEPTION
    WHEN OTHERS THEN
      vCurrentJobId := NULL;
    END;
  END LOOP;
  IF curs%ISOPEN
  THEN
    close curs;
  END IF;
  logInfo('Ende Export aller erzeugten Jobs');
EXCEPTION
  WHEN OTHERS THEN
    IF curs%ISOPEN
    THEN
      close curs;
    END IF;
    PKG_LIB_EX_BASIC_NEWAPI.WRITELOG(P_EVENTTYPE      => PKG_LIB_EX_BASIC_NEWAPI.EventT_SysError,
                              P_SESSION        => vEX_SESSION,
                              P_IDENTIFICATION => $$PLSQL_UNIT,
                              P_TXT            => 'Systemfehler');
END exportJobs;


/*------------------------------------------------------------------------------
Name:         checkSTL

Parameter:    keine

Beschreibung: Prüfungen für eine Steuerliste
------------------------------------------------------------------------------*/
function checkSTL RETURN BOOLEAN IS
vERR     NUMBER;
vAnz     NUMBER;
vAnzEAP  NUMBER;
BEGIN
     -- Pruefen ob EXEMPT ACCESS POLICY da.
     SELECT count(*)
       INTO vAnzEAP
       FROM DBA_SYS_PRIVS
      WHERE GRANTEE = 'STAT_EXP_DATAMART_MCFG'
        AND PRIVILEGE ='EXEMPT ACCESS POLICY';

    -- Wenn nicht muss abgebrochen werden, da sonst nur Blödsinn exportiert wird
    IF vAnzEAP = 0  THEN
      pkg_lib_ex_basic_newapi.WRITELOG(P_EVENTTYPE      => pkg_lib_ex_basic_newapi.EventT_FuncError,
                                P_SESSION        => vEX_SESSION,
                                P_IDENTIFICATION => 'PKG_EX_LOAD_STL',
                                P_TXT            => 'Fehler, dem User "STAT_EXP_DATAMART_MCFG" fehlt das System-Privileg "EXEMPT ACCESS POLICY"');
    END IF;

    -- Prüfen ob Systemfehler
    SELECT count(*)
      INTO vERR
      FROM STAT_EXP_DATAMART_MCFG.ERR_EX_PROT
     WHERE IDENTIFICATION = 'PKG_EX_LOAD_STL'
       AND EVENT_TYPE IN(pkg_lib_ex_basic_newapi.EventT_FuncError, pkg_lib_ex_basic_newapi.EventT_SysError)
       AND EX_SESSION = vEX_SESSION;

    IF vERR > 0 THEN
      logInfo('Beim Laden der Steuerliste sind Fehler aufgetreten');
    END IF;

    -- Prüfen ob nichts geladen aus Steuerliste
    SELECT count(*)
      INTO vANZ
      FROM STAT_EXP_DATAMART_MCFG.WRK_EX_OBJ
     WHERE STL_STATUS = 'L'
       AND STL_SESSION = vEX_SESSION;



    IF vANZ = 0 THEN
      logInfo('Es wurden keine Daten aus der Steuerliste geladen');
    END IF;

    -- fuer alle geladenen Steuerlisten-Objekte
    FOR cOBJ IN
    (SELECT *
       FROM STAT_EXP_DATAMART_MCFG.WRK_EX_OBJ
      WHERE STL_STATUS = 'L'
        AND STL_SESSION = vEX_SESSION
    )
    LOOP
      setStatusOBJ (cOBJ.STL_ID, 'G');  --Status geprueft setzen
    END LOOP;

    IF vANZ = 0 OR vERR > 0 THEN
      RETURN TRUE;
    ELSE
      RETURN FALSE;
    END IF;
EXCEPTION
  WHEN OTHERS THEN
    pkg_lib_ex_basic_newapi.WRITELOG(P_EVENTTYPE      => pkg_lib_ex_basic_newapi.EventT_SysError,
                              P_SESSION        => vEX_SESSION,
                              P_IDENTIFICATION => $$PLSQL_UNIT,
                              P_TXT            => 'Systemfehler');
    RETURN TRUE;
END checkSTL;

/*------------------------------------------------------------------------------
Name:         createSemaphores

Parameter:    keine

Beschreibung: Erzeugt ".ok" oder ".error" Semaphore fuer alle Jobs.
------------------------------------------------------------------------------*/
PROCEDURE createSemaphores
IS
l_file          UTL_FILE.FILE_TYPE;
vFilenameTmp    VARCHAR2(200);
vFilenameFix    VARCHAR2(200);
vFilenameNeu    VARCHAR2(200);
vDirName        VARCHAR2(200);
vError          STAT_EXP_DATAMART_MCFG.ERR_EX_PROT.TEXT%TYPE;
vPath           VARCHAR2(2000);
vUnixCom        VARCHAR2(4000);
vExists         INTEGER;
vTouchScript    VARCHAR2(200);
BEGIN

  logInfo ('Start erzeugen Semaphore für exportierte Objekte');


     vTouchScript   := 'EXP' ||vEX_SESSION  || 'stati';

    l_file := FOPEN(
            location => 'EX_TEMP',
            filename => vTouchscript
            /*,open_mode => 'w',
            max_linesize => 32767*/);



  -- fuer alle generierten Jobs
  FOR cJOB IN
  (SELECT *
     FROM STAT_EXP_DATAMART_MCFG.WRK_EX_JOB
    WHERE JOB_SESSION = vEX_SESSION
      AND JOB_STATUS NOT IN ( 'I', 'E')
  )
  LOOP


      IF cJob.JOB_PART = 'NO' THEN
        vFilenameFix := cJob.JOB_SCHEMA || '.' || cJob.JOB_OBJ_NAME;
      ELSE
        vFilenameFix := cJob.JOB_SCHEMA || '.' || cJob.JOB_OBJ_NAME || '_' || cJob.JOB_PART;
      END IF;

    vDirName  := getTargetDir (cJob.JOB_KZ, cJob.JOB_SCHEMA);
    vPath    := pkg_lib_ex_basic_newapi.getPath(vDirName);

    -- loeschen altes Semaphor
    UTL_FILE.PUT_LINE(l_file,'find '|| vPath || ' -type f -iname ' || vFilenameFix || '.*.unl.status.*' || ' -mmin +1 -delete');
    --pkg_lib_ex_basic_newapi.CALLUNIXASYNC(vUnixCom, vEX_SESSION);

    --dbms_output.put_line(cJob.Job_id||':'||vFilenameFix||':'||cJob.JOB_STATUS);
    IF cJob.JOB_STATUS = 'A' THEN
       vFilenameNeu := vFilenameFix || '.' || cJob.JOB_ANZ_DUMP || '.1.unl.status.ok';
    ELSIF cJob.JOB_STATUS = 'I' THEN
      NULL;  --Ignorieren
    ELSIF cJob.JOB_STATUS = 'E' THEN
      NULL;  --Ignorieren
    ELSE
      vFilenameNeu := vFilenameFix || '.1.unl.status.error';
      vError := getJobError(cJob.JOB_ID);
      UTL_FILE.PUT_LINE(l_file, 'echo "'||replace(vError,'"','\"')||'">'||vPath||'/'||vFilenameNeu);

    END IF;
    UTL_FILE.PUT_LINE(l_file,'touch ' || vPath || '/' || vFilenameNeu);
    UTL_FILE.PUT_LINE(l_file,'chmod ' || vMODE || ' ' || vPath || '/' || vFilenameNeu);

  END LOOP;
  UTL_FILE.FCLOSE(l_file);
  vUnixCom:='sh '||pkg_lib_ex_basic_newapi.getPath('EX_TEMP')||'/'||vTouchScript||chr(10)||
  'rm '||pkg_lib_ex_basic_newapi.getPath('EX_TEMP')||'/'||vTouchScript;
  pkg_lib_ex_basic_newapi.CALLUNIX(vUnixCom, vEX_SESSION);

--  vUnixCom:='rm '||pkg_lib_ex_basic_newapi.getPath('EX_TEMP')||'/'||vTouchScript;
--  pkg_lib_ex_basic_newapi.CALLUNIXASYNC(vUnixCom, vEX_SESSION);



  logInfo ('Ende erzeugen Semaphore für exportierte Objekte');

  EXCEPTION
  WHEN UTL_FILE.INVALID_PATH THEN
    pkg_lib_ex_basic_newapi.WRITELOG(P_EVENTTYPE      => pkg_lib_ex_basic_newapi.EventT_FuncError,
                              P_SESSION        => vEX_SESSION,
                              P_IDENTIFICATION => $$PLSQL_UNIT,
                              P_TXT => 'Fehler, Pfad fuer Semaphore-Datei ist nicht valide');
  WHEN UTL_FILE.RENAME_FAILED THEN
    pkg_lib_ex_basic_newapi.WRITELOG(P_EVENTTYPE      => pkg_lib_ex_basic_newapi.EventT_FuncError,
                              P_SESSION        => vEX_SESSION,
                              P_IDENTIFICATION => $$PLSQL_UNIT,
                              P_TXT            => 'Fehler, Semaphore-Datei kann nicht umbenannt werden');
  WHEN UTL_FILE.INVALID_OPERATION THEN
    pkg_lib_ex_basic_newapi.WRITELOG(P_EVENTTYPE      => pkg_lib_ex_basic_newapi.EventT_FuncError,
                              P_SESSION        => vEX_SESSION,
                              P_IDENTIFICATION => $$PLSQL_UNIT,
                              P_TXT            => 'Fehler, Semaphore-Datei kann nicht geschrieben werden');
  WHEN OTHERS THEN
    UTL_FILE.FCLOSE(l_file);
    pkg_lib_ex_basic_newapi.WRITELOG(P_EVENTTYPE      => pkg_lib_ex_basic_newapi.EventT_SysError,
                              P_SESSION        => vEX_SESSION,
                              P_IDENTIFICATION => $$PLSQL_UNIT,
                              P_TXT            => 'Systemfehler');
END createSemaphores;

/*------------------------------------------------------------------------------
Name:         createSemaphoresJob

Parameter:    keine

Beschreibung: Erzeugt ".ok" oder ".error" Semaphore fuer einen Job.
------------------------------------------------------------------------------*/
PROCEDURE createSemaphoresJob(pJobId in Number)
IS
l_file          UTL_FILE.FILE_TYPE;
vFilenameTmp    VARCHAR2(200);
vFilenameFix    VARCHAR2(200);
vFilenameNeu    VARCHAR2(200);
vDirName        VARCHAR2(200);
vError          STAT_EXP_DATAMART_MCFG.ERR_EX_PROT.TEXT%TYPE;
vPath           VARCHAR2(2000);
vUnixCom        VARCHAR2(4000);
vExists         INTEGER;
vTouchScript    VARCHAR2(200);
BEGIN

  logInfo ('Start erzeugen Semaphore für exportierte Job-Objekte');


     vTouchScript   := 'EXP' ||vEX_SESSION  || 'stati';

    l_file := FOPEN(
            location => 'EX_TEMP',
            filename => vTouchscript
            /*,open_mode => 'w',
            max_linesize => 32767*/);



  -- fuer alle generierten Jobs
  FOR cJOB IN
  (SELECT *
     FROM STAT_EXP_DATAMART_MCFG.WRK_EX_JOB
    WHERE JOB_SESSION = vEX_SESSION
      AND JOB_STATUS NOT IN ( 'I', 'E')
      AND JOB_ID=pJobId
  )
  LOOP


      IF cJob.JOB_PART = 'NO' THEN
        vFilenameFix := cJob.JOB_SCHEMA || '.' || cJob.JOB_OBJ_NAME;
      ELSE
        vFilenameFix := cJob.JOB_SCHEMA || '.' || cJob.JOB_OBJ_NAME || '_' || cJob.JOB_PART;
      END IF;

    vDirName  := getTargetDir (cJob.JOB_KZ, cJob.JOB_SCHEMA);
    vPath    := pkg_lib_ex_basic_newapi.getPath(vDirName);

    -- loeschen altes Semaphor
    UTL_FILE.PUT_LINE(l_file,'find '|| vPath || ' -type f -iname ' || vFilenameFix || '.*.unl.status.*' || ' -mmin +1 -delete');
    --pkg_lib_ex_basic_newapi.CALLUNIXASYNC(vUnixCom, vEX_SESSION);
    -- loeschen evtl. vorhandene Pipes
    UTL_FILE.PUT_LINE(l_file,'find '|| pkg_lib_ex_basic_newapi.getPath('EX_TEMP') || ' -type p -iname ' || 'EXP'||cJob.Job_Id||'.*.pipe' || ' -delete');

    --dbms_output.put_line(cJob.Job_id||':'||vFilenameFix||':'||cJob.JOB_STATUS);
    IF cJob.JOB_STATUS = 'A' THEN
       vFilenameNeu := vFilenameFix || '.' || cJob.JOB_ANZ_DUMP || '.1.unl.status.ok';
    ELSIF cJob.JOB_STATUS = 'I' THEN
      NULL;  --Ignorieren
    ELSIF cJob.JOB_STATUS = 'E' THEN
      NULL;  --Ignorieren
    ELSE
      vFilenameNeu := vFilenameFix || '.1.unl.status.error';
      vError := getJobError(cJob.JOB_ID);
      UTL_FILE.PUT_LINE(l_file, 'echo "'||replace(vError,'"','\"')||'">'||vPath||'/'||vFilenameNeu);

    END IF;
    UTL_FILE.PUT_LINE(l_file,'touch ' || vPath || '/' || vFilenameNeu);
    UTL_FILE.PUT_LINE(l_file,'chmod ' || vMODE || ' ' || vPath || '/' || vFilenameNeu);

  END LOOP;
  UTL_FILE.FCLOSE(l_file);
  vUnixCom:='sh '||pkg_lib_ex_basic_newapi.getPath('EX_TEMP')||'/'||vTouchScript;
  vUnixCom:=vUnixCom||chr(10)||'rm          '||pkg_lib_ex_basic_newapi.getPath('EX_TEMP')||'/'||vTouchScript;
  pkg_lib_ex_basic_newapi.CALLUNIXASYNC(vUnixCom, vEX_SESSION);

--  vUnixCom:=' rm '||pkg_lib_ex_basic_newapi.getPath('EX_TEMP')||'/'||vTouchScript;
--  pkg_lib_ex_basic_newapi.CALLUNIXASYNC(vUnixCom, vEX_SESSION);



  logInfo ('Ende erzeugen Semaphore für exportierte Job-Objekte');

  EXCEPTION
  WHEN UTL_FILE.INVALID_PATH THEN
    pkg_lib_ex_basic_newapi.WRITELOG(P_EVENTTYPE      => pkg_lib_ex_basic_newapi.EventT_FuncError,
                              P_SESSION        => vEX_SESSION,
                              P_IDENTIFICATION => $$PLSQL_UNIT,
                              P_TXT => 'Fehler, Pfad fuer Semaphore-Datei ist nicht valide');
  WHEN UTL_FILE.RENAME_FAILED THEN
    pkg_lib_ex_basic_newapi.WRITELOG(P_EVENTTYPE      => pkg_lib_ex_basic_newapi.EventT_FuncError,
                              P_SESSION        => vEX_SESSION,
                              P_IDENTIFICATION => $$PLSQL_UNIT,
                              P_TXT            => 'Fehler, Semaphore-Datei kann nicht umbenannt werden');
  WHEN UTL_FILE.INVALID_OPERATION THEN
    pkg_lib_ex_basic_newapi.WRITELOG(P_EVENTTYPE      => pkg_lib_ex_basic_newapi.EventT_FuncError,
                              P_SESSION        => vEX_SESSION,
                              P_IDENTIFICATION => $$PLSQL_UNIT,
                              P_TXT            => 'Fehler, Semaphore-Datei kann nicht geschrieben werden');
  WHEN OTHERS THEN
    UTL_FILE.FCLOSE(l_file);
    pkg_lib_ex_basic_newapi.WRITELOG(P_EVENTTYPE      => pkg_lib_ex_basic_newapi.EventT_SysError,
                              P_SESSION        => vEX_SESSION,
                              P_IDENTIFICATION => $$PLSQL_UNIT,
                              P_TXT            => 'Systemfehler');
END createSemaphoresJob;


/*------------------------------------------------------------------------------
Name:         createStatusList

Parameter:    p_RELOAD   Reload
              p_STL_NAME Name der Steuerliste

Beschreibung: Erzeugt eine Statusliste fuer die Steuerliste.
              Aufbau: - Schemaname
                      - Objektname
                      - Anzahl Datensätze
                      - Dateiname
------------------------------------------------------------------------------*/
PROCEDURE createStatusList (p_RELOAD BOOLEAN, p_STL_NAME VARCHAR2)
IS
l_file              UTL_FILE.FILE_TYPE;
vFilename1 constant VARCHAR2(200) := 'status.liste.' || vEX_SESSION;
vFilename2          VARCHAR2(200);
vFilenameUnl        VARCHAR2(200);
vSchema             VARCHAR2(128);
vSTL_KZ             CHAR(1);
vKenz               VARCHAR2(4);
vLine               VARCHAR2(2000);
vDirName            VARCHAR2(200);
vPath               VARCHAR2(2000);
vUnixCom            VARCHAR2(4000);
pPar                INTEGER;
vAlt_EX_SESSION     STAT_EXP_DATAMART_MCFG.WRK_EX_OBJ.STL_SESSION%TYPE;

BEGIN
  logInfo ('Start erzeugen Statusliste');
  l_file := FOPEN(
            location => 'EX_TEMP',
            filename => vFilename1
            /*,open_mode => 'w',
            max_linesize => 32767*/);




    SELECT DISTINCT STL_KZ INTO vSTL_KZ
       FROM STAT_EXP_DATAMART_MCFG.WRK_EX_OBJ
      WHERE STL_SESSION = vEX_SESSION;

     IF vSTL_KZ = 'i' THEN
        vDirName  := 'EX_IAB';
     ELSE
       SELECT DISTINCT upper(STL_SCHEMA) INTO vSchema
         FROM STAT_EXP_DATAMART_MCFG.WRK_EX_OBJ
        WHERE STL_SESSION = vEX_SESSION;
       vDirName  := getTargetDir (vSTL_KZ, vSchema);
     END IF;


     IF vSTL_KZ = 'i' THEN
        vKenz := 'iab';
     ELSIF vSTL_KZ = 'd' THEN
        vKenz := 'dim';
     ELSE
        vKenz := 'fakt';
     END IF;

     vFilename2 := 'status.liste.' || vKenz || '.' || lower(vSchema) || '.' || TO_CHAR(SYSDATE, 'YYYYMMDD.HH24MISS');


    -- fuer alle aktuell generierten Jobs
    FOR cJOB IN
    (SELECT *
       FROM STAT_EXP_DATAMART_MCFG.WRK_EX_JOB
      WHERE JOB_SESSION = vEX_SESSION
        AND JOB_STATUS = 'A'
    )
    LOOP


      -- wenn Parallelitaet nicht vorgegeben dann automatisch ermitteln
      IF vParallel IS NULL THEN
        pPar := getParallel(cJob);
      ELSE
        pPar := vParallel;
      END IF;

      IF cJob.JOB_PART = 'NO' THEN
         vFilenameUnl := cJob.JOB_OBJ_NAME || '.' || pPar ||'.unl.gz';
         -- aktuell verarbeitete Zeile eintragen
         vLine := lower(cJOB.JOB_SCHEMA) || ':' || cJOB.JOB_OBJ_NAME || ':' || cJOB.JOB_ANZ_DUMP || ':' || vFilenameUnl;
      ELSE
         vFilenameUnl := cJob.JOB_OBJ_NAME || '_' || cJob.JOB_PART || '.' || pPar ||'.unl.gz';
         -- aktuell verarbeitete Zeile eintragen
         vLine := lower(cJOB.JOB_SCHEMA) || ':' || cJOB.JOB_OBJ_NAME || '_' || cJob.JOB_PART || ':' || cJOB.JOB_ANZ_DUMP || ':' || vFilenameUnl;
      END IF;

      IF cJob.JOB_KZ = 'b' THEN
          vLine := replaceName_b(vLine);
      ELSIF cJob.JOB_KZ = 'v' THEN
          vLine := replaceName_v(vLine);
      END IF;


      UTL_FILE.PUT_LINE(l_file, vLine);

    END LOOP;

    -- wenn Reload
    IF p_RELOAD THEN

      --Session vom letzten lauf der Steuerliste ermitteln
      SELECT STL_SESSION
      INTO vAlt_EX_SESSION
      FROM
        (SELECT STL_SESSION,
                STL_ZBDAT
           FROM STAT_EXP_DATAMART_MCFG.WRK_EX_OBJ
          WHERE STL_FILE_NAME = p_STL_NAME
            AND STL_SESSION  <> vEX_SESSION
          ORDER BY STL_ZBDAT DESC
        )
      WHERE rownum = 1;

      -- fuer alle vor RELOAD generierten Jobs
      -- nur Jobs von Abgeschlossenen Listeneintraegen
      FOR cJOB IN
      (SELECT job.*
         FROM STAT_EXP_DATAMART_MCFG.WRK_EX_JOB job, STAT_EXP_DATAMART_MCFG.WRK_EX_OBJ obj
        WHERE JOB_SESSION = vAlt_EX_SESSION
          AND JOB_STL_ID = STL_ID
          AND STL_STATUS = 'A'
      )
      LOOP

        -- wenn Parallelitaet nicht vorgegeben dann automatisch ermitteln
        IF vParallel IS NULL THEN
          pPar := getParallel(cJob);
        ELSE
          pPar := vParallel;
        END IF;

        IF cJob.JOB_PART = 'NO' THEN
           vFilenameUnl := cJob.JOB_OBJ_NAME || '.' || pPar ||'.unl.gz';
           --  vor RELOAD verarbeitete Zeile eintragen
           vLine := lower(cJOB.JOB_SCHEMA) || ':' || cJOB.JOB_OBJ_NAME || ':' || cJOB.JOB_ANZ_DUMP || ':' || vFilenameUnl;
        ELSE
           vFilenameUnl := cJob.JOB_OBJ_NAME || '_' || cJob.JOB_PART || '.' || pPar ||'.unl.gz';
           --  vor RELOAD verarbeitete Zeile eintragen
           vLine := lower(cJOB.JOB_SCHEMA) || ':' || cJOB.JOB_OBJ_NAME || '_' || cJob.JOB_PART || ':' || cJOB.JOB_ANZ_DUMP || ':' || vFilenameUnl;
        END IF;

        IF cJob.JOB_KZ = 'b' THEN
            vLine := replaceName_b(vLine);
        ELSIF cJob.JOB_KZ = 'v' THEN
            vLine := replaceName_v(vLine);
        END IF;

        UTL_FILE.PUT_LINE(l_file, vLine);

      END LOOP;

    END IF;




  UTL_FILE.FCLOSE(l_file);

  UTL_FILE.FRENAME('EX_TEMP', vFilename1, vDirName, vFilename2, TRUE);

  vPath := pkg_lib_ex_basic_newapi.getPath(vDirName);
--  vUnixCom := 'chmod ' || vMODE || ' ' || vPath || '/' || vFilename2;
--  pkg_lib_ex_basic_newapi.CALLUNIX(vUnixCom, vEX_SESSION);

  logInfo ('Ende erzeugen Statusliste');
EXCEPTION
  WHEN UTL_FILE.INVALID_PATH THEN
    pkg_lib_ex_basic_newapi.WRITELOG(P_EVENTTYPE      => pkg_lib_ex_basic_newapi.EventT_FuncError,
                              P_SESSION        => vEX_SESSION,
                              P_IDENTIFICATION => $$PLSQL_UNIT,
                              P_TXT            => 'Fehler, Pfad für StatusList-Datei ist nicht valide');
  WHEN UTL_FILE.RENAME_FAILED THEN
    pkg_lib_ex_basic_newapi.WRITELOG(P_EVENTTYPE      => pkg_lib_ex_basic_newapi.EventT_FuncError,
                              P_SESSION        => vEX_SESSION,
                              P_IDENTIFICATION => $$PLSQL_UNIT,
                              P_TXT            => 'Fehler, StatusList-Datei kann nicht umbenannt werden');
  WHEN UTL_FILE.INVALID_OPERATION THEN
    pkg_lib_ex_basic_newapi.WRITELOG(P_EVENTTYPE      => pkg_lib_ex_basic_newapi.EventT_FuncError,
                              P_SESSION        => vEX_SESSION,
                              P_IDENTIFICATION => $$PLSQL_UNIT,
                              P_TXT            => 'Fehler, StatusList-Datei kann nicht geschrieben werden');
WHEN OTHERS THEN
  UTL_FILE.FCLOSE(l_file);
    pkg_lib_ex_basic_newapi.WRITELOG(P_EVENTTYPE      => pkg_lib_ex_basic_newapi.EventT_SysError,
                              P_SESSION        => vEX_SESSION,
                              P_IDENTIFICATION => $$PLSQL_UNIT,
                              P_TXT            => 'Systemfehler');
END createStatusList;


/*------------------------------------------------------------------------------
Name:         createStatusProtokoll

Parameter:    p_STL_NAME   Steuerlistenname

Beschreibung: Erzeugt ein Protokoll mit allen Logeintaegen.
------------------------------------------------------------------------------*/
PROCEDURE createStatusProtokoll (p_STL_NAME VARCHAR2)
IS
l_file              UTL_FILE.FILE_TYPE;
vFilename1 constant VARCHAR2(200) := 'status.protokoll.' || vEX_SESSION;
vFilename2 constant VARCHAR2(200) := 'status.export.protokoll.txt';
vSchema             VARCHAR2(128);
vSTL_KZ             CHAR(1);
vLine               VARCHAR2(4000);
vDirName            VARCHAR2(200);
vTYPE               VARCHAR2(9);
i                   INTEGER;
vExistsIAB          INTEGER;
vPath               VARCHAR2(4000);
vUnixCom            VARCHAR2(4000);
BEGIN
  logInfo ('Start erzeugen Statusprotokoll');
  l_file := FOPEN(
            location => 'EX_TEMP',
            filename => vFilename1
            /*,open_mode => 'w',
            max_linesize => 32767*/);

       -- Infos aus Dateinamen befuellen (Zerlegung des Dateinamens)
      i := 1;
      FOR cParam IN
      (SELECT * FROM TABLE(STAT_EXP_DATAMART_MCFG.pkg_lib_ex_basic_newapi.splitSTL( p_STL_NAME))
      )
      LOOP
        IF i                 = 1 THEN
           IF cParam.COLUMN_VALUE IN ( 'b', 'f', 'd', 'v') THEN
              vSTL_KZ     := cParam.COLUMN_VALUE;
           ELSE
              vSTL_KZ     := '0';
           END IF;
        ELSIF i              = 2 THEN
          NULL ; --wird hier nicht gebraucht.
        ELSIF i              = 3 THEN
          vSchema := upper(cParam.COLUMN_VALUE);
        ELSE
          NULL;
        END IF;
        i := i + 1;
      END LOOP; -- Infos aus Dateinamen befuellen (Zerlegung des Dateinamens)

    vExistsIAB  := dbms_lob.fileexists(bfilename('EX_STL_IAB',  p_STL_NAME));
    vDirName  := getTargetDir (vSTL_KZ, vSchema);

    IF vExistsIAB <> 0 THEN
      vDirName  := 'EX_IAB';
    END IF;

    -- fuer alle Protokolleintraege zu dem Lauf
    FOR cProt IN
    (SELECT *
       FROM STAT_EXP_DATAMART_MCFG.ERR_EX_PROT
      WHERE EX_SESSION = to_char(vEX_SESSION)
      ORDER BY ID
    )
    LOOP
      IF cProt.EVENT_TYPE IN ('S','E') THEN
         vTYPE := ' ERROR   ';
      ELSIF cProt.EVENT_TYPE IN ('W') THEN
         vTYPE := ' WARNING ';
      ELSIF cProt.EVENT_TYPE IN ('I') THEN
         vTYPE := '         ';
      ELSE
         vTYPE := '         ';
      END IF;

      vLine := substr(to_char (cProt.DATUM, 'RRRR-MM-DD HH24:MI:SS' ),1, 4000) || vTYPE || cProt.TEXT;
      UTL_FILE.PUT_LINE(l_file, vLine);
    END LOOP;

    vLine := to_char (sysdate, 'RRRR-MM-DD HH24:MI:SS' ) || '         ' || 'E N D E - D M _ E X P O R T';
    UTL_FILE.PUT_LINE(l_file, vLine);

  UTL_FILE.FCLOSE(l_file);
  logInfo(vDirName||':'||vFilename2);
  UTL_FILE.FRENAME('EX_TEMP', vFilename1, vDirName, vFilename2, TRUE);

  vPath := pkg_lib_ex_basic_newapi.getPath(vDirName);
--  vUnixCom := 'chmod ' || vMODE || ' ' || vPath || '/' || vFilename2;
--  pkg_lib_ex_basic_newapi.CALLUNIX(vUnixCom, vEX_SESSION);

  logInfo ('Ende erzeugen Statusprotokoll');
EXCEPTION
  WHEN UTL_FILE.INVALID_PATH THEN
    pkg_lib_ex_basic_newapi.WRITELOG(P_EVENTTYPE      => pkg_lib_ex_basic_newapi.EventT_FuncError,
                              P_SESSION        => vEX_SESSION,
                              P_IDENTIFICATION => $$PLSQL_UNIT,
                              P_TXT            => 'Fehler, Pfad für Protokoll-Datei ist nicht valide');
  WHEN UTL_FILE.RENAME_FAILED THEN
    pkg_lib_ex_basic_newapi.WRITELOG(P_EVENTTYPE      => pkg_lib_ex_basic_newapi.EventT_FuncError,
                              P_SESSION        => vEX_SESSION,
                              P_IDENTIFICATION => $$PLSQL_UNIT,
                              P_TXT            => 'Fehler, Protokoll-Datei kann nicht umbenannt werden');
  WHEN UTL_FILE.INVALID_OPERATION THEN
    pkg_lib_ex_basic_newapi.WRITELOG(P_EVENTTYPE      => pkg_lib_ex_basic_newapi.EventT_FuncError,
                              P_SESSION        => vEX_SESSION,
                              P_IDENTIFICATION => $$PLSQL_UNIT,
                              P_TXT            => 'Fehler, Protokoll-Datei kann nicht geschrieben werden');
WHEN OTHERS THEN
  UTL_FILE.FCLOSE(l_file);
    pkg_lib_ex_basic_newapi.WRITELOG(P_EVENTTYPE      => pkg_lib_ex_basic_newapi.EventT_SysError,
                              P_SESSION        => vEX_SESSION,
                              P_IDENTIFICATION => $$PLSQL_UNIT,
                              P_TXT            => 'Systemfehler');
END createStatusProtokoll;

/*------------------------------------------------------------------------------
Name:         excludeJobs

Parameter:    keine

Beschreibung: Setzt Jobs auf Status "I" anhand einert Ausschlussliste
              Diese Jobs werden nicht exportiert, dafuer beim checken der
              Jobs aber protokolliert.
------------------------------------------------------------------------------*/
PROCEDURE excludeJobs
IS

 vSTL_LINE     STAT_EXP_DATAMART_MCFG.WRK_EX_OBJ.STL_LINE%TYPE;
 vSTL_OBJ_NAME VARCHAR2(200);
 vSTL_SCHEMA VARCHAR2(200);
BEGIN

  logInfo('Start prüfen Ausschluss-Jobs');

  -- fuer alle Ausschluss-Objekte.
  FOR cOBJ IN
  (SELECT  STL_OBJ_NAME,STL_ID,STL_SCHEMA
  FROM STAT_EXP_DATAMART_MCFG.WRK_EX_OBJ
  WHERE STL_STATUS = 'E'
  AND STL_SESSION   = vEX_SESSION
  ORDER BY STL_ID DESC
  )
  LOOP
    vSTL_OBJ_NAME :=  '^' || REPLACE(star2regex(upper(cOBJ.STL_OBJ_NAME)),'?','(.{1,1})') ||  '$';
    vSTL_SCHEMA :=  '^' || REPLACE(star2regex(upper(cOBJ.STL_SCHEMA)),'?','(.{1,1})') ||  '$';
    -- fuer alle Ausschluss-Jobs.
    FOR cJOB IN
    (SELECT   JOB_ID
    FROM STAT_EXP_DATAMART_MCFG.WRK_EX_JOB
    WHERE JOB_SESSION = vEX_SESSION
      AND REGEXP_LIKE (upper(job_schema),vSTL_SCHEMA)
      AND REGEXP_LIKE ( upper(JOB_OBJ_NAME), vSTL_OBJ_NAME)
                       )
    LOOP
	  UPDATE STAT_EXP_DATAMART_MCFG.WRK_EX_JOB SET JOB_STATUS = 'E',JOB_STL_ID_EXCL=cOBJ.STL_ID WHERE JOB_ID = cJOB.JOB_ID AND JOB_STL_ID_EXCL IS NULL;
      --setStatusJOB(cJOB.JOB_ID, 'E');
      pkg_lib_ex_basic_newapi.WRITELOG(P_EVENTTYPE      => pkg_lib_ex_basic_newapi.EventT_Information,
                                P_SESSION        => vEX_SESSION,
                                P_IDENTIFICATION => $$PLSQL_UNIT || '_JOB_ID_' || cJOB.JOB_ID,
                                P_TXT            => 'Job (JOB_ID=' || cJOB.JOB_ID || ') wird wegen Treffer in Ausschluss-Liste ignoriert' );
    END LOOP; -- cJOB
  END LOOP; --cOBJ
  commit;

  logInfo('Ende pruefen Ausschluss-Jobs');

EXCEPTION
WHEN OTHERS THEN
    pkg_lib_ex_basic_newapi.WRITELOG(P_EVENTTYPE      => pkg_lib_ex_basic_newapi.EventT_SysError,
                              P_SESSION        => vEX_SESSION,
                              P_IDENTIFICATION => $$PLSQL_UNIT,
                              P_TXT            => 'Systemfehler');
END excludeJobs;

/*------------------------------------------------------------------------------
Name:         excludeMultipleJobs

Parameter:    keine

Beschreibung: Setzt merfach vorkommentde (Fachlich) Jobs auf Status "I".
              Diese Jobs werden nicht exportiert, dafuer beim checken der
              Jobs aber protokolliert.
------------------------------------------------------------------------------*/
PROCEDURE excludeMultipleJobs
IS

 vSTL_LINE  STAT_EXP_DATAMART_MCFG.WRK_EX_OBJ.STL_LINE%TYPE;

BEGIN

  logInfo('Start prüfen mehrfach vorkommende Jobs');

  -- fuer alle Jobs die merfach vorkommentde.
  FOR cJOB IN
  (SELECT *
  FROM STAT_EXP_DATAMART_MCFG.WRK_EX_JOB
  WHERE JOB_SESSION = vEX_SESSION
  AND JOB_ID NOT   IN
    (SELECT MIN(JOB_ID) AS JOB_ID
    FROM STAT_EXP_DATAMART_MCFG.WRK_EX_JOB ok
    WHERE JOB_SESSION = vEX_SESSION
    GROUP BY JOB_OBJ_NAME,
      JOB_OBJ_TYPE,
      JOB_PART_COL,
      JOB_PART,
      JOB_KZ,
      JOB_SCHEMA,
      JOB_SESSION,
      JOB_STATUS,
      JOB_ANZ_MAT,
      JOB_ANZ_DUMP
    )
  )
  LOOP
      SELECT STL_LINE
        INTO vSTL_LINE
        FROM STAT_EXP_DATAMART_MCFG.WRK_EX_OBJ
       WHERE STL_ID = cJOB.JOB_STL_ID;

      setStatusJOB(cJOB.JOB_ID, 'I');
      pkg_lib_ex_basic_newapi.WRITELOG(P_EVENTTYPE      => pkg_lib_ex_basic_newapi.EventT_Warning,
                                P_SESSION        => vEX_SESSION,
                                P_IDENTIFICATION => $$PLSQL_UNIT || '_JOB_ID_' || cJOB.JOB_ID,
                                P_TXT            => 'Steuerlisteneintag "' || vSTL_LINE || '" produziert Job (JOB_ID=' || cJOB.JOB_ID || ') fuer ein Objekt welches in dieser Steuerliste schon exportiert wird.' );
  END LOOP;

  logInfo('Ende prüfen mehrfach vorkommende Jobs');

EXCEPTION
WHEN OTHERS THEN
    pkg_lib_ex_basic_newapi.WRITELOG(P_EVENTTYPE      => pkg_lib_ex_basic_newapi.EventT_SysError,
                              P_SESSION        => vEX_SESSION,
                              P_IDENTIFICATION => $$PLSQL_UNIT,
                              P_TXT            => 'Systemfehler');
END excludeMultipleJobs;


/*------------------------------------------------------------------------------
Name:         checkJobs

Parameter:    keine

Beschreibung: ueberprueft anzahl exportierter Datensaetze fuer Alle Jobs.
------------------------------------------------------------------------------*/
PROCEDURE checkJobs
IS
  vStatus STAT_EXP_DATAMART_MCFG.WRK_EX_OBJ.STL_STATUS%TYPE;
  vSTLID  STAT_EXP_DATAMART_MCFG.WRK_EX_OBJ.STL_ID%TYPE;
  vAnz    INTEGER;
BEGIN
  logInfo('Start prüfen, ob alle Jobs exportiert wurden');
  -- fuer alle Jobs bei denen nicht alle Datensaestze exportiert wurden.
  FOR cJOB IN
  (SELECT *
     FROM STAT_EXP_DATAMART_MCFG.WRK_EX_JOB
    WHERE JOB_SESSION = vEX_SESSION
      AND JOB_STATUS  = 'A'
      AND JOB_OBJ_TYPE NOT IN ('t', 'm')  -- Nur da woch die Objekte auch Materialisiert wurden.
      AND JOB_ANZ_DUMP <> JOB_ANZ_MAT
  )
  LOOP
    setStatusJOB(cJOB.JOB_ID, 'F');
    pkg_lib_ex_basic_newapi.WRITELOG(P_EVENTTYPE      => pkg_lib_ex_basic_newapi.EventT_FuncError,
                              P_SESSION        => vEX_SESSION,
                              P_IDENTIFICATION => $$PLSQL_UNIT,
                              P_TXT            => 'Fehler, anzahl materialisierter Datensaeze stimmt nicht ueberein mit anzahl exportierter Datensaetze');
  END LOOP;

  -- fuer alle Steuerlisten-Objekte mit generierten Jobs
  FOR cOBJ IN
  (SELECT *
     FROM STAT_EXP_DATAMART_MCFG.WRK_EX_OBJ
    WHERE STL_STATUS = 'J'
      AND STL_SESSION = vEX_SESSION
  )
  LOOP
    vSTLID := cOBJ.STL_ID;
    -- pruefung ob Statuslisten-Eintraege ohne Jobs
    SELECT count(*)
      INTO vAnz
      FROM STAT_EXP_DATAMART_MCFG.WRK_EX_JOB
      WHERE JOB_STL_ID = cOBJ.STL_ID
      AND JOB_SESSION = vEX_SESSION;
    IF vAnz = 0 THEN
      setStatusOBJ (cOBJ.STL_ID, 'F'); --Status fehlerhaft setzen
      logInfo('Für Statuslisten-Eintrag ' || cOBJ.STL_LINE || ' konnten keine Jobs generiert werden');
    END IF;
    -- pruefung ob fehlerhafte Jobs
    BEGIN --Pruefung
      SELECT DISTINCT JOB_STATUS
      INTO vStatus
      FROM STAT_EXP_DATAMART_MCFG.WRK_EX_JOB
      WHERE JOB_STL_ID = cOBJ.STL_ID
      AND JOB_STATUS   = 'F'
      AND JOB_SESSION = vEX_SESSION;
      setStatusOBJ (cOBJ.STL_ID, vStatus); --Status fehlerhaft setzen
      logInfo('Statuslisten-Objekt ' || cOBJ.STL_OBJ_NAME || ' hat fehlerhafte Jobs');

    EXCEPTION
    WHEN NO_DATA_FOUND THEN
      setStatusOBJ (vSTLID, 'A'); --Status Abgeschlossen setzen
    END;                         --Pruefung
  END LOOP;

    -- fuer alle ausgeschlossenen Steuerlisten-Objekte
  FOR cOBJ IN
  (SELECT *
     FROM STAT_EXP_DATAMART_MCFG.WRK_EX_OBJ
    WHERE STL_STATUS = 'E'
      AND STL_SESSION = vEX_SESSION
  )
  LOOP
    vSTLID := cOBJ.STL_ID;

    -- pruefung ob Statuslisten-Eintraege ohne Jobs
    SELECT count(*)
      INTO vAnz
      FROM STAT_EXP_DATAMART_MCFG.WRK_EX_JOB
      WHERE JOB_STL_ID = cOBJ.STL_ID
      AND JOB_SESSION = vEX_SESSION;

    IF vAnz = 0 THEN
      pkg_lib_ex_basic_newapi.WRITELOG(P_EVENTTYPE      => pkg_lib_ex_basic_newapi.EventT_Warning,
                              P_SESSION        => vEX_SESSION,
                              P_IDENTIFICATION => $$PLSQL_UNIT,
                              P_TXT            => 'Für Ausschluss-Statuslisten-Eintrag ' || cOBJ.STL_LINE || ' konnten keine Jobs generiert werden');
    END IF;

  END LOOP;

  logInfo('Ende prüfen, ob alle Jobs exportiert wurden');
EXCEPTION
WHEN OTHERS THEN
    pkg_lib_ex_basic_newapi.WRITELOG(P_EVENTTYPE      => pkg_lib_ex_basic_newapi.EventT_SysError,
                              P_SESSION        => vEX_SESSION,
                              P_IDENTIFICATION => $$PLSQL_UNIT,
                              P_TXT            => 'Systemfehler');
END checkJobs;
/*------------------------------------------------------------------------------
Name:        dumpJobList

Parameter:    p_STL_NAME   Steuerlistenname

Beschreibung: Erzeugt ein Protokoll mit allen Logeintaegen.
------------------------------------------------------------------------------*/
PROCEDURE dumpJobList (p_STL_NAME VARCHAR2)
IS
l_file              UTL_FILE.FILE_TYPE;
vFilename1 constant VARCHAR2(200) := 'dump.joblist.' || vEX_SESSION;
vFilename2 constant VARCHAR2(200) := replace(replace(p_STL_NAME,'.csv',''),'.txt','')||'.sim.csv';
vSchema             VARCHAR2(128);
vSTL_KZ             CHAR(1);
vLine               VARCHAR2(4000);
vDirName            VARCHAR2(200);
vTYPE               VARCHAR2(9);
i                   INTEGER;
vExistsIAB          INTEGER;
vPath               VARCHAR2(4000);
vPathTgt            VARCHAR2(4000);
vUnixCom            VARCHAR2(4000);
BEGIN
  logInfo ('Beginn erzeugen DumpJobList');



  l_file := UTL_FILE.FOPEN(
            location => 'EX_TEMP',
            filename => vFilename1
            ,open_mode => 'w',
            max_linesize => 32767);

       -- Infos aus Dateinamen befuellen (Zerlegung des Dateinamens)
      i := 1;
      FOR cParam IN
      (SELECT * FROM TABLE(STAT_EXP_DATAMART_MCFG.PKG_LIB_EX_BASIC_NEWAPI.splitSTL( p_STL_NAME))
      )
      LOOP
        IF i                 = 1 THEN
           IF cParam.COLUMN_VALUE IN ( 'b', 'f', 'd', 'v') THEN
              vSTL_KZ     := cParam.COLUMN_VALUE;
           ELSE
              vSTL_KZ     := '0';
           END IF;
        ELSIF i              = 2 THEN
          NULL ; --wird hier nicht gebraucht.
        ELSIF i              = 3 THEN
          vSchema := upper(cParam.COLUMN_VALUE);
        ELSE
          NULL;
        END IF;
        i := i + 1;
      END LOOP; -- Infos aus Dateinamen befuellen (Zerlegung des Dateinamens)

    vExistsIAB  := dbms_lob.fileexists(bfilename('EX_STL_IAB',  p_STL_NAME));
    vDirName  := getTargetDir (vSTL_KZ, vSchema);

    IF vExistsIAB <> 0 THEN
      vDirName  := 'EX_IAB';
    END IF;

      vLine:='stlId;stlLine;stlLineTyp;stlSchema;stlObjName;jobSchema;jobObjName;jobPartitionColumn;jobPartition;jobStatus;jobStlIdExcl';
      UTL_FILE.PUT_LINE(l_file, vLine);

    -- fuer alle Protokolleintraege zu dem Lauf
    FOR cJL IN
    (
         select stl_id,stl_line
            ,decode(stl_status,'E','E','I')stl_status
            ,stl_schema
            ,o.stl_obj_name
            ,job_schema
            ,j.job_obj_name
            ,case when job_part_col='NO' then null else job_part_col end job_part_col
            ,case when job_part='NO' then null else job_part end job_part
            ,decode(job_status,'C','resolved','I','ignored,Duplicate','E','excluded',null) job_status
            ,job_stl_id_excl
        from wrk_ex_obj o
        left outer join wrk_ex_job j on (decode(stl_status,'E',j.job_stl_id_excl,j.job_stl_id)=o.stl_id and j.job_session=o.stl_session and job_status!='X')
        where o.stl_session=vEX_SESSION
        order by stl_session,stl_id,job_schema,job_obj_name,job_part
    )
    LOOP

      vLine := cJL.stl_id||';'||cJL.stl_line||';'||cJL.stl_status||';'||cJL.stl_schema||';'||cJL.stl_obj_name||';'||cJL.job_schema||';'
               || cJL.job_obj_name||';'||cJL.job_part_col||';'||cJL.job_part||';'||cJL.job_status||';'||cJL.job_stl_id_excl;
      UTL_FILE.PUT_LINE(l_file, vLine);
    END LOOP;


  UTL_FILE.FCLOSE(l_file);

  logInfo(vDirName||':'||vFilename2);
  UTL_FILE.FRENAME('EX_TEMP', vFilename1, vDirName, vFilename2, TRUE);

--  vPath := PKG_LIB_EX_BASIC_NEWAPI.getPath('EX_TEMP');
--  vPathTgt := PKG_LIB_EX_BASIC_NEWAPI.getPath(vDirName);
--  vUnixCom := 'mv '||vPath||'/'||vFilename1||' '||vPathTgt||'/'||vFilename2;
--  PKG_LIB_EX_BASIC_NEWAPI.CALLUNIX(vUnixCom, vEX_SESSION);

  logInfo ('Ende erzeugen DumpJobList');
EXCEPTION
  WHEN UTL_FILE.INVALID_PATH THEN
    PKG_LIB_EX_BASIC_NEWAPI.WRITELOG(P_EVENTTYPE      => PKG_LIB_EX_BASIC_NEWAPI.EventT_FuncError,
                              P_SESSION        => vEX_SESSION,
                              P_IDENTIFICATION => $$PLSQL_UNIT,
                              P_TXT            => 'Fehler, Pfad fuer Dump-Datei ist nicht valide');
  WHEN UTL_FILE.RENAME_FAILED THEN
    PKG_LIB_EX_BASIC_NEWAPI.WRITELOG(P_EVENTTYPE      => PKG_LIB_EX_BASIC_NEWAPI.EventT_FuncError,
                              P_SESSION        => vEX_SESSION,
                              P_IDENTIFICATION => $$PLSQL_UNIT,
                              P_TXT            => 'Fehler, Dump-Datei kann nicht umbenannt werden');
  WHEN UTL_FILE.INVALID_OPERATION THEN
    PKG_LIB_EX_BASIC_NEWAPI.WRITELOG(P_EVENTTYPE      => PKG_LIB_EX_BASIC_NEWAPI.EventT_FuncError,
                              P_SESSION        => vEX_SESSION,
                              P_IDENTIFICATION => $$PLSQL_UNIT,
                              P_TXT            => 'Fehler, Dump-Datei kann nicht geschrieben werden');
WHEN OTHERS THEN
  UTL_FILE.FCLOSE(l_file);
    PKG_LIB_EX_BASIC_NEWAPI.WRITELOG(P_EVENTTYPE      => PKG_LIB_EX_BASIC_NEWAPI.EventT_SysError,
                              P_SESSION        => vEX_SESSION,
                              P_IDENTIFICATION => $$PLSQL_UNIT,
                              P_TXT            => 'Systemfehler');
END dumpJobList;


/*------------------------------------------------------------------------------
Name:         loadSTL

Parameter:    p_STL_NAME  Name der Steuerliste die exportiert werden soll
              p_RELOAD    Steuert den Ablauf, Normal- oder Wiederanlauf

Beschreibung: Prozedur pr?ob IAB- oder Normale-Steuerliste und startet die
              entsprechende Verarbeitung.
------------------------------------------------------------------------------*/
procedure loadSTL( p_STL_NAME VARCHAR2, p_RELOAD BOOLEAN DEFAULT FALSE, p_Ref_Date date default null) IS
vExists    INTEGER;
vExistsIAB INTEGER;
BEGIN
   logInfo('Start laden Steuerliste ' || p_STL_NAME);

   vExists     := dbms_lob.fileexists(bfilename('EX_STL_TODO', p_STL_NAME));
   vExistsIAB  := dbms_lob.fileexists(bfilename('EX_STL_IAB',  p_STL_NAME));

   IF vExists <> 0 THEN
     PKG_EX_LOAD_STL.getSTL(vEX_SESSION, p_STL_NAME, p_Ref_Date);
     logInfo('Normale Steuerliste ' || p_STL_NAME);
   ELSIF vExistsIAB <> 0 THEN
     PKG_EX_LOAD_STL.getSTLIAB(vEX_SESSION, p_STL_NAME,p_Ref_Date);
     logInfo('IAB Steuerliste ' || p_STL_NAME);
   ELSE
     pkg_lib_ex_basic_newapi.WRITELOG(P_EVENTTYPE      => pkg_lib_ex_basic_newapi.EventT_FuncError,
                              P_SESSION        => vEX_SESSION,
                              P_IDENTIFICATION => $$PLSQL_UNIT,
                              P_TXT            => 'Fehler, Steuerliste ' || p_STL_NAME || ' kann nicht gelesen werden');
   END IF;

   -- wenn reload dann nur die die im vorherigen Lauf nicht exportiert wurden
   IF p_RELOAD THEN
     PKG_EX_LOAD_STL.setReload(vEX_SESSION, p_STL_NAME,p_Ref_Date);
     logInfo('Verarbeitung der Steuerliste ' || p_STL_NAME || ' erfolgte im Modus "RELOAD"');
   END IF;
   logInfo('Ende laden Steuerliste '   || p_STL_NAME);
EXCEPTION
  WHEN OTHERS THEN
    pkg_lib_ex_basic_newapi.WRITELOG(P_EVENTTYPE      => pkg_lib_ex_basic_newapi.EventT_SysError,
                              P_SESSION        => vEX_SESSION,
                              P_IDENTIFICATION => $$PLSQL_UNIT,
                              P_TXT            => 'Systemfehler');
    RAISE;
END loadSTL;


/*------------------------------------------------------------------------------
Name:         doneSTL

Parameter:    p_STL_NAME  Name der Steuerliste die exportiert werden soll

Beschreibung: Prozedur verschiebt Steuerliste ins ToDo Verzeichis wenn sie
              fehlerfrei exportiert wurde.
------------------------------------------------------------------------------*/
procedure doneSTL( p_STL_NAME VARCHAR2) IS
BEGIN
  logInfo('Start abschliessen der Steuerliste ' || p_STL_NAME);
  PKG_EX_LOAD_STL.setSTLDone(vEX_SESSION, p_STL_NAME);
  logInfo('Ende abschliessen der Steuerliste '  || p_STL_NAME);
EXCEPTION
  WHEN OTHERS THEN
    pkg_lib_ex_basic_newapi.WRITELOG(P_EVENTTYPE      => pkg_lib_ex_basic_newapi.EventT_SysError,
                              P_SESSION        => vEX_SESSION,
                              P_IDENTIFICATION => $$PLSQL_UNIT,
                              P_TXT            => 'Systemfehler');
END doneSTL;

/*------------------------------------------------------------------------------
Name:         checkUnixAPI
Parameter:

Beschreibung: Prozedur führt einen Test-Aufruf über das UNIX-API durch.
------------------------------------------------------------------------------*/
procedure checkUnixAPI
as
begin
  logInfo('Start checkUnixAPI ');
  PKG_LIB_EX_BASIC_NEWAPI.checkUnixAPI(vEX_SESSION);
  logInfo('Ende checkUnixAPI');
EXCEPTION
  WHEN OTHERS THEN
    pkg_lib_ex_basic_newapi.WRITELOG(P_EVENTTYPE      => pkg_lib_ex_basic_newapi.EventT_SysError,
                              P_SESSION        => vEX_SESSION,
                              P_IDENTIFICATION => $$PLSQL_UNIT,
                              P_TXT            => 'Systemfehler');
    RAISE;
END;



/*------------------------------------------------------------------------------
Name:         startWork

Parameter: p_STL_NAME  Name der Steuerliste die exportiert werden soll.
                       (ohne Pfad)
           p_RELOAD    Steuert den Ablauf, Normal- oder Wiederanlauf
                       (Default FALSE also KEIN Wiederanlauf)
           p_PARALLEL  Steuert die Paralleliaet mit welcher der DM-Export
                       ausgeführt wird. (Default NULL bedeutet dass die
                       Parallelität automatisch ermittelt wird)
           p_DELIM     Trennzeichen zwischen Spalten (Default "|" erlaubt sind:
                       chr(xxx), ;, |, ...)
           p_PIPE      Steuert, ob der Export über Pipes läuft oder über Files
                       (Default ist über Pipes)
           p_VIEWOBJ   Steuert ob Views, bei Dimensionen, Materialisiert
                       exportiert werden oder ob die Views in Einzelobjejte
                       aufgelöst werden.(Default TRUE also Dimensions-Views
                       werden materialisiert)
           p_EXP       Steuert ob exportiert werden soll oder ob nach dem
                       aufarbeiten der Steuerliste aufgehört werden soll.
                       Wird nur für Test- und Wartungszweck benötigt.
                       (Default TRUE also vollständig exportiert)
           p_REFDATE   Referenzdatum für MONID-Ermittlung,Default ist SYSDATE

           p_DO_UTF8_UNLOOAD
                       Default:false, True: Unload erfolgt in UTF8

Beschreibung: Prozedur startet die Verarbeitung.
              - Erzeugt die benötigten Verzeichnisse Falls nicht vorhanden
              - Räumt alle Dateien weg, die älter als 90 Tage sind
              - Lädt die Steuerliste
              - Erzeugt die Export-Jobs
              - Erzeugt die Verzeichnisse für die Jobs falls nicht vorhanden
              - Exportiert die Jobs
              - Statusliste Semaphore und Protokoll werden erzeugt.
------------------------------------------------------------------------------*/
procedure startWork( p_STL_NAME VARCHAR2,
                     p_RELOAD   BOOLEAN DEFAULT FALSE,
                     p_PARALLEL INTEGER DEFAULT NULL,
                     p_DELIM    VARCHAR2 DEFAULT 'chr(124)',
                     p_PIPE     BOOLEAN DEFAULT TRUE,
                     p_VIEW_MAT BOOLEAN DEFAULT TRUE,
                     p_EXP      BOOLEAN DEFAULT TRUE,
                     p_REFDATE  DATE DEFAULT SYSDATE,
                     p_DO_UTF8_Unload BOOLEAN DEFAULT FALSE) IS
vERR BOOLEAN;
BEGIN
   logInfo('S T A R T - D M _ E X P O R T');
   DBMS_OUTPUT.PUT_LINE('S T A R T - D M _ E X P O R T');

   -- Referenzdatum für MONID-Ersetzung
   vRefDate:=p_REFDATE;

   -- Prüfen, ob UnixAPI verfügbar ist
   checkUnixAPI;

   createDirDefault;
   loadSTL(p_STL_NAME,p_RELOAD,p_REFDATE);
   vERR := checkSTL;
   IF NOT vERR THEN
     vPARALLEL     := p_PARALLEL;
     vPIPE         := p_PIPE;
     vVIEW_MAT     := p_VIEW_MAT;
     vDoUtf8Unload := p_DO_UTF8_UNLOAD;
     setDelimiter(p_DELIM);
     createJobs;
     createDirJobs(p_STL_NAME);
     excludeJobs;
     excludeMultipleJobs;
     IF p_EXP THEN
       exportJobs;
       checkJobs;
       -- 08.10.2020 Semaphores werden direkt nach dem Entladen eines Jobs angelegt
       --createSemaphores;
       createStatusList(p_RELOAD, p_STL_NAME);
       doneSTL(p_STL_NAME);
     ELSE
       dumpJobList(p_STL_NAME);
     END IF;
     createStatusProtokoll(p_STL_NAME);
   ELSE
     createStatusProtokoll(p_STL_NAME);
     Raise_Application_Error (-20344, 'STEUERLISTEN-FEHLER ');
   END IF;
   doHousekeeping;
   logInfo ('E N D E - D M _ E X P O R T');
   DBMS_OUTPUT.PUT_LINE('Für mehr Infos siehe status.export.protokoll.txt');
   DBMS_OUTPUT.PUT_LINE('E N D E - D M _ E X P O R T');
EXCEPTION
  WHEN OTHERS THEN
    DBMS_OUTPUT.PUT_LINE('Für mehr Infos siehe status.export.protokoll.txt');
    DBMS_OUTPUT.PUT_LINE('F E H L E R - E N D E - D M _ E X P O R T');
    pkg_lib_ex_basic_newapi.WRITELOG(P_EVENTTYPE      => pkg_lib_ex_basic_newapi.EventT_SysError,
                              P_SESSION        => vEX_SESSION,
                              P_IDENTIFICATION => $$PLSQL_UNIT,
                              P_TXT            => 'Systemfehler');
    RAISE;
END startWork;


procedure testSemaphore(pSession in number) IS


begin
vEX_SESSION:=pSession;

createSemaphores;

end;


END PKG_EX_WORK_NEWAPI;
/
