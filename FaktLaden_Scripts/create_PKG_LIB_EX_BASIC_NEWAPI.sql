CREATE OR REPLACE PACKAGE                          "PKG_LIB_EX_BASIC_NEWAPI" IS
-- $Id: create_PKG_LIB_EX_BASIC_NEWAPI.sql 1178 2022-04-26 05:02:18Z seideld006 $
-- ---------------------------------------------------------------------------------
--@START DOKU
-- ---------------------------------------------------------------------------------
-- MODUL
-- Revision     $LastChangedRevision: 1178 $
-- Last Revised $LastChangedDate: 2022-04-26 07:02:18 +0200 (Di, 26 Apr 2022) $
-- Author       $LastChangedBy: seideld006 $
-- ---------------------------------------------------------------------------------
-- DATEI
--  File         $HeadURL: https://svn.sdst.sbaintern.de/put/trunk/Anwendungen/DMExport/src/stat_exp_datamart_mcfg/sql/packages/create_PKG_LIB_EX_BASIC_NEWAPI.sql $
--
-- SYNTAX
--
-- BENOETIGT
-- Tabelle ERR_EX_PROT
--         TMP_EX_DIR_LIST
--
--
-- BESCHREIBUNG
--  Allgemeine Hilfsfunktionen für die Exportschnitstelle
--  Beinhaltet unter anderm eine Logging-Funktion die  benötigte Typen von
--  Events in einer Tabelle protokolliert.
--
--  EventT_SysError    Systemfehler
--  EventT_FuncError   Fachliche Fehler
--  EventT_Warning     Warnungen
--  EventT_Information Informationen

--  04.03.2015 CebucL - Initiale Version
--  15.06.2015 CebucL - Kleinere Erweiterungen
--  10.09.2018 SeidelD006
--                    - Unix-Kommandos über Scheduler-API
--  07.07.2020 SeidelD006
--                    - Neuer Parameter JobId für writeLog
--  19.01.2021 SeidelD006
--                    - Referenzen auf PKG_EX_WORK auf PKG_EX_WORK_NEWAPI geändert
--                    - Neue Funktionalität fopen mit Batchanlage von leeren Dateien
--                      damit die Penalty beim FOPEN durch den Scheduler-Einsatz
--                      nahezu eliminiert wird.
--  05.04.2022 SeidelD006
--                    - Prozedur checkUnixAPI zur Überprüfung des Unix-APIs angelegt
--  26.04.2022 SeidelD006
--                    - Fehler in Prozedur checkUnixAPI behoben
--
--@ENDE DOKU
-- ---------------------------------------------------------------------------------
  CurrentUser VARCHAR2(64) := USER;

TYPE split_tbl_stl
IS
  TABLE OF           VARCHAR2(60);

  EventT_SysError    CONSTANT VARCHAR2(2 BYTE) := 'S';
  EventT_FuncError   CONSTANT VARCHAR2(2 BYTE) := 'E';
  EventT_Warning     CONSTANT VARCHAR2(2 BYTE) := 'W';
  EventT_Information CONSTANT VARCHAR2(2 BYTE) := 'I';


/*------------------------------------------------------------------------------
Name:         writeLog

Parameter: P_EVENTTYPE      Event-Typ z.B. SysError (definier als Package konstanten)
           P_IDENTIFICATION Identifikator z.B. Packagename wo das Event stattgefunden hat
           P_TXT            Text für das Event
           P_PARAMSTRING1   Zusatztext für das Event optional
           P_PARAMSTRING2   Zusatztext für das Event optional
           P_PARAMSTRING3   Zusatztext für das Event optional
           P_PARAMSTRING4   Zusatztext für das Event optional
           P_PARAMSTRING5   Zusatztext für das Event optional
           P_PARAMSTRING6   Zusatztext für das Event optional
           P_PARAMSTRING7   Zusatztext für das Event optional
           P_PARAMSTRING8   Zusatztext für das Event optional
           P_PARAMSTRING9   Zusatztext für das Event optional

Beschreibung: Prozedur legt einen neuen Eintrag für die übergebene Protokolltabelle
              in der Tabelle UEB_HOUSEKEEPING an. Falls der Insert fehlschlägt,
              wird der Parameter p_success auf FALSE gesetzt.

------------------------------------------------------------------------------*/
  PROCEDURE writeLog(
      P_EVENTTYPE      IN VARCHAR2 DEFAULT EventT_SysError ,
      P_SESSION        IN VARCHAR2 DEFAULT dbms_session.unique_session_id ,
      P_JOB_ID         IN NUMBER DEFAULT NULL,
      P_IDENTIFICATION IN ERR_EX_PROT.IDENTIFICATION%TYPE DEFAULT $$PLSQL_UNIT ,
      P_TXT            IN ERR_EX_PROT.TEXT%TYPE ,
      P_PARAMSTRING1   IN VARCHAR2 := '(?)' ,
      P_PARAMSTRING2   IN VARCHAR2 := '(?)' ,
      P_PARAMSTRING3   IN VARCHAR2 := '(?)' ,
      P_PARAMSTRING4   IN VARCHAR2 := '(?)' ,
      P_PARAMSTRING5   IN VARCHAR2 := '(?)' ,
      P_PARAMSTRING6   IN VARCHAR2 := '(?)' ,
      P_PARAMSTRING7   IN VARCHAR2 := '(?)' ,
      P_PARAMSTRING8   IN VARCHAR2 := '(?)' ,
      P_PARAMSTRING9   IN VARCHAR2 := '(?)' );

/*-----------------------------------------------------------------------------
Name:         splitSTL

Parameter:    p_list          String der gesplittet werden soll
              p_del           Zeichen an Hand dem gesplittet werden soll

Return:       split_tbl_stl Tabelle mit den gesplitteten Teil-Strings

Beschreibung: Liefert eine Tabelle vom Typ split_tbl_stl mit den gesplitteten
              Teil-Strings. Default bei dem Zeichen, an Hand dem gesplittet werden
              soll, ist '.'
-----------------------------------------------------------------------------*/
  FUNCTION splitSTL(
      p_list    IN VARCHAR2,
      p_del     IN VARCHAR2 := '.',
      P_SESSION IN VARCHAR2 DEFAULT dbms_session.unique_session_id )
    RETURN split_tbl_stl pipelined;

/*------------------------------------------------------------------------------
Name:         getPath

Parameter: pDir     Name eines Directory-Objekts

Beschreibung: Prozedur gibt den Pfad  eines Directory-Objekts zurück. Wenn es das
              Directory-Objekt nicht gibt wird NULL zurückgegeben.
------------------------------------------------------------------------------*/
  FUNCTION getPath(
      pDir      IN  VARCHAR2,
      P_SESSION IN VARCHAR2 DEFAULT dbms_session.unique_session_id )
    RETURN VARCHAR2;

/*------------------------------------------------------------------------------
Name:         callUnix

Parameter: pCommand     Unix Komando

Beschreibung: Die Prozedur nimmt das Kommando als Parameter entgegen, führt es
              aus und gibt die ersten 32 Kb der Ausgabe an den Aufrufer zurück.
------------------------------------------------------------------------------*/
  PROCEDURE callUnix( pCommand  IN VARCHAR2,
                      P_SESSION IN VARCHAR2 DEFAULT dbms_session.unique_session_id );

/*------------------------------------------------------------------------------
Name:         callUnixAsync

Parameter: pCommand     Unix Komando

Beschreibung: Die Prozedur nimmt das Kommando als Parameter entgegen, führt es
              aus und gibt die ersten 32 Kb der Ausgabe an den Aufrufer zurück.
              Es wird nicht gewartet bis der Afruf abgeschlossen ist
------------------------------------------------------------------------------*/
  PROCEDURE callUnixAsync( pCommand  IN VARCHAR2,
                           P_SESSION IN VARCHAR2 DEFAULT dbms_session.unique_session_id );


/*------------------------------------------------------------------------------
Name:         checkUnixAPI

Parameter: pCommand     Unix Komando

Beschreibung: Die Prozedur nimmt das Kommando als Parameter entgegen, führt es
              aus und gibt die ersten 32 Kb der Ausgabe an den Aufrufer zurück.
------------------------------------------------------------------------------*/
  PROCEDURE checkUnixAPI(P_SESSION IN VARCHAR2 DEFAULT dbms_session.unique_session_id );
/*------------------------------------------------------------------------------
Name:         fopen

Parameter: pSession        Session-Id des laufenden DM-Exports
           pTempDir        Name des temporären Directory-Objects
           pLocation       Ziel-Directory
           pFileName       Name der zu öffenden Datei

Return:    Filehandle

Beschreibung: Die Funktion benennt eine Datei aus dem im Vorfeld angelegten
            Datei-Pool um und verschiebt sie in das Ziel-Verzeichnis.
            Falls keine vorangelegten Dateien mehr vorhanden sind, so werden
            wieder welche angelegt.
------------------------------------------------------------------------------*/
FUNCTION fopen(pSession IN NUMBER,pTempDir in varchar2,pMode in varchar2,pLocation in varchar2,pFileName in varchar2) return UTL_FILE.FILE_TYPE;


/*------------------------------------------------------------------------------
Name:         prc_RemovePreCreateFiles

Parameter: pSession        Session-Id des laufenden DM-Exports
           pTempDir        Name des temporären Directory-Objects


Beschreibung: Die Prozedur räumt zu viel angelete Dateien aus dem Pool
              wieder weg.
------------------------------------------------------------------------------*/
PROCEDURE prc_RemovePreCreateFiles(pSession IN NUMBER,pTempDir in varchar2);

END PKG_LIB_EX_BASIC_NEWAPI;
/


CREATE OR REPLACE PACKAGE BODY                          "PKG_LIB_EX_BASIC_NEWAPI"
IS

-- Zähler für FilePreCreate
gFPCCurrentId NUMBER := 0;
gFPCMaxId NUMBER := 0;
gFPCMaxFiles NUMBER := 30;
gFPCFileTemplate varchar2(100) := 'FilePreCreate.<SESSION>.<ID>';


/*------------------------------------------------------------------------------
Name:         writeLog

Parameter: P_EVENTTYPE      Event-Typ z.B. SysError (definier als Package konstanten)
           P_IDENTIFICATION Identifikator z.B. Packagename wo das Event stattgefunden hat
           P_TXT            Text für das Event
           P_PARAMSTRING1   Zusatztext für das Event optional
           P_PARAMSTRING2   Zusatztext für das Event optional
           P_PARAMSTRING3   Zusatztext für das Event optional
           P_PARAMSTRING4   Zusatztext für das Event optional
           P_PARAMSTRING5   Zusatztext für das Event optional
           P_PARAMSTRING6   Zusatztext für das Event optional
           P_PARAMSTRING7   Zusatztext für das Event optional
           P_PARAMSTRING8   Zusatztext für das Event optional
           P_PARAMSTRING9   Zusatztext für das Event optional

Beschreibung: Prozedur legt einen neuen Eintrag für die übergebene Protokolltabelle
              in der Tabelle UEB_HOUSEKEEPING an. Falls der Insert fehlschlägt,
              wird der Parameter p_success auf FALSE gesetzt.

------------------------------------------------------------------------------*/
  PROCEDURE writeLog(
      P_EVENTTYPE      IN VARCHAR2 DEFAULT EventT_SysError ,
      P_SESSION        IN VARCHAR2 DEFAULT dbms_session.unique_session_id ,
      P_JOB_ID         IN NUMBER   DEFAULT NULL ,
      P_IDENTIFICATION IN ERR_EX_PROT.IDENTIFICATION%TYPE DEFAULT $$PLSQL_UNIT ,
      P_TXT            IN ERR_EX_PROT.TEXT%TYPE ,
      P_PARAMSTRING1   IN VARCHAR2 := '(?)' ,
      P_PARAMSTRING2   IN VARCHAR2 := '(?)' ,
      P_PARAMSTRING3   IN VARCHAR2 := '(?)' ,
      P_PARAMSTRING4   IN VARCHAR2 := '(?)' ,
      P_PARAMSTRING5   IN VARCHAR2 := '(?)' ,
      P_PARAMSTRING6   IN VARCHAR2 := '(?)' ,
      P_PARAMSTRING7   IN VARCHAR2 := '(?)' ,
      P_PARAMSTRING8   IN VARCHAR2 := '(?)' ,
      P_PARAMSTRING9   IN VARCHAR2 := '(?)' )
  IS
    PRAGMA AUTONOMOUS_TRANSACTION;
    vIDENTIFICATION ERR_EX_PROT.IDENTIFICATION%TYPE;
    vTXT            ERR_EX_PROT.TEXT%TYPE;
    vPARAMSTRING    ERR_EX_PROT.TEXT%TYPE;
  BEGIN
    vIdentification := p_Identification;
    vTxt            := p_Txt;

    IF p_EventType   = EventT_SysError THEN
      vTxt          := SUBSTR( vTxt||CHR(10)|| '_____ Error Stack ____'||CHR(10)|| DBMS_UTILITY.FORMAT_ERROR_STACK()||CHR(10)|| DBMS_UTILITY.FORMAT_CALL_STACK(), 1, 4000 );
    END IF;

    IF NOT (P_PARAMSTRING1 = '(?)') THEN
       vPARAMSTRING := P_PARAMSTRING1;
    ELSIF NOT (P_PARAMSTRING2 = '(?)') THEN
       vPARAMSTRING := vPARAMSTRING || CHR(10) || P_PARAMSTRING2;
    ELSIF NOT (P_PARAMSTRING3 = '(?)') THEN
       vPARAMSTRING := vPARAMSTRING || CHR(10) || P_PARAMSTRING3;
    ELSIF NOT (P_PARAMSTRING4 = '(?)') THEN
       vPARAMSTRING := vPARAMSTRING || CHR(10) || P_PARAMSTRING4;
    ELSIF NOT (P_PARAMSTRING5 = '(?)') THEN
       vPARAMSTRING := vPARAMSTRING || CHR(10) || P_PARAMSTRING5;
    ELSIF NOT (P_PARAMSTRING6 = '(?)') THEN
       vPARAMSTRING := vPARAMSTRING || CHR(10) || P_PARAMSTRING6;
    ELSIF NOT (P_PARAMSTRING7 = '(?)') THEN
       vPARAMSTRING := vPARAMSTRING || CHR(10) || P_PARAMSTRING7;
    ELSIF NOT (P_PARAMSTRING8 = '(?)') THEN
       vPARAMSTRING := vPARAMSTRING || CHR(10) || P_PARAMSTRING8;
    ELSIF NOT (P_PARAMSTRING9 = '(?)') THEN
       vPARAMSTRING := vPARAMSTRING || CHR(10) || P_PARAMSTRING9;
    END IF;

    IF vPARAMSTRING IS NOT NULL THEN
      vTxt := vTxt || CHR(10) || vPARAMSTRING;
    END IF;

    dbms_transaction.begin_discrete_transaction;
    WHILE TRUE
    LOOP
      BEGIN
        INSERT
        INTO ERR_EX_PROT
          (
            id,
            ex_session,
            job_id,
            datum,
            dbuser,
            event_type,
            identification,
            text
          )
          VALUES
          (
            SEQ_ERR_ID.nextval,
            P_SESSION,
            P_JOB_ID,
            SYSDATE,
            USER,
            p_EventType,
            vIdentification,
            vTxt
          );
        COMMIT;
        EXIT;
      EXCEPTION
      WHEN dbms_transaction.discrete_transaction_failed THEN
        ROLLBACK;
      END;
    END LOOP;
  EXCEPTION
  WHEN OTHERS THEN
    vTxt := SQLERRM;
    INSERT
    INTO ERR_EX_PROT
      (
        id,
      ex_session,
        datum,
        dbuser,
        event_type,
        identification,
        text
      )
      VALUES
      (
        SEQ_ERR_ID.nextval,
        dbms_session.unique_session_id,
        SYSDATE,
        USER,
        EventT_SysError,
        'PKG_LIB_EX_BASIC_NEWAPI.writeLog',
        SUBSTR( p_EventType
        || CHR(10)
        || p_Identification
        || CHR(10)
        || p_Txt
        || CHR(10)
        || vTxt, --SQLERRM
        1, 4000 )
      );
    COMMIT;
  END;

 /*-----------------------------------------------------------------------------
Name:         splitSTL

Parameter:    p_list          String der gesplittet werden soll
              p_del           Zeichen an hand dem gesplittet werden soll

Return:       split_tbl_stl Tabelle mit den gesplitteten Teil-Strings

Beschreibung: Liefert eine Tabelle vom Typ split_tbl_stl mit den gesplitteten
              Teil-Strings. Default bei dem Zeichen an Hand dem gesplittet werden
              soll ist '.'
-----------------------------------------------------------------------------*/
  FUNCTION splitSTL
    (
      p_list    IN VARCHAR2,
      p_del     IN VARCHAR2 := '.',
      P_SESSION IN VARCHAR2 DEFAULT dbms_session.unique_session_id
    )
    RETURN split_tbl_stl pipelined
  IS
    l_idx pls_integer;
    l_list  VARCHAR2(32767) := p_list;
    l_value VARCHAR2(32767);
  BEGIN
    LOOP
      l_idx   := instr(l_list,p_del);
      IF l_idx > 0 THEN
        pipe row(SUBSTR(l_list,1,l_idx-1));
        l_list := SUBSTR(l_list,l_idx +LENGTH(p_del));
      ELSE
        pipe row(l_list);
        EXIT;
      END IF;
    END LOOP;
    RETURN;
  EXCEPTION
  WHEN OTHERS THEN
    PKG_LIB_EX_BASIC_NEWAPI.WRITELOG(P_EVENTTYPE      => PKG_LIB_EX_BASIC_NEWAPI.EventT_SysError,
                              P_SESSION        => P_SESSION,
                              P_IDENTIFICATION => $$PLSQL_UNIT,
                              P_TXT            => 'Systemfehler');
  END splitSTL;

/*------------------------------------------------------------------------------
Name:         getPath

Parameter: pDir     Name eines Directory-Objekts

Beschreibung: Prozedur den Pfad  eines Directory-Objekts zurück. Wenn es das
              Directory-Objekt nicht gibt wird NULL zurückgegeben.
------------------------------------------------------------------------------*/
  FUNCTION getPath(
      pDir      IN VARCHAR2,
      P_SESSION IN VARCHAR2 DEFAULT dbms_session.unique_session_id )
    RETURN VARCHAR2
  AS
  vPath VARCHAR2(2000);
  BEGIN
     SELECT directory_path into vPath
       FROM all_directories where directory_name = pDir;
    RETURN vPath ;
  EXCEPTION
  WHEN no_data_found THEN
    RETURN NULL;
  WHEN OTHERS THEN
    PKG_LIB_EX_BASIC_NEWAPI.WRITELOG(P_EVENTTYPE      => PKG_LIB_EX_BASIC_NEWAPI.EventT_SysError,
                              P_SESSION        => P_SESSION,
                              P_IDENTIFICATION => $$PLSQL_UNIT,
                              P_TXT            => 'Systemfehler');
  END getPath;



/*------------------------------------------------------------------------------
Name:         checkUnixAPI

Parameter: pCommand     Unix Komando

Beschreibung: Die Prozedur nimmt das Kommando als Parameter entgegen, führt es
              aus und gibt die ersten 32 Kb der Ausgabe an den Aufrufer zurück.
------------------------------------------------------------------------------*/
  PROCEDURE checkUnixAPI(P_SESSION IN VARCHAR2 DEFAULT dbms_session.unique_session_id ) IS
  vResult        VARCHAR2(32000);
  BEGIN
--    vUnixErr := PKG_LIB_EX_BASIC_NEWAPI.CALLUNIX(pCommand);
    vResult := replace('|'||PUT_SCHEDULER_UTILS.PKG_UNIX_API.fnc_call_unix('ls $DWH_BASE/work|grep DM_EXPORT'),chr(10),'|');

    IF (instr(vResult,'DM_EXPORT')=0) THEN
      Raise_Application_Error (-20343, 'UNIX-API nicht verfügbar!');
    END IF;
  EXCEPTION
  WHEN OTHERS THEN
    PKG_LIB_EX_BASIC_NEWAPI.WRITELOG(P_EVENTTYPE      => PKG_LIB_EX_BASIC_NEWAPI.EventT_SysError,
                              P_SESSION        => P_SESSION,
                              P_IDENTIFICATION => $$PLSQL_UNIT,
                              P_TXT            => 'Unix-API nicht verfügbar!');
    RAISE;
  END checkUnixAPI;




/*------------------------------------------------------------------------------
Name:         callUnix

Parameter: pCommand     Unix Komando

Beschreibung: Die Funktion nimmt das Kommando als Parameter entgegen, führt es
              aus und gibt die ersten 32 Kb der Ausgabe an den Aufrufer zurück.
------------------------------------------------------------------------------*/
--  FUNCTION callUnix( pCommand     IN VARCHAR2 )
--    RETURN VARCHAR2
--   IS language java name 'ExExternalCall.call_unix(java.lang.String) return java.lang.String';

/*------------------------------------------------------------------------------
Name:         callUnix

Parameter: pCommand     Unix Komando

Beschreibung: Die Prozedur nimmt das Kommando als Parameter entgegen, führt es
              aus und gibt die ersten 32 Kb der Ausgabe an den Aufrufer zurück.
------------------------------------------------------------------------------*/
  PROCEDURE callUnix( pCommand  IN VARCHAR2,
                      P_SESSION IN VARCHAR2 DEFAULT dbms_session.unique_session_id ) IS
  vUnixErr        VARCHAR2(32000);
  BEGIN
--    vUnixErr := PKG_LIB_EX_BASIC_NEWAPI.CALLUNIX(pCommand);
    vUnixErr := PUT_SCHEDULER_UTILS.PKG_UNIX_API.fnc_call_unix(pCommand);

    IF (vUnixErr IS NOT NULL) THEN
      Raise_Application_Error (-20343, 'UNIX-AUFRUF-FEHLER ' || vUnixErr);
    END IF;
  EXCEPTION
  WHEN OTHERS THEN
    PKG_LIB_EX_BASIC_NEWAPI.WRITELOG(P_EVENTTYPE      => PKG_LIB_EX_BASIC_NEWAPI.EventT_SysError,
                              P_SESSION        => P_SESSION,
                              P_IDENTIFICATION => $$PLSQL_UNIT,
                              P_TXT            => 'Systemfehler');
  END callUnix;

/*------------------------------------------------------------------------------
Name:         callUnixAsync

Parameter: pCommand     Unix Komando

Beschreibung: Die Funktion nimmt das Kommando als Parameter entgegen, führt es
              aus und gibt die ersten 32 Kb der Ausgabe an den Aufrufer zurück.
              Es wird nicht gewartet bis der Aufruf abgeschlossen ist.
------------------------------------------------------------------------------*/
--  FUNCTION callUnixAsync( pCommand     IN VARCHAR2 )
--    RETURN VARCHAR2
--   IS language java name 'ExExternalCallAsync.call_unix(java.lang.String) return java.lang.String';

/*------------------------------------------------------------------------------
Name:         callUnixAsync

Parameter: pCommand     Unix Komando

Beschreibung: Die Prozedur nimmt das Kommando als Parameter entgegen, führt es
              aus und gibt die ersten 32 Kb der Ausgabe an den Aufrufer zurück.
              Es wird nicht gewartet bis der Aufruf abgeschlossen ist
------------------------------------------------------------------------------*/
  PROCEDURE callUnixAsync( pCommand  IN VARCHAR2,
                           P_SESSION IN VARCHAR2 DEFAULT dbms_session.unique_session_id ) IS
  vUnixErr        VARCHAR2(32000) := NULL;
  BEGIN
--    vUnixErr := PKG_LIB_EX_BASIC_NEWAPI.callUnixAsync(pCommand);
    PUT_SCHEDULER_UTILS.PKG_UNIX_API.prc_call_unix_async(pCommand);
    IF (vUnixErr IS NOT NULL) THEN
      Raise_Application_Error (-20343, 'UNIX-ASYNCHRON-AUFRUF-FEHLER ' || vUnixErr);
    END IF;
  EXCEPTION
  WHEN OTHERS THEN
    PKG_LIB_EX_BASIC_NEWAPI.WRITELOG(P_EVENTTYPE      => PKG_LIB_EX_BASIC_NEWAPI.EventT_SysError,
                              P_SESSION        => P_SESSION,
                              P_IDENTIFICATION => $$PLSQL_UNIT,
                              P_TXT            => 'Systemfehler');
  END callUnixAsync;

/*------------------------------------------------------------------------------
Name:         prc_PreCreateFiles

Parameter: pSession        Session-Id des laufenden DM-Exports
           pTempDir        Name des temporären Directory-Objects
           pMode           Default-File-Rechte

Beschreibung: Legt einen Vorrat von leeren Dateien mit nur einem Scheduler-
              Aufruf an.
------------------------------------------------------------------------------*/
  PROCEDURE prc_PreCreateFiles(pSession IN NUMBER,pTempDir in varchar2,pMode in varchar2)
  IS
  vCmd VARCHAR2(4000);
  vResult VARCHAR2(32762);
  BEGIN
    -- Anzahl angelegter Dateien um gFPCMaxFiles erhöhen
    gFPCMaxId:=gFPCMaxId+gFPCMaxFiles;
    vCmd:='( cd '||getPath(pTempDir,pSession)||chr(10);
    for i in gFPCCurrentId+1 .. gFPCMaxId
    loop
 --       dbms_output.put_line(i);
        vCmd := vCmd || 'touch '||replace(replace(gFPCFileTemplate,'<SESSION>',pSession),'<ID>',i)|| chr(10);
    end loop;
    vCmd := vCmd || 'chmod ' || pMODE ||' '||replace(replace(gFPCFileTemplate,'<SESSION>',pSession),'<ID>','*');
    vCmd := vCmd || chr(10)||')';
 --   dbms_output.put_line(vCmd);
 --   dbms_output.put_line(length(vCmd));
    callUnix(vCmd,pSession);

  END;

/*------------------------------------------------------------------------------
Name:         prc_RemovePreCreateFiles

Parameter: pSession        Session-Id des laufenden DM-Exports
           pTempDir        Name des temporären Directory-Objects


Beschreibung: Die Prozedur räumt zu viel angelete Dateien aus dem Pool
              wieder weg.
------------------------------------------------------------------------------*/
PROCEDURE prc_RemovePreCreateFiles(pSession IN NUMBER,pTempDir in varchar2)
  IS
  vCmd VARCHAR2(4000);
  vResult VARCHAR2(32762);
  BEGIN
    -- angelegter Dateien wieder löschen
    vCmd:='rm '||getPath(pTempDir,pSession)||'/'||replace(replace(gFPCFileTemplate,'<SESSION>',pSession),'<ID>','*');
    callUnix(vCmd,pSession);
  END;



/*------------------------------------------------------------------------------
Name:         fnc_getNextFile

Parameter: pSession        Session-Id des laufenden DM-Exports
           pTempDir        Name des temporären Directory-Objects
           pMode           Standard Filesystemrechte

Return:    Filename

Beschreibung: Diese Funktion gibt die nächste Datei aus dem Dateipool zurück.
              Falls es keine Datei mehr gibt, so wird der Pool vergrößert.
------------------------------------------------------------------------------*/
  FUNCTION fnc_getNextFile(pSession IN NUMBER,pTempDir in varchar2,pMode in varchar2) return varchar2
  IS
  vFileName VARCHAR2(100);
  BEGIN

    if gFPCCurrentId=0 or gFPCCurrentId=gFPCMaxId
    then
        prc_PreCreateFiles(pSession,pTempDir,pMode);
    end if;
    gFPCCurrentId := gFPCCurrentId+1;

    return replace(replace(gFPCFileTemplate,'<SESSION>',pSession),'<ID>',gFPCCurrentId);
  END;

/*------------------------------------------------------------------------------
Name:         fopen

Parameter: pSession        Session-Id des laufenden DM-Exports
           pTempDir        Name des temporären Directory-Objects
           pLocation       Ziel-Directory
           pFileName       Name der zu öffenden Datei

Return:    Filehandle

Beschreibung: Die Funktion benennt eine Datei aus dem im Vorfeld angelegten
            Datei-Pool um und verschiebt sie in das Ziel-Verzeichnis.
            Falls keine vorangelegten Dateien mehr vorhanden sind, so werden
            wieder welche angelegt.
------------------------------------------------------------------------------*/
FUNCTION fopen(pSession IN NUMBER,pTempDir in varchar2,pMode in varchar2,pLocation in varchar2,pFileName in varchar2) return UTL_FILE.FILE_TYPE
AS
lFile      UTL_FILE.FILE_TYPE;
lPath       VARCHAR2(4000);
lUnixCom    VARCHAR2(4000);
lFileName   VARCHAR2(4000);
BEGIN

  -- Neues vorher angelegtes File holen
  lFileName := fnc_getNextFile(pSession,pTempDir,pMode);
--   PKG_LIB_EX_BASIC_NEWAPI.WRITELOG(P_EVENTTYPE      => PKG_LIB_EX_BASIC_NEWAPI.EventT_Information,
--                              P_SESSION        => pSession,
--                              P_IDENTIFICATION => $$PLSQL_UNIT,
--                              P_TXT            => lFileName||':'||pFileName);

  -- File umbenennen
  UTL_FILE.FRENAME(pTempDir,lFileName, pLocation,pFileName, TRUE);

  -- File öffnen
  lFile := UTL_FILE.FOPEN(
            location => pLocation,
            filename => pFileName,
            open_mode => 'a',
            max_linesize => 32767);
  return lFile;
END;


  END PKG_LIB_EX_BASIC_NEWAPI;
/
