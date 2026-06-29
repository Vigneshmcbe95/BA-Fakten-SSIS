CREATE OR REPLACE PACKAGE                          "PKG_EX_LOAD_STL" AS
-- $Id: create_PKG_EX_LOAD_STL.sql 1395 2023-03-31 11:20:58Z seideld006 $
-- -----------------------------------------------------------------------------
--@START DOKU
-- -----------------------------------------------------------------------------
-- MODUL
-- Revision     $LastChangedRevision: 1395 $
-- Last Revised $LastChangedDate: 2023-03-31 13:20:58 +0200 (Fr, 31 Mrz 2023) $
-- Author       $LastChangedBy: seideld006 $
-- -----------------------------------------------------------------------------
-- DATEI
--  File         $HeadURL: https://svn.sdst.sbaintern.de/put/trunk/Anwendungen/DMExport/src/stat_exp_datamart_mcfg/sql/packages/create_PKG_EX_LOAD_STL.sql $
--
-- SYNTAX
--
-- BENOETIGT
-- Tabelle WRK_EX_OBJ
--         TMP_EX_DIR_LIST
--
--
-- BESCHREIBUNG
--  Liest die Steuerlisten für die Exportschnitstelle und bereitet die gelesenen
--  Daten in der Tabelle WRK_EX_OBJ auf.
--
--  04.03.2015 CebucL - Initiale Version
--  15.06.2015 CebucL - Kleinere Erweiterungen
--  05.08.2020 SeidelD006 - Abgleich mit ESTATX
--  19.01.2021 SeidelD006 - Referenzen auf PKG_EX_WORK auf PKG_EX_WORK_NEWAPI geändert
--  19.10.2021 SeidelD006 - Logging-Informationen für eine bessere Fehleranalyse hinzugefügt (BIDW-439)
--  31.03.2023 SeidelD006 - Verarbeitung von IAB-Steuerlisten deaktiviert (BIDW-471)
--
--@ENDE DOKU
-- -----------------------------------------------------------------------------

/*------------------------------------------------------------------------------
Name:         getSTL

Parameter:    p_EX_SESSION     Nummer zu welchem Lauf die Steuerlisten gehören.
              p_STL_NAME       Name der zu ladenden Steuerliste
              p_REF_DATE       Referenzdatum für MON_ID-Berechnungen

Beschreibung: Prozedur liest die Inhalte einer Steuerlisten und legt sie in
              der Tabelle WRK_EX_OBJ.
------------------------------------------------------------------------------*/
  PROCEDURE getSTL(p_EX_SESSION IN NUMBER, p_STL_NAME VARCHAR2 ,p_Ref_Date IN DATE);


/*------------------------------------------------------------------------------
Name:         getSTLIAB

Parameter:    p_EX_SESSION     Nummer zu welchem Lauf die Steuerlisten gehören.
              p_STL_NAME       Name der zu ladenden Steuerliste
              p_REF_DATE       Referenzdatum für MON_ID-Berechnungen

Beschreibung: Prozedur liest die Inhalte einer Steuerlisten und legt sie in
              der Tabelle WRK_EX_OBJ.
------------------------------------------------------------------------------*/
  PROCEDURE getSTLIAB(p_EX_SESSION IN NUMBER, p_STL_NAME VARCHAR2, p_Ref_Date IN DATE);


/*------------------------------------------------------------------------------
Name:         setSTLDone

Parameter:    p_EX_SESSION     Nummer zu welchem Lauf die Steuerlisten gehören.
              p_STL_NAME       Name der zu ladenden Steuerliste

Beschreibung: Prozedur verschiebt Steuerliste ins ToDo Verzeichis wenn soe
              fehlerfrei exportiert wurde.
------------------------------------------------------------------------------*/
  PROCEDURE setSTLDone(p_EX_SESSION IN NUMBER, p_STL_NAME VARCHAR2 );

/*------------------------------------------------------------------------------
Name:         setReload

Parameter:    p_EX_SESSION     Nummer zu welchem Lauf die Steuerlisten gehoeren.
              p_STL_NAME       Name der zu ladenden Steuerliste

Beschreibung: Objekte einer Steuerliste, die im letzten Lauf fehlerfrei
              abgearbeitet wurden, werden nicht noch einmal exportiert.
              Sie erhalten den Status  "Abgeschlossen".
------------------------------------------------------------------------------*/
  PROCEDURE setReload(p_EX_SESSION IN NUMBER, p_STL_NAME VARCHAR2,p_Ref_Date in date );

END PKG_EX_LOAD_STL;
/


CREATE OR REPLACE PACKAGE BODY                          "PKG_EX_LOAD_STL"
IS

-- Verzeichnis in welchem die Steuerlisten liegen
vDirTodo  VARCHAR2(30) :='EX_STL_TODO';

-- Verzeichnis in welchem die IAB-Steuerlisten liegen
vDirIAB  VARCHAR2(30) :='EX_STL_IAB';

-- Verzeichnis in welches die erledigten Steuerlisten verschoben werden
vDirDone  VARCHAR2(30) :='EX_STL_DONE';

vVerbose boolean := true;
/*------------------------------------------------------------------------------
Name:         getPath

Parameter:    Directory-Name

Beschreibung: Gibt den Pfad des übergebenen Directories zurück
------------------------------------------------------------------------------*/
function getPath(pDirName in varchar2) return varchar2
is
 lPath varchar2(4000);
begin
    select directory_path
    into lPath
    from all_directories
    where upper(directory_name)=upper(pDirName);

    return lPath;
exception when others then return null;
end;




/*------------------------------------------------------------------------------
Name:         logInfo

Parameter:    beliebiger Text

Beschreibung: Der übergebene Text wird noch ergänzt und es wird ein Log-Eintrag
              daraus erzeugt..
------------------------------------------------------------------------------*/
procedure logInfo (pText VARCHAR2,p_EX_SESSION IN NUMBER) IS
BEGIN
  pkg_lib_ex_basic_newapi.WRITELOG(P_EVENTTYPE      => pkg_lib_ex_basic_newapi.EventT_Information,
                            P_SESSION        => p_EX_SESSION,
                            P_IDENTIFICATION => $$PLSQL_UNIT,
                            P_TXT            => pText);
EXCEPTION
  WHEN OTHERS THEN
    pkg_lib_ex_basic_newapi.WRITELOG(P_EVENTTYPE      => pkg_lib_ex_basic_newapi.EventT_SysError,
                              P_SESSION        => p_EX_SESSION,
                              P_IDENTIFICATION => $$PLSQL_UNIT,
                              P_TXT            => 'Systemfehler');
END logInfo;
/*------------------------------------------------------------------------------
Name:         logVerbose

Parameter:    beliebiger Text

Beschreibung: Der übergebene Text wird noch ergänzt und es wird ein Log-Eintrag
              daraus erzeugt..
------------------------------------------------------------------------------*/
procedure logVerbose (pText VARCHAR2,p_EX_SESSION IN NUMBER) IS
BEGIN

    if vVerbose then
        pkg_lib_ex_basic_newapi.WRITELOG(P_EVENTTYPE      => pkg_lib_ex_basic_newapi.EventT_Information,
                            P_SESSION        => p_EX_SESSION,
                            P_IDENTIFICATION => $$PLSQL_UNIT,
                            P_TXT            => pText);
    end if;

EXCEPTION
  WHEN OTHERS THEN
    pkg_lib_ex_basic_newapi.WRITELOG(P_EVENTTYPE      => pkg_lib_ex_basic_newapi.EventT_SysError,
                              P_SESSION        => p_EX_SESSION,
                              P_IDENTIFICATION => $$PLSQL_UNIT,
                              P_TXT            => 'Systemfehler');
END logVerbose;



/*------------------------------------------------------------------------------
Name:      getPosStern

Parameter:    pLine      Name des Directory-Objekts in welchem die STL sind.

Return:       Postition des gefundenen Musters '*'

Beschreibung: Function, prueft ob der Eingabestring einem gewissen Muster
              entspricht
------------------------------------------------------------------------------*/
  FUNCTION getPosStern(
      pLine        IN VARCHAR2,
      p_EX_SESSION IN NUMBER)
    RETURN INTEGER
  AS
    vPos         INTEGER;
    vRegStern    VARCHAR2(60) := '*';   -- z.B *
  BEGIN

    -- finden '*'
    SELECT INSTR(pLine, vRegStern)
    INTO vPos
    FROM DUAL;

    RETURN vPos;
  EXCEPTION
  WHEN no_data_found THEN
     RETURN vPos;
  WHEN OTHERS THEN
    PKG_LIB_EX_BASIC_NEWAPI.WRITELOG(P_EVENTTYPE      => PKG_LIB_EX_BASIC_NEWAPI.EventT_SysError,
                              P_SESSION        => p_EX_SESSION,
                              P_IDENTIFICATION => $$PLSQL_UNIT,
                              P_TXT            => 'Systemfehler');
  END getPosStern;


/*------------------------------------------------------------------------------
Name:         getPosSeparator

Parameter:    pLine      Name des Directory-Objekts in welchem die STL sind.

Return:       Postition des gefundenen Trenners ': oder .'

Beschreibung: Function, prueft ob der Eingabestring einem gewissen Muster
              entspricht
------------------------------------------------------------------------------*/
  FUNCTION getPosSeparator(
      pLine        IN VARCHAR2,
      p_EX_SESSION IN NUMBER)
    RETURN INTEGER
  AS
    vPos         INTEGER;
    vRegDPunkt   VARCHAR2(60) := ':';
    vRegPunkt    VARCHAR2(60) := '.';
  BEGIN

    -- finden ':'
    SELECT INSTR(pLine, vRegDPunkt)
    INTO vPos
    FROM DUAL;

    IF vPos = 0 THEN
      SELECT INSTR(pLine, vRegPunkt)
      INTO vPos
      FROM DUAL;
    END IF;

    RETURN vPos;
  EXCEPTION
  WHEN no_data_found THEN
     RETURN vPos;
  WHEN OTHERS THEN
    PKG_LIB_EX_BASIC_NEWAPI.WRITELOG(P_EVENTTYPE      => PKG_LIB_EX_BASIC_NEWAPI.EventT_SysError,
                              P_SESSION        => p_EX_SESSION,
                              P_IDENTIFICATION => $$PLSQL_UNIT,
                              P_TXT            => 'Systemfehler');
  END getPosSeparator;


  /*------------------------------------------------------------------------------
Name:      getPosStrich

Parameter:    pLine      Name des Directory-Objekts in welchem die STL sind.

Return:       Postition des gefundenen Musters '_'

Beschreibung: Function, prueft ob der Eingabestring einem gewissen Muster
              entspricht
------------------------------------------------------------------------------*/
  FUNCTION getPosStrich(
      pLine        IN VARCHAR2,
      p_EX_SESSION IN NUMBER)
    RETURN INTEGER
  AS
    vPos         INTEGER;
    vRegStern    VARCHAR2(60) := '_';   -- z.B *
  BEGIN

    -- finden '*'
    SELECT INSTR(pLine, vRegStern)
    INTO vPos
    FROM DUAL;

    RETURN vPos;
  EXCEPTION
  WHEN no_data_found THEN
     RETURN vPos;
  WHEN OTHERS THEN
    PKG_LIB_EX_BASIC_NEWAPI.WRITELOG(P_EVENTTYPE      => PKG_LIB_EX_BASIC_NEWAPI.EventT_SysError,
                              P_SESSION        => p_EX_SESSION,
                              P_IDENTIFICATION => $$PLSQL_UNIT,
                              P_TXT            => 'Systemfehler');
  END getPosStrich;

/*------------------------------------------------------------------------------
Name:      exTrim

Parameter:    pLine      Zeile die getrimmt werden soll

Return:       Getrimmte Zeile

Beschreibung: Function entfernt vom Anfang und Ende einer Zeile diverse
              Sonderzeichen z.B. HT, CR, SP oder ;
------------------------------------------------------------------------------*/
  FUNCTION exTrim(pLine        IN VARCHAR2,
                  p_EX_SESSION IN NUMBER)
    RETURN VARCHAR2
  AS
    vLine VARCHAR2(2000) := pLine;
  BEGIN
    SELECT TRIM(CHR(9)  FROM vLine) into vLine FROM dual;  -- HT
    SELECT TRIM(CHR(13) FROM vLine) into vLine FROM dual;  -- CR
    SELECT TRIM(CHR(32) FROM vLine) into vLine FROM dual;  -- SP
    SELECT TRIM(CHR(59) FROM vLine) into vLine FROM dual;  -- ;
    RETURN vLine;
  EXCEPTION
  WHEN OTHERS THEN
    PKG_LIB_EX_BASIC_NEWAPI.WRITELOG(P_EVENTTYPE      => PKG_LIB_EX_BASIC_NEWAPI.EventT_SysError,
                              P_SESSION        => p_EX_SESSION,
                              P_IDENTIFICATION => $$PLSQL_UNIT,
                              P_TXT            => 'Systemfehler');
  END exTrim;

/*------------------------------------------------------------------------------
Name:         getSTL

Parameter:    p_EX_SESSION     Nummer zu welchem Lauf die Steuerlisten gehoeren.
              p_STL_NAME       Name der zu ladenden Steuerliste

Beschreibung: Prozedur liest die Inhalte einer Steuerlisten und legt sie in
              der Tabelle WRK_EX_OBJ ab.
------------------------------------------------------------------------------*/
  PROCEDURE getSTL(p_EX_SESSION IN NUMBER, p_STL_NAME VARCHAR2, p_Ref_Date IN DATE )
  IS
    i          INTEGER;
    vLine      VARCHAR2(2000);
    vPath      VARCHAR2(2000);
    vTest      VARCHAR2(30);
    vFile      UTL_FILE.FILE_TYPE;
    vPerm      INTEGER;
    vPos       INTEGER;
    vPosStern  INTEGER;
    vPosStrich INTEGER;
    stl_rec    WRK_EX_OBJ%ROWTYPE;
    vOBJ_NAME  STAT_EXP_DATAMART_MCFG.WRK_EX_OBJ.STL_OBJ_NAME%TYPE;
  BEGIN

    logInfo('Verarbeite Steuerliste >'||p_STL_NAME||'<',p_EX_SESSION);
      i       := 1;
      stl_rec := NULL;
      -- Infos aus Dateinamen befuellen (Zerlegung des Dateinamens)
      FOR cParam IN
      (SELECT * FROM TABLE(STAT_EXP_DATAMART_MCFG.PKG_LIB_EX_BASIC_NEWAPI.splitSTL( p_STL_NAME))
      )
      LOOP
        IF i                 = 1 THEN
           IF cParam.COLUMN_VALUE IN ( 'b', 'f', 'd', 'v') THEN
              stl_rec.STL_KZ     := cParam.COLUMN_VALUE;
           ELSE
              stl_rec.STL_KZ     := '0';
           END IF;
        ELSIF i              = 2 THEN
          stl_rec.STL_SYSTEM := cParam.COLUMN_VALUE;
        ELSIF i              = 3 THEN
          stl_rec.STL_SCHEMA := cParam.COLUMN_VALUE;
        ELSE
          NULL;
        END IF;
        i := i + 1;
      END LOOP; -- Infos aus Dateinamen befuellen (Zerlegung des Dateinamens)

      logInfo('STL_KZ: >'|| stl_rec.STL_KZ||'< STL_SYSTEM: >'||stl_rec.STL_SYSTEM ||'< STL_SCHEMA: >'||stl_rec.STL_SCHEMA ||'<',p_EX_SESSION);

      -- Datenbank pruefen (EXCEPTION no_data_found wenn nicht vorhanden)
      --SELECT NAME INTO vTest FROM V$DATABASE WHERE upper(NAME) like upper(stl_rec.STL_SYSTEM) || '%';
      SELECT NAME INTO vTest from (select UPPER(sys_context('USERENV','DB_NAME')) as NAME from DUAL) where name like '%STAT%';
      -- Schema pruefen (EXCEPTION no_data_found wenn nicht vorhanden)
      SELECT USERNAME INTO vTest FROM all_users WHERE upper(USERNAME) = upper(stl_rec.STL_SCHEMA);

      vPerm := NULL;
      SELECT INSTR(p_STL_NAME, '_perm', 1, 1) INTO vPerm FROM DUAL;
      IF vPerm          != 0 THEN
        stl_rec.STL_PERM := 1;
      ELSE
        stl_rec.STL_PERM := 0;
      END IF;
      stl_rec.STL_FILE_NAME := p_STL_NAME;
      stl_rec.STL_SESSION   := p_EX_SESSION;
      stl_rec.STL_ZBDAT     := sysdate;
      logInfo('STL_PERM: >'||stl_rec.stl_perm||'< vDirTodo: >'||getPath(vDirTodo)||'<',p_EX_SESSION);
      logVerbose('Open file',p_EX_SESSION);

      vFile                 := UTL_FILE.FOPEN(vDirTodo, p_STL_NAME, 'r');
      -- Schleife inerhalb der einzelne Datei. GET_LINE bringt NO_DATA_FOUND
      -- wenn die Datei am Ende ist. das wird benutzt fuer Schleifenende
      LOOP
        BEGIN

          vLine                   := NULL;
          UTL_FILE.GET_LINE(vFile, vLine);
          logVerbose('Line: >'||vLine||'<',p_EX_SESSION);
          vLine                   := exTrim(vLine, p_EX_SESSION);

          -- Leer-Zeilen ignorieren also zum ende der Schleife
          IF vLine IS NULL THEN
             GOTO end_loop;
          END IF;
          -- Auskomentierte-Zeilen ignorieren also zum ende der Schleife
          IF vLine LIKE '#%' THEN
             GOTO end_loop;
          END IF;

          stl_rec.STL_ID          := SEQ_STL_IDS.NEXTVAL;
          stl_rec.STL_OBJ_NAME    := LOWER(vLine);
          stl_rec.STL_LINE        := vLine;
          stl_rec.STL_REF_DATE    := p_Ref_Date;
          stl_rec.STL_STATUS      := 'L';  -- geladen

          logVerbose('STL_ID: >'||stl_rec.STL_ID ||'< STL_OBJ_NAME: >'||stl_rec.STL_OBJ_NAME ||'< STL_LINE: >'||stl_rec.STL_LINE ||'< STL_STATUS: >'||stl_rec.STL_STATUS ||'<'||' STL_REF_DATE: >'||to_char(stl_rec.stl_ref_date,'DD.MM.YYYY HH24:MI:SS') ||'<',p_EX_SESSION);
          -- insert
          INSERT INTO WRK_EX_OBJ VALUES stl_rec;

          <<end_loop>>
          NULL;
        EXCEPTION
        WHEN no_data_found THEN
          logVerbose('Close file',p_EX_SESSION);
          UTL_FILE.FCLOSE(vFile);
          EXIT;
        END;
      END LOOP; -- Schleife innerhalb der einzelne Datei.
    COMMIT;
  EXCEPTION
  WHEN no_data_found THEN
    PKG_LIB_EX_BASIC_NEWAPI.WRITELOG(P_EVENTTYPE      => PKG_LIB_EX_BASIC_NEWAPI.EventT_FuncError,
                              P_SESSION        => p_EX_SESSION,
                              P_IDENTIFICATION => $$PLSQL_UNIT,
                              P_TXT            => 'Fehler, Steuerliste' || p_STL_NAME || ' nicht valide, System oder Datenbankschema nicht gefunden');
  WHEN UTL_FILE.INVALID_PATH THEN
    PKG_LIB_EX_BASIC_NEWAPI.WRITELOG(P_EVENTTYPE      => PKG_LIB_EX_BASIC_NEWAPI.EventT_FuncError,
                              P_SESSION        => p_EX_SESSION,
                              P_IDENTIFICATION => $$PLSQL_UNIT,
                              P_TXT            => 'Fehler, Pfad fuer Steuerliste ' || p_STL_NAME ||' ist nicht valide');
  WHEN UTL_FILE.INVALID_OPERATION THEN
    PKG_LIB_EX_BASIC_NEWAPI.WRITELOG(P_EVENTTYPE      => PKG_LIB_EX_BASIC_NEWAPI.EventT_FuncError,
                              P_SESSION        => p_EX_SESSION,
                              P_IDENTIFICATION => $$PLSQL_UNIT,
                              P_TXT            => 'Fehler, Steuerliste ' || p_STL_NAME || ' kann nicht gelesen werden');
  WHEN OTHERS THEN
    UTL_FILE.FCLOSE(vFile);
    PKG_LIB_EX_BASIC_NEWAPI.WRITELOG(P_EVENTTYPE      => PKG_LIB_EX_BASIC_NEWAPI.EventT_SysError,
                              P_SESSION        => p_EX_SESSION,
                              P_IDENTIFICATION => $$PLSQL_UNIT,
                              P_TXT            => 'Systemfehler');
  END getSTL;


/*------------------------------------------------------------------------------
Name:         getSTLIAB

Parameter:    p_EX_SESSION     Nummer zu welchem Lauf die Steuerlisten gehoeren.
              p_STL_NAME       Name der zu ladenden Steuerliste

Beschreibung: Prozedur liest die Inhalte einer Steuerlisten und legt sie in
              der Tabelle WRK_EX_OBJ_IAB ab.
------------------------------------------------------------------------------*/
  PROCEDURE getSTLIAB(p_EX_SESSION IN NUMBER, p_STL_NAME VARCHAR2, p_REF_DATE IN DATE )
  IS

    vLine            VARCHAR2(2000);
    vPosDoppelP      INTEGER;
    vTO_BE_SAVED     BOOLEAN := FALSE;
    vNOT_TO_BE_SAVED BOOLEAN := FALSE;
    vFile            UTL_FILE.FILE_TYPE;
    iab_rec          WRK_EX_OBJ%ROWTYPE;
    vOBJ_NAME        STAT_EXP_DATAMART_MCFG.WRK_EX_OBJ.STL_OBJ_NAME%TYPE;
    vSystem          STAT_EXP_DATAMART_MCFG.WRK_EX_OBJ.STL_SYSTEM%TYPE;
  BEGIN

    logInfo('Verarbeite Steuerliste >'||p_STL_NAME||'<',p_EX_SESSION);

    RAISE_APPLICATION_ERROR(-20001,'IAB-Steuerlisten wurden durch die IAB-Treuhänderlösung abgelöst. Verarbeitung von IAB-Steuerlisten wird nicht mehr ausgeführt!');


      iab_rec.STL_FILE_NAME := p_STL_NAME;
      iab_rec.STL_SESSION   := p_EX_SESSION;
      iab_rec.STL_ZBDAT     := sysdate;
      logInfo('vDirIAB: >'||getPath(vDirIAB)||'<',p_EX_SESSION);
      logVerbose('Open file',p_EX_SESSION);

      vFile                 := UTL_FILE.FOPEN(vDirIAB, p_STL_NAME, 'r');
      -- Schleife inerhalb der einzelne Datei. GET_LINE bringt NO_DATA_FOUND
      -- wenn die Datei am Ende ist. das wird benutzt fuer Schleifenende
      LOOP
        BEGIN
          iab_rec.STL_OBJ_NAME    := NULL;
          vLine                   := NULL;
          UTL_FILE.GET_LINE(vFile, vLine);
          logVerbose('Line: >'||vLine||'<',p_EX_SESSION);
          vLine                   := exTrim(vLine, p_EX_SESSION);

          -- Leer-Zeilen ignorieren also zum ende der Schleife
          IF vLine IS NULL THEN
             GOTO end_loop;
          END IF;
          -- Auskomentierte-Zeilen ignorieren also zum ende der Schleife
          IF vLine LIKE '#%' THEN
             GOTO end_loop;
          END IF;

          -- echte Auftraege in der Steuerliste
          IF vLine = 'START_SECTION_TABLES_TO_BE_SAVED' THEN
          logVerbose('Start zu exportierende Objekte',p_EX_SESSION);
             vTO_BE_SAVED := TRUE;
             GOTO end_loop;
          END IF;

          -- echte Auftraege in der Steuerliste ENDE
          IF vLine = 'END_SECTION_TABLES_TO_BE_SAVED' THEN
             vTO_BE_SAVED := FALSE;
          logVerbose('Ende zu exportierende Objekte',p_EX_SESSION);
             GOTO end_loop;
          END IF;

          -- echte Auftraege in der Steuerliste fuer Aussschluss
          IF vLine = 'START_SECTION_TABLES_NOT_TO_BE_SAVED' THEN
             vNOT_TO_BE_SAVED := TRUE;
            logVerbose('Start nicht zu exportierende Objekte',p_EX_SESSION);
             GOTO end_loop;
          END IF;

          -- echte Auftraege in der Steuerliste fuer Aussschluss ENDE
          IF vLine = 'END_SECTION_TABLES_NOT_TO_BE_SAVED' THEN
             vNOT_TO_BE_SAVED := FALSE;
            logVerbose('Ende nicht zu exportierende Objekte',p_EX_SESSION);
             GOTO end_loop;
          END IF;

          IF vTO_BE_SAVED OR vNOT_TO_BE_SAVED THEN
            -- normaler Satz
            IF vTO_BE_SAVED THEN
              iab_rec.STL_STATUS      := 'L';  -- geladen
            END IF;

            -- Ausschluss-Eintrag
            IF vNOT_TO_BE_SAVED THEN
              iab_rec.STL_STATUS      := 'E';  -- aussgeschlosen
            END IF;

            vPosDoppelP             := getPosSeparator(vLine, p_EX_SESSION);

            SELECT
              CASE
                WHEN upper(NAME) LIKE 'ESTAT%'
                THEN 'ESTAT'
                WHEN upper(NAME) LIKE 'FSTAT%'
                THEN 'FSTAT'
                WHEN upper(NAME) LIKE 'ISTAT%'
                THEN 'ISTAT'
                WHEN upper(NAME) LIKE 'PSTAT%'
                THEN 'PSTAT'
                WHEN upper(NAME) LIKE 'LSTAT%'
                THEN 'LSTAT'
                WHEN upper(NAME) LIKE 'ASTAT%'
                THEN 'ASTAT'
                ELSE NULL
              END
            INTO vSystem
  FROM (select UPPER(sys_context('USERENV','DB_NAME')) as NAME from DUAL);
         --dbms_output.put_line('System:'||vSystem);
            logVerbose('vSystem: >'||vSystem||'<',p_EX_SESSION);

            iab_rec.STL_SCHEMA   := LOWER(SUBSTR(vLine, 1, vPosDoppelP - 1)); -- ohne ':'
            iab_rec.STL_OBJ_NAME := SUBSTR(vLine, vPosDoppelP + 1, LENGTH(vLine));   -- ohne '_'
            iab_rec.STL_ID          := SEQ_STL_IDS.NEXTVAL;
            iab_rec.STL_LINE        := vLine;
            iab_rec.STL_SYSTEM      := lower(vSystem);
            iab_rec.STL_PERM        := 1;
            iab_rec.STL_KZ          := 'i';
            iab_rec.STL_REF_DATE    := p_REF_DATE;
            logVerbose('STL_SCHEMA: >'|| iab_rec.STL_SCHEMA||'< STL_OBJ_NAME: >'|| iab_rec.STL_OBJ_NAME||'< STL_ID: >'|| iab_rec.STL_ID||'< STL_LINE: >'|| iab_rec.STL_LINE||'< STL_STATUS: >'|| iab_rec.STL_STATUS||'< STL_SYSTEM: >'|| iab_rec.STL_SYSTEM||'< STL_PERM: >'|| iab_rec.STL_PERM||'< STL_KZ: >'|| iab_rec.STL_KZ||'< '||' STL_REF_DATE: >'||to_char(iab_rec.stl_ref_date,'DD.MM.YYYY HH24:MI:SS') ||'<',p_EX_SESSION);

            IF iab_rec.STL_SCHEMA IS NULL OR iab_rec.STL_OBJ_NAME IS NULL THEN
                PKG_LIB_EX_BASIC_NEWAPI.WRITELOG(P_EVENTTYPE  => PKG_LIB_EX_BASIC_NEWAPI.EventT_FuncError,
                              P_SESSION                => p_EX_SESSION,
                              P_IDENTIFICATION         => $$PLSQL_UNIT,
                              P_TXT                    => 'Fehler, Steuerlisteneintrag ' || vLine ||' ist nicht valide');
               GOTO end_loop;
            END IF;

            -- insert
            INSERT INTO WRK_EX_OBJ VALUES iab_rec;

          END IF;


          <<end_loop>>
          NULL;
        EXCEPTION
        WHEN no_data_found THEN
          UTL_FILE.FCLOSE(vFile);
          EXIT;
        END;
      END LOOP; -- Schleife inerhalb der einzelne Datei.
    COMMIT;
  EXCEPTION
  WHEN no_data_found THEN
    PKG_LIB_EX_BASIC_NEWAPI.WRITELOG(P_EVENTTYPE      => PKG_LIB_EX_BASIC_NEWAPI.EventT_FuncError,
                              P_SESSION        => p_EX_SESSION,
                              P_IDENTIFICATION => $$PLSQL_UNIT,
                              P_TXT            => 'Fehler, Steuerliste' || p_STL_NAME || ' nicht valide, System oder Datenbankschema nicht gefunden');
  WHEN UTL_FILE.INVALID_PATH THEN
    PKG_LIB_EX_BASIC_NEWAPI.WRITELOG(P_EVENTTYPE      => PKG_LIB_EX_BASIC_NEWAPI.EventT_FuncError,
                              P_SESSION        => p_EX_SESSION,
                              P_IDENTIFICATION => $$PLSQL_UNIT,
                              P_TXT            => 'Fehler, Pfad fuer Steuerliste ' || p_STL_NAME ||' ist nicht valide');
  WHEN UTL_FILE.INVALID_OPERATION THEN
    PKG_LIB_EX_BASIC_NEWAPI.WRITELOG(P_EVENTTYPE      => PKG_LIB_EX_BASIC_NEWAPI.EventT_FuncError,
                              P_SESSION        => p_EX_SESSION,
                              P_IDENTIFICATION => $$PLSQL_UNIT,
                              P_TXT            => 'Fehler, Steuerliste ' || p_STL_NAME || ' kann nicht gelesen werden');
  WHEN OTHERS THEN
    UTL_FILE.FCLOSE(vFile);
    PKG_LIB_EX_BASIC_NEWAPI.WRITELOG(P_EVENTTYPE      => PKG_LIB_EX_BASIC_NEWAPI.EventT_SysError,
                              P_SESSION        => p_EX_SESSION,
                              P_IDENTIFICATION => $$PLSQL_UNIT,
                              P_TXT            => 'Systemfehler');
    RAISE;
  END getSTLIAB;


/*------------------------------------------------------------------------------
Name:         setSTLDone

Parameter:    p_EX_SESSION     Nummer zu welchem Lauf die Steuerlisten gehoeren.
              p_STL_NAME       Name der zu ladenden Steuerliste

Beschreibung: Prozedur verschiebt Steuerliste ins DONE Verzeichis wenn sie
              fehlerfrei exportiert wurde.
------------------------------------------------------------------------------*/
  PROCEDURE setSTLDone(p_EX_SESSION IN NUMBER, p_STL_NAME VARCHAR2 )
  IS
  vAnz NUMBER;
  vExistsIAB INTEGER;
  BEGIN

  -- Fehler: Steuerlisteneintrag mit fehlerhafen oder fehlende Objekten
  FOR cOBJ IN ( SELECT *
                  FROM STAT_EXP_DATAMART_MCFG.WRK_EX_OBJ
                 WHERE STL_STATUS NOT IN ('A', 'E')
                   AND STL_FILE_NAME = p_STL_NAME
                   AND STL_SESSION = p_EX_SESSION)
  LOOP
    PKG_LIB_EX_BASIC_NEWAPI.WRITELOG(P_EVENTTYPE      => PKG_LIB_EX_BASIC_NEWAPI.EventT_FuncError,
                              P_SESSION        => p_EX_SESSION,
                              P_IDENTIFICATION => $$PLSQL_UNIT,
                              P_TXT            => 'Fehler, fuer Steuerlisteneintrag ' || chr(34) || cOBJ.STL_LINE || chr(34) || ' wurden nicht alle (oder keine) Objekte exportiert');
  END LOOP;

  vExistsIAB  := dbms_lob.fileexists(bfilename('EX_STL_IAB',  p_STL_NAME));

  SELECT count(*) INTO vAnz
     FROM STAT_EXP_DATAMART_MCFG.WRK_EX_OBJ
    WHERE STL_STATUS <> 'A'
      AND STL_FILE_NAME = p_STL_NAME
      AND STL_SESSION = p_EX_SESSION;

     IF  vExistsIAB <> 0 THEN
           PKG_LIB_EX_BASIC_NEWAPI.WRITELOG(P_EVENTTYPE      => PKG_LIB_EX_BASIC_NEWAPI.EventT_Information,
                                     P_SESSION        => p_EX_SESSION,
                                     P_IDENTIFICATION => $$PLSQL_UNIT,
                                     P_TXT            => 'IAB-Steuerliste ' || p_STL_NAME || ' wurden NICHT nach DONE verschoben');
     ELSIF vAnz = 0 AND p_STL_NAME NOT LIKE '%_perm.csv' THEN
       UTL_FILE.FRENAME(vDirTodo, p_STL_NAME, vDirDone, p_STL_NAME, TRUE);
       PKG_LIB_EX_BASIC_NEWAPI.WRITELOG(P_EVENTTYPE      => PKG_LIB_EX_BASIC_NEWAPI.EventT_Information,
                                 P_SESSION        => p_EX_SESSION,
                                 P_IDENTIFICATION => $$PLSQL_UNIT,
                                 P_TXT            => 'Steuerliste ' || p_STL_NAME || ' wurden nach DONE verschoben');
     ELSE
       IF  p_STL_NAME NOT LIKE '%_perm.csv' THEN
           PKG_LIB_EX_BASIC_NEWAPI.WRITELOG(P_EVENTTYPE      => PKG_LIB_EX_BASIC_NEWAPI.EventT_Warning,
                                     P_SESSION        => p_EX_SESSION,
                                     P_IDENTIFICATION => $$PLSQL_UNIT,
                                     P_TXT            => 'Steuerliste ' || p_STL_NAME || ' wurden NICHT nach DONE verschoben');
       ELSE
           PKG_LIB_EX_BASIC_NEWAPI.WRITELOG(P_EVENTTYPE      => PKG_LIB_EX_BASIC_NEWAPI.EventT_Information,
                                     P_SESSION        => p_EX_SESSION,
                                     P_IDENTIFICATION => $$PLSQL_UNIT,
                                     P_TXT            => 'Steuerliste ' || p_STL_NAME || ' wurden NICHT nach DONE verschoben');
       END IF;
     END IF;
  EXCEPTION
  WHEN UTL_FILE.INVALID_PATH THEN
    PKG_LIB_EX_BASIC_NEWAPI.WRITELOG(P_EVENTTYPE      => PKG_LIB_EX_BASIC_NEWAPI.EventT_FuncError,
                              P_SESSION        => p_EX_SESSION,
                              P_IDENTIFICATION => $$PLSQL_UNIT,
                              P_TXT            => 'Fehler, Pfad fuer Steuerliste ' || p_STL_NAME ||' ist nicht valide');
  WHEN UTL_FILE.RENAME_FAILED THEN
    PKG_LIB_EX_BASIC_NEWAPI.WRITELOG(P_EVENTTYPE      => PKG_LIB_EX_BASIC_NEWAPI.EventT_FuncError,
                              P_SESSION        => p_EX_SESSION,
                              P_IDENTIFICATION => $$PLSQL_UNIT,
                              P_TXT            => 'Fehler, Steuerliste ' || p_STL_NAME ||' kann nicht verschoben werden');
  WHEN UTL_FILE.INVALID_OPERATION THEN
    PKG_LIB_EX_BASIC_NEWAPI.WRITELOG(P_EVENTTYPE      => PKG_LIB_EX_BASIC_NEWAPI.EventT_FuncError,
                              P_SESSION        => p_EX_SESSION,
                              P_IDENTIFICATION => $$PLSQL_UNIT,
                              P_TXT            => 'Fehler, Steuerliste ' || p_STL_NAME ||' kann nicht verschoben werden');
  WHEN OTHERS THEN
    PKG_LIB_EX_BASIC_NEWAPI.WRITELOG(P_EVENTTYPE      => PKG_LIB_EX_BASIC_NEWAPI.EventT_SysError,
                              P_SESSION        => p_EX_SESSION,
                              P_IDENTIFICATION => $$PLSQL_UNIT,
                              P_TXT            => 'Systemfehler');
  END setSTLDone;


/*------------------------------------------------------------------------------
Name:         setReload

Parameter:    p_EX_SESSION     Nummer zu welchem Lauf die Steuerlisten gehoeren.
              p_STL_NAME       Name der zu ladenden Steuerliste

Beschreibung: Prozedur setzt alle Steuerlisten-Objekte einer Steuerliste auf
              "Abgeschlossen" die im letzten lauf der Steuerliste fehlerfrei
              abgearbeitet wurden.
------------------------------------------------------------------------------*/
  PROCEDURE setReload(p_EX_SESSION IN NUMBER, p_STL_NAME VARCHAR2,p_Ref_Date in date )
  IS
  vAnz            NUMBER;
  vAnzAltAbg      NUMBER;
  vAlt_EX_SESSION STAT_EXP_DATAMART_MCFG.WRK_EX_OBJ.STL_SESSION%TYPE;
  vAlt_ZBDAT      STAT_EXP_DATAMART_MCFG.WRK_EX_OBJ.STL_ZBDAT%TYPE;
  BEGIN

  --Session vom letzten lauf der Steuerliste ermitteln
  SELECT STL_SESSION,STL_ZBDAT
  INTO vAlt_EX_SESSION,vAlt_ZBDAT
  FROM
    (SELECT STL_SESSION,
            STL_ZBDAT
       FROM STAT_EXP_DATAMART_MCFG.WRK_EX_OBJ
      WHERE STL_FILE_NAME = p_STL_NAME
        AND STL_SESSION  <> p_EX_SESSION
        AND STL_REF_DATE = p_Ref_Date
      ORDER BY STL_ZBDAT DESC
    )
  WHERE rownum = 1;

  LogInfo('Zu '||p_stl_name||' alte Session >'||vAlt_EX_SESSION||'< von >'||to_char(vAlt_ZBDAT,'DD.MM.YYYY HH24:MI:SS')||'< gefunden!',p_EX_SESSION);

  -- fuer alle Eintraege der neuen Steuerliste
  FOR cOBJ IN ( SELECT *
                  FROM STAT_EXP_DATAMART_MCFG.WRK_EX_OBJ
                 WHERE STL_STATUS = 'L'
                   AND STL_FILE_NAME = p_STL_NAME
                   AND STL_SESSION   = p_EX_SESSION)
  LOOP

    -- Pruefen ob Steuerlisteneintrag der alte Steuerliste fehlerfrei verarbeitet wurde
    SELECT count(*)
      INTO vAnzAltAbg
      FROM STAT_EXP_DATAMART_MCFG.WRK_EX_OBJ
     WHERE STL_STATUS = 'A'
      AND STL_FILE_NAME = p_STL_NAME
      AND STL_LINE      = cOBJ.STL_LINE
      AND STL_SESSION   = vAlt_EX_SESSION;

     -- wenn alte Verearbeitung fehlerfrei dann abschliessen und Protokollieren
     IF vAnzAltAbg <> 0 THEN
       UPDATE STAT_EXP_DATAMART_MCFG.WRK_EX_OBJ SET STL_STATUS = 'A'  WHERE STL_ID = cOBJ.STL_ID;
       PKG_LIB_EX_BASIC_NEWAPI.WRITELOG(P_EVENTTYPE      => PKG_LIB_EX_BASIC_NEWAPI.EventT_Information,
                                 P_SESSION        => p_EX_SESSION,
                                 P_IDENTIFICATION => $$PLSQL_UNIT,
                                 P_TXT            => 'Steuerlisteneintrag ' || chr(34) || cOBJ.STL_LINE || chr(34) || ' wird bei R E L O A D nicht noch mal veraebeitet');
     END IF;
  END LOOP;

  EXCEPTION
  WHEN OTHERS THEN
    PKG_LIB_EX_BASIC_NEWAPI.WRITELOG(P_EVENTTYPE      => PKG_LIB_EX_BASIC_NEWAPI.EventT_SysError,
                              P_SESSION        => p_EX_SESSION,
                              P_IDENTIFICATION => $$PLSQL_UNIT,
                              P_TXT            => 'Systemfehler');
  END setReload;

END PKG_EX_LOAD_STL;
/
