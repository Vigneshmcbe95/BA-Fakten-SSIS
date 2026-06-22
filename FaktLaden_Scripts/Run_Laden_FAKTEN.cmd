@echo off
setlocal enabledelayedexpansion

:: ============================================================================================
:: DATEINAME:     Run_Laden_FAKTEN.cmd
:: SPEICHERORT:   E:\BDATA\Batches\Ladepakete\Run_Laden_FAKTEN.cmd
:: PROJEKT:       Data Warehouse - Ladeprozess (SSIS)
:: ZWECK:         Fuehrt das SSIS-Paket "Laden_FAKTEN.dtsx" zentral aus
::                fuer ALLE Verfahren (DWH, AST, BB, BST, FST, LST, etc.)
::
:: AUTOR:         Nils Ruemmeli
:: ERSTELLT AM:   06.05.2026
:: ZULETZT GEAENDERT: 07.05.2026
::
:: BESCHREIBUNG:
::   ZENTRALES Ladeskript - liegt einmalig unter Ladepakete.
::   UC4/Automic ruft dieses eine CMD fuer jedes Verfahren auf.
::   Das SSIS-Paket liegt ebenfalls zentral unter Ladepakete\Prog.
::
::   %1 = Oracle-Benutzername
::   Oracle-Kennwort wird ueber UC4-Variable %POLYBASEAUTH% bereitgestellt
::   %2 = SQL Server Instanz (z.B. Mig2022)
::   %3 = Vollstaendiger Windows-Pfad zur Steuerliste-CSV-Datei
::        (z.B. C:\Fakten\STL\Valide\steuerliste_fakten.csv)
::        -> BA::STLOrdner    = Ordner aus %3  (z.B. C:\Fakten\STL\Valide)
::        -> BA::STLDateiname = Dateiname aus %3 (z.B. steuerliste_fakten.csv)
::        -> BA::OrdnerArchiv = STL-Ordner + \Archiv (z.B. C:\Fakten\STL\Valide\Archiv)
::   %4 = Umgebung (z.B. ESTAT, ISTAT, PSTAT)
::        -> BA::ExtTableLocation = %4.STATRT.VM_DDL_SQL_SERVER
::        -> BA::ExtSourceName    = Oracle-%4
::   %5 = Oracle Hostname:Port (primaer) (z.B. cman-idst.idst.ibaintern.de:55436)
::        -> BA::ExtSourceLocation = oracle://%5
::   %6 = Oracle Hostname:Port (sekundaer)
::        -> BA::ExtConnectionOptions = Failover-String aus %6 (nur intern/Logging,
::           wird NICHT per /SET ans SSIS-Paket uebergeben - Variable existiert nicht im Paket)
::   %7 = Verfahren (z.B. FAKTEN, DWH, AST, BB, BST, FST, LST)
::        -> BA::Verfahren = %7 (direkt, ohne Suffix)
::        -> Logfile wird nach \\Server\Batches\%7\prog\log\ geschrieben
::   %8 = Datenbankname (z.B. msi_dm_lst_sgbii)
::        -> BA::Datenbank, BA::Datamart, BA::ProtokollDB, BA::ParameterDB verwenden diesen Wert
::        -> BA::Parametertabelle = tm_%8_param
::
:: VERWENDUNG:
::   Per UC4-Job:
::     Run_Laden_FAKTEN.cmd <User> <Instanz> <STL-Dateipfad> <Umgebung> <Oracle-Host:Port-Primaer> <Oracle-Host:Port-Sekundaer> <Verfahren> <Datenbankname>
::
::   Manuell (Test/Notfall):
::     set POLYBASEAUTH=meinKennwort
::     Run_Laden_FAKTEN.cmd myuser Mig2022 C:\Fakten\STL\Valide\steuerliste_fakten.csv ESTAT cman-idst.idst.ibaintern.de:55436 "" FAKTEN msi_dm_lst_sgbii
::
:: UNTERSCHIEDE ZU Run_Laden_DIM.cmd:
::   - SSIS-Paket: Laden_FAKTEN.dtsx (statt Laden_DIM.dtsx)
::   - OrdnerArchiv: STLOrdner + \Archiv (statt \done)
::   - SteuerlistenTabelle: tm_steuerlistenfile_Fakten (statt tm_steuerlistenfile_dimensionen)
::   - Zusaetzliche Variablen: BA::ParameterDB, BA::Parametertabelle
::     -> ParameterDB      = Datenbankname (%8)
::     -> Parametertabelle = tm_%8_param
::
:: TECHNISCHE HINWEISE:
::   - Das Oracle-Kennwort wird ueber die UC4-Variable %POLYBASEAUTH% bereitgestellt.
::   - Bei manuellem Test muss die Umgebungsvariable POLYBASEAUTH vorher gesetzt werden.
::   - Nach jeder Zeile mit ^ duerfen KEINE Leerzeichen folgen (CMD-Zeilenfortsetzung).
::   - Paket-Variablen die mit /SET ueberschrieben werden muessen auf Package-Ebene existieren.
::   - Das Logfile liegt unter \\Server\Batches\<Verfahren>\prog\log\ mit Timestamp und DB-Name.
::   - Verfahren steckt im Ordnerpfad, nicht nochmal im Dateinamen.
::   - Fehlende Pflichtparameter werden einzeln mit Namen ins Log geschrieben.
::   - DTExec und SSIS-Paket muessen existieren - fehlende Dateien fuehren zum Abbruch.
::   - Keine Passwoerter sind fest im Skript kodiert.
::   - PARAM_SERVER wird automatisch vom Computernamen abgeleitet.
::   - %3 ist ein Windows-Pfad mit backslashes - wird nativ verarbeitet.
::   - Pfad-Splitting erfolgt durch CMD-eigene Funktionen.
::   - Archiv-Ordner: STL-Ordner + \Archiv (C:\Fakten\STL\Valide -> C:\Fakten\STL\Valide\Archiv).
::   - EXT_TABLE_LOCATION wird aus %4 abgeleitet: %4.STATRT.VM_DDL_SQL_SERVER
::   - EXT_SOURCE_NAME wird aus %4 abgeleitet: Oracle-%4
::   - EXT_SOURCE_LOCATION wird aus %5 zusammengesetzt: oracle://%5
::   - EXT_CONNECTION_OPTIONS wird aus %6 als Failover-String generiert (nur Logging).
::   - BA::ExtConnectionOptions existiert NICHT im SSIS-Paket und wird nicht per /SET gesetzt.
::   - PARAM_VERFAHREN verwendet %7 direkt (z.B. FAKTEN).
::   - PARAM_DATENBANK, PARAM_DATAMART, PARAM_PROTOKOLL_DB und PARAM_PARAMETER_DB verwenden %8.
::   - PARAM_PARAMETERTABELLE wird abgeleitet: tm_%8_param
::   - Lockfile ist pro Verfahren+Datenbank, damit parallele Laeufe moeglich sind.
::   - ABSCHNITT 4 (Blockierungspruefung) laeuft VOR dem normalen Logging,
::     damit kein verwaistes Logfile entsteht.
:: ============================================================================================


:: ============================================================================================
:: ABSCHNITT 1 - UC4-JOB-PARAMETER
:: ============================================================================================

set ATOMIC_ORACLE_USERNAME=%~1
set ATOMIC_ORACLE_PASSWORD=%POLYBASEAUTH%
set ATOMIC_INSTANCE=%~2
set ATOMIC_STL_FULLPATH=%~3
set ATOMIC_ENVIRONMENT=%~4
set ATOMIC_ORACLE_HOSTPORT=%~5
set ATOMIC_ORACLE_HOSTPORT_SECONDARY=%~6
set ATOMIC_VERFAHREN=%~7
set ATOMIC_DATENBANKNAME=%~8

:: Dateiname extrahieren (z.B. steuerliste_fakten.csv)
for %%F in ("!ATOMIC_STL_FULLPATH!") do set ATOMIC_STL_DATEINAME=%%~nxF

:: Ordner extrahieren - verwende den Pfad EXAKT wie uebergeben
:: Entferne nur den Dateinamen am Ende
set ATOMIC_STL_ORDNER=!ATOMIC_STL_FULLPATH!
for %%F in ("!ATOMIC_STL_FULLPATH!") do set ATOMIC_STL_ORDNER=!ATOMIC_STL_ORDNER:%%~nxF=!
:: Entferne abschliessenden Backslash
if "!ATOMIC_STL_ORDNER:~-1!"=="\" set ATOMIC_STL_ORDNER=!ATOMIC_STL_ORDNER:~0,-1!


:: ============================================================================================
:: ABSCHNITT 2 - STATISCHE KONFIGURATION
:: ============================================================================================

:: --- Pfade (zentral unter Ladepakete) ---
set DTEXEC_PATH=C:\Program Files\Microsoft SQL Server\160\DTS\Binn\DTExec.exe
set SSIS_PACKAGE=E:\BDATA\Batches\Ladepakete\Fakten\Laden_FAKTEN.dtsx

:: --- Logfile mit Timestamp und Datenbankname ---
set LOG_DATE=%DATE:~6,4%%DATE:~3,2%%DATE:~0,2%
set LOG_TIME_RAW=%TIME: =0%
set LOG_TIME=%LOG_TIME_RAW:~0,2%%LOG_TIME_RAW:~3,2%%LOG_TIME_RAW:~6,2%

set LOG_FILENAME=Laden_FAKTEN_!ATOMIC_DATENBANKNAME!_!LOG_DATE!_!LOG_TIME!.log
set LOG_DIR=\\!COMPUTERNAME!\Batches\!ATOMIC_VERFAHREN!\prog\log
set LOGFILE=!LOG_DIR!\!LOG_FILENAME!

:: --- Lockfile pro Verfahren + Datenbank ---
set LOCKFILE_DIR=E:\BDATA\Batches\Ladepakete\log
set LOCKFILE=!LOCKFILE_DIR!\SSIS_FAKTEN_!ATOMIC_VERFAHREN!_!ATOMIC_DATENBANKNAME!.lock

:: --- SQL Server / Datenbank ---
set PARAM_SERVER=%COMPUTERNAME%
set PARAM_CONNECTION_SERVER=!PARAM_SERVER!\!ATOMIC_INSTANCE!
set PARAM_DATENBANK=!ATOMIC_DATENBANKNAME!
set PARAM_VERFAHREN=!ATOMIC_VERFAHREN!
set PARAM_DATAMART=!ATOMIC_DATENBANKNAME!
set PARAM_PROTOKOLL_DB=!ATOMIC_DATENBANKNAME!
set PARAM_PROTOKOLL_SP=usp_SSIS_Protokoll
set PARAM_PROTOKOLL_TABELLE=tm_!ATOMIC_DATENBANKNAME!_prot

:: --- Zusaetzliche FAKTEN-Parameter ---
set PARAM_PARAMETER_DB=!ATOMIC_DATENBANKNAME!
set PARAM_PARAMETERTABELLE=tm_!ATOMIC_DATENBANKNAME!_param

:: --- Dateipfade ---
set PARAM_STL_ORDNER=!ATOMIC_STL_ORDNER!
set PARAM_STL_DATEINAME=!ATOMIC_STL_DATEINAME!
:: FAKTEN verwendet \Archiv als Archiv-Unterordner (DIM verwendet \done)
set PARAM_ORDNER_ARCHIV=!ATOMIC_STL_ORDNER!\Archiv

:: --- Verarbeitung ---
set PARAM_MAXPARALLEL=8

:: --- Index-Erstellung (SCR16): MAXDOP je einzelnem CI/CCI-Build ---
:: Unabhaengig von MAXPARALLEL. Gesamtlast ~ MAXPARALLEL * INDEX_MAXDOP.
set PARAM_INDEX_MAXDOP=4

:: --- Externe Oracle-Quelle ---
set PARAM_EXT_SOURCE_NAME=Oracle-!ATOMIC_ENVIRONMENT!
set PARAM_EXT_SOURCE_LOCATION=oracle://!ATOMIC_ORACLE_HOSTPORT!
set PARAM_EXT_TABLE_LOCATION=!ATOMIC_ENVIRONMENT!.STATRT.VM_DDL_SQL_SERVER
set PARAM_EXT_TABLE_SCHEMA=ext
set PARAM_EXT_TABLE_NAME=vm_ddl_sql_server

:: --- Oracle-Schema der Partitionssicht V_PARTITION_INFO (Owner-Filter in SCR11) ---
set PARAM_PARTITION_SCHEMA=BI_DM_EXPORT

:: --- CONNECTION_OPTIONS fuer Failover (nur Logging - nicht ans SSIS-Paket uebergeben) ---
set PARAM_EXT_CONNECTION_OPTIONS=
set FAILOVER_CONFIGURED=0

if not "!ATOMIC_ORACLE_HOSTPORT_SECONDARY!"=="" (
    if not "!ATOMIC_ORACLE_HOSTPORT_SECONDARY!"=="""" (
        set FAILOVER_CONFIGURED=1
        for /f "tokens=1,2 delims=:" %%a in ("!ATOMIC_ORACLE_HOSTPORT_SECONDARY!") do (
            set SECONDARY_HOST=%%a
            set SECONDARY_PORT=%%b
        )
        set "RAW_OPTIONS=AlternateServers=(HostName=!SECONDARY_HOST!:PortNumber=!SECONDARY_PORT!:ServiceName=!ATOMIC_ENVIRONMENT!);FailoverMode=0;LoadBalancing=0"
        set "PARAM_EXT_CONNECTION_OPTIONS=!RAW_OPTIONS:;=%%3B!"
    )
)

:: --- Steuerliste ---
:: FAKTEN verwendet tm_steuerlistenfile_Fakten (DIM verwendet tm_steuerlistenfile_dimensionen)
set PARAM_STEUERLISTEN_TABELLE=tm_steuerlistenfile_Fakten


:: ============================================================================================
:: ABSCHNITT 4 - SCHUTZ VOR MEHRFACHAUSFUEHRUNG (pro Instanz + Datenbank + Verfahren)
::
:: Prueft ob bereits ein DTExec-Prozess fuer GENAU diese Kombination laeuft.
:: Andere Verfahren/Datenbanken koennen parallel laufen.
:: Laeuft VOR dem normalen Logging damit kein verwaistes Logfile entsteht.
:: ============================================================================================

set PROCESS_RUNNING=0
set BLOCKING_PID=

for /f "tokens=2 delims==" %%P in ('wmic process where "name='DTExec.exe'" get processid /format:list 2^>nul ^| find "="') do (
    set CURRENT_PID=%%P
    for /f "tokens=*" %%C in ('wmic process where "processid='!CURRENT_PID!'" get commandline /format:list 2^>nul ^| find /I "BA::Datenbank].Value;""!ATOMIC_DATENBANKNAME!"""') do (
        set PROCESS_RUNNING=1
        set BLOCKING_PID=!CURRENT_PID!
    )
)

if !PROCESS_RUNNING!==1 (
    if not exist "!LOG_DIR!" mkdir "!LOG_DIR!" 2>nul

    set LOG_FILENAME=BLOCKED_Laden_FAKTEN_!ATOMIC_DATENBANKNAME!_!LOG_DATE!_!LOG_TIME!.log
    set LOGFILE=!LOG_DIR!\!LOG_FILENAME!

    echo.
    echo ========================================================
    echo  BLOCKIERT: Paket laeuft bereits!
    echo  Verfahren  : !ATOMIC_VERFAHREN!
    echo  Datenbank  : !ATOMIC_DATENBANKNAME!
    echo  Instanz    : !PARAM_CONNECTION_SERVER!
    echo  Prozess-ID : !BLOCKING_PID!
    echo  Logfile    : !LOGFILE!
    echo ========================================================
    echo.

    echo [%DATE% %TIME%] ================================================================= > "!LOGFILE!"
    echo [%DATE% %TIME%] BLOCKIERT - SSIS-Paket konnte nicht gestartet werden               >> "!LOGFILE!"
    echo [%DATE% %TIME%] ================================================================= >> "!LOGFILE!"
    echo.                                                                                   >> "!LOGFILE!"
    echo [%DATE% %TIME%] WAS WURDE VERSUCHT:                                                >> "!LOGFILE!"
    echo [%DATE% %TIME%]   Das SSIS-Paket Laden_FAKTEN.dtsx sollte gestartet werden fuer:   >> "!LOGFILE!"
    echo [%DATE% %TIME%]     Verfahren  : !ATOMIC_VERFAHREN!                                >> "!LOGFILE!"
    echo [%DATE% %TIME%]     Datenbank  : !ATOMIC_DATENBANKNAME!                            >> "!LOGFILE!"
    echo [%DATE% %TIME%]     Instanz    : !PARAM_CONNECTION_SERVER!                         >> "!LOGFILE!"
    echo [%DATE% %TIME%]     STL Datei  : !PARAM_STL_DATEINAME!                             >> "!LOGFILE!"
    echo.                                                                                   >> "!LOGFILE!"
    echo [%DATE% %TIME%] WARUM BLOCKIERT:                                                    >> "!LOGFILE!"
    echo [%DATE% %TIME%]   Fuer die Kombination Verfahren=!ATOMIC_VERFAHREN! und             >> "!LOGFILE!"
    echo [%DATE% %TIME%]   Datenbank=!ATOMIC_DATENBANKNAME! laeuft bereits ein               >> "!LOGFILE!"
    echo [%DATE% %TIME%]   DTExec-Prozess auf diesem Server.                                 >> "!LOGFILE!"
    echo [%DATE% %TIME%]   Prozess-ID: !BLOCKING_PID!                                       >> "!LOGFILE!"
    echo [%DATE% %TIME%]   Zwei Laeufe fuer dieselbe Kombination wuerden zu                  >> "!LOGFILE!"
    echo [%DATE% %TIME%]   Dateninkonsistenzen und Sperrkonflikten fuehren.                  >> "!LOGFILE!"
    echo.                                                                                   >> "!LOGFILE!"
    echo [%DATE% %TIME%] WAS SIE TUN KOENNEN:                                               >> "!LOGFILE!"
    echo [%DATE% %TIME%]                                                                     >> "!LOGFILE!"
    echo [%DATE% %TIME%]   Option 1: Warten bis der laufende Prozess beendet ist             >> "!LOGFILE!"
    echo [%DATE% %TIME%]             und danach erneut starten.                              >> "!LOGFILE!"
    echo [%DATE% %TIME%]                                                                     >> "!LOGFILE!"
    echo [%DATE% %TIME%]   Option 2: Blockierenden Prozess pruefen und beenden:              >> "!LOGFILE!"
    echo [%DATE% %TIME%]                                                                     >> "!LOGFILE!"
    echo [%DATE% %TIME%]     Schritt 1 - Prozess inspizieren:                                >> "!LOGFILE!"
    echo [%DATE% %TIME%]       wmic process where "processid='!BLOCKING_PID!'" get processid,creationdate,status >> "!LOGFILE!"
    echo [%DATE% %TIME%]                                                                     >> "!LOGFILE!"
    echo [%DATE% %TIME%]     Schritt 2 - Prozess beenden:                                   >> "!LOGFILE!"
    echo [%DATE% %TIME%]       taskkill /F /PID !BLOCKING_PID!                               >> "!LOGFILE!"
    echo [%DATE% %TIME%]                                                                     >> "!LOGFILE!"
    echo [%DATE% %TIME%]     Schritt 3 - Paket erneut starten.                              >> "!LOGFILE!"
    echo [%DATE% %TIME%]                                                                     >> "!LOGFILE!"
    echo [%DATE% %TIME%]   Option 3: In UC4 den vorherigen Job pruefen und ggf.              >> "!LOGFILE!"
    echo [%DATE% %TIME%]             abbrechen, dann diesen Job erneut einplanen.            >> "!LOGFILE!"
    echo.                                                                                   >> "!LOGFILE!"
    echo [%DATE% %TIME%] LAUFENDE DTEXEC-PROZESSE AUF DIESEM SERVER:                        >> "!LOGFILE!"
    echo [%DATE% %TIME%] -----------------------------------------------------------------  >> "!LOGFILE!"
    wmic process where "name='DTExec.exe'" get processid,creationdate /format:list 2>nul >> "!LOGFILE!"
    echo [%DATE% %TIME%] -----------------------------------------------------------------  >> "!LOGFILE!"
    echo.                                                                                   >> "!LOGFILE!"
    echo [%DATE% %TIME%] ================================================================= >> "!LOGFILE!"
    echo [%DATE% %TIME%] Exitcode: 2 ^(BLOCKIERT^)                                          >> "!LOGFILE!"
    echo [%DATE% %TIME%] ================================================================= >> "!LOGFILE!"

    exit /b 2
)

:: Kein laufender Prozess - Lockfile erstellen mit Details
echo !PARAM_CONNECTION_SERVER! !ATOMIC_DATENBANKNAME! !ATOMIC_VERFAHREN! %DATE% %TIME% > "!LOCKFILE!"


:: ============================================================================================
:: ABSCHNITT 2b - LOG-ORDNER UND ARCHIV-ORDNER ERSTELLEN
:: ============================================================================================

if not exist "!LOG_DIR!" (
    mkdir "!LOG_DIR!" 2>nul
)

echo [%DATE% %TIME%] =============================================== > "!LOGFILE!"
echo [%DATE% %TIME%] SSIS Faktenladen - Zentrales Ladeskript        >> "!LOGFILE!"
echo [%DATE% %TIME%] Verfahren   : !ATOMIC_VERFAHREN!               >> "!LOGFILE!"
echo [%DATE% %TIME%] Datenbank   : !ATOMIC_DATENBANKNAME!           >> "!LOGFILE!"
echo [%DATE% %TIME%] Logdatei    : !LOG_FILENAME!                   >> "!LOGFILE!"
echo [%DATE% %TIME%] =============================================== >> "!LOGFILE!"

echo [%DATE% %TIME%] Archiv-Ordner (Windows-Pfad): !PARAM_ORDNER_ARCHIV! >> "!LOGFILE!"

if not exist "!PARAM_ORDNER_ARCHIV!" (
    echo [%DATE% %TIME%] [INFO]     Archiv-Ordner existiert nicht - wird erstellt... >> "!LOGFILE!"
    mkdir "!PARAM_ORDNER_ARCHIV!" 2>nul
    if !ERRORLEVEL!==0 (
        echo [%DATE% %TIME%] [OK]       Archiv-Ordner erfolgreich erstellt. >> "!LOGFILE!"
    ) else (
        echo [%DATE% %TIME%] [FEHLER]   Archiv-Ordner konnte nicht erstellt werden! >> "!LOGFILE!"
    )
) else (
    echo [%DATE% %TIME%] [OK]       Archiv-Ordner existiert bereits. >> "!LOGFILE!"
)


:: ============================================================================================
:: ABSCHNITT 3 - VORABPRUEFUNG DER PARAMETER
:: ============================================================================================

set VALIDATION_FAILED=0

echo [%DATE% %TIME%] =============================================== >> "!LOGFILE!"
echo [%DATE% %TIME%] VORABPRUEFUNG DER PARAMETER GESTARTET         >> "!LOGFILE!"
echo [%DATE% %TIME%] =============================================== >> "!LOGFILE!"

if "!ATOMIC_ORACLE_USERNAME!"=="" (
    echo [%DATE% %TIME%] [FEHLEND]  ATOMIC_ORACLE_USERNAME ^(%%1^) >> "!LOGFILE!"
    set VALIDATION_FAILED=1
) else (
    echo [%DATE% %TIME%] [OK]       ATOMIC_ORACLE_USERNAME ist gesetzt >> "!LOGFILE!"
)

if "!ATOMIC_ORACLE_PASSWORD!"=="" (
    echo [%DATE% %TIME%] [FEHLEND]  ATOMIC_ORACLE_PASSWORD ^(UC4-Variable POLYBASEAUTH nicht gesetzt^) >> "!LOGFILE!"
    set VALIDATION_FAILED=1
) else (
    echo [%DATE% %TIME%] [OK]       ATOMIC_ORACLE_PASSWORD ist gesetzt via POLYBASEAUTH [Wert verborgen] >> "!LOGFILE!"
)

if "!ATOMIC_INSTANCE!"=="" (
    echo [%DATE% %TIME%] [FEHLEND]  ATOMIC_INSTANCE ^(%%2^) >> "!LOGFILE!"
    set VALIDATION_FAILED=1
) else (
    echo [%DATE% %TIME%] [OK]       ATOMIC_INSTANCE          = !ATOMIC_INSTANCE! >> "!LOGFILE!"
    echo [%DATE% %TIME%] [OK]       PARAM_SERVER             = !PARAM_SERVER! ^(Computername^) >> "!LOGFILE!"
    echo [%DATE% %TIME%] [OK]       PARAM_CONNECTION_SERVER  = !PARAM_CONNECTION_SERVER! >> "!LOGFILE!"
)

if "!ATOMIC_STL_ORDNER!"=="" (
    echo [%DATE% %TIME%] [FEHLEND]  ATOMIC_STL_ORDNER ^(Ordner aus %%3^) >> "!LOGFILE!"
    set VALIDATION_FAILED=1
) else (
    echo [%DATE% %TIME%] [OK]       PARAM_STL_ORDNER         = !PARAM_STL_ORDNER! >> "!LOGFILE!"
    echo [%DATE% %TIME%] [OK]       PARAM_ORDNER_ARCHIV      = !PARAM_ORDNER_ARCHIV! >> "!LOGFILE!"
)

if "!ATOMIC_STL_DATEINAME!"=="" (
    echo [%DATE% %TIME%] [FEHLEND]  ATOMIC_STL_DATEINAME ^(Dateiname aus %%3^) >> "!LOGFILE!"
    set VALIDATION_FAILED=1
) else (
    echo [%DATE% %TIME%] [OK]       PARAM_STL_DATEINAME      = !PARAM_STL_DATEINAME! >> "!LOGFILE!"
)

if "!ATOMIC_ENVIRONMENT!"=="" (
    echo [%DATE% %TIME%] [FEHLEND]  ATOMIC_ENVIRONMENT ^(%%4^) >> "!LOGFILE!"
    set VALIDATION_FAILED=1
) else (
    echo [%DATE% %TIME%] [OK]       ATOMIC_ENVIRONMENT       = !ATOMIC_ENVIRONMENT! >> "!LOGFILE!"
    echo [%DATE% %TIME%] [OK]       PARAM_EXT_TABLE_LOCATION = !PARAM_EXT_TABLE_LOCATION! >> "!LOGFILE!"
    echo [%DATE% %TIME%] [OK]       PARAM_EXT_SOURCE_NAME    = !PARAM_EXT_SOURCE_NAME! ^(aus Umgebung abgeleitet^) >> "!LOGFILE!"
)

if "!ATOMIC_ORACLE_HOSTPORT!"=="" (
    echo [%DATE% %TIME%] [FEHLEND]  ATOMIC_ORACLE_HOSTPORT ^(%%5^) >> "!LOGFILE!"
    set VALIDATION_FAILED=1
) else (
    echo [%DATE% %TIME%] [OK]       ATOMIC_ORACLE_HOSTPORT   = !ATOMIC_ORACLE_HOSTPORT! >> "!LOGFILE!"
    echo [%DATE% %TIME%] [OK]       PARAM_EXT_SOURCE_LOCATION= !PARAM_EXT_SOURCE_LOCATION! >> "!LOGFILE!"
)

if "!ATOMIC_ORACLE_HOSTPORT_SECONDARY!"=="" (
    echo [%DATE% %TIME%] [WARNUNG]  ATOMIC_ORACLE_HOSTPORT_SECONDARY ^(%%6^) nicht angegeben - Failover deaktiviert >> "!LOGFILE!"
) else (
    if "!ATOMIC_ORACLE_HOSTPORT_SECONDARY!"=="""" (
        echo [%DATE% %TIME%] [WARNUNG]  ATOMIC_ORACLE_HOSTPORT_SECONDARY ^(%%6^) ist leer - Failover deaktiviert >> "!LOGFILE!"
    ) else (
        echo [%DATE% %TIME%] [OK]       ATOMIC_ORACLE_HOSTPORT_SECONDARY = !ATOMIC_ORACLE_HOSTPORT_SECONDARY! >> "!LOGFILE!"
        if !FAILOVER_CONFIGURED!==1 (
            echo [%DATE% %TIME%] [INFO]     PARAM_EXT_CONNECTION_OPTIONS ^(nur Logging, nicht ans Paket^) = !PARAM_EXT_CONNECTION_OPTIONS! >> "!LOGFILE!"
        )
    )
)

if "!ATOMIC_VERFAHREN!"=="" (
    echo [%DATE% %TIME%] [FEHLEND]  ATOMIC_VERFAHREN ^(%%7^) >> "!LOGFILE!"
    set VALIDATION_FAILED=1
) else (
    echo [%DATE% %TIME%] [OK]       ATOMIC_VERFAHREN         = !ATOMIC_VERFAHREN! >> "!LOGFILE!"
    echo [%DATE% %TIME%] [OK]       PARAM_VERFAHREN          = !PARAM_VERFAHREN! >> "!LOGFILE!"
)

if "!ATOMIC_DATENBANKNAME!"=="" (
    echo [%DATE% %TIME%] [FEHLEND]  ATOMIC_DATENBANKNAME ^(%%8^) >> "!LOGFILE!"
    set VALIDATION_FAILED=1
) else (
    echo [%DATE% %TIME%] [OK]       ATOMIC_DATENBANKNAME     = !ATOMIC_DATENBANKNAME! >> "!LOGFILE!"
    echo [%DATE% %TIME%] [OK]       PARAM_DATENBANK          = !PARAM_DATENBANK! >> "!LOGFILE!"
    echo [%DATE% %TIME%] [OK]       PARAM_DATAMART           = !PARAM_DATAMART! >> "!LOGFILE!"
    echo [%DATE% %TIME%] [OK]       PARAM_PROTOKOLL_DB       = !PARAM_PROTOKOLL_DB! >> "!LOGFILE!"
    echo [%DATE% %TIME%] [OK]       PARAM_PARAMETER_DB       = !PARAM_PARAMETER_DB! >> "!LOGFILE!"
    echo [%DATE% %TIME%] [OK]       PARAM_PARAMETERTABELLE   = !PARAM_PARAMETERTABELLE! ^(abgeleitet aus Datenbankname^) >> "!LOGFILE!"
)

if "!DTEXEC_PATH!"=="" (
    echo [%DATE% %TIME%] [FEHLEND]  DTEXEC_PATH ist leer >> "!LOGFILE!"
    set VALIDATION_FAILED=1
) else (
    echo [%DATE% %TIME%] [OK]       DTEXEC_PATH              = !DTEXEC_PATH! >> "!LOGFILE!"
    if not exist "!DTEXEC_PATH!" (
        echo [%DATE% %TIME%] [FEHLER]   DTExec.exe nicht gefunden unter: !DTEXEC_PATH! >> "!LOGFILE!"
        set VALIDATION_FAILED=1
    )
)

if "!SSIS_PACKAGE!"=="" (
    echo [%DATE% %TIME%] [FEHLEND]  SSIS_PACKAGE ist leer >> "!LOGFILE!"
    set VALIDATION_FAILED=1
) else (
    echo [%DATE% %TIME%] [OK]       SSIS_PACKAGE             = !SSIS_PACKAGE! >> "!LOGFILE!"
    if not exist "!SSIS_PACKAGE!" (
        echo [%DATE% %TIME%] [FEHLER]   SSIS-Paket nicht gefunden unter: !SSIS_PACKAGE! >> "!LOGFILE!"
        set VALIDATION_FAILED=1
    )
)

echo [%DATE% %TIME%] [OK]       PARAM_PROTOKOLL_SP       = !PARAM_PROTOKOLL_SP! >> "!LOGFILE!"
echo [%DATE% %TIME%] [OK]       PARAM_PROTOKOLL_TABELLE  = !PARAM_PROTOKOLL_TABELLE! >> "!LOGFILE!"

if "!PARAM_MAXPARALLEL!"=="" (
    echo [%DATE% %TIME%] [FEHLEND]  PARAM_MAXPARALLEL >> "!LOGFILE!"
    set VALIDATION_FAILED=1
) else (
    echo [%DATE% %TIME%] [OK]       PARAM_MAXPARALLEL        = !PARAM_MAXPARALLEL! >> "!LOGFILE!"
)

:: PARAM_INDEX_MAXDOP ist optional - das Paket faellt bei fehlendem Wert auf 4 zurueck.
if "!PARAM_INDEX_MAXDOP!"=="" (
    echo [%DATE% %TIME%] [WARNUNG]  PARAM_INDEX_MAXDOP nicht gesetzt - Paket verwendet Standard 4 >> "!LOGFILE!"
) else (
    echo [%DATE% %TIME%] [OK]       PARAM_INDEX_MAXDOP       = !PARAM_INDEX_MAXDOP! >> "!LOGFILE!"
)

if "!PARAM_EXT_SOURCE_NAME!"=="" (
    echo [%DATE% %TIME%] [FEHLEND]  PARAM_EXT_SOURCE_NAME >> "!LOGFILE!"
    set VALIDATION_FAILED=1
) else (
    echo [%DATE% %TIME%] [OK]       PARAM_EXT_SOURCE_NAME    = !PARAM_EXT_SOURCE_NAME! >> "!LOGFILE!"
)

if "!PARAM_EXT_TABLE_SCHEMA!"=="" (
    echo [%DATE% %TIME%] [FEHLEND]  PARAM_EXT_TABLE_SCHEMA >> "!LOGFILE!"
    set VALIDATION_FAILED=1
) else (
    echo [%DATE% %TIME%] [OK]       PARAM_EXT_TABLE_SCHEMA   = !PARAM_EXT_TABLE_SCHEMA! >> "!LOGFILE!"
)

if "!PARAM_EXT_TABLE_NAME!"=="" (
    echo [%DATE% %TIME%] [FEHLEND]  PARAM_EXT_TABLE_NAME >> "!LOGFILE!"
    set VALIDATION_FAILED=1
) else (
    echo [%DATE% %TIME%] [OK]       PARAM_EXT_TABLE_NAME     = !PARAM_EXT_TABLE_NAME! >> "!LOGFILE!"
)

if "!PARAM_PARTITION_SCHEMA!"=="" (
    echo [%DATE% %TIME%] [FEHLEND]  PARAM_PARTITION_SCHEMA >> "!LOGFILE!"
    set VALIDATION_FAILED=1
) else (
    echo [%DATE% %TIME%] [OK]       PARAM_PARTITION_SCHEMA   = !PARAM_PARTITION_SCHEMA! >> "!LOGFILE!"
)

if "!PARAM_STEUERLISTEN_TABELLE!"=="" (
    echo [%DATE% %TIME%] [FEHLEND]  PARAM_STEUERLISTEN_TABELLE >> "!LOGFILE!"
    set VALIDATION_FAILED=1
) else (
    echo [%DATE% %TIME%] [OK]       PARAM_STEUERLISTEN_TABELLE= !PARAM_STEUERLISTEN_TABELLE! >> "!LOGFILE!"
)

echo [%DATE% %TIME%] =============================================== >> "!LOGFILE!"
echo [%DATE% %TIME%] VALIDATION_FAILED = !VALIDATION_FAILED!        >> "!LOGFILE!"
echo [%DATE% %TIME%] =============================================== >> "!LOGFILE!"

if !VALIDATION_FAILED!==1 (
    echo [%DATE% %TIME%] [ABGEBROCHEN] Siehe [FEHLEND]/[FEHLER] Eintraege oben. >> "!LOGFILE!"
    echo [%DATE% %TIME%]               DTExec wird NICHT gestartet.             >> "!LOGFILE!"
    echo [%DATE% %TIME%] =============================================== >> "!LOGFILE!"
    del "!LOCKFILE!" 2>nul
    exit /b 1
)

echo [%DATE% %TIME%] VORABPRUEFUNG BESTANDEN - Ausfuehrung startet. >> "!LOGFILE!"
echo [%DATE% %TIME%] =============================================== >> "!LOGFILE!"


:: ============================================================================================
:: ABSCHNITT 5 - SSIS-PAKET AUSFUEHREN
:: ============================================================================================

echo [%DATE% %TIME%] SSIS-Ausfuehrung gestartet       >> "!LOGFILE!"
echo [%DATE% %TIME%] Verfahren       : !ATOMIC_VERFAHREN!    >> "!LOGFILE!"
echo [%DATE% %TIME%] Datenbankname   : !ATOMIC_DATENBANKNAME! >> "!LOGFILE!"
echo [%DATE% %TIME%] STL Ordner      : !PARAM_STL_ORDNER!    >> "!LOGFILE!"
echo [%DATE% %TIME%] STL Dateiname   : !PARAM_STL_DATEINAME! >> "!LOGFILE!"
echo [%DATE% %TIME%] Archiv Ordner   : !PARAM_ORDNER_ARCHIV! >> "!LOGFILE!"
echo [%DATE% %TIME%] Umgebung        : !ATOMIC_ENVIRONMENT!  >> "!LOGFILE!"
echo [%DATE% %TIME%] Oracle Host:Port (primaer)   : !ATOMIC_ORACLE_HOSTPORT! >> "!LOGFILE!"
echo [%DATE% %TIME%] Oracle Host:Port (sekundaer) : !ATOMIC_ORACLE_HOSTPORT_SECONDARY! >> "!LOGFILE!"
echo [%DATE% %TIME%] EXT Source Name : !PARAM_EXT_SOURCE_NAME! >> "!LOGFILE!"
echo [%DATE% %TIME%] EXT Source Loc  : !PARAM_EXT_SOURCE_LOCATION! >> "!LOGFILE!"
echo [%DATE% %TIME%] EXT Table Loc   : !PARAM_EXT_TABLE_LOCATION! >> "!LOGFILE!"
echo [%DATE% %TIME%] Parameter DB    : !PARAM_PARAMETER_DB! >> "!LOGFILE!"
echo [%DATE% %TIME%] Parametertabelle: !PARAM_PARAMETERTABELLE! >> "!LOGFILE!"
echo ----------------------------------------------- >> "!LOGFILE!"

"!DTEXEC_PATH!" ^
  /F "!SSIS_PACKAGE!" ^
  /REPORTING IEW ^
  /SET \Package.Variables[BA::Server].Value;"!PARAM_SERVER!" ^
  /SET \Package.Variables[BA::ConnectionServerName].Value;"!PARAM_CONNECTION_SERVER!" ^
  /SET \Package.Variables[BA::Datenbank].Value;"!PARAM_DATENBANK!" ^
  /SET \Package.Variables[BA::Verfahren].Value;"!PARAM_VERFAHREN!" ^
  /SET \Package.Variables[BA::Datamart].Value;"!PARAM_DATAMART!" ^
  /SET \Package.Variables[BA::ProtokollDB].Value;"!PARAM_PROTOKOLL_DB!" ^
  /SET \Package.Variables[BA::ProtokollSP].Value;"!PARAM_PROTOKOLL_SP!" ^
  /SET \Package.Variables[BA::Protokolltabelle].Value;"!PARAM_PROTOKOLL_TABELLE!" ^
  /SET \Package.Variables[BA::STLOrdner].Value;"!PARAM_STL_ORDNER!" ^
  /SET \Package.Variables[BA::STLDateiname].Value;"!PARAM_STL_DATEINAME!" ^
  /SET \Package.Variables[BA::OrdnerArchiv].Value;"!PARAM_ORDNER_ARCHIV!" ^
  /SET \Package.Variables[BA::CredBenutzername].Value;"!ATOMIC_ORACLE_USERNAME!" ^
  /SET \Package.Variables[BA::CredKennwort].Value;"!ATOMIC_ORACLE_PASSWORD!" ^
  /SET \Package.Variables[BA::Maxparallel].Value;"!PARAM_MAXPARALLEL!" ^
  /SET \Package.Variables[BA::IndexMaxDop].Value;"!PARAM_INDEX_MAXDOP!" ^
  /SET \Package.Variables[BA::ExtSourceName].Value;"!PARAM_EXT_SOURCE_NAME!" ^
  /SET \Package.Variables[BA::ExtSourceLocation].Value;"!PARAM_EXT_SOURCE_LOCATION!" ^
  /SET \Package.Variables[BA::ExtTableLocation].Value;"!PARAM_EXT_TABLE_LOCATION!" ^
  /SET \Package.Variables[BA::ExtTableSchema].Value;"!PARAM_EXT_TABLE_SCHEMA!" ^
  /SET \Package.Variables[BA::ExtTableName].Value;"!PARAM_EXT_TABLE_NAME!" ^
  /SET \Package.Variables[BA::partition_schema].Value;"!PARAM_PARTITION_SCHEMA!" ^
  /SET \Package.Variables[BA::SteuerlistenTabelle].Value;"!PARAM_STEUERLISTEN_TABELLE!" ^
  /SET \Package.Variables[BA::ParameterDB].Value;"!PARAM_PARAMETER_DB!" ^
  /SET \Package.Variables[BA::Parametertabelle].Value;"!PARAM_PARAMETERTABELLE!" ^
  >> "!LOGFILE!" 2>&1

set SSIS_EXIT=!ERRORLEVEL!


:: ============================================================================================
:: ABSCHNITT 6 - PROTOKOLLABSCHLUSS UND AUFRAEUMEN
:: ============================================================================================

echo. >> "!LOGFILE!"
echo [%DATE% %TIME%] SSIS-Ausfuehrung abgeschlossen >> "!LOGFILE!"
echo [%DATE% %TIME%] Exitcode: !SSIS_EXIT!           >> "!LOGFILE!"
echo ----------------------------------------------- >> "!LOGFILE!"

del "!LOCKFILE!" 2>nul

exit /b !SSIS_EXIT!
