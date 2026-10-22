@echo off
setlocal enabledelayedexpansion

:: ============================================================================================
:: DATEINAME:     Kill_Laden_FAKTEN.cmd
:: SPEICHERORT:   E:\BDATA\Batches\Ladepakete\Kill_Laden_FAKTEN.cmd
:: PROJEKT:       Data Warehouse - Ladeprozess (SSIS)
:: ZWECK:         Beendet den laufenden DTExec-Prozess (Laden_FAKTEN.dtsx) fuer eine
::                Kombination Verfahren + Datenbank, wenn der Abbruch ueber UC4/Automic
::                das SSIS-Paket nicht beendet hat.
::
:: AUTOR:         Nils Ruemmeli
:: ERSTELLT AM:   22.10.2026
:: ZULETZT GEAENDERT: 22.10.2026
::
:: BESCHREIBUNG:
::   Ein Abbruch in UC4/Automic beendet nur die cmd.exe des Jobs (Run_Laden_FAKTEN.cmd).
::   Der Kindprozess DTExec.exe laeuft weiter. Dieses Skript sucht den DTExec-Prozess
::   mit derselben Methode wie ABSCHNITT 4 in Run_Laden_FAKTEN.cmd (Kommandozeile enthaelt
::   Laden_FAKTEN.dtsx und BA::Datenbank = Datenbankname) und beendet ihn samt
::   Kindprozessen. Anschliessend wird das Lockfile geloescht.
::
::   %1 = Verfahren (z.B. FAKTEN, DWH, AST, BB, BST, FST, LST)
::        -> wie %7 bei Run_Laden_FAKTEN.cmd (bestimmt den Log-Ordner)
::   %2 = Datenbankname (z.B. msi_dm_lst_sgbii)
::        -> wie %8 bei Run_Laden_FAKTEN.cmd
::
:: VERWENDUNG:
::   Per UC4-Job (Post-Process bei Abbruch oder eigener Job im Abbruch-Zweig):
::     Kill_Laden_FAKTEN.cmd <Verfahren> <Datenbankname>
::
::   Manuell (Test/Notfall):
::     Kill_Laden_FAKTEN.cmd FAKTEN msi_dm_lst_sgbii
::
:: EXITCODES:
::   0 = Prozess beendet ODER kein laufender Prozess gefunden (kein Fehler, damit der
::       UC4-Job nicht zusaetzlich fehlschlaegt)
::   1 = Pflichtparameter fehlt
::   4 = Prozess konnte nicht beendet werden
::
:: TECHNISCHE HINWEISE:
::   - Die Suche erfolgt per PowerShell (Get-CimInstance), da wmic in neueren
::     Windows-Versionen entfernt ist.
::   - taskkill /F /T beendet den Prozess inklusive aller Kindprozesse.
::   - Beim Beenden von DTExec werden die SQL-Verbindungen getrennt; SQL Server rollt
::     laufende Statements (z.B. SELECT INTO ..._LOADING) zurueck. Das kann dauern.
::   - Uebrig gebliebene *_LOADING-Tabellen muessen vor dem naechsten Lauf bereinigt werden.
::   - Oracle-Sessions enden normalerweise mit dem Trennen der Verbindung - ggf. pruefen.
::   - Lockfile ist pro Verfahren+Datenbank (identisch zu Run_Laden_FAKTEN.cmd).
:: ============================================================================================


:: ============================================================================================
:: ABSCHNITT 1 - PARAMETER
:: ============================================================================================

set ATOMIC_VERFAHREN=%~1
set ATOMIC_DATENBANKNAME=%~2

:: Sonderfall FST (identisch zu Run_Laden_FAKTEN.cmd)
if /i "!ATOMIC_DATENBANKNAME!"=="msi_fst" set ATOMIC_DATENBANKNAME=FST


:: ============================================================================================
:: ABSCHNITT 2 - STATISCHE KONFIGURATION
:: ============================================================================================

set LOG_DATE=%DATE:~6,4%%DATE:~3,2%%DATE:~0,2%
set LOG_TIME_RAW=%TIME: =0%
set LOG_TIME=%LOG_TIME_RAW:~0,2%%LOG_TIME_RAW:~3,2%%LOG_TIME_RAW:~6,2%

set LOG_FILENAME=Kill_Laden_FAKTEN_!ATOMIC_DATENBANKNAME!_!LOG_DATE!_!LOG_TIME!.log
set LOG_DIR=\\!COMPUTERNAME!\Batches\!ATOMIC_VERFAHREN!\prog\log
set LOGFILE=!LOG_DIR!\!LOG_FILENAME!

set LOCKFILE_DIR=E:\BDATA\Batches\Ladepakete\log
set LOCKFILE=!LOCKFILE_DIR!\SSIS_FAKTEN_!ATOMIC_VERFAHREN!_!ATOMIC_DATENBANKNAME!.lock


:: ============================================================================================
:: ABSCHNITT 3 - VORABPRUEFUNG DER PARAMETER
:: ============================================================================================

if "!ATOMIC_VERFAHREN!"=="" (
    echo [%DATE% %TIME%] [FEHLEND]  ATOMIC_VERFAHREN ^(%%1^)
    exit /b 1
)

if "!ATOMIC_DATENBANKNAME!"=="" (
    echo [%DATE% %TIME%] [FEHLEND]  ATOMIC_DATENBANKNAME ^(%%2^)
    exit /b 1
)

if not exist "!LOG_DIR!" mkdir "!LOG_DIR!" 2>nul

echo [%DATE% %TIME%] =============================================== > "!LOGFILE!"
echo [%DATE% %TIME%] SSIS Faktenladen - Prozess beenden             >> "!LOGFILE!"
echo [%DATE% %TIME%] Verfahren   : !ATOMIC_VERFAHREN!               >> "!LOGFILE!"
echo [%DATE% %TIME%] Datenbank   : !ATOMIC_DATENBANKNAME!           >> "!LOGFILE!"
echo [%DATE% %TIME%] Logdatei    : !LOG_FILENAME!                   >> "!LOGFILE!"
echo [%DATE% %TIME%] =============================================== >> "!LOGFILE!"


:: ============================================================================================
:: ABSCHNITT 4 - DTEXEC-PROZESS SUCHEN UND BEENDEN
::
:: Gleiche Kriterien wie die Blockierungspruefung in Run_Laden_FAKTEN.cmd:
::   DTExec.exe, Kommandozeile enthaelt Laden_FAKTEN.dtsx und BA::Datenbank = Datenbankname
:: ============================================================================================

set KILL_DB=!ATOMIC_DATENBANKNAME!
set FOUND=0
set KILL_FAILED=0

for /f "usebackq delims=" %%P in (`powershell -NoProfile -Command "Get-CimInstance Win32_Process -Filter ('Name=' + [char]39 + 'DTExec.exe' + [char]39) | Where-Object { $_.CommandLine -and $_.CommandLine.Contains('Laden_FAKTEN.dtsx') -and $_.CommandLine.Contains('BA::Datenbank].Value;' + [char]34 + $env:KILL_DB + [char]34) } | ForEach-Object { $_.ProcessId }"`) do (
    set FOUND=1
    echo [%DATE% %TIME%] [INFO]     DTExec gefunden - Prozess-ID: %%P >> "!LOGFILE!"
    taskkill /F /T /PID %%P >> "!LOGFILE!" 2>&1
    if errorlevel 1 (
        echo [%DATE% %TIME%] [FEHLER]   Prozess %%P konnte nicht beendet werden. >> "!LOGFILE!"
        set KILL_FAILED=1
    ) else (
        echo [%DATE% %TIME%] [OK]       Prozess %%P beendet. >> "!LOGFILE!"
    )
)

if !FOUND!==0 (
    echo [%DATE% %TIME%] [INFO]     Kein laufender DTExec-Prozess fuer !ATOMIC_VERFAHREN! / !ATOMIC_DATENBANKNAME! gefunden. >> "!LOGFILE!"
)


:: ============================================================================================
:: ABSCHNITT 5 - LOCKFILE LOESCHEN UND ABSCHLUSS
:: ============================================================================================

if !KILL_FAILED!==1 (
    echo [%DATE% %TIME%] [ABGEBROCHEN] Lockfile bleibt bestehen: !LOCKFILE! >> "!LOGFILE!"
    echo [%DATE% %TIME%] Exitcode: 4 >> "!LOGFILE!"
    type "!LOGFILE!"
    exit /b 4
)

if exist "!LOCKFILE!" (
    del "!LOCKFILE!" 2>nul
    echo [%DATE% %TIME%] [OK]       Lockfile geloescht: !LOCKFILE! >> "!LOGFILE!"
) else (
    echo [%DATE% %TIME%] [INFO]     Kein Lockfile vorhanden: !LOCKFILE! >> "!LOGFILE!"
)

echo [%DATE% %TIME%] =============================================== >> "!LOGFILE!"
echo [%DATE% %TIME%] Exitcode: 0 >> "!LOGFILE!"

type "!LOGFILE!"

exit /b 0
