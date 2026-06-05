Option Explicit On
Option Strict On

Imports System
Imports System.Data.SqlClient
Imports System.Text
Imports Microsoft.SqlServer.Dts.Runtime

' =============================================================================
' PAKET  : Fakten Laden
' SKRIPT : SCR01_Parameter_Einlesen
' ZWECK  : 1. Alle SSIS-Variablen lesen (werden per CMD /SET Ã¼bergeben)
'          2. Jeden Parameter einzeln prÃ¼fen â [OK] oder [FEHLEND]
'          3. Bei fehlendem Pflichtparameter â Paket abbrechen
' =============================================================================
<Microsoft.SqlServer.Dts.Tasks.ScriptTask.SSISScriptTaskEntryPointAttribute()>
<CLSCompliant(False)>
Partial Public Class ScriptMain
    Inherits Microsoft.SqlServer.Dts.Tasks.ScriptTask.VSTARTScriptObjectModelBase

    ' -------------------------------------------------------------------------
    ' Konstanten
    ' -------------------------------------------------------------------------
    Private Const SKRIPT_NAME As String = "SCR01_Parameter_Einlesen"

    ' -------------------------------------------------------------------------
    ' Einstiegspunkt
    ' -------------------------------------------------------------------------
    Public Sub Main()

        Log("ââââââââââââââââââââââââââââââââââââââââââââââââââââââââ")
        Log("SCR01_Parameter_Einlesen â Start")
        Log("Zeitpunkt: " & DateTime.Now.ToString("dd.MM.yyyy HH:mm:ss"))
        Log("ââââââââââââââââââââââââââââââââââââââââââââââââââââââââ")

        Dim validierungOK As Boolean = True

        ' âââââââââââââââââââââââââââââââââââââââââââââââââââââââââ
        ' ABSCHNITT 1 â Verbindung / Server
        ' âââââââââââââââââââââââââââââââââââââââââââââââââââââââââ
        Log("ââ Verbindung / Server")
        validierungOK = PruefeParam("BA::Server", "PARAM_SERVER", False) And validierungOK
        validierungOK = PruefeParam("BA::ConnectionServerName", "PARAM_CONNECTION_SERVER", False) And validierungOK
        validierungOK = PruefeParam("BA::Datenbank", "PARAM_DATENBANK", False) And validierungOK
        validierungOK = PruefeParam("BA::Verfahren", "PARAM_VERFAHREN", False) And validierungOK
        validierungOK = PruefeParam("BA::Datamart", "PARAM_DATAMART", False) And validierungOK

        ' âââââââââââââââââââââââââââââââââââââââââââââââââââââââââ
        ' ABSCHNITT 2 â Oracle Credentials
        ' âââââââââââââââââââââââââââââââââââââââââââââââââââââââââ
        Log("ââ Oracle Credentials")
        validierungOK = PruefeParam("BA::CredBenutzername", "ATOMIC_ORACLE_USERNAME", False) And validierungOK
        validierungOK = PruefeParam("BA::CredKennwort", "ATOMIC_ORACLE_PASSWORD", True) And validierungOK  ' Wert verborgen

        ' âââââââââââââââââââââââââââââââââââââââââââââââââââââââââ
        ' ABSCHNITT 3 â Steuerlisten
        ' âââââââââââââââââââââââââââââââââââââââââââââââââââââââââ
        Log("ââ Steuerlisten")
        validierungOK = PruefeParam("BA::STLOrdner", "PARAM_STL_ORDNER", False) And validierungOK
        validierungOK = PruefeParam("BA::SteuerlistenTabelle", "PARAM_STEUERLISTEN_TABELLE", False) And validierungOK
        'validierungOK = PruefeParam("BA::PartitionCacheTabelle", "PARAM_PARTITION_CACHE_TAB", False) And validierungOK

        ' âââââââââââââââââââââââââââââââââââââââââââââââââââââââââ
        ' ABSCHNITT 4 â Parametertabelle
        ' âââââââââââââââââââââââââââââââââââââââââââââââââââââââââ
        Log("ââ Parametertabelle")
        validierungOK = PruefeParam("BA::ParameterDB", "PARAM_PARAMETER_DB", False) And validierungOK
        validierungOK = PruefeParam("BA::Parametertabelle", "PARAM_PARAMETERTABELLE", False) And validierungOK

        ' âââââââââââââââââââââââââââââââââââââââââââââââââââââââââ
        ' ABSCHNITT 5 â PolyBase / External Source
        ' âââââââââââââââââââââââââââââââââââââââââââââââââââââââââ
        Log("ââ PolyBase / External Source")
        validierungOK = PruefeParam("BA::ExtSourceName", "PARAM_EXT_SOURCE_NAME", False) And validierungOK
        validierungOK = PruefeParam("BA::ExtSourceLocation", "PARAM_EXT_SOURCE_LOCATION", False) And validierungOK
        validierungOK = PruefeParam("BA::ExtTableLocation", "PARAM_EXT_TABLE_LOCATION", False) And validierungOK
        validierungOK = PruefeParam("BA::ExtTableSchema", "PARAM_EXT_TABLE_SCHEMA", False) And validierungOK
        validierungOK = PruefeParam("BA::ExtTableName", "PARAM_EXT_TABLE_NAME", False) And validierungOK

        ' âââââââââââââââââââââââââââââââââââââââââââââââââââââââââ
        ' ABSCHNITT 6 â Verarbeitung
        ' âââââââââââââââââââââââââââââââââââââââââââââââââââââââââ
        Log("ââ Verarbeitung")
        validierungOK = PruefeParam("BA::Maxparallel", "PARAM_MAXPARALLEL", False) And validierungOK

        ' âââââââââââââââââââââââââââââââââââââââââââââââââââââââââ
        ' ABSCHNITT 7 â Protokollierung
        ' âââââââââââââââââââââââââââââââââââââââââââââââââââââââââ
        Log("ââ Protokollierung")
        validierungOK = PruefeParam("BA::ProtokollDB", "PARAM_PROTOKOLL_DB", False) And validierungOK
        validierungOK = PruefeParam("BA::ProtokollSP", "PARAM_PROTOKOLL_SP", False) And validierungOK
        validierungOK = PruefeParam("BA::Protokolltabelle", "PARAM_PROTOKOLL_TABELLE", False) And validierungOK

        ' âââââââââââââââââââââââââââââââââââââââââââââââââââââââââ
        ' ERGEBNIS
        ' âââââââââââââââââââââââââââââââââââââââââââââââââââââââââ
        Log("ââââââââââââââââââââââââââââââââââââââââââââââââââââââââ")
        If Not validierungOK Then
            Log("VALIDATION_FAILED = 1")
            Log("[ABGEBROCHEN] Pflichtparameter fehlen â Paket wird nicht gestartet.")
            Log("ââââââââââââââââââââââââââââââââââââââââââââââââââââââââ")
            LogFehler("SCR01: Pflichtparameter fehlen â Paket abgebrochen.")
            Dts.TaskResult = ScriptResults.Failure
        Else
            Log("VALIDATION_FAILED = 0")
            Log("VORABPRUEFUNG BESTANDEN â Paket wird fortgesetzt.")
            Log("ââââââââââââââââââââââââââââââââââââââââââââââââââââââââ")
            Dts.TaskResult = ScriptResults.Success
        End If

    End Sub

    ' =========================================================================
    ' Einzelnen Parameter prÃ¼fen und loggen
    ' variableName  = SSIS-Variablenname (z.B. "BA::Server")
    ' paramLabel    = Anzeigename im Log  (z.B. "PARAM_SERVER")
    ' istPasswort   = True â Wert wird als [Wert verborgen] geloggt
    ' =========================================================================
    Private Function PruefeParam(variableName As String,
                                  paramLabel As String,
                                  istPasswort As Boolean) As Boolean
        Try
            Dim wert As String = Dts.Variables(variableName).Value.ToString().Trim()

            If String.IsNullOrEmpty(wert) Then
                Log(String.Format("[FEHLEND]  {0,-40} ({1})", paramLabel, variableName))
                Return False
            End If

            If istPasswort Then
                Log(String.Format("[OK]       {0,-40} = [Wert verborgen]", paramLabel))
            Else
                Log(String.Format("[OK]       {0,-40} = {1}", paramLabel, wert))
            End If

            Return True

        Catch ex As Exception
            Log(String.Format("[FEHLEND]  {0,-40} ({1}) â Variable nicht gefunden: {2}",
                paramLabel, variableName, ex.Message))
            Return False
        End Try
    End Function

    ' =========================================================================
    ' Variable sicher lesen â leerer String bei Fehler
    ' =========================================================================
    Private Function LeseVariable(name As String) As String
        Try
            Return Dts.Variables(name).Value.ToString().Trim()
        Catch
            Return String.Empty
        End Try
    End Function

    ' =========================================================================
    ' Logging
    ' =========================================================================
    Private Sub Log(nachricht As String)
        Dim fireAgain As Boolean = False
        Dts.Events.FireInformation(0, SKRIPT_NAME, nachricht, "", 0, fireAgain)
    End Sub

    Private Sub LogFehler(nachricht As String)
        Dts.Events.FireError(0, SKRIPT_NAME, nachricht, "", 0)
    End Sub

    ' =========================================================================
    ' Ergebnistypen
    ' =========================================================================
    Public Enum ScriptResults
        Success = DTSExecResult.Success
        Failure = DTSExecResult.Failure
    End Enum

End Class
