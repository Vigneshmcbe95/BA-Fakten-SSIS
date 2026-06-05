Option Explicit On
Option Strict On

Imports System
Imports System.Data.SqlClient
Imports System.Text
Imports Microsoft.SqlServer.Dts.Runtime

' =============================================================================
'  Script   : SCR01_Parameter_Einlesen
'  Package  : Fakten Laden (SSIS)
'  Purpose  : Reads the package parameters from the parameter table into the
'             SSIS variables (BA::*).
'  Logging  : SSIS events only (FireInformation / FireError)
' =============================================================================
<Microsoft.SqlServer.Dts.Tasks.ScriptTask.SSISScriptTaskEntryPointAttribute()>
<CLSCompliant(False)>
Partial Public Class ScriptMain
    Inherits Microsoft.SqlServer.Dts.Tasks.ScriptTask.VSTARTScriptObjectModelBase

    ' -------------------------------------------------------------------------
    ' Konstanten
    ' -------------------------------------------------------------------------
    Private Const SKRIPT_NAME As String = "SCR01_Parameter_Einlesen"

    ' -----------------------------------------------------------------------
    ' Main - Entry point - orchestrates the script flow.
    ' -----------------------------------------------------------------------
    Public Sub Main()

        Log("SCR01_Parameter_Einlesen - Start")

        Dim validierungOK As Boolean = True

        ' ─────────────────────────────────────────────────────────
        ' ABSCHNITT 1 – Verbindung / Server
        ' ─────────────────────────────────────────────────────────
        Log("Verbindung / Server")
        validierungOK = PruefeParam("BA::Server", "PARAM_SERVER", False) And validierungOK
        validierungOK = PruefeParam("BA::ConnectionServerName", "PARAM_CONNECTION_SERVER", False) And validierungOK
        validierungOK = PruefeParam("BA::Datenbank", "PARAM_DATENBANK", False) And validierungOK
        validierungOK = PruefeParam("BA::Verfahren", "PARAM_VERFAHREN", False) And validierungOK
        validierungOK = PruefeParam("BA::Datamart", "PARAM_DATAMART", False) And validierungOK

        ' ─────────────────────────────────────────────────────────
        ' ABSCHNITT 2 – Oracle Credentials
        ' ─────────────────────────────────────────────────────────
        Log("Oracle Credentials")
        validierungOK = PruefeParam("BA::CredBenutzername", "ATOMIC_ORACLE_USERNAME", False) And validierungOK
        validierungOK = PruefeParam("BA::CredKennwort", "ATOMIC_ORACLE_PASSWORD", True) And validierungOK  ' Wert verborgen

        ' ─────────────────────────────────────────────────────────
        ' ABSCHNITT 3 – Steuerlisten
        ' ─────────────────────────────────────────────────────────
        Log("Steuerlisten")
        validierungOK = PruefeParam("BA::STLOrdner", "PARAM_STL_ORDNER", False) And validierungOK
        validierungOK = PruefeParam("BA::SteuerlistenTabelle", "PARAM_STEUERLISTEN_TABELLE", False) And validierungOK
        'validierungOK = PruefeParam("BA::PartitionCacheTabelle", "PARAM_PARTITION_CACHE_TAB", False) And validierungOK

        ' ─────────────────────────────────────────────────────────
        ' ABSCHNITT 4 – Parametertabelle
        ' ─────────────────────────────────────────────────────────
        Log("Parametertabelle")
        validierungOK = PruefeParam("BA::ParameterDB", "PARAM_PARAMETER_DB", False) And validierungOK
        validierungOK = PruefeParam("BA::Parametertabelle", "PARAM_PARAMETERTABELLE", False) And validierungOK

        ' ─────────────────────────────────────────────────────────
        ' ABSCHNITT 5 – PolyBase / External Source
        ' ─────────────────────────────────────────────────────────
        Log("PolyBase / External Source")
        validierungOK = PruefeParam("BA::ExtSourceName", "PARAM_EXT_SOURCE_NAME", False) And validierungOK
        validierungOK = PruefeParam("BA::ExtSourceLocation", "PARAM_EXT_SOURCE_LOCATION", False) And validierungOK
        validierungOK = PruefeParam("BA::ExtTableLocation", "PARAM_EXT_TABLE_LOCATION", False) And validierungOK
        validierungOK = PruefeParam("BA::ExtTableSchema", "PARAM_EXT_TABLE_SCHEMA", False) And validierungOK
        validierungOK = PruefeParam("BA::ExtTableName", "PARAM_EXT_TABLE_NAME", False) And validierungOK

        ' ─────────────────────────────────────────────────────────
        ' ABSCHNITT 6 – Verarbeitung
        ' ─────────────────────────────────────────────────────────
        Log("Verarbeitung")
        validierungOK = PruefeParam("BA::Maxparallel", "PARAM_MAXPARALLEL", False) And validierungOK

        ' ─────────────────────────────────────────────────────────
        ' ABSCHNITT 7 – Protokollierung
        ' ─────────────────────────────────────────────────────────
        Log("Protokollierung")
        validierungOK = PruefeParam("BA::ProtokollDB", "PARAM_PROTOKOLL_DB", False) And validierungOK
        validierungOK = PruefeParam("BA::ProtokollSP", "PARAM_PROTOKOLL_SP", False) And validierungOK
        validierungOK = PruefeParam("BA::Protokolltabelle", "PARAM_PROTOKOLL_TABELLE", False) And validierungOK

        ' ─────────────────────────────────────────────────────────
        ' ERGEBNIS
        ' ─────────────────────────────────────────────────────────
        If Not validierungOK Then
            Log("VALIDATION_FAILED = 1")
            Log("[ABGEBROCHEN] Pflichtparameter fehlen Paket wird nicht gestartet.")
            LogFehler("SCR01: Pflichtparameter fehlen Paket abgebrochen.")
            Dts.TaskResult = ScriptResults.Failure
        Else
            Log("VALIDATION_FAILED = 0")
            Log("VORABPRUEFUNG BESTANDEN Paket wird fortgesetzt.")
            Dts.TaskResult = ScriptResults.Success
        End If

    End Sub

    ' -----------------------------------------------------------------------
    ' PruefeParam - Checks a single parameter for existence and value.
    ' -----------------------------------------------------------------------
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
            Log(String.Format("[FEHLEND]  {0,-40} ({1}) Variable nicht gefunden: {2}",
                paramLabel, variableName, ex.Message))
            Return False
        End Try
    End Function

    ' -----------------------------------------------------------------------
    ' LeseVariable - Reads a single SSIS variable (with fallback).
    ' -----------------------------------------------------------------------
    Private Function LeseVariable(name As String) As String
        Try
            Return Dts.Variables(name).Value.ToString().Trim()
        Catch
            Return String.Empty
        End Try
    End Function

    ' -----------------------------------------------------------------------
    ' Log - Writes an information message to the SSIS log
    ' (FireInformation).
    ' -----------------------------------------------------------------------
    Private Sub Log(nachricht As String)
        Dim fireAgain As Boolean = False
        Dts.Events.FireInformation(0, SKRIPT_NAME, nachricht, "", 0, fireAgain)
    End Sub

    ' -----------------------------------------------------------------------
    ' LogFehler - Writes an error message to the SSIS log (FireError).
    ' -----------------------------------------------------------------------
    Private Sub LogFehler(nachricht As String)
        Dts.Events.FireError(0, SKRIPT_NAME, nachricht, "", 0)
    End Sub

    ' -----------------------------------------------------------------------
    ' ScriptResults - SSIS task result codes.
    ' -----------------------------------------------------------------------
    Public Enum ScriptResults
        Success = DTSExecResult.Success
        Failure = DTSExecResult.Failure
    End Enum

End Class
