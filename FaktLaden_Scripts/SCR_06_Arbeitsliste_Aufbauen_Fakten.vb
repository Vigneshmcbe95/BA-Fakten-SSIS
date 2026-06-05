Option Explicit On
Option Strict On

Imports System
Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports Microsoft.SqlServer.Dts.Runtime

' =============================================================================
'  Script   : SCR_06_Arbeitsliste_Aufbauen_Fakten
'  Package  : Fakten Laden (SSIS)
'  Purpose  : Builds the work list for this run: inserts new Verfahren,
'             resets FEHLER rows to AUSSTEHEND and stamps the current RunID.
'  Logging  : SSIS events only (FireInformation / FireError)
' =============================================================================
<Microsoft.SqlServer.Dts.Tasks.ScriptTask.SSISScriptTaskEntryPointAttribute()>
<CLSCompliant(False)>
Partial Public Class ScriptMain
    Inherits Microsoft.SqlServer.Dts.Tasks.ScriptTask.VSTARTScriptObjectModelBase

    Private Const SKRIPT_NAME As String = "SCR_06_Arbeitsliste_Aufbauen_Fakten"
    Private Const CONN_NAME As String = "Verbindung"

    Private _runID As Integer = 0
    Private _parameterDB As String = String.Empty
    Private _parametertabelle As String = String.Empty
    Private _steuerlistenTabelle As String = String.Empty
    Private _connectionString As String = String.Empty

    ' -----------------------------------------------------------------------
    ' Main - Entry point - orchestrates the script flow.
    ' -----------------------------------------------------------------------
    Public Sub Main()

        Try
        Log("SCR_06_Arbeitsliste_Aufbauen_Fakten - Start")

            ' Variablen laden
            VariablenLaden()

            ' Pflichtfelder prüfen
            If Not PflichtfelderPruefen() Then
                Dts.TaskResult = ScriptResults.Failure
                Return
            End If

            _connectionString = HoleVerbindungszeichenfolge()

            ' ┌─────────────────────────────────────────────────────────┐
            ' │ SCHRITT 1: Neue Verfahren aus Steuerliste eintragen     │
            ' └─────────────────────────────────────────────────────────┘
            Log("Schritt 1: Neue Verfahren aus Steuerliste eintragen")
            Dim neueVerfahren As Integer = NeueVerfahrenEintragen()
            Log("Neue Verfahren eingetragen: " & neueVerfahren.ToString())

            ' ┌─────────────────────────────────────────────────────────┐
            ' │ SCHRITT 2: Status analysieren und zurücksetzen         │
            ' │           KEINE RunID-Filterung!                       │
            ' │           Betrachtet ALLE Verfahren aus Steuerliste     │
            ' └─────────────────────────────────────────────────────────┘
            Log("Schritt 2: Status analysieren und zuruecksetzen")
            Dim resetResult As String = StatusAnalysierenUndZuruecksetzen()
            Log(resetResult)

            ' ┌─────────────────────────────────────────────────────────┐
            ' │ SCHRITT 3: Aktuelle RunID für alle Verfahren setzen     │
            ' └─────────────────────────────────────────────────────────┘
            Dim updatedRunID As Integer = RunIDAktualisieren()

            ' ┌─────────────────────────────────────────────────────────┐
            ' │ SCHRITT 4: Zusammenfassung                              │
            ' └─────────────────────────────────────────────────────────┘
            Log("ZUSAMMENFASSUNG")
            Log("Neue Verfahren : " & neueVerfahren.ToString())
            Log("SCR_Arbeitsliste_Aufbauen_Fakten erfolgreich abgeschlossen OK")

            Dts.TaskResult = ScriptResults.Success

        Catch ex As Exception
            LogFehler("KRITISCHER FEHLER: " & ex.Message)
            LogFehler("Stack Trace: " & ex.StackTrace)
            Dts.TaskResult = ScriptResults.Failure
        End Try

    End Sub

    ' -----------------------------------------------------------------------
    ' VariablenLaden - Reads the required SSIS variables into module
    ' fields.
    ' -----------------------------------------------------------------------
    Private Sub VariablenLaden()
        _runID = Convert.ToInt32(Dts.Variables("BA::RunID").Value)
        _parameterDB = Dts.Variables("BA::ParameterDB").Value.ToString().Trim()
        _parametertabelle = Dts.Variables("BA::Parametertabelle").Value.ToString().Trim()
        _steuerlistenTabelle = Dts.Variables("BA::SteuerlistenTabelle").Value.ToString().Trim()

    End Sub

    ' -----------------------------------------------------------------------
    ' PflichtfelderPruefen - Validates that all mandatory variables /
    ' parameters are present.
    ' -----------------------------------------------------------------------
    Private Function PflichtfelderPruefen() As Boolean
        Dim fehlend As New System.Text.StringBuilder()
        If _runID <= 0 Then fehlend.AppendLine("  → BA::RunID (ungültig)")
        If String.IsNullOrEmpty(_parameterDB) Then fehlend.AppendLine("  → BA::ParameterDB")
        If String.IsNullOrEmpty(_parametertabelle) Then fehlend.AppendLine("  → BA::Parametertabelle")
        If String.IsNullOrEmpty(_steuerlistenTabelle) Then fehlend.AppendLine("  → BA::SteuerlistenTabelle")
        If fehlend.Length > 0 Then
            LogFehler("Pflichtfelder fehlen:" & Environment.NewLine & fehlend.ToString())
            Return False
        End If
        Log("Pflichtfelder-Pruefung: alle Variablen vorhanden OK")
        Return True
    End Function

    ' -----------------------------------------------------------------------
    ' NeueVerfahrenEintragen - Inserts new Verfahren from the control list
    ' into the work list.
    ' -----------------------------------------------------------------------
    Private Function NeueVerfahrenEintragen() As Integer

        Dim sql As String = "
INSERT INTO dbo.ETL_Fkt_Arbeitsliste
       (RunID, Verfahren, Themengebiet, Status, LetzterSchritt, Versuche, AktualisiertAm)
SELECT DISTINCT
    @runID,
    LOWER(p.Verfahren),
    LOWER(f.themengebiet),
    'AUSSTEHEND',
    NULL,
    0,
    GETDATE()
FROM " & _parameterDB & ".dbo." & _parametertabelle & " p
INNER JOIN dbo." & _steuerlistenTabelle & " f
    ON LOWER(LTRIM(RTRIM(f.tabelle))) = LOWER(LTRIM(RTRIM(p.Verfahren)))
WHERE p.Verfahren IS NOT NULL
  AND NOT EXISTS (
      SELECT 1 FROM dbo.ETL_Fkt_Arbeitsliste a
      WHERE a.Verfahren = LOWER(p.Verfahren)
  );"

        Try
            Using conn As New SqlConnection(_connectionString)
                conn.Open()
                Using cmd As New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@runID", _runID)
                    cmd.CommandTimeout = 0
                    Dim rows As Integer = cmd.ExecuteNonQuery()
                    Return rows
                End Using
            End Using
        Catch ex As Exception
            LogFehler("FEHLER beim Eintragen neuer Verfahren: " & ex.Message)
            Throw New Exception("Neue Verfahren konnten nicht eingetragen werden: " & ex.Message, ex)
        End Try

    End Function

    ' -----------------------------------------------------------------------
    ' StatusAnalysierenUndZuruecksetzen - Analyzes the work list status and
    ' resets rows for the new run.
    ' -----------------------------------------------------------------------
    Private Function StatusAnalysierenUndZuruecksetzen() As String

        ' 1. Alle relevanten Verfahren und deren Status aus der Steuerliste holen
        Dim sqlAnalyse As String = "
SELECT
    SUM(CASE WHEN a.Status = 'ERFOLG' THEN 1 ELSE 0 END) AS CountErfolg,
    SUM(CASE WHEN a.Status = 'FEHLER' THEN 1 ELSE 0 END) AS CountFehler,
    SUM(CASE WHEN a.Status NOT IN ('ERFOLG', 'FEHLER', 'AUSSTEHEND') AND a.Status IS NOT NULL THEN 1 ELSE 0 END) AS CountOther,
    COUNT(*) AS Total
FROM " & _parameterDB & ".dbo." & _parametertabelle & " p
INNER JOIN dbo." & _steuerlistenTabelle & " f
    ON LOWER(LTRIM(RTRIM(f.tabelle))) = LOWER(LTRIM(RTRIM(p.Verfahren)))
LEFT JOIN dbo.ETL_Fkt_Arbeitsliste a
    ON a.Verfahren = LOWER(p.Verfahren)
WHERE p.Verfahren IS NOT NULL;"

        Dim countErfolg As Integer = 0
        Dim countFehler As Integer = 0
        Dim countOther As Integer = 0
        Dim total As Integer = 0

        Try
            Using conn As New SqlConnection(_connectionString)
                conn.Open()
                Using cmd As New SqlCommand(sqlAnalyse, conn)
                    cmd.CommandTimeout = 0
                    Using reader As SqlDataReader = cmd.ExecuteReader()
                        If reader.Read() Then
                            countErfolg = If(reader("CountErfolg") Is DBNull.Value, 0, Convert.ToInt32(reader("CountErfolg")))
                            countFehler = If(reader("CountFehler") Is DBNull.Value, 0, Convert.ToInt32(reader("CountFehler")))
                            countOther = If(reader("CountOther") Is DBNull.Value, 0, Convert.ToInt32(reader("CountOther")))
                            total = If(reader("Total") Is DBNull.Value, 0, Convert.ToInt32(reader("Total")))
                        End If
                    End Using
                End Using
            End Using
        Catch ex As Exception
            LogFehler("FEHLER bei der Statusanalyse: " & ex.Message)
            Throw New Exception("Statusanalyse fehlgeschlagen: " & ex.Message, ex)
        End Try

        Log("Status-Analyse (alle Verfahren aus Steuerliste):")
        Log("  - Total    : " & total.ToString())
        Log("  - ERFOLG   : " & countErfolg.ToString())
        Log("  - FEHLER   : " & countFehler.ToString())
        Log("  - ANDERE   : " & countOther.ToString())

        If total = 0 Then
            Return "Keine Verfahren in Steuerliste gefunden."
        End If

        Dim rowsAffected As Integer = 0
        Dim resetSql As String = ""
        Dim resetReason As String = ""

        ' 2. Entscheidung basierend auf Analyse
        If countErfolg > 0 AndAlso countFehler = 0 AndAlso countOther = 0 Then
            ' Fall 1: NUR ERFOLG → Alle ERFOLG zu AUSSTEHEND
            resetSql = "
UPDATE a
SET a.Status = 'AUSSTEHEND',
    a.LetzterSchritt = NULL,
    a.Fehlermeldung = NULL,
    a.Versuche = 0,
    a.AktualisiertAm = GETDATE()
FROM dbo.ETL_Fkt_Arbeitsliste a
INNER JOIN " & _parameterDB & ".dbo." & _parametertabelle & " p ON LOWER(p.Verfahren) = a.Verfahren
INNER JOIN dbo." & _steuerlistenTabelle & " f ON LOWER(LTRIM(RTRIM(f.tabelle))) = LOWER(LTRIM(RTRIM(p.Verfahren)))
WHERE a.Status = 'ERFOLG'
  AND p.Verfahren IS NOT NULL;"
            resetReason = "Nur ERFOLG vorhanden → Alle ERFOLG zu AUSSTEHEND (kompletter Neulauf)"

        ElseIf countFehler > 0 Then
            ' Fall 2: FEHLER vorhanden → Nur FEHLER zu AUSSTEHEND
            resetSql = "
UPDATE a
SET a.Status = 'AUSSTEHEND',
    a.LetzterSchritt = NULL,
    a.Fehlermeldung = NULL,
    a.Versuche = 0,
    a.AktualisiertAm = GETDATE()
FROM dbo.ETL_Fkt_Arbeitsliste a
INNER JOIN " & _parameterDB & ".dbo." & _parametertabelle & " p ON LOWER(p.Verfahren) = a.Verfahren
INNER JOIN dbo." & _steuerlistenTabelle & " f ON LOWER(LTRIM(RTRIM(f.tabelle))) = LOWER(LTRIM(RTRIM(p.Verfahren)))
WHERE a.Status = 'FEHLER'
  AND p.Verfahren IS NOT NULL;"
            resetReason = "FEHLER vorhanden → Nur FEHLER zu AUSSTEHEND (Erfolge bleiben)"

        ElseIf countOther > 0 Then
            ' Fall 3: ANDERE Status (DATA_LOADING, SCHEMA_KOPIERT, etc.)
            resetSql = "
UPDATE a
SET a.Status = 'AUSSTEHEND',
    a.LetzterSchritt = NULL,
    a.Fehlermeldung = NULL,
    a.Versuche = 0,
    a.AktualisiertAm = GETDATE()
FROM dbo.ETL_Fkt_Arbeitsliste a
INNER JOIN " & _parameterDB & ".dbo." & _parametertabelle & " p ON LOWER(p.Verfahren) = a.Verfahren
INNER JOIN dbo." & _steuerlistenTabelle & " f ON LOWER(LTRIM(RTRIM(f.tabelle))) = LOWER(LTRIM(RTRIM(p.Verfahren)))
WHERE a.Status NOT IN ('ERFOLG', 'AUSSTEHEND')
  AND p.Verfahren IS NOT NULL;"
            resetReason = "ANDERE Status vorhanden → Alle NICHT-ERFOLG zu AUSSTEHEND"

        Else
            Return "Kein Reset nötig (nur AUSSTEHEND oder keine Daten)"
        End If

        ' 3. Reset durchführen
        Try
            Using conn As New SqlConnection(_connectionString)
                conn.Open()
                Using cmd As New SqlCommand(resetSql, conn)
                    cmd.CommandTimeout = 0
                    rowsAffected = cmd.ExecuteNonQuery()
                End Using
            End Using
        Catch ex As Exception
            LogFehler("FEHLER beim Zuruecksetzen: " & ex.Message)
            Throw New Exception("Reset fehlgeschlagen: " & ex.Message, ex)
        End Try

        Return resetReason & " → " & rowsAffected.ToString() & " Zeilen zurückgesetzt"

    End Function

    ' -----------------------------------------------------------------------
    ' RunIDAktualisieren - Stamps the current RunID on all Verfahren of
    ' this run.
    ' -----------------------------------------------------------------------
    Private Function RunIDAktualisieren() As Integer

        Dim sql As String = "
UPDATE a
SET a.RunID = @runID,
    a.AktualisiertAm = GETDATE()
FROM dbo.ETL_Fkt_Arbeitsliste a
INNER JOIN " & _parameterDB & ".dbo." & _parametertabelle & " p ON LOWER(p.Verfahren) = a.Verfahren
INNER JOIN dbo." & _steuerlistenTabelle & " f ON LOWER(LTRIM(RTRIM(f.tabelle))) = LOWER(LTRIM(RTRIM(p.Verfahren)))
WHERE p.Verfahren IS NOT NULL;"

        Try
            Using conn As New SqlConnection(_connectionString)
                conn.Open()
                Using cmd As New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@runID", _runID)
                    cmd.CommandTimeout = 0
                    Dim rows As Integer = cmd.ExecuteNonQuery()
                    Return rows
                End Using
            End Using
        Catch ex As Exception
            LogFehler("FEHLER beim Aktualisieren der RunID: " & ex.Message)
            Throw New Exception("RunID Aktualisierung fehlgeschlagen: " & ex.Message, ex)
        End Try

    End Function

    ' -----------------------------------------------------------------------
    ' HoleVerbindungszeichenfolge - Returns the connection string of the
    ' package connection manager.
    ' -----------------------------------------------------------------------
    Private Function HoleVerbindungszeichenfolge() As String
        Return Dts.Connections(CONN_NAME).ConnectionString
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
