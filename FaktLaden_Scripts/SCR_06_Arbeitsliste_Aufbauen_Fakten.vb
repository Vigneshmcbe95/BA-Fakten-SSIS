Option Explicit On
Option Strict On

Imports System
Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports Microsoft.SqlServer.Dts.Runtime

' =============================================================================
' PAKET  : Fakten Laden
' SKRIPT : SCR_Arbeitsliste_Aufbauen_Fakten
' ZWECK  : 1. Neue Verfahren aus Steuerliste in ETL_Fkt_Arbeitsliste eintragen
'          2. Basierend auf vorhandenen Status zurücksetzen
'          3. KEINE RunID-Filterung für Reset (weil alte Einträge NULL haben)
'          4. KEINE where_klausel/partition_wert Prüfung - nur direkter JOIN
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

    Public Sub Main()

        Try
            Log("════════════════════════════════════════════════════════")
            Log("SCR_Arbeitsliste_Aufbauen_Fakten – Start")
            Log("Zeitpunkt: " & DateTime.Now.ToString("dd.MM.yyyy HH:mm:ss"))
            Log("════════════════════════════════════════════════════════")

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
            Log("── Schritt 1: Neue Verfahren aus Steuerliste eintragen")
            Dim neueVerfahren As Integer = NeueVerfahrenEintragen()
            Log("Neue Verfahren eingetragen: " & neueVerfahren.ToString())

            ' ┌─────────────────────────────────────────────────────────┐
            ' │ SCHRITT 2: Status analysieren und zurücksetzen         │
            ' │           KEINE RunID-Filterung!                       │
            ' │           Betrachtet ALLE Verfahren aus Steuerliste     │
            ' └─────────────────────────────────────────────────────────┘
            Log("── Schritt 2: Status analysieren und zurücksetzen")
            Dim resetResult As String = StatusAnalysierenUndZuruecksetzen()
            Log(resetResult)

            ' ┌─────────────────────────────────────────────────────────┐
            ' │ SCHRITT 3: Aktuelle RunID für alle Verfahren setzen     │
            ' └─────────────────────────────────────────────────────────┘
            Log("── Schritt 3: RunID für alle Verfahren aktualisieren")
            Dim updatedRunID As Integer = RunIDAktualisieren()
            Log("RunID aktualisiert für: " & updatedRunID.ToString() & " Verfahren")

            ' ┌─────────────────────────────────────────────────────────┐
            ' │ SCHRITT 4: Zusammenfassung                              │
            ' └─────────────────────────────────────────────────────────┘
            Log("════════════════════════════════════════════════════════")
            Log("ZUSAMMENFASSUNG")
            Log("════════════════════════════════════════════════════════")
            Log("Neue Verfahren : " & neueVerfahren.ToString())
            Log("RunID          : " & _runID.ToString())
            Log("════════════════════════════════════════════════════════")
            Log("SCR_Arbeitsliste_Aufbauen_Fakten erfolgreich abgeschlossen ✓")

            Dts.TaskResult = ScriptResults.Success

        Catch ex As Exception
            LogFehler("KRITISCHER FEHLER: " & ex.Message)
            LogFehler("Stack Trace: " & ex.StackTrace)
            Dts.TaskResult = ScriptResults.Failure
        End Try

    End Sub

    ' =========================================================================
    ' Variablen aus SSIS laden
    ' =========================================================================
    Private Sub VariablenLaden()
        _runID = Convert.ToInt32(Dts.Variables("BA::RunID").Value)
        _parameterDB = Dts.Variables("BA::ParameterDB").Value.ToString().Trim()
        _parametertabelle = Dts.Variables("BA::Parametertabelle").Value.ToString().Trim()
        _steuerlistenTabelle = Dts.Variables("BA::SteuerlistenTabelle").Value.ToString().Trim()

        Log("────────────────────────────────────────────────────────")
        Log("RunID                : " & _runID.ToString())
        Log("ParameterDB          : " & _parameterDB)
        Log("Parametertabelle     : " & _parametertabelle)
        Log("Steuerlisten-Tabelle : " & _steuerlistenTabelle)
        Log("────────────────────────────────────────────────────────")
    End Sub

    ' =========================================================================
    ' Pflichtfelder prüfen
    ' =========================================================================
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
        Log("Pflichtfelder-Prüfung: alle Variablen vorhanden ✓")
        Return True
    End Function

    ' =========================================================================
    ' NEUE VERFAHREN EINTRAGEN
    ' =========================================================================
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

    ' =========================================================================
    ' STATUS ANALYSIEREN UND ZURÜCKSETZEN
    ' Analysiert ALLE Verfahren aus der Steuerliste (nicht nach RunID)
    ' =========================================================================
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
            LogFehler("FEHLER beim Zurücksetzen: " & ex.Message)
            Throw New Exception("Reset fehlgeschlagen: " & ex.Message, ex)
        End Try

        Return resetReason & " → " & rowsAffected.ToString() & " Zeilen zurückgesetzt"

    End Function

    ' =========================================================================
    ' RUNID FÜR ALLE VERFAHREN AKTUALISIEREN
    ' Setzt die aktuelle RunID für alle Verfahren aus der Steuerliste
    ' =========================================================================
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

    ' =========================================================================
    ' Hilfsfunktionen
    ' =========================================================================
    Private Function HoleVerbindungszeichenfolge() As String
        Return Dts.Connections(CONN_NAME).ConnectionString
    End Function

    Private Sub Log(nachricht As String)
        Dim fireAgain As Boolean = False
        Dts.Events.FireInformation(0, SKRIPT_NAME, nachricht, "", 0, fireAgain)
    End Sub

    Private Sub LogFehler(nachricht As String)
        Dts.Events.FireError(0, SKRIPT_NAME, nachricht, "", 0)
    End Sub

    Public Enum ScriptResults
        Success = DTSExecResult.Success
        Failure = DTSExecResult.Failure
    End Enum

End Class
