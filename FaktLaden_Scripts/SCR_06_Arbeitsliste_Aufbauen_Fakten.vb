Option Explicit On
Option Strict On

Imports System
Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports Microsoft.SqlServer.Dts.Runtime

' =============================================================================
'  Skript       : SCR_06_Arbeitsliste_Aufbauen_Fakten
'  Paket        : Fakten Laden (SSIS)
'  Zweck        : Baut die Arbeitsliste fuer den Lauf auf: traegt neue
'                 Verfahren ein, setzt FEHLER-Zeilen auf AUSSTEHEND zurueck
'                 und aktualisiert die RunID.
'  Protokoll    : Nur SSIS-Events (FireInformation / FireError)
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
    Private _stlDateiname As String = String.Empty
    Private _connectionString As String = String.Empty

    ' -----------------------------------------------------------------------
    ' DateiFilter - Sicherheitsbedingung: nur Steuerlisten-Zeilen der
    ' aktuell geladenen STL-Datei (BA::STLDateiname). SCR03 leert die
    ' Arbeitstabelle vor jedem Lauf komplett, daher enthaelt sie ohnehin
    ' nur die aktuelle Datei - der Filter schuetzt zusaetzlich gegen
    ' Altbestaende, falls SCR03 nicht lief.
    ' -----------------------------------------------------------------------
    Private Function DateiFilter() As String
        Return " AND LOWER(LTRIM(RTRIM(f.FILE_NAME))) = '" & _stlDateiname.Trim().ToLower().Replace("'", "''") & "'"
    End Function

    ' -----------------------------------------------------------------------
    ' SteuerlisteLoggen - Protokolliert Dateiname, Anzahl und Tabellenliste
    ' des aktuellen Laufs aus der (von SCR03 frisch befuellten)
    ' Steuerlisten-Tabelle.
    ' -----------------------------------------------------------------------
    Private Sub SteuerlisteLoggen()
        Dim tabellen As New List(Of String)()
        Dim sql As String =
            "SELECT DISTINCT LOWER(LTRIM(RTRIM(f.tabelle))) FROM dbo." & _steuerlistenTabelle &
            " f WHERE 1=1" & DateiFilter() & " ORDER BY 1"
        Using conn As New SqlConnection(_connectionString)
            conn.Open()
            Using cmd As New SqlCommand(sql, conn)
                cmd.CommandTimeout = 0
                Using rdr As SqlDataReader = cmd.ExecuteReader()
                    While rdr.Read()
                        If Not rdr.IsDBNull(0) Then tabellen.Add(rdr.GetString(0))
                    End While
                End Using
            End Using
        End Using
        Log("Aktueller Lauf: Datei [" & _stlDateiname & "] | Tabellen in Steuerliste (" &
            tabellen.Count.ToString() & "): " & String.Join(", ", tabellen))
    End Sub

    ' -----------------------------------------------------------------------
    ' Main - Einstiegspunkt - steuert den Ablauf des Skripts.
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

            ' Geladene Datei + Tabellenliste des aktuellen Laufs protokollieren
            SteuerlisteLoggen()

            ' ┌─────────────────────────────────────────────────────────┐
            ' │ SCHRITT 1: Neue Verfahren aus Steuerliste eintragen     │
            ' └─────────────────────────────────────────────────────────┘
            Log("Schritt 1: Neue Verfahren aus Steuerliste eintragen")
            Dim neueVerfahren As Integer = NeueVerfahrenEintragen()
            Log("Neue Verfahren eingetragen: " & neueVerfahren.ToString())

            ' ┌─────────────────────────────────────────────────────────┐
            ' │ SCHRITT 2: Status analysieren und zurücksetzen         │
            ' │           KEINE RunID-Filterung!                       │
            ' │           Betrachtet NUR Verfahren der aktuellen        │
            ' │           STL-Datei (BA::STLDateiname)                  │
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
    ' VariablenLaden - Liest die benoetigten SSIS-Variablen in Modulfelder
    ' ein.
    ' -----------------------------------------------------------------------
    Private Sub VariablenLaden()
        _runID = Convert.ToInt32(Dts.Variables("BA::RunID").Value)
        _parameterDB = Dts.Variables("BA::ParameterDB").Value.ToString().Trim()
        _parametertabelle = Dts.Variables("BA::Parametertabelle").Value.ToString().Trim()
        _steuerlistenTabelle = Dts.Variables("BA::SteuerlistenTabelle").Value.ToString().Trim()
        _stlDateiname = Dts.Variables("BA::STLDateiname").Value.ToString().Trim()

    End Sub

    ' -----------------------------------------------------------------------
    ' PflichtfelderPruefen - Prueft, ob alle Pflichtvariablen / -parameter
    ' vorhanden sind.
    ' -----------------------------------------------------------------------
    Private Function PflichtfelderPruefen() As Boolean
        Dim fehlend As New System.Text.StringBuilder()
        If _runID <= 0 Then fehlend.AppendLine("  → BA::RunID (ungültig)")
        If String.IsNullOrEmpty(_parameterDB) Then fehlend.AppendLine("  → BA::ParameterDB")
        If String.IsNullOrEmpty(_parametertabelle) Then fehlend.AppendLine("  → BA::Parametertabelle")
        If String.IsNullOrEmpty(_steuerlistenTabelle) Then fehlend.AppendLine("  → BA::SteuerlistenTabelle")
        If String.IsNullOrEmpty(_stlDateiname) Then fehlend.AppendLine("  → BA::STLDateiname")
        If fehlend.Length > 0 Then
            LogFehler("Pflichtfelder fehlen:" & Environment.NewLine & fehlend.ToString())
            Return False
        End If
        Log("Pflichtfelder-Pruefung: alle Variablen vorhanden OK")
        Return True
    End Function

    ' -----------------------------------------------------------------------
    ' NeueVerfahrenEintragen - Traegt neue Verfahren aus der Steuerliste in
    ' die Arbeitsliste ein.
    ' -----------------------------------------------------------------------
    Private Function NeueVerfahrenEintragen() As Integer

        Dim sql As String = "
INSERT INTO dbo.ETL_Fkt_Arbeitsliste
       (RunID, Verfahren, Themengebiet, Status, LetzterSchritt, Versuche, AktualisiertAm)
SELECT DISTINCT
    @runID,
    LOWER(LTRIM(RTRIM(f.tabelle))),
    LOWER(f.themengebiet),
    'AUSSTEHEND',
    NULL,
    0,
    GETDATE()
FROM " & _parameterDB & ".dbo." & _parametertabelle & " p
INNER JOIN dbo." & _steuerlistenTabelle & " f
    ON dbo.fn_ParamVerfahren(LOWER(LTRIM(RTRIM(f.tabelle)))) = LOWER(LTRIM(RTRIM(p.Verfahren)))" & DateiFilter() & "
WHERE p.Verfahren IS NOT NULL
  AND NOT EXISTS (
      SELECT 1 FROM dbo.ETL_Fkt_Arbeitsliste a
      WHERE dbo.fn_ParamVerfahren(a.Verfahren) = LOWER(p.Verfahren)
        AND LOWER(LTRIM(RTRIM(a.Themengebiet))) = LOWER(LTRIM(RTRIM(f.themengebiet)))
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
    ' StatusAnalysierenUndZuruecksetzen - Analysiert den
    ' Arbeitslisten-Status und setzt Zeilen fuer den neuen Lauf zurueck.
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
    ON dbo.fn_ParamVerfahren(LOWER(LTRIM(RTRIM(f.tabelle)))) = LOWER(LTRIM(RTRIM(p.Verfahren)))" & DateiFilter() & "
LEFT JOIN dbo.ETL_Fkt_Arbeitsliste a
    ON dbo.fn_ParamVerfahren(a.Verfahren) = LOWER(p.Verfahren)
    AND LOWER(LTRIM(RTRIM(a.Themengebiet))) = LOWER(LTRIM(RTRIM(f.themengebiet)))
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
            ' Datei wurde geladen, aber keine der Tabellen existiert als
            ' Verfahren in der Parametertabelle -> Konfigurationsfehler,
            ' lauter Abbruch statt leerem "Erfolgs"-Lauf.
            Throw New Exception("Keine der aus [" & _stlDateiname & "] geladenen Tabellen " &
                "ist als Verfahren in der Parametertabelle [" &
                _parameterDB & ".dbo." & _parametertabelle & "] vorhanden " &
                "(Tabellenliste siehe Log oben).")
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
INNER JOIN " & _parameterDB & ".dbo." & _parametertabelle & " p ON LOWER(p.Verfahren) = dbo.fn_ParamVerfahren(a.Verfahren)
INNER JOIN dbo." & _steuerlistenTabelle & " f ON dbo.fn_ParamVerfahren(LOWER(LTRIM(RTRIM(f.tabelle)))) = LOWER(LTRIM(RTRIM(p.Verfahren)))
    AND LOWER(LTRIM(RTRIM(f.themengebiet))) = LOWER(LTRIM(RTRIM(a.Themengebiet)))" & DateiFilter() & "
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
INNER JOIN " & _parameterDB & ".dbo." & _parametertabelle & " p ON LOWER(p.Verfahren) = dbo.fn_ParamVerfahren(a.Verfahren)
INNER JOIN dbo." & _steuerlistenTabelle & " f ON dbo.fn_ParamVerfahren(LOWER(LTRIM(RTRIM(f.tabelle)))) = LOWER(LTRIM(RTRIM(p.Verfahren)))
    AND LOWER(LTRIM(RTRIM(f.themengebiet))) = LOWER(LTRIM(RTRIM(a.Themengebiet)))" & DateiFilter() & "
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
INNER JOIN " & _parameterDB & ".dbo." & _parametertabelle & " p ON LOWER(p.Verfahren) = dbo.fn_ParamVerfahren(a.Verfahren)
INNER JOIN dbo." & _steuerlistenTabelle & " f ON dbo.fn_ParamVerfahren(LOWER(LTRIM(RTRIM(f.tabelle)))) = LOWER(LTRIM(RTRIM(p.Verfahren)))
    AND LOWER(LTRIM(RTRIM(f.themengebiet))) = LOWER(LTRIM(RTRIM(a.Themengebiet)))" & DateiFilter() & "
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
    ' RunIDAktualisieren - Aktualisiert die RunID fuer alle Verfahren
    ' dieses Laufs.
    ' -----------------------------------------------------------------------
    Private Function RunIDAktualisieren() As Integer

        Dim sql As String = "
UPDATE a
SET a.RunID = @runID,
    a.AktualisiertAm = GETDATE()
FROM dbo.ETL_Fkt_Arbeitsliste a
INNER JOIN " & _parameterDB & ".dbo." & _parametertabelle & " p ON LOWER(p.Verfahren) = dbo.fn_ParamVerfahren(a.Verfahren)
INNER JOIN dbo." & _steuerlistenTabelle & " f ON dbo.fn_ParamVerfahren(LOWER(LTRIM(RTRIM(f.tabelle)))) = LOWER(LTRIM(RTRIM(p.Verfahren)))
    AND LOWER(LTRIM(RTRIM(f.themengebiet))) = LOWER(LTRIM(RTRIM(a.Themengebiet)))" & DateiFilter() & "
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
    ' HoleVerbindungszeichenfolge - Liefert den Connection String des
    ' Paket-Verbindungsmanagers.
    ' -----------------------------------------------------------------------
    Private Function HoleVerbindungszeichenfolge() As String
        Return Dts.Connections(CONN_NAME).ConnectionString
    End Function

    ' -----------------------------------------------------------------------
    ' Log - Schreibt eine Informationsmeldung in das SSIS-Protokoll
    ' (FireInformation).
    ' -----------------------------------------------------------------------
    Private Sub Log(nachricht As String)
        Dim fireAgain As Boolean = False
        Dts.Events.FireInformation(0, SKRIPT_NAME, nachricht, "", 0, fireAgain)
    End Sub

    ' -----------------------------------------------------------------------
    ' LogFehler - Schreibt eine Fehlermeldung in das SSIS-Protokoll
    ' (FireError).
    ' -----------------------------------------------------------------------
    Private Sub LogFehler(nachricht As String)
        Dts.Events.FireError(0, SKRIPT_NAME, nachricht, "", 0)
    End Sub

    ' -----------------------------------------------------------------------
    ' ScriptResults - SSIS-Task-Ergebniscodes.
    ' -----------------------------------------------------------------------
    Public Enum ScriptResults
        Success = DTSExecResult.Success
        Failure = DTSExecResult.Failure
    End Enum

End Class
