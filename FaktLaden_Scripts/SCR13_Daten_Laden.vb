Option Explicit On
Option Strict On
Imports System
Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports System.Collections.Concurrent
Imports System.Threading
Imports System.Threading.Tasks
Imports Microsoft.SqlServer.Dts.Runtime
Imports Microsoft.SqlServer.Dts.Tasks.ScriptTask
Imports System.Linq

<Microsoft.SqlServer.Dts.Tasks.ScriptTask.SSISScriptTaskEntryPointAttribute()>
<CLSCompliant(False)>
Partial Public Class ScriptMain
    Inherits Microsoft.SqlServer.Dts.Tasks.ScriptTask.VSTARTScriptObjectModelBase

    Private Const SKRIPT_NAME As String = "SCR11_Daten_Laden"
    Private Const CONN_NAME As String = "Verbindung"
    Private Const MAX_VERSUCHE As Integer = 10
    Private Const WARTE_SEK As Integer = 30

    Private _runID As Integer = 0
    Private _parameterDB As String = String.Empty
    Private _parametertab As String = String.Empty
    Private _maxparallel As Integer = 0
    Private _datenbank As String = String.Empty

    Private ReadOnly _logSperre As New Object()
    Private ReadOnly _fehlerListe As New ConcurrentBag(Of String)
    Private _cntOK As Integer = 0
    Private _cntFehler As Integer = 0

    ' ==========================================================================================
    ' HAUPTABLAUF
    ' ==========================================================================================
    Public Sub Main()

        Log("════════════════════════════════════════════════════════")
        Log("SCR11_Daten_Laden – Start")
        Log("Zeitpunkt: " & DateTime.Now.ToString("dd.MM.yyyy HH:mm:ss"))
        Log("════════════════════════════════════════════════════════")

        Try
            _runID = Convert.ToInt32(Dts.Variables("BA::RunID").Value)
            _parameterDB = Dts.Variables("BA::ParameterDB").Value.ToString().Trim()
            _parametertab = Dts.Variables("BA::Parametertabelle").Value.ToString().Trim()
            _datenbank = Dts.Variables("BA::Datenbank").Value.ToString().Trim()

            ' Maximale ParallelitÃ¤t aus SSIS-Variable lesen
            _maxparallel = CInt(Dts.Variables("BA::Maxparallel").Value)
            Log("Maximale Parallelitaet: " & _maxparallel.ToString())
            Log("Datenbank             : " & _datenbank)

            Dim connStr As String = HoleVerbindungszeichenfolge()
            Dim verfahren As List(Of VerfahrenInfo) = VerfahrenLaden(connStr)
            Log("Verfahren: " & verfahren.Count.ToString())

            If verfahren.Count = 0 Then
                Log("Keine Verfahren zur Verarbeitung vorhanden.")
                Dts.TaskResult = ScriptResults.Success
                Return
            End If

            ' Parallel verarbeiten mit MaxDegreeOfParallelism
            Dim opts As New ParallelOptions With {.MaxDegreeOfParallelism = _maxparallel}

            Parallel.ForEach(verfahren, opts,
                Sub(v As VerfahrenInfo)
                    Try
                        VerarbeiteVerfahren(v, connStr)
                    Catch ex As Exception
                        Log("Verfahren " & v.Verfahren & " uebersprungen wegen kritischem Fehler: " & ex.Message)
                    End Try
                End Sub)

            Log("Erfolgreich: " & _cntOK.ToString() & " | Fehler: " & _cntFehler.ToString())

            If _fehlerListe.Count > 0 Then
                Log("════════════════════════════════════════════════════════")
                Log("FEHLER: VERFAHREN MIT VERARBEITUNGSFEHLERN:")
                Log("────────────────────────────────────────────────────────")
                For Each f As String In _fehlerListe.OrderBy(Function(x) x)
                    Log("  →¢ " & f)
                Next
                Log("════════════════════════════════════════════════════════")
            End If

            Dts.TaskResult = If(_cntFehler > 0, ScriptResults.Failure, ScriptResults.Success)

        Catch ex As Exception
            LogFehler("Kritischer Fehler: " & ex.Message)
            Dts.TaskResult = ScriptResults.Failure
        End Try

    End Sub

    ' ==========================================================================================
    ' EINZELNES VERFAHREN VERARBEITEN
    ' ==========================================================================================
    Private Sub VerarbeiteVerfahren(v As VerfahrenInfo, connStr As String)

        SyncLock _logSperre
            Log("────────────────────────────────────────────────────────")
            Log("Verfahren: " & v.Verfahren & " | Spalte: " & v.PartitionColumn)
        End SyncLock

        If v.LetzterSchritt = "DATEN_GELADEN" Then
            Log("  → bereits abgeschlossen → Ã¼bersprungen ✓")
            Return
        End If

        Try
            StatusSetzen(connStr, v.ID, "DATEN_LADEN")

            ' Alle _in Tabellen finden
            Dim inTables As List(Of String) = InTabellenLaden(connStr, v.Faktentabelle)

            SyncLock _logSperre
                Log("  _in Tabellen: " & inTables.Count.ToString())
            End SyncLock

            ' Thread-safe ZeilenzÃ¤hler
            Dim cntGesamtZeilen As Long = 0

            ' Inner loop jetzt auch PARALLEL
            Dim innerOpts As New ParallelOptions With {.MaxDegreeOfParallelism = _maxparallel}

            Parallel.ForEach(inTables, innerOpts,
            Sub(inTable As String)
                Try
                    ' -- RESUME LOGIC --
                    ' Wenn die finale Tabelle bereits existiert, wurde sie erfolgreich geladen (sp_rename erfolgt erst am Ende)
                    Dim finalTableExists As Boolean = Convert.ToInt32(SqlSkalar(connStr,
                        "SELECT COUNT(*) FROM sys.tables WHERE schema_id=SCHEMA_ID('dbo') AND name='" & inTable & "'",
                        "final table check")) > 0

                    If finalTableExists Then
                        Log("  → Partition " & inTable & " bereits geladen. Übersprungen ✓")
                        Return
                    End If

                    Dim loadingTable As String = inTable & "_LOADING"
                    Dim pvStr As String = inTable.Replace(v.Faktentabelle.ToLower() & "_in_", "")
                    Dim extTable As String = v.Faktentabelle.ToLower()

                    ' Ext Tabelle prüfen
                    Dim extExists As Boolean = Convert.ToInt32(SqlSkalar(connStr,
                        "SELECT COUNT(*) FROM sys.external_tables WHERE schema_id=SCHEMA_ID('ext') AND name='" & extTable & "'",
                        "ext prüfen")) > 0

                    If Not extExists Then
                        Log("  WARNUNG: ext." & extTable & " nicht gefunden → übersprungen")
                        Return
                    End If

                    ' -- ATOMIC LOAD (RENAME TRICK) --
                    ' 1. LOADING Tabelle sicherheitshalber löschen (falls ein vorheriger Versuch abgebrochen ist)
                    SqlAusfuehren(connStr, "IF OBJECT_ID('dbo.[" & loadingTable & "]','U') IS NOT NULL DROP TABLE dbo.[" & loadingTable & "];", "DROP " & loadingTable)

                    ' 2. SELECT INTO in die LOADING Tabelle
                    Dim zeilen As Integer = DatenLadenSelectInto(connStr, v, loadingTable, extTable, pvStr)

                    ' 3. RENAME nur bei Erfolg
                    SqlAusfuehren(connStr, "EXEC sp_rename 'dbo.[" & loadingTable & "]', '" & inTable & "';", "RENAME to " & inTable)

                    Log("  → Eingefügte Zeilen (" & inTable & "): " & zeilen.ToString())
                    LogSchreiben(connStr, v.Verfahren, "LOAD_" & pvStr,
                        "SELECT INTO dbo." & inTable & " | Zeilen: " & zeilen.ToString())

                    ' Thread-safe addieren
                    Interlocked.Add(cntGesamtZeilen, CLng(zeilen))

                Catch ex As Exception
                    Interlocked.Increment(_cntFehler)
                    _fehlerListe.Add("FEHLER '" & v.Verfahren & "." & inTable & "': " & ex.Message)
                    LogFehler("FEHLER bei " & inTable & ": " & ex.Message)
                End Try
            End Sub)

            StatusSetzen(connStr, v.ID, "DATEN_GELADEN")
            LogSchreiben(connStr, v.Verfahren, "SCHRITT_6",
            "Daten geladen. Partitionen: " & inTables.Count.ToString() & " | Zeilen gesamt: " & cntGesamtZeilen.ToString())

            Interlocked.Increment(_cntOK)
            Log("  → Schritt 6 abgeschlossen ✓")

        Catch ex As Exception
            Interlocked.Increment(_cntFehler)
            _fehlerListe.Add("FEHLER '" & v.Verfahren & "': " & ex.Message)
            FehlerSetzen(connStr, v.ID, ex.Message)
            LogSchreiben(connStr, v.Verfahren, "FEHLER_SCR11", ex.Message)
            LogFehler("FEHLER '" & v.Verfahren & "': " & ex.Message)
        End Try

    End Sub

    ' ==========================================================================================
    ' DATEN LADEN MIT SELECT INTO - verwendet DBO Spalten aus Template
    ' ==========================================================================================
    Private Function DatenLadenSelectInto(connStr As String, v As VerfahrenInfo,
                                          inTable As String, extTable As String,
                                          partitionValue As String) As Integer

        ' SELECT-Liste aus Template (via INFORMATION_SCHEMA) + Metadaten laden
        Dim templateName As String = v.Faktentabelle.ToLower() & "_template"
        
        Dim sqlSelectList As String = "
        SELECT STRING_AGG(CAST(m.columns_dbo AS nvarchar(max)), ',' + CHAR(13) + CHAR(10)) WITHIN GROUP (ORDER BY m.colno)
        FROM [" & _datenbank & "].INFORMATION_SCHEMA.COLUMNS c
        JOIN dbo.tm_polybase_struktur m ON UPPER(LTRIM(RTRIM(m.colname))) = UPPER(LTRIM(RTRIM(c.COLUMN_NAME)))
        WHERE c.TABLE_SCHEMA = 'dbo' 
          AND c.TABLE_NAME = '" & templateName & "'
          AND m.tabname = @tab
          AND m.themengebiet = @thema"

        Dim selectList As String = Nothing
        Dim versuch As Integer = 0

        While versuch < MAX_VERSUCHE
            versuch += 1
            Try
                Using conn As New SqlConnection(connStr)
                    conn.Open()
                    Using cmd As New SqlCommand(sqlSelectList, conn)
                        cmd.Parameters.AddWithValue("@tab", v.Verfahren.ToLower())
                        cmd.Parameters.AddWithValue("@thema", v.Themengebiet.ToLower())
                        Dim result As Object = cmd.ExecuteScalar()
                        If result IsNot Nothing AndAlso result IsNot DBNull.Value Then
                            selectList = result.ToString()
                        End If
                    End Using
                End Using
                Exit While
            Catch ex As Exception
                Log(String.Format("WARNUNG [Template Query] Versuch {0}/{1}: {2}", versuch, MAX_VERSUCHE, ex.Message))
                If versuch < MAX_VERSUCHE Then Thread.Sleep(WARTE_SEK * 1000) Else Throw
            End Try
        End While

        If String.IsNullOrEmpty(selectList) Then
            Throw New Exception("SELECT-Liste konnte nicht für Template " & templateName & " geladen werden.")
        End If

        ' SELECT INTO aufbauen
        Dim sql As String =
"SELECT" & vbCrLf &
selectList & vbCrLf &
"INTO dbo.[" & inTable & "]" & vbCrLf &
"FROM ext.[" & extTable & "]" & vbCrLf &
"WHERE [" & v.PartitionColumn & "] = " & partitionValue & ";"

        ' SQL loggen
        Log("SQL SELECT INTO:" & vbCrLf & sql)

        SqlAusfuehren(connStr, sql, "SELECT INTO " & inTable)

        ' Zeilen zÃ¤hlen
        Dim zeilen As Integer = Convert.ToInt32(SqlSkalar(connStr,
            "SELECT COUNT(*) FROM dbo.[" & inTable & "]",
            "Count " & inTable))

        Return zeilen
    End Function

    ' ==========================================================================================
    ' _IN TABELLEN LADEN
    ' ==========================================================================================
    Private Function InTabellenLaden(connStr As String, faktentabelle As String) As List(Of String)
        Dim liste As New List(Of String)()
        Dim versuch As Integer = 0
        While versuch < MAX_VERSUCHE
            versuch += 1
            Try
                Using conn As New SqlConnection(connStr)
                    conn.Open()
                    Using cmd As New SqlCommand(
                        "SELECT name FROM sys.tables WHERE schema_id=SCHEMA_ID('dbo') AND name LIKE '" &
                        faktentabelle.ToLower() & "_in_%' ORDER BY name", conn)
                        cmd.CommandTimeout = 0
                        Using rdr As SqlDataReader = cmd.ExecuteReader()
                            While rdr.Read()
                                liste.Add(rdr(0).ToString())
                            End While
                        End Using
                    End Using
                End Using
                Return liste
            Catch ex As Exception
                Log(String.Format("WARNUNG [_in Tabellen] Versuch {0}/{1}: {2}", versuch, MAX_VERSUCHE, ex.Message))
                If versuch < MAX_VERSUCHE Then Thread.Sleep(WARTE_SEK * 1000) Else Throw
            End Try
        End While
        Return liste
    End Function

    ' ==========================================================================================
    ' VERFAHREN LADEN
    ' ==========================================================================================
    Private Function VerfahrenLaden(connStr As String) As List(Of VerfahrenInfo)
        Dim liste As New List(Of VerfahrenInfo)()
        Dim sql As String =
"SELECT a.ID, a.Verfahren, a.Themengebiet, a.LetzterSchritt,
        pf.Wert AS Faktentabelle, pp.Wert AS PartitionColumn
 FROM   dbo.ETL_Fkt_Arbeitsliste a
 JOIN   " & _parameterDB & ".dbo." & _parametertab & " pf ON pf.Verfahren=a.Verfahren AND pf.Parameter='Faktentabelle'
 JOIN   " & _parameterDB & ".dbo." & _parametertab & " pp ON pp.Verfahren=a.Verfahren AND pp.Parameter='Faktenpartitionsspalte'
 WHERE  a.Status IN ('STAGING_ERSTELLT','DATEN_LADEN')
 AND    a.RunID = " & _runID & " ORDER BY a.Verfahren"

        Dim versuch As Integer = 0
        While versuch < MAX_VERSUCHE
            versuch += 1
            Try
                Using conn As New SqlConnection(connStr)
                    conn.Open()
                    Using cmd As New SqlCommand(sql, conn)
                        cmd.CommandTimeout = 0
                        Using rdr As SqlDataReader = cmd.ExecuteReader()
                            While rdr.Read()
                                Dim rawPart As String = rdr(5).ToString().Trim()
                                liste.Add(New VerfahrenInfo With {
                                    .ID = rdr.GetInt32(0),
                                    .Verfahren = rdr(1).ToString().Trim(),
                                    .Themengebiet = rdr(2).ToString().Trim(),
                                    .LetzterSchritt = If(rdr.IsDBNull(3), "", rdr(3).ToString().Trim()),
                                    .Faktentabelle = rdr(4).ToString().Trim(),
                                    .PartitionColumn = If(rawPart.Contains("|"),
                                        rawPart.Substring(0, rawPart.IndexOf("|")), rawPart)
                                })
                            End While
                        End Using
                    End Using
                End Using
                Return liste
            Catch ex As Exception
                Log(String.Format("WARNUNG [Verfahren laden] Versuch {0}/{1}: {2}", versuch, MAX_VERSUCHE, ex.Message))
                If versuch < MAX_VERSUCHE Then Thread.Sleep(WARTE_SEK * 1000) Else Throw
            End Try
        End While
        Return liste
    End Function

    ' ==========================================================================================
    ' STATUS METHODEN
    ' ==========================================================================================
    Private Sub StatusSetzen(connStr As String, id As Integer, status As String)
        SqlAusfuehren(connStr,
            "UPDATE dbo.ETL_Fkt_Arbeitsliste SET Status='" & status &
            "',LetzterSchritt='" & status & "',AktualisiertAm=GETDATE() WHERE ID=" & id,
            "Status")
    End Sub

    Private Sub FehlerSetzen(connStr As String, id As Integer, msg As String)
        Try
            Using conn As New SqlConnection(connStr)
                conn.Open()
                Using cmd As New SqlCommand(
                    "UPDATE dbo.ETL_Fkt_Arbeitsliste SET Status='FEHLER',Fehlermeldung=@m,AktualisiertAm=GETDATE() WHERE ID=@id",
                    conn)
                    cmd.Parameters.AddWithValue("@m", If(msg.Length > 3900, msg.Substring(0, 3900), msg))
                    cmd.Parameters.AddWithValue("@id", id)
                    cmd.ExecuteNonQuery()
                End Using
            End Using
        Catch
        End Try
    End Sub

    Private Sub LogSchreiben(connStr As String, verfahren As String, schritt As String, meldung As String)
        Try
            Using conn As New SqlConnection(connStr)
                conn.Open()
                Using cmd As New SqlCommand(
                    "INSERT INTO dbo.tm_fakten_load_log(verfahren,schritt,meldung) VALUES(@v,@s,@m)",
                    conn)
                    cmd.Parameters.AddWithValue("@v", verfahren)
                    cmd.Parameters.AddWithValue("@s", schritt)
                    cmd.Parameters.AddWithValue("@m", meldung)
                    cmd.ExecuteNonQuery()
                End Using
            End Using
        Catch
        End Try
    End Sub

    ' ==========================================================================================
    ' SQL HILFSMETHODEN
    ' ==========================================================================================
    Private Function SqlAusfuehren(connStr As String, sql As String, beschreibung As String) As Integer
        Dim versuch As Integer = 0
        Dim letzterFehler As Exception = Nothing
        While versuch < MAX_VERSUCHE
            versuch += 1
            Try
                Using conn As New SqlConnection(connStr)
                    conn.Open()
                    Using cmd As New SqlCommand(sql, conn)
                        cmd.CommandTimeout = 0
                        Return cmd.ExecuteNonQuery()
                    End Using
                End Using
            Catch ex As Exception
                letzterFehler = ex
                Log(String.Format("WARNUNG [{0}] Versuch {1}/{2}: {3}", beschreibung, versuch, MAX_VERSUCHE, ex.Message))
                If versuch < MAX_VERSUCHE Then Thread.Sleep(WARTE_SEK * 1000)
            End Try
        End While
        Throw New Exception(String.Format("[{0}] fehlgeschlagen: {1}", beschreibung,
            If(letzterFehler IsNot Nothing, letzterFehler.Message, "Unbekannt")))
    End Function

    Private Function SqlSkalar(connStr As String, sql As String, beschreibung As String) As Object
        Dim versuch As Integer = 0
        While versuch < MAX_VERSUCHE
            versuch += 1
            Try
                Using conn As New SqlConnection(connStr)
                    conn.Open()
                    Using cmd As New SqlCommand(sql, conn)
                        cmd.CommandTimeout = 0
                        Return cmd.ExecuteScalar()
                    End Using
                End Using
            Catch ex As Exception
                Log(String.Format("WARNUNG [{0}] Versuch {1}/{2}: {3}", beschreibung, versuch, MAX_VERSUCHE, ex.Message))
                If versuch < MAX_VERSUCHE Then Thread.Sleep(WARTE_SEK * 1000) Else Throw
            End Try
        End While
        Return Nothing
    End Function

    Private Function HoleVerbindungszeichenfolge() As String
        Return Dts.Connections(CONN_NAME).ConnectionString
    End Function

    ' ==========================================================================================
    ' LOGGING
    ' ==========================================================================================
    Private Sub Log(n As String)
        SyncLock _logSperre
            Dim f As Boolean = False
            Dts.Events.FireInformation(0, SKRIPT_NAME, n, "", 0, f)
        End SyncLock
    End Sub

    Private Sub LogFehler(n As String)
        SyncLock _logSperre
            Dts.Events.FireError(0, SKRIPT_NAME, n, "", 0)
        End SyncLock
    End Sub

    ' ==========================================================================================
    ' DATENSTRUKTUREN
    ' ==========================================================================================
    Private Class VerfahrenInfo
        Public Property ID As Integer
        Public Property Verfahren As String
        Public Property Themengebiet As String
        Public Property LetzterSchritt As String
        Public Property Faktentabelle As String
        Public Property PartitionColumn As String
    End Class

    Public Enum ScriptResults
        Success = DTSExecResult.Success
        Failure = DTSExecResult.Failure
    End Enum

End Class
