Option Explicit On
Option Strict On
Imports System
Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports Microsoft.SqlServer.Dts.Runtime
' =============================================================================
' SCR17_Komprimierung_Out
' ZWECK: PAGE/ROW Komprimierung auf alle _out Tabellen (nur bei CI, nicht CCI)
' Status: KOMPRIMIERUNG → KOMPRIMIERUNG_ERSTELLT
' =============================================================================
<Microsoft.SqlServer.Dts.Tasks.ScriptTask.SSISScriptTaskEntryPointAttribute()>
<CLSCompliant(False)>
Partial Public Class ScriptMain
    Inherits Microsoft.SqlServer.Dts.Tasks.ScriptTask.VSTARTScriptObjectModelBase

    Private Const SKRIPT_NAME As String = "SCR13_Komprimierung_Out"
    Private Const CONN_NAME As String = "Verbindung"
    Private Const MAX_VERSUCHE As Integer = 10
    Private Const WARTE_SEK As Integer = 30
    Private _runID As Integer = 0
    Private _parameterDB As String = String.Empty
    Private _parametertab As String = String.Empty

    Public Sub Main()
        Log("════════════════════════════════════════════════════════")
        Log("SCR13_Komprimierung_Out – Start")
        Log("Zeitpunkt: " & DateTime.Now.ToString("dd.MM.yyyy HH:mm:ss"))
        Log("════════════════════════════════════════════════════════")
        Try
            _runID = Convert.ToInt32(Dts.Variables("BA::RunID").Value)
            _parameterDB = Dts.Variables("BA::ParameterDB").Value.ToString().Trim()
            _parametertab = Dts.Variables("BA::Parametertabelle").Value.ToString().Trim()
            Dim connStr As String = HoleVerbindungszeichenfolge()
            Dim verfahren As List(Of VerfahrenInfo) = VerfahrenLaden(connStr)
            Log("Verfahren: " & verfahren.Count.ToString())
            Dim cntOK As Integer = 0
            Dim cntFehler As Integer = 0
            For Each v As VerfahrenInfo In verfahren
                Log("────────────────────────────────────────────────────────")
                Log("Verfahren: " & v.Verfahren & " | Komprimierung: " & v.Compression & " | IndexTyp: " & v.IndexType)
                If v.LetzterSchritt = "KOMPRIMIERUNG_ERSTELLT" Then
                    Log("  → bereits abgeschlossen → Ã¼bersprungen ✓")
                    Continue For
                End If
                Try
                    StatusSetzen(connStr, v.ID, "KOMPRIMIERUNG")
                    If v.IndexType = "CCI" Then
                        Log("  → CCI aktiv → PAGE/ROW nicht anwendbar → Ã¼bersprungen")
                        StatusSetzen(connStr, v.ID, "KOMPRIMIERUNG_ERSTELLT")
                        LogSchreiben(connStr, v.Verfahren, "SCHRITT_7B", "CCI aktiv → Komprimierung Ã¼bersprungen")
                        cntOK += 1
                        Continue For
                    End If
                    If v.Compression <> "PAGE" AndAlso v.Compression <> "ROW" Then
                        Log("  → Keine Komprimierung konfiguriert → Ã¼bersprungen")
                        StatusSetzen(connStr, v.ID, "KOMPRIMIERUNG_ERSTELLT")
                        LogSchreiben(connStr, v.Verfahren, "SCHRITT_7B", "Keine Komprimierung konfiguriert → Ã¼bersprungen")
                        cntOK += 1
                        Continue For
                    End If
                    ' Alle _out Tabellen
                    Dim cntTbl As Integer = 0
                    Dim versuch2 As Integer = 0
                    While versuch2 < MAX_VERSUCHE
                        versuch2 += 1
                        Try
                            Using conn As New SqlConnection(connStr)
                                conn.Open()
                                Using cmd As New SqlCommand("SELECT name FROM sys.tables WHERE schema_id=SCHEMA_ID('dbo') AND name LIKE '" & v.Faktentabelle.ToLower() & "_out_%' ORDER BY name", conn)
                                    cmd.CommandTimeout = 0
                                    Using rdr As SqlDataReader = cmd.ExecuteReader()
                                        While rdr.Read()
                                            Dim tbl As String = rdr(0).ToString()
                                            SqlAusfuehren(connStr, "ALTER TABLE dbo.[" & tbl & "] REBUILD WITH (DATA_COMPRESSION=" & v.Compression & ");", "Komprimierung " & tbl)
                                            Log("  → Komprimierung " & v.Compression & " auf: " & tbl & " ✓")
                                            cntTbl += 1
                                        End While
                                    End Using
                                End Using
                            End Using
                            Exit While
                        Catch ex As Exception
                            If versuch2 < MAX_VERSUCHE Then System.Threading.Thread.Sleep(WARTE_SEK * 1000) Else Throw
                        End Try
                    End While
                    StatusSetzen(connStr, v.ID, "KOMPRIMIERUNG_ERSTELLT")
                    LogSchreiben(connStr, v.Verfahren, "SCHRITT_7B", "Komprimierung " & v.Compression & " auf " & cntTbl.ToString() & " _out Tabellen")
                    cntOK += 1
                    Log("  → Schritt 7b abgeschlossen ✓")
                Catch ex As Exception
                    cntFehler += 1
                    FehlerSetzen(connStr, v.ID, ex.Message)
                    LogSchreiben(connStr, v.Verfahren, "FEHLER_SCR13", ex.Message)
                    LogFehler("FEHLER '" & v.Verfahren & "': " & ex.Message)
                End Try
            Next
            Log("Erfolgreich: " & cntOK.ToString() & " | Fehler: " & cntFehler.ToString())
            Dts.TaskResult = If(cntFehler > 0, ScriptResults.Failure, ScriptResults.Success)
        Catch ex As Exception
            LogFehler("Kritischer Fehler: " & ex.Message)
            Dts.TaskResult = ScriptResults.Failure
        End Try
    End Sub

    Private Function VerfahrenLaden(connStr As String) As List(Of VerfahrenInfo)
        Dim liste As New List(Of VerfahrenInfo)()
        Dim sql As String =
"SELECT a.ID, a.Verfahren, a.LetzterSchritt, pf.Wert AS Faktentabelle,
        UPPER(ISNULL(pc.Wert,'NONE')) AS Compression, UPPER(pi.Wert) AS IndexType
 FROM   dbo.ETL_Fkt_Arbeitsliste a
 JOIN   " & _parameterDB & ".dbo." & _parametertab & " pf ON pf.Verfahren=a.Verfahren AND pf.Parameter='Faktentabelle'
 LEFT JOIN " & _parameterDB & ".dbo." & _parametertab & " pc ON pc.Verfahren=a.Verfahren AND pc.Parameter='Faktenkomprimierung'
 JOIN   " & _parameterDB & ".dbo." & _parametertab & " pi ON pi.Verfahren=a.Verfahren AND pi.Parameter='FaktenClusteredIndex'
 WHERE  a.Status IN ('INDEX_IN_OUT_ERSTELLT','KOMPRIMIERUNG') AND a.RunID=" & _runID & " ORDER BY a.Verfahren"
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
                                liste.Add(New VerfahrenInfo With {
                                    .ID = rdr.GetInt32(0), .Verfahren = rdr(1).ToString().Trim(),
                                    .LetzterSchritt = If(rdr.IsDBNull(2), "", rdr(2).ToString().Trim()),
                                    .Faktentabelle = rdr(3).ToString().Trim(),
                                    .Compression = rdr(4).ToString().Trim(),
                                    .IndexType = rdr(5).ToString().Trim()})
                            End While
                        End Using
                    End Using
                End Using
                Return liste
            Catch ex As Exception
                If versuch < MAX_VERSUCHE Then System.Threading.Thread.Sleep(WARTE_SEK * 1000) Else Throw
            End Try
        End While
        Return liste
    End Function

    Private Sub StatusSetzen(connStr As String, id As Integer, status As String)
        SqlAusfuehren(connStr, "UPDATE dbo.ETL_Fkt_Arbeitsliste SET Status='" & status & "',LetzterSchritt='" & status & "',AktualisiertAm=GETDATE() WHERE ID=" & id, "Status")
    End Sub

    Private Sub FehlerSetzen(connStr As String, id As Integer, msg As String)
        Try
            Using conn As New SqlConnection(connStr)
                conn.Open()
                Using cmd As New SqlCommand("UPDATE dbo.ETL_Fkt_Arbeitsliste SET Status='FEHLER',Fehlermeldung=@m,AktualisiertAm=GETDATE() WHERE ID=@id", conn)
                    cmd.Parameters.AddWithValue("@m", If(msg.Length > 3900, msg.Substring(0, 3900), msg))
                    cmd.Parameters.AddWithValue("@id", id)
                    cmd.ExecuteNonQuery()
                End Using
            End Using
        Catch
        End Try
    End Sub

    Private Sub LogSchreiben(connStr As String, verfahren As String, schritt As String, meldung As String)
        ' Kein DB-Log: Logging laeuft vollstaendig ueber SSIS Events (Eventhandler)
        ' FEHLER_* -> FireError | alles andere -> FireInformation
        If schritt.StartsWith("FEHLER", StringComparison.OrdinalIgnoreCase) Then
            LogFehler("[" & schritt & "] " & verfahren & ": " & meldung)
        Else
            Log("[" & schritt & "] " & verfahren & ": " & meldung)
        End If
    End Sub

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
                If versuch < MAX_VERSUCHE Then System.Threading.Thread.Sleep(WARTE_SEK * 1000)
            End Try
        End While
        Throw New Exception(String.Format("[{0}] fehlgeschlagen: {1}", beschreibung, If(letzterFehler IsNot Nothing, letzterFehler.Message, "Unbekannt")))
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
                If versuch < MAX_VERSUCHE Then System.Threading.Thread.Sleep(WARTE_SEK * 1000) Else Throw
            End Try
        End While
        Return Nothing
    End Function

    Private Function HoleVerbindungszeichenfolge() As String
        Return Dts.Connections(CONN_NAME).ConnectionString
    End Function

    Private Sub Log(n As String)
        Dim f As Boolean = False
        Dts.Events.FireInformation(0, SKRIPT_NAME, n, "", 0, f)
    End Sub

    Private Sub LogFehler(n As String)
        Dts.Events.FireError(0, SKRIPT_NAME, n, "", 0)
    End Sub

    Private Class VerfahrenInfo
        Public Property ID As Integer
        Public Property Verfahren As String
        Public Property LetzterSchritt As String
        Public Property Faktentabelle As String
        Public Property Compression As String
        Public Property IndexType As String
    End Class

    Public Enum ScriptResults
        Success = DTSExecResult.Success
        Failure = DTSExecResult.Failure
    End Enum

End Class
