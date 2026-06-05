Option Explicit On
Option Strict On
Imports System
Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports Microsoft.SqlServer.Dts.Runtime
' =============================================================================
' SCR19_Partitionstausch"
' ZWECK: Pro Verfahren pro _in Tabelle:
'        SWITCH OUT Faktentabelle â _out
'        CHECK Constraint auf _in
'        SWITCH IN _in â Faktentabelle
'        DROP _out + _in
'        Abschlussstatus: MSSQL MIN/MAX/COUNT loggen
'        Status: PARTITIONSTAUSCH â ERFOLG
' =============================================================================
<Microsoft.SqlServer.Dts.Tasks.ScriptTask.SSISScriptTaskEntryPointAttribute()>
<CLSCompliant(False)>
Partial Public Class ScriptMain
    Inherits Microsoft.SqlServer.Dts.Tasks.ScriptTask.VSTARTScriptObjectModelBase

    Private Const SKRIPT_NAME As String = "SCR15_Partitionstausch"
    Private Const CONN_NAME As String = "Verbindung"
    Private Const MAX_VERSUCHE As Integer = 10
    Private Const WARTE_SEK As Integer = 30
    Private _runID As Integer = 0
    Private _parameterDB As String = String.Empty
    Private _parametertab As String = String.Empty

    Public Sub Main()
        Log("ââââââââââââââââââââââââââââââââââââââââââââââââââââââââ")
        Log("SCR15_Partitionstausch â Start")
        Log("Zeitpunkt: " & DateTime.Now.ToString("dd.MM.yyyy HH:mm:ss"))
        Log("ââââââââââââââââââââââââââââââââââââââââââââââââââââââââ")
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
                Log("ââââââââââââââââââââââââââââââââââââââââââââââââââââââââ")
                Log("Verfahren: " & v.Verfahren & " | Tabelle: " & v.Faktentabelle)
                If v.LetzterSchritt = "PARTITIONSTAUSCH_ERFOLG" Then
                    Log("  â bereits abgeschlossen â Ã¼bersprungen â")
                    Continue For
                End If
                Try
                    StatusSetzen(connStr, v.ID, "PARTITIONSTAUSCH")
                    Dim pf As String = "PF_" & v.PartitionColumn & "_" & v.Faktentabelle
                    ' Alle _in Tabellen
                    Dim inTables As List(Of String) = InTabellenLaden(connStr, v.Faktentabelle)
                    Log("  _in Tabellen: " & inTables.Count.ToString())
                    For Each inTable As String In inTables
                        Dim pvStr As String = inTable.Replace(v.Faktentabelle.ToLower() & "_in_", "")
                        Dim outTable As String = v.Faktentabelle.ToLower() & "_out_" & pvStr
                        Log("  â Partition: " & pvStr)
                        ' Partitionsnummer
                        Dim pnr As Object = SqlSkalar(connStr,
                            "SELECT sprv.boundary_id FROM sys.partition_functions spf JOIN sys.partition_range_values sprv ON sprv.function_id=spf.function_id WHERE spf.name='" & pf & "' AND sprv.value=" & pvStr,
                            "Partitionsnummer")
                        If pnr Is Nothing OrElse pnr Is DBNull.Value OrElse Convert.ToInt32(pnr) = 0 Then
                            Log("  FEHLER: Partitionsnummer nicht gefunden fÃ¼r: " & pvStr)
                            LogSchreiben(connStr, v.Verfahren, "FEHLER_SWITCH", "Partition nicht gefunden: " & pvStr)
                            Continue For
                        End If
                        Dim pnrVal As Integer = Convert.ToInt32(pnr)
                        Log("  Partitionsnummer: " & pnrVal.ToString())
                        ' SWITCH OUT
                        SqlAusfuehren(connStr, "ALTER TABLE dbo.[" & v.Faktentabelle & "] SWITCH PARTITION " & pnrVal & " TO dbo.[" & outTable & "];", "SWITCH OUT")
                        Log("  â SWITCH OUT â " & outTable & " â")
                        ' CHECK Constraint auf _in
                        Dim ckName As String = v.PartitionColumn & "_" & pvStr & "_" & v.Faktentabelle & "_CK"
                        If Convert.ToInt32(SqlSkalar(connStr, "SELECT COUNT(*) FROM sys.check_constraints WHERE parent_object_id=OBJECT_ID('dbo." & inTable & "') AND name='" & ckName & "'", "CK prÃ¼fen")) > 0 Then
                            SqlAusfuehren(connStr, "ALTER TABLE dbo.[" & inTable & "] DROP CONSTRAINT [" & ckName & "];", "CK lÃ¶schen")
                        End If
                        SqlAusfuehren(connStr, "ALTER TABLE dbo.[" & inTable & "] ADD CONSTRAINT [" & ckName & "] CHECK([" & v.PartitionColumn & "]=" & pvStr & ");", "CK setzen")
                        Log("  â CHECK Constraint: " & ckName & " â")
                        ' SWITCH IN
                        SqlAusfuehren(connStr, "ALTER TABLE dbo.[" & inTable & "] SWITCH TO dbo.[" & v.Faktentabelle & "] PARTITION " & pnrVal & ";", "SWITCH IN")
                        Log("  â SWITCH IN â " & v.Faktentabelle & " Partition " & pnrVal.ToString() & " â")
                        ' Cleanup
                        If Convert.ToInt32(SqlSkalar(connStr, "SELECT CASE WHEN OBJECT_ID('dbo." & outTable & "','U') IS NOT NULL THEN 1 ELSE 0 END", "out prÃ¼fen")) = 1 Then
                            SqlAusfuehren(connStr, "DROP TABLE dbo.[" & outTable & "];", "drop _out")
                            Log("  â _out gelÃ¶scht â")
                        End If
                        If Convert.ToInt32(SqlSkalar(connStr, "SELECT CASE WHEN OBJECT_ID('dbo." & inTable & "','U') IS NOT NULL THEN 1 ELSE 0 END", "in prÃ¼fen")) = 1 Then
                            SqlAusfuehren(connStr, "DROP TABLE dbo.[" & inTable & "];", "drop _in")
                            Log("  â _in gelÃ¶scht â")
                        End If
                        LogSchreiben(connStr, v.Verfahren, "SWITCH_" & pvStr, "SWITCH IN erfolgreich â " & v.Faktentabelle & " Partition " & pnrVal.ToString())
                    Next
                    ' Abschlussstatus MSSQL
                    Dim finalMin As Object = SqlSkalar(connStr, "SELECT MIN([" & v.PartitionColumn & "]) FROM dbo.[" & v.Faktentabelle & "]", "Final MIN")
                    Dim finalMax As Object = SqlSkalar(connStr, "SELECT MAX([" & v.PartitionColumn & "]) FROM dbo.[" & v.Faktentabelle & "]", "Final MAX")
                    Dim finalCnt As Object = SqlSkalar(connStr, "SELECT COUNT(*) FROM dbo.[" & v.Faktentabelle & "]", "Final COUNT")
                    Log("  âââââââââââââââââââââââââââââââââââââââââââââââââââââââ")
                    Log("  â  ABSCHLUSSSTATUS: " & v.Faktentabelle)
                    Log("  â  Zeilen: " & Convert.ToString(finalCnt))
                    Log("  â  MIN:    " & If(finalMin Is Nothing OrElse finalMin Is DBNull.Value, "NULL", Convert.ToString(finalMin)))
                    Log("  â  MAX:    " & If(finalMax Is Nothing OrElse finalMax Is DBNull.Value, "NULL", Convert.ToString(finalMax)))
                    Log("  âââââââââââââââââââââââââââââââââââââââââââââââââââââââ")
                    LogSchreiben(connStr, v.Verfahren, "ABSCHLUSS",
                        "Fertig. Zeilen: " & Convert.ToString(finalCnt) &
                        " | MIN: " & If(finalMin Is Nothing OrElse finalMin Is DBNull.Value, "NULL", Convert.ToString(finalMin)) &
                        " | MAX: " & If(finalMax Is Nothing OrElse finalMax Is DBNull.Value, "NULL", Convert.ToString(finalMax)))
                    StatusSetzenErfolg(connStr, v.ID)
                    cntOK += 1
                    Log("  â Verfahren erfolgreich abgeschlossen â")
                Catch ex As Exception
                    cntFehler += 1
                    FehlerSetzen(connStr, v.ID, ex.Message)
                    LogSchreiben(connStr, v.Verfahren, "FEHLER_SCR15", ex.Message)
                    LogFehler("FEHLER '" & v.Verfahren & "': " & ex.Message)
                End Try
            Next
            Log("ââââââââââââââââââââââââââââââââââââââââââââââââââââââââ")
            Log("Erfolgreich: " & cntOK.ToString() & " | Fehler: " & cntFehler.ToString())
            Log("ââââââââââââââââââââââââââââââââââââââââââââââââââââââââ")
            Dts.TaskResult = If(cntFehler > 0, ScriptResults.Failure, ScriptResults.Success)
        Catch ex As Exception
            LogFehler("Kritischer Fehler: " & ex.Message)
            Dts.TaskResult = ScriptResults.Failure
        End Try
    End Sub

    Private Function InTabellenLaden(connStr As String, faktentabelle As String) As List(Of String)
        Dim liste As New List(Of String)()
        Dim versuch As Integer = 0
        While versuch < MAX_VERSUCHE
            versuch += 1
            Try
                Using conn As New SqlConnection(connStr)
                    conn.Open()
                    Using cmd As New SqlCommand("SELECT name FROM sys.tables WHERE schema_id=SCHEMA_ID('dbo') AND name LIKE '" & faktentabelle.ToLower() & "_in_%' ORDER BY name", conn)
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
                If versuch < MAX_VERSUCHE Then System.Threading.Thread.Sleep(WARTE_SEK * 1000) Else Throw
            End Try
        End While
        Return liste
    End Function

    Private Function VerfahrenLaden(connStr As String) As List(Of VerfahrenInfo)
        Dim liste As New List(Of VerfahrenInfo)()
        Dim sql As String =
"SELECT a.ID, a.Verfahren, a.LetzterSchritt, pf.Wert AS Faktentabelle, pp.Wert AS PartitionColumn
 FROM   dbo.ETL_Fkt_Arbeitsliste a
 JOIN   " & _parameterDB & ".dbo." & _parametertab & " pf ON pf.Verfahren=a.Verfahren AND pf.Parameter='Faktentabelle'
 JOIN   " & _parameterDB & ".dbo." & _parametertab & " pp ON pp.Verfahren=a.Verfahren AND pp.Parameter='Faktenpartitionsspalte'
 WHERE  a.Status IN ('NCCI_OUT_ERSTELLT','PARTITIONSTAUSCH') AND a.RunID=" & _runID & " ORDER BY a.Verfahren"
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
                                Dim rawPart As String = rdr(4).ToString().Trim()
                                liste.Add(New VerfahrenInfo With {
                                    .ID = rdr.GetInt32(0), .Verfahren = rdr(1).ToString().Trim(),
                                    .LetzterSchritt = If(rdr.IsDBNull(2), "", rdr(2).ToString().Trim()),
                                    .Faktentabelle = rdr(3).ToString().Trim(),
                                    .PartitionColumn = If(rawPart.Contains("|"), rawPart.Substring(0, rawPart.IndexOf("|")), rawPart)})
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

    Private Sub StatusSetzenErfolg(connStr As String, id As Integer)
        SqlAusfuehren(connStr, "UPDATE dbo.ETL_Fkt_Arbeitsliste SET Status='ERFOLG',LetzterSchritt='ERFOLG',AktualisiertAm=GETDATE() WHERE ID=" & id, "ERFOLG")
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
        Try
            Using conn As New SqlConnection(connStr)
                conn.Open()
                Using cmd As New SqlCommand("INSERT INTO dbo.tm_fakten_load_log(verfahren,schritt,meldung) VALUES(@v,@s,@m)", conn)
                    cmd.Parameters.AddWithValue("@v", verfahren)
                    cmd.Parameters.AddWithValue("@s", schritt)
                    cmd.Parameters.AddWithValue("@m", meldung)
                    cmd.ExecuteNonQuery()
                End Using
            End Using
        Catch
        End Try
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
        Public Property PartitionColumn As String
    End Class

    Public Enum ScriptResults
        Success = DTSExecResult.Success
        Failure = DTSExecResult.Failure
    End Enum

End Class
