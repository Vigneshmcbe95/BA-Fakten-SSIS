Option Explicit On
Option Strict On

Imports System
Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports Microsoft.SqlServer.Dts.Runtime

' =============================================================================
'  Skript       : SCR21_usp_Faktenladen_Nachlauf
'  Paket        : Fakten Laden (SSIS)
'  Zweck        : Fuehrt die optionale Abschluss-Prozedur
'                 usp_Faktenladen_Nachlauf_<fakt> je Verfahren nach
'                 erfolgreichem Lauf aus, sofern vorhanden.
'  Ablauf       : NACHLAUF_PART_ERFOLG -> NACHLAUF_VERF_ERFOLG
'  Wiederholung : 3 Versuche je SQL-Anweisung, 30 s Wartezeit
'  Protokoll    : Nur SSIS-Events (FireInformation / FireError)
' =============================================================================
<Microsoft.SqlServer.Dts.Tasks.ScriptTask.SSISScriptTaskEntryPointAttribute()>
<CLSCompliant(False)>
Partial Public Class ScriptMain
    Inherits Microsoft.SqlServer.Dts.Tasks.ScriptTask.VSTARTScriptObjectModelBase

    Private Const SKRIPT_NAME As String = "SCR21_usp_Faktenladen_Nachlauf"
    Private Const CONN_NAME As String = "Verbindung"
    Private Const MAX_VERSUCHE As Integer = 3
    Private Const WARTE_SEK As Integer = 30

    Private Const STATUS_START As String = "NACHLAUF_PART_ERFOLG"
    Private Const STATUS_RUN As String = "NACHLAUF_VERF"
    Private Const STATUS_OK As String = "NACHLAUF_VERF_ERFOLG"

    Private _runID As Integer = 0
    Private _parameterDB As String = String.Empty
    Private _parametertab As String = String.Empty
    Private _server As String = String.Empty
    Private _datenbank As String = String.Empty
    Private _datamart As String = String.Empty
    Private _protokollDB As String = String.Empty
    Private _protokolltabelle As String = String.Empty
    Private _protokollSP As String = String.Empty
    Private _userName As String = String.Empty
    Private _packageName As String = String.Empty

    ' -----------------------------------------------------------------------
    ' Main - Einstiegspunkt - steuert den Ablauf des Skripts.
    ' -----------------------------------------------------------------------
    Public Sub Main()

        Log("SCR21_usp_Faktenladen_Nachlauf - Start")

        Try
            _runID = Convert.ToInt32(Dts.Variables("BA::RunID").Value)
            _parameterDB = Dts.Variables("BA::ParameterDB").Value.ToString().Trim()
            _parametertab = Dts.Variables("BA::Parametertabelle").Value.ToString().Trim()
            _server = Dts.Variables("BA::Server").Value.ToString().Trim()
            _datenbank = Dts.Variables("BA::Datenbank").Value.ToString().Trim()
            _datamart = Dts.Variables("BA::Datamart").Value.ToString().Trim()
            _protokollDB = Dts.Variables("BA::ProtokollDB").Value.ToString().Trim()
            _protokolltabelle = Dts.Variables("BA::Protokolltabelle").Value.ToString().Trim()
            _protokollSP = Dts.Variables("BA::ProtokollSP").Value.ToString().Trim()
            _userName = System.Environment.UserName
            _packageName = System.Environment.MachineName


            ' BA::objPartitionValues lesen
            Dim partObjekt As Object = Dts.Variables("BA::objPartitionValues").Value
            If partObjekt Is Nothing Then
                Log("BA::objPartitionValues ist leer -> keine Nachlauf-Aufrufe noetig")
                Dts.TaskResult = ScriptResults.Success
                Return
            End If

            Dim partArray(,) As String = CType(partObjekt, String(,))
            Dim anzahl As Integer = partArray.GetLength(0)
            Log("Partitionswerte: " & anzahl.ToString() & " Eintraege")

            ' Nach Verfahren gruppieren (nur mow_ids sammeln)
            Dim verfahrenWerte As New Dictionary(Of String, List(Of String))()
            For i As Integer = 0 To anzahl - 1
                Dim verf As String = partArray(i, 0)
                Dim wert As String = partArray(i, 1)
                Dim modus As String = partArray(i, 2)
                If Not verfahrenWerte.ContainsKey(verf) Then
                    verfahrenWerte(verf) = New List(Of String)()
                End If
                verfahrenWerte(verf).Add(wert)
            Next

            Dim connStr As String = HoleVerbindungszeichenfolge()

            ' Verfahren aus ETL_Fkt_Arbeitslisteladen
            Dim verfahren As List(Of VerfahrenInfo) = VerfahrenLaden(connStr)
            Log("Verfahren mit Status '" & STATUS_START & "': " & verfahren.Count.ToString())

            If verfahren.Count = 0 Then
                Log("Keine Verfahren zu verarbeiten")
                Dts.TaskResult = ScriptResults.Success
                Return
            End If

            Dim cntOK As Integer = 0
            Dim cntFehler As Integer = 0

            For Each v As VerfahrenInfo In verfahren
                Log("Verfahren: " & v.Verfahren)

                Dim meineWerte As List(Of String) = Nothing
                If Not verfahrenWerte.TryGetValue(v.Verfahren, meineWerte) OrElse meineWerte.Count = 0 Then
                    Log("  WARNUNG: Keine Partitionswerte -> uebersprungen")
                    StatusSetzen(connStr, v.ID, STATUS_OK)
                    cntOK += 1
                    Continue For
                End If

                ' Komma-getrennte mow_id Liste bauen
                Dim mowIdListe As String = String.Join(",", meineWerte.ToArray())
                Log("  MowIdListe: " & mowIdListe)

                Dim procName As String = "usp_Faktenladen_Nachlauf_" & v.Verfahren
                Log("  Prozedur: " & procName)

                If Not ProzedurExistiert(connStr, procName) Then
                    Log("  Info: Prozedur [" & procName & "] nicht aktiviert -> uebersprungen")
                    ProtokollSchreiben(connStr, v.Verfahren, "INFO_SCR17",
                        "Prozedur " & procName & " nicht aktiviert")
                    StatusSetzen(connStr, v.ID, STATUS_OK)
                    cntOK += 1
                    Continue For
                End If

                Log("  Prozedur existiert")

                Try
                    StatusSetzen(connStr, v.ID, STATUS_RUN)
                    Dim protokollID As String = "{" & Guid.NewGuid().ToString() & "}"

                    ' 11 Parameter: originale 10 + @MowIdListe
                    Dim sql As String = "EXEC [" & procName & "]" &
                        " '" & _userName & "'" &
                        ", '" & _packageName & "'" &
                        ", '" & _server & "'" &
                        ", '" & _datenbank & "'" &
                        ", '" & _datamart & "'" &
                        ", '" & v.Verfahren & "'" &
                        ", '" & protokollID & "'" &
                        ", '" & _protokollDB & "'" &
                        ", '" & _protokolltabelle & "'" &
                        ", '" & _protokollSP & "'" &
                        ", '" & mowIdListe & "'"

                    Log("  SQL: " & sql)
                    Log("  ProtokollID: " & protokollID)
                    Log("  MowIdListe:  " & mowIdListe)

                    Dim zeilen As Integer = SqlAusfuehren(connStr, sql, "Nachlauf " & v.Verfahren)
                    Log("  -> Ausgefuehrt, Zeilen: " & zeilen.ToString())
                    ProtokollSchreiben(connStr, v.Verfahren, "NACHLAUF_VERF",
                        "Nachlauf Verfahren: " & procName & " | MowIdListe: " & mowIdListe & " | Zeilen: " & zeilen.ToString())

                    StatusSetzen(connStr, v.ID, STATUS_OK)
                    cntOK += 1
                    Log("  Nachlauf Verfahren erfolgreich")

                Catch ex As Exception
                    cntFehler += 1
                    FehlerSetzen(connStr, v.ID, ex.Message)
                    ProtokollSchreiben(connStr, v.Verfahren, "FEHLER_SCR17", ex.Message)
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

    ' -----------------------------------------------------------------------
    ' VerfahrenLaden - Laedt die zu verarbeitenden Verfahren aus der
    ' Arbeitsliste (Join mit der Parametertabelle).
    ' -----------------------------------------------------------------------
    Private Function VerfahrenLaden(connStr As String) As List(Of VerfahrenInfo)
        Dim liste As New List(Of VerfahrenInfo)()
        Dim sql As String =
"SELECT a.ID, a.Verfahren, a.Themengebiet, a.LetzterSchritt" &
" FROM  dbo.ETL_Fkt_Arbeitsliste a" &
" WHERE a.Status IN ('" & STATUS_START & "','" & STATUS_RUN & "')" &
" AND   a.RunID = " & _runID &
" ORDER BY a.Verfahren"

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
                                    .ID = rdr.GetInt32(0),
                                    .Verfahren = rdr(1).ToString().Trim(),
                                    .Themengebiet = rdr(2).ToString().Trim(),
                                    .LetzterSchritt = If(rdr.IsDBNull(3), "", rdr(3).ToString().Trim())})
                            End While
                        End Using
                    End Using
                End Using
                Return liste
            Catch ex As Exception
                Log(String.Format("WARNUNG [Verfahren laden] Versuch {0}/{1}: {2}", versuch, MAX_VERSUCHE, ex.Message))
                If versuch < MAX_VERSUCHE Then System.Threading.Thread.Sleep(WARTE_SEK * 1000) Else Throw
            End Try
        End While
        Return liste
    End Function

    ' -----------------------------------------------------------------------
    ' ProzedurExistiert - Prueft, ob eine Stored Procedure existiert.
    ' -----------------------------------------------------------------------
    Private Function ProzedurExistiert(connStr As String, procName As String) As Boolean
        Using c As New SqlConnection(connStr)
            c.Open()
            Using cmd As New SqlCommand("SELECT COUNT(*) FROM sys.objects WHERE object_id=OBJECT_ID(@p) AND type IN ('P','PC')", c)
                cmd.Parameters.AddWithValue("@p", procName)
                Return CInt(cmd.ExecuteScalar()) > 0
            End Using
        End Using
    End Function

    ' -----------------------------------------------------------------------
    ' StatusSetzen - Aktualisiert Status / LetzterSchritt einer
    ' Arbeitslisten-Zeile.
    ' -----------------------------------------------------------------------
    Private Sub StatusSetzen(connStr As String, id As Integer, status As String)
        SqlAusfuehren(connStr, "UPDATE dbo.ETL_Fkt_Arbeitsliste SET Status='" & status & "',LetzterSchritt='" & status & "',AktualisiertAm=GETDATE() WHERE ID=" & id, "Status")
    End Sub

    ' -----------------------------------------------------------------------
    ' FehlerSetzen - Markiert eine Arbeitslisten-Zeile als FEHLER und
    ' speichert die Fehlermeldung.
    ' -----------------------------------------------------------------------
    Private Sub FehlerSetzen(connStr As String, id As Integer, meldung As String)
        Try
            Using conn As New SqlConnection(connStr)
                conn.Open()
                Using cmd As New SqlCommand("UPDATE dbo.ETL_Fkt_Arbeitsliste SET Status='FEHLER',Fehlermeldung=@m,AktualisiertAm=GETDATE() WHERE ID=@id", conn)
                    cmd.Parameters.AddWithValue("@m", If(meldung.Length > 3900, meldung.Substring(0, 3900), meldung))
                    cmd.Parameters.AddWithValue("@id", id)
                    cmd.ExecuteNonQuery()
                End Using
            End Using
        Catch
        End Try
    End Sub

    ' -----------------------------------------------------------------------
    ' ProtokollSchreiben - Leitet Protokollmeldungen an SSIS-Events weiter:
    ' FEHLER_* -> FireError, alles andere -> FireInformation.
    ' -----------------------------------------------------------------------
    Private Sub ProtokollSchreiben(connStr As String, verfahren As String, schritt As String, meldung As String)
        ' Kein DB-Log: Logging laeuft vollstaendig ueber SSIS Events (Eventhandler)
        ' FEHLER_* -> FireError | alles andere -> FireInformation
        If schritt.StartsWith("FEHLER", StringComparison.OrdinalIgnoreCase) Then
            LogFehler("[" & schritt & "] " & verfahren & ": " & meldung)
        Else
            Log("[" & schritt & "] " & verfahren & ": " & meldung)
        End If
    End Sub

    ' -----------------------------------------------------------------------
    ' SqlAusfuehren - Fuehrt eine SQL-Anweisung (Non-Query) mit
    ' Wiederholung aus; protokolliert Warnung und vollstaendiges
    ' SQL-Statement bei Fehlern.
    ' -----------------------------------------------------------------------
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
                If versuch = 1 Then Log("SQL Statement [" & beschreibung & "]: " & sql)
                If versuch < MAX_VERSUCHE Then System.Threading.Thread.Sleep(WARTE_SEK * 1000)
            End Try
        End While
        Throw New Exception(String.Format("[{0}] fehlgeschlagen: {1}", beschreibung, If(letzterFehler IsNot Nothing, letzterFehler.Message, "Unbekannt")))
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
    Private Sub Log(n As String)
        Dim f As Boolean = False
        Dts.Events.FireInformation(0, SKRIPT_NAME, n, "", 0, f)
    End Sub

    ' -----------------------------------------------------------------------
    ' LogFehler - Schreibt eine Fehlermeldung in das SSIS-Protokoll
    ' (FireError).
    ' -----------------------------------------------------------------------
    Private Sub LogFehler(n As String)
        Dts.Events.FireError(0, SKRIPT_NAME, n, "", 0)
    End Sub

    ' -----------------------------------------------------------------------
    ' VerfahrenInfo - Datencontainer fuer ein Verfahren der Arbeitsliste.
    ' -----------------------------------------------------------------------
    Private Class VerfahrenInfo
        Public Property ID As Integer
        Public Property Verfahren As String
        Public Property Themengebiet As String
        Public Property LetzterSchritt As String
    End Class

    ' -----------------------------------------------------------------------
    ' ScriptResults - SSIS-Task-Ergebniscodes.
    ' -----------------------------------------------------------------------
    Public Enum ScriptResults
        Success = DTSExecResult.Success
        Failure = DTSExecResult.Failure
    End Enum

End Class
