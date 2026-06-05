Option Explicit On
Option Strict On

Imports System
Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports Microsoft.SqlServer.Dts.Runtime

' =============================================================================
' PAKET  : Fakten Laden
' SKRIPT : SCR15_Partitionladen_Nachlauf
' ZWECK  : Pro Verfahren + Partition: usp_Faktenladen_Partitionladen_aufrufen
'          15 Parameter aus Parametertabelle + BA:: Variablen + objPartitionValues
'          Wenn Prozedur nicht existiert -> Info, kein Fehler
'          Status: ERFOLG -> NACHLAUF_PART -> NACHLAUF_PART_ERFOLG
' =============================================================================
<Microsoft.SqlServer.Dts.Tasks.ScriptTask.SSISScriptTaskEntryPointAttribute()>
<CLSCompliant(False)>
Partial Public Class ScriptMain
    Inherits Microsoft.SqlServer.Dts.Tasks.ScriptTask.VSTARTScriptObjectModelBase

    Private Const SKRIPT_NAME As String = "SCR16_Partitionladen_Nachlauf"
    Private Const CONN_NAME As String = "Verbindung"
    Private Const MAX_VERSUCHE As Integer = 10
    Private Const WARTE_SEK As Integer = 30

    Private Const STATUS_START As String = "ERFOLG"
    Private Const STATUS_RUN As String = "NACHLAUF_PART"
    Private Const STATUS_OK As String = "NACHLAUF_PART_ERFOLG"

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

    Public Sub Main()

        Log("════════════════════════════════════════════════════════")
        Log("SCR16_Partitionladen_Nachlauf – Start")
        Log("Zeitpunkt: " & DateTime.Now.ToString("dd.MM.yyyy HH:mm:ss"))
        Log("════════════════════════════════════════════════════════")

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

            Log("Server           : " & _server)
            Log("Datenbank        : " & _datenbank)
            Log("Datamart         : " & _datamart)
            Log("ProtokollDB      : " & _protokollDB)
            Log("Protokolltabelle : " & _protokolltabelle)
            Log("ProtokollSP      : " & _protokollSP)
            Log("UserName         : " & _userName)
            Log("PackageName      : " & _packageName)

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

            ' Nach Verfahren gruppieren
            Dim verfahrenWerte As New Dictionary(Of String, List(Of String))()
            For i As Integer = 0 To anzahl - 1
                Dim verf As String = partArray(i, 0)
                Dim wert As String = partArray(i, 1)
                Dim modus As String = partArray(i, 2)
                If Not verfahrenWerte.ContainsKey(verf) Then
                    verfahrenWerte(verf) = New List(Of String)()
                End If
                verfahrenWerte(verf).Add(wert)
                Log("  Gelesen: " & verf & " | " & wert & " | " & modus)
            Next

            Dim connStr As String = HoleVerbindungszeichenfolge()

            ' Verfahren aus ETL_Fkt_Arbeitsliste laden
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
                Log("────────────────────────────────────────────────────────")
                Log("Verfahren        : " & v.Verfahren)
                Log("Faktentabelle    : " & v.Faktentabelle)
                Log("Partitionsspalte : " & v.PartitionsSpalte)
                Log("ClusteredIndex   : " & v.ClusteredIndex)
                Log("Komprimierung    : " & v.Komprimierung)

                Dim meineWerte As List(Of String) = Nothing
                If Not verfahrenWerte.TryGetValue(v.Verfahren, meineWerte) OrElse meineWerte.Count = 0 Then
                    Log("  WARNUNG: Keine Partitionswerte -> uebersprungen")
                    StatusSetzen(connStr, v.ID, STATUS_OK)
                    cntOK += 1
                    Continue For
                End If

                Log("  Partitionen: " & String.Join(", ", meineWerte.ToArray()))

                Dim procName As String = "usp_Faktenladen_Partitionladen_" & v.Verfahren
                Log("  Prozedur: " & procName)

                If Not ProzedurExistiert(connStr, procName) Then
                    Log("  Info: Prozedur [" & procName & "] nicht aktiviert -> uebersprungen")
                    ProtokollSchreiben(connStr, v.Verfahren, "INFO_SCR16",
                        "Prozedur " & procName & " nicht aktiviert")
                    StatusSetzen(connStr, v.ID, STATUS_OK)
                    cntOK += 1
                    Continue For
                End If

                Log("  Prozedur existiert")

                Try
                    StatusSetzen(connStr, v.ID, STATUS_RUN)
                    Dim protokollID As String = "{" & Guid.NewGuid().ToString() & "}"

                    For Each partWert As String In meineWerte
                        Log("  ────────────────────────────────────────────────")
                        Log("  Partition      : " & partWert)
                        Log("  ProtokollID    : " & protokollID)

                        ' 15 Parameter (wie im Original)
                        Dim sql As String = "EXEC [" & procName & "]" &
                            " '" & partWert & "'" &
                            ", '" & v.Faktentabelle & "'" &
                            ", '" & v.PartitionsSpalte & "'" &
                            ", '" & v.ClusteredIndex & "'" &
                            ", '" & v.Komprimierung & "'" &
                            ", '" & _userName & "'" &
                            ", '" & _packageName & "'" &
                            ", '" & _server & "'" &
                            ", '" & _datenbank & "'" &
                            ", '" & _datamart & "'" &
                            ", '" & v.Verfahren & "'" &
                            ", '" & protokollID & "'" &
                            ", '" & _protokollDB & "'" &
                            ", '" & _protokolltabelle & "'" &
                            ", '" & _protokollSP & "'"

                        Log("  SQL: " & sql)

                        Dim zeilen As Integer = SqlAusfuehren(connStr, sql, "Nachlauf " & partWert)
                        Log("  -> Ausgefuehrt, Zeilen: " & zeilen.ToString())
                        ProtokollSchreiben(connStr, v.Verfahren, "NACHLAUF_PART_" & partWert,
                            "Nachlauf Partition " & partWert & " | Zeilen: " & zeilen.ToString())
                    Next

                    StatusSetzen(connStr, v.ID, STATUS_OK)
                    ProtokollSchreiben(connStr, v.Verfahren, "SCHRITT_NACHLAUF_PART",
                        "Partitionladen Nachlauf erfolgreich: " & meineWerte.Count.ToString() & " Partition(en)")
                    cntOK += 1
                    Log("  Nachlauf erfolgreich fuer " & meineWerte.Count.ToString() & " Partition(en)")

                Catch ex As Exception
                    cntFehler += 1
                    FehlerSetzen(connStr, v.ID, ex.Message)
                    ProtokollSchreiben(connStr, v.Verfahren, "FEHLER_SCR16", ex.Message)
                    LogFehler("FEHLER '" & v.Verfahren & "': " & ex.Message)
                End Try
            Next

            Log("════════════════════════════════════════════════════════")
            Log("Erfolgreich: " & cntOK.ToString() & " | Fehler: " & cntFehler.ToString())
            Log("════════════════════════════════════════════════════════")
            Dts.TaskResult = If(cntFehler > 0, ScriptResults.Failure, ScriptResults.Success)

        Catch ex As Exception
            LogFehler("Kritischer Fehler: " & ex.Message)
            Dts.TaskResult = ScriptResults.Failure
        End Try
    End Sub

    Private Function VerfahrenLaden(connStr As String) As List(Of VerfahrenInfo)
        Dim liste As New List(Of VerfahrenInfo)()
        Dim sql As String =
"SELECT a.ID, a.Verfahren, a.Themengebiet, a.LetzterSchritt," &
"       pf.Wert AS Faktentabelle," &
"       pp.Wert AS PartitionsSpalte," &
"       ISNULL(pci.Wert, 'FALSE') AS ClusteredIndex," &
"       ISNULL(pk.Wert, 'PAGE') AS Komprimierung" &
" FROM  dbo.ETL_Fkt_Arbeitsliste a" &
" JOIN  " & _parameterDB & ".dbo." & _parametertab & " pf ON pf.Verfahren=a.Verfahren AND pf.Parameter='Faktentabelle'" &
" JOIN  " & _parameterDB & ".dbo." & _parametertab & " pp ON pp.Verfahren=a.Verfahren AND pp.Parameter='Faktenpartitionsspalte'" &
" LEFT JOIN " & _parameterDB & ".dbo." & _parametertab & " pci ON pci.Verfahren=a.Verfahren AND pci.Parameter='FaktenClusteredIndex'" &
" LEFT JOIN " & _parameterDB & ".dbo." & _parametertab & " pk ON pk.Verfahren=a.Verfahren AND pk.Parameter='Faktenkomprimierung'" &
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
                                Dim rohPart As String = rdr(5).ToString().Trim()
                                liste.Add(New VerfahrenInfo With {
                                    .ID = rdr.GetInt32(0),
                                    .Verfahren = rdr(1).ToString().Trim(),
                                    .Themengebiet = rdr(2).ToString().Trim(),
                                    .LetzterSchritt = If(rdr.IsDBNull(3), "", rdr(3).ToString().Trim()),
                                    .Faktentabelle = rdr(4).ToString().Trim(),
                                    .PartitionsSpalte = If(rohPart.Contains("|"), rohPart.Substring(0, rohPart.IndexOf("|")), rohPart),
                                    .ClusteredIndex = rdr(6).ToString().Trim(),
                                    .Komprimierung = rdr(7).ToString().Trim()})
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

    Private Function ProzedurExistiert(connStr As String, procName As String) As Boolean
        Using c As New SqlConnection(connStr)
            c.Open()
            Using cmd As New SqlCommand("SELECT COUNT(*) FROM sys.objects WHERE object_id=OBJECT_ID(@p) AND type IN ('P','PC')", c)
                cmd.Parameters.AddWithValue("@p", procName)
                Return CInt(cmd.ExecuteScalar()) > 0
            End Using
        End Using
    End Function

    Private Sub StatusSetzen(connStr As String, id As Integer, status As String)
        SqlAusfuehren(connStr, "UPDATE dbo.ETL_Fkt_Arbeitsliste SET Status='" & status & "',LetzterSchritt='" & status & "',AktualisiertAm=GETDATE() WHERE ID=" & id, "Status")
    End Sub

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

    Private Sub ProtokollSchreiben(connStr As String, verfahren As String, schritt As String, meldung As String)
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
        Public Property Themengebiet As String
        Public Property LetzterSchritt As String
        Public Property Faktentabelle As String
        Public Property PartitionsSpalte As String
        Public Property ClusteredIndex As String
        Public Property Komprimierung As String
    End Class

    Public Enum ScriptResults
        Success = DTSExecResult.Success
        Failure = DTSExecResult.Failure
    End Enum

End Class
