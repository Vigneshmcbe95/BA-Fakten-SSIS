Option Explicit On
Option Strict On

Imports System
Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports Microsoft.SqlServer.Dts.Runtime

' =============================================================================
'  Skript       : SCR14_StagetabellenDML_Aufruf
'  Paket        : Fakten Laden (SSIS)
'  Zweck        : Fuehrt die optionale DML-Prozedur
'                 usp_Faktenladen_StagetabellenDML_<fakt> je Verfahren auf
'                 den Staging-Daten aus.
'  Ablauf       : DATEN_GELADEN -> STAGE_DML_ERFOLG
'  Wiederholung : 3 Versuche je SQL-Anweisung, 30 s Wartezeit
'  Protokoll    : Nur SSIS-Events (FireInformation / FireError)
' =============================================================================
<Microsoft.SqlServer.Dts.Tasks.ScriptTask.SSISScriptTaskEntryPointAttribute()>
<CLSCompliant(False)>
Partial Public Class ScriptMain
    Inherits Microsoft.SqlServer.Dts.Tasks.ScriptTask.VSTARTScriptObjectModelBase

    Private Const SKRIPT_NAME As String = "SCR14_StagetabellenDML_Aufruf"
    Private Const CONN_NAME As String = "Verbindung"
    Private Const MAX_VERSUCHE As Integer = 3
    Private Const WARTE_SEK As Integer = 30

    Private Const STATUS_START As String = "DATEN_GELADEN"
    Private Const STATUS_RUN As String = "STAGE_DML"
    Private Const STATUS_OK As String = "STAGE_DML_ERFOLG"

    Private _runID As Integer = 0
    Private _parameterDB As String = String.Empty
    Private _parametertab As String = String.Empty

    Private _server As String = String.Empty
    Private _datenbank As String = String.Empty
    Private _datamart As String = String.Empty
    Private _verfahrenBA As String = String.Empty
    Private _protokollDB As String = String.Empty
    Private _protokolltabelle As String = String.Empty
    Private _protokollSP As String = String.Empty
    Private _userName As String = String.Empty
    Private _packageName As String = String.Empty

    ' -----------------------------------------------------------------------
    ' Main - Einstiegspunkt - steuert den Ablauf des Skripts.
    ' -----------------------------------------------------------------------
    Public Sub Main()

        Log("SCR14_StagetabellenDML_Aufruf - Start")

        Try
            _runID = Convert.ToInt32(Dts.Variables("BA::RunID").Value)
            _parameterDB = Dts.Variables("BA::ParameterDB").Value.ToString().Trim()
            _parametertab = Dts.Variables("BA::Parametertabelle").Value.ToString().Trim()
            _server = Dts.Variables("BA::Server").Value.ToString().Trim()
            _datenbank = Dts.Variables("BA::Datenbank").Value.ToString().Trim()
            _datamart = Dts.Variables("BA::Datamart").Value.ToString().Trim()
            _verfahrenBA = Dts.Variables("BA::Verfahren").Value.ToString().Trim()
            _protokollDB = Dts.Variables("BA::ProtokollDB").Value.ToString().Trim()
            _protokolltabelle = Dts.Variables("BA::Protokolltabelle").Value.ToString().Trim()
            _protokollSP = Dts.Variables("BA::ProtokollSP").Value.ToString().Trim()
            _userName = Dts.Variables("System::UserName").Value.ToString().Trim()
            _packageName = Dts.Variables("System::PackageName").Value.ToString().Trim()


            ' BA::objPartitionValues lesen
            Dim partObjekt As Object = Dts.Variables("BA::objPartitionValues").Value
            If partObjekt Is Nothing Then
                Log("BA::objPartitionValues ist leer (Nothing) -> keine DML Aufrufe noetig")
                Dts.TaskResult = ScriptResults.Success
                Return
            End If

            Dim partArray(,) As String = CType(partObjekt, String(,))
            Dim anzahlEintraege As Integer = partArray.GetLength(0)
            Log("Partitionswerte aus SCR09: " & anzahlEintraege.ToString() & " Eintraege")

            Dim verfahrenWerte As New Dictionary(Of String, List(Of PartitionsEintrag))()
            For i As Integer = 0 To anzahlEintraege - 1
                Dim verf As String = partArray(i, 0)
                Dim wert As String = partArray(i, 1)
                Dim modus As String = partArray(i, 2)
                If Not verfahrenWerte.ContainsKey(verf) Then
                    verfahrenWerte(verf) = New List(Of PartitionsEintrag)()
                End If
                verfahrenWerte(verf).Add(New PartitionsEintrag With {.Wert = wert, .Modus = modus})
            Next

            Dim connStr As String = HoleVerbindungszeichenfolge()
            Dim verfahren As List(Of VerfahrenInfo) = VerfahrenLaden(connStr)
            Log("Verfahren mit Status '" & STATUS_START & "': " & verfahren.Count.ToString())

            If verfahren.Count = 0 Then
                Log("Keine Verfahren zu verarbeiten - beende erfolgreich.")
                Dts.TaskResult = ScriptResults.Success
                Return
            End If

            Dim cntOK As Integer = 0
            Dim cntFehler As Integer = 0
            Dim fehlerListe As New List(Of String)()

            For Each v As VerfahrenInfo In verfahren
                Log("Verfahren        : " & v.Verfahren)
                Log("Faktentabelle    : " & v.Faktentabelle)
                Log("Partitionsspalte : " & v.PartitionsSpalte)
                Log("ClusteredIndex   : " & v.ClusteredIndex)
                Log("Komprimierung    : " & v.Komprimierung)

                Dim meineWerte As List(Of PartitionsEintrag) = Nothing
                If Not verfahrenWerte.TryGetValue(v.Verfahren, meineWerte) OrElse meineWerte.Count = 0 Then
                    Log("  WARNUNG: Keine Partitionswerte in BA::objPartitionValues -> uebersprungen")
                    Continue For
                End If

                Dim procName As String = "usp_Faktenladen_StagetabellenDML_" & v.Verfahren
                Log("  Prozedur: " & procName)

                If Not ProzedurExistiert(connStr, procName) Then
                    Log("  Info: Prozedur [" & procName & "] existiert nicht -> uebersprungen")
                    ProtokollSchreiben(connStr, v.Verfahren, "INFO_SCR11b",
                        "Prozedur " & procName & " nicht vorhanden -> uebersprungen")
                    StatusSetzen(connStr, v.ID, STATUS_OK)
                    cntOK += 1
                    Continue For
                End If

                Log("  Prozedur " & procName & " existiert")

                Try
                    StatusSetzen(connStr, v.ID, STATUS_RUN)

                    Dim protokollID As String = "{" & Guid.NewGuid().ToString() & "}"

                    Dim dmlAnzahl As Integer = 0
                    For Each pe As PartitionsEintrag In meineWerte
                        Dim partWert As String = pe.Wert
                        Dim inTabelle As String = v.Faktentabelle.ToLower() & "_in_" & partWert

                        Log("  Partition      : " & partWert)
                        Log("  Modus          : " & pe.Modus)
                        Log("  _in Tabelle    : " & inTabelle)
                        Log("  ProtokollID    : " & protokollID)

                        Dim sql As String = "EXEC [" & procName & "]" &
                            " '" & partWert & "'" &
                            ", '" & inTabelle & "'" &
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

                        Dim zeilenBetroffen As Integer = SqlAusfuehren(connStr, sql, "DML " & inTabelle)
                        Log("  -> DML ausgefuehrt, Zeilen betroffen: " & zeilenBetroffen.ToString())
                        ProtokollSchreiben(connStr, v.Verfahren, "DML_" & partWert,
                            "DML ausgefuehrt: " & inTabelle & " | Modus: " & pe.Modus & " | Zeilen: " & zeilenBetroffen.ToString())

                        dmlAnzahl += 1
                    Next

                    StatusSetzen(connStr, v.ID, STATUS_OK)
                    ProtokollSchreiben(connStr, v.Verfahren, "SCHRITT_DML",
                        "StagetabellenDML erfolgreich: " & dmlAnzahl.ToString() & " Partition(en)")
                    cntOK += 1
                    Log("  DML erfolgreich fuer " & dmlAnzahl.ToString() & " Partition(en)")

                Catch ex As Exception
                    cntFehler += 1
                    fehlerListe.Add(v.Verfahren)
                    FehlerSetzen(connStr, v.ID, ex.Message)
                    ProtokollSchreiben(connStr, v.Verfahren, "FEHLER_SCR11b", ex.Message)
                    LogFehler("FEHLER '" & v.Verfahren & "': " & ex.Message)
                End Try
            Next

            Log("ZUSAMMENFASSUNG SCR11b_StagetabellenDML_Aufruf")
            Log("Erfolgreich: " & cntOK.ToString())
            Log("Fehler:      " & cntFehler.ToString())
            If fehlerListe.Count > 0 Then
                Log("Fehlerhafte Verfahren:")
                For Each fv As String In fehlerListe
                    Log("  - " & fv)
                Next
            End If

            If cntFehler > 0 Then
                LogFehler("SCR11b: " & cntFehler.ToString() & " Verfahren mit Fehlern")
                Dts.TaskResult = ScriptResults.Failure
            Else
                Log("SCR11b_StagetabellenDML_Aufruf erfolgreich abgeschlossen")
                Dts.TaskResult = ScriptResults.Success
            End If

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
"SELECT a.ID, a.Verfahren, a.Themengebiet, a.LetzterSchritt," &
"       pf.Wert AS Faktentabelle," &
"       pp.Wert AS PartitionsSpalte," &
"       ISNULL(pci.Wert, 'FALSE') AS ClusteredIndex," &
"       ISNULL(pk.Wert, 'PAGE') AS Komprimierung" &
" FROM  dbo.ETL_Fakt_Arbeitsliste a" &
" JOIN  " & _parameterDB & ".dbo." & _parametertab & " pf ON pf.Verfahren=dbo.fn_ParamVerfahren(a.Verfahren) AND pf.Parameter='Faktentabelle'" &
" JOIN  " & _parameterDB & ".dbo." & _parametertab & " pp ON pp.Verfahren=dbo.fn_ParamVerfahren(a.Verfahren) AND pp.Parameter='Faktenpartitionsspalte'" &
" LEFT JOIN " & _parameterDB & ".dbo." & _parametertab & " pci ON pci.Verfahren=dbo.fn_ParamVerfahren(a.Verfahren) AND pci.Parameter='FaktenClusteredIndex'" &
" LEFT JOIN " & _parameterDB & ".dbo." & _parametertab & " pk ON pk.Verfahren=dbo.fn_ParamVerfahren(a.Verfahren) AND pk.Parameter='Faktenkomprimierung'" &
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

    ' -----------------------------------------------------------------------
    ' ProzedurExistiert - Prueft, ob eine Stored Procedure existiert.
    ' -----------------------------------------------------------------------
    Private Function ProzedurExistiert(connStr As String, procName As String) As Boolean
        Dim sql As String =
            "SELECT COUNT(*) FROM sys.objects " &
            "WHERE object_id = OBJECT_ID(@proc) " &
            "AND type IN ('P', 'PC')"

        Dim versuch As Integer = 0
        While versuch < MAX_VERSUCHE
            versuch += 1
            Try
                Using conn As New SqlConnection(connStr)
                    conn.Open()
                    Using cmd As New SqlCommand(sql, conn)
                        cmd.CommandTimeout = 0
                        cmd.Parameters.AddWithValue("@proc", procName)
                        Dim cnt As Integer = Convert.ToInt32(cmd.ExecuteScalar())
                        Return cnt > 0
                    End Using
                End Using
            Catch ex As Exception
                Log(String.Format("WARNUNG [ProzedurExistiert] Versuch {0}/{1}: {2}", versuch, MAX_VERSUCHE, ex.Message))
                If versuch >= MAX_VERSUCHE Then Return False
                System.Threading.Thread.Sleep(WARTE_SEK * 1000)
            End Try
        End While
        Return False
    End Function

    ' -----------------------------------------------------------------------
    ' StatusSetzen - Aktualisiert Status / LetzterSchritt einer
    ' Arbeitslisten-Zeile.
    ' -----------------------------------------------------------------------
    Private Sub StatusSetzen(connStr As String, id As Integer, status As String)
        SqlAusfuehren(connStr,
            "UPDATE dbo.ETL_Fakt_Arbeitsliste SET Status='" & status & "', LetzterSchritt='" & status & "', AktualisiertAm=GETDATE() WHERE ID=" & id,
            "StatusSetzen")
    End Sub

    ' -----------------------------------------------------------------------
    ' FehlerSetzen - Markiert eine Arbeitslisten-Zeile als FEHLER und
    ' speichert die Fehlermeldung.
    ' -----------------------------------------------------------------------
    Private Sub FehlerSetzen(connStr As String, id As Integer, meldung As String)
        Try
            Using conn As New SqlConnection(connStr)
                conn.Open()
                Using cmd As New SqlCommand(
                    "UPDATE dbo.ETL_Fakt_Arbeitsliste SET Status='FEHLER', Fehlermeldung=@m, AktualisiertAm=GETDATE() WHERE ID=@id", conn)
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
    ' PartitionsEintrag - Datencontainer fuer einen Partitionswert.
    ' -----------------------------------------------------------------------
    Private Class PartitionsEintrag
        Public Property Wert As String
        Public Property Modus As String
    End Class

    ' -----------------------------------------------------------------------
    ' VerfahrenInfo - Datencontainer fuer ein Verfahren der Arbeitsliste.
    ' -----------------------------------------------------------------------
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

    ' -----------------------------------------------------------------------
    ' ScriptResults - SSIS-Task-Ergebniscodes.
    ' -----------------------------------------------------------------------
    Public Enum ScriptResults
        Success = DTSExecResult.Success
        Failure = DTSExecResult.Failure
    End Enum

End Class
