Option Explicit On
Option Strict On
Imports System
Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports Microsoft.SqlServer.Dts.Runtime

' =============================================================================
'  Skript       : SCR16_Stagingtabelle___Clustered_Index_erstellen
'  Paket        : Fakten Laden (SSIS)
'  Zweck        : Erstellt den Clustered (Columnstore) Index auf allen _in_-
'                 / _out_-Staging-Tabellen passend zur Struktur der
'                 Faktentabelle.
'  Ablauf       : STAGE_DML_ERFOLG -> INDEX_IN_OUT_ERSTELLT
'  Wiederholung : 3 Versuche je SQL-Anweisung, 30 s Wartezeit
'  Protokoll    : Nur SSIS-Events (FireInformation / FireError)
' =============================================================================
<Microsoft.SqlServer.Dts.Tasks.ScriptTask.SSISScriptTaskEntryPointAttribute()>
<CLSCompliant(False)>
Partial Public Class ScriptMain
    Inherits Microsoft.SqlServer.Dts.Tasks.ScriptTask.VSTARTScriptObjectModelBase

    Private Const SKRIPT_NAME As String = "SCR16_Stagingtabelle___Clustered_Index_erstellen"
    Private Const CONN_NAME As String = "Verbindung"
    Private Const MAX_VERSUCHE As Integer = 3
    Private Const WARTE_SEK As Integer = 30
    Private _runID As Integer = 0
    Private _parameterDB As String = String.Empty
    Private _parametertab As String = String.Empty

    ' -----------------------------------------------------------------------
    ' Main - Einstiegspunkt - steuert den Ablauf des Skripts.
    ' -----------------------------------------------------------------------
    Public Sub Main()
        Log("SCR16_Stagingtabelle___Clustered_Index_erstellen - Start")
        Try
            _runID = Convert.ToInt32(Dts.Variables("BA::RunID").Value)
            _parameterDB = Dts.Variables("BA::ParameterDB").Value.ToString().Trim()
            _parametertab = Dts.Variables("BA::Parametertabelle").Value.ToString().Trim()
            Dim connStr As String = HoleVerbindungszeichenfolge()
            ' Partitionswerte des aktuellen Laufs aus BA::objPartitionValues (gesetzt von SCR09)
            ' Nur diese Tabellen verarbeiten — kein sys.tables LIKE-Scan ueber alle Laeufe
            Dim verfahrenWerte As New Dictionary(Of String, List(Of String))()
            Dim partObjekt As Object = Dts.Variables("BA::objPartitionValues").Value
            If partObjekt IsNot Nothing Then
                Dim partArray(,) As String = CType(partObjekt, String(,))
                For i As Integer = 0 To partArray.GetLength(0) - 1
                    Dim verf As String = partArray(i, 0).Trim().ToLower()
                    Dim wert As String = partArray(i, 1).Trim()
                    If Not verfahrenWerte.ContainsKey(verf) Then verfahrenWerte(verf) = New List(Of String)()
                    verfahrenWerte(verf).Add(wert)
                Next
            End If
            Log("Partitionswerte geladen: " & verfahrenWerte.Count.ToString() & " Verfahren")
            Dim verfahren As List(Of VerfahrenInfo) = VerfahrenLaden(connStr)
            Log("Verfahren: " & verfahren.Count.ToString())
            Dim cntOK As Integer = 0
            Dim cntFehler As Integer = 0
            For Each v As VerfahrenInfo In verfahren
                Log("Verfahren: " & v.Verfahren & " | IndexTyp: " & v.IndexType)
                If v.LetzterSchritt = "INDEX_IN_OUT_ERSTELLT" Then
                    Log("  bereits abgeschlossen uebersprungen OK")
                    Continue For
                End If
                Try
                    StatusSetzen(connStr, v.ID, "INDEX_IN_OUT")
                    ' CI Spalten aus Faktentabelle
                    Dim ciCols As String = Nothing
                    If v.IndexType = "TRUE" Then
                        ciCols = Convert.ToString(SqlSkalar(connStr,
                            "SELECT STUFF((SELECT ', '+QUOTENAME(c.name) FROM sys.index_columns ic JOIN sys.columns c ON c.object_id=ic.object_id AND c.column_id=ic.column_id WHERE ic.object_id=OBJECT_ID('dbo." & v.Faktentabelle & "') AND ic.index_id=1 ORDER BY ic.key_ordinal FOR XML PATH(''),TYPE).value('.','nvarchar(max)'),1,2,'')",
                            "CI Spalten"))
                        If String.IsNullOrEmpty(ciCols) Then ciCols = "[" & v.PartitionColumn & "]"
                    End If
                    ' Nur Tabellen des aktuellen Laufs — exakte Namen aus BA::objPartitionValues
                    Dim verfKey As String = v.Verfahren.Trim().ToLower()
                    Dim werteListe As List(Of String) = Nothing
                    If Not verfahrenWerte.TryGetValue(verfKey, werteListe) OrElse werteListe.Count = 0 Then
                        Log("  WARNUNG: Keine Partitionswerte in BA::objPartitionValues -> uebersprungen")
                        Continue For
                    End If
                    Dim cntTbl As Integer = 0
                    For Each wert As String In werteListe
                        For Each tbl As String In New String() {
                                v.Faktentabelle.ToLower() & "_in_"  & wert,
                                v.Faktentabelle.ToLower() & "_out_" & wert}
                            ' Leere Grenzpartition (keine Oracle-Daten, z.B. 202606): _in_-Tabelle
                            ' wurde von SCR13 nicht erzeugt -> ueberspringen statt Fehler.
                            If Convert.ToInt32(SqlSkalar(connStr, "SELECT CASE WHEN OBJECT_ID('dbo.[" & tbl & "]','U') IS NULL THEN 0 ELSE 1 END", "Tabelle pruefen")) = 0 Then
                                Log("  Tabelle nicht vorhanden (keine Daten geladen) uebersprungen: " & tbl)
                                Continue For
                            End If
                            Log("  Index auf: " & tbl)
                            If v.IndexType = "TRUE" Then
                                If Not IndexVorhanden(connStr, tbl, "CI_" & tbl) Then
                                    SqlAusfuehren(connStr, "CREATE CLUSTERED INDEX [CI_" & tbl & "] ON dbo.[" & tbl & "] (" & ciCols & ") WITH (FILLFACTOR=100,SORT_IN_TEMPDB=ON);", "CI " & tbl)
                                    Log("    CI angelegt OK")
                                Else
                                    Log("    CI bereits vorhanden uebersprungen")
                                End If
                            ElseIf v.IndexType = "CCI" Then
                                If Not IndexVorhanden(connStr, tbl, "CCI_" & tbl) Then
                                    SqlAusfuehren(connStr, "CREATE CLUSTERED COLUMNSTORE INDEX [CCI_" & tbl & "] ON dbo.[" & tbl & "];", "CCI " & tbl)
                                    Log("    CCI angelegt OK")
                                Else
                                    Log("    CCI bereits vorhanden uebersprungen")
                                End If
                            End If
                            cntTbl += 1
                        Next
                    Next
                    StatusSetzen(connStr, v.ID, "INDEX_IN_OUT_ERSTELLT")
                    LogSchreiben(connStr, v.Verfahren, "SCHRITT_7A", "Index _in/_out erstellt: " & cntTbl.ToString() & " Tabellen")
                    cntOK += 1
                    Log("  Schritt 7a abgeschlossen OK")
                Catch ex As Exception
                    cntFehler += 1
                    FehlerSetzen(connStr, v.ID, ex.Message)
                    LogSchreiben(connStr, v.Verfahren, "FEHLER_SCR12", ex.Message)
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
    ' IndexVorhanden - Prueft, ob ein Index auf einer Tabelle existiert.
    ' -----------------------------------------------------------------------
    Private Function IndexVorhanden(connStr As String, tbl As String, idxName As String) As Boolean
        Return Convert.ToInt32(SqlSkalar(connStr, "SELECT COUNT(*) FROM sys.indexes WHERE object_id=OBJECT_ID('dbo." & tbl & "') AND name='" & idxName & "'", "Index prÃ¼fen")) > 0
    End Function

    ' -----------------------------------------------------------------------
    ' VerfahrenLaden - Laedt die zu verarbeitenden Verfahren aus der
    ' Arbeitsliste (Join mit der Parametertabelle).
    ' -----------------------------------------------------------------------
    Private Function VerfahrenLaden(connStr As String) As List(Of VerfahrenInfo)
        Dim liste As New List(Of VerfahrenInfo)()
        Dim sql As String =
"SELECT a.ID, a.Verfahren, a.LetzterSchritt,
        pf.Wert AS Faktentabelle,
        UPPER(pi.Wert) AS IndexType,
        pp.Wert AS PartitionColumn
 FROM   dbo.ETL_Fkt_Arbeitsliste a
 JOIN   " & _parameterDB & ".dbo." & _parametertab & " pf ON pf.Verfahren=a.Verfahren AND pf.Parameter='Faktentabelle'
 JOIN   " & _parameterDB & ".dbo." & _parametertab & " pi ON pi.Verfahren=a.Verfahren AND pi.Parameter='FaktenClusteredIndex'
 JOIN   " & _parameterDB & ".dbo." & _parametertab & " pp ON pp.Verfahren=a.Verfahren AND pp.Parameter='Faktenpartitionsspalte'
 WHERE  a.Status IN ('STAGE_DML_ERFOLG','INDEX_IN_OUT') AND a.RunID=" & _runID & " ORDER BY a.Verfahren"
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
                                    .ID = rdr.GetInt32(0), .Verfahren = rdr(1).ToString().Trim(),
                                    .LetzterSchritt = If(rdr.IsDBNull(2), "", rdr(2).ToString().Trim()),
                                    .Faktentabelle = rdr(3).ToString().Trim(),
                                    .IndexType = rdr(4).ToString().Trim(),
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

    ' -----------------------------------------------------------------------
    ' LogSchreiben - Leitet Protokollmeldungen an SSIS-Events weiter:
    ' FEHLER_* -> FireError, alles andere -> FireInformation.
    ' -----------------------------------------------------------------------
    Private Sub LogSchreiben(connStr As String, verfahren As String, schritt As String, meldung As String)
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
    ' SqlSkalar - Fuehrt eine skalare SQL-Abfrage mit Wiederholung aus;
    ' protokolliert Warnung und vollstaendiges SQL-Statement bei Fehlern.
    ' -----------------------------------------------------------------------
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
                If versuch = 1 Then Log("SQL Statement [" & beschreibung & "]: " & sql)
                If versuch < MAX_VERSUCHE Then System.Threading.Thread.Sleep(WARTE_SEK * 1000) Else Throw
            End Try
        End While
        Return Nothing
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
        Public Property LetzterSchritt As String
        Public Property Faktentabelle As String
        Public Property IndexType As String
        Public Property PartitionColumn As String
    End Class

    ' -----------------------------------------------------------------------
    ' ScriptResults - SSIS-Task-Ergebniscodes.
    ' -----------------------------------------------------------------------
    Public Enum ScriptResults
        Success = DTSExecResult.Success
        Failure = DTSExecResult.Failure
    End Enum

End Class
