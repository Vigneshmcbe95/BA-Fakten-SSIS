Option Explicit On
Option Strict On
Imports System
Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports Microsoft.SqlServer.Dts.Runtime

' =============================================================================
'  Skript       : SCR19_Partitionstausch
'  Paket        : Fakten Laden (SSIS)
'  Zweck        : Fuehrt den Partitionstausch je geladener Partition durch:
'                 SWITCH OUT nach _out_, CHECK-Constraint auf _in_, SWITCH
'                 IN, Aufraeumen und Endstatus.
'  Ablauf       : NCCI_OUT_ERSTELLT -> ERFOLG
'  Wiederholung : 3 Versuche je SQL-Anweisung, 30 s Wartezeit
'  Protokoll    : Nur SSIS-Events (FireInformation / FireError)
' =============================================================================
<Microsoft.SqlServer.Dts.Tasks.ScriptTask.SSISScriptTaskEntryPointAttribute()>
<CLSCompliant(False)>
Partial Public Class ScriptMain
    Inherits Microsoft.SqlServer.Dts.Tasks.ScriptTask.VSTARTScriptObjectModelBase

    Private Const SKRIPT_NAME As String = "SCR19_Partitionstausch"
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
        Log("SCR19_Partitionstausch - Start")
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
                Log("Verfahren: " & v.Verfahren & " | Tabelle: " & v.Faktentabelle)
                If v.LetzterSchritt = "PARTITIONSTAUSCH_ERFOLG" Then
                    Log("  bereits abgeschlossen uebersprungen OK")
                    Continue For
                End If
                Try
                    StatusSetzen(connStr, v.ID, "PARTITIONSTAUSCH")
                    Dim pf As String = "PF_" & v.PartitionColumn & "_" & v.Faktentabelle
                    ' Nur Tabellen des aktuellen Laufs — exakte Namen aus BA::objPartitionValues
                    Dim verfKey As String = v.Verfahren.Trim().ToLower()
                    Dim werteListe As List(Of String) = Nothing
                    If Not verfahrenWerte.TryGetValue(verfKey, werteListe) OrElse werteListe.Count = 0 Then
                        Log("  WARNUNG: Keine Partitionswerte in BA::objPartitionValues -> uebersprungen")
                        Continue For
                    End If
                    Log("  Partitionen: " & werteListe.Count.ToString())
                    For Each pvStr As String In werteListe
                        Dim inTable  As String = v.Faktentabelle.ToLower() & "_in_"  & pvStr
                        Dim outTable As String = v.Faktentabelle.ToLower() & "_out_" & pvStr
                        Log("  Partition: " & pvStr)
                        ' Leere Grenzpartition (keine Oracle-Daten, z.B. 202606): _in_-Tabelle
                        ' wurde von SCR13 nicht erzeugt -> Tausch ueberspringen statt Fehler.
                        ' Leere _out_-Huelle aufraeumen, damit kein Waisenobjekt bleibt.
                        If Convert.ToInt32(SqlSkalar(connStr, "SELECT CASE WHEN OBJECT_ID('dbo.[" & inTable & "]','U') IS NULL THEN 0 ELSE 1 END", "in vorhanden")) = 0 Then
                            Log("  Keine Daten geladen (leere Partition) -> Tausch uebersprungen: " & pvStr)
                            LogSchreiben(connStr, v.Verfahren, "LEER_" & pvStr, "Keine Daten in Oracle - Partitionstausch uebersprungen")
                            If Convert.ToInt32(SqlSkalar(connStr, "SELECT CASE WHEN OBJECT_ID('dbo.[" & outTable & "]','U') IS NOT NULL THEN 1 ELSE 0 END", "out vorhanden")) = 1 Then
                                SqlAusfuehren(connStr, "DROP TABLE dbo.[" & outTable & "];", "drop leeres _out")
                            End If
                            Continue For
                        End If
                        ' Partitionsnummer per $partition-Funktion — gibt direkt die korrekte Nummer zurueck
                        Dim pnr As Object = SqlSkalar(connStr,
                            "SELECT $partition.[" & pf & "](" & pvStr & ")",
                            "Partitionsnummer")
                        If pnr Is Nothing OrElse pnr Is DBNull.Value OrElse Convert.ToInt32(pnr) = 0 Then
                            Log("  FEHLER: Partitionsnummer nicht gefunden fuer: " & pvStr)
                            LogSchreiben(connStr, v.Verfahren, "FEHLER_SWITCH", "Partition nicht gefunden: " & pvStr)
                            Continue For
                        End If
                        Dim pnrVal As Integer = Convert.ToInt32(pnr)
                        Log("  Partitionsnummer: " & pnrVal.ToString())
                        ' SWITCH OUT
                        SqlAusfuehren(connStr, "ALTER TABLE dbo.[" & v.Faktentabelle & "] SWITCH PARTITION " & pnrVal & " TO dbo.[" & outTable & "];", "SWITCH OUT")
                        Log("  SWITCH OUT " & outTable & " OK")
                        ' CHECK Constraint auf _in — beide Grenzen explizit (RANGE LEFT benoetigt > untere Grenze)
                        Dim ckName As String = v.PartitionColumn & "_" & pvStr & "_" & v.Faktentabelle & "_CK"
                        If Convert.ToInt32(SqlSkalar(connStr, "SELECT COUNT(*) FROM sys.check_constraints WHERE parent_object_id=OBJECT_ID('dbo." & inTable & "') AND name='" & ckName & "'", "CK pruefen")) > 0 Then
                            SqlAusfuehren(connStr, "ALTER TABLE dbo.[" & inTable & "] DROP CONSTRAINT [" & ckName & "];", "CK loeschen")
                        End If
                        ' Untere Partitionsgrenze aus PF lesen (groesster Grenzwert der kleiner ist als pvStr)
                        Dim lbObj As Object = SqlSkalar(connStr,
                            "SELECT ISNULL(MAX(CAST(sprv.value AS bigint)), -2147483648) " &
                            "FROM sys.partition_functions spf " &
                            "JOIN sys.partition_range_values sprv ON sprv.function_id=spf.function_id " &
                            "WHERE spf.name='" & pf & "' AND CAST(sprv.value AS bigint) < " & pvStr,
                            "Untere Grenze")
                        Dim lowerBound As Long = Convert.ToInt64(lbObj)
                        SqlAusfuehren(connStr,
                            "ALTER TABLE dbo.[" & inTable & "] ADD CONSTRAINT [" & ckName & "] " &
                            "CHECK([" & v.PartitionColumn & "] IS NOT NULL AND [" & v.PartitionColumn & "] > " & lowerBound.ToString() & " AND [" & v.PartitionColumn & "] <= " & pvStr & ");",
                            "CK setzen")
                        Log("  CHECK Constraint: " & ckName & " (" & v.PartitionColumn & " IS NOT NULL AND > " & lowerBound.ToString() & " AND <= " & pvStr & ") OK")
                        ' SWITCH IN
                        SqlAusfuehren(connStr, "ALTER TABLE dbo.[" & inTable & "] SWITCH TO dbo.[" & v.Faktentabelle & "] PARTITION " & pnrVal & ";", "SWITCH IN")
                        Log("  SWITCH IN " & v.Faktentabelle & " Partition " & pnrVal.ToString() & " OK")
                        ' Cleanup
                        If Convert.ToInt32(SqlSkalar(connStr, "SELECT CASE WHEN OBJECT_ID('dbo." & outTable & "','U') IS NOT NULL THEN 1 ELSE 0 END", "out prÃ¼fen")) = 1 Then
                            SqlAusfuehren(connStr, "DROP TABLE dbo.[" & outTable & "];", "drop _out")
                            Log("  _out geloescht OK")
                        End If
                        If Convert.ToInt32(SqlSkalar(connStr, "SELECT CASE WHEN OBJECT_ID('dbo." & inTable & "','U') IS NOT NULL THEN 1 ELSE 0 END", "in prÃ¼fen")) = 1 Then
                            SqlAusfuehren(connStr, "DROP TABLE dbo.[" & inTable & "];", "drop _in")
                            Log("  _in geloescht OK")
                        End If
                        LogSchreiben(connStr, v.Verfahren, "SWITCH_" & pvStr, "SWITCH IN erfolgreich " & v.Faktentabelle & " Partition " & pnrVal.ToString())
                    Next
                    ' Abschlussstatus MSSQL
                    Dim finalMin As Object = SqlSkalar(connStr, "SELECT MIN([" & v.PartitionColumn & "]) FROM dbo.[" & v.Faktentabelle & "]", "Final MIN")
                    Dim finalMax As Object = SqlSkalar(connStr, "SELECT MAX([" & v.PartitionColumn & "]) FROM dbo.[" & v.Faktentabelle & "]", "Final MAX")
                    Dim finalCnt As Object = SqlSkalar(connStr, "SELECT COUNT_BIG(*) FROM dbo.[" & v.Faktentabelle & "]", "Final COUNT")
                    Log("  ABSCHLUSSSTATUS: " & v.Faktentabelle)
                    Log("  Zeilen: " & Convert.ToString(finalCnt))
                    Log("  MIN:    " & If(finalMin Is Nothing OrElse finalMin Is DBNull.Value, "NULL", Convert.ToString(finalMin)))
                    Log("  MAX:    " & If(finalMax Is Nothing OrElse finalMax Is DBNull.Value, "NULL", Convert.ToString(finalMax)))
                    LogSchreiben(connStr, v.Verfahren, "ABSCHLUSS",
                        "Fertig. Zeilen: " & Convert.ToString(finalCnt) &
                        " | MIN: " & If(finalMin Is Nothing OrElse finalMin Is DBNull.Value, "NULL", Convert.ToString(finalMin)) &
                        " | MAX: " & If(finalMax Is Nothing OrElse finalMax Is DBNull.Value, "NULL", Convert.ToString(finalMax)))
                    StatusSetzenErfolg(connStr, v.ID)
                    cntOK += 1
                    Log("  Verfahren erfolgreich abgeschlossen OK")
                Catch ex As Exception
                    cntFehler += 1
                    FehlerSetzen(connStr, v.ID, ex.Message)
                    LogSchreiben(connStr, v.Verfahren, "FEHLER_SCR15", ex.Message)
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
"SELECT a.ID, a.Verfahren, a.LetzterSchritt, pf.Wert AS Faktentabelle, pp.Wert AS PartitionColumn
 FROM   dbo.ETL_Fakt_Arbeitsliste a
 JOIN   " & _parameterDB & ".dbo." & _parametertab & " pf ON pf.Verfahren=dbo.fn_ParamVerfahren(a.Verfahren) AND pf.Parameter='Faktentabelle'
 JOIN   " & _parameterDB & ".dbo." & _parametertab & " pp ON pp.Verfahren=dbo.fn_ParamVerfahren(a.Verfahren) AND pp.Parameter='Faktenpartitionsspalte'
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

    ' -----------------------------------------------------------------------
    ' StatusSetzen - Aktualisiert Status / LetzterSchritt einer
    ' Arbeitslisten-Zeile.
    ' -----------------------------------------------------------------------
    Private Sub StatusSetzen(connStr As String, id As Integer, status As String)
        SqlAusfuehren(connStr, "UPDATE dbo.ETL_Fakt_Arbeitsliste SET Status='" & status & "',LetzterSchritt='" & status & "',AktualisiertAm=GETDATE() WHERE ID=" & id, "Status")
    End Sub

    ' -----------------------------------------------------------------------
    ' StatusSetzenErfolg - Setzt den Endstatus ERFOLG auf einer
    ' Arbeitslisten-Zeile.
    ' -----------------------------------------------------------------------
    Private Sub StatusSetzenErfolg(connStr As String, id As Integer)
        SqlAusfuehren(connStr, "UPDATE dbo.ETL_Fakt_Arbeitsliste SET Status='ERFOLG',LetzterSchritt='ERFOLG',AktualisiertAm=GETDATE() WHERE ID=" & id, "ERFOLG")
    End Sub

    ' -----------------------------------------------------------------------
    ' FehlerSetzen - Markiert eine Arbeitslisten-Zeile als FEHLER und
    ' speichert die Fehlermeldung.
    ' -----------------------------------------------------------------------
    Private Sub FehlerSetzen(connStr As String, id As Integer, msg As String)
        Try
            Using conn As New SqlConnection(connStr)
                conn.Open()
                Using cmd As New SqlCommand("UPDATE dbo.ETL_Fakt_Arbeitsliste SET Status='FEHLER',Fehlermeldung=@m,AktualisiertAm=GETDATE() WHERE ID=@id", conn)
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
