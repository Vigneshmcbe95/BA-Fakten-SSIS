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

' =============================================================================
'  Skript       : SCR13_Daten_Laden
'  Paket        : Fakten Laden (SSIS)
'  Zweck        : Laedt die Daten partitionsweise aus ext.<fakt> in
'                 _in_-Tabellen: SELECT INTO _LOADING, danach atomares
'                 Umbenennen. Parallel und wiederaufsetzbar.
'  Ablauf       : STAGING_ERSTELLT -> DATEN_GELADEN
'  Wiederholung : 3 Versuche je SQL-Anweisung, 30 s Wartezeit
'  Protokoll    : Nur SSIS-Events (FireInformation / FireError)
' =============================================================================
<Microsoft.SqlServer.Dts.Tasks.ScriptTask.SSISScriptTaskEntryPointAttribute()>
<CLSCompliant(False)>
Partial Public Class ScriptMain
    Inherits Microsoft.SqlServer.Dts.Tasks.ScriptTask.VSTARTScriptObjectModelBase

    Private Const SKRIPT_NAME As String = "SCR13_Daten_Laden"
    Private Const CONN_NAME As String = "Verbindung"
    Private Const MAX_VERSUCHE As Integer = 3
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

    ' -----------------------------------------------------------------------
    ' Main - Einstiegspunkt - steuert den Ablauf des Skripts.
    ' -----------------------------------------------------------------------
    Public Sub Main()

        Log("SCR13_Daten_Laden - Start")

        Try
            _runID = Convert.ToInt32(Dts.Variables("BA::RunID").Value)
            _parameterDB = Dts.Variables("BA::ParameterDB").Value.ToString().Trim()
            _parametertab = Dts.Variables("BA::Parametertabelle").Value.ToString().Trim()
            _datenbank = Dts.Variables("BA::Datenbank").Value.ToString().Trim()
            _maxparallel = CInt(Dts.Variables("BA::Maxparallel").Value)


            ' ─────────────────────────────────────────────────────────────────
            ' Partitionswerte aus BA::objPartitionValues lesen
            ' ─────────────────────────────────────────────────────────────────
            Dim partDict As Dictionary(Of String, List(Of String)) = LesePartitionValues()
            Log("Partitionswerte gesamt: " & partDict.Values.Sum(Function(l) l.Count).ToString() & " Eintraege")

            Dim connStr As String = HoleVerbindungszeichenfolge()
            Dim verfahren As List(Of VerfahrenInfo) = VerfahrenLaden(connStr)
            Log("Verfahren: " & verfahren.Count.ToString())

            If verfahren.Count = 0 Then
                Log("Keine Verfahren zur Verarbeitung vorhanden.")
                Dts.TaskResult = ScriptResults.Success
                Return
            End If

            ' ─────────────────────────────────────────────────────────────────
            ' Phase 1 (sequentiell): Arbeitspakete (Verfahren x Partition)
            ' aufbauen - EIN gemeinsames Parallel-Limit statt verschachtelter
            ' Parallel.ForEach (verhindert Ueberbuchung der Oracle-Sessions:
            ' frueher Verfahren x Partitionen = bis zu Maxparallel^2 Sessions)
            ' ─────────────────────────────────────────────────────────────────
            Dim arbeitspakete As New List(Of ArbeitsPaket)()
            Dim aktiveVerfahren As New List(Of VerfahrenInfo)()

            For Each v As VerfahrenInfo In verfahren
                Log("Verfahren: " & v.Verfahren & " | Spalte: " & v.PartitionColumn)

                If v.LetzterSchritt = "DATEN_GELADEN" Then
                    Log("  bereits abgeschlossen uebersprungen OK")
                    Continue For
                End If

                Dim partWerte As List(Of String) = Nothing
                If Not partDict.TryGetValue(v.Verfahren, partWerte) OrElse partWerte.Count = 0 Then
                    Log("  WARNUNG: Keine Partitionswerte in BA::objPartitionValues fuer '" & v.Verfahren & "' uebersprungen")
                    Continue For
                End If

                Log("  Partitionswerte: " & partWerte.Count.ToString())

                Try
                    StatusSetzen(connStr, v.ID, "DATEN_LADEN")
                    aktiveVerfahren.Add(v)
                    For Each pv As String In partWerte
                        arbeitspakete.Add(New ArbeitsPaket With {.V = v, .Pv = pv})
                    Next
                Catch ex As Exception
                    Interlocked.Increment(_cntFehler)
                    _fehlerListe.Add("FEHLER '" & v.Verfahren & "': " & ex.Message)
                    FehlerSetzen(connStr, v.ID, ex.Message)
                    LogFehler("FEHLER '" & v.Verfahren & "': " & ex.Message)
                End Try
            Next

            Log("Arbeitspakete gesamt (Verfahren x Partition): " & arbeitspakete.Count.ToString() &
                " | Parallelitaet: " & _maxparallel.ToString())

            ' ─────────────────────────────────────────────────────────────────
            ' Phase 2 (parallel): alle Partitionen aller Verfahren ueber EINE
            ' Drossel laden - maximal _maxparallel gleichzeitige Oracle-Sessions
            ' ─────────────────────────────────────────────────────────────────
            Dim zeilenJeVerfahren As New ConcurrentDictionary(Of Integer, Long)()
            Dim fehlerJeVerfahren As New ConcurrentDictionary(Of Integer, Integer)()

            Dim opts As New ParallelOptions With {.MaxDegreeOfParallelism = _maxparallel}
            Parallel.ForEach(arbeitspakete, opts,
                Sub(p As ArbeitsPaket)
                    Try
                        Dim zeilen As Long = VerarbeitePartition(p.V, p.Pv, connStr)
                        zeilenJeVerfahren.AddOrUpdate(p.V.ID, zeilen, Function(k, alt) alt + zeilen)
                    Catch ex As Exception
                        fehlerJeVerfahren.AddOrUpdate(p.V.ID, 1, Function(k, alt) alt + 1)
                        Interlocked.Increment(_cntFehler)
                        _fehlerListe.Add("FEHLER '" & p.V.Verfahren & "' pv=" & p.Pv & ": " & ex.Message)
                        LogFehler("FEHLER bei pv=" & p.Pv & " (" & p.V.Faktentabelle.ToLower() & "_in_" & p.Pv & "): " & ex.Message)
                    End Try
                End Sub)

            ' ─────────────────────────────────────────────────────────────────
            ' Phase 3 (sequentiell): Status je Verfahren abschliessen
            ' ─────────────────────────────────────────────────────────────────
            For Each v As VerfahrenInfo In aktiveVerfahren
                Dim fehlerAnzahl As Integer = 0
                fehlerJeVerfahren.TryGetValue(v.ID, fehlerAnzahl)
                Dim gesamtZeilen As Long = 0
                zeilenJeVerfahren.TryGetValue(v.ID, gesamtZeilen)

                If fehlerAnzahl > 0 Then
                    FehlerSetzen(connStr, v.ID, fehlerAnzahl.ToString() & " Partition(en) fehlgeschlagen")
                    LogSchreiben(connStr, v.Verfahren, "FEHLER_SCR13",
                        fehlerAnzahl.ToString() & " Partition(en) fehlgeschlagen")
                Else
                    StatusSetzen(connStr, v.ID, "DATEN_GELADEN")
                    LogSchreiben(connStr, v.Verfahren, "SCHRITT_6",
                        "Daten geladen. Zeilen gesamt: " & gesamtZeilen.ToString())
                    Interlocked.Increment(_cntOK)
                    Log("  Schritt 6 abgeschlossen OK (" & v.Verfahren & ")")
                End If
            Next

            Log("Erfolgreich: " & _cntOK.ToString() & " | Fehler: " & _cntFehler.ToString())

            If _fehlerListe.Count > 0 Then
                Log("FEHLER-DETAILS:")
                For Each f As String In _fehlerListe.OrderBy(Function(x) x)
                    Log("  " & f)
                Next
            End If

            Dts.TaskResult = If(_cntFehler > 0, ScriptResults.Failure, ScriptResults.Success)

        Catch ex As Exception
            LogFehler("Kritischer Fehler: " & ex.Message)
            Dts.TaskResult = ScriptResults.Failure
        End Try

    End Sub

    ' -----------------------------------------------------------------------
    ' LesePartitionValues - Liest BA::objPartitionValues in ein Dictionary
    ' je Verfahren ein.
    ' -----------------------------------------------------------------------
    Private Function LesePartitionValues() As Dictionary(Of String, List(Of String))
        Dim dict As New Dictionary(Of String, List(Of String))(StringComparer.OrdinalIgnoreCase)
        Dim partObjekt As Object = Dts.Variables("BA::objPartitionValues").Value
        If partObjekt Is Nothing Then
            Log("BA::objPartitionValues ist leer (Nothing)")
            Return dict
        End If
        Dim partArray(,) As String = CType(partObjekt, String(,))
        For i As Integer = 0 To partArray.GetLength(0) - 1
            Dim verf As String = partArray(i, 0)
            Dim wert As String = partArray(i, 1)
            If Not dict.ContainsKey(verf) Then dict(verf) = New List(Of String)()
            dict(verf).Add(wert)
        Next
        Return dict
    End Function

    ' -----------------------------------------------------------------------
    ' VerarbeitePartition - Laedt EINE Partition eines Verfahrens (paralleler
    ' Worker, ein Arbeitspaket). Liefert die geladenen / vorhandenen Zeilen
    ' zurueck. Resume: bereits gefuellte _in_-Tabellen werden uebersprungen.
    ' -----------------------------------------------------------------------
    Private Function VerarbeitePartition(v As VerfahrenInfo, pv As String,
                                         connStr As String) As Long

        Dim inTable As String = v.Faktentabelle.ToLower() & "_in_" & pv
        Dim loadingTable As String = inTable & "_LOADING"

        ' ─── RESUME: inTable existiert UND hat Zeilen → bereits geladen ───
        Dim inExists As Boolean = Convert.ToInt32(SqlSkalar(connStr,
            "SELECT COUNT(*) FROM sys.tables WHERE schema_id=SCHEMA_ID('dbo') AND name='" & inTable & "'",
            "inTable check")) > 0

        If inExists Then
            Dim zeilenCount As Integer = Convert.ToInt32(SqlSkalar(connStr,
                "SELECT COUNT(*) FROM dbo.[" & inTable & "]",
                "rows check"))
            If zeilenCount > 0 Then
                Log("  " & inTable & " bereits gefuellt (" & zeilenCount.ToString() & " Zeilen) uebersprungen OK")
                Return CLng(zeilenCount)
            End If
        End If

        ' ─── CLEANUP: _LOADING von abgebrochenem Lauf loeschen ───
        SqlAusfuehren(connStr,
            "IF OBJECT_ID('dbo.[" & loadingTable & "]','U') IS NOT NULL DROP TABLE dbo.[" & loadingTable & "];",
            "DROP _LOADING " & pv)

        ' ─── CLEANUP: leere _in_ Tabelle loeschen ───
        SqlAusfuehren(connStr,
            "IF OBJECT_ID('dbo.[" & inTable & "]','U') IS NOT NULL DROP TABLE dbo.[" & inTable & "];",
            "DROP _in_ leer " & pv)

        ' ─── Ext-Tabelle pruefen ───
        ' Oracle-Quelle = Verfahrensname aus der Steuerliste (Oracle-Objektname,
        ' z.B. View vf_stea). Geladen wird in die Zieltabelle (_in_<Faktentabelle>).
        Dim extTable As String = v.Verfahren.ToLower()
        Dim extExists As Boolean = Convert.ToInt32(SqlSkalar(connStr,
            "SELECT COUNT(*) FROM sys.external_tables WHERE schema_id=SCHEMA_ID('ext') AND name='" & extTable & "'",
            "ext pruefen")) > 0

        If Not extExists Then
            Throw New Exception("ext." & extTable & " nicht gefunden")
        End If

        ' ─── SELECT INTO _LOADING (atomar) ───
        Dim zeilen As Integer = DatenLadenSelectInto(connStr, v, loadingTable, extTable, pv)

        ' ─── 0 Zeilen darf hier NICHT vorkommen ───
        ' SCR11 uebergibt nur Partitionswerte MIT Daten - leere Grenzpartitionen
        ' (z.B. unterhalb des MSSQL-Minimums oder die offene obere Grenze) werden
        ' bereits dort ausgeschlossen. Kommt hier dennoch eine Partition ohne
        ' Daten an, ist das ein echter Fehler (falsche Partitionsspalte, fehlende
        ' Oracle-Daten, falscher Verfahrensname o.ae.) -> Lauf mit klarer Meldung
        ' abbrechen statt still zu ueberspringen.
        If zeilen = 0 Then
            SqlAusfuehren(connStr,
                "IF OBJECT_ID('dbo.[" & loadingTable & "]','U') IS NOT NULL DROP TABLE dbo.[" & loadingTable & "];",
                "DROP _LOADING leer " & pv)
            Throw New Exception("Partition " & v.PartitionColumn & " = " & pv & ": [ext." & extTable &
                "] lieferte 0 Zeilen. Diese Partition haette nach der SCR11-Berechnung Daten enthalten muessen. " &
                "Bitte pruefen: (1) existiert der Wert in Oracle (SELECT DISTINCT " & v.PartitionColumn &
                " FROM ext.[" & extTable & "])? (2) ist die Faktenpartitionsspalte [" & v.PartitionColumn &
                "] in der Parametertabelle fuer Verfahren [" & v.Verfahren & "] korrekt?")
        End If

        ' ─── RENAME _LOADING → _in_ (atomar abschliessen) ───
        SqlAusfuehren(connStr,
            "EXEC sp_rename 'dbo.[" & loadingTable & "]', '" & inTable & "';",
            "RENAME " & pv)

        Log("  Zeilen geladen (" & inTable & "): " & zeilen.ToString())
        LogSchreiben(connStr, v.Verfahren, "LOAD_" & pv,
            "SELECT INTO dbo." & inTable & " | Zeilen: " & zeilen.ToString())

        Return CLng(zeilen)

    End Function

    ' -----------------------------------------------------------------------
    ' DatenLadenSelectInto - Laedt eine Partition per SELECT INTO _LOADING
    ' und benennt sie atomar in _in_ um.
    ' -----------------------------------------------------------------------
    Private Function DatenLadenSelectInto(connStr As String, v As VerfahrenInfo,
                                          loadingTable As String, extTable As String,
                                          partitionValue As String) As Integer

        ' SELECT-Liste aus columns_dbo (tm_polybase_struktur): die
        ' ISNULL/CONVERT-Ausdruecke erzeugen exakt die dbo-Typen und die
        ' Nullability des Templates / der Faktentabelle. Blosse Spaltennamen
        ' FROM ext.[] wuerden Typen/Nullability der ext-Tabelle erben
        ' (alles NULL) -> SWITCH IN scheitert an abweichender Nullability.
        Dim selectList As String = Nothing
        Dim sqlCols As String =
            "SELECT STRING_AGG(CAST(columns_dbo AS nvarchar(max)), ',' + CHAR(13) + CHAR(10)) WITHIN GROUP (ORDER BY colno) " &
            "FROM dbo.tm_polybase_struktur " &
            "WHERE tabname = '" & v.Verfahren.ToLower().Replace("'", "''") & "' " &
            "AND themengebiet = '" & v.Themengebiet.Trim().ToLower().Replace("'", "''") & "'"

        Dim versuch As Integer = 0
        While versuch < MAX_VERSUCHE
            versuch += 1
            Try
                Using conn As New SqlConnection(connStr)
                    conn.Open()
                    Using cmd As New SqlCommand(sqlCols, conn)
                        cmd.CommandTimeout = 0
                        Dim result As Object = cmd.ExecuteScalar()
                        If result IsNot Nothing AndAlso result IsNot DBNull.Value Then
                            selectList = result.ToString()
                        End If
                    End Using
                End Using
                Exit While
            Catch ex As Exception
                SyncLock _logSperre
                    Log(String.Format("WARNUNG [Template Spalten] Versuch {0}/{1}: {2}", versuch, MAX_VERSUCHE, ex.Message))
                End SyncLock
                If versuch < MAX_VERSUCHE Then Thread.Sleep(WARTE_SEK * 1000) Else Throw
            End Try
        End While

        If String.IsNullOrEmpty(selectList) Then
            Throw New Exception("columns_dbo nicht gefunden in tm_polybase_struktur fuer " &
                                v.Themengebiet & "." & v.Faktentabelle.ToLower())
        End If

        Dim sql As String =
            "SELECT" & vbCrLf &
            selectList & vbCrLf &
            "INTO dbo.[" & loadingTable & "]" & vbCrLf &
            "FROM ext.[" & extTable & "]" & vbCrLf &
            "WHERE [" & v.PartitionColumn & "] = " & partitionValue & ";"

        SyncLock _logSperre
            Log("  SELECT INTO [" & loadingTable & "] WHERE " & v.PartitionColumn & " = " & partitionValue & " (wird zu _in_ umbenannt)")
        End SyncLock

        SqlAusfuehren(connStr, sql, "SELECT INTO " & loadingTable)

        Return Convert.ToInt32(SqlSkalar(connStr,
            "SELECT COUNT(*) FROM dbo.[" & loadingTable & "]",
            "Count " & loadingTable))
    End Function

    ' -----------------------------------------------------------------------
    ' VerfahrenLaden - Laedt die zu verarbeitenden Verfahren aus der
    ' Arbeitsliste (Join mit der Parametertabelle).
    ' -----------------------------------------------------------------------
    Private Function VerfahrenLaden(connStr As String) As List(Of VerfahrenInfo)
        Dim liste As New List(Of VerfahrenInfo)()
        Dim sql As String =
            "SELECT a.ID, a.Verfahren, a.Themengebiet, a.LetzterSchritt," &
            "       pf.Wert AS Faktentabelle, pp.Wert AS PartitionColumn" &
            " FROM   dbo.ETL_Fkt_Arbeitsliste a" &
            " JOIN   " & _parameterDB & ".dbo." & _parametertab & " pf ON pf.Verfahren=dbo.fn_ParamVerfahren(a.Verfahren) AND pf.Parameter='Faktentabelle'" &
            " JOIN   " & _parameterDB & ".dbo." & _parametertab & " pp ON pp.Verfahren=dbo.fn_ParamVerfahren(a.Verfahren) AND pp.Parameter='Faktenpartitionsspalte'" &
            " WHERE  a.Status IN ('STAGING_ERSTELLT','DATEN_LADEN')" &
            " AND    a.RunID = " & _runID & " ORDER BY a.Verfahren"

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

    ' -----------------------------------------------------------------------
    ' StatusSetzen - Aktualisiert Status / LetzterSchritt einer
    ' Arbeitslisten-Zeile.
    ' -----------------------------------------------------------------------
    Private Sub StatusSetzen(connStr As String, id As Integer, status As String)
        SqlAusfuehren(connStr,
            "UPDATE dbo.ETL_Fkt_Arbeitsliste SET Status='" & status &
            "',LetzterSchritt='" & status & "',AktualisiertAm=GETDATE() WHERE ID=" & id,
            "Status")
    End Sub

    ' -----------------------------------------------------------------------
    ' FehlerSetzen - Markiert eine Arbeitslisten-Zeile als FEHLER und
    ' speichert die Fehlermeldung.
    ' -----------------------------------------------------------------------
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
                SyncLock _logSperre
                    Log(String.Format("WARNUNG [{0}] Versuch {1}/{2}: {3}", beschreibung, versuch, MAX_VERSUCHE, ex.Message))
                If versuch = 1 Then Log("SQL Statement [" & beschreibung & "]: " & sql)
                End SyncLock
                If versuch < MAX_VERSUCHE Then Thread.Sleep(WARTE_SEK * 1000)
            End Try
        End While
        Throw New Exception(String.Format("[{0}] fehlgeschlagen: {1}", beschreibung,
            If(letzterFehler IsNot Nothing, letzterFehler.Message, "Unbekannt")))
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
                SyncLock _logSperre
                    Log(String.Format("WARNUNG [{0}] Versuch {1}/{2}: {3}", beschreibung, versuch, MAX_VERSUCHE, ex.Message))
                If versuch = 1 Then Log("SQL Statement [" & beschreibung & "]: " & sql)
                End SyncLock
                If versuch < MAX_VERSUCHE Then Thread.Sleep(WARTE_SEK * 1000) Else Throw
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
        SyncLock _logSperre
            Dim f As Boolean = False
            Dts.Events.FireInformation(0, SKRIPT_NAME, n, "", 0, f)
        End SyncLock
    End Sub

    ' -----------------------------------------------------------------------
    ' LogFehler - Schreibt eine Fehlermeldung in das SSIS-Protokoll
    ' (FireError).
    ' -----------------------------------------------------------------------
    Private Sub LogFehler(n As String)
        SyncLock _logSperre
            Dts.Events.FireError(0, SKRIPT_NAME, n, "", 0)
        End SyncLock
    End Sub

    ' -----------------------------------------------------------------------
    ' ArbeitsPaket - Datencontainer fuer ein Arbeitspaket
    ' (Verfahren x Partitionswert) der flachen Parallel-Verarbeitung.
    ' -----------------------------------------------------------------------
    Private Class ArbeitsPaket
        Public Property V As VerfahrenInfo
        Public Property Pv As String
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
