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
'                 Optional je Tabelle (Parameter FaktenStageCCIInsert = TRUE,
'                 nur bei CCI): _LOADING leer mit CCI anlegen, dann
'                 INSERT INTO ... WITH (TABLOCK) - Daten werden direkt
'                 komprimiert geladen.
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
    Private Const MAX_FEHLER_OHNE_ERFOLG As Integer = 5
    Private Const MAX_VERBINDUNGSFEHLER_FOLGE As Integer = 20
    Private Const VORAB_MB_JE_STREAM As Integer = 10240
    Private _verbindungsFehlerFolge As Integer = 0
    Private _globalerAbbruch As Boolean = False
    Private _abbruchGrund As String = ""

    ' SQL-Fehler 1105 und 1101 = "Dateigruppe voll". Tritt auf, wenn mehrere
    ' parallele SELECT INTO gleichzeitig Platz brauchen, waehrend die
    ' Datendatei waechst (SQL Server laesst je Datei nur EIN Wachstum
    ' gleichzeitig zu - die anderen Sessions bekommen den Fehler, statt zu warten).
    Private Const SQL_FEHLER_FILEGROUP_VOLL As Integer = 1105
    Private Const SQL_FEHLER_SEITE_NICHT_ZUWEISBAR As Integer = 1101

    Private Const DATEIWACHSTUM_SCHRITT_MB As Integer = 10240 ' 10 GB je Wachstumsschritt
    Private ReadOnly _wachstumsSperre As New Object()

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

                Log("  Partitionswerte: " & partWerte.Count.ToString() &
                    " | Ladeart: " & If(v.StageCciInsert, "CCI-INSERT (FaktenStageCCIInsert)", "SELECT INTO"))

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
            Dim erledigtJeVerfahren As New ConcurrentDictionary(Of Integer, Integer)()
            Dim abbruchVerfahren As New ConcurrentDictionary(Of Integer, Boolean)()
            Dim erwartetJeVerfahren As Dictionary(Of Integer, Integer) =
                arbeitspakete.GroupBy(Function(p) p.V.ID).ToDictionary(Function(g) g.Key, Function(g) g.Count())

            DateienVorabVergroessern(connStr)

            ' Legacy Cardinality Estimator nur fuer den Ladevorgang deaktivieren,
            ' falls er aktiv ist - danach wieder in den Ausgangszustand
            ' zurueckversetzen. War er bereits deaktiviert, bleibt er unangetastet.
            Dim legacyCEWarAn As Boolean = LegacyCEPruefenUndDeaktivieren(connStr)
            Try
                Dim opts As New ParallelOptions With {.MaxDegreeOfParallelism = _maxparallel}
                Parallel.ForEach(arbeitspakete, opts,
                    Sub(p As ArbeitsPaket, schleife As ParallelLoopState)
                        If _globalerAbbruch OrElse abbruchVerfahren.ContainsKey(p.V.ID) Then Return
                        Try
                            Dim zeilen As Long = VerarbeitePartition(p.V, p.Pv, connStr)
                            zeilenJeVerfahren.AddOrUpdate(p.V.ID, zeilen, Function(k, alt) alt + zeilen)
                            erledigtJeVerfahren.AddOrUpdate(p.V.ID, 1, Function(k, alt) alt + 1)
                            Interlocked.Exchange(_verbindungsFehlerFolge, 0)
                        Catch ex As Exception
                            Dim fehlerAnz As Integer = fehlerJeVerfahren.AddOrUpdate(p.V.ID, 1, Function(k, alt) alt + 1)
                            Interlocked.Increment(_cntFehler)
                            _fehlerListe.Add("FEHLER '" & p.V.Verfahren & "' pv=" & p.Pv & ": " & ex.Message)
                            LogFehler("FEHLER bei pv=" & p.Pv & " (" & p.V.Faktentabelle.ToLower() & "_in_" & p.Pv & "): " & ex.Message)
                            If Not zeilenJeVerfahren.ContainsKey(p.V.ID) AndAlso fehlerAnz >= MAX_FEHLER_OHNE_ERFOLG Then
                                If abbruchVerfahren.TryAdd(p.V.ID, True) Then Log("Verfahren abgebrochen: " & p.V.Verfahren)
                            End If
                            If IstVerbindungsfehler(ex) Then
                                If Interlocked.Increment(_verbindungsFehlerFolge) >= MAX_VERBINDUNGSFEHLER_FOLGE Then
                                    _abbruchGrund = "Verbindungsfehler in Folge"
                                    _globalerAbbruch = True
                                    schleife.Stop()
                                End If
                            End If
                        End Try
                    End Sub)
            Finally
                If legacyCEWarAn Then LegacyCEWiederAktivieren(connStr)
            End Try

            If _globalerAbbruch Then
                LogFehler("Abbruch: " & _abbruchGrund)
                Interlocked.Increment(_cntFehler)
            End If

            ' ─────────────────────────────────────────────────────────────────
            ' Phase 3 (sequentiell): Status je Verfahren abschliessen
            ' ─────────────────────────────────────────────────────────────────
            For Each v As VerfahrenInfo In aktiveVerfahren
                Dim fehlerAnzahl As Integer = 0
                fehlerJeVerfahren.TryGetValue(v.ID, fehlerAnzahl)
                Dim gesamtZeilen As Long = 0
                zeilenJeVerfahren.TryGetValue(v.ID, gesamtZeilen)

                Dim erledigt As Integer = 0
                erledigtJeVerfahren.TryGetValue(v.ID, erledigt)
                Dim erwartet As Integer = 0
                erwartetJeVerfahren.TryGetValue(v.ID, erwartet)

                If fehlerAnzahl > 0 Then
                    FehlerSetzen(connStr, v.ID, fehlerAnzahl.ToString() & " Partition(en) fehlgeschlagen")
                    LogSchreiben(connStr, v.Verfahren, "FEHLER_SCR13",
                        fehlerAnzahl.ToString() & " Partition(en) fehlgeschlagen")
                ElseIf erledigt < erwartet Then
                    FehlerSetzen(connStr, v.ID, (erwartet - erledigt).ToString() & " Partition(en) offen")
                    LogSchreiben(connStr, v.Verfahren, "FEHLER_SCR13",
                        (erwartet - erledigt).ToString() & " Partition(en) offen")
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
    ' LegacyCEPruefenUndDeaktivieren - Prueft den aktuellen Zustand von
    ' LEGACY_CARDINALITY_ESTIMATION. War er AN, wird er fuer den Ladevorgang
    ' AUS geschaltet und True zurueckgegeben (spaeter wieder herstellen). War
    ' er bereits AUS, bleibt er unangetastet und es wird False zurueckgegeben.
    ' -----------------------------------------------------------------------
    Private Function LegacyCEPruefenUndDeaktivieren(connStr As String) As Boolean
        Try
            Dim wert As Object = SqlSkalar(connStr,
                "SELECT value FROM sys.database_scoped_configurations WHERE name = 'LEGACY_CARDINALITY_ESTIMATION';",
                "Legacy CE Status pruefen")

            If wert IsNot Nothing AndAlso wert IsNot DBNull.Value AndAlso Convert.ToInt32(wert) = 1 Then
                SqlAusfuehren(connStr,
                    "ALTER DATABASE SCOPED CONFIGURATION SET LEGACY_CARDINALITY_ESTIMATION = OFF;",
                    "Legacy CE deaktivieren")
                Log("Legacy Cardinality Estimator war AN - fuer den Ladevorgang AUS geschaltet.")
                Return True
            End If
        Catch ex As Exception
            Log("WARNUNG: Legacy CE Status konnte nicht geprueft/gesetzt werden: " & ex.Message)
        End Try
        Return False
    End Function

    ' -----------------------------------------------------------------------
    ' LegacyCEWiederAktivieren - Schaltet LEGACY_CARDINALITY_ESTIMATION wieder
    ' AN (nur aufgerufen, wenn er vor dem Laden AN war).
    ' -----------------------------------------------------------------------
    Private Sub LegacyCEWiederAktivieren(connStr As String)
        Try
            SqlAusfuehren(connStr,
                "ALTER DATABASE SCOPED CONFIGURATION SET LEGACY_CARDINALITY_ESTIMATION = ON;",
                "Legacy CE wieder aktivieren")
            Log("Legacy Cardinality Estimator wieder AN geschaltet.")
        Catch ex As Exception
            LogFehler("WARNUNG: Legacy CE konnte nicht wiederhergestellt werden: " & ex.Message)
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

        ' Leere Partition wird uebertragen
        If zeilen = 0 Then
            Log("  Info: " & v.PartitionColumn & " = " & pv & " liefert 0 Zeilen -> leere Partition")
            LogSchreiben(connStr, v.Verfahren, "LEER_" & pv, "0 Zeilen - leere Partition wird uebertragen")
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

        ' SELECT-Liste direkt aus der FAKTENTABELLE (sys.columns) erzeugen:
        ' Jede Spalte wird per CONVERT exakt auf den Ziel-Datentyp der
        ' Faktentabelle gebracht (Typ, Praezision/Skala, Laenge, Collation) und
        ' bei NOT NULL per ISNULL mit Default geschlossen. Dadurch ist die
        ' _in_-Staging-Struktur 1:1 identisch zum SWITCH-Ziel - keine
        ' Typ-/Praezision-/Nullability-/Collation-Abweichung mehr (z.B.
        ' decimal(38,0) der ext-Quelle vs. decimal(17,10) der Faktentabelle:
        ' der Wert wird auf decimal(17,10) konvertiert). Die Faktentabelle ist
        ' die massgebliche Struktur (Template wird manuell gepflegt).
        Dim selectList As String = Nothing
        Dim sqlCols As String =
"SELECT STRING_AGG(CAST(
    CASE WHEN c.is_nullable = 0 THEN 'ISNULL(' ELSE '' END
  + 'CONVERT(' + ty.name
  + CASE WHEN ty.name IN ('varchar','char') THEN '(' + CASE WHEN c.max_length = -1 THEN 'MAX' ELSE CAST(c.max_length AS varchar(10)) END + ')'
         WHEN ty.name IN ('nvarchar','nchar') THEN '(' + CASE WHEN c.max_length = -1 THEN 'MAX' ELSE CAST(c.max_length/2 AS varchar(10)) END + ')'
         WHEN ty.name IN ('decimal','numeric') THEN '(' + CAST(c.precision AS varchar(10)) + ',' + CAST(c.scale AS varchar(10)) + ')'
         WHEN ty.name IN ('datetime2','time','datetimeoffset') THEN '(' + CAST(c.scale AS varchar(10)) + ')'
         WHEN ty.name = 'varbinary' THEN '(' + CASE WHEN c.max_length = -1 THEN 'MAX' ELSE CAST(c.max_length AS varchar(10)) END + ')'
         ELSE '' END
  + ', [' + c.name + '])'
  + CASE WHEN ty.name IN ('varchar','char','nvarchar','nchar') AND c.collation_name IS NOT NULL
         THEN ' COLLATE ' + c.collation_name ELSE '' END
  + CASE WHEN c.is_nullable = 0 THEN
            CASE WHEN ty.name LIKE '%char%' THEN ', '''')'
                 WHEN ty.name LIKE '%date%' OR ty.name IN ('time','datetimeoffset') THEN ', ''1900-01-01'')'
                 WHEN ty.name IN ('float','real','decimal','numeric','int','bigint','smallint','tinyint','bit','money','smallmoney') THEN ', 0)'
                 ELSE ', NULL)' END
        ELSE '' END
  + ' AS [' + c.name + ']'
  AS nvarchar(max)), ',' + CHAR(13) + CHAR(10)) WITHIN GROUP (ORDER BY c.column_id)
FROM sys.columns c
JOIN sys.types ty ON ty.user_type_id = c.user_type_id
WHERE c.object_id = OBJECT_ID('dbo.[" & v.Faktentabelle.ToLower().Replace("'", "''") & "]') AND c.is_computed = 0;"

        ' =====================================================================
        ' ALTERNATIVE (DEAKTIVIERT): SELECT-Liste direkt aus columns_dbo der
        ' Metadatentabelle dbo.tm_polybase_struktur lesen (urspruengliche
        ' Variante). columns_dbo wird in SCR07 aus den Oracle-DDL-Metadaten
        ' aufgebaut - ein einzelner Metadaten-Read statt Aufbau aus sys.columns.
        '
        '   Dim sqlCols As String =
        '       "SELECT STRING_AGG(CAST(columns_dbo AS nvarchar(max)), ',' + CHAR(13) + CHAR(10)) WITHIN GROUP (ORDER BY colno) " &
        '       "FROM dbo.tm_polybase_struktur " &
        '       "WHERE tabname = '" & v.Verfahren.ToLower().Replace("'", "''") & "' " &
        '       "AND themengebiet = '" & v.Themengebiet.Trim().ToLower().Replace("'", "''") & "'"
        '
        ' WARUM DERZEIT DEAKTIVIERT:
        '   columns_dbo erbt Typ/Praezision/Skala aus der Oracle-Quelle. Weicht
        '   diese von der (manuell gepflegten) Faktentabelle ab - z.B. Oracle
        '   NUMBER(38,0) vs. Faktentabelle decimal(17,10) - dann ist die
        '   _in_-Staging-Spalte != SWITCH-Ziel und SCR19 (Partitionstausch)
        '   bricht ab ("different data type / precision"). Die aktive Variante
        '   oben konvertiert stattdessen exakt auf die Faktentabellen-Typen.
        '
        ' WANN WIEDER VERWENDBAR:
        '   Sobald die Oracle-Seite bereinigt ist und Oracle- und SQL-Server-
        '   Datentypen (inkl. Praezision/Skala/Laenge) IDENTISCH sind, kann
        '   wieder columns_dbo verwendet werden (schlanker). Dann diesen Block
        '   aktivieren und die obige sys.columns-Variante auskommentieren.
        ' =====================================================================

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
            Throw New Exception("Keine Spalten aus der Faktentabelle dbo." & v.Faktentabelle.ToLower() &
                                " gelesen (sys.columns leer/Tabelle fehlt) - SELECT-Liste konnte nicht erzeugt werden.")
        End If

        ' EINZELFALL-PATCH nur fuer tf_lst_tkug_pers_verbl: Oracle hat die Spalte
        ' unter FTB_EZP_ID statt FTV_EZP_ID (Schreibfehler in der Quelle, nicht
        ' kurzfristig auf Oracle-Seite behebbar). Betrifft ausschliesslich diese
        ' eine Tabelle/Spalte - alle anderen Tabellen bleiben von der obigen,
        ' unveraenderten sys.columns-Logik unberuehrt.
        If extTable.ToLower() = "tf_lst_tkug_pers_verbl" Then
            selectList = selectList.Replace("[ftv_ezp_id])", "[FTB_EZP_ID])")
        End If

        ' DROP vorab im selben Batch: falls ein vorheriger Versuch die Tabelle
        ' schon angelegt hat und erst danach fehlgeschlagen ist (z.B. Verbindungsfehler
        ' nach erfolgreichem SELECT INTO), macht das jeden Retry selbst-bereinigend -
        ' sonst schlaegt jeder weitere Versuch mit "already an object" fehl.
        Dim sql As String
        Dim ladeart As String
        If v.StageCciInsert Then
            ' CCI-INSERT: _LOADING leer als Strukturkopie der Faktentabelle anlegen,
            ' CCI auf die leere Tabelle, dann einfuegen. Der CCI bekommt gleich den
            ' Namen der spaeteren _in_-Tabelle - SCR16 erkennt ihn und ueberspringt.
            Dim inTable As String = loadingTable.Substring(0, loadingTable.Length - "_LOADING".Length)
            Dim spaltenListe As String = Convert.ToString(SqlSkalar(connStr,
                "SELECT STRING_AGG(CAST(QUOTENAME(c.name) AS nvarchar(max)), ',') WITHIN GROUP (ORDER BY c.column_id) " &
                "FROM sys.columns c WHERE c.object_id = OBJECT_ID('dbo.[" & v.Faktentabelle.ToLower().Replace("'", "''") & "]') AND c.is_computed = 0;",
                "Spaltenliste " & loadingTable))
            sql =
                "IF OBJECT_ID('dbo.[" & loadingTable & "]','U') IS NOT NULL DROP TABLE dbo.[" & loadingTable & "];" & vbCrLf &
                "SELECT TOP 0 * INTO dbo.[" & loadingTable & "] FROM dbo.[" & v.Faktentabelle.ToLower() & "];" & vbCrLf &
                "CREATE CLUSTERED COLUMNSTORE INDEX [CCI_" & inTable & "] ON dbo.[" & loadingTable & "];" & vbCrLf &
                "INSERT INTO dbo.[" & loadingTable & "] WITH (TABLOCK) (" & spaltenListe & ")" & vbCrLf &
                "SELECT" & vbCrLf &
                selectList & vbCrLf &
                "FROM ext.[" & extTable & "]" & vbCrLf &
                "WHERE [" & v.PartitionColumn & "] = " & partitionValue & ";"
            ladeart = "CCI-INSERT"
        Else
            sql =
                "IF OBJECT_ID('dbo.[" & loadingTable & "]','U') IS NOT NULL DROP TABLE dbo.[" & loadingTable & "];" & vbCrLf &
                "SELECT" & vbCrLf &
                selectList & vbCrLf &
                "INTO dbo.[" & loadingTable & "]" & vbCrLf &
                "FROM ext.[" & extTable & "]" & vbCrLf &
                "WHERE [" & v.PartitionColumn & "] = " & partitionValue & ";"
            ladeart = "SELECT INTO"
        End If

        SyncLock _logSperre
            Log("  " & ladeart & " [" & loadingTable & "] WHERE " & v.PartitionColumn & " = " & partitionValue & " (wird zu _in_ umbenannt)")
        End SyncLock

        Dim betroffen As Integer = SqlAusfuehren(connStr, sql, ladeart & " " & loadingTable)
        If betroffen > 0 Then Return betroffen

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
            "       pf.Wert AS Faktentabelle, pp.Wert AS PartitionColumn," &
            "       pi.Wert AS IndexTyp, ps.Wert AS StageCciInsert" &
            " FROM   dbo.ETL_Fakt_Arbeitsliste a" &
            " JOIN   " & _parameterDB & ".dbo." & _parametertab & " pf ON pf.Verfahren=dbo.fn_ParamVerfahren(a.Verfahren) AND pf.Parameter='Faktentabelle'" &
            " JOIN   " & _parameterDB & ".dbo." & _parametertab & " pp ON pp.Verfahren=dbo.fn_ParamVerfahren(a.Verfahren) AND pp.Parameter='Faktenpartitionsspalte'" &
            " LEFT JOIN " & _parameterDB & ".dbo." & _parametertab & " pi ON pi.Verfahren=dbo.fn_ParamVerfahren(a.Verfahren) AND pi.Parameter='FaktenClusteredIndex'" &
            " LEFT JOIN " & _parameterDB & ".dbo." & _parametertab & " ps ON ps.Verfahren=dbo.fn_ParamVerfahren(a.Verfahren) AND ps.Parameter='FaktenStageCCIInsert'" &
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
                                Dim indexTyp As String = If(rdr.IsDBNull(6), "", rdr(6).ToString().Trim().ToUpper())
                                Dim cciSchalter As String = If(rdr.IsDBNull(7), "", rdr(7).ToString().Trim().ToUpper())
                                liste.Add(New VerfahrenInfo With {
                                    .ID = rdr.GetInt32(0),
                                    .Verfahren = rdr(1).ToString().Trim(),
                                    .Themengebiet = rdr(2).ToString().Trim(),
                                    .LetzterSchritt = If(rdr.IsDBNull(3), "", rdr(3).ToString().Trim()),
                                    .Faktentabelle = rdr(4).ToString().Trim(),
                                    .PartitionColumn = If(rawPart.Contains("|"),
                                        rawPart.Substring(0, rawPart.IndexOf("|")), rawPart),
                                    .StageCciInsert = (indexTyp = "CCI" AndAlso cciSchalter = "TRUE")
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
            "UPDATE dbo.ETL_Fakt_Arbeitsliste SET Status='" & status &
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
                    "UPDATE dbo.ETL_Fakt_Arbeitsliste SET Status='FEHLER',Fehlermeldung=@m,AktualisiertAm=GETDATE() WHERE ID=@id",
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
                Dim istFilegroupVoll As Boolean =
                    (TypeOf ex Is SqlException) AndAlso
                    (CType(ex, SqlException).Number = SQL_FEHLER_FILEGROUP_VOLL OrElse
                     CType(ex, SqlException).Number = SQL_FEHLER_SEITE_NICHT_ZUWEISBAR)
                SyncLock _logSperre
                    Log(String.Format("WARNUNG [{0}] Versuch {1}/{2}: {3}", beschreibung, versuch, MAX_VERSUCHE, ex.Message))
                If versuch = 1 Then Log("SQL Statement [" & beschreibung & "]: " & sql)
                End SyncLock
                If IstNichtWiederholbar(ex) Then Exit While
                If istFilegroupVoll Then
                    DateiWachstumErzwingen(connStr)
                ElseIf versuch < MAX_VERSUCHE Then
                    Thread.Sleep(WARTE_SEK * 1000)
                End If
                If _globalerAbbruch Then Exit While
            End Try
        End While
        Throw New Exception(String.Format("[{0}] fehlgeschlagen: {1}", beschreibung,
            If(letzterFehler IsNot Nothing, letzterFehler.Message, "Unbekannt")), letzterFehler)
    End Function

    Private Function SqlNummer(ex As Exception) As Integer
        Dim e As Exception = ex
        While e IsNot Nothing
            If TypeOf e Is SqlException Then Return CType(e, SqlException).Number
            e = e.InnerException
        End While
        Return 0
    End Function

    Private Function IstNichtWiederholbar(ex As Exception) As Boolean
        Select Case SqlNummer(ex)
            Case 102, 156, 207, 208, 245, 2714
                Return True
        End Select
        Return False
    End Function

    Private Function IstVerbindungsfehler(ex As Exception) As Boolean
        Select Case SqlNummer(ex)
            Case 53, 64, 233, 4060, 10053, 10054, 10060, 18456
                Return True
        End Select
        Return ex.Message.Contains("Cannot open database") OrElse ex.Message.Contains("The login failed")
    End Function

    Private Sub DateienVorabVergroessern(connStr As String)
        Try
            Dim sql As String =
"DECLARE @bedarfMB bigint = " & (CLng(_maxparallel) * VORAB_MB_JE_STREAM).ToString() & ";
DECLARE @sql nvarchar(max) = N'';
WITH f AS (
    SELECT f.name,
           CAST(f.size AS bigint) / 128 AS groesseMB,
           CAST(FILEPROPERTY(f.name, 'SpaceUsed') AS bigint) / 128 AS belegtMB,
           v.available_bytes / 1048576 AS platzMB,
           COUNT(*) OVER () AS anz
    FROM sys.database_files f
    JOIN sys.filegroups fg ON fg.data_space_id = f.data_space_id AND fg.is_default = 1
    CROSS APPLY sys.dm_os_volume_stats(DB_ID(), f.file_id) v
    WHERE f.type_desc = 'ROWS'
), g AS (
    SELECT name, groesseMB, platzMB,
           @bedarfMB / anz - (groesseMB - belegtMB) AS zuwachs
    FROM f
)
SELECT @sql = @sql + N'ALTER DATABASE ' + QUOTENAME(DB_NAME()) + N' MODIFY FILE (NAME = ' + QUOTENAME(name) +
              N', SIZE = ' + CAST(groesseMB + CASE WHEN zuwachs > platzMB * 0.9 THEN CAST(platzMB * 0.9 AS bigint) ELSE zuwachs END AS varchar(20)) + N'MB); '
FROM g
WHERE zuwachs > 0 AND platzMB * 0.9 >= 1;
IF @sql <> N'' EXEC sp_executesql @sql;"
            SqlAusfuehren(connStr, sql, "Dateien vergroessern")
            Log("Dateien vorab vergroessert.")
        Catch ex As Exception
            Log("WARNUNG [Dateien vergroessern]: " & ex.Message)
        End Try
    End Sub

    ' -----------------------------------------------------------------------
    ' DateiWachstumErzwingen - Reagiert auf SQL-Fehler 1105/1101 ("Dateigruppe
    ' voll"): vergroessert die Datendateien der Standard-Dateigruppe um einen
    ' festen Schritt. Per SyncLock serialisiert, damit bei mehreren
    ' gleichzeitig fehlschlagenden Partitionen nur EIN Thread vergroessert -
    ' die anderen warten kurz und versuchen es danach erneut.
    ' -----------------------------------------------------------------------
    Private Sub DateiWachstumErzwingen(connStr As String)
        SyncLock _wachstumsSperre
            Try
                Dim sql As String =
"DECLARE @sql nvarchar(max) = N'';
SELECT @sql = @sql + N'ALTER DATABASE ' + QUOTENAME(DB_NAME()) + N' MODIFY FILE (NAME = ' + QUOTENAME(f.name) +
              N', SIZE = ' + CAST(CEILING(f.size / 128.0) + " & DATEIWACHSTUM_SCHRITT_MB & " AS varchar(20)) + N'MB); '
FROM sys.database_files f
JOIN sys.filegroups fg ON fg.data_space_id = f.data_space_id AND fg.is_default = 1
WHERE f.type_desc = 'ROWS';
EXEC sp_executesql @sql;"
                Using conn As New SqlConnection(connStr)
                    conn.Open()
                    Using cmd As New SqlCommand(sql, conn)
                        cmd.CommandTimeout = 0
                        cmd.ExecuteNonQuery()
                    End Using
                End Using
                SyncLock _logSperre
                    Log("  Dateiwachstum ausgefuehrt (+" & DATEIWACHSTUM_SCHRITT_MB.ToString() & " MB je Datendatei) nach Fehler 1105/1101")
                End SyncLock
            Catch ex As Exception
                SyncLock _logSperre
                    Log("WARNUNG [Dateiwachstum]: Datendatei konnte nicht vergroessert werden - Laufwerk vermutlich voll: " & ex.Message)
                End SyncLock
                If ex.Message.Contains("operating system error 112") OrElse ex.Message.Contains("not enough space on the disk") Then
                    _abbruchGrund = "Laufwerk voll"
                    _globalerAbbruch = True
                End If
            End Try
        End SyncLock
    End Sub

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
                If versuch < MAX_VERSUCHE AndAlso Not IstNichtWiederholbar(ex) Then Thread.Sleep(WARTE_SEK * 1000) Else Throw
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
        Public Property StageCciInsert As Boolean
    End Class

    ' -----------------------------------------------------------------------
    ' ScriptResults - SSIS-Task-Ergebniscodes.
    ' -----------------------------------------------------------------------
    Public Enum ScriptResults
        Success = DTSExecResult.Success
        Failure = DTSExecResult.Failure
    End Enum

End Class
