Option Explicit On
Option Strict On

Imports System
Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports System.Linq
Imports Microsoft.SqlServer.Dts.Runtime

' =============================================================================
' PAKET  : Fakten Laden
' SKRIPT : SCR09_Partitionsgrenzen
' 
' ZWECK  : HYBRID MODE - Unterstuetzt BEIDE Modi:
'          
'          MODE 1 - MANUAL (partition_wert in CSV vorhanden):
'            - Verwendet nur die vom Benutzer angegebenen Werte
'            - Validiert gegen Oracle
'            - AKTUALISIERUNG fuer existierende Werte
'            - NEU fuer fehlende Werte
'          
'          MODE 2 - AUTOMATIC (partition_wert ist NULL/leer):
'            - Liest ALLE distinct Werte aus Oracle
'            - FULL LOAD wenn MSSQL leer
'            - APPEND: Laedt ALLE Oracle-Werte die NICHT in MSSQL sind
'              (vor, zwischen und nach - keine Luecken!)
'
' FIXES  : - CCI/CI Index detection for temp tables in SWITCH operations
'          - APPEND loads ALL missing values (not just > MAX)
'          - Compact logging (no individual partition lines)
'
' Status: PARTITIONSGRENZEN â PARTITIONSGRENZEN_ERSTELLT
' =============================================================================
<Microsoft.SqlServer.Dts.Tasks.ScriptTask.SSISScriptTaskEntryPointAttribute()>
<CLSCompliant(False)>
Partial Public Class ScriptMain
    Inherits Microsoft.SqlServer.Dts.Tasks.ScriptTask.VSTARTScriptObjectModelBase

    Private Const SKRIPT_NAME As String = "SCR11_Partitionsgrenzen"
    Private Const CONN_NAME As String = "Verbindung"
    Private Const MAX_VERSUCHE As Integer = 10
    Private Const WARTE_SEK As Integer = 30

    Private _runID As Integer = 0
    Private _parameterDB As String = String.Empty
    Private _parametertab As String = String.Empty
    Private _stlTabelle As String = String.Empty

    Public Sub Main()

        Log("ââââââââââââââââââââââââââââââââââââââââââââââââââââââââ")
        Log("SCR09_Partitionsgrenzen â Start (v6: FINAL - All Fixes)")
        Log("Zeitpunkt: " & DateTime.Now.ToString("dd.MM.yyyy HH:mm:ss"))
        Log("ââââââââââââââââââââââââââââââââââââââââââââââââââââââââ")

        Try
            _runID = Convert.ToInt32(Dts.Variables("BA::RunID").Value)
            _parameterDB = Dts.Variables("BA::ParameterDB").Value.ToString().Trim()
            _parametertab = Dts.Variables("BA::Parametertabelle").Value.ToString().Trim()
            _stlTabelle = Dts.Variables("BA::SteuerlistenTabelle").Value.ToString().Trim()

            Dim connStr As String = HoleVerbindungszeichenfolge()
            Dim verfahren As List(Of VerfahrenInfo) = VerfahrenLaden(connStr)
            Log("Verfahren: " & verfahren.Count.ToString())

            Dim gesamtErgebnis As New Dictionary(Of String, List(Of PartitionsEintrag))()

            Dim cntOK As Integer = 0
            Dim cntFehler As Integer = 0

            For Each v As VerfahrenInfo In verfahren
                Log("ââââââââââââââââââââââââââââââââââââââââââââââââââââââââ")
                Log("Verfahren: " & v.Verfahren & " | Tabelle: " & v.Faktentabelle & " | Spalte: " & v.PartitionsSpalte)

                If v.LetzterSchritt = "PARTITIONSGRENZEN_ERSTELLT" Then
                    Log("  â bereits abgeschlossen â uebersprungen â")
                    Continue For
                End If

                Try
                    StatusSetzen(connStr, v.ID, "PARTITIONSGRENZEN")

                    ' âââââââââââââââââââââââââââââââââââââââââââââââââââââââââ
                    ' SCHRITT 1: Versuche partition_wert aus CSV zu lesen
                    ' âââââââââââââââââââââââââââââââââââââââââââââââââââââââââ
                    Dim benutzerWerte As List(Of Integer) = PartitionWerteLaden(connStr, v.Verfahren)

                    Dim zuVerarbeiten As New List(Of PartitionsEintrag)()
                    Dim modus As String = String.Empty
                    Dim oracleAlleWerte As List(Of Integer) = Nothing
                    Dim mssqlWerte As List(Of Integer) = Nothing

                    If benutzerWerte.Count > 0 Then
                        ' âââââââââââââââââââââââââââââââââââââââââââââââââââââ
                        ' MODE 1: MANUAL (partition_wert aus CSV)
                        ' âââââââââââââââââââââââââââââââââââââââââââââââââââââ
                        modus = "MANUAL"
                        Log("  ââââââââââââââââââââââââââââââââââââââââââââââââââ")
                        Log("  MODE: MANUAL (partition_wert aus CSV)")
                        Log("  ââââââââââââââââââââââââââââââââââââââââââââââââââ")
                        Log("  Benutzer partition_wert: " & benutzerWerte.Count.ToString() & " Werte")
                        Log("  MIN: " & benutzerWerte.Min().ToString() & " | MAX: " & benutzerWerte.Max().ToString())

                        ' Alle Oracle-Werte laden
                        oracleAlleWerte = OracleAlleWerteLaden(connStr, v)
                        Log("  Oracle Werte gesamt: " & oracleAlleWerte.Count.ToString())

                        ' Validierung - Benutzer-Werte muessen in Oracle existieren
                        Dim nichtInOracle As List(Of Integer) = benutzerWerte.Where(Function(w) Not oracleAlleWerte.Contains(w)).ToList()
                        If nichtInOracle.Count > 0 Then
                            Dim fehlermeldung As String =
                                "FEHLER: " & nichtInOracle.Count.ToString() & " partition_wert nicht in Oracle: " &
                                String.Join(", ", nichtInOracle.Take(10).Select(Function(x) x.ToString()).ToArray()) &
                                If(nichtInOracle.Count > 10, " ...", "")
                            Log("  " & fehlermeldung)
                            FehlerSetzen(connStr, v.ID, fehlermeldung)
                            ProtokollSchreiben(connStr, v.Verfahren, "FEHLER_SCR09", fehlermeldung)
                            LogFehler(fehlermeldung)
                            Dts.TaskResult = ScriptResults.Failure
                            Return
                        End If
                        Log("  â Alle Benutzer-Werte in Oracle vorhanden â")

                        ' MSSQL Status pruefen
                        mssqlWerte = MssqlWerteLaden(connStr, v)
                        Log("  MSSQL Werte gesamt: " & mssqlWerte.Count.ToString())

                        ' Klassifizierung - AKTUALISIERUNG vs NEU
                        Dim cntAktualisierung As Integer = 0
                        Dim cntNeu As Integer = 0

                        For Each bw As Integer In benutzerWerte
                            If mssqlWerte.Contains(bw) Then
                                zuVerarbeiten.Add(New PartitionsEintrag With {.Wert = bw, .Modus = "AKTUALISIERUNG"})
                                cntAktualisierung += 1
                            Else
                                zuVerarbeiten.Add(New PartitionsEintrag With {.Wert = bw, .Modus = "NEU"})
                                cntNeu += 1
                            End If
                        Next

                        Log("  Klassifizierung: AKTUALISIERUNG=" & cntAktualisierung.ToString() & " | NEU=" & cntNeu.ToString())

                    Else
                        ' âââââââââââââââââââââââââââââââââââââââââââââââââââââ
                        ' MODE 2: AUTOMATIC (kein partition_wert in CSV)
                        ' âââââââââââââââââââââââââââââââââââââââââââââââââââââ
                        modus = "AUTOMATIC"
                        Log("  ââââââââââââââââââââââââââââââââââââââââââââââââââ")
                        Log("  MODE: AUTOMATIC (partition_wert aus Oracle)")
                        Log("  ââââââââââââââââââââââââââââââââââââââââââââââââââ")

                        ' Alle Oracle-Werte laden
                        oracleAlleWerte = OracleAlleWerteLaden(connStr, v)
                        Log("  Oracle Werte gesamt: " & oracleAlleWerte.Count.ToString())

                        If oracleAlleWerte.Count = 0 Then
                            Log("  WARNUNG: Keine Daten in Oracle â uebersprungen")
                            ProtokollSchreiben(connStr, v.Verfahren, "WARNUNG_SCR09", "Keine Daten in Oracle")
                            StatusSetzen(connStr, v.ID, "PARTITIONSGRENZEN_ERSTELLT")
                            cntOK += 1
                            Continue For
                        End If

                        Log("  Oracle MIN: " & oracleAlleWerte.Min().ToString() & " | MAX: " & oracleAlleWerte.Max().ToString())

                        ' MSSQL Status pruefen
                        mssqlWerte = MssqlWerteLaden(connStr, v)
                        Log("  MSSQL Werte gesamt: " & mssqlWerte.Count.ToString())

                        If mssqlWerte.Count = 0 Then
                            ' âââ FULL LOAD: MSSQL ist leer âââ
                            Log("  ââââââââââââââââââââââââââââââââââââââââââââââ")
                            Log("  â Entscheidung: FULL LOAD (MSSQL leer)")
                            Log("  â Lade alle " & oracleAlleWerte.Count.ToString() & " Oracle-Werte")
                            Log("  ââââââââââââââââââââââââââââââââââââââââââââââ")

                            For Each ow As Integer In oracleAlleWerte
                                zuVerarbeiten.Add(New PartitionsEintrag With {.Wert = ow, .Modus = "NEU"})
                            Next

                        Else
                            ' âââ APPEND: Lade ALLE fehlenden Oracle-Werte âââ
                            Log("  MSSQL MIN: " & mssqlWerte.Min().ToString() & " | MAX: " & mssqlWerte.Max().ToString())
                            Log("  ââââââââââââââââââââââââââââââââââââââââââââââ")
                            Log("  â Entscheidung: APPEND (alle fehlenden Werte)")
                            Log("  ââââââââââââââââââââââââââââââââââââââââââââââ")

                            ' âââââââââââââââââââââââââââââââââââââââââââââââââââââââââ
                            ' FIX: Load ALL Oracle values NOT in MSSQL
                            ' This catches:
                            ' - Historical values before MSSQL MIN
                            ' - Gap values between MSSQL MIN and MAX
                            ' - Future values after MSSQL MAX
                            ' âââââââââââââââââââââââââââââââââââââââââââââââââââââââââ
                            Dim neueWerte As List(Of Integer) = oracleAlleWerte.Where(Function(w) Not mssqlWerte.Contains(w)).ToList()

                            Log("  Oracle gesamt: " & oracleAlleWerte.Count.ToString())
                            Log("  MSSQL vorhanden: " & mssqlWerte.Count.ToString())
                            Log("  Fehlende Werte: " & neueWerte.Count.ToString())

                            If neueWerte.Count = 0 Then
                                Log("  â Keine fehlenden Werte â uebersprungen")
                                StatusSetzen(connStr, v.ID, "PARTITIONSGRENZEN_ERSTELLT")
                                cntOK += 1
                                Continue For
                            End If

                            ' Show where gaps are
                            Log("  Fehlend MIN: " & neueWerte.Min().ToString() & " | MAX: " & neueWerte.Max().ToString())

                            ' Count gaps in different regions
                            Dim vorMin As Integer = neueWerte.Where(Function(w) w < mssqlWerte.Min()).Count()
                            Dim nachMax As Integer = neueWerte.Where(Function(w) w > mssqlWerte.Max()).Count()
                            Dim dazwischen As Integer = neueWerte.Count - vorMin - nachMax

                            Log("  Fehlend VOR MSSQL MIN: " & vorMin.ToString())
                            Log("  Fehlend ZWISCHEN MIN-MAX: " & dazwischen.ToString())
                            Log("  Fehlend NACH MSSQL MAX: " & nachMax.ToString())

                            For Each nw As Integer In neueWerte
                                zuVerarbeiten.Add(New PartitionsEintrag With {.Wert = nw, .Modus = "NEU"})
                            Next
                        End If
                    End If

                    ' âââââââââââââââââââââââââââââââââââââââââââââââââââââââââ
                    ' ZUSAMMENFASSUNG (COMPACT)
                    ' âââââââââââââââââââââââââââââââââââââââââââââââââââââââââ
                    zuVerarbeiten = zuVerarbeiten.OrderBy(Function(z) z.Wert).ToList()

                    Log("  ââââââââââââââââââââââââââââââââââââââââââââââââââ")
                    Log("  ZUSAMMENFASSUNG: " & v.Verfahren)
                    Log("  ââââââââââââââââââââââââââââââââââââââââââââââââââ")
                    Log("  Modus: " & modus)
                    Log("  Gesamt zu verarbeiten: " & zuVerarbeiten.Count.ToString())

                    If modus = "MANUAL" Then
                        Log("  AKTUALISIERUNG: " & zuVerarbeiten.Where(Function(z) z.Modus = "AKTUALISIERUNG").Count().ToString())
                    End If
                    Log("  NEU: " & zuVerarbeiten.Where(Function(z) z.Modus = "NEU").Count().ToString())

                    ' Show value range in ONE line
                    If zuVerarbeiten.Count > 0 Then
                        Dim minVal As Integer = zuVerarbeiten.Min(Function(z) z.Wert)
                        Dim maxVal As Integer = zuVerarbeiten.Max(Function(z) z.Wert)
                        Log("  Werte-Bereich: " & minVal.ToString() & " bis " & maxVal.ToString() &
                            " (" & zuVerarbeiten.Count.ToString() & " Partitionen)")
                    End If
                    Log("  ââââââââââââââââââââââââââââââââââââââââââââââââââ")

                    If zuVerarbeiten.Count = 0 Then
                        Log("  â Keine Partitionswerte zu verarbeiten â uebersprungen")
                        StatusSetzen(connStr, v.ID, "PARTITIONSGRENZEN_ERSTELLT")
                        cntOK += 1
                        Continue For
                    End If

                    ' âââââââââââââââââââââââââââââââââââââââââââââââââââââââââ
                    ' PARTITIONSGRENZEN ERSTELLEN (nur fuer NEU Werte)
                    ' âââââââââââââââââââââââââââââââââââââââââââââââââââââââââ
                    Dim dateigruppe As String = Convert.ToString(SqlSkalar(connStr,
                        "SELECT name FROM sys.filegroups WHERE is_default=1", "Dateigruppe"))
                    Dim pf As String = "PF_" & v.PartitionsSpalte & "_" & v.Faktentabelle
                    Dim ps As String = "PS_" & v.PartitionsSpalte & "_" & v.Faktentabelle

                    Dim cntSplit As Integer = 0
                    Dim cntSkip As Integer = 0

                    For Each pe As PartitionsEintrag In zuVerarbeiten
                        If pe.Modus = "NEU" Then
                            PartitionSplitDurchfuehren(connStr, v, pf, ps, dateigruppe, pe.Wert)
                            cntSplit += 1
                        Else
                            cntSkip += 1
                        End If
                    Next

                    Log("  â SPLIT ausgefuehrt: " & cntSplit.ToString() & " | Uebersprungen: " & cntSkip.ToString())

                    ' âââââââââââââââââââââââââââââââââââââââââââââââââââââââââ
                    ' ERGEBNIS SPEICHERN
                    ' âââââââââââââââââââââââââââââââââââââââââââââââââââââââââ
                    gesamtErgebnis(v.Verfahren) = zuVerarbeiten

                    StatusSetzen(connStr, v.ID, "PARTITIONSGRENZEN_ERSTELLT")

                    ' Compact protocol entry
                    Dim protMsg As String = "Partitionsgrenzen erstellt. Modus=" & modus &
                        " | Gesamt=" & zuVerarbeiten.Count.ToString() &
                        If(modus = "MANUAL", " | AKTUALISIERUNG=" & zuVerarbeiten.Where(Function(z) z.Modus = "AKTUALISIERUNG").Count().ToString(), "") &
                        " | NEU=" & zuVerarbeiten.Where(Function(z) z.Modus = "NEU").Count().ToString() &
                        " | Bereich=" & zuVerarbeiten.Min(Function(z) z.Wert).ToString() & "-" & zuVerarbeiten.Max(Function(z) z.Wert).ToString()

                    ProtokollSchreiben(connStr, v.Verfahren, "SCHRITT_4", protMsg)
                    cntOK += 1
                    Log("  â Schritt 4 abgeschlossen â")

                Catch ex As Exception
                    cntFehler += 1
                    FehlerSetzen(connStr, v.ID, ex.Message)
                    ProtokollSchreiben(connStr, v.Verfahren, "FEHLER_SCR09", ex.Message)
                    LogFehler("FEHLER '" & v.Verfahren & "': " & ex.Message)
                End Try
            Next

            ' âââââââââââââââââââââââââââââââââââââââââââââââââââââââââ
            ' BA::objPartitionValues SETZEN
            ' âââââââââââââââââââââââââââââââââââââââââââââââââââââââââ
            Dim gesamtAnzahl As Integer = 0
            For Each kvp As KeyValuePair(Of String, List(Of PartitionsEintrag)) In gesamtErgebnis
                gesamtAnzahl += kvp.Value.Count
            Next

            If gesamtAnzahl > 0 Then
                Dim partArray(gesamtAnzahl - 1, 2) As String
                Dim idx As Integer = 0
                For Each kvp As KeyValuePair(Of String, List(Of PartitionsEintrag)) In gesamtErgebnis
                    For Each pe As PartitionsEintrag In kvp.Value
                        partArray(idx, 0) = kvp.Key
                        partArray(idx, 1) = pe.Wert.ToString()
                        partArray(idx, 2) = pe.Modus
                        idx += 1
                    Next
                Next
                Dts.Variables("BA::objPartitionValues").Value = partArray

                Log("ââââââââââââââââââââââââââââââââââââââââââââââââââââââââ")
                Log("BA::objPartitionValues gesetzt: " & gesamtAnzahl.ToString() & " Eintraege")

                ' Show summary by Verfahren (compact)
                For Each kvp As KeyValuePair(Of String, List(Of PartitionsEintrag)) In gesamtErgebnis
                    Dim minV As Integer = kvp.Value.Min(Function(p) p.Wert)
                    Dim maxV As Integer = kvp.Value.Max(Function(p) p.Wert)
                    Log("  " & kvp.Key & ": " & kvp.Value.Count.ToString() & " Partitionen (" & minV.ToString() & "-" & maxV.ToString() & ")")
                Next

                Log("ââââââââââââââââââââââââââââââââââââââââââââââââââââââââ")
            Else
                Dts.Variables("BA::objPartitionValues").Value = Nothing
                Log("BA::objPartitionValues: leer (Nothing)")
            End If

            Log("ââââââââââââââââââââââââââââââââââââââââââââââââââââââââ")
            Log("Erfolgreich: " & cntOK.ToString() & " | Fehler: " & cntFehler.ToString())
            Log("ââââââââââââââââââââââââââââââââââââââââââââââââââââââââ")
            Dts.TaskResult = If(cntFehler > 0, ScriptResults.Failure, ScriptResults.Success)

        Catch ex As Exception
            LogFehler("Kritischer Fehler: " & ex.Message)
            Dts.TaskResult = ScriptResults.Failure
        End Try

    End Sub

    ' =========================================================================
    ' partition_wert aus Steuerlisten-Tabelle laden
    ' =========================================================================
    Private Function PartitionWerteLaden(connStr As String, verfahren As String) As List(Of Integer)
        Dim alleWerte As New HashSet(Of Integer)()
        Dim sql As String =
            "SELECT partition_wert FROM dbo." & _stlTabelle &
            " WHERE LOWER(LTRIM(RTRIM(tabelle))) = @verf" &
            " AND partition_wert IS NOT NULL AND LTRIM(RTRIM(partition_wert)) <> ''"

        Dim versuch As Integer = 0
        While versuch < MAX_VERSUCHE
            versuch += 1
            Try
                Using conn As New SqlConnection(connStr)
                    conn.Open()
                    Using cmd As New SqlCommand(sql, conn)
                        cmd.CommandTimeout = 0
                        cmd.Parameters.AddWithValue("@verf", verfahren.ToLower().Trim())
                        Using rdr As SqlDataReader = cmd.ExecuteReader()
                            While rdr.Read()
                                Dim rohWert As String = rdr(0).ToString().Trim()
                                Dim teile() As String = rohWert.Split(","c)
                                For Each teil As String In teile
                                    Dim sauber As String = teil.Trim()
                                    Dim intWert As Integer
                                    If Integer.TryParse(sauber, intWert) Then
                                        alleWerte.Add(intWert)
                                    End If
                                Next
                            End While
                        End Using
                    End Using
                End Using
                Return alleWerte.OrderBy(Function(w) w).ToList()
            Catch ex As Exception
                If versuch < MAX_VERSUCHE Then System.Threading.Thread.Sleep(WARTE_SEK * 1000) Else Throw
            End Try
        End While
        Return alleWerte.OrderBy(Function(w) w).ToList()
    End Function

    ' =========================================================================
    ' ALLE distinct Partitionswerte aus Oracle laden
    ' =========================================================================
    Private Function OracleAlleWerteLaden(connStr As String, v As VerfahrenInfo) As List(Of Integer)
        Dim liste As New List(Of Integer)()
        Dim sql As String = "SELECT DISTINCT [" & v.PartitionsSpalte & "] FROM ext.[" & v.Faktentabelle.ToLower() &
            "] WHERE [" & v.PartitionsSpalte & "] IS NOT NULL ORDER BY [" & v.PartitionsSpalte & "]"

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
                                If Not rdr.IsDBNull(0) Then liste.Add(rdr.GetInt32(0))
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

    ' =========================================================================
    ' ALLE distinct Partitionswerte aus MSSQL laden
    ' =========================================================================
    Private Function MssqlWerteLaden(connStr As String, v As VerfahrenInfo) As List(Of Integer)
        Dim liste As New List(Of Integer)()
        Dim sql As String = "SELECT DISTINCT [" & v.PartitionsSpalte & "] FROM dbo.[" & v.Faktentabelle &
            "] WHERE [" & v.PartitionsSpalte & "] IS NOT NULL ORDER BY [" & v.PartitionsSpalte & "]"

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
                                If Not rdr.IsDBNull(0) Then liste.Add(rdr.GetInt32(0))
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

    ' =========================================================================
    ' Partition SPLIT durchfuehren (MIT CCI/CI FIX)
    ' =========================================================================
    Private Sub PartitionSplitDurchfuehren(connStr As String, v As VerfahrenInfo,
                                            pf As String, ps As String,
                                            dateigruppe As String, partWert As Integer)

        Dim sqlInfo As String =
"SELECT MAX(CASE WHEN CONVERT(int,r.value)=@pv THEN 1 ELSE 0 END) AS treffer,
        MAX(CASE WHEN p.rows=0 THEN 1 ELSE 0 END) AS leer,
        MAX(ISNULL(CONVERT(int,r.value),2147483647)) AS partname,
        MAX(p.partition_number) AS partid
 FROM sys.indexes i
 JOIN sys.tables t ON i.object_id=t.object_id
 JOIN sys.partitions p ON i.object_id=p.object_id AND i.index_id=p.index_id AND p.index_id<2
 JOIN sys.data_spaces d ON i.data_space_id=d.data_space_id
 LEFT JOIN sys.partition_schemes s ON d.name=s.name
 LEFT JOIN sys.partition_functions f ON s.function_id=f.function_id
 LEFT JOIN sys.partition_range_values r ON r.function_id=f.function_id AND r.boundary_id+f.boundary_value_on_right=p.partition_number
 LEFT JOIN sys.partition_range_values vv ON vv.function_id=f.function_id AND CONVERT(int,vv.value)<ISNULL(CONVERT(int,r.value),2147483647)
 WHERE t.schema_id=SCHEMA_ID('dbo') AND t.name=@ft AND t.type='U'
 GROUP BY r.value,p.partition_number,p.rows
 HAVING @pv>ISNULL(MAX(CONVERT(int,vv.value)),-2147483648) AND @pv<=ISNULL(CONVERT(int,r.value),2147483647)"

        Dim treffer As Integer = 0
        Dim leer As Integer = 0
        Dim partName As Integer = 0
        Dim partId As Integer = 0

        Dim versuch As Integer = 0
        While versuch < MAX_VERSUCHE
            versuch += 1
            Try
                Using conn As New SqlConnection(connStr)
                    conn.Open()
                    Using cmd As New SqlCommand(sqlInfo, conn)
                        cmd.CommandTimeout = 0
                        cmd.Parameters.AddWithValue("@pv", partWert)
                        cmd.Parameters.AddWithValue("@ft", v.Faktentabelle)
                        Using rdr As SqlDataReader = cmd.ExecuteReader()
                            If rdr.Read() Then
                                treffer = If(rdr.IsDBNull(0), 0, rdr.GetInt32(0))
                                leer = If(rdr.IsDBNull(1), 0, rdr.GetInt32(1))
                                partName = If(rdr.IsDBNull(2), 0, rdr.GetInt32(2))
                                partId = If(rdr.IsDBNull(3), 0, rdr.GetInt32(3))
                            End If
                        End Using
                    End Using
                End Using
                Exit While
            Catch ex As Exception
                If versuch < MAX_VERSUCHE Then System.Threading.Thread.Sleep(WARTE_SEK * 1000) Else Throw
            End Try
        End While

        If treffer = 1 Then
            Return  ' Boundary already exists
        End If

        If leer = 1 Then
            ' Direct SPLIT - partition is empty
            SqlAusfuehren(connStr,
                "ALTER PARTITION SCHEME [" & ps & "] NEXT USED [" & dateigruppe & "];" &
                "ALTER PARTITION FUNCTION [" & pf & "]() SPLIT RANGE(" & partWert & ");",
                "SPLIT direkt")
        Else
            ' âââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââ
            ' SWITCH OUT / SPLIT / SWITCH IN (WITH CCI/CI FIX)
            ' âââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââ
            Dim tmpTabelle As String = v.Faktentabelle & "_tmp_" & partName.ToString()
            SqlAusfuehren(connStr, "IF OBJECT_ID('dbo.[" & tmpTabelle & "]','U') IS NOT NULL DROP TABLE dbo.[" & tmpTabelle & "];", "Tmp loeschen")

            Dim spaltenDef As String = HoleSpaltendefinition(connStr, v.Faktentabelle)

            ' Get compression
            Dim komprimierung As String = Convert.ToString(SqlSkalar(connStr,
                "SELECT TOP 1 p.data_compression_desc FROM sys.partitions p JOIN sys.indexes i ON p.object_id=i.object_id AND p.index_id=i.index_id JOIN sys.tables t ON t.object_id=p.object_id WHERE t.name='" & v.Faktentabelle & "' AND p.partition_number=1",
                "Komprimierung"))
            Dim kompStr As String = If(komprimierung = "PAGE" OrElse komprimierung = "ROW", " WITH (DATA_COMPRESSION=" & komprimierung & ")", "")

            ' âââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââ
            ' FIX: Detect source table index type
            ' âââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââ
            Dim sqlIndexInfo As String =
                "SELECT i.type_desc, i.name " &
                "FROM sys.indexes i " &
                "JOIN sys.tables t ON i.object_id = t.object_id " &
                "WHERE t.name = '" & v.Faktentabelle & "' AND i.index_id = 1"

            Dim indexType As String = ""
            Dim indexName As String = ""

            versuch = 0
            While versuch < MAX_VERSUCHE
                versuch += 1
                Try
                    Using conn As New SqlConnection(connStr)
                        conn.Open()
                        Using cmd As New SqlCommand(sqlIndexInfo, conn)
                            cmd.CommandTimeout = 0
                            Using rdr As SqlDataReader = cmd.ExecuteReader()
                                If rdr.Read() Then
                                    indexType = If(rdr.IsDBNull(0), "", rdr.GetString(0))
                                    indexName = If(rdr.IsDBNull(1), "", rdr.GetString(1))
                                End If
                            End Using
                        End Using
                    End Using
                    Exit While
                Catch ex As Exception
                    If versuch < MAX_VERSUCHE Then System.Threading.Thread.Sleep(WARTE_SEK * 1000) Else Throw
                End Try
            End While

            ' Create temp table (HEAP first)
            SqlAusfuehren(connStr,
                "CREATE TABLE dbo.[" & tmpTabelle & "] (" & spaltenDef & ")" & kompStr & ";",
                "Tmp erstellen")

            ' âââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââ
            ' Create matching clustered index on temp table
            ' âââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââ
            If indexType = "CLUSTERED COLUMNSTORE" Then
                ' Source has Clustered Columnstore Index (CCI)
                SqlAusfuehren(connStr,
                    "CREATE CLUSTERED COLUMNSTORE INDEX [CCI_" & tmpTabelle & "] ON dbo.[" & tmpTabelle & "];",
                    "Tmp CCI erstellen")

            ElseIf indexType = "CLUSTERED" Then
                ' Source has regular Clustered Index (CI) - get key columns
                Dim sqlIndexCols As String =
                    "SELECT STUFF((SELECT ', ' + QUOTENAME(c.name) " &
                    "FROM sys.index_columns ic " &
                    "JOIN sys.columns c ON ic.object_id = c.object_id AND ic.column_id = c.column_id " &
                    "WHERE ic.object_id = OBJECT_ID('dbo." & v.Faktentabelle & "') AND ic.index_id = 1 " &
                    "ORDER BY ic.key_ordinal " &
                    "FOR XML PATH(''), TYPE).value('.','NVARCHAR(MAX)'), 1, 2, '')"

                Dim indexColumns As String = Convert.ToString(SqlSkalar(connStr, sqlIndexCols, "Index Columns"))

                SqlAusfuehren(connStr,
                    "CREATE CLUSTERED INDEX [CI_" & tmpTabelle & "] ON dbo.[" & tmpTabelle & "] (" & indexColumns & ")" & kompStr & ";",
                    "Tmp CI erstellen")
            End If
            ' If HEAP (no index_id=1), temp table stays as HEAP

            ' Now SWITCH OUT will work - indexes match!
            SqlAusfuehren(connStr,
                "ALTER TABLE dbo.[" & v.Faktentabelle & "] SWITCH PARTITION " & partId & " TO dbo.[" & tmpTabelle & "];",
                "SWITCH OUT")

            ' SPLIT
            SqlAusfuehren(connStr,
                "ALTER PARTITION SCHEME [" & ps & "] NEXT USED [" & dateigruppe & "];" &
                "ALTER PARTITION FUNCTION [" & pf & "]() SPLIT RANGE(" & partWert & ");",
                "SPLIT")

            ' Find new partition ID
            Dim neuePartId As Object = SqlSkalar(connStr,
                "SELECT sprv.boundary_id FROM sys.partition_functions spf JOIN sys.partition_range_values sprv ON sprv.function_id=spf.function_id WHERE spf.name='" & pf & "' AND sprv.value=" & partName,
                "Neue PartID")

            ' CHECK constraint
            SqlAusfuehren(connStr,
                "ALTER TABLE dbo.[" & tmpTabelle & "] WITH CHECK ADD CONSTRAINT [CK_" & tmpTabelle & "] CHECK([" & v.PartitionsSpalte & "]<=" & partName & " AND [" & v.PartitionsSpalte & "]>" & partWert & " AND [" & v.PartitionsSpalte & "] IS NOT NULL);",
                "CHECK Constraint")

            ' SWITCH IN
            SqlAusfuehren(connStr,
                "ALTER TABLE dbo.[" & tmpTabelle & "] SWITCH TO dbo.[" & v.Faktentabelle & "] PARTITION " & Convert.ToInt32(neuePartId) & ";",
                "SWITCH IN")

            ' Cleanup
            SqlAusfuehren(connStr, "DROP TABLE dbo.[" & tmpTabelle & "];", "Tmp loeschen")
        End If

    End Sub

    ' =========================================================================
    ' Spaltendefinition
    ' =========================================================================
    Private Function HoleSpaltendefinition(connStr As String, faktentabelle As String) As String
        Dim sql As String =
"SELECT STUFF((SELECT ', '+QUOTENAME(c.name)+' '+
    CASE WHEN y.name IN ('char','nchar','binary') THEN y.name+'('+LTRIM(STR(c.max_length))+')'
         WHEN y.name IN ('varchar','nvarchar','varbinary') THEN y.name+'('+CASE WHEN c.max_length=-1 THEN 'max' ELSE LTRIM(STR(c.max_length)) END+')'
         WHEN y.name IN ('decimal','numeric') THEN y.name+'('+LTRIM(STR(c.precision))+','+LTRIM(STR(c.scale))+')'
         WHEN y.name IN ('datetime2','datetimeoffset','time') THEN y.name+'('+LTRIM(STR(c.scale))+')'
         ELSE y.name END+' '+CASE WHEN c.is_nullable=1 THEN 'NULL' ELSE 'NOT NULL' END
    FROM sys.columns c JOIN sys.types y ON c.user_type_id=y.user_type_id
    WHERE c.object_id=OBJECT_ID('dbo." & faktentabelle & "') AND c.is_computed=0
    ORDER BY c.column_id FOR XML PATH(''),TYPE).value('.','nvarchar(max)'),1,2,'')"
        Return Convert.ToString(SqlSkalar(connStr, sql, "Spaltendefinition"))
    End Function

    ' =========================================================================
    ' Verfahren laden
    ' =========================================================================
    Private Function VerfahrenLaden(connStr As String) As List(Of VerfahrenInfo)
        Dim liste As New List(Of VerfahrenInfo)()
        Dim sql As String =
"SELECT a.ID, a.Verfahren, a.Themengebiet, a.LetzterSchritt,
        pf.Wert AS Faktentabelle, pp.Wert AS PartitionsSpalte
 FROM   dbo.ETL_Fkt_Arbeitsliste a
 JOIN   " & _parameterDB & ".dbo." & _parametertab & " pf ON pf.Verfahren=a.Verfahren AND pf.Parameter='Faktentabelle'
 JOIN   " & _parameterDB & ".dbo." & _parametertab & " pp ON pp.Verfahren=a.Verfahren AND pp.Parameter='Faktenpartitionsspalte'
 WHERE  a.Status IN ('FAKTENTABELLE_ERSTELLT','PARTITIONSGRENZEN')
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
                                Dim rohPart As String = rdr(5).ToString().Trim()
                                liste.Add(New VerfahrenInfo With {
                                    .ID = rdr.GetInt32(0), .Verfahren = rdr(1).ToString().Trim(),
                                    .Themengebiet = rdr(2).ToString().Trim(),
                                    .LetzterSchritt = If(rdr.IsDBNull(3), "", rdr(3).ToString().Trim()),
                                    .Faktentabelle = rdr(4).ToString().Trim(),
                                    .PartitionsSpalte = If(rohPart.Contains("|"), rohPart.Substring(0, rohPart.IndexOf("|")), rohPart)})
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

    ' =========================================================================
    ' Helper functions
    ' =========================================================================
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

    ' =========================================================================
    ' Datenklassen
    ' =========================================================================
    Private Class PartitionsEintrag
        Public Property Wert As Integer
        Public Property Modus As String
    End Class

    Private Class VerfahrenInfo
        Public Property ID As Integer
        Public Property Verfahren As String
        Public Property Themengebiet As String
        Public Property LetzterSchritt As String
        Public Property Faktentabelle As String
        Public Property PartitionsSpalte As String
    End Class

    Public Enum ScriptResults
        Success = DTSExecResult.Success
        Failure = DTSExecResult.Failure
    End Enum

End Class
