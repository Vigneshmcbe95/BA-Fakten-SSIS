Option Explicit On
Option Strict On

Imports System
Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports System.Linq
Imports Microsoft.SqlServer.Dts.Runtime

' =============================================================================
'  Skript       : SCR11_Partitionsgrenzen
'  Paket        : Fakten Laden (SSIS)
'  Zweck        : Ermittelt die zu ladenden Partitionswerte (Oracle vs.
'                 MSSQL Delta), fuehrt SPLIT auf der Partitionsfunktion aus
'                 und setzt BA::objPartitionValues.
'  Ablauf       : FAKTENTABELLE_ERSTELLT -> PARTITIONSGRENZEN_ERSTELLT
'  Wiederholung : 3 Versuche je SQL-Anweisung, 30 s Wartezeit
'  Protokoll    : Nur SSIS-Events (FireInformation / FireError)
' =============================================================================
<Microsoft.SqlServer.Dts.Tasks.ScriptTask.SSISScriptTaskEntryPointAttribute()>
<CLSCompliant(False)>
Partial Public Class ScriptMain
    Inherits Microsoft.SqlServer.Dts.Tasks.ScriptTask.VSTARTScriptObjectModelBase

    Private Const SKRIPT_NAME As String = "SCR11_Partitionsgrenzen"
    Private Const CONN_NAME As String = "Verbindung"
    Private Const MAX_VERSUCHE As Integer = 3
    Private Const WARTE_SEK As Integer = 30

    Private _runID As Integer = 0
    Private _parameterDB As String = String.Empty
    Private _parametertab As String = String.Empty
    Private _stlTabelle As String = String.Empty
    Private _partitionSchema As String = String.Empty

    ' -----------------------------------------------------------------------
    ' Main - Einstiegspunkt - steuert den Ablauf des Skripts.
    ' -----------------------------------------------------------------------
    Public Sub Main()

        Log("SCR11_Partitionsgrenzen - Start")

        Try
            _runID = Convert.ToInt32(Dts.Variables("BA::RunID").Value)
            _parameterDB = Dts.Variables("BA::ParameterDB").Value.ToString().Trim()
            _parametertab = Dts.Variables("BA::Parametertabelle").Value.ToString().Trim()
            _stlTabelle = Dts.Variables("BA::SteuerlistenTabelle").Value.ToString().Trim()
            _partitionSchema = Dts.Variables("BA::partition_schema").Value.ToString().Trim()

            Dim connStr As String = HoleVerbindungszeichenfolge()
            Dim verfahren As List(Of VerfahrenInfo) = VerfahrenLaden(connStr)
            Log("Verfahren: " & verfahren.Count.ToString())

            Dim gesamtErgebnis As New Dictionary(Of String, List(Of PartitionsEintrag))()

            Dim cntOK As Integer = 0
            Dim cntFehler As Integer = 0

            For Each v As VerfahrenInfo In verfahren
                Log("Verfahren: " & v.Verfahren & " | Tabelle: " & v.Faktentabelle & " | Spalte: " & v.PartitionsSpalte)

                If v.LetzterSchritt = "PARTITIONSGRENZEN_ERSTELLT" Then
                    Log("  bereits abgeschlossen uebersprungen OK")
                    Continue For
                End If

                Try
                    StatusSetzen(connStr, v.ID, "PARTITIONSGRENZEN")

                    ' ═════════════════════════════════════════════════════════
                    ' SCHRITT 1: partition_wert pruefen → Modus bestimmen
                    ' (MANUAL sobald der Benutzer in der STL-Datei einen
                    '  konkreten Wert mitgegeben hat - SC04 schreibt ihn in
                    '  partition_wert; where_klausel wird nicht mehr genutzt)
                    ' ═════════════════════════════════════════════════════════
                    Dim benutzerWerte As List(Of Integer) = PartitionWerteLaden(connStr, v.Verfahren)

                    Dim zuVerarbeiten As New List(Of PartitionsEintrag)()
                    Dim modus As String = String.Empty
                    Dim oracleAlleWerte As List(Of Integer) = Nothing
                    Dim mssqlWerte As List(Of Integer) = Nothing

                    If benutzerWerte.Count > 0 Then
                        ' ═════════════════════════════════════════════════════
                        ' MODE 1: MANUAL (partition_wert gesetzt)
                        ' Oracle-Pruefung wird uebersprungen - partition_wert
                        ' direkt aus Steuerliste verwenden
                        ' ═════════════════════════════════════════════════════
                        modus = "MANUAL"
                        Log("  MODE: MANUAL (partition_wert gesetzt - Oracle-Pruefung uebersprungen)")
                        Log("  Benutzer partition_wert: " & benutzerWerte.Count.ToString() & " Werte")
                        If benutzerWerte.Count > 0 Then
                            Log("  MIN: " & benutzerWerte.Min().ToString() & " | MAX: " & benutzerWerte.Max().ToString())
                        End If

                        ' MSSQL Status pruefen - NEU vs AKTUALISIERUNG
                        mssqlWerte = MssqlWerteLaden(connStr, v)
                        Log("  MSSQL Werte gesamt: " & mssqlWerte.Count.ToString())

                        ' Stelligkeit von Eingabe- und MSSQL-Werten muss zusammenpassen,
                        ' sonst ist die Partitionsspalte nicht stimmig -> Abbruch mit Meldung.
                        StelligkeitPruefen(benutzerWerte, mssqlWerte, v)

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
                        ' ═════════════════════════════════════════════════════
                        ' MODE 2: AUTOMATIC (kein partition_wert in CSV)
                        ' ═════════════════════════════════════════════════════
                        modus = "AUTOMATIC"
                        Log("  MODE: AUTOMATIC (partition_wert aus Oracle)")

                        ' Alle Oracle-Werte laden
                        oracleAlleWerte = OracleAlleWerteLaden(connStr, v)
                        Log("  Oracle Werte gesamt: " & oracleAlleWerte.Count.ToString())

                        If oracleAlleWerte.Count = 0 Then
                            Dim msg As String = "Keine Partitionen in Oracle gefunden (View v_partition_info)"

                            LogFehler("  FEHLER: " & msg)
                            ProtokollSchreiben(connStr, v.Verfahren, "FEHLER_SCR11", msg)

                            FehlerSetzen(connStr, v.ID, msg)

                            Throw New Exception(msg)
                        End If

                        ' MSSQL Status pruefen
                        mssqlWerte = MssqlWerteLaden(connStr, v)

                        ' Stelligkeit von Oracle- und MSSQL-Werten muss zusammenpassen,
                        ' sonst ist die Partitionsspalte nicht stimmig -> Abbruch mit Meldung
                        ' (statt still als "bereits geladen" zu ueberspringen).
                        StelligkeitPruefen(oracleAlleWerte, mssqlWerte, v)

                        If mssqlWerte.Count = 0 Then
                            ' ─── FULL LOAD: MSSQL ist leer ───
                            ' Kein MSSQL-Minimum zum Vergleich vorhanden -> die
                            ' niedrigsten Oracle-Werte koennen leere Grenzpartitionen
                            ' sein (1. Partition ohne Daten). Diese von unten abschneiden:
                            ' den jeweils niedrigsten Wert per Einzel-Partition-Pruefung
                            ' testen (WHERE partcol = wert liest dank Pushdown nur EINE
                            ' Partition, kein Full Scan) und bei 0 Zeilen entfernen.
                            Dim vollWerte As List(Of Integer) = oracleAlleWerte.OrderBy(Function(w) w).ToList()
                            Dim idx As Integer = 0
                            While idx < vollWerte.Count
                                If PartitionHatDaten(connStr, v, vollWerte(idx)) Then Exit While
                                Log("  Leere untere Grenzpartition entfernt: " & vollWerte(idx).ToString())
                                idx += 1
                            End While
                            Dim ladeWerte As List(Of Integer) = vollWerte.Skip(idx).ToList()

                            If ladeWerte.Count = 0 Then
                                Log("  Keine Daten in Oracle (alle Grenzpartitionen leer) - kein Ladevorgang noetig.")
                                ProtokollSchreiben(connStr, v.Verfahren, "SCHRITT_4",
                                    "Keine Daten in Oracle - alle Grenzpartitionen leer.")
                                StatusSetzen(connStr, v.ID, "PARTITIONSGRENZEN_ERSTELLT")
                                cntOK += 1
                                Continue For
                            End If

                            Log("  FULL LOAD (MSSQL leer) | Oracle: MIN=" & ladeWerte.Min().ToString() &
                                " MAX=" & ladeWerte.Max().ToString() &
                                " | " & ladeWerte.Count.ToString() & " Werte")
                            For Each ow As Integer In ladeWerte
                                zuVerarbeiten.Add(New PartitionsEintrag With {.Wert = ow, .Modus = "NEU"})
                            Next

                        Else
                            ' ─── APPEND: nur echte neue Werte laden ───
                            Dim oMax As Integer = oracleAlleWerte.Max()
                            Dim mMin As Integer = mssqlWerte.Min()
                            Dim mMax As Integer = mssqlWerte.Max()

                            ' Zu laden = Oracle-Werte, die in MSSQL fehlen UND nicht
                            ' unterhalb des MSSQL-Minimums liegen. Werte unter mMin sind
                            ' leere Oracle-Grenzpartitionen (1. Partition ohne Daten -
                            ' 6- und 8-stelliges Schema) und werden ausgeblendet, daher
                            ' wird die Oracle-MIN/MAX-Angabe nach dieser Bereinigung
                            ' ermittelt (kein Fehlalarm).
                            Dim zuLaden As List(Of Integer) =
                                oracleAlleWerte.Where(Function(w) w >= mMin AndAlso Not mssqlWerte.Contains(w)) _
                                               .OrderBy(Function(w) w).ToList()

                            ' Oracle-MIN nach Bereinigung (leere Grenzpartitionen < mMin ausgeblendet)
                            Dim oMinEff As Integer = oracleAlleWerte.Where(Function(w) w >= mMin).DefaultIfEmpty(mMin).Min()

                            Log("  Oracle: MIN=" & oMinEff.ToString() & " MAX=" & oMax.ToString() &
                                " | MSSQL: MIN=" & mMin.ToString() & " MAX=" & mMax.ToString())

                            If zuLaden.Count = 0 Then
                                Log("  Alle Daten bereits geladen - kein neuer Ladevorgang noetig.")
                                ProtokollSchreiben(connStr, v.Verfahren, "SCHRITT_4",
                                    "Alle Daten bereits geladen - kein neuer Ladevorgang noetig. " &
                                    "Oracle " & oMinEff.ToString() & "-" & oMax.ToString() &
                                    " | MSSQL " & mMin.ToString() & "-" & mMax.ToString())
                                StatusSetzen(connStr, v.ID, "PARTITIONSGRENZEN_ERSTELLT")
                                cntOK += 1
                                Continue For
                            End If

                            Log("  Zu laden: " & zuLaden.Count.ToString() & " Partition(en) (" &
                                zuLaden.Min().ToString() & "-" & zuLaden.Max().ToString() & ")")

                            For Each nw As Integer In zuLaden
                                zuVerarbeiten.Add(New PartitionsEintrag With {.Wert = nw, .Modus = "NEU"})
                            Next
                        End If
                    End If

                    ' ═════════════════════════════════════════════════════════
                    ' ZUSAMMENFASSUNG (COMPACT)
                    ' ═════════════════════════════════════════════════════════
                    zuVerarbeiten = zuVerarbeiten.OrderBy(Function(z) z.Wert).ToList()

                    Log("  ZUSAMMENFASSUNG: " & v.Verfahren)
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

                    If zuVerarbeiten.Count = 0 Then
                        Log("  Keine Partitionswerte zu verarbeiten uebersprungen")
                        StatusSetzen(connStr, v.ID, "PARTITIONSGRENZEN_ERSTELLT")
                        cntOK += 1
                        Continue For
                    End If

                    ' ═════════════════════════════════════════════════════════
                    ' PARTITIONSGRENZEN ERSTELLEN (nur fuer NEU Werte)
                    ' ═════════════════════════════════════════════════════════
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

                    Log("  SPLIT ausgefuehrt: " & cntSplit.ToString() & " | Uebersprungen: " & cntSkip.ToString())

                    ' ═════════════════════════════════════════════════════════
                    ' ERGEBNIS SPEICHERN
                    ' ═════════════════════════════════════════════════════════
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
                    Log("  Schritt 4 abgeschlossen OK")

                Catch ex As Exception
                    cntFehler += 1
                    FehlerSetzen(connStr, v.ID, ex.Message)
                    ProtokollSchreiben(connStr, v.Verfahren, "FEHLER_SCR09", ex.Message)
                    LogFehler("FEHLER '" & v.Verfahren & "': " & ex.Message)
                End Try
            Next

            ' ═════════════════════════════════════════════════════════
            ' BA::objPartitionValues SETZEN
            ' ═════════════════════════════════════════════════════════
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

                Log("BA::objPartitionValues gesetzt: " & gesamtAnzahl.ToString() & " Eintraege")

                ' Show summary by Verfahren (compact)
                For Each kvp As KeyValuePair(Of String, List(Of PartitionsEintrag)) In gesamtErgebnis
                    Dim minV As Integer = kvp.Value.Min(Function(p) p.Wert)
                    Dim maxV As Integer = kvp.Value.Max(Function(p) p.Wert)
                    Log("  " & kvp.Key & ": " & kvp.Value.Count.ToString() & " Partitionen (" & minV.ToString() & "-" & maxV.ToString() & ")")
                Next

            Else
                Dts.Variables("BA::objPartitionValues").Value = Nothing
                Log("BA::objPartitionValues: leer (Nothing)")
            End If

            Log("Erfolgreich: " & cntOK.ToString() & " | Fehler: " & cntFehler.ToString())
            Dts.TaskResult = If(cntFehler > 0, ScriptResults.Failure, ScriptResults.Success)

        Catch ex As Exception
            LogFehler("Kritischer Fehler: " & ex.Message)
            Dts.TaskResult = ScriptResults.Failure
        End Try

    End Sub

    ' -----------------------------------------------------------------------
    ' PartitionWerteLaden - Laedt die Partitionswerte eines Verfahrens.
    ' Sind Werte vorhanden, wird MANUAL-Modus ohne Oracle-Pruefung verwendet.
    ' -----------------------------------------------------------------------
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

    ' -----------------------------------------------------------------------
    ' OracleAlleWerteLaden - Laedt alle eindeutigen Partitionswerte aus der
    ' Oracle-Partitionssicht ext.v_partition_info (DISTINCT HIGH_VALUE je
    ' Faktentabelle, Owner = BA::partition_schema). HIGH_VALUE enthaelt den
    ' bereits bereinigten ganzzahligen Partitionswert; nicht-numerische Werte
    ' werden uebersprungen.
    ' -----------------------------------------------------------------------
    Private Function OracleAlleWerteLaden(connStr As String, v As VerfahrenInfo) As List(Of Integer)
        Dim liste As New List(Of Integer)()
        Dim sql As String =
            "SELECT DISTINCT HIGH_VALUE FROM ext.[v_partition_info] " &
            "WHERE TABLE_NAME = UPPER('" & v.Verfahren & "') " &
            "  AND OWNER      = UPPER('" & v.Themengebiet & "') " &
            "  AND HIGH_VALUE IS NOT NULL"

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
                                If Not rdr.IsDBNull(0) Then
                                    Dim intWert As Integer
                                    If Integer.TryParse(rdr(0).ToString().Trim(), intWert) Then
                                        ' ─────────────────────────────────────────────────────
                                        ' Zwei Partitionsschemata, abhaengig vom Wertformat:
                                        '
                                        ' (A) 6-stellige Werte (YYYYMM, z.B. mon_id):
                                        '     HIGH_VALUE entspricht direkt dem Datenwert.
                                        '     Ausnahme: aelterer Jahresend-Marker "YYYY13" ->
                                        '     Januar des Folgejahres (200713 -> 200801,
                                        '     201713 -> 201801). Ab 2018/2019 erscheint Januar
                                        '     direkt als YYYY01, keine Umrechnung noetig.
                                        '
                                        ' (B) 8-stellige Werte (YYYYMM + 2-stellige Anhangzahl,
                                        '     z.B. mow_id): HIGH_VALUE ist die exklusive
                                        '     Obergrenze; der tatsaechliche Datenwert ist
                                        '     HIGH_VALUE - 1 (Anhang 07->06, 13->12, 19->18).
                                        '     Der unterste Anker mit Anhang 00 (z.B. 19990100)
                                        '     traegt keine Daten und wird uebersprungen.
                                        ' ─────────────────────────────────────────────────────
                                        Dim gueltig As Boolean = True
                                        If intWert > 999999 Then
                                            ' (B) 8-stellig: mow_id = HIGH_VALUE - 1
                                            If (intWert Mod 100) = 0 Then
                                                gueltig = False
                                            Else
                                                intWert = intWert - 1
                                            End If
                                        ElseIf (intWert Mod 100) = 13 Then
                                            ' (A) 6-stellig: YYYY13 -> Januar Folgejahr
                                            intWert = (intWert \ 100 + 1) * 100 + 1
                                        End If
                                        If gueltig Then liste.Add(intWert)
                                    End If
                                End If
                            End While
                        End Using
                    End Using
                End Using
                ' Leere obere Grenzpartition nur beim 6-stelligen Schema entfernen:
                ' Dort gilt HIGH_VALUE = Datenwert, der hoechste Wert (z.B. 202606)
                ' ist die offene Folgemonats-Grenze ohne Daten.
                ' Beim 8-stelligen Schema gilt Datenwert = HIGH_VALUE - 1, d.h. der
                ' hoechste Wert ist bereits der letzte echte Datenmonat -> NICHT
                ' entfernen (sonst ginge der aktuellste Monat verloren).
                If liste.Count > 1 Then
                    Dim maxWert As Integer = liste.Max()
                    If maxWert <= 999999 Then
                        liste.RemoveAll(Function(w) w = maxWert)
                        Log("  Leere obere Grenzpartition entfernt: " & maxWert.ToString())
                    End If
                End If
                If liste.Count > 0 Then Return liste
                ' Keine Partitionsmetadaten in v_partition_info - typisch fuer eine
                ' VIEW-Quelle (vf_*): Views haben keine Eintraege in
                ' dba_tab_partitions. Fallback: tatsaechliche Partitionswerte direkt
                ' aus der ext-Quelle lesen (DISTINCT). Diese Werte sind echte
                ' Datenwerte -> keine Umrechnung. Hinweis: dieser Pfad scannt die Quelle.
                Log("  Keine Partitionsmetadaten in v_partition_info fuer [" & v.Verfahren & "] (z.B. View-Quelle) -> Fallback: DISTINCT " & v.PartitionsSpalte & " aus ext.[" & v.Verfahren.ToLower() & "]")
                Return OracleDistinctWerteLaden(connStr, v)
            Catch ex As Exception
                If versuch < MAX_VERSUCHE Then System.Threading.Thread.Sleep(WARTE_SEK * 1000) Else Throw
            End Try
        End While
        Return liste
    End Function

    ' -----------------------------------------------------------------------
    ' OracleDistinctWerteLaden - Fallback fuer Quellen ohne Partitionsmetadaten
    ' (z.B. Views vf_*): liest die tatsaechlich vorhandenen Partitionswerte
    ' direkt aus der ext-Quelle (SELECT DISTINCT <partcol>). Liefert echte
    ' Datenwerte (keine HIGH_VALUE-Umrechnung noetig). Scannt die Quelle.
    ' -----------------------------------------------------------------------
    Private Function OracleDistinctWerteLaden(connStr As String, v As VerfahrenInfo) As List(Of Integer)
        Dim liste As New List(Of Integer)()
        Dim sql As String = "SELECT DISTINCT [" & v.PartitionsSpalte & "] FROM ext.[" & v.Verfahren.ToLower() &
                            "] WHERE [" & v.PartitionsSpalte & "] IS NOT NULL"
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
                                If Not rdr.IsDBNull(0) Then
                                    Dim intWert As Integer
                                    If Integer.TryParse(rdr(0).ToString().Trim(), intWert) Then liste.Add(intWert)
                                End If
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
    ' MssqlWerteLaden - Laedt die bereits vorhandenen Partitionswerte aus
    ' der MSSQL-Faktentabelle.
    ' -----------------------------------------------------------------------
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

    ' -----------------------------------------------------------------------
    ' PartitionSplitDurchfuehren - Ergaenzt fehlende Partitionsgrenzen per
    ' ALTER PARTITION FUNCTION ... SPLIT.
    ' -----------------------------------------------------------------------
    Private Sub PartitionSplitDurchfuehren(connStr As String, v As VerfahrenInfo,
                                            pf As String, ps As String,
                                            dateigruppe As String, partWert As Integer)

        ' Direct check: is this boundary already in the partition function?
        Dim sqlBoundExists As String =
            "SELECT COUNT(*) FROM sys.partition_range_values prv " &
            "JOIN sys.partition_functions pfn ON prv.function_id=pfn.function_id " &
            "WHERE pfn.name='" & pf & "' AND CONVERT(int,prv.value)=" & partWert
        Dim boundCount As Object = SqlSkalar(connStr, sqlBoundExists, "Boundary prüfen")
        If Convert.ToInt32(If(boundCount, 0)) > 0 Then
            Log("    Boundary " & partWert & " bereits vorhanden uebersprungen")
            Return
        End If

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
            ' ═════════════════════════════════════════════════════════════
            ' SWITCH OUT / SPLIT / SWITCH IN (WITH CCI/CI FIX)
            ' ═════════════════════════════════════════════════════════════
            Dim tmpTabelle As String = v.Faktentabelle & "_tmp_" & partName.ToString()
            SqlAusfuehren(connStr, "IF OBJECT_ID('dbo.[" & tmpTabelle & "]','U') IS NOT NULL DROP TABLE dbo.[" & tmpTabelle & "];", "Tmp loeschen")

            Dim spaltenDef As String = HoleSpaltendefinition(connStr, v.Faktentabelle)

            ' Get compression
            Dim komprimierung As String = Convert.ToString(SqlSkalar(connStr,
                "SELECT TOP 1 p.data_compression_desc FROM sys.partitions p JOIN sys.indexes i ON p.object_id=i.object_id AND p.index_id=i.index_id JOIN sys.tables t ON t.object_id=p.object_id WHERE t.name='" & v.Faktentabelle & "' AND p.partition_number=1",
                "Komprimierung"))
            Dim kompStr As String = If(komprimierung = "PAGE" OrElse komprimierung = "ROW", " WITH (DATA_COMPRESSION=" & komprimierung & ")", "")

            ' ═════════════════════════════════════════════════════════════
            ' FIX: Detect source table index type
            ' ═════════════════════════════════════════════════════════════
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

            ' ═════════════════════════════════════════════════════════════
            ' Create matching clustered index on temp table
            ' ═════════════════════════════════════════════════════════════
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

    ' -----------------------------------------------------------------------
    ' HoleSpaltendefinition - Baut die Spalten-DDL-Definition aus den
    ' Metadaten auf.
    ' -----------------------------------------------------------------------
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

    ' -----------------------------------------------------------------------
    ' VerfahrenLaden - Laedt die zu verarbeitenden Verfahren aus der
    ' Arbeitsliste (Join mit der Parametertabelle).
    ' -----------------------------------------------------------------------
    Private Function VerfahrenLaden(connStr As String) As List(Of VerfahrenInfo)
        Dim liste As New List(Of VerfahrenInfo)()
        Dim sql As String =
"SELECT a.ID, a.Verfahren, a.Themengebiet, a.LetzterSchritt,
        pf.Wert AS Faktentabelle, pp.Wert AS PartitionsSpalte
 FROM   dbo.ETL_Fakt_Arbeitsliste a
 JOIN   " & _parameterDB & ".dbo." & _parametertab & " pf ON pf.Verfahren=dbo.fn_ParamVerfahren(a.Verfahren) AND pf.Parameter='Faktentabelle'
 JOIN   " & _parameterDB & ".dbo." & _parametertab & " pp ON pp.Verfahren=dbo.fn_ParamVerfahren(a.Verfahren) AND pp.Parameter='Faktenpartitionsspalte'
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

    ' -----------------------------------------------------------------------
    ' StatusSetzen - Aktualisiert Status / LetzterSchritt einer
    ' Arbeitslisten-Zeile.
    ' -----------------------------------------------------------------------
    Private Sub StatusSetzen(connStr As String, id As Integer, status As String)
        SqlAusfuehren(connStr, "UPDATE dbo.ETL_Fakt_Arbeitsliste SET Status='" & status & "',LetzterSchritt='" & status & "',AktualisiertAm=GETDATE() WHERE ID=" & id, "Status")
    End Sub

    ' -----------------------------------------------------------------------
    ' FehlerSetzen - Markiert eine Arbeitslisten-Zeile als FEHLER und
    ' speichert die Fehlermeldung.
    ' -----------------------------------------------------------------------
    Private Sub FehlerSetzen(connStr As String, id As Integer, meldung As String)
        Try
            Using conn As New SqlConnection(connStr)
                conn.Open()
                Using cmd As New SqlCommand("UPDATE dbo.ETL_Fakt_Arbeitsliste SET Status='FEHLER',Fehlermeldung=@m,AktualisiertAm=GETDATE() WHERE ID=@id", conn)
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
    ' StelligkeitPruefen - Vergleicht die Stellenzahl (Anzahl Ziffern) der
    ' Quell-/Eingabewerte mit der der MSSQL-Werte. Weichen sie ab, passt die
    ' Partitionsspalte zwischen Oracle und MSSQL nicht zusammen (z.B. mon_id
    ' 6-stellig vs mow_id 8-stellig). Dann wird mit klarer Meldung abgebrochen,
    ' statt still als "bereits geladen" zu ueberspringen. Bei leerer MSSQL-Seite
    ' (Erstbefuellung) gibt es nichts zu vergleichen.
    ' -----------------------------------------------------------------------
    Private Sub StelligkeitPruefen(quelleWerte As List(Of Integer), mssqlWerte As List(Of Integer), v As VerfahrenInfo)
        If quelleWerte Is Nothing OrElse quelleWerte.Count = 0 Then Return
        If mssqlWerte Is Nothing OrElse mssqlWerte.Count = 0 Then Return

        Dim qLen As Integer = quelleWerte.Max().ToString().Length
        Dim mLen As Integer = mssqlWerte.Max().ToString().Length

        If qLen <> mLen Then
            Throw New Exception(
                "Partitionsspalten-Stelligkeit weicht ab fuer Verfahren '" & v.Verfahren & "': " &
                "Oracle/Quelle-Werte sind " & qLen.ToString() & "-stellig (z.B. " & quelleWerte.Max().ToString() & "), " &
                "MSSQL-Werte sind " & mLen.ToString() & "-stellig (z.B. " & mssqlWerte.Max().ToString() & "). " &
                "Die Partitionsspalte [" & v.PartitionsSpalte & "] passt nicht zwischen Oracle und MSSQL zusammen " &
                "(mon_id = 6-stellig, mow_id = 8-stellig). Bitte die Partitionsspalte in Oracle auf die korrekte " &
                "Spalte umstellen und den Lauf erneut starten.")
        End If
    End Sub

    ' -----------------------------------------------------------------------
    ' PartitionHatDaten - Prueft per Einzel-Partition-Pushdown, ob die
    ' Oracle-Tabelle fuer den Partitionswert mindestens eine Zeile liefert.
    ' WHERE <partcol> = <wert> liest dank Pushdown nur EINE Partition
    ' (kein Full Scan). True = Daten vorhanden, False = leere Partition.
    ' -----------------------------------------------------------------------
    Private Function PartitionHatDaten(connStr As String, v As VerfahrenInfo, wert As Integer) As Boolean
        ' ext-Tabelle = Verfahrensname (Oracle-Objektname aus der Steuerliste)
        Dim sql As String = "SELECT TOP 1 1 FROM ext.[" & v.Verfahren.ToLower() &
                            "] WHERE [" & v.PartitionsSpalte & "] = " & wert.ToString()
        Dim obj As Object = SqlSkalar(connStr, sql, "Pruefe Daten " & v.Verfahren.ToLower() & "=" & wert.ToString())
        Return obj IsNot Nothing AndAlso obj IsNot DBNull.Value
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
        Public Property Wert As Integer
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
    End Class

    ' -----------------------------------------------------------------------
    ' ScriptResults - SSIS-Task-Ergebniscodes.
    ' -----------------------------------------------------------------------
    Public Enum ScriptResults
        Success = DTSExecResult.Success
        Failure = DTSExecResult.Failure
    End Enum

End Class
