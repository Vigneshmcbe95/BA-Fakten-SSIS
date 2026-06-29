Option Explicit On
Option Strict On

Imports System
Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports System.Text.RegularExpressions
Imports Microsoft.SqlServer.Dts.Runtime

' =============================================================================
'  Skript       : SC04_Whereclause
'  Paket        : Fakten Laden (SSIS)
'  Zweck        : Baut je Verfahren die WHERE-Klausel aus der Steuerliste
'                 auf (basierend auf dem Referenzdatum) und schreibt sie
'                 zurueck.
'  Wiederholung : 3 Versuche je SQL-Anweisung, 30 s Wartezeit
'  Protokoll    : Nur SSIS-Events (FireInformation / FireError)
' =============================================================================
<Microsoft.SqlServer.Dts.Tasks.ScriptTask.SSISScriptTaskEntryPointAttribute()>
<CLSCompliant(False)>
Partial Public Class ScriptMain
    Inherits Microsoft.SqlServer.Dts.Tasks.ScriptTask.VSTARTScriptObjectModelBase

    Private Const SKRIPT_NAME As String = "SC04_Whereclause"
    Private Const CONN_NAME As String = "Verbindung"
    Private Const MAX_VERSUCHE As Integer = 3
    Private Const WARTE_SEK As Integer = 30

    ' Referenzdatum = aktueller Systemmonat (1. des laufenden Monats).
    ' Tag fix auf 1 gesetzt, damit AddMonths/AddYears keine Monatsend-
    ' Verschiebungen erzeugt (z.B. 31.03 -> 28.02). Relative Filter
    ' (:LAST_*, :MONID(-n), :YEAR(-n)) werden so vom echten Laufdatum
    ' aus berechnet, nicht von einem fest verdrahteten Datum.
    Private ReadOnly _refDatum As DateTime = New Date(Date.Today.Year, Date.Today.Month, 1)
    Private _stlTabelle As String = String.Empty
    Private _parameterDB As String = String.Empty
    Private _parametertab As String = String.Empty

    ' -----------------------------------------------------------------------
    ' Main - Einstiegspunkt - steuert den Ablauf des Skripts.
    ' -----------------------------------------------------------------------
    Public Sub Main()

        Log("SC04_Whereclause - Start")
        Log("Referenzdatum: " & _refDatum.ToString("dd.MM.yyyy HH:mm:ss"))

        Try
            _stlTabelle = Dts.Variables("BA::SteuerlistenTabelle").Value.ToString().Trim()
            _parameterDB = Dts.Variables("BA::ParameterDB").Value.ToString().Trim()
            _parametertab = Dts.Variables("BA::Parametertabelle").Value.ToString().Trim()
            Log("Steuerlisten-Tabelle : dbo." & _stlTabelle)

            Dim connStr As String = HoleVerbindungszeichenfolge()

            SpaltenSicherstellen(connStr)

            ' Es wird nur partition_wert befuellt - die Partitionsspalte
            ' kommt in allen Folgeskripten aus der Parametertabelle und
            ' wird hier nicht benoetigt. SCR11 erkennt den MANUAL-Modus
            ' am gefuellten partition_wert.
            Dim zeilen As List(Of Zeile) = ZeilenLaden(connStr)
            Log("Steuerlisten-Zeilen : " & zeilen.Count.ToString())

            Dim cntWert As Integer = 0
            Dim cntMuster As Integer = 0
            Dim cntOhne As Integer = 0

            For Each z As Zeile In zeilen

                ' Ergebnis: ENTWEDER ein Einzelwert (partition_wert -> Cut-Off in SCR11)
                ' ODER ein Regex-Muster (where_klausel -> Mengen-Treffer in SCR11).
                Dim wert As String = Nothing
                Dim muster As String = Nothing
                FilterAnalysieren(z.TabnameFilter, wert, muster)

                If muster IsNot Nothing Then
                    cntMuster += 1
                ElseIf wert IsNot Nothing Then
                    cntWert += 1
                Else
                    cntOhne += 1
                End If

                Log(String.Format("  {0,-50} -> wert={1} | muster={2}",
                    z.TabnameFilter,
                    If(wert IsNot Nothing, wert, "NULL"),
                    If(muster IsNot Nothing, muster, "NULL")))

                ZurueckSchreiben(connStr, z.TabnameFilter, z.FileName, muster, wert)
            Next

            Log("partition_wert (Cut-Off) gefuellt: " & cntWert.ToString())
            Log("where_klausel  (Regex-Muster)    : " & cntMuster.ToString())
            Log("ohne Treffer                     : " & cntOhne.ToString())
            Dts.TaskResult = ScriptResults.Success

        Catch ex As Exception
            LogFehler("Kritischer Fehler: " & ex.Message)
            Dts.TaskResult = ScriptResults.Failure
        End Try

    End Sub

    ' -----------------------------------------------------------------------
    ' FilterAnalysieren - Wandelt einen Steuerlisten-Filter in GENAU EINES um:
    '   * partWert : ein einzelner numerischer Wert (Einzelmonat/-jahr, Datums-
    '                token oder fester Wert) -> in SCR11 als Cut-Off interpretiert
    '                (Full Load aller Partitionen <= Wert).
    '   * regexMuster: ein .NET-Regex, der gegen die REALEN Partitionswerte
    '                  matcht (Wildcards * ? +, Klammern [..], IN-Listen,
    '                  LAST_*-Fenster) -> in SCR11 als Mengen-Treffer expandiert.
    ' Spiegelt die Oracle-Logik (MonId/YearId/getKlausel/getPosReg + star2regex).
    ' -----------------------------------------------------------------------
    Private Sub FilterAnalysieren(filter As String, ByRef partWert As String, ByRef regexMuster As String)

        partWert = Nothing
        regexMuster = Nothing
        If String.IsNullOrEmpty(filter) Then Return

        ' Datumstoken (:MONID(n) :MONID6(n) :MONID4(n) :YEAR(n)) zuerst aufloesen
        Dim f As String = TokenAufloesen(filter)
        Dim m As Match

        ' :LAST_MM / :LAST_YYMM / :LAST_YYYYMM(n) -> Fenster der letzten n Monate
        m = Regex.Match(f, ":LAST_(?:MM|YYMM|YYYYMM)\((\d+)\)", RegexOptions.IgnoreCase)
        If m.Success Then
            regexMuster = MonatsfensterRegex(CInt(m.Groups(1).Value))
            Return
        End If

        ' :LAST_YY / :LAST_YYYY(n) -> Fenster der letzten n Jahre
        m = Regex.Match(f, ":LAST_(?:YY|YYYY)\((\d+)\)", RegexOptions.IgnoreCase)
        If m.Success Then
            regexMuster = JahresfensterRegex(CInt(m.Groups(1).Value))
            Return
        End If

        ' :YYYYMM(v1,v2,...) -> Praefix-IN-Liste (6-stellig)
        m = Regex.Match(f, ":YYYYMM\(([^)]+)\)", RegexOptions.IgnoreCase)
        If m.Success Then
            regexMuster = ListeZuRegex(m.Groups(1).Value)
            Return
        End If

        ' :YYYY(v1,v2,...) -> Praefix-IN-Liste (4-stellig)
        m = Regex.Match(f, ":YYYY\(([^)]+)\)", RegexOptions.IgnoreCase)
        If m.Success Then
            regexMuster = ListeZuRegex(m.Groups(1).Value)
            Return
        End If

        ' Partitionsteil bestimmen: nach letztem ':' (Token) bzw. letztem '_' (Suffix)
        Dim sel As String = f
        Dim pc As Integer = sel.LastIndexOf(":"c)
        If pc >= 0 Then
            sel = sel.Substring(pc + 1)
        Else
            Dim uc As Integer = sel.LastIndexOf("_"c)
            If uc >= 0 Then sel = sel.Substring(uc + 1)
        End If
        sel = sel.Trim()

        ' Enthaelt der Selektor Regex-/Wildcard-Metazeichen? -> Mengen-Muster
        If Regex.IsMatch(sel, "[\*\?\+\[\]\-]") Then
            regexMuster = MusterZuRegex(sel)
            Return
        End If

        ' Sonst: einzelner fester Wert (Ziffern) -> Cut-Off
        m = Regex.Match(f, "((?:19|20)\d{4}\d*)")
        If m.Success Then
            partWert = m.Groups(1).Value
            Return
        End If

        ' Bare 4-stelliges YYMM-Suffix (z.B. _1502)
        m = Regex.Match(f, "_(\d{4})$")
        If m.Success Then partWert = m.Groups(1).Value

    End Sub

    ' -----------------------------------------------------------------------
    ' MusterZuRegex - Wandelt ein Steuerlisten-Wildcardmuster in einen
    ' verankerten .NET-Regex. '*' -> '.{0,128}' (wie Oracle star2regex);
    ' '?', '+', '[..]', '-' werden direkt als Regex uebernommen, Ziffern
    ' bleiben literal.
    ' -----------------------------------------------------------------------
    Private Function MusterZuRegex(muster As String) As String
        Return "^" & muster.Replace("*", ".{0,128}") & "$"
    End Function

    ' -----------------------------------------------------------------------
    ' ListeZuRegex - Wandelt eine IN-Liste '(v1,v2,...)' bzw. 'v1,v2,...' in
    ' eine Praefix-Alternation '^(v1|v2|...)' um (matcht Partitionswerte, die
    ' mit einem der Werte beginnen - entspricht Oracle substr(col,1,len) IN).
    ' -----------------------------------------------------------------------
    Private Function ListeZuRegex(liste As String) As String
        Dim roh As String = liste.Trim().TrimStart("("c).TrimEnd(")"c)
        Dim teile As New List(Of String)()
        For Each t As String In roh.Split(","c)
            Dim s As String = t.Trim().Trim("'"c).Trim()
            If s.Length > 0 Then teile.Add(Regex.Escape(s))
        Next
        If teile.Count = 0 Then Return Nothing
        Return "^(" & String.Join("|", teile) & ")"
    End Function

    ' -----------------------------------------------------------------------
    ' MonatsfensterRegex - Erzeugt eine Praefix-Alternation der letzten n
    ' Monate (inkl. Referenzmonat), z.B. ^(202601|202602|...).
    ' -----------------------------------------------------------------------
    Private Function MonatsfensterRegex(n As Integer) As String
        Dim teile As New List(Of String)()
        For i As Integer = n To 0 Step -1
            teile.Add(_refDatum.AddMonths(-i).ToString("yyyyMM"))
        Next
        Return "^(" & String.Join("|", teile) & ")"
    End Function

    ' -----------------------------------------------------------------------
    ' JahresfensterRegex - Erzeugt eine Praefix-Alternation der letzten n
    ' Jahre (inkl. Referenzjahr), z.B. ^(2024|2025|2026).
    ' -----------------------------------------------------------------------
    Private Function JahresfensterRegex(n As Integer) As String
        Dim teile As New List(Of String)()
        For i As Integer = n To 0 Step -1
            teile.Add(_refDatum.AddYears(-i).ToString("yyyy"))
        Next
        Return "^(" & String.Join("|", teile) & ")"
    End Function

    ' -----------------------------------------------------------------------
    ' TokenAufloesen - Loest Platzhalter-Token in einer Vorlage auf.
    ' -----------------------------------------------------------------------
    Private Function TokenAufloesen(f As String) As String
        Dim r As String = f
        Dim m As Match

        ' :MONID4(n) -> 2-stelliges Jahr + Monat (YYMM). ZUERST aufloesen, damit
        ' das :MONID6?-Muster es nicht faelschlich anfasst.
        m = Regex.Match(r, ":MONID4\((-?\d+)\)", RegexOptions.IgnoreCase)
        While m.Success
            r = r.Replace(m.Value, _refDatum.AddMonths(CInt(m.Groups(1).Value)).ToString("yyMM"))
            m = Regex.Match(r, ":MONID4\((-?\d+)\)", RegexOptions.IgnoreCase)
        End While

        ' :MONID(n) / :MONID6(n) -> 4-stelliges Jahr + Monat (YYYYMM)
        m = Regex.Match(r, ":MONID6?\((-?\d+)\)", RegexOptions.IgnoreCase)
        While m.Success
            r = r.Replace(m.Value, _refDatum.AddMonths(CInt(m.Groups(1).Value)).ToString("yyyyMM"))
            m = Regex.Match(r, ":MONID6?\((-?\d+)\)", RegexOptions.IgnoreCase)
        End While

        m = Regex.Match(r, ":YEAR\((-?\d+)\)", RegexOptions.IgnoreCase)
        While m.Success
            r = r.Replace(m.Value, _refDatum.AddYears(CInt(m.Groups(1).Value)).ToString("yyyy"))
            m = Regex.Match(r, ":YEAR\((-?\d+)\)", RegexOptions.IgnoreCase)
        End While

        Return r
    End Function

    ' -----------------------------------------------------------------------
    ' SpaltenSicherstellen - Stellt sicher, dass benoetigte Spalten auf der
    ' Zieltabelle existieren (ergaenzt fehlende).
    ' -----------------------------------------------------------------------
    Private Sub SpaltenSicherstellen(connStr As String)
        SqlRun(connStr,
"IF NOT EXISTS(SELECT 1 FROM sys.columns WHERE object_id=OBJECT_ID('dbo." & _stlTabelle & "') AND name='where_klausel')
    ALTER TABLE dbo." & _stlTabelle & " ADD where_klausel  NVARCHAR(1000) NULL;
IF NOT EXISTS(SELECT 1 FROM sys.columns WHERE object_id=OBJECT_ID('dbo." & _stlTabelle & "') AND name='partition_wert')
    ALTER TABLE dbo." & _stlTabelle & " ADD partition_wert NVARCHAR(500)  NULL;")
        Log("Spalten geprueft/angelegt OK")
    End Sub

    ' -----------------------------------------------------------------------
    ' ZeilenLaden - Laedt die Zeilen einer SQL-Abfrage in eine Liste.
    ' -----------------------------------------------------------------------
    Private Function ZeilenLaden(connStr As String) As List(Of Zeile)
        Dim liste As New List(Of Zeile)()
        Dim sql As String =
            "SELECT tabname_filter, tabelle, FILE_NAME " &
            "FROM   dbo." & _stlTabelle & " " &
            "WHERE  tabname_filter IS NOT NULL " &
            "AND    LTRIM(RTRIM(tabname_filter)) <> '' " &
            "ORDER  BY FILE_NAME, tabname_filter"

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
                                liste.Add(New Zeile With {
                                    .TabnameFilter = rdr(0).ToString(),
                                    .Tabelle = rdr(1).ToString(),
                                    .FileName = rdr(2).ToString()})
                            End While
                        End Using
                    End Using
                End Using
                Return liste
            Catch ex As Exception
                Log(String.Format("WARNUNG [Laden] {0}/{1}: {2}", versuch, MAX_VERSUCHE, ex.Message))
                If versuch < MAX_VERSUCHE Then System.Threading.Thread.Sleep(WARTE_SEK * 1000) Else Throw
            End Try
        End While
        Return liste
    End Function

    ' -----------------------------------------------------------------------
    ' ZurueckSchreiben - Schreibt die erzeugte WHERE-Klausel in die
    ' Steuerliste zurueck.
    ' -----------------------------------------------------------------------
    Private Sub ZurueckSchreiben(connStr As String, tabFilter As String, fileName As String,
                                  whereKlausel As String, partWert As String)
        ' Explizit NULL schreiben wenn where_klausel oder partition_wert nicht ermittelt
        Dim versuch As Integer = 0
        While versuch < MAX_VERSUCHE
            versuch += 1
            Try
                Using conn As New SqlConnection(connStr)
                    conn.Open()
                    Using cmd As New SqlCommand(
                        "UPDATE dbo." & _stlTabelle &
                        " SET where_klausel=@w, partition_wert=@p" &
                        " WHERE tabname_filter=@z AND FILE_NAME=@f", conn)
                        cmd.CommandTimeout = 0
                        cmd.Parameters.AddWithValue("@w", If(whereKlausel Is Nothing, CObj(DBNull.Value), CObj(whereKlausel)))
                        cmd.Parameters.AddWithValue("@p", If(partWert Is Nothing, CObj(DBNull.Value), CObj(partWert)))
                        cmd.Parameters.AddWithValue("@z", tabFilter)
                        cmd.Parameters.AddWithValue("@f", fileName)
                        cmd.ExecuteNonQuery()
                        Return
                    End Using
                End Using
            Catch ex As Exception
                If versuch < MAX_VERSUCHE Then System.Threading.Thread.Sleep(WARTE_SEK * 1000) Else Throw
            End Try
        End While
    End Sub

    ' -----------------------------------------------------------------------
    ' SqlRun - Fuehrt eine SQL-Anweisung mit Wiederholung aus.
    ' -----------------------------------------------------------------------
    Private Sub SqlRun(connStr As String, sql As String)
        Using conn As New SqlConnection(connStr)
            conn.Open()
            Using cmd As New SqlCommand(sql, conn)
                cmd.CommandTimeout = 0
                cmd.ExecuteNonQuery()
            End Using
        End Using
    End Sub

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
    ' Zeile
    ' -----------------------------------------------------------------------
    Private Class Zeile
        Public Property TabnameFilter As String
        Public Property Tabelle As String
        Public Property FileName As String
    End Class

    ' -----------------------------------------------------------------------
    ' ScriptResults - SSIS-Task-Ergebniscodes.
    ' -----------------------------------------------------------------------
    Public Enum ScriptResults
        Success = DTSExecResult.Success
        Failure = DTSExecResult.Failure
    End Enum

End Class
