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

    Private ReadOnly _refDatum As DateTime = New Date(2026, 3, 1)
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

            Dim cntMit As Integer = 0
            Dim cntOhne As Integer = 0

            For Each z As Zeile In zeilen

                Dim wert As String = WertExtrahieren(z.TabnameFilter)

                If wert IsNot Nothing Then
                    cntMit += 1
                Else
                    cntOhne += 1
                End If

                Log(String.Format("  {0,-50} -> wert={1}",
                    z.TabnameFilter,
                    If(wert IsNot Nothing, wert, "NULL")))

                ZurueckSchreiben(connStr, z.TabnameFilter, z.FileName, Nothing, wert)
            Next

            Log("partition_wert gefuellt: " & cntMit.ToString())
            Log("partition_wert NULL    : " & cntOhne.ToString())
            Dts.TaskResult = ScriptResults.Success

        Catch ex As Exception
            LogFehler("Kritischer Fehler: " & ex.Message)
            Dts.TaskResult = ScriptResults.Failure
        End Try

    End Sub

    ' -----------------------------------------------------------------------
    ' WertExtrahieren - Extrahiert einen Einzelwert aus einem getrennten
    ' Parameterstring.
    ' -----------------------------------------------------------------------
    Private Function WertExtrahieren(filter As String) As String

        If String.IsNullOrEmpty(filter) Then Return Nothing

        Dim f As String = TokenAufloesen(filter)
        Dim m As Match

        ' :LAST_MM / :LAST_YYMM / :LAST_YYYYMM(n) → YYYYMM
        m = Regex.Match(f, ":LAST_(?:MM|YYMM|YYYYMM)\((\d+)\)", RegexOptions.IgnoreCase)
        If m.Success Then Return _refDatum.AddMonths(-CInt(m.Groups(1).Value)).ToString("yyyyMM")

        ' :LAST_YY / :LAST_YYYY(n) → YYYY
        m = Regex.Match(f, ":LAST_(?:YY|YYYY)\((\d+)\)", RegexOptions.IgnoreCase)
        If m.Success Then Return _refDatum.AddYears(-CInt(m.Groups(1).Value)).ToString("yyyy")

        ' :YYYYMM(val,...) → wert direkt
        m = Regex.Match(f, ":YYYYMM\(([^)]+)\)", RegexOptions.IgnoreCase)
        If m.Success Then Return m.Groups(1).Value.Trim()

        ' :YYYY(val,...)
        m = Regex.Match(f, ":YYYY\(([^)]+)\)", RegexOptions.IgnoreCase)
        If m.Success Then Return m.Groups(1).Value.Trim()

        ' Nach Token-Aufloesung: ALLE Ziffern aus dem String extrahieren
        ' v2: Erfasst alle aufeinanderfolgenden Ziffern nach dem Jahrespraefix
        ' z.B. tf_lstp_bg_bs_bedarfe:20260400 → 20260400 (alle 8 Ziffern)
        ' z.B. tf_zkt_bb_20260301              → 20260301 (alle 8 Ziffern)
        m = Regex.Match(f, "((?:19|20)\d{4}\d*)", RegexOptions.IgnoreCase)
        If m.Success Then Return m.Groups(1).Value

        Return Nothing

    End Function

    ' -----------------------------------------------------------------------
    ' TokenAufloesen - Loest Platzhalter-Token in einer Vorlage auf.
    ' -----------------------------------------------------------------------
    Private Function TokenAufloesen(f As String) As String
        Dim r As String = f
        Dim m As Match

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
