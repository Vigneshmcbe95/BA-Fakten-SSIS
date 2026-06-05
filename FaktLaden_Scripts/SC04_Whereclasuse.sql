Option Explicit On
Option Strict On

Imports System
Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports System.Text.RegularExpressions
Imports Microsoft.SqlServer.Dts.Runtime


<Microsoft.SqlServer.Dts.Tasks.ScriptTask.SSISScriptTaskEntryPointAttribute()>
<CLSCompliant(False)>
Partial Public Class ScriptMain
    Inherits Microsoft.SqlServer.Dts.Tasks.ScriptTask.VSTARTScriptObjectModelBase

    Private Const SKRIPT_NAME As String = "SCR04_WhereKlausel_Aufbauen"
    Private Const CONN_NAME As String = "Verbindung"
    Private Const MAX_VERSUCHE As Integer = 10
    Private Const WARTE_SEK As Integer = 30

    Private ReadOnly _refDatum As DateTime = New Date(2026, 3, 1)
    Private _stlTabelle As String = String.Empty
    Private _parameterDB As String = String.Empty
    Private _parametertab As String = String.Empty

    ' =========================================================================
    Public Sub Main()

        Log("ââââââââââââââââââââââââââââââââââââââââââââââââââââââââ")
        Log("SCR04_WhereKlausel_Aufbauen â Start (v2: vollstaendige Ziffernerfassung)")
        Log("Referenzdatum: " & _refDatum.ToString("dd.MM.yyyy HH:mm:ss"))
        Log("ââââââââââââââââââââââââââââââââââââââââââââââââââââââââ")

        Try
            _stlTabelle = Dts.Variables("BA::SteuerlistenTabelle").Value.ToString().Trim()
            _parameterDB = Dts.Variables("BA::ParameterDB").Value.ToString().Trim()
            _parametertab = Dts.Variables("BA::Parametertabelle").Value.ToString().Trim()
            Log("Steuerlisten-Tabelle : dbo." & _stlTabelle)
            Log("Parametertabelle     : " & _parameterDB & ".dbo." & _parametertab)

            Dim connStr As String = HoleVerbindungszeichenfolge()

            SpaltenSicherstellen(connStr)

            ' ââ Partitionsspalte direkt aus Parametertabelle (vom Benutzer gepflegt)
            Log("ââ Partitionsspalte aus Parametertabelle laden ...")
            Dim partLookup As Dictionary(Of String, String) = PartitionsspalteAusParameterLaden(connStr)
            Log("Partitionsspalten geladen: " & partLookup.Count.ToString() & " Verfahren")

            Dim zeilen As List(Of Zeile) = ZeilenLaden(connStr)
            Log("Steuerlisten-Zeilen : " & zeilen.Count.ToString())

            Dim cntMit As Integer = 0
            Dim cntOhne As Integer = 0
            Dim cntNull As Integer = 0

            For Each z As Zeile In zeilen

                Dim partCol As String = String.Empty
                partLookup.TryGetValue(z.Tabelle.ToLower().Trim(), partCol)

                Dim wert As String = WertExtrahieren(z.TabnameFilter)

                Dim where As String = Nothing
                If Not String.IsNullOrEmpty(partCol) AndAlso wert IsNot Nothing Then
                    where = "WHERE " & partCol & " = '" & wert & "'"
                    cntMit += 1
                ElseIf wert IsNot Nothing AndAlso String.IsNullOrEmpty(partCol) Then
                    cntNull += 1
                Else
                    cntOhne += 1
                End If

                Log(String.Format("  {0,-50} -> wert={1} | partCol={2}",
                    z.TabnameFilter,
                    If(wert IsNot Nothing, wert, "NULL"),
                    If(String.IsNullOrEmpty(partCol), "NULL", partCol)))

                ZurueckSchreiben(connStr, z.TabnameFilter, z.FileName, where, wert)
            Next

            Log("ââââââââââââââââââââââââââââââââââââââââââââââââââââââââ")
            Log("where_klausel gefuellt       : " & cntMit.ToString())
            Log("partition_wert NULL          : " & cntOhne.ToString())
            Log("kein Partitionsspalte-Eintrag: " & cntNull.ToString())
            Log("ââââââââââââââââââââââââââââââââââââââââââââââââââââââââ")
            Dts.TaskResult = ScriptResults.Success

        Catch ex As Exception
            LogFehler("Kritischer Fehler: " & ex.Message)
            Dts.TaskResult = ScriptResults.Failure
        End Try

    End Sub

    ' =========================================================================
    ' KERN: Wert aus tabname_filter extrahieren per Regex
    '  1. :MONID/:YEAR Token aufloesen
    '  2. :LAST_MM(n)  :LAST_YY(n)  :YYYYMM(...)  :YYYY(...)
    '  3. Nach Token-Aufloesung: ALLE Ziffern aus aufgeloestem String extrahieren
    '  4. Kein Treffer â Nothing (NULL in DB)
    '
    ' AENDERUNG v2:
    '  Fallback-Regex geaendert von ((?:19|20)\d{4})\d* auf ((?:19|20)\d{4}\d*)
    '  Damit werden alle Ziffern erfasst, z.B.:
    '    tf_lstp_bg_bs_bedarfe:MONID(0)00 â nach Aufloesung: tf_lstp_bg_bs_bedarfe:20260400
    '    Alt: 202604   (nur 6 Ziffern, trailing 00 ignoriert)
    '    Neu: 20260400 (alle 8 Ziffern, stimmt mit Oracle mow_id ueberein)
    ' =========================================================================
    Private Function WertExtrahieren(filter As String) As String

        If String.IsNullOrEmpty(filter) Then Return Nothing

        Dim f As String = TokenAufloesen(filter)
        Dim m As Match

        ' :LAST_MM / :LAST_YYMM / :LAST_YYYYMM(n) â YYYYMM
        m = Regex.Match(f, ":LAST_(?:MM|YYMM|YYYYMM)\((\d+)\)", RegexOptions.IgnoreCase)
        If m.Success Then Return _refDatum.AddMonths(-CInt(m.Groups(1).Value)).ToString("yyyyMM")

        ' :LAST_YY / :LAST_YYYY(n) â YYYY
        m = Regex.Match(f, ":LAST_(?:YY|YYYY)\((\d+)\)", RegexOptions.IgnoreCase)
        If m.Success Then Return _refDatum.AddYears(-CInt(m.Groups(1).Value)).ToString("yyyy")

        ' :YYYYMM(val,...) â wert direkt
        m = Regex.Match(f, ":YYYYMM\(([^)]+)\)", RegexOptions.IgnoreCase)
        If m.Success Then Return m.Groups(1).Value.Trim()

        ' :YYYY(val,...)
        m = Regex.Match(f, ":YYYY\(([^)]+)\)", RegexOptions.IgnoreCase)
        If m.Success Then Return m.Groups(1).Value.Trim()

        ' Nach Token-Aufloesung: ALLE Ziffern aus dem String extrahieren
        ' v2: Erfasst alle aufeinanderfolgenden Ziffern nach dem Jahrespraefix
        ' z.B. tf_lstp_bg_bs_bedarfe:20260400 â 20260400 (alle 8 Ziffern)
        ' z.B. tf_zkt_bb_20260301              â 20260301 (alle 8 Ziffern)
        m = Regex.Match(f, "((?:19|20)\d{4}\d*)", RegexOptions.IgnoreCase)
        If m.Success Then Return m.Groups(1).Value

        Return Nothing

    End Function

    ' =========================================================================
    ' Partitionsspalte EINMALIG aus Parametertabelle laden
    ' Dictionary: Verfahren (lower) â Faktenpartitionsspalte
    ' Direkt vom Benutzer gepflegter Wert â kein vm_ddl benoetigt
    ' =========================================================================
    Private Function PartitionsspalteAusParameterLaden(connStr As String) As Dictionary(Of String, String)
        Dim dict As New Dictionary(Of String, String)()
        Dim sql As String =
            "SELECT LOWER(LTRIM(RTRIM(Verfahren))), LTRIM(RTRIM(Wert)) " &
            "FROM   " & _parameterDB & ".dbo." & _parametertab & " " &
            "WHERE  LOWER(LTRIM(RTRIM(Parameter))) = 'faktenpartitionsspalte' " &
            "AND    Wert IS NOT NULL AND LTRIM(RTRIM(Wert)) <> ''"

        Dim versuch As Integer = 0
        Dim letzterFehler As Exception = Nothing
        While versuch < MAX_VERSUCHE
            versuch += 1
            Try
                Using conn As New SqlConnection(connStr)
                    conn.Open()
                    Using cmd As New SqlCommand(sql, conn)
                        cmd.CommandTimeout = 0
                        Using rdr As SqlDataReader = cmd.ExecuteReader()
                            While rdr.Read()
                                Dim key As String = rdr(0).ToString().Trim().ToLower()
                                Dim val As String = rdr(1).ToString().Trim()
                                If Not dict.ContainsKey(key) Then dict(key) = val
                            End While
                        End Using
                    End Using
                End Using
                Return dict
            Catch ex As Exception
                letzterFehler = ex
                Log(String.Format("WARNUNG [PartitionsspalteAusParameterLaden] Versuch {0}/{1}: {2}",
                    versuch, MAX_VERSUCHE, ex.Message))
                If versuch < MAX_VERSUCHE Then System.Threading.Thread.Sleep(WARTE_SEK * 1000)
            End Try
        End While
        Throw New Exception(String.Format(
            "[PartitionsspalteAusParameterLaden] fehlgeschlagen nach {0} Versuchen: {1}",
            MAX_VERSUCHE,
            If(letzterFehler IsNot Nothing, letzterFehler.Message, "Unbekannt")))
    End Function

    ' =========================================================================
    ' :MONID(n) â YYYYMM  |  :YEAR(n) â YYYY
    ' =========================================================================
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

    ' =========================================================================
    Private Sub SpaltenSicherstellen(connStr As String)
        SqlRun(connStr,
"IF NOT EXISTS(SELECT 1 FROM sys.columns WHERE object_id=OBJECT_ID('dbo." & _stlTabelle & "') AND name='where_klausel')
    ALTER TABLE dbo." & _stlTabelle & " ADD where_klausel  NVARCHAR(1000) NULL;
IF NOT EXISTS(SELECT 1 FROM sys.columns WHERE object_id=OBJECT_ID('dbo." & _stlTabelle & "') AND name='partition_wert')
    ALTER TABLE dbo." & _stlTabelle & " ADD partition_wert NVARCHAR(500)  NULL;
IF NOT EXISTS(SELECT 1 FROM sys.columns WHERE object_id=OBJECT_ID('dbo." & _stlTabelle & "') AND name='ref_datum')
    ALTER TABLE dbo." & _stlTabelle & " ADD ref_datum      DATETIME       NULL;")
        Log("Spalten geprueft/angelegt OK")
    End Sub

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
                        " SET where_klausel=@w, partition_wert=@p, ref_datum=@r" &
                        " WHERE tabname_filter=@z AND FILE_NAME=@f", conn)
                        cmd.CommandTimeout = 0
                        cmd.Parameters.AddWithValue("@w", If(whereKlausel Is Nothing, CObj(DBNull.Value), CObj(whereKlausel)))
                        cmd.Parameters.AddWithValue("@p", If(partWert Is Nothing, CObj(DBNull.Value), CObj(partWert)))
                        cmd.Parameters.AddWithValue("@r", _refDatum)
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

    Private Sub SqlRun(connStr As String, sql As String)
        Using conn As New SqlConnection(connStr)
            conn.Open()
            Using cmd As New SqlCommand(sql, conn)
                cmd.CommandTimeout = 0
                cmd.ExecuteNonQuery()
            End Using
        End Using
    End Sub

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

    Private Class Zeile
        Public Property TabnameFilter As String
        Public Property Tabelle As String
        Public Property FileName As String
    End Class

    Public Enum ScriptResults
        Success = DTSExecResult.Success
        Failure = DTSExecResult.Failure
    End Enum

End Class
