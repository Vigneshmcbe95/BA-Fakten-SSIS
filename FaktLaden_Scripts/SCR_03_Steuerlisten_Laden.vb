Option Explicit On
Option Strict On

Imports System
Imports System.IO
Imports System.Data.SqlClient
Imports System.Text.RegularExpressions
Imports Microsoft.SqlServer.Dts.Runtime

' =============================================================================
'  Skript       : SCR_03_Steuerlisten_Laden
'  Paket        : Fakten Laden (SSIS)
'  Zweck        : Laedt die Steuerliste in zwei Tabellen:
'                 1. Audit-Tabelle (<Tabelle>_audit): nur INSERT, nie UPDATE
'                    oder DELETE - vollstaendige Historie (wer, wann, was).
'                 2. Arbeitstabelle (<Tabelle>): DELETE (nur aktuelle Datei)
'                    + INSERT je Datei - die Datei ist die Wahrheit; SC04
'                    fuellt hier where_klausel / partition_wert.
'                 Danach wird die Datei in den Ordner 'Done' verschoben;
'                 Dateien mit Endung _perm.csv bleiben liegen.
'                 Jeder Fehler (Datei fehlt / nicht lesbar / SQL) fuehrt zum
'                 Abbruch des Tasks.
'  Wiederholung : 3 Versuche je SQL-Anweisung, 30 s Wartezeit
'  Protokoll    : Nur SSIS-Events (FireInformation / FireError)
' =============================================================================
<Microsoft.SqlServer.Dts.Tasks.ScriptTask.SSISScriptTaskEntryPointAttribute()>
<CLSCompliant(False)>
Partial Public Class ScriptMain
    Inherits Microsoft.SqlServer.Dts.Tasks.ScriptTask.VSTARTScriptObjectModelBase

    ' -------------------------------------------------------------------------
    ' Konstanten
    ' -------------------------------------------------------------------------
    Private Const SKRIPT_NAME As String = "SCR_03_Steuerlisten_Laden"
    Private Const CONN_NAME As String = "Verbindung"
    Private Const MAX_VERSUCHE As Integer = 3
    Private Const WARTE_SEK As Integer = 30
    Private Const DONE_ORDNER As String = "Done"
    Private Const PERM_ENDUNG As String = "_perm.csv"

    ' -------------------------------------------------------------------------
    ' SSIS-Variablen
    ' -------------------------------------------------------------------------
    Private _stlOrdner As String = String.Empty
    Private _stlDateiname As String = String.Empty
    Private _steuerlistenTabelle As String = String.Empty
    Private _auditTabelle As String = String.Empty
    Private _bearbeiter As String = String.Empty

    Private _geladeneTabellen As New List(Of String)()

    ' -----------------------------------------------------------------------
    ' Main - Einstiegspunkt - steuert den Ablauf des Skripts.
    ' -----------------------------------------------------------------------
    Public Sub Main()

        Log("SCR_03_Steuerlisten_Laden - Start")

        Try
            VariablenLaden()

            If Not PflichtfelderPruefen() Then
                Dts.TaskResult = ScriptResults.Failure
                Return
            End If

            If Not _stlOrdner.EndsWith("\") Then _stlOrdner &= "\"

            ' Zentrale Insert-Only-Audit-Tabelle fuer FAKTEN und DIMENSIONEN
            ' (beide STL-Loader schreiben hier hinein; nur INSERT, kein DELETE/UPDATE).
            _auditTabelle = "tm_steuerliste_audit"
            Log("Arbeitstabelle : dbo." & _steuerlistenTabelle)
            Log("Audit-Tabelle  : dbo." & _auditTabelle)

            Dim connStr As String = HoleVerbindungszeichenfolge()
            TabellenSicherstellen(connStr)

            Dim dateipfad As String = Path.Combine(_stlOrdner, _stlDateiname)
            Log("Vollstaendiger Pfad : " & dateipfad)

            If Not File.Exists(dateipfad) Then
                LogFehler("Datei nicht gefunden: " & dateipfad)
                Dts.TaskResult = ScriptResults.Failure
                Return
            End If

            Dim dateiname As String = Path.GetFileName(dateipfad)
            Log("Verarbeite Datei: " & dateiname)

            VerarbeiteDatei(dateipfad, dateiname, connStr)
            Log("Datei erfolgreich verarbeitet: " & dateiname & " OK")

            Log("Geladene Tabellen aus [" & dateiname & "] (" &
                _geladeneTabellen.Count.ToString() & "): " & String.Join(", ", _geladeneTabellen))
            If _geladeneTabellen.Count = 0 Then
                LogFehler("STL-Datei [" & dateiname & "] enthielt keine verwertbaren Tabellenzeilen - Abbruch.")
                Dts.TaskResult = ScriptResults.Failure
                Return
            End If

            DateiVerschieben(dateipfad, dateiname)

            Log("SCR_03_Steuerlisten_Laden erfolgreich abgeschlossen OK")
            Dts.TaskResult = ScriptResults.Success

        Catch ex As Exception
            LogFehler("Kritischer Fehler in " & SKRIPT_NAME & ": " & ex.Message)
            Dts.TaskResult = ScriptResults.Failure
        End Try

    End Sub

    ' -----------------------------------------------------------------------
    ' VariablenLaden
    ' -----------------------------------------------------------------------
    Private Sub VariablenLaden()
        _stlOrdner = Dts.Variables("BA::STLOrdner").Value.ToString().Trim()
        _stlDateiname = Dts.Variables("BA::STLDateiname").Value.ToString().Trim()
        _steuerlistenTabelle = Dts.Variables("BA::SteuerlistenTabelle").Value.ToString().Trim()
        _bearbeiter = Dts.Variables("System::CreatorName").Value.ToString().Trim()
    End Sub

    ' -----------------------------------------------------------------------
    ' PflichtfelderPruefen
    ' -----------------------------------------------------------------------
    Private Function PflichtfelderPruefen() As Boolean
        Dim fehlend As New System.Text.StringBuilder()
        If String.IsNullOrEmpty(_stlOrdner) Then fehlend.AppendLine("  - BA::STLOrdner")
        If String.IsNullOrEmpty(_stlDateiname) Then fehlend.AppendLine("  - BA::STLDateiname")
        If String.IsNullOrEmpty(_steuerlistenTabelle) Then fehlend.AppendLine("  - BA::SteuerlistenTabelle")
        If fehlend.Length > 0 Then
            LogFehler("Pflichtfelder fehlen:" & Environment.NewLine & fehlend.ToString())
            Return False
        End If
        Log("Pflichtfelder-Pruefung: alle Variablen vorhanden OK")
        Return True
    End Function

    ' -----------------------------------------------------------------------
    ' TabellenSicherstellen
    ' -----------------------------------------------------------------------
    Private Sub TabellenSicherstellen(connStr As String)

        Dim w As String = _steuerlistenTabelle
        Dim a As String = _auditTabelle

        Dim sql As String =
"IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = '" & w & "' AND schema_id = SCHEMA_ID('dbo'))
BEGIN
    CREATE TABLE dbo." & w & "
    (
        tabelle        NVARCHAR(255)  NULL,
        FILE_NAME      NVARCHAR(255)  NULL,
        tabellentyp    NVARCHAR(255)  NULL,
        umgebung       NVARCHAR(255)  NULL,
        load_Date      DATETIME       NULL,
        themengebiet   NVARCHAR(255)  NULL,
        bearbeiter     NVARCHAR(255)  NULL,
        tabname_filter NVARCHAR(500)  NULL,
        where_klausel  NVARCHAR(1000) NULL,
        partition_wert NVARCHAR(500)  NULL
    );
END;

IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = '" & a & "' AND schema_id = SCHEMA_ID('dbo'))
BEGIN
    CREATE TABLE dbo." & a & "
    (
        stlid          INT IDENTITY(1,1) NOT NULL
            CONSTRAINT PK_" & a & " PRIMARY KEY,
        bereich        NVARCHAR(10)   NULL,
        tabelle        NVARCHAR(255)  NULL,
        FILE_NAME      NVARCHAR(255)  NULL,
        tabellentyp    NVARCHAR(255)  NULL,
        umgebung       NVARCHAR(255)  NULL,
        load_Date      DATETIME       NULL,
        themengebiet   NVARCHAR(255)  NULL,
        bearbeiter     NVARCHAR(255)  NULL,
        tabname_filter NVARCHAR(500)  NULL
    );
END;"

        SqlAusfuehren(connStr, sql, "Tabellen sicherstellen")
        Log("Arbeitstabelle und Audit-Tabelle geprueft/angelegt OK")

    End Sub

    ' -----------------------------------------------------------------------
    ' VerarbeiteDatei
    '   Audit-Tabelle  : INSERT only – tabelle = parsed name,
    '                                  tabname_filter = raw line
    '   Arbeitstabelle : DELETE by FILE_NAME, then INSERT
    ' -----------------------------------------------------------------------
    Private Sub VerarbeiteDatei(dateipfad As String,
                                dateiname As String,
                                connStr As String)

        Dim teile() As String =
            Path.GetFileNameWithoutExtension(dateiname).Split("."c)

        Dim tabellentyp As String = If(teile.Length > 0, teile(0).ToLower().Trim(), "")
        Dim umgebung As String = If(teile.Length > 1, teile(1).ToLower().Trim(), "")
        Dim themengebiet As String = If(teile.Length > 2, teile(2).ToLower().Trim(), "")

        ' "_perm" ist nur eine Datei-Konvention (Datei bleibt im Quellordner) -
        ' das fachliche Themengebiet in Oracle traegt den Suffix nicht.
        If themengebiet.EndsWith("_perm") Then
            themengebiet = themengebiet.Substring(0, themengebiet.Length - "_perm".Length)
            Log("  Themengebiet : '_perm'-Suffix entfernt (Datei-Konvention)")
        End If

        Log("  Tabellentyp  : " & tabellentyp)
        Log("  Umgebung     : " & umgebung)
        Log("  Themengebiet : " & themengebiet)

        ' Arbeitstabelle: komplett leeren - spiegelt ausschliesslich den aktuellen Lauf
        SqlAusfuehren(connStr,
            "DELETE FROM dbo." & _steuerlistenTabelle,
            "DELETE Arbeitstabelle")
        Log("  Arbeitstabelle: vollstaendig geleert")

        Dim ladeZeit As DateTime = DateTime.Now
        Dim zeilenNr As Integer = 0
        Dim cntInsert As Integer = 0
        Dim cntUebersp As Integer = 0

        For Each rohzeile As String In File.ReadAllLines(dateipfad)
            zeilenNr += 1

            Dim zeile As String = rohzeile _
                .Replace(Chr(13), "") _
                .Replace(Chr(10), "") _
                .Replace(Chr(9), "") _
                .Trim() _
                .TrimEnd(";"c) _
                .Trim()

            If String.IsNullOrEmpty(zeile) OrElse zeile.StartsWith("#") Then
                cntUebersp += 1
                Continue For
            End If

            Dim tabellenname As String = TabellennameBestimmen(zeile)

            Try
                ' 1. Audit-Tabelle: tabelle = parsed name, tabname_filter = raw line
                SqlMitParameternAusfuehren(connStr,
                    "INSERT INTO dbo." & _auditTabelle &
                    " (bereich, tabelle, FILE_NAME, tabellentyp, umgebung," &
                    "  load_Date, themengebiet, bearbeiter, tabname_filter)" &
                    " VALUES ('FAKT', @tab, @f, @typ, @umb, @dat, @thm, @bea, @filter)",
                    "INSERT Audit",
                    New With {
                        .tab = tabellenname,
                        .f = dateiname,
                        .typ = tabellentyp,
                        .umb = umgebung,
                        .dat = ladeZeit,
                        .thm = themengebiet,
                        .bea = _bearbeiter,
                        .filter = zeile
                    })

                ' 2. Arbeitstabelle: tabelle = parsed name, tabname_filter = raw line
                SqlMitParameternAusfuehren(connStr,
                    "INSERT INTO dbo." & _steuerlistenTabelle &
                    " (tabelle, FILE_NAME, tabellentyp, umgebung," &
                    "  load_Date, themengebiet, bearbeiter, tabname_filter)" &
                    " VALUES (@tab, @f, @typ, @umb, @dat, @thm, @bea, @filter)",
                    "INSERT Arbeitstabelle",
                    New With {
                        .tab = tabellenname,
                        .f = dateiname,
                        .typ = tabellentyp,
                        .umb = umgebung,
                        .dat = ladeZeit,
                        .thm = themengebiet,
                        .bea = _bearbeiter,
                        .filter = zeile
                    })

                cntInsert += 1
                If Not _geladeneTabellen.Contains(tabellenname) Then
                    _geladeneTabellen.Add(tabellenname)
                End If

            Catch ex As Exception
                Throw New Exception(
                    String.Format("Zeile {0} ('{1}'): {2}", zeilenNr, zeile, ex.Message), ex)
            End Try
        Next

        Log(String.Format("  Zeilen gesamt: {0} | INSERT: {1} | Uebersprungen: {2}",
            zeilenNr, cntInsert, cntUebersp))

    End Sub

    ' -----------------------------------------------------------------------
    ' DateiVerschieben
    ' -----------------------------------------------------------------------
    Private Sub DateiVerschieben(dateipfad As String, dateiname As String)

        If dateiname.ToLower().EndsWith(PERM_ENDUNG) Then
            Log("Datei endet auf " & PERM_ENDUNG & " -> bleibt im Quellordner")
            Return
        End If

        Dim doneOrdner As String = Path.Combine(_stlOrdner, DONE_ORDNER)
        If Not Directory.Exists(doneOrdner) Then
            Directory.CreateDirectory(doneOrdner)
            Log("Done-Ordner angelegt: " & doneOrdner)
        End If

        Dim zielPfad As String = Path.Combine(doneOrdner, dateiname)
        If File.Exists(zielPfad) Then
            zielPfad = Path.Combine(doneOrdner,
                Path.GetFileNameWithoutExtension(dateiname) &
                "_" & DateTime.Now.ToString("yyyyMMdd_HHmmss") &
                Path.GetExtension(dateiname))
        End If

        File.Move(dateipfad, zielPfad)
        Log("Datei verschoben nach: " & zielPfad)

    End Sub

    ' -----------------------------------------------------------------------
    ' TabellennameBestimmen - Ermittelt den reinen Tabellen-/Verfahrensnamen
    ' aus einer STL-Zeile und validiert das Format.
    ' Unterstuetzte Schreibweisen:
    '   tabelle                 - einfacher Name
    '   tabelle_*               - Wildcard: alle Partitionen (AUTOMATIC)
    '   tabelle_<JJJJMM[..]>    - fester Datums-Suffix
    '   tabelle:<JJJJMM[..]>    - fester Datumswert nach Doppelpunkt
    '   tabelle:<TOKEN>(...)    - Token (:YYYYMM(), :LAST_MM(), :MONID..., :YEAR)
    ' Bleibt nach dem Parsen ein Sonderzeichen uebrig (unbekanntes Format),
    ' wird mit einer Hinweismeldung abgebrochen.
    ' -----------------------------------------------------------------------
    Private Function TabellennameBestimmen(zeile As String) As String

        Dim basis As String = RohenTabellennamenExtrahieren(zeile)
        basis = basis.ToLower().Trim().TrimEnd("_"c)

        ' Ein gueltiger Tabellen-/Verfahrensname besteht nur aus Buchstaben,
        ' Ziffern und Unterstrichen. Sonst ist das Format unbekannt -> Abbruch.
        If basis = "" OrElse Not Regex.IsMatch(basis, "^[a-z0-9_]+$") Then
            Throw New Exception(
                "Tabellenname '" & zeile & "' hat ein unbekanntes Format. " &
                "Erlaubte Schreibweisen: 'tabelle', 'tabelle_*', " &
                "'tabelle_<JJJJMM>' oder 'tabelle:<TOKEN>(...)' " &
                "(z.B. :YYYYMM(), :LAST_MM(), :MONID6(), :YEAR()). " &
                "Bitte den Tabellennamen in der STL-Datei entsprechend angeben.")
        End If

        Return basis

    End Function

    ' -----------------------------------------------------------------------
    ' RohenTabellennamenExtrahieren - Schneidet Token / Datums-Suffix /
    ' Wildcard ab und liefert den Basisnamen (noch ohne Validierung).
    ' -----------------------------------------------------------------------
    Private Function RohenTabellennamenExtrahieren(zeile As String) As String

        ' :YYYYMM( / :YYYY( / :YY( / :MM( etc.
        Dim t As Match = Regex.Match(zeile, ":[YM]{2,6}\(", RegexOptions.IgnoreCase)
        If Not t.Success Then
            ' :LAST_YYYYMM( / :LAST_YYYY( / :LAST_MM( etc.
            t = Regex.Match(zeile, ":LAST_[YM]{2,6}\(", RegexOptions.IgnoreCase)
        End If
        If t.Success Then Return zeile.Substring(0, t.Index)

        ' bare date digits after colon: tf_table:20260400 or tf_table:202604
        t = Regex.Match(zeile, ":((?:19|20)\d{4,})", RegexOptions.IgnoreCase)
        If t.Success Then Return zeile.Substring(0, t.Index)

        ' date suffix embedded in table name: tf_table_20240606
        t = Regex.Match(zeile, "_((?:19|20)\d{4,})$", RegexOptions.IgnoreCase)
        If t.Success Then Return zeile.Substring(0, t.Index)

        ' :MONID token (with or without leading underscore)
        t = Regex.Match(zeile, "_?:MONID", RegexOptions.IgnoreCase)
        If t.Success Then Return zeile.Substring(0, t.Index)

        ' :YEAR token (with or without leading underscore)
        t = Regex.Match(zeile, "_?:YEAR", RegexOptions.IgnoreCase)
        If t.Success Then Return zeile.Substring(0, t.Index)

        ' Wildcard-Suffix: tf_table_*  -> alle Partitionen (AUTOMATIC)
        t = Regex.Match(zeile, "_?\*+$")
        If t.Success Then Return zeile.Substring(0, t.Index)

        ' einfacher Tabellenname
        Return zeile

    End Function

    ' -----------------------------------------------------------------------
    ' SqlAusfuehren
    ' -----------------------------------------------------------------------
    Private Function SqlAusfuehren(connStr As String,
                                   sql As String,
                                   beschreibung As String) As Integer
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
                Log(String.Format("WARNUNG [{0}] Versuch {1}/{2}: {3}",
                    beschreibung, versuch, MAX_VERSUCHE, ex.Message))
                If versuch = 1 Then Log("SQL Statement [" & beschreibung & "]: " & sql)
                If versuch < MAX_VERSUCHE Then
                    System.Threading.Thread.Sleep(WARTE_SEK * 1000)
                End If
            End Try
        End While

        Throw New Exception(String.Format(
            "[{0}] fehlgeschlagen nach {1} Versuchen: {2}",
            beschreibung, MAX_VERSUCHE,
            If(letzterFehler IsNot Nothing, letzterFehler.Message, "Unbekannt")))
    End Function

    ' -----------------------------------------------------------------------
    ' SqlMitParameternAusfuehren
    ' -----------------------------------------------------------------------
    Private Sub SqlMitParameternAusfuehren(connStr As String,
                                           sql As String,
                                           beschreibung As String,
                                           params As Object)
        Dim versuch As Integer = 0
        Dim letzterFehler As Exception = Nothing

        While versuch < MAX_VERSUCHE
            versuch += 1
            Try
                Using conn As New SqlConnection(connStr)
                    conn.Open()
                    Using cmd As New SqlCommand(sql, conn)
                        cmd.CommandTimeout = 0
                        ParameterHinzufuegen(cmd, params)
                        cmd.ExecuteNonQuery()
                        Return
                    End Using
                End Using
            Catch ex As Exception
                letzterFehler = ex
                Log(String.Format("WARNUNG [{0}] Versuch {1}/{2}: {3}",
                    beschreibung, versuch, MAX_VERSUCHE, ex.Message))
                If versuch = 1 Then Log("SQL Statement [" & beschreibung & "]: " & sql)
                If versuch < MAX_VERSUCHE Then
                    System.Threading.Thread.Sleep(WARTE_SEK * 1000)
                End If
            End Try
        End While

        Throw New Exception(String.Format(
            "[{0}] fehlgeschlagen nach {1} Versuchen: {2}",
            beschreibung, MAX_VERSUCHE,
            If(letzterFehler IsNot Nothing, letzterFehler.Message, "Unbekannt")))
    End Sub

    ' -----------------------------------------------------------------------
    ' ParameterHinzufuegen
    ' -----------------------------------------------------------------------
    Private Sub ParameterHinzufuegen(cmd As SqlCommand, params As Object)
        If params Is Nothing Then Return
        For Each prop As System.Reflection.PropertyInfo In params.GetType().GetProperties()
            Dim wert As Object = prop.GetValue(params, Nothing)
            cmd.Parameters.AddWithValue(
                "@" & prop.Name,
                If(wert Is Nothing, CObj(DBNull.Value), wert))
        Next
    End Sub

    ' -----------------------------------------------------------------------
    ' HoleVerbindungszeichenfolge
    ' -----------------------------------------------------------------------
    Private Function HoleVerbindungszeichenfolge() As String
        Return Dts.Connections(CONN_NAME).ConnectionString
    End Function

    ' -----------------------------------------------------------------------
    ' Log
    ' -----------------------------------------------------------------------
    Private Sub Log(nachricht As String)
        Dim fireAgain As Boolean = False
        Dts.Events.FireInformation(0, SKRIPT_NAME, nachricht, "", 0, fireAgain)
    End Sub

    ' -----------------------------------------------------------------------
    ' LogFehler
    ' -----------------------------------------------------------------------
    Private Sub LogFehler(nachricht As String)
        Dts.Events.FireError(0, SKRIPT_NAME, nachricht, "", 0)
    End Sub

    ' -----------------------------------------------------------------------
    ' ScriptResults
    ' -----------------------------------------------------------------------
    Public Enum ScriptResults
        Success = DTSExecResult.Success
        Failure = DTSExecResult.Failure
    End Enum

End Class
