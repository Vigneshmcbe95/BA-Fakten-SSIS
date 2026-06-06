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
'                 2. Arbeitstabelle (<Tabelle>): DELETE + INSERT je Datei -
'                    die Datei ist die Wahrheit; SC04 fuellt hier
'                    where_klausel / partition_wert.
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

            ' Ordnerpfad normalisieren
            If Not _stlOrdner.EndsWith("\") Then _stlOrdner &= "\"

            _auditTabelle = _steuerlistenTabelle & "_audit"
            Log("Arbeitstabelle : dbo." & _steuerlistenTabelle)
            Log("Audit-Tabelle  : dbo." & _auditTabelle)

            ' Tabellen sicherstellen (Arbeitstabelle + Audit)
            Dim connStr As String = HoleVerbindungszeichenfolge()
            TabellenSicherstellen(connStr)

            ' Exakte Datei aus BA::STLDateiname - fehlt sie, ist das ein Fehler
            Dim dateipfad As String = Path.Combine(_stlOrdner, _stlDateiname)
            Log("Vollstaendiger Pfad : " & dateipfad)

            If Not File.Exists(dateipfad) Then
                LogFehler("Datei nicht gefunden: " & dateipfad)
                Dts.TaskResult = ScriptResults.Failure
                Return
            End If

            ' Datei verarbeiten - jeder Fehler fuehrt zum Abbruch
            Dim dateiname As String = Path.GetFileName(dateipfad)
            Log("Verarbeite Datei: " & dateiname)

            VerarbeiteDatei(dateipfad, dateiname, connStr)
            Log("Datei erfolgreich verarbeitet: " & dateiname & " OK")

            ' Datei nach Done verschieben (ausser *_perm.csv)
            DateiVerschieben(dateipfad, dateiname)

            Log("SCR_03_Steuerlisten_Laden erfolgreich abgeschlossen OK")
            Dts.TaskResult = ScriptResults.Success

        Catch ex As Exception
            LogFehler("Kritischer Fehler in " & SKRIPT_NAME & ": " & ex.Message)
            Dts.TaskResult = ScriptResults.Failure
        End Try

    End Sub

    ' -----------------------------------------------------------------------
    ' VariablenLaden - Liest die benoetigten SSIS-Variablen in Modulfelder
    ' ein.
    ' -----------------------------------------------------------------------
    Private Sub VariablenLaden()
        _stlOrdner = Dts.Variables("BA::STLOrdner").Value.ToString().Trim()
        _stlDateiname = Dts.Variables("BA::STLDateiname").Value.ToString().Trim()
        _steuerlistenTabelle = Dts.Variables("BA::SteuerlistenTabelle").Value.ToString().Trim()
        _bearbeiter = Dts.Variables("System::CreatorName").Value.ToString().Trim()
    End Sub

    ' -----------------------------------------------------------------------
    ' PflichtfelderPruefen - Prueft, ob alle Pflichtvariablen / -parameter
    ' vorhanden sind.
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
    ' TabellenSicherstellen - Stellt Arbeitstabelle und Audit-Tabelle sicher
    ' (CREATE IF NOT EXISTS).
    ' Arbeitstabelle: technische Tabelle, SC04 fuellt where_klausel /
    '                 partition_wert.
    ' Audit-Tabelle : stlid IDENTITY PRIMARY KEY, nur INSERT.
    ' -----------------------------------------------------------------------
    Private Sub TabellenSicherstellen(connStr As String)

        Dim w As String = _steuerlistenTabelle
        Dim a As String = _auditTabelle

        Dim sql As String =
"-- =====================================================================
-- 1. Arbeitstabelle sicherstellen
-- =====================================================================
IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = '" & w & "' AND schema_id = SCHEMA_ID('dbo'))
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

-- =====================================================================
-- 2. Audit-Tabelle sicherstellen (nur INSERT, nie UPDATE / DELETE)
-- =====================================================================
IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = '" & a & "' AND schema_id = SCHEMA_ID('dbo'))
BEGIN
    CREATE TABLE dbo." & a & "
    (
        stlid          INT IDENTITY(1,1) NOT NULL
            CONSTRAINT PK_" & a & " PRIMARY KEY,
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
    ' VerarbeiteDatei - Verarbeitet eine einzelne Steuerlisten-Datei:
    ' 1. INSERT jeder Zeile in die Audit-Tabelle (Historie)
    ' 2. DELETE + INSERT in der Arbeitstabelle je FILE_NAME
    '    (die Datei ist die Wahrheit - entfernte Zeilen verschwinden)
    ' -----------------------------------------------------------------------
    Private Sub VerarbeiteDatei(dateipfad As String,
                                dateiname As String,
                                connStr As String)

        ' Metadaten aus Dateiname ableiten
        ' Format: tabellentyp.umgebung.themengebiet_perm.csv
        Dim teile() As String =
            Path.GetFileNameWithoutExtension(dateiname).Split("."c)

        Dim tabellentyp As String = If(teile.Length > 0, teile(0).ToLower().Trim(), "")
        Dim umgebung As String = If(teile.Length > 1, teile(1).ToLower().Trim(), "")
        Dim themengebiet As String = If(teile.Length > 2, teile(2).ToLower().Trim(), "")

        Log("  Tabellentyp  : " & tabellentyp)
        Log("  Umgebung     : " & umgebung)
        Log("  Themengebiet : " & themengebiet)

        ' Arbeitstabelle: alte Zeilen dieser Datei entfernen (Datei = Wahrheit)
        SqlMitParameternAusfuehren(connStr,
            "DELETE FROM dbo." & _steuerlistenTabelle & " WHERE FILE_NAME = @f",
            "DELETE Arbeitstabelle",
            New With {.f = dateiname})
        Log("  Arbeitstabelle: alte Zeilen der Datei entfernt")

        Dim ladeZeit As DateTime = DateTime.Now
        Dim zeilenNr As Integer = 0
        Dim cntInsert As Integer = 0
        Dim cntUebersp As Integer = 0

        For Each rohzeile As String In File.ReadAllLines(dateipfad)
            zeilenNr += 1

            ' Zeile bereinigen
            Dim zeile As String = rohzeile _
                .Replace(Chr(13), "") _
                .Replace(Chr(10), "") _
                .Replace(Chr(9), "") _
                .Trim() _
                .TrimEnd(";"c) _
                .Trim()

            ' Leerzeilen und Kommentare ueberspringen
            If String.IsNullOrEmpty(zeile) OrElse zeile.StartsWith("#") Then
                cntUebersp += 1
                Continue For
            End If

            ' Tabellenname aus Zeile ableiten
            Dim tabellenname As String = TabellennameBestimmen(zeile)

            Try
                ' 1. Audit-Tabelle: immer INSERT (Historie, nie aendern)
                SqlMitParameternAusfuehren(connStr,
                    "INSERT INTO dbo." & _auditTabelle &
                    " (tabelle, FILE_NAME, tabellentyp, umgebung," &
                    "  load_Date, themengebiet, bearbeiter, tabname_filter)" &
                    " VALUES (@tab, @f, @typ, @umb, @dat, @thm, @bea, @z)",
                    "INSERT Audit",
                    New With {
                        .tab = tabellenname,
                        .f = dateiname,
                        .typ = tabellentyp,
                        .umb = umgebung,
                        .dat = ladeZeit,
                        .thm = themengebiet,
                        .bea = _bearbeiter,
                        .z = zeile
                    })

                ' 2. Arbeitstabelle: INSERT (vorher je Datei geleert)
                SqlMitParameternAusfuehren(connStr,
                    "INSERT INTO dbo." & _steuerlistenTabelle &
                    " (tabelle, FILE_NAME, tabellentyp, umgebung," &
                    "  load_Date, themengebiet, bearbeiter, tabname_filter)" &
                    " VALUES (@tab, @f, @typ, @umb, @dat, @thm, @bea, @z)",
                    "INSERT Arbeitstabelle",
                    New With {
                        .tab = tabellenname,
                        .f = dateiname,
                        .typ = tabellentyp,
                        .umb = umgebung,
                        .dat = ladeZeit,
                        .thm = themengebiet,
                        .bea = _bearbeiter,
                        .z = zeile
                    })

                cntInsert += 1

            Catch ex As Exception
                Throw New Exception(
                    String.Format("Zeile {0} ('{1}'): {2}", zeilenNr, zeile, ex.Message), ex)
            End Try
        Next

        Log(String.Format("  Zeilen gesamt: {0} | INSERT: {1} | Uebersprungen: {2}",
            zeilenNr, cntInsert, cntUebersp))

    End Sub

    ' -----------------------------------------------------------------------
    ' DateiVerschieben - Verschiebt die verarbeitete Datei in den Ordner
    ' 'Done'. Dateien mit Endung _perm.csv bleiben liegen. Bei Namens-
    ' kollision im Done-Ordner wird ein Zeitstempel angehaengt.
    ' Jeder Fehler beim Verschieben fuehrt zum Abbruch.
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
    ' TabellennameBestimmen - Ermittelt den Zieltabellennamen eines
    ' Verfahrens.
    ' -----------------------------------------------------------------------
    Private Function TabellennameBestimmen(zeile As String) As String

        Dim t As Match = Regex.Match(zeile, ":[YM]{2,6}\(", RegexOptions.IgnoreCase)
        If Not t.Success Then
            t = Regex.Match(zeile, ":LAST_[YM]{2,6}\(", RegexOptions.IgnoreCase)
        End If
        If t.Success Then Return zeile.Substring(0, t.Index).ToLower().Trim()

        t = Regex.Match(zeile, "_:MONID", RegexOptions.IgnoreCase)
        If t.Success Then Return zeile.Substring(0, t.Index).ToLower().Trim()

        t = Regex.Match(zeile, "_:YEAR", RegexOptions.IgnoreCase)
        If t.Success Then Return zeile.Substring(0, t.Index).ToLower().Trim()

        Return zeile.ToLower().Trim()

    End Function

    ' -----------------------------------------------------------------------
    ' SqlAusfuehren - Fuehrt eine SQL-Anweisung (Non-Query) mit
    ' Wiederholung aus; protokolliert Warnung und vollstaendiges
    ' SQL-Statement bei Fehlern.
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
    ' SqlMitParameternAusfuehren - Fuehrt eine parametrisierte
    ' SQL-Anweisung aus.
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
    ' ParameterHinzufuegen - Fuegt einem SQL-Command einen Parameter hinzu.
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
    Private Sub Log(nachricht As String)
        Dim fireAgain As Boolean = False
        Dts.Events.FireInformation(0, SKRIPT_NAME, nachricht, "", 0, fireAgain)
    End Sub

    ' -----------------------------------------------------------------------
    ' LogFehler - Schreibt eine Fehlermeldung in das SSIS-Protokoll
    ' (FireError).
    ' -----------------------------------------------------------------------
    Private Sub LogFehler(nachricht As String)
        Dts.Events.FireError(0, SKRIPT_NAME, nachricht, "", 0)
    End Sub

    ' -----------------------------------------------------------------------
    ' ScriptResults - SSIS-Task-Ergebniscodes.
    ' -----------------------------------------------------------------------
    Public Enum ScriptResults
        Success = DTSExecResult.Success
        Failure = DTSExecResult.Failure
    End Enum

End Class
