Option Explicit On
Option Strict On

Imports System
Imports System.IO
Imports System.Data.SqlClient
Imports System.Text.RegularExpressions
Imports Microsoft.SqlServer.Dts.Runtime

' =============================================================================
' PAKET  : Fakten Laden
' SKRIPT : SCR_03_Steuerlisten_Laden
' ZWECK  : 1. Exakte CSV-Datei aus BA::STLDateiname einlesen
'          2. Zeilenweise verarbeiten → Tabellenname ableiten
'          3. INSERT / UPDATE in dbo.tm_steuerlistenfile_Fakten
'          4. Pflichtfelder (where_klausel, partition_wert) sicherstellen
'          5. Zusammenfassung ausgeben
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

    ' -------------------------------------------------------------------------
    ' SSIS-Variablen
    ' -------------------------------------------------------------------------
    Private _stlOrdner As String = String.Empty
    Private _stlDateiname As String = String.Empty
    Private _steuerlistenTabelle As String = String.Empty
    Private _bearbeiter As String = String.Empty

    ' -------------------------------------------------------------------------
    ' Einstiegspunkt
    ' -------------------------------------------------------------------------
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

            Log("Steuerlisten-Tabelle: dbo." & _steuerlistenTabelle)

            ' Spalten sicherstellen
            Dim connStr As String = HoleVerbindungszeichenfolge()
            SpaltenSicherstellen(connStr)

            ' Exakte Datei aus BA::STLDateiname
            Dim dateipfad As String = Path.Combine(_stlOrdner, _stlDateiname)
            Log("Vollstaendiger Pfad : " & dateipfad)

            If Not File.Exists(dateipfad) Then
                LogFehler("Datei nicht gefunden: " & dateipfad)
                Dts.TaskResult = ScriptResults.Failure
                Return
            End If

            Dim alleDateien() As String = {dateipfad}
            Dim gesamt As Integer = 1
            Log("Gefundene Dateien  : 1 (exakt per BA::STLDateiname)")

            ' Datei verarbeiten
            Dim cntErfolgreich As Integer = 0
            Dim cntFehlerhaft As Integer = 0

            For Each dp As String In alleDateien
                Dim dateiname As String = Path.GetFileName(dp)
                Log(String.Format("Verarbeite Datei 1/1: {0}", dateiname))
                Try
                    VerarbeiteDatei(dp, dateiname, connStr)
                    cntErfolgreich += 1
                    Log("Datei erfolgreich verarbeitet: " & dateiname & " OK")
                Catch ex As Exception
                    cntFehlerhaft += 1
                    LogFehler(String.Format("FEHLER in Datei '{0}': {1}", dateiname, ex.Message))
                End Try
            Next

            ' Zusammenfassung
            Log("ZUSAMMENFASSUNG SCR_02_Steuerlisten_Laden")
            Log("Dateien gesamt      : " & gesamt.ToString())
            Log("Erfolgreich         : " & cntErfolgreich.ToString())
            Log("Fehlerhaft          : " & cntFehlerhaft.ToString())

            If cntFehlerhaft > 0 Then
                LogFehler(cntFehlerhaft.ToString() & " Datei(en) konnten nicht verarbeitet werden.")
                Dts.TaskResult = ScriptResults.Failure
            Else
                Log("Datei erfolgreich verarbeitet OK")
                Dts.TaskResult = ScriptResults.Success
            End If

        Catch ex As Exception
            LogFehler("Kritischer Fehler in " & SKRIPT_NAME & ": " & ex.Message)
            Dts.TaskResult = ScriptResults.Failure
        End Try

    End Sub

    ' =========================================================================
    ' Variablen aus SSIS laden
    ' =========================================================================
    Private Sub VariablenLaden()
        _stlOrdner = Dts.Variables("BA::STLOrdner").Value.ToString().Trim()
        _stlDateiname = Dts.Variables("BA::STLDateiname").Value.ToString().Trim()
        _steuerlistenTabelle = Dts.Variables("BA::SteuerlistenTabelle").Value.ToString().Trim()
        _bearbeiter = Dts.Variables("System::CreatorName").Value.ToString().Trim()
    End Sub

    ' =========================================================================
    ' Pflichtfelder prüfen
    ' =========================================================================
    Private Function PflichtfelderPruefen() As Boolean
        Dim fehlend As New System.Text.StringBuilder()
        If String.IsNullOrEmpty(_stlOrdner) Then fehlend.AppendLine("  → BA::STLOrdner")
        If String.IsNullOrEmpty(_stlDateiname) Then fehlend.AppendLine("  → BA::STLDateiname")
        If String.IsNullOrEmpty(_steuerlistenTabelle) Then fehlend.AppendLine("  → BA::SteuerlistenTabelle")
        If fehlend.Length > 0 Then
            LogFehler("Pflichtfelder fehlen:" & Environment.NewLine & fehlend.ToString())
            Return False
        End If
        Log("Pflichtfelder-Pruefung: alle Variablen vorhanden OK")
        Return True
    End Function

    ' =========================================================================
    ' Zusätzliche Spalten in tm_steuerlistenfile_Fakten sicherstellen
    ' =========================================================================
    Private Sub SpaltenSicherstellen(connStr As String)

        Dim sql As String =
"-- tm_steuerlistenfile_Fakten Tabelle sicherstellen
IF NOT EXISTS (
    SELECT 1 FROM sys.tables
    WHERE  name      = '" & _steuerlistenTabelle & "'
    AND    schema_id = SCHEMA_ID('dbo')
)
BEGIN
    CREATE TABLE dbo." & _steuerlistenTabelle & "
    (
        tabelle        VARCHAR(255)   NULL,
        FILE_NAME      NVARCHAR(255)  NULL,
        tabellentyp    NVARCHAR(50)    NULL,
        umgebung       NVARCHAR(50)   NULL,
        load_Date      DATETIME       NULL,
        themengebiet   NVARCHAR(100)  NULL,
        bearbeiter     NVARCHAR(100)  NULL,
        tabname_filter NVARCHAR(500)  NULL,
        where_klausel  NVARCHAR(1000) NULL,
        partition_wert NVARCHAR(500)  NULL,
        obj_gefunden   VARCHAR(50)     NULL,
        ref_datum      DATETIME       NULL
    );
    PRINT 'Tabelle " & _steuerlistenTabelle & " neu angelegt.';
END;

-- Fehlende Spalten nachträglich hinzufügen (Idempotent)
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo." & _steuerlistenTabelle & "') AND name = 'tabname_filter')
    ALTER TABLE dbo." & _steuerlistenTabelle & " ADD tabname_filter NVARCHAR(500) NULL;
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo." & _steuerlistenTabelle & "') AND name = 'where_klausel')
    ALTER TABLE dbo." & _steuerlistenTabelle & " ADD where_klausel NVARCHAR(1000) NULL;
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo." & _steuerlistenTabelle & "') AND name = 'partition_wert')
    ALTER TABLE dbo." & _steuerlistenTabelle & " ADD partition_wert NVARCHAR(500) NULL;
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo." & _steuerlistenTabelle & "') AND name = 'obj_gefunden')
    ALTER TABLE dbo." & _steuerlistenTabelle & " ADD obj_gefunden VARCHAR(3) NULL;
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo." & _steuerlistenTabelle & "') AND name = 'ref_datum')
    ALTER TABLE dbo." & _steuerlistenTabelle & " ADD ref_datum DATETIME NULL;"

        SqlAusfuehren(connStr, sql, "Spalten sicherstellen")
        Log("Tabelle und Spalten geprueft/angelegt OK")

    End Sub

    ' =========================================================================
    ' Einzelne Datei verarbeiten
    ' =========================================================================
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

        Dim zeilenNr As Integer = 0
        Dim cntInsert As Integer = 0
        Dim cntUpdate As Integer = 0
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

            ' Leerzeilen und Kommentare überspringen
            If String.IsNullOrEmpty(zeile) OrElse zeile.StartsWith("#") Then
                cntUebersp += 1
                Continue For
            End If

            ' Tabellenname aus Zeile ableiten
            Dim tabellenname As String = TabellennameBestimmen(zeile)

            Try
                ' Existiert bereits?
                Dim anzahl As Integer = Convert.ToInt32(
                    SqlSkalarAusfuehren(connStr,
                        "SELECT COUNT(*) FROM dbo." & _steuerlistenTabelle &
                        " WHERE tabname_filter = @z AND FILE_NAME = @f",
                        "Zeilenprüfung",
                        New With {.z = zeile, .f = dateiname}))

                If anzahl > 0 Then
                    ' UPDATE
                    SqlMitParameternAusfuehren(connStr,
                        "UPDATE dbo." & _steuerlistenTabelle &
                        " SET   tabelle      = @tab" &
                        "     , tabellentyp  = @typ" &
                        "     , umgebung     = @umb" &
                        "     , load_Date    = @dat" &
                        "     , themengebiet = @thm" &
                        "     , bearbeiter   = @bea" &
                        " WHERE tabname_filter = @z" &
                        " AND   FILE_NAME      = @f",
                        "UPDATE Zeile",
                        New With {
                            .tab = tabellenname,
                            .typ = tabellentyp,
                            .umb = umgebung,
                            .dat = DateTime.Now,
                            .thm = themengebiet,
                            .bea = _bearbeiter,
                            .z = zeile,
                            .f = dateiname
                        })
                    cntUpdate += 1
                Else
                    ' INSERT
                    SqlMitParameternAusfuehren(connStr,
                        "INSERT INTO dbo." & _steuerlistenTabelle &
                        " (tabelle, FILE_NAME, tabellentyp, umgebung," &
                        "  load_Date, themengebiet, bearbeiter, tabname_filter)" &
                        " VALUES (@tab, @f, @typ, @umb, @dat, @thm, @bea, @z)",
                        "INSERT Zeile",
                        New With {
                            .tab = tabellenname,
                            .f = dateiname,
                            .typ = tabellentyp,
                            .umb = umgebung,
                            .dat = DateTime.Now,
                            .thm = themengebiet,
                            .bea = _bearbeiter,
                            .z = zeile
                        })
                    cntInsert += 1
                End If

            Catch ex As Exception
                Throw New Exception(
                    String.Format("Zeile {0} ('{1}'): {2}", zeilenNr, zeile, ex.Message), ex)
            End Try
        Next

        Log(String.Format("  Zeilen gesamt: {0} | INSERT: {1} | UPDATE: {2} | Uebersprungen: {3}",
            zeilenNr, cntInsert, cntUpdate, cntUebersp))

    End Sub

    ' =========================================================================
    ' Tabellenname aus STL-Zeile ableiten
    ' =========================================================================
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

    ' =========================================================================
    ' SQL-Helfer: NonQuery mit Retry
    ' =========================================================================
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
                If versuch = 1 Then Log("SQL Statement [" & beschreibung & "]: " & sql)
                    beschreibung, versuch, MAX_VERSUCHE, ex.Message))
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

    ' =========================================================================
    ' SQL-Helfer: Scalar mit Retry + anonyme Parameter
    ' =========================================================================
    Private Function SqlSkalarAusfuehren(connStr As String,
                                         sql As String,
                                         beschreibung As String,
                                         params As Object) As Object
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
                        Return cmd.ExecuteScalar()
                    End Using
                End Using
            Catch ex As Exception
                letzterFehler = ex
                Log(String.Format("WARNUNG [{0}] Versuch {1}/{2}: {3}",
                    beschreibung, versuch, MAX_VERSUCHE, ex.Message))
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

    ' =========================================================================
    ' SQL-Helfer: NonQuery mit Retry + anonyme Parameter
    ' =========================================================================
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

    ' =========================================================================
    ' Parameter aus anonymem Objekt zu SqlCommand hinzufügen
    ' =========================================================================
    Private Sub ParameterHinzufuegen(cmd As SqlCommand, params As Object)
        If params Is Nothing Then Return
        For Each prop As System.Reflection.PropertyInfo In params.GetType().GetProperties()
            Dim wert As Object = prop.GetValue(params, Nothing)
            cmd.Parameters.AddWithValue(
                "@" & prop.Name,
                If(wert Is Nothing, CObj(DBNull.Value), wert))
        Next
    End Sub

    ' =========================================================================
    ' Verbindungszeichenfolge aus SSIS Connection Manager holen
    ' =========================================================================
    Private Function HoleVerbindungszeichenfolge() As String
        Return Dts.Connections(CONN_NAME).ConnectionString
    End Function

    ' =========================================================================
    ' Logging
    ' =========================================================================
    Private Sub Log(nachricht As String)
        Dim fireAgain As Boolean = False
        Dts.Events.FireInformation(0, SKRIPT_NAME, nachricht, "", 0, fireAgain)
    End Sub

    Private Sub LogFehler(nachricht As String)
        Dts.Events.FireError(0, SKRIPT_NAME, nachricht, "", 0)
    End Sub

    ' =========================================================================
    ' Ergebnistypen
    ' =========================================================================
    Public Enum ScriptResults
        Success = DTSExecResult.Success
        Failure = DTSExecResult.Failure
    End Enum

End Class
