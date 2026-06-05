Option Explicit On
Option Strict On

Imports System
Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports System.Text
Imports Microsoft.SqlServer.Dts.Runtime

' =============================================================================
' PAKET  : Fakten Laden
' SKRIPT : SCR_04b_Parameter_Vollstaendigkeitspruefung
' ZWECK  : PrÃ¼ft VOR dem Aufbau der Arbeitsliste ob:
'          1. Die Steuerlisten-Tabelle existiert und EintrÃ¤ge enthÃ¤lt
'          2. Alle Verfahren aus der Steuerliste (tabname_filter IS NOT NULL)
'             in der Parametertabelle vorhanden sind
'          3. Alle 7 Pflichtparameter pro Verfahren vorhanden und befÃ¼llt sind
'          â Bei JEDEM Fehler: Paket wird mit detaillierter Fehlermeldung abgebrochen
' ============================================================================='
<Microsoft.SqlServer.Dts.Tasks.ScriptTask.SSISScriptTaskEntryPointAttribute()>
<CLSCompliant(False)>
Partial Public Class ScriptMain
    Inherits Microsoft.SqlServer.Dts.Tasks.ScriptTask.VSTARTScriptObjectModelBase

    ' -------------------------------------------------------------------------
    ' Konstanten
    ' -------------------------------------------------------------------------
    Private Const SKRIPT_NAME As String = "SCR_05_Parameter_Vollstaendigkeitspruefung"
    Private Const CONN_NAME As String = "Verbindung"
    Private Const MAX_VERSUCHE As Integer = 10
    Private Const WARTE_SEK As Integer = 30

    ' -------------------------------------------------------------------------
    ' Pflichtparameter â jeder Eintrag MUSS pro Verfahren in der
    ' Parametertabelle vorhanden und befÃ¼llt sein (Spalte "Parameter")
    ' -------------------------------------------------------------------------
    Private Shared ReadOnly PFLICHT_PARAMETER As String() = New String() {
        "FaktentabelleTemplate",
        "Faktentabelle",
        "Faktenpartitionsspalte",
        "Faktenkomprimierung",
        "FaktenClusteredIndex",
        "DateFormat",
        "Anzahl_ParallelTasks"
    }

    ' -------------------------------------------------------------------------
    ' SSIS-Variablen (werden in VariablenLaden() befÃ¼llt)
    ' -------------------------------------------------------------------------
    Private _parameterDB As String = String.Empty
    Private _parametertabelle As String = String.Empty
    Private _steuerlistenTabelle As String = String.Empty

    ' =========================================================================
    ' Einstiegspunkt
    ' =========================================================================
    Public Sub Main()

        Log("ââââââââââââââââââââââââââââââââââââââââââââââââââââââââ")
        Log("SCR_04b_Parameter_Vollstaendigkeitspruefung â Start")
        Log("Zeitpunkt: " & DateTime.Now.ToString("dd.MM.yyyy HH:mm:ss"))
        Log("ââââââââââââââââââââââââââââââââââââââââââââââââââââââââ")

        Try
            ' -- SSIS-Variablen laden
            VariablenLaden()

            ' -- Pflichtfelder des Skripts selbst prÃ¼fen
            If Not PflichtfelderPruefen() Then
                Dts.TaskResult = ScriptResults.Failure
                Return
            End If

            Dim connStr As String = HoleVerbindungszeichenfolge()

            ' ââ Schritt 1: Steuerlisten-Tabelle prÃ¼fen âââââââââââââââââââââââ
            Log("ââ Schritt 1: Steuerlisten-Tabelle prÃ¼fen")
            If Not SteuerlistenTabellePruefen(connStr) Then
                Dts.TaskResult = ScriptResults.Failure
                Return
            End If

            ' ââ Schritt 2: Verfahren aus Steuerliste lesen âââââââââââââââââââ
            Log("ââ Schritt 2: Verfahren aus Steuerliste laden")
            Dim verfahrenListe As List(Of String) = VerfahrenAusSteuerlisteHolen(connStr)

            If verfahrenListe.Count = 0 Then
                Dim fehler As String =
                    "FEHLER: Keine gÃ¼ltigen Verfahren in der Steuerlisten-Tabelle gefunden." &
                    Environment.NewLine &
                    "  GeprÃ¼fte Tabelle : dbo." & _steuerlistenTabelle &
                    Environment.NewLine &
                    "  Bedingungen      : tabname_filter IS NOT NULL" &
                    Environment.NewLine &
                    "                     tabelle beginnt mit 'tf_' oder 'tt_'" &
                    Environment.NewLine &
                    "â Die Steuerliste enthÃ¤lt keine Zeilen mit befÃ¼lltem 'tabname_filter'" &
                    Environment.NewLine &
                    "  und einem Verfahrensnamen mit PrÃ¤fix 'tf_' oder 'tt_'." &
                    Environment.NewLine &
                    "  Bitte laden Sie die Steuerliste fÃ¼r das jeweilige Verfahren hoch."
                LogFehler(fehler)
                Dts.TaskResult = ScriptResults.Success
                Return
            End If

            Log("Verfahren in Steuerliste gefunden: " & verfahrenListe.Count.ToString())

            ' ââ Schritt 3: Pro Verfahren Parametertabelle prÃ¼fen âââââââââââââ
            Log("ââ Schritt 3: Parametertabelle vollstÃ¤ndig prÃ¼fen")
            Dim gesamtFehler As Boolean = False
            Dim cntOK As Integer = 0
            Dim cntFehler As Integer = 0

            For Each verfahren As String In verfahrenListe

                Log("ââââââââââââââââââââââââââââââââââââââââââââââââââââââââ")
                Log("PrÃ¼fe Verfahren: " & verfahren)

                ' 3a: Existiert das Verfahren Ã¼berhaupt in der Parametertabelle?
                If Not VerfahrenExistiertInParametertabelle(connStr, verfahren) Then
                    ' Verfahren fehlt komplett â Fehlermeldung + weiter mit nÃ¤chstem
                    Dim fehler As String =
                        "FEHLER â Verfahren nicht in Parametertabelle gefunden:" &
                        Environment.NewLine &
                        "  Verfahren        : " & verfahren &
                        Environment.NewLine &
                        "  Parametertabelle : " & _parameterDB & ".dbo." & _parametertabelle &
                        Environment.NewLine &
                        "â Bitte tragen Sie das Verfahren '" & verfahren & "' mit allen " &
                        "Pflichtparametern in die Parametertabelle [" &
                        _parameterDB & "].[dbo].[" & _parametertabelle & "] ein."
                    LogFehler(fehler)
                    gesamtFehler = True
                    cntFehler += 1
                    Continue For
                End If

                ' 3b: Jeden Pflichtparameter einzeln prÃ¼fen
                Dim verfahrenHatFehler As Boolean = False

                For Each param As String In PFLICHT_PARAMETER

                    Dim pruefErgebnis As String = ParameterPruefen(connStr, verfahren, param)

                    Select Case pruefErgebnis

                        Case "OK"
                            Log("  [OK]      " & param)

                        Case "FEHLEND"
                            ' Zeile existiert nicht in der Parametertabelle
                            Dim fehler As String =
                                "FEHLER â Pflichtparameter fehlt in Parametertabelle:" &
                                Environment.NewLine &
                                "  Verfahren        : " & verfahren &
                                Environment.NewLine &
                                "  Fehlender Param  : " & param &
                                Environment.NewLine &
                                "  Parametertabelle : " & _parameterDB & ".dbo." & _parametertabelle &
                                Environment.NewLine &
                                "â Bitte fÃ¼gen Sie den Parameter '" & param & "' fÃ¼r das Verfahren '" &
                                verfahren & "' in die Parametertabelle [" &
                                _parameterDB & "].[dbo].[" & _parametertabelle & "] ein."
                            LogFehler(fehler)
                            Log("  [FEHLEND] " & param & "  â Zeile in Parametertabelle fehlt komplett")
                            verfahrenHatFehler = True
                            gesamtFehler = True

                        Case "LEER"
                            ' Zeile existiert, aber Spalte 'Wert' ist NULL oder leer
                            Dim fehler As String =
                                "FEHLER â Parameterwert ist leer (NULL oder ''):" &
                                Environment.NewLine &
                                "  Verfahren        : " & verfahren &
                                Environment.NewLine &
                                "  Parameter        : " & param &
                                Environment.NewLine &
                                "  Parametertabelle : " & _parameterDB & ".dbo." & _parametertabelle &
                                Environment.NewLine &
                                "â Bitte tragen Sie einen gÃ¼ltigen Wert fÃ¼r den Parameter '" & param &
                                "' beim Verfahren '" & verfahren & "' in der Parametertabelle [" &
                                _parameterDB & "].[dbo].[" & _parametertabelle & "] ein."
                            LogFehler(fehler)
                            Log("  [LEER]    " & param & "  â Wert ist NULL oder leer string")
                            verfahrenHatFehler = True
                            gesamtFehler = True

                        Case Else
                            Log("  [UNBEKANNT] " & param & " â PrÃ¼fergebnis: " & pruefErgebnis)

                    End Select

                Next

                If Not verfahrenHatFehler Then
                    Log("  â Alle Pflichtparameter fÃ¼r '" & verfahren & "' vorhanden und befÃ¼llt.")
                    cntOK += 1
                Else
                    cntFehler += 1
                End If

            Next

            ' ââ Schritt 4: Zusammenfassung und Gesamtergebnis ââââââââââââââââ
            Log("ââââââââââââââââââââââââââââââââââââââââââââââââââââââââ")
            Log("ZUSAMMENFASSUNG SCR_04b_Parameter_Vollstaendigkeitspruefung")
            Log("ââââââââââââââââââââââââââââââââââââââââââââââââââââââââ")
            Log("Verfahren geprÃ¼ft        : " & verfahrenListe.Count.ToString())
            Log("VollstÃ¤ndig (OK)         : " & cntOK.ToString())
            Log("Mit Fehlern              : " & cntFehler.ToString())

            If gesamtFehler Then
                Log("ERGEBNIS: FEHLGESCHLAGEN â Pflichtparameter unvollstÃ¤ndig.")
                Log("ââââââââââââââââââââââââââââââââââââââââââââââââââââââââ")
                LogFehler(
                    "SCR_04b: ParameterprÃ¼fung fehlgeschlagen â " & cntFehler.ToString() &
                    " Verfahren haben fehlende oder leere Pflichtparameter. " &
                    "Bitte prÃ¼fen Sie die obigen Fehlermeldungen und ergÃ¤nzen Sie " &
                    "die fehlenden EintrÃ¤ge in der Parametertabelle [" &
                    _parameterDB & "].[dbo].[" & _parametertabelle & "] " &
                    "bzw. laden Sie die Steuerliste in [dbo].[" & _steuerlistenTabelle & "] hoch.")
                Dts.TaskResult = ScriptResults.Failure
            Else
                Log("ERGEBNIS: BESTANDEN â Alle Parameter vollstÃ¤ndig und befÃ¼llt â")
                Log("ââââââââââââââââââââââââââââââââââââââââââââââââââââââââ")

                ' ââ Schritt 4: where_klausel mit echter Partitionsspalte befÃ¼llen ââ
                Log("ââ Schritt 4: where_klausel in Steuerliste befÃ¼llen")
                Dim cntWhereOK As Integer = 0
                Dim cntWhereNull As Integer = 0

                For Each verfahren As String In verfahrenListe

                    ' Faktenpartitionsspalte aus Parametertabelle lesen
                    Dim partCol As String = PartitionsspalteHolen(connStr, verfahren)

                    If String.IsNullOrEmpty(partCol) Then
                        Log("  WARNUNG: Faktenpartitionsspalte leer fÃ¼r '" & verfahren & "' â where_klausel bleibt unverÃ¤ndert")
                        cntWhereNull += 1
                        Continue For
                    End If

                    ' where_klausel = "WHERE <partCol> = '<partition_wert>'"
                    ' nur fÃ¼r Zeilen wo partition_wert bereits gesetzt ist (durch SCR04)
                    WhereKlauselFuellen(connStr, verfahren, partCol)
                    Log("  where_klausel gefÃ¼llt: " & verfahren & " | Spalte: " & partCol)
                    cntWhereOK += 1

                Next

                Log("ââââââââââââââââââââââââââââââââââââââââââââââââââââââââ")
                Log("where_klausel gefÃ¼llt : " & cntWhereOK.ToString())
                Log("Ohne Partitionsspalte : " & cntWhereNull.ToString())
                Log("ââââââââââââââââââââââââââââââââââââââââââââââââââââââââ")
                Dts.TaskResult = ScriptResults.Success
            End If

        Catch ex As Exception
            LogFehler("Kritischer Fehler in " & SKRIPT_NAME & ": " & ex.Message)
            Dts.TaskResult = ScriptResults.Failure
        End Try

    End Sub

    ' =========================================================================
    ' SSIS-Variablen laden
    ' =========================================================================
    Private Sub VariablenLaden()
        _parameterDB = Dts.Variables("BA::ParameterDB").Value.ToString().Trim()
        _parametertabelle = Dts.Variables("BA::Parametertabelle").Value.ToString().Trim()
        _steuerlistenTabelle = Dts.Variables("BA::SteuerlistenTabelle").Value.ToString().Trim()

        Log("ââââââââââââââââââââââââââââââââââââââââââââââââââââââââ")
        Log("ParameterDB          : " & _parameterDB)
        Log("Parametertabelle     : " & _parametertabelle)
        Log("Steuerlisten-Tabelle : " & _steuerlistenTabelle)
        Log("Pflichtparameter     : " & String.Join(", ", PFLICHT_PARAMETER))
        Log("ââââââââââââââââââââââââââââââââââââââââââââââââââââââââ")
    End Sub

    ' =========================================================================
    ' Pflichtfelder des Skripts selbst prÃ¼fen (SSIS-Variablen)
    ' =========================================================================
    Private Function PflichtfelderPruefen() As Boolean
        Dim fehlend As New StringBuilder()
        If String.IsNullOrEmpty(_parameterDB) Then fehlend.AppendLine("  â BA::ParameterDB")
        If String.IsNullOrEmpty(_parametertabelle) Then fehlend.AppendLine("  â BA::Parametertabelle")
        If String.IsNullOrEmpty(_steuerlistenTabelle) Then fehlend.AppendLine("  â BA::SteuerlistenTabelle")
        If fehlend.Length > 0 Then
            LogFehler("SSIS-Pflichtfelder fehlen:" & Environment.NewLine & fehlend.ToString())
            Return False
        End If
        Log("SSIS-Variablen-PrÃ¼fung: alle vorhanden â")
        Return True
    End Function

    ' =========================================================================
    ' Schritt 4a: Faktenpartitionsspalte fÃ¼r ein Verfahren aus Parametertabelle lesen
    ' =========================================================================
    Private Function PartitionsspalteHolen(connStr As String, verfahren As String) As String
        Dim sql As String =
            "SELECT LTRIM(RTRIM(ISNULL(Wert,''))) " &
            "FROM   " & _parameterDB & ".dbo." & _parametertabelle & " " &
            "WHERE  LOWER(LTRIM(RTRIM(Verfahren))) = @v " &
            "AND    LOWER(LTRIM(RTRIM(Parameter))) = 'faktenpartitionsspalte'"
        Dim versuch As Integer = 0
        Dim letzterFehler As Exception = Nothing
        While versuch < MAX_VERSUCHE
            versuch += 1
            Try
                Using conn As New SqlConnection(connStr)
                    conn.Open()
                    Using cmd As New SqlCommand(sql, conn)
                        cmd.CommandTimeout = 0
                        cmd.Parameters.AddWithValue("@v", verfahren.ToLower())
                        Dim r As Object = cmd.ExecuteScalar()
                        Return If(r Is Nothing OrElse r Is DBNull.Value, String.Empty, r.ToString().Trim())
                    End Using
                End Using
            Catch ex As Exception
                letzterFehler = ex
                If versuch < MAX_VERSUCHE Then System.Threading.Thread.Sleep(WARTE_SEK * 1000)
            End Try
        End While
        Throw New Exception(String.Format(
            "[PartitionsspalteHolen '{0}'] fehlgeschlagen nach {1} Versuchen: {2}",
            verfahren, MAX_VERSUCHE,
            If(letzterFehler IsNot Nothing, letzterFehler.Message, "Unbekannt")))
    End Function

    ' =========================================================================
    ' Schritt 4b: where_klausel in Steuerliste befÃ¼llen
    '             WHERE <partCol> = '<partition_wert>'  â nur wo partition_wert gesetzt
    ' =========================================================================
    Private Sub WhereKlauselFuellen(connStr As String, verfahren As String, partCol As String)
        ' UPDATE alle Zeilen dieses Verfahrens wo partition_wert bereits gesetzt (durch SCR04)
        Dim sql As String =
            "UPDATE dbo." & _steuerlistenTabelle & " " &
            "SET    where_klausel = 'WHERE " & partCol & " = ''' + partition_wert + '''' " &
            "WHERE  LOWER(LTRIM(RTRIM(tabelle))) = @v "
        Dim versuch As Integer = 0
        Dim letzterFehler As Exception = Nothing
        While versuch < MAX_VERSUCHE
            versuch += 1
            Try
                Using conn As New SqlConnection(connStr)
                    conn.Open()
                    Using cmd As New SqlCommand(sql, conn)
                        cmd.CommandTimeout = 0
                        cmd.Parameters.AddWithValue("@v", verfahren.ToLower())
                        Dim rows As Integer = cmd.ExecuteNonQuery()
                        Log("    â " & rows.ToString() & " Zeile(n) aktualisiert")
                        Return
                    End Using
                End Using
            Catch ex As Exception
                letzterFehler = ex
                If versuch < MAX_VERSUCHE Then System.Threading.Thread.Sleep(WARTE_SEK * 1000)
            End Try
        End While
        Throw New Exception(String.Format(
            "[WhereKlauselFuellen '{0}'] fehlgeschlagen nach {1} Versuchen: {2}",
            verfahren, MAX_VERSUCHE,
            If(letzterFehler IsNot Nothing, letzterFehler.Message, "Unbekannt")))
    End Sub

    ' =========================================================================
    ' Schritt 1: PrÃ¼ft ob Steuerlisten-Tabelle existiert und EintrÃ¤ge hat
    ' =========================================================================
    Private Function SteuerlistenTabellePruefen(connStr As String) As Boolean

        ' 1a: Tabelle Ã¼berhaupt vorhanden?
        Dim sqlExistenz As String =
            "SELECT COUNT(1) FROM sys.tables " &
            "WHERE name = '" & _steuerlistenTabelle & "' " &
            "AND schema_id = SCHEMA_ID('dbo')"

        Dim tabelleExistiert As Integer = Convert.ToInt32(
            SqlSkalar(connStr, sqlExistenz, "Steuerlisten-Tabelle ExistenzprÃ¼fung"))

        If tabelleExistiert = 0 Then
            LogFehler(
                "FEHLER â Steuerlisten-Tabelle nicht gefunden:" &
                Environment.NewLine &
                "  Gesuchte Tabelle : dbo." & _steuerlistenTabelle &
                Environment.NewLine &
                "â Die Steuerlisten-Tabelle [dbo].[" & _steuerlistenTabelle & "] existiert nicht " &
                "in der Datenbank. Bitte laden Sie zunÃ¤chst die Steuerliste mit dem " &
                "entsprechenden Verfahrensnamen hoch (SCR_03_Steuerlisten_Laden).")
            Return False
        End If

        Log("Steuerlisten-Tabelle [dbo." & _steuerlistenTabelle & "] vorhanden â")

        ' 1b: EnthÃ¤lt die Tabelle Ã¼berhaupt Zeilen?
        Dim sqlAnzahl As String =
            "SELECT COUNT(1) FROM dbo." & _steuerlistenTabelle

        Dim anzahlZeilen As Integer = Convert.ToInt32(
            SqlSkalar(connStr, sqlAnzahl, "Steuerlisten-Tabelle Zeilenanzahl"))

        If anzahlZeilen = 0 Then
            LogFehler(
                "FEHLER â Steuerlisten-Tabelle ist leer:" &
                Environment.NewLine &
                "  Tabelle : dbo." & _steuerlistenTabelle &
                Environment.NewLine &
                "â Die Steuerlisten-Tabelle [dbo].[" & _steuerlistenTabelle & "] enthÃ¤lt keine " &
                "DatensÃ¤tze. Bitte laden Sie die Steuerliste fÃ¼r das jeweilige Verfahren hoch.")
            Return False
        End If

        Log("Steuerlisten-Tabelle enthÃ¤lt " & anzahlZeilen.ToString() & " Zeile(n) â")
        Return True

    End Function

    ' =========================================================================
    ' Schritt 2: Verfahren aus Steuerliste wo where_klausel IS NOT NULL
    '            JOIN Parametertabelle â nur Verfahren die dort existieren
    ' =========================================================================
    Private Function VerfahrenAusSteuerlisteHolen(connStr As String) As List(Of String)

        Dim liste As New List(Of String)()
        Dim sql As String =
        "SELECT DISTINCT LOWER(LTRIM(RTRIM(s.tabelle))) " &
        "FROM   dbo." & _steuerlistenTabelle & " s " &
        "INNER JOIN " & _parameterDB & ".dbo." & _parametertabelle & " p " &
        "    ON LOWER(LTRIM(RTRIM(s.tabelle))) = LOWER(LTRIM(RTRIM(p.Verfahren))) " &
        "ORDER  BY 1"

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
                                If Not rdr.IsDBNull(0) Then
                                    liste.Add(rdr.GetString(0).Trim())
                                End If
                            End While
                        End Using
                    End Using
                End Using
                Return liste
            Catch ex As Exception
                letzterFehler = ex
                Log(String.Format("WARNUNG [Verfahren aus Steuerliste] Versuch {0}/{1}: {2}",
                    versuch, MAX_VERSUCHE, ex.Message))
                If versuch < MAX_VERSUCHE Then System.Threading.Thread.Sleep(WARTE_SEK * 1000)
            End Try
        End While

        Throw New Exception(String.Format(
            "[Verfahren aus Steuerliste laden] fehlgeschlagen nach {0} Versuchen: {1}",
            MAX_VERSUCHE,
            If(letzterFehler IsNot Nothing, letzterFehler.Message, "Unbekannt")))

    End Function

    ' =========================================================================
    ' Schritt 3a: PrÃ¼ft ob ein Verfahren Ã¼berhaupt in der Parametertabelle
    '             existiert (mindestens eine Zeile mit diesem Verfahren)
    ' =========================================================================
    Private Function VerfahrenExistiertInParametertabelle(connStr As String,
                                                           verfahren As String) As Boolean
        Dim sql As String =
            "SELECT COUNT(1) " &
            "FROM   " & _parameterDB & ".dbo." & _parametertabelle & " " &
            "WHERE  LOWER(LTRIM(RTRIM(Verfahren))) = @v"

        Dim versuch As Integer = 0
        Dim letzterFehler As Exception = Nothing

        While versuch < MAX_VERSUCHE
            versuch += 1
            Try
                Using conn As New SqlConnection(connStr)
                    conn.Open()
                    Using cmd As New SqlCommand(sql, conn)
                        cmd.CommandTimeout = 0
                        cmd.Parameters.AddWithValue("@v", verfahren.ToLower())
                        Dim anzahl As Integer = Convert.ToInt32(cmd.ExecuteScalar())
                        Return anzahl > 0
                    End Using
                End Using
            Catch ex As Exception
                letzterFehler = ex
                Log(String.Format("WARNUNG [Verfahren Existenz] Versuch {0}/{1}: {2}",
                    versuch, MAX_VERSUCHE, ex.Message))
                If versuch < MAX_VERSUCHE Then System.Threading.Thread.Sleep(WARTE_SEK * 1000)
            End Try
        End While

        Throw New Exception(String.Format(
            "[Verfahren Existenz prÃ¼fen] fehlgeschlagen nach {0} Versuchen: {1}",
            MAX_VERSUCHE,
            If(letzterFehler IsNot Nothing, letzterFehler.Message, "Unbekannt")))
    End Function

    ' =========================================================================
    ' Schritt 3b: PrÃ¼ft einen einzelnen Pflichtparameter fÃ¼r ein Verfahren
    '  RÃ¼ckgabe:
    '    "OK"       â Zeile vorhanden, Wert befÃ¼llt
    '    "FEHLEND"  â Zeile in Parametertabelle nicht vorhanden
    '    "LEER"     â Zeile vorhanden, Wert ist NULL oder leer
    ' =========================================================================
    Private Function ParameterPruefen(connStr As String,
                                       verfahren As String,
                                       parameter As String) As String

        ' PrÃ¼ft ob die Zeile existiert UND ob Wert befÃ¼llt ist in einem Query
        Dim sql As String =
            "SELECT CASE " &
            "  WHEN COUNT(1) = 0                          THEN 'FEHLEND' " &
            "  WHEN SUM(CASE WHEN LTRIM(RTRIM(ISNULL(CAST(Wert AS NVARCHAR(MAX)),'')))" &
            "       = '' THEN 1 ELSE 0 END) > 0           THEN 'LEER' " &
            "  ELSE 'OK' " &
            "END " &
            "FROM " & _parameterDB & ".dbo." & _parametertabelle & " " &
            "WHERE LOWER(LTRIM(RTRIM(Verfahren)))  = @v " &
            "AND   LOWER(LTRIM(RTRIM(Parameter)))  = @p"

        Dim versuch As Integer = 0
        Dim letzterFehler As Exception = Nothing

        While versuch < MAX_VERSUCHE
            versuch += 1
            Try
                Using conn As New SqlConnection(connStr)
                    conn.Open()
                    Using cmd As New SqlCommand(sql, conn)
                        cmd.CommandTimeout = 0
                        cmd.Parameters.AddWithValue("@v", verfahren.ToLower())
                        cmd.Parameters.AddWithValue("@p", parameter.ToLower())
                        Dim ergebnis As Object = cmd.ExecuteScalar()
                        If ergebnis Is Nothing OrElse ergebnis Is DBNull.Value Then
                            Return "FEHLEND"
                        End If
                        Return ergebnis.ToString().Trim()
                    End Using
                End Using
            Catch ex As Exception
                letzterFehler = ex
                Log(String.Format("WARNUNG [Parameter prÃ¼fen '{0}'] Versuch {1}/{2}: {3}",
                    parameter, versuch, MAX_VERSUCHE, ex.Message))
                If versuch < MAX_VERSUCHE Then System.Threading.Thread.Sleep(WARTE_SEK * 1000)
            End Try
        End While

        Throw New Exception(String.Format(
            "[Parameter prÃ¼fen '{0}'] fehlgeschlagen nach {1} Versuchen: {2}",
            parameter, MAX_VERSUCHE,
            If(letzterFehler IsNot Nothing, letzterFehler.Message, "Unbekannt")))
    End Function

    ' =========================================================================
    ' Hilfsfunktion: Skalaren SQL-Wert lesen mit Retry-Logik
    ' =========================================================================
    Private Function SqlSkalar(connStr As String,
                                sql As String,
                                beschreibung As String) As Object

        Dim versuch As Integer = 0
        Dim letzterFehler As Exception = Nothing

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
                letzterFehler = ex
                Log(String.Format("WARNUNG [{0}] Versuch {1}/{2}: {3}",
                    beschreibung, versuch, MAX_VERSUCHE, ex.Message))
                If versuch < MAX_VERSUCHE Then System.Threading.Thread.Sleep(WARTE_SEK * 1000)
            End Try
        End While

        Throw New Exception(String.Format(
            "[{0}] fehlgeschlagen nach {1} Versuchen: {2}",
            beschreibung, MAX_VERSUCHE,
            If(letzterFehler IsNot Nothing, letzterFehler.Message, "Unbekannt")))
    End Function

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
