Option Explicit On
Option Strict On

Imports System
Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports System.Text
Imports Microsoft.SqlServer.Dts.Runtime

' =============================================================================
'  Skript       : SCR_05_Parameter_Vollstaendigkeitspruefung
'  Paket        : Fakten Laden (SSIS)
'  Zweck        : Prueft, ob jedes Verfahren alle Pflichtparameter in der
'                 Parametertabelle besitzt.
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
    Private Const SKRIPT_NAME As String = "SCR_05_Parameter_Vollstaendigkeitspruefung"
    Private Const CONN_NAME As String = "Verbindung"
    Private Const MAX_VERSUCHE As Integer = 3
    Private Const WARTE_SEK As Integer = 30

    ' -------------------------------------------------------------------------
    ' Pflichtparameter – jeder Eintrag MUSS pro Verfahren in der
    ' Parametertabelle vorhanden und befüllt sein (Spalte "Parameter")
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
    ' SSIS-Variablen (werden in VariablenLaden() befüllt)
    ' -------------------------------------------------------------------------
    Private _parameterDB As String = String.Empty
    Private _parametertabelle As String = String.Empty
    Private _steuerlistenTabelle As String = String.Empty

    ' -----------------------------------------------------------------------
    ' Main - Einstiegspunkt - steuert den Ablauf des Skripts.
    ' -----------------------------------------------------------------------
    Public Sub Main()

        Log("SCR_05_Parameter_Vollstaendigkeitspruefung - Start")

        Try
            ' -- SSIS-Variablen laden
            VariablenLaden()

            ' -- Pflichtfelder des Skripts selbst prüfen
            If Not PflichtfelderPruefen() Then
                Dts.TaskResult = ScriptResults.Failure
                Return
            End If

            Dim connStr As String = HoleVerbindungszeichenfolge()

            ' ── Schritt 1: Steuerlisten-Tabelle prüfen ───────────────────────
            Log("Schritt 1: Steuerlisten-Tabelle pruefen")
            If Not SteuerlistenTabellePruefen(connStr) Then
                Dts.TaskResult = ScriptResults.Failure
                Return
            End If

            ' ── Schritt 2: Verfahren aus Steuerliste lesen ───────────────────
            Log("Schritt 2: Verfahren aus Steuerliste laden")
            Dim verfahrenListe As List(Of String) = VerfahrenAusSteuerlisteHolen(connStr)

            If verfahrenListe.Count = 0 Then
                LogFehler("FEHLER: Die Steuerlisten-Tabelle [dbo." & _steuerlistenTabelle &
                          "] enthaelt keine Tabellen. Bitte Steuerliste hochladen.")
                Dts.TaskResult = ScriptResults.Failure
                Return
            End If

            Log("Verfahren in Steuerliste gefunden: " & verfahrenListe.Count.ToString())

            ' ── Schritt 3: Pro Verfahren Parametertabelle prüfen ─────────────
            Log("Schritt 3: Parametertabelle vollstaendig pruefen")
            Dim gesamtFehler As Boolean = False
            Dim cntOK As Integer = 0
            Dim cntFehler As Integer = 0

            For Each verfahren As String In verfahrenListe

                Log("Pruefe Verfahren: " & verfahren)

                ' 3a: Existiert das Verfahren überhaupt in der Parametertabelle?
                If Not VerfahrenExistiertInParametertabelle(connStr, verfahren) Then
                    Dim fehler As String =
                        "FEHLER – Verfahren nicht in Parametertabelle gefunden:" &
                        Environment.NewLine &
                        "  Verfahren        : " & verfahren &
                        Environment.NewLine &
                        "  Parametertabelle : " & _parameterDB & ".dbo." & _parametertabelle &
                        Environment.NewLine &
                        "→ Bitte tragen Sie das Verfahren '" & verfahren & "' mit allen " &
                        "Pflichtparametern in die Parametertabelle [" &
                        _parameterDB & "].[dbo].[" & _parametertabelle & "] ein."
                    LogFehler(fehler)
                    gesamtFehler = True
                    cntFehler += 1
                    Continue For
                End If

                ' 3b: Jeden Pflichtparameter einzeln prüfen
                Dim verfahrenHatFehler As Boolean = False

                For Each param As String In PFLICHT_PARAMETER

                    Dim pruefErgebnis As String = ParameterPruefen(connStr, verfahren, param)

                    Select Case pruefErgebnis

                        Case "OK"
                            Log("  [OK]      " & param)

                        Case "FEHLEND"
                            Dim fehler As String =
                                "FEHLER – Pflichtparameter fehlt in Parametertabelle:" &
                                Environment.NewLine &
                                "  Verfahren        : " & verfahren &
                                Environment.NewLine &
                                "  Fehlender Param  : " & param &
                                Environment.NewLine &
                                "  Parametertabelle : " & _parameterDB & ".dbo." & _parametertabelle &
                                Environment.NewLine &
                                "→ Bitte fügen Sie den Parameter '" & param & "' für das Verfahren '" &
                                verfahren & "' in die Parametertabelle [" &
                                _parameterDB & "].[dbo].[" & _parametertabelle & "] ein."
                            LogFehler(fehler)
                            Log("  [FEHLEND] " & param & "  <- Zeile in Parametertabelle fehlt komplett")
                            verfahrenHatFehler = True
                            gesamtFehler = True

                        Case "LEER"
                            Dim fehler As String =
                                "FEHLER – Parameterwert ist leer (NULL oder ''):" &
                                Environment.NewLine &
                                "  Verfahren        : " & verfahren &
                                Environment.NewLine &
                                "  Parameter        : " & param &
                                Environment.NewLine &
                                "  Parametertabelle : " & _parameterDB & ".dbo." & _parametertabelle &
                                Environment.NewLine &
                                "→ Bitte tragen Sie einen gültigen Wert für den Parameter '" & param &
                                "' beim Verfahren '" & verfahren & "' in der Parametertabelle [" &
                                _parameterDB & "].[dbo].[" & _parametertabelle & "] ein."
                            LogFehler(fehler)
                            Log("  [LEER]    " & param & "  <- Wert ist NULL oder leer string")
                            verfahrenHatFehler = True
                            gesamtFehler = True

                        Case Else
                            Log("  [UNBEKANNT] " & param & " Pruefergebnis: " & pruefErgebnis)

                    End Select

                Next

                If Not verfahrenHatFehler Then
                    Log("  OK Alle Pflichtparameter fuer '" & verfahren & "' vorhanden und befuellt.")
                    cntOK += 1
                Else
                    cntFehler += 1
                End If

            Next

            ' ── Schritt 4: Zusammenfassung und Gesamtergebnis ────────────────
            Log("ZUSAMMENFASSUNG SCR_05_Parameter_Vollstaendigkeitspruefung")
            Log("Verfahren geprueft        : " & verfahrenListe.Count.ToString())
            Log("Vollstaendig (OK)         : " & cntOK.ToString())
            Log("Mit Fehlern              : " & cntFehler.ToString())

            If gesamtFehler Then
                Log("ERGEBNIS: FEHLGESCHLAGEN - Pflichtparameter unvollstaendig.")
                LogFehler(
                    "SCR_05: Parameterprüfung fehlgeschlagen – " & cntFehler.ToString() &
                    " Verfahren haben fehlende oder leere Pflichtparameter. " &
                    "Bitte prüfen Sie die obigen Fehlermeldungen und ergänzen Sie " &
                    "die fehlenden Einträge in der Parametertabelle [" &
                    _parameterDB & "].[dbo].[" & _parametertabelle & "] " &
                    "bzw. laden Sie die Steuerliste in [dbo].[" & _steuerlistenTabelle & "] hoch.")
                Dts.TaskResult = ScriptResults.Failure
            Else
                Log("ERGEBNIS: BESTANDEN - Alle Parameter vollstaendig und befuellt OK")
                Dts.TaskResult = ScriptResults.Success
            End If

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
        _parameterDB = Dts.Variables("BA::ParameterDB").Value.ToString().Trim()
        _parametertabelle = Dts.Variables("BA::Parametertabelle").Value.ToString().Trim()
        _steuerlistenTabelle = Dts.Variables("BA::SteuerlistenTabelle").Value.ToString().Trim()

        Log("Pflichtparameter     : " & String.Join(", ", PFLICHT_PARAMETER))
    End Sub

    ' -----------------------------------------------------------------------
    ' PflichtfelderPruefen - Prueft, ob alle Pflichtvariablen / -parameter
    ' vorhanden sind.
    ' -----------------------------------------------------------------------
    Private Function PflichtfelderPruefen() As Boolean
        Dim fehlend As New StringBuilder()
        If String.IsNullOrEmpty(_parameterDB) Then fehlend.AppendLine("  → BA::ParameterDB")
        If String.IsNullOrEmpty(_parametertabelle) Then fehlend.AppendLine("  → BA::Parametertabelle")
        If String.IsNullOrEmpty(_steuerlistenTabelle) Then fehlend.AppendLine("  → BA::SteuerlistenTabelle")
        If fehlend.Length > 0 Then
            LogFehler("SSIS-Pflichtfelder fehlen:" & Environment.NewLine & fehlend.ToString())
            Return False
        End If
        Log("SSIS-Variablen-Pruefung: alle vorhanden OK")
        Return True
    End Function

    ' -----------------------------------------------------------------------
    ' SteuerlistenTabellePruefen - Stellt sicher, dass die
    ' Steuerlisten-Tabelle existiert.
    ' -----------------------------------------------------------------------
    Private Function SteuerlistenTabellePruefen(connStr As String) As Boolean

        Dim sqlExistenz As String =
            "SELECT COUNT(1) FROM sys.tables " &
            "WHERE name = '" & _steuerlistenTabelle & "' " &
            "AND schema_id = SCHEMA_ID('dbo')"

        Dim tabelleExistiert As Integer = Convert.ToInt32(
            SqlSkalar(connStr, sqlExistenz, "Steuerlisten-Tabelle Existenzprüfung"))

        If tabelleExistiert = 0 Then
            LogFehler(
                "FEHLER – Steuerlisten-Tabelle nicht gefunden:" &
                Environment.NewLine &
                "  Gesuchte Tabelle : dbo." & _steuerlistenTabelle &
                Environment.NewLine &
                "→ Die Steuerlisten-Tabelle [dbo].[" & _steuerlistenTabelle & "] existiert nicht " &
                "in der Datenbank. Bitte laden Sie zunächst die Steuerliste mit dem " &
                "entsprechenden Verfahrensnamen hoch (SCR_03_Steuerlisten_Laden).")
            Return False
        End If

        Log("Steuerlisten-Tabelle [dbo." & _steuerlistenTabelle & "] vorhanden OK")

        Dim sqlAnzahl As String =
            "SELECT COUNT(1) FROM dbo." & _steuerlistenTabelle

        Dim anzahlZeilen As Integer = Convert.ToInt32(
            SqlSkalar(connStr, sqlAnzahl, "Steuerlisten-Tabelle Zeilenanzahl"))

        If anzahlZeilen = 0 Then
            LogFehler(
                "FEHLER – Steuerlisten-Tabelle ist leer:" &
                Environment.NewLine &
                "  Tabelle : dbo." & _steuerlistenTabelle &
                Environment.NewLine &
                "→ Die Steuerlisten-Tabelle [dbo].[" & _steuerlistenTabelle & "] enthält keine " &
                "Datensätze. Bitte laden Sie die Steuerliste für das jeweilige Verfahren hoch.")
            Return False
        End If

        Log("Steuerlisten-Tabelle enthaelt " & anzahlZeilen.ToString() & " Zeile(n) OK")
        Return True

    End Function

    ' -----------------------------------------------------------------------
    ' VerfahrenAusSteuerlisteHolen - Laedt die Verfahren-Eintraege aus der
    ' Steuerlisten-Tabelle.
    ' -----------------------------------------------------------------------
    Private Function VerfahrenAusSteuerlisteHolen(connStr As String) As List(Of String)

        ' Bewusst OHNE Join auf die Parametertabelle: JEDES Verfahren aus der
        ' Steuerliste wird in Schritt 3 gegen die Parametertabelle geprueft -
        ' fehlt es dort, gibt es pro Verfahren eine klare Fehlermeldung.
        Dim liste As New List(Of String)()
        Dim sql As String =
        "SELECT DISTINCT LOWER(LTRIM(RTRIM(s.tabelle))) " &
        "FROM   dbo." & _steuerlistenTabelle & " s " &
        "WHERE  s.tabelle IS NOT NULL AND LTRIM(RTRIM(s.tabelle)) <> '' " &
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

    ' -----------------------------------------------------------------------
    ' VerfahrenExistiertInParametertabelle - Prueft, ob ein Verfahren in
    ' der Parametertabelle existiert.
    ' -----------------------------------------------------------------------
    Private Function VerfahrenExistiertInParametertabelle(connStr As String,
                                                           verfahren As String) As Boolean
        Dim sql As String =
            "SELECT COUNT(1) " &
            "FROM   " & _parameterDB & ".dbo." & _parametertabelle & " " &
            "WHERE  LOWER(LTRIM(RTRIM(Verfahren))) = LOWER(dbo.fn_ParamVerfahren(@v))"

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
            "[Verfahren Existenz prüfen] fehlgeschlagen nach {0} Versuchen: {1}",
            MAX_VERSUCHE,
            If(letzterFehler IsNot Nothing, letzterFehler.Message, "Unbekannt")))
    End Function

    ' -----------------------------------------------------------------------
    ' ParameterPruefen - Prueft den Parametersatz eines Verfahrens.
    ' -----------------------------------------------------------------------
    Private Function ParameterPruefen(connStr As String,
                                       verfahren As String,
                                       parameter As String) As String

        Dim sql As String =
            "SELECT CASE " &
            "  WHEN COUNT(1) = 0                          THEN 'FEHLEND' " &
            "  WHEN SUM(CASE WHEN LTRIM(RTRIM(ISNULL(CAST(Wert AS NVARCHAR(MAX)),'')))= '' THEN 1 ELSE 0 END) > 0           THEN 'LEER' " &
            "  ELSE 'OK' " &
            "END " &
            "FROM " & _parameterDB & ".dbo." & _parametertabelle & " " &
            "WHERE LOWER(LTRIM(RTRIM(Verfahren)))  = LOWER(dbo.fn_ParamVerfahren(@v)) " &
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
                Log(String.Format("WARNUNG [Parameter pruefen '{0}'] Versuch {1}/{2}: {3}",
                    parameter, versuch, MAX_VERSUCHE, ex.Message))
                If versuch < MAX_VERSUCHE Then System.Threading.Thread.Sleep(WARTE_SEK * 1000)
            End Try
        End While

        Throw New Exception(String.Format(
            "[Parameter prüfen '{0}'] fehlgeschlagen nach {1} Versuchen: {2}",
            parameter, MAX_VERSUCHE,
            If(letzterFehler IsNot Nothing, letzterFehler.Message, "Unbekannt")))
    End Function

    ' -----------------------------------------------------------------------
    ' SqlSkalar - Fuehrt eine skalare SQL-Abfrage mit Wiederholung aus.
    ' -----------------------------------------------------------------------
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
                If versuch = 1 Then Log("SQL Statement [" & beschreibung & "]: " & sql)
                If versuch < MAX_VERSUCHE Then System.Threading.Thread.Sleep(WARTE_SEK * 1000)
            End Try
        End While

        Throw New Exception(String.Format(
            "[{0}] fehlgeschlagen nach {1} Versuchen: {2}",
            beschreibung, MAX_VERSUCHE,
            If(letzterFehler IsNot Nothing, letzterFehler.Message, "Unbekannt")))
    End Function

    ' -----------------------------------------------------------------------
    ' HoleVerbindungszeichenfolge - Liefert den Connection String.
    ' -----------------------------------------------------------------------
    Private Function HoleVerbindungszeichenfolge() As String
        Return Dts.Connections(CONN_NAME).ConnectionString
    End Function

    ' -----------------------------------------------------------------------
    ' Log - Schreibt eine Informationsmeldung in das SSIS-Protokoll.
    ' -----------------------------------------------------------------------
    Private Sub Log(nachricht As String)
        Dim fireAgain As Boolean = False
        Dts.Events.FireInformation(0, SKRIPT_NAME, nachricht, "", 0, fireAgain)
    End Sub

    ' -----------------------------------------------------------------------
    ' LogFehler - Schreibt eine Fehlermeldung in das SSIS-Protokoll.
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
