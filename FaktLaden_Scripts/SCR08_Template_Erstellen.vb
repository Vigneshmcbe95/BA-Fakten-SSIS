Option Explicit On
Option Strict On

Imports System
Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports Microsoft.SqlServer.Dts.Runtime

' =============================================================================
'  Skript       : SCR08_Template_Erstellen
'  Paket        : Fakten Laden (SSIS)
'  Zweck        : Erstellt die strukturelle Template-Tabelle
'                 dbo.<fakt>_template per dynamischem SQL: SELECT TOP 0
'                 <columns_dbo> INTO Template FROM #ext_struct.
'                 #ext_struct wird aus columns_ext (tm_polybase_struktur)
'                 aufgebaut - ext.<fakt> existiert erst ab SCR09.
'  Ablauf       : SCHEMADATEN_KOPIERT -> TEMPLATE_ERSTELLT
'  Wiederholung : 3 Versuche je SQL-Anweisung, 30 s Wartezeit
'  Protokoll    : Nur SSIS-Events (FireInformation / FireError)
' =============================================================================
<Microsoft.SqlServer.Dts.Tasks.ScriptTask.SSISScriptTaskEntryPointAttribute()>
<CLSCompliant(False)>
Partial Public Class ScriptMain
    Inherits Microsoft.SqlServer.Dts.Tasks.ScriptTask.VSTARTScriptObjectModelBase

    Private Const SKRIPT_NAME As String = "SCR08_Template_Erstellen"
    Private Const CONN_NAME As String = "Verbindung"
    Private Const MAX_VERSUCHE As Integer = 3
    Private Const WARTE_SEK As Integer = 30

    Private _runID As Integer = 0
    Private _parameterDB As String = String.Empty
    Private _parametertab As String = String.Empty
    Private _datenbank As String = String.Empty
    Private _extTableSchema As String = String.Empty
    Private _extTableName As String = String.Empty

    ' -----------------------------------------------------------------------
    ' Main - Einstiegspunkt - steuert den Ablauf des Skripts.
    ' -----------------------------------------------------------------------
    Public Sub Main()

        Log("SCR08_Template_Erstellen - Start")

        Try
            _runID = Convert.ToInt32(Dts.Variables("BA::RunID").Value)
            _parameterDB = Dts.Variables("BA::ParameterDB").Value.ToString().Trim()
            _parametertab = Dts.Variables("BA::Parametertabelle").Value.ToString().Trim()
            _datenbank = Dts.Variables("BA::Datenbank").Value.ToString().Trim()
            _extTableSchema = Dts.Variables("BA::ExtTableSchema").Value.ToString().Trim()
            _extTableName = Dts.Variables("BA::ExtTableName").Value.ToString().Trim()


            Dim connStr As String = HoleVerbindungszeichenfolge()
            Dim verfahren As List(Of VerfahrenInfo) = VerfahrenLaden(connStr)
            Log("Verfahren zur Verarbeitung: " & verfahren.Count.ToString())

            Dim cntOK As Integer = 0
            Dim cntFehler As Integer = 0

            For Each v As VerfahrenInfo In verfahren
                Log("Verfahren: " & v.Verfahren & " | Themengebiet: " & v.Themengebiet)

                If v.LetzterSchritt = "TEMPLATE_ERSTELLT" Then
                    Log("  bereits abgeschlossen uebersprungen OK")
                    Continue For
                End If

                Try
                    StatusSetzen(connStr, v.ID, "TEMPLATE_ERSTELLEN")
                    TemplateErstellen(connStr, v)
                    StatusSetzen(connStr, v.ID, "TEMPLATE_ERSTELLT")
                    LogSchreiben(connStr, v.Verfahren, "SCHRITT_1",
                        "Template erstellt/geprueft: " & _datenbank & ".dbo." & v.Faktentabelle.ToLower() & "_template")
                    cntOK += 1
                    Log("  Template erstellt OK")
                Catch ex As Exception
                    cntFehler += 1
                    FehlerSetzen(connStr, v.ID, ex.Message)
                    LogSchreiben(connStr, v.Verfahren, "FEHLER_SCR08", ex.Message)
                    LogFehler("FEHLER Verfahren '" & v.Verfahren & "': " & ex.Message)
                End Try
            Next

            Log("Erfolgreich: " & cntOK.ToString() & " | Fehler: " & cntFehler.ToString())
            Dts.TaskResult = If(cntFehler > 0, ScriptResults.Failure, ScriptResults.Success)

        Catch ex As Exception
            LogFehler("Kritischer Fehler: " & ex.Message)
            Dts.TaskResult = ScriptResults.Failure
        End Try

    End Sub

    ' -----------------------------------------------------------------------
    ' TemplateErstellen - Erstellt dbo.<fakt>_template per dynamischem SQL:
    ' SELECT TOP 0 <columns_dbo> INTO Template FROM #ext_struct.
    ' #ext_struct wird aus columns_ext nachgebildet, da die externe Tabelle
    ' ext.<fakt> erst in SCR09 angelegt wird.
    ' -----------------------------------------------------------------------
    Private Sub TemplateErstellen(connStr As String, v As VerfahrenInfo)
        Dim sql As String =
            "DECLARE @t  nvarchar(128) = N'" & v.Faktentabelle.ToLower() & "';" & vbCrLf &
            "DECLARE @s  nvarchar(128) = N'" & v.Themengebiet.Trim().ToLower() & "';" & vbCrLf &
            "DECLARE @db nvarchar(128) = N'" & _datenbank & "';" & vbCrLf &
            "DECLARE @cols nvarchar(max), @colsExt nvarchar(max), @sql nvarchar(max), @tmpl nvarchar(300), @tmplObj int, @diff nvarchar(max);" & vbCrLf & vbCrLf &
            "-- Soll-Spalten (columns_dbo) und ext-Struktur (columns_ext) aus tm_polybase_struktur laden" & vbCrLf &
            "SELECT @cols    = STRING_AGG(CAST(columns_dbo AS nvarchar(max)), CONCAT(N',', CHAR(13), CHAR(10)))" & vbCrLf &
            "                    WITHIN GROUP (ORDER BY colno)," & vbCrLf &
            "       @colsExt = STRING_AGG(CAST(columns_ext AS nvarchar(max)), CONCAT(N',', CHAR(13), CHAR(10)))" & vbCrLf &
            "                    WITHIN GROUP (ORDER BY colno)" & vbCrLf &
            "FROM dbo.tm_polybase_struktur" & vbCrLf &
            "WHERE tabname = LOWER(@t) AND themengebiet = @s;" & vbCrLf & vbCrLf &
            "IF @cols IS NULL OR @colsExt IS NULL" & vbCrLf &
            "    THROW 50001, 'Keine columns_dbo/columns_ext Metadaten gefunden', 1;" & vbCrLf & vbCrLf &
            "SET @tmpl    = CONCAT('[', @db, '].dbo.[', @t, '_template]');" & vbCrLf &
            "SET @tmplObj = OBJECT_ID(@tmpl);" & vbCrLf & vbCrLf &
            "IF @tmplObj IS NULL" & vbCrLf &
            "BEGIN" & vbCrLf &
            "    -- Template existiert nicht -> aus columns_dbo erstellen (nur Struktur)" & vbCrLf &
            "    -- ext.<t> existiert erst ab SCR09 -> Struktur lokal als #ext_struct nachbilden" & vbCrLf &
            "    SET @sql = N'CREATE TABLE #ext_struct (' + @colsExt + N');' +" & vbCrLf &
            "               N'SELECT TOP 0 ' + @cols + N' INTO ' + @tmpl + N' FROM #ext_struct;';" & vbCrLf &
            "    EXEC sp_executesql @sql;" & vbCrLf &
            "END" & vbCrLf &
            "ELSE" & vbCrLf &
            "BEGIN" & vbCrLf &
            "    -- Template existiert -> Struktur (Spaltenname/Typ/Laenge/Nullable) gegen Soll pruefen" & vbCrLf &
            "    SET @sql =" & vbCrLf &
            "        N'CREATE TABLE #ext_struct (' + @colsExt + N');' +" & vbCrLf &
            "        N'SELECT TOP 0 ' + @cols + N' INTO #soll FROM #ext_struct;' +" & vbCrLf &
            "        N';WITH soll AS (SELECT c.name COLLATE DATABASE_DEFAULT AS nm, ty.name COLLATE DATABASE_DEFAULT AS typ, c.max_length AS ml, c.precision AS pr, c.scale AS sc, c.is_nullable AS nu' +" & vbCrLf &
            "          N' FROM tempdb.sys.columns c JOIN tempdb.sys.types ty ON ty.user_type_id=c.user_type_id WHERE c.object_id=OBJECT_ID(N''tempdb..#soll'')),' +" & vbCrLf &
            "          N' ist AS (SELECT c.name COLLATE DATABASE_DEFAULT AS nm, ty.name COLLATE DATABASE_DEFAULT AS typ, c.max_length AS ml, c.precision AS pr, c.scale AS sc, c.is_nullable AS nu' +" & vbCrLf &
            "          N' FROM ' + QUOTENAME(@db) + N'.sys.columns c JOIN ' + QUOTENAME(@db) + N'.sys.types ty ON ty.user_type_id=c.user_type_id WHERE c.object_id=@po)' +" & vbCrLf &
            "          N' SELECT @diff = STRING_AGG(z, CHAR(13)+CHAR(10)) FROM (' +" & vbCrLf &
            "          N'   SELECT CONCAT(COALESCE(s.nm,i.nm), '': '',' +" & vbCrLf &
            "          N'     CASE WHEN s.nm IS NULL THEN ''FEHLT in Oracle (nur im Template vorhanden)''' +" & vbCrLf &
            "          N'          WHEN i.nm IS NULL THEN ''FEHLT im Template (nur in Oracle vorhanden)''' +" & vbCrLf &
            "          N'          ELSE CONCAT(''Oracle='', s.typ, '' len='', s.ml, '' pr='', s.pr, '' sc='', s.sc, '' null='', s.nu, '' <> Template='', i.typ, '' len='', i.ml, '' pr='', i.pr, '' sc='', i.sc, '' null='', i.nu) END) AS z' +" & vbCrLf &
            "          N'   FROM soll s FULL OUTER JOIN ist i ON s.nm=i.nm' +" & vbCrLf &
            "          N'   WHERE s.nm IS NULL OR i.nm IS NULL OR s.typ<>i.typ OR s.ml<>i.ml OR s.pr<>i.pr OR s.sc<>i.sc OR s.nu<>i.nu) d;';" & vbCrLf &
            "    EXEC sp_executesql @sql, N'@po int, @diff nvarchar(max) OUTPUT', @po=@tmplObj, @diff=@diff OUTPUT;" & vbCrLf &
            "    IF @diff IS NOT NULL" & vbCrLf &
            "    BEGIN" & vbCrLf &
            "        DECLARE @msg nvarchar(2048) = LEFT(CONCAT('Template ', @tmpl, ': Struktur weicht von columns_dbo (Soll) ab. Nicht uebereinstimmende Spalten (Spalte | SOLL=Oracle | IST=Template):', CHAR(13), CHAR(10), @diff), 2048);" & vbCrLf &
            "        THROW 50010, @msg, 1;" & vbCrLf &
            "    END" & vbCrLf &
            "END"

        Log("  Template fuer: " & v.Faktentabelle.ToLower())
        SqlAusfuehren(connStr, sql, "Template erstellen/pruefen")
        Log("  Template erstellt bzw. Struktur geprueft OK")
    End Sub

    ' -----------------------------------------------------------------------
    ' VerfahrenLaden - Laedt die zu verarbeitenden Verfahren aus der
    ' Arbeitsliste (Join mit der Parametertabelle).
    ' -----------------------------------------------------------------------
    Private Function VerfahrenLaden(connStr As String) As List(Of VerfahrenInfo)
        Dim liste As New List(Of VerfahrenInfo)()
        Dim sql As String =
"SELECT a.ID, a.Verfahren, a.Themengebiet, a.LetzterSchritt,
        pf.Wert AS Faktentabelle
 FROM   dbo.ETL_Fkt_Arbeitsliste a
 JOIN   " & _parameterDB & ".dbo." & _parametertab & " pf
        ON pf.Verfahren = dbo.fn_ParamVerfahren(a.Verfahren)
        AND pf.Parameter = 'Faktentabelle'
 WHERE  a.Status IN ('SCHEMADATEN_KOPIERT','TEMPLATE_ERSTELLEN')
 AND    a.RunID = " & _runID.ToString() & "
 ORDER  BY a.Verfahren"

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
                                liste.Add(New VerfahrenInfo With {
                                    .ID = rdr.GetInt32(0),
                                    .Verfahren = rdr(1).ToString().Trim(),
                                    .Themengebiet = rdr(2).ToString().Trim(),
                                    .LetzterSchritt = If(rdr.IsDBNull(3), "", rdr(3).ToString().Trim()),
                                    .Faktentabelle = rdr(4).ToString().Trim()
                                })
                            End While
                        End Using
                    End Using
                End Using
                Return liste
            Catch ex As Exception
                Log(String.Format("WARNUNG [Verfahren laden] Versuch {0}/{1}: {2}", versuch, MAX_VERSUCHE, ex.Message))
                If versuch < MAX_VERSUCHE Then
                    System.Threading.Thread.Sleep(WARTE_SEK * 1000)
                Else
                    Throw
                End If
            End Try
        End While
        Return liste
    End Function

    ' -----------------------------------------------------------------------
    ' StatusSetzen - Aktualisiert Status / LetzterSchritt einer
    ' Arbeitslisten-Zeile.
    ' -----------------------------------------------------------------------
    Private Sub StatusSetzen(connStr As String, id As Integer, status As String)
        SqlAusfuehren(connStr,
            "UPDATE dbo.ETL_Fkt_Arbeitsliste SET Status='" & status & "', LetzterSchritt='" & status & "', AktualisiertAm=GETDATE() WHERE ID=" & id.ToString(),
            "Status setzen")
    End Sub

    ' -----------------------------------------------------------------------
    ' FehlerSetzen - Markiert eine Arbeitslisten-Zeile als FEHLER und
    ' speichert die Fehlermeldung.
    ' -----------------------------------------------------------------------
    Private Sub FehlerSetzen(connStr As String, id As Integer, msg As String)
        Dim kurz As String = If(msg.Length > 3900, msg.Substring(0, 3900), msg)
        Try
            Using conn As New SqlConnection(connStr)
                conn.Open()
                Using cmd As New SqlCommand("UPDATE dbo.ETL_Fkt_Arbeitsliste SET Status='FEHLER',Fehlermeldung=@m,AktualisiertAm=GETDATE() WHERE ID=@id", conn)
                    cmd.Parameters.AddWithValue("@m", kurz)
                    cmd.Parameters.AddWithValue("@id", id)
                    cmd.ExecuteNonQuery()
                End Using
            End Using
        Catch
        End Try
    End Sub

    ' -----------------------------------------------------------------------
    ' LogSchreiben - Leitet Protokollmeldungen an SSIS-Events weiter:
    ' FEHLER_* -> FireError, alles andere -> FireInformation.
    ' -----------------------------------------------------------------------
    Private Sub LogSchreiben(connStr As String, verfahren As String, schritt As String, meldung As String)
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
                If versuch < MAX_VERSUCHE Then
                    System.Threading.Thread.Sleep(WARTE_SEK * 1000)
                End If
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
                If versuch < MAX_VERSUCHE Then
                    System.Threading.Thread.Sleep(WARTE_SEK * 1000)
                Else
                    Throw
                End If
            End Try
        End While
        Return Nothing
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
    ' VerfahrenInfo - Datencontainer fuer ein Verfahren der Arbeitsliste.
    ' -----------------------------------------------------------------------
    Private Class VerfahrenInfo
        Public Property ID As Integer
        Public Property Verfahren As String
        Public Property Themengebiet As String
        Public Property LetzterSchritt As String
        Public Property Faktentabelle As String
    End Class

    ' -----------------------------------------------------------------------
    ' ScriptResults - SSIS-Task-Ergebniscodes.
    ' -----------------------------------------------------------------------
    Public Enum ScriptResults
        Success = DTSExecResult.Success
        Failure = DTSExecResult.Failure
    End Enum

End Class
