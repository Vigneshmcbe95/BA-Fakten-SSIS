Option Explicit On
Option Strict On

Imports System
Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports Microsoft.SqlServer.Dts.Runtime

' =============================================================================
' PAKET  : Fakten Laden
' SKRIPT : SCR08_Template_Erstellen (v5 - Direct from Oracle)
' ZWECK  : Pro Verfahren: Template direkt aus Oracle ext table erstellen
'          Baut columns_dbo und columns_ext on-the-fly
'          Status: SCHEMADATEN_KOPIERT → TEMPLATE_ERSTELLT
' =============================================================================
<Microsoft.SqlServer.Dts.Tasks.ScriptTask.SSISScriptTaskEntryPointAttribute()>
<CLSCompliant(False)>
Partial Public Class ScriptMain
    Inherits Microsoft.SqlServer.Dts.Tasks.ScriptTask.VSTARTScriptObjectModelBase

    Private Const SKRIPT_NAME As String = "SCR08_Template_Erstellen"
    Private Const CONN_NAME As String = "Verbindung"
    Private Const MAX_VERSUCHE As Integer = 10
    Private Const WARTE_SEK As Integer = 30

    Private _runID As Integer = 0
    Private _parameterDB As String = String.Empty
    Private _parametertab As String = String.Empty
    Private _datenbank As String = String.Empty
    Private _extTableSchema As String = String.Empty
    Private _extTableName As String = String.Empty

    Public Sub Main()

        Log("════════════════════════════════════════════════════════")
        Log("SCR06_Template_Erstellen – Start (v5: Direct from Oracle)")
        Log("Zeitpunkt: " & DateTime.Now.ToString("dd.MM.yyyy HH:mm:ss"))
        Log("════════════════════════════════════════════════════════")

        Try
            _runID = Convert.ToInt32(Dts.Variables("BA::RunID").Value)
            _parameterDB = Dts.Variables("BA::ParameterDB").Value.ToString().Trim()
            _parametertab = Dts.Variables("BA::Parametertabelle").Value.ToString().Trim()
            _datenbank = Dts.Variables("BA::Datenbank").Value.ToString().Trim()
            _extTableSchema = Dts.Variables("BA::ExtTableSchema").Value.ToString().Trim()
            _extTableName = Dts.Variables("BA::ExtTableName").Value.ToString().Trim()

            Log("Parameter DB       : " & _parameterDB)
            Log("Parameter Tabelle  : " & _parametertab)
            Log("Datenbank          : " & _datenbank)
            Log("Ext Schema         : " & _extTableSchema)
            Log("Ext Tabelle        : " & _extTableName)

            Dim connStr As String = HoleVerbindungszeichenfolge()
            Dim verfahren As List(Of VerfahrenInfo) = VerfahrenLaden(connStr)
            Log("Verfahren zur Verarbeitung: " & verfahren.Count.ToString())

            Dim cntOK As Integer = 0
            Dim cntFehler As Integer = 0

            For Each v As VerfahrenInfo In verfahren
                Log("────────────────────────────────────────────────────────")
                Log("Verfahren: " & v.Verfahren & " | Themengebiet: " & v.Themengebiet)

                If v.LetzterSchritt = "TEMPLATE_ERSTELLT" Then
                    Log("  → bereits abgeschlossen → übersprungen ✓")
                    Continue For
                End If

                Try
                    StatusSetzen(connStr, v.ID, "TEMPLATE_ERSTELLEN")
                    TemplateErstellen(connStr, v)
                    StatusSetzen(connStr, v.ID, "TEMPLATE_ERSTELLT")
                    LogSchreiben(connStr, v.Verfahren, "SCHRITT_1",
                        "Template erstellt: dwh.dbo." & v.Faktentabelle.ToLower() & "_template")
                    cntOK += 1
                    Log("  → Template erstellt ✓")
                Catch ex As Exception
                    cntFehler += 1
                    FehlerSetzen(connStr, v.ID, ex.Message)
                    LogSchreiben(connStr, v.Verfahren, "FEHLER_SCR06", ex.Message)
                    LogFehler("FEHLER Verfahren '" & v.Verfahren & "': " & ex.Message)
                End Try
            Next

            Log("════════════════════════════════════════════════════════")
            Log("Erfolgreich: " & cntOK.ToString() & " | Fehler: " & cntFehler.ToString())
            Log("════════════════════════════════════════════════════════")
            Dts.TaskResult = If(cntFehler > 0, ScriptResults.Failure, ScriptResults.Success)

        Catch ex As Exception
            LogFehler("Kritischer Fehler: " & ex.Message)
            Dts.TaskResult = ScriptResults.Failure
        End Try

    End Sub

    ' =============================================================================
    ' TEMPLATE ERSTELLEN
    ' Gleiche Dynamik wie SCR07: STRING_AGG(columns_dbo) -> @cols -> SET @sql -> EXEC sp_executesql
    ' SELECT TOP 0 (nur Struktur, keine Daten) — Template nur wenn noch nicht vorhanden
    ' Ergebnis: [db].dbo.[tabelle_template] mit echten Spalten (aam_id, jhw_id ...)
    ' =============================================================================
    Private Sub TemplateErstellen(connStr As String, v As VerfahrenInfo)
        Dim sql As String =
            "DECLARE @t  nvarchar(128) = N'" & v.Faktentabelle.ToLower() & "';" & vbCrLf &
            "DECLARE @s  nvarchar(128) = N'" & v.Themengebiet.Trim().ToLower() & "';" & vbCrLf &
            "DECLARE @db nvarchar(128) = N'" & _datenbank & "';" & vbCrLf &
            "DECLARE @cols nvarchar(max), @sql nvarchar(max);" & vbCrLf & vbCrLf &
            "-- columns_dbo aus tm_polybase_struktur laden" & vbCrLf &
            "SELECT @cols = STRING_AGG(" & vbCrLf &
            "    CAST(columns_dbo AS nvarchar(max))," & vbCrLf &
            "    CONCAT(N',', CHAR(13), CHAR(10))" & vbCrLf &
            ") WITHIN GROUP (ORDER BY colno)" & vbCrLf &
            "FROM dbo.tm_polybase_struktur" & vbCrLf &
            "WHERE tabname = LOWER(@t)" & vbCrLf &
            "  AND themengebiet = @s;" & vbCrLf & vbCrLf &
            "-- Validierung" & vbCrLf &
            "IF @cols IS NULL" & vbCrLf &
            "    THROW 50001, 'Keine columns_dbo Metadaten gefunden', 1;" & vbCrLf & vbCrLf &
            "-- Template nur erstellen wenn noch nicht vorhanden" & vbCrLf &
            "IF OBJECT_ID(CONCAT('[', @db, '].dbo.[', @t, '_template]'), 'U') IS NULL" & vbCrLf &
            "BEGIN" & vbCrLf &
            "    SET @sql = N'SELECT TOP 0 ' + @cols +" & vbCrLf &
            "               N' INTO [' + @db + N'].dbo.[' + @t + N'_template]' +" & vbCrLf &
            "               N' FROM ext.[' + @t + N'];';" & vbCrLf &
            "    EXEC sp_executesql @sql;" & vbCrLf &
            "END"

        Log("  Template fuer: " & v.Faktentabelle.ToLower())
        SqlAusfuehren(connStr, sql, "Template erstellen")
        Log("  Template erstellt / bereits vorhanden ✓")
    End Sub

    ' =============================================================================
    ' VERFAHREN LADEN
    ' =============================================================================
    Private Function VerfahrenLaden(connStr As String) As List(Of VerfahrenInfo)
        Dim liste As New List(Of VerfahrenInfo)()
        Dim sql As String =
"SELECT a.ID, a.Verfahren, a.Themengebiet, a.LetzterSchritt,
        pf.Wert AS Faktentabelle
 FROM   dbo.ETL_Fkt_Arbeitsliste a
 JOIN   " & _parameterDB & ".dbo." & _parametertab & " pf
        ON pf.Verfahren = a.Verfahren
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

    ' =============================================================================
    ' STATUS MANAGEMENT
    ' =============================================================================
    Private Sub StatusSetzen(connStr As String, id As Integer, status As String)
        SqlAusfuehren(connStr,
            "UPDATE dbo.ETL_Fkt_Arbeitsliste SET Status='" & status & "', LetzterSchritt='" & status & "', AktualisiertAm=GETDATE() WHERE ID=" & id.ToString(),
            "Status setzen")
    End Sub

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

    ' =============================================================================
    ' LOGGING
    ' =============================================================================
    Private Sub LogSchreiben(connStr As String, verfahren As String, schritt As String, meldung As String)
        ' Kein DB-Log: Logging laeuft vollstaendig ueber SSIS Events (Eventhandler)
        ' FEHLER_* -> FireError | alles andere -> FireInformation
        If schritt.StartsWith("FEHLER", StringComparison.OrdinalIgnoreCase) Then
            LogFehler("[" & schritt & "] " & verfahren & ": " & meldung)
        Else
            Log("[" & schritt & "] " & verfahren & ": " & meldung)
        End If
    End Sub

    ' =============================================================================
    ' SQL HELPER FUNCTIONS
    ' =============================================================================
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
                If versuch < MAX_VERSUCHE Then
                    System.Threading.Thread.Sleep(WARTE_SEK * 1000)
                End If
            End Try
        End While
        Throw New Exception(String.Format("[{0}] fehlgeschlagen: {1}", beschreibung, If(letzterFehler IsNot Nothing, letzterFehler.Message, "Unbekannt")))
    End Function

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
                If versuch < MAX_VERSUCHE Then
                    System.Threading.Thread.Sleep(WARTE_SEK * 1000)
                Else
                    Throw
                End If
            End Try
        End While
        Return Nothing
    End Function

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

    ' =============================================================================
    ' DATA CLASSES
    ' =============================================================================
    Private Class VerfahrenInfo
        Public Property ID As Integer
        Public Property Verfahren As String
        Public Property Themengebiet As String
        Public Property LetzterSchritt As String
        Public Property Faktentabelle As String
    End Class

    Public Enum ScriptResults
        Success = DTSExecResult.Success
        Failure = DTSExecResult.Failure
    End Enum

End Class
