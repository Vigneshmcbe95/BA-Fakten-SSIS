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
    ' TEMPLATE ERSTELLEN - Direct from Oracle external table
    ' =============================================================================
    Private Sub TemplateErstellen(connStr As String, v As VerfahrenInfo)
        Dim tabelle As String = v.Faktentabelle.ToLower()
        ' Use _datenbank parameter for the template location
        Dim ziel As String = "[" & _datenbank & "].dbo." & tabelle & "_template"

        Log("  Prüfe/Erstelle Template: " & ziel)

        ' 1. Check if template already exists
        Dim exists As Integer = Convert.ToInt32(SqlSkalar(connStr,
            "SELECT COUNT(*) FROM [" & _datenbank & "].sys.tables WHERE schema_id=SCHEMA_ID('dbo') AND name='" & tabelle & "_template'",
            "Template Check"))

        If exists > 0 Then
            Log("  Template existiert bereits. Vergleiche Spaltenliste mit tm_polybase_struktur...")
            
            ' Validierung: Prüfen ob die Business-Spalten im Template mit den 'colname' Einträgen in der Metadaten-Tabelle übereinstimmen
            Dim sqlCompare As String = "
            WITH NewSchema AS (
                SELECT STRING_AGG(UPPER(LTRIM(RTRIM(colname))), '|') WITHIN GROUP (ORDER BY colno) as ColStr
                FROM dbo.tm_polybase_struktur
                WHERE themengebiet = @thema AND tabname = @tab
            ),
            OldSchema AS (
                SELECT STRING_AGG(UPPER(LTRIM(RTRIM(COLUMN_NAME))), '|') WITHIN GROUP (ORDER BY ORDINAL_POSITION) as ColStr
                FROM [" & _datenbank & "].INFORMATION_SCHEMA.COLUMNS
                WHERE TABLE_SCHEMA = 'dbo' AND TABLE_NAME = '" & tabelle & "_template'
            )
            SELECT CASE WHEN n.ColStr = o.ColStr THEN 1 ELSE 0 END
            FROM NewSchema n, OldSchema o"

            Dim match As Integer = 0
            Try
                Using conn As New SqlConnection(connStr)
                    conn.Open()
                    Using cmd As New SqlCommand(sqlCompare, conn)
                        cmd.Parameters.AddWithValue("@thema", v.Themengebiet.Trim().ToLower())
                        cmd.Parameters.AddWithValue("@tab", v.Verfahren.Trim().ToLower())
                        Dim res = cmd.ExecuteScalar()
                        match = If(res IsNot Nothing AndAlso res IsNot DBNull.Value, Convert.ToInt32(res), 0)
                    End Using
                End Using
            Catch ex As Exception
                Log("  Warnung beim Strukturvergleich: " & ex.Message & " -> Sicherhaltshalber Neuanlage empfohlen.")
                match = 0
            End Try

            If match = 1 Then
                Log("  ✓ Spaltenliste identisch. Bestehendes Template wird wiederverwendet.")
                Return
            Else
                LogFehler("  !!! STRUKTUR-MISSMATCH !!!")
                LogFehler("  Die Spalten im bestehenden Template passen nicht zur Definition in tm_polybase_struktur.")
                LogFehler("  Bitte Template-Tabelle löschen, damit sie neu erstellt werden kann.")
                Throw New Exception("Struktur-Konflikt in " & ziel)
            End If
        End If

        ' 2. Create Template from tm_polybase_struktur if not exists
        Log("  Erstelle neues Template aus tm_polybase_struktur...")
        
        Dim sqlCreate As String = "
        SELECT 
            themengebiet,
            tabname,
            colname,
            colno,
            columns_dbo,
            columns_ext
        INTO " & ziel & "
        FROM dbo.tm_polybase_struktur
        WHERE themengebiet = @thema AND tabname = @tab
        ORDER BY colno"

        Using conn As New SqlConnection(connStr)
            conn.Open()
            Using cmd As New SqlCommand(sqlCreate, conn)
                cmd.Parameters.AddWithValue("@thema", v.Themengebiet.Trim().ToLower())
                cmd.Parameters.AddWithValue("@tab", v.Verfahren.Trim().ToLower())
                cmd.ExecuteNonQuery()
            End Using
        End Using

        ' Validate
        Dim cnt As Integer = Convert.ToInt32(SqlSkalar(connStr,
            "SELECT COUNT(*) FROM " & ziel,
            "Template Count"))

        If cnt = 0 Then
            Throw New Exception("Keine Daten in tm_polybase_struktur für " & v.Verfahren & " gefunden!")
        End If

        Log("  ✓ Template neu erstellt: " & ziel & " | Spalten: " & cnt.ToString())
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
        Try
            Using conn As New SqlConnection(connStr)
                conn.Open()
                Using cmd As New SqlCommand("INSERT INTO dbo.tm_fakten_load_log(verfahren,schritt,meldung) VALUES(@v,@s,@m)", conn)
                    cmd.Parameters.AddWithValue("@v", verfahren)
                    cmd.Parameters.AddWithValue("@s", schritt)
                    cmd.Parameters.AddWithValue("@m", meldung)
                    cmd.ExecuteNonQuery()
                End Using
            End Using
        Catch
        End Try
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
