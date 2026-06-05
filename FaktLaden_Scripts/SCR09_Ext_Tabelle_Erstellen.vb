Option Explicit On
Option Strict On

Imports System
Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports Microsoft.SqlServer.Dts.Runtime

' =============================================================================
' PAKET  : Fakten Laden
' SKRIPT : SCR09_Ext_Tabelle_Erstellen (v2 - using tm_polybase_struktur)
' ZWECK  : Pro Verfahren: externe Tabelle aus tm_polybase_struktur.columns_ext erstellen
'          Oracle LOCATION → UPPER CASE
'          Status: TEMPLATE_ERSTELLT → EXT_TABELLE_ERSTELLT
' =============================================================================
<Microsoft.SqlServer.Dts.Tasks.ScriptTask.SSISScriptTaskEntryPointAttribute()>
<CLSCompliant(False)>
Partial Public Class ScriptMain
    Inherits Microsoft.SqlServer.Dts.Tasks.ScriptTask.VSTARTScriptObjectModelBase

    Private Const SKRIPT_NAME As String = "SCR09_Ext_Tabelle_Erstellen"
    Private Const CONN_NAME As String = "Verbindung"
    Private Const MAX_VERSUCHE As Integer = 10
    Private Const WARTE_SEK As Integer = 30

    Private _runID As Integer = 0
    Private _parameterDB As String = String.Empty
    Private _parametertab As String = String.Empty
    Private _extSchema As String = String.Empty
    Private _extSourceName As String = String.Empty
    Private _extTableLocation As String = String.Empty

    Public Sub Main()
        Log("════════════════════════════════════════════════════════")
        Log("SCR07_Ext_Tabelle_Erstellen – Start (v2: tm_polybase_struktur)")
        Log("Zeitpunkt: " & DateTime.Now.ToString("dd.MM.yyyy HH:mm:ss"))
        Log("════════════════════════════════════════════════════════")

        Try
            _runID = Convert.ToInt32(Dts.Variables("BA::RunID").Value)
            _parameterDB = Dts.Variables("BA::ParameterDB").Value.ToString().Trim()
            _parametertab = Dts.Variables("BA::Parametertabelle").Value.ToString().Trim()
            _extSchema = Dts.Variables("BA::ExtTableSchema").Value.ToString().Trim()
            _extSourceName = Dts.Variables("BA::ExtSourceName").Value.ToString().Trim()
            _extTableLocation = Dts.Variables("BA::ExtTableLocation").Value.ToString().Trim()

            Log("RunID              : " & _runID.ToString())
            Log("ParameterDB        : " & _parameterDB)
            Log("Parametertabelle   : " & _parametertab)
            Log("ExtSchema          : " & _extSchema)
            Log("ExtSourceName      : " & _extSourceName)
            Log("ExtTableLocation   : " & _extTableLocation)

            Dim connStr As String = HoleVerbindungszeichenfolge()
            Dim verfahren As List(Of VerfahrenInfo) = VerfahrenLaden(connStr)

            Log("Verfahren zur Verarbeitung: " & verfahren.Count.ToString())

            Dim cntOK As Integer = 0
            Dim cntFehler As Integer = 0

            For Each v As VerfahrenInfo In verfahren
                Log("────────────────────────────────────────────────────────")
                Log("Verfahren: " & v.Verfahren & " | Themengebiet: " & v.Themengebiet)

                If v.LetzterSchritt = "EXT_TABELLE_ERSTELLT" Then
                    Log("  → bereits abgeschlossen → Ã¼bersprungen ✓")
                    Continue For
                End If

                Try
                    StatusSetzen(connStr, v.ID, "EXT_TABELLE_ERSTELLEN")
                    ExtTabelleErstellen(connStr, v)
                    StatusSetzen(connStr, v.ID, "EXT_TABELLE_ERSTELLT")
                    LogSchreiben(connStr, v.Verfahren, "SCHRITT_2",
                        "Externe Tabelle erstellt: " & _extSchema & "." & v.Faktentabelle.ToLower())
                    cntOK += 1
                    Log("  → Externe Tabelle erstellt ✓")
                Catch ex As Exception
                    cntFehler += 1
                    FehlerSetzen(connStr, v.ID, ex.Message)
                    LogSchreiben(connStr, v.Verfahren, "FEHLER_SCR07", ex.Message)
                    LogFehler("FEHLER '" & v.Verfahren & "': " & ex.Message)
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
    ' EXTERNE TABELLE ERSTELLEN - using columns_ext from tm_polybase_struktur
    ' =============================================================================
    Private Sub ExtTabelleErstellen(connStr As String, v As VerfahrenInfo)
        Dim extName As String = v.Faktentabelle.ToLower()
        Dim extFullName As String = _extSchema & ".[" & extName & "]"

        Log("  Erstelle externe Tabelle: " & extFullName)

        ' Extract Oracle environment from ExtTableLocation
        ' Example: "ISTAT.STATRT.VM_DDL_SQL_SERVER" → "ISTAT"
        Dim oracleEnvironment As String = _extTableLocation.Split("."c)(0).ToUpper()

        ' Oracle LOCATION: <ENVIRONMENT>.<THEMENGEBIET>.<FAKTENTABELLE> (all UPPER)
        Dim location As String = oracleEnvironment & "." & v.Themengebiet.ToUpper() & "." & v.Faktentabelle.ToUpper()

        Log("  Oracle Location: " & location)

        ' Build dynamic SQL using columns_ext directly from tm_polybase_struktur
        Dim sqlBuild As String =
"DECLARE @drop nvarchar(max), @crt nvarchar(max);

SELECT
    @drop = CONCAT(N'IF EXISTS(SELECT 1 FROM sys.external_tables WHERE schema_id=SCHEMA_ID(''', 
                   '" & _extSchema & "', ''') AND name=''', '" & extName & "', ''') ',
                   'DROP EXTERNAL TABLE " & extFullName & ";'),
    @crt  = CONCAT(N'CREATE EXTERNAL TABLE " & extFullName & " (', CHAR(13), CHAR(10),
                   STRING_AGG(CAST(columns_ext AS nvarchar(max)), CONCAT(N',',CHAR(13),CHAR(10))) WITHIN GROUP (ORDER BY colno),
                   CHAR(13), CHAR(10), N') WITH (DATA_SOURCE=[" & _extSourceName & "], LOCATION=''', 
                   '" & location & "', ''');')
FROM dbo.tm_polybase_struktur
WHERE themengebiet COLLATE Latin1_General_100_CI_AS_SC_UTF8 = @thema COLLATE Latin1_General_100_CI_AS_SC_UTF8
  AND tabname COLLATE Latin1_General_100_CI_AS_SC_UTF8 = @tab COLLATE Latin1_General_100_CI_AS_SC_UTF8;

IF @crt IS NULL 
BEGIN
    DECLARE @msg nvarchar(500);
    SET @msg = 'Keine Schemadaten in tm_polybase_struktur für Theme=' + @thema + ', Tabelle=' + @tab;
    THROW 50002, @msg, 1;
END

EXEC(@drop);
EXEC(@crt);"

        ' Execute the dynamic SQL
        Dim versuch As Integer = 0
        While versuch < MAX_VERSUCHE
            versuch += 1
            Try
                Using conn As New SqlConnection(connStr)
                    conn.Open()
                    Using cmd As New SqlCommand(sqlBuild, conn)
                        cmd.CommandTimeout = 0
                        cmd.Parameters.AddWithValue("@thema", v.Themengebiet.Trim().ToLower())
                        cmd.Parameters.AddWithValue("@tab", v.Verfahren.Trim().ToLower())
                        cmd.ExecuteNonQuery()
                    End Using
                End Using
                Exit While
            Catch ex As Exception
                Log(String.Format("WARNUNG [Ext Tabelle erstellen] Versuch {0}/{1}: {2}", versuch, MAX_VERSUCHE, ex.Message))
                If versuch < MAX_VERSUCHE Then
                    System.Threading.Thread.Sleep(WARTE_SEK * 1000)
                Else
                    Throw
                End If
            End Try
        End While

        ' Validate
        Dim checkSQL As String =
"SELECT COUNT(*) 
FROM sys.external_tables 
WHERE schema_id = SCHEMA_ID('" & _extSchema & "') 
  AND name = '" & extName & "';"

        Dim exists As Integer = Convert.ToInt32(SqlSkalar(connStr, checkSQL, "Ext Tabelle prÃ¼fen"))

        If exists = 0 Then
            Throw New Exception("Externe Tabelle wurde nicht erstellt!")
        End If

        Log("  ✓ Externe Tabelle erfolgreich: " & extFullName)
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
 WHERE  a.Status IN ('TEMPLATE_ERSTELLT','EXT_TABELLE_ERSTELLEN')
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
            "UPDATE dbo.ETL_Fkt_Arbeitsliste SET Status='" & status & "',LetzterSchritt='" & status & "',AktualisiertAm=GETDATE() WHERE ID=" & id.ToString(),
            "Status setzen")
    End Sub

    Private Sub FehlerSetzen(connStr As String, id As Integer, msg As String)
        Dim kurz As String = If(msg.Length > 3900, msg.Substring(0, 3900), msg)
        Try
            Using conn As New SqlConnection(connStr)
                conn.Open()
                Using cmd As New SqlCommand(
                    "UPDATE dbo.ETL_Fkt_Arbeitsliste SET Status='FEHLER',Fehlermeldung=@m,AktualisiertAm=GETDATE() WHERE ID=@id", conn)
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
                Using cmd As New SqlCommand(
                    "INSERT INTO dbo.tm_fakten_load_log(verfahren,schritt,meldung) VALUES(@v,@s,@m)", conn)
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

        Throw New Exception(String.Format("[{0}] fehlgeschlagen: {1}", beschreibung,
            If(letzterFehler IsNot Nothing, letzterFehler.Message, "Unbekannt")))
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
