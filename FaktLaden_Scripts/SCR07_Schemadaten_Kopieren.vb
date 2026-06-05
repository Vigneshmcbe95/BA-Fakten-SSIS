Option Explicit On
Option Strict On

Imports System
Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports Microsoft.SqlServer.Dts.Runtime

' =============================================================================
'  Script   : SCR07_Schemadaten_Kopieren
'  Package  : Fakten Laden (SSIS)
'  Purpose  : Copies the Oracle DDL schema metadata from
'             ext.vm_ddl_sql_server into dbo.tm_polybase_struktur
'             (columns_dbo / columns_ext per fact table).
'  Workflow : AUSSTEHEND -> SCHEMADATEN_KOPIERT
'  Retry    : 3 attempts per SQL statement, 30 s delay
'  Logging  : SSIS events only (FireInformation / FireError)
' =============================================================================
<Microsoft.SqlServer.Dts.Tasks.ScriptTask.SSISScriptTaskEntryPointAttribute()>
<CLSCompliant(False)>
Partial Public Class ScriptMain
    Inherits Microsoft.SqlServer.Dts.Tasks.ScriptTask.VSTARTScriptObjectModelBase

    Private Const SKRIPT_NAME As String = "SCR07_Schemadaten_Kopieren"
    Private Const CONN_NAME As String = "Verbindung"
    Private Const MAX_VERSUCHE As Integer = 3
    Private Const WARTE_SEK As Integer = 30

    Private _runID As Integer = 0
    Private _datenbank As String = String.Empty
    Private _extTableSchema As String = String.Empty
    Private _extTableName As String = String.Empty

    ' -----------------------------------------------------------------------
    ' Main - Entry point - orchestrates the script flow.
    ' -----------------------------------------------------------------------
    Public Sub Main()

        Log("SCR07_Schemadaten_Kopieren - Start")

        Try
            _runID = Convert.ToInt32(Dts.Variables("BA::RunID").Value)
            _datenbank = Dts.Variables("BA::Datenbank").Value.ToString().Trim()
            _extTableSchema = Dts.Variables("BA::ExtTableSchema").Value.ToString().Trim()
            _extTableName = Dts.Variables("BA::ExtTableName").Value.ToString().Trim()


            Dim connStr As String = HoleVerbindungszeichenfolge()

            ' 1. Ziel-Tabelle sicherstellen
            EnsureSchemadataTableExists(connStr)

            ' 2. Verfahren aus Arbeitsliste laden
            Dim verfahren As List(Of VerfahrenInfo) = VerfahrenLaden(connStr)
            Log("Verfahren zur Verarbeitung: " & verfahren.Count.ToString())

            If verfahren.Count = 0 Then
                Log("Keine Verfahren zur Verarbeitung vorhanden.")
                Dts.TaskResult = ScriptResults.Success
                Return
            End If

            ' 3. Oracle-Daten einmalig in Staging laden
            StageOracleData(connStr)

            ' 4. Für jedes Verfahren: Daten aus Staging in tm_polybase_struktur kopieren
            Dim cntOK As Integer = 0
            Dim cntFehler As Integer = 0

            For Each v As VerfahrenInfo In verfahren
                Log("Verfahren: " & v.Verfahren & " | Themengebiet: " & v.Themengebiet)

                If v.LetzterSchritt = "SCHEMADATEN_KOPIERT" Then
                    Log("  bereits abgeschlossen uebersprungen OK")
                    Continue For
                End If

                Try
                    StatusSetzen(connStr, v.ID, "SCHEMADATEN_KOPIEREN")

                    Dim rows As Integer = InsertTabelle(connStr, v.Themengebiet, v.Verfahren)

                    If rows = 0 Then
                        Log("  WARNUNG: Keine Schemadaten fuer " & v.Themengebiet & "." & v.Verfahren)
                        FehlerSetzen(connStr, v.ID, "Keine Schemadaten in Oracle DDL gefunden")
                        LogSchreiben(connStr, v.Verfahren, "WARNUNG_SCR05",
                            "Keine Schemadaten in Oracle DDL gefunden")
                        cntFehler += 1
                    Else
                        StatusSetzen(connStr, v.ID, "SCHEMADATEN_KOPIERT")
                        LogSchreiben(connStr, v.Verfahren, "SCHRITT_5",
                            "Schemadaten kopiert: " & rows.ToString() & " Spalten")
                        cntOK += 1
                        Log("  " & rows.ToString() & " Spalten kopiert OK")
                    End If

                Catch ex As Exception
                    cntFehler += 1
                    FehlerSetzen(connStr, v.ID, ex.Message)
                    LogSchreiben(connStr, v.Verfahren, "FEHLER_SCR05", ex.Message)
                    LogFehler("FEHLER bei '" & v.Verfahren & "': " & ex.Message)
                End Try
            Next

            ' 5. STAGING BLEIBT ERHALTEN - nicht löschen!
            Log("Staging-Tabelle dbo.ddl_staging bleibt fuer Debugging erhalten")

            Log("Erfolgreich: " & cntOK.ToString() & " | Fehler: " & cntFehler.ToString())
            Dts.TaskResult = If(cntFehler > 0, ScriptResults.Failure, ScriptResults.Success)

        Catch ex As Exception
            LogFehler("Kritischer Fehler: " & ex.Message)
            Dts.TaskResult = ScriptResults.Failure
        End Try

    End Sub

    ' -----------------------------------------------------------------------
    ' StageOracleData - Loads the Oracle DDL data into the staging table.
    ' -----------------------------------------------------------------------
    Private Sub StageOracleData(connStr As String)
        Log("Oracle-Daten werden in Staging-Tabelle geladen...")
        Log("Quelle: [" & _datenbank & "].[" & _extTableSchema & "].[" & _extTableName & "]")

        ' Prüfen ob Staging bereits existiert
        Dim existsSQL As String =
            "SELECT COUNT(*) FROM sys.tables WHERE name = 'ddl_staging' AND schema_id = SCHEMA_ID('dbo');"

        Dim stagingExists As Integer = Convert.ToInt32(SqlSkalar(connStr, existsSQL, "Staging Check"))

        If stagingExists > 0 Then
            Dim rowCount As Integer = Convert.ToInt32(SqlSkalar(connStr,
                "SELECT COUNT(*) FROM dbo.ddl_staging;", "Staging Count"))

            If rowCount > 0 Then
                Log("Staging-Tabelle existiert bereits mit " & rowCount.ToString() & " Zeilen - wird wiederverwendet")
                Return
            Else
                Log("Staging-Tabelle existiert aber ist leer - wird neu geladen")
                SqlAusfuehren(connStr, "DROP TABLE dbo.ddl_staging;", "Staging DROP")
            End If
        End If

        Log("Oracle-Abfrage laeuft (kann 20-60 Sekunden dauern)...")

        Dim sqlSelect As String =
"SELECT
    CAST(RTRIM(THMNAME) AS VARCHAR(100)) COLLATE Latin1_General_100_CI_AS_SC_UTF8 AS THMNAME,
    CAST(RTRIM(TABNAME) AS VARCHAR(100)) COLLATE Latin1_General_100_CI_AS_SC_UTF8 AS TABNAME,
    COLNAME,
    COLNO,
    TYPNAME,
    COLLENGTH,
    PRECISION,
    SCALE,
    IS_NULLABLE
INTO dbo.ddl_staging
FROM [" & _datenbank & "].[" & _extTableSchema & "].[" & _extTableName & "];"

        SqlAusfuehren(connStr, sqlSelect, "Staging SELECT INTO")

        Dim sqlIndex As String =
            "CREATE INDEX ix_ddl_staging_thm_tab ON dbo.ddl_staging (THMNAME, TABNAME);"
        SqlAusfuehren(connStr, sqlIndex, "Staging INDEX")

        Dim rows As Integer = Convert.ToInt32(SqlSkalar(connStr,
            "SELECT COUNT(*) FROM dbo.ddl_staging;", "Staging COUNT"))

        Log("Staging abgeschlossen: " & rows.ToString() & " Zeilen geladen.")

        If rows = 0 Then
            Throw New Exception("Staging-Tabelle ist leer - Oracle-View hat keine Daten.")
        End If
    End Sub

    ' -----------------------------------------------------------------------
    ' InsertTabelle - Inserts a row into the target table.
    ' -----------------------------------------------------------------------
    Private Function InsertTabelle(connStr As String, thema As String, tab As String) As Integer
        ' DELETE old data
        Dim sqlDelete As String =
"DELETE FROM dbo.tm_polybase_struktur
WHERE themengebiet = @thema
  AND tabname = @tab;"

        Using conn As New SqlConnection(connStr)
            conn.Open()
            Using cmd As New SqlCommand(sqlDelete, conn)
                cmd.CommandTimeout = 0
                cmd.Parameters.AddWithValue("@thema", thema.Trim().ToLower())
                cmd.Parameters.AddWithValue("@tab", tab.Trim().ToLower())
                cmd.ExecuteNonQuery()
            End Using
        End Using

        ' INSERT new data with columns_dbo and columns_ext
        Dim sqlInsert As String =
"INSERT INTO dbo.tm_polybase_struktur
    (themengebiet, tabname, colname, colno, columns_dbo, columns_ext)
SELECT DISTINCT
    ddl.THMNAME,
    ddl.TABNAME,
    ddl.COLNAME,
    ddl.COLNO,
    -- columns_dbo: DBO-Tabellen Definition
    CONCAT(
        CHAR(9), LOWER(ddl.COLNAME), ' = ',
        CASE WHEN ddl.IS_NULLABLE = 0 THEN 'ISNULL(' ELSE '' END,
        CASE WHEN ddl.TYPNAME IN ('nvarchar','varchar','nchar','char')
             THEN CONCAT('CONVERT(', ddl.TYPNAME COLLATE Latin1_General_100_CI_AS_SC_UTF8,
                         '(', ddl.COLLENGTH, '), ')
             ELSE ''
        END,
        UPPER(ddl.COLNAME),
        CASE WHEN ddl.TYPNAME IN ('nvarchar','varchar','nchar','char')
             THEN ' COLLATE Latin1_General_100_CI_AS_SC_UTF8)'
             ELSE ''
        END,
        CASE WHEN ddl.IS_NULLABLE = 0 AND ddl.TYPNAME LIKE '%char%'
                THEN ', '''')'
             WHEN ddl.IS_NULLABLE = 0 AND (ddl.TYPNAME LIKE 'float%'
                  OR ddl.TYPNAME IN ('numeric','decimal')
                  OR ddl.TYPNAME LIKE '%int%')
                THEN ', 0)'
             WHEN ddl.IS_NULLABLE = 0 AND ddl.TYPNAME LIKE '%date%'
                THEN ', ''1900-01-01'')'
             ELSE ''
        END
    ),
    -- columns_ext: EXT-Tabellen Definition
    CONCAT(
        CHAR(9), UPPER(ddl.COLNAME), ' ', ddl.TYPNAME,
        CASE WHEN ddl.TYPNAME IN ('nvarchar','varchar','nchar','char')
             THEN CONCAT('(',
                    CASE WHEN ddl.COLLENGTH * 4 > 4000
                         THEN 4000
                         ELSE ddl.COLLENGTH * 4
                    END,
                    ') COLLATE Latin1_General_100_CS_AS_SC_UTF8')
             WHEN ddl.TYPNAME = 'varbinary'
                THEN CONCAT('(', ddl.COLLENGTH, ')')
             WHEN ddl.TYPNAME IN ('decimal','numeric')
                THEN CONCAT('(', ddl.PRECISION, ',', ddl.SCALE, ')')
             ELSE ''
        END,
        ' NULL'
    )
FROM dbo.ddl_staging ddl
WHERE ddl.THMNAME = @thema
  AND ddl.TABNAME = @tab;"

        Using conn As New SqlConnection(connStr)
            conn.Open()
            Using cmd As New SqlCommand(sqlInsert, conn)
                cmd.CommandTimeout = 0
                cmd.Parameters.AddWithValue("@thema", thema.Trim().ToLower())
                cmd.Parameters.AddWithValue("@tab", tab.Trim().ToLower())
                cmd.ExecuteNonQuery()
            End Using
        End Using

        ' COUNT check
        Dim sqlCount As String =
"SELECT COUNT(*)
FROM dbo.tm_polybase_struktur
WHERE themengebiet = @thema
  AND tabname = @tab;"

        Using conn As New SqlConnection(connStr)
            conn.Open()
            Using cmd As New SqlCommand(sqlCount, conn)
                cmd.CommandTimeout = 0
                cmd.Parameters.AddWithValue("@thema", thema.Trim().ToLower())
                cmd.Parameters.AddWithValue("@tab", tab.Trim().ToLower())
                Return Convert.ToInt32(cmd.ExecuteScalar())
            End Using
        End Using
    End Function

    ' -----------------------------------------------------------------------
    ' VerfahrenLaden - Loads the Verfahren to process from the work list
    ' (joined with the parameter table).
    ' -----------------------------------------------------------------------
    Private Function VerfahrenLaden(connStr As String) As List(Of VerfahrenInfo)
        Dim liste As New List(Of VerfahrenInfo)()
        Dim sql As String =
"SELECT ID, Verfahren, Themengebiet, LetzterSchritt
FROM dbo.ETL_Fkt_Arbeitsliste
WHERE Status IN ('AUSSTEHEND','SCHEMADATEN_KOPIEREN')
  AND RunID = " & _runID.ToString() & "
ORDER BY Verfahren"

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
                                    .LetzterSchritt = If(rdr.IsDBNull(3), "", rdr(3).ToString().Trim())
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
    ' EnsureSchemadataTableExists - Ensures dbo.tm_polybase_struktur
    ' exists.
    ' -----------------------------------------------------------------------
    Private Sub EnsureSchemadataTableExists(connStr As String)
        Dim sql As String =
"IF NOT EXISTS (
    SELECT 1 FROM sys.tables
    WHERE name = 'tm_polybase_struktur'
      AND schema_id = SCHEMA_ID('dbo')
)
BEGIN
    CREATE TABLE dbo.tm_polybase_struktur (
        themengebiet  VARCHAR(100),
        tabname       VARCHAR(100),
        colname       VARCHAR(128),
        colno         INT,
        columns_dbo   NVARCHAR(MAX),
        columns_ext   NVARCHAR(MAX),
        PRIMARY KEY (themengebiet, tabname, colname)
    );
END;"

        SqlAusfuehren(connStr, sql, "tm_polybase_struktur CREATE")
        Log("Ziel-Tabelle tm_polybase_struktur geprueft/erstellt.")
    End Sub

    ' -----------------------------------------------------------------------
    ' StatusSetzen - Updates Status / LetzterSchritt of a work list row.
    ' -----------------------------------------------------------------------
    Private Sub StatusSetzen(connStr As String, id As Integer, status As String)
        SqlAusfuehren(connStr,
            "UPDATE dbo.ETL_Fkt_Arbeitsliste SET Status='" & status & "', LetzterSchritt='" & status & "', AktualisiertAm=GETDATE() WHERE ID=" & id.ToString(),
            "Status setzen")
    End Sub

    ' -----------------------------------------------------------------------
    ' FehlerSetzen - Marks a work list row as FEHLER and stores the error
    ' message.
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
    ' LogSchreiben - Routes protocol messages to SSIS events: FEHLER_* ->
    ' FireError, everything else -> FireInformation.
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
    ' SqlAusfuehren - Executes a non-query SQL statement with retry; logs
    ' warning and the full SQL statement on failure.
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
    ' SqlSkalar - Executes a scalar SQL query with retry; logs warning and
    ' the full SQL statement on failure.
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
    ' HoleVerbindungszeichenfolge - Returns the connection string of the
    ' package connection manager.
    ' -----------------------------------------------------------------------
    Private Function HoleVerbindungszeichenfolge() As String
        Return Dts.Connections(CONN_NAME).ConnectionString
    End Function

    ' -----------------------------------------------------------------------
    ' Log - Writes an information message to the SSIS log
    ' (FireInformation).
    ' -----------------------------------------------------------------------
    Private Sub Log(n As String)
        Dim f As Boolean = False
        Dts.Events.FireInformation(0, SKRIPT_NAME, n, "", 0, f)
    End Sub

    ' -----------------------------------------------------------------------
    ' LogFehler - Writes an error message to the SSIS log (FireError).
    ' -----------------------------------------------------------------------
    Private Sub LogFehler(n As String)
        Dts.Events.FireError(0, SKRIPT_NAME, n, "", 0)
    End Sub

    ' -----------------------------------------------------------------------
    ' VerfahrenInfo - Data container for one Verfahren work item.
    ' -----------------------------------------------------------------------
    Private Class VerfahrenInfo
        Public Property ID As Integer
        Public Property Verfahren As String
        Public Property Themengebiet As String
        Public Property LetzterSchritt As String
    End Class

    ' -----------------------------------------------------------------------
    ' ScriptResults - SSIS task result codes.
    ' -----------------------------------------------------------------------
    Public Enum ScriptResults
        Success = DTSExecResult.Success
        Failure = DTSExecResult.Failure
    End Enum

End Class
