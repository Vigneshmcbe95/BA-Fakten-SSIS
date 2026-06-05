Option Explicit On
Option Strict On

Imports System
Imports System.Data.SqlClient
Imports Microsoft.SqlServer.Dts.Runtime

' =============================================================================
'  Script   : SCR_02_Vorbereitungen
'  Package  : Fakten Laden (SSIS)
'  Purpose  : One-time environment preparation: work list / error history
'             tables, PolyBase master key, credential, external data source,
'             ext schema, external DDL table and its local dbo copy.
'  Retry    : 3 attempts per SQL statement, 30 s delay
'  Logging  : SSIS events only (FireInformation / FireError)
' =============================================================================
<Microsoft.SqlServer.Dts.Tasks.ScriptTask.SSISScriptTaskEntryPointAttribute()>
<CLSCompliant(False)>
Partial Public Class ScriptMain
    Inherits Microsoft.SqlServer.Dts.Tasks.ScriptTask.VSTARTScriptObjectModelBase

    ' -------------------------------------------------------------------------
    ' Konstanten
    ' -------------------------------------------------------------------------
    Private Const SKRIPT_NAME As String = "SCR_02_Vorbereitungen"
    Private Const CONN_NAME As String = "Verbindung"
    Private Const MAX_VERSUCHE As Integer = 3
    Private Const WARTE_SEK As Integer = 30

    ' -------------------------------------------------------------------------
    ' SSIS-Variablen (werden einmalig geladen)
    ' -------------------------------------------------------------------------
    Private _server As String = String.Empty
    Private _datenbank As String = String.Empty
    Private _credBenutzer As String = String.Empty
    Private _credKennwort As String = String.Empty
    Private _extSourceName As String = String.Empty
    Private _extSourceLocation As String = String.Empty
    Private _extTabSchema As String = String.Empty
    Private _extTabDDLName As String = String.Empty
    Private _extTabDDLLocation As String = String.Empty
    Private _steuerlistenTabelle As String = String.Empty
    Private _credName As String = String.Empty

    ' -----------------------------------------------------------------------
    ' Main - Entry point - orchestrates the script flow.
    ' -----------------------------------------------------------------------
    Public Sub Main()

        Log("SCR_02_Vorbereitungen - Start")

        Try
            ' Variablen laden
            VariablenLaden()

            ' Pflichtfelder prüfen
            If Not PflichtfelderPruefen() Then
                Dts.TaskResult = ScriptResults.Failure
                Return
            End If

            Dim connStr As String = HoleVerbindungszeichenfolge()

            ' -- Schritt 1: ETL_Fkt_Arbeitsliste sicherstellen -------------------
            Log("Schritt 1: ETL_Fkt_Arbeitsliste sicherstellen")
            ArbeitslisteSicherstellen(connStr)

            ' -- Schritt 2: ETL_Fkt_FehlerHistorie sicherstellen -----------------
            Log("Schritt 2: ETL_Fkt_FehlerHistorie sicherstellen")
            FehlerHistorieSicherstellen(connStr)

            ' -- Schritt 3: PolyBase Master Key ------------------------------
            Log("Schritt 3: PolyBase Master Key pruefen")
            MasterKeyPruefen(connStr)

            ' -- Schritt 4: PolyBase Credential ------------------------------
            Log("Schritt 4: PolyBase Credential pruefen")
            CredentialPruefen(connStr)

            ' -- Schritt 5: External Data Source -----------------------------
            Log("Schritt 5: External Data Source pruefen")
            ExtDataSourcePruefen(connStr)

            ' -- Schritt 6: ext Schema ---------------------------------------
            Log("Schritt 6: Schema [" & _extTabSchema & "] pruefen")
            SchemaPruefen(connStr)

            ' -- Schritt 7: Externe DDL-Tabelle ------------------------------
            Log("Schritt 7: Externe DDL-Tabelle [" & _extTabSchema & "." & _extTabDDLName & "] pruefen")
            ExtDDLTabellePruefen(connStr)

            ' -- Schritt 8: Lokale DBO-Kopie der DDL-Tabelle erstellen -------
            Log("Schritt 8: Lokale dbo-Kopie [dbo." & _extTabDDLName & "] erstellen")
            DboKopieErstellen(connStr)

            Log("SCR_01_Vorbereitungen erfolgreich abgeschlossen.")

            Dts.TaskResult = ScriptResults.Success

        Catch ex As Exception
            LogFehler("Kritischer Fehler in " & SKRIPT_NAME & ": " & ex.Message)
            Dts.TaskResult = ScriptResults.Failure
        End Try

    End Sub

    ' -----------------------------------------------------------------------
    ' DboKopieErstellen - Creates / refreshes the local dbo copy of the
    ' Oracle DDL view.
    ' -----------------------------------------------------------------------
    Private Sub DboKopieErstellen(connStr As String)

        Dim zielTabelle As String = "dbo." & _extTabDDLName
        Dim quellTabelle As String = _extTabSchema & "." & _extTabDDLName

        ' Ziel-Tabelle löschen falls vorhanden
        Dim sqlDrop As String =
            "IF OBJECT_ID(N'" & zielTabelle & "', N'U') IS NOT NULL " &
            "    DROP TABLE " & zielTabelle & ";"
        SqlAusfuehren(connStr, sqlDrop, "dbo-Kopie löschen")
        Log("  Alte dbo-Kopie geloescht (falls vorhanden)")

        ' Neu befüllen per SELECT INTO — COLNAME lowercase für Vergleich in SCR04
        Dim sqlSelectInto As String =
            "SELECT " &
            "    LOWER(LTRIM(RTRIM(THMNAME)))  AS THMNAME, " &
            "    LOWER(LTRIM(RTRIM(TABNAME)))  AS TABNAME, " &
            "    LOWER(LTRIM(RTRIM(COLNAME)))  AS COLNAME, " &
            "    COLNO, IS_NULLABLE, COLLENGTH, PRECISION, SCALE, TYPNAME " &
            "INTO " & zielTabelle & " " &
            "FROM " & quellTabelle & ";"
        Dim zeilen As Integer = SqlAusfuehren(connStr, sqlSelectInto, "dbo-Kopie SELECT INTO")
        Log("  dbo-Kopie erstellt: " & zielTabelle & " | Zeilen: " & zeilen.ToString())

        ' Index auf TABNAME + COLNAME:
        '   → Non-Clustered Index — optimal für den Lookup in SCR04:
        '     WHERE TABNAME = @t AND COLNAME IN ('mow_id','monid')
        '   → TABNAME als führende Spalte (Equality-Prädikat)
        '   → COLNAME als zweite Spalte (IN-Prädikat)
        Dim idxName As String = "IX_" & _extTabDDLName & "_TABNAME_COLNAME"
        Dim sqlIndex As String =
            "CREATE NONCLUSTERED INDEX [" & idxName & "] " &
            "ON " & zielTabelle & " (TABNAME, COLNAME);"
        SqlAusfuehren(connStr, sqlIndex, "Index erstellen")
        Log("  Index erstellt: " & idxName & " (TABNAME, COLNAME)")

    End Sub

    ' -----------------------------------------------------------------------
    ' VariablenLaden - Reads the required SSIS variables into module
    ' fields.
    ' -----------------------------------------------------------------------
    Private Sub VariablenLaden()

        _server = Dts.Variables("BA::Server").Value.ToString().Trim()
        _datenbank = Dts.Variables("BA::Datenbank").Value.ToString().Trim()
        _credBenutzer = Dts.Variables("BA::CredBenutzername").Value.ToString().Trim()
        _credKennwort = Dts.Variables("BA::CredKennwort").Value.ToString().Trim()
        _extSourceName = Dts.Variables("BA::ExtSourceName").Value.ToString().Trim()
        _extSourceLocation = Dts.Variables("BA::ExtSourceLocation").Value.ToString().Trim()
        _extTabSchema = Dts.Variables("BA::ExtTableSchema").Value.ToString().Trim()
        _extTabDDLName = Dts.Variables("BA::ExtTableName").Value.ToString().Trim()
        _extTabDDLLocation = Dts.Variables("BA::ExtTableLocation").Value.ToString().Trim()
        _steuerlistenTabelle = Dts.Variables("BA::SteuerlistenTabelle").Value.ToString().Trim()
        _credName = _server & "_" & _credBenutzer

        Log("Steuerlisten-Tabelle: dbo." & _steuerlistenTabelle)

    End Sub

    ' -----------------------------------------------------------------------
    ' PflichtfelderPruefen - Validates that all mandatory variables /
    ' parameters are present.
    ' -----------------------------------------------------------------------
    Private Function PflichtfelderPruefen() As Boolean

        Dim fehlend As New System.Text.StringBuilder()

        If String.IsNullOrEmpty(_server) Then fehlend.AppendLine("  → BA::Server")
        If String.IsNullOrEmpty(_datenbank) Then fehlend.AppendLine("  → BA::Datenbank")
        If String.IsNullOrEmpty(_credBenutzer) Then fehlend.AppendLine("  → BA::CredBenutzername")
        If String.IsNullOrEmpty(_credKennwort) Then fehlend.AppendLine("  → BA::CredKennwort")
        If String.IsNullOrEmpty(_extSourceName) Then fehlend.AppendLine("  → BA::ExtSourceName")
        If String.IsNullOrEmpty(_extSourceLocation) Then fehlend.AppendLine("  → BA::ExtSourceLocation")
        If String.IsNullOrEmpty(_extTabSchema) Then fehlend.AppendLine("  → BA::ExtTableSchema")
        If String.IsNullOrEmpty(_extTabDDLName) Then fehlend.AppendLine("  → BA::ExtTableName")
        If String.IsNullOrEmpty(_extTabDDLLocation) Then fehlend.AppendLine("  → BA::ExtTableLocation")

        If fehlend.Length > 0 Then
            LogFehler("Pflichtfelder fehlen:" & Environment.NewLine & fehlend.ToString())
            Return False
        End If

        Log("Pflichtfelder-Pruefung: alle Variablen vorhanden ")
        Return True

    End Function

    ' -----------------------------------------------------------------------
    ' ArbeitslisteSicherstellen - Ensures dbo.ETL_Fkt_Arbeitsliste exists.
    ' -----------------------------------------------------------------------
    Private Sub ArbeitslisteSicherstellen(connStr As String)

        Dim sql As String =
"IF NOT EXISTS (
    SELECT 1 FROM sys.tables
    WHERE  name      = 'ETL_Fkt_Arbeitsliste'
    AND    schema_id = SCHEMA_ID('dbo')
)
BEGIN
    CREATE TABLE dbo.ETL_Fkt_Arbeitsliste
    (
        ID             INT           IDENTITY(1,1) PRIMARY KEY,
        RunID          INT           NULL,
        Verfahren      VARCHAR(200)  NOT NULL,
        Themengebiet   VARCHAR(200)  NULL,
        Status         VARCHAR(50)   NOT NULL DEFAULT 'AUSSTEHEND',
        LetzterSchritt VARCHAR(100)  NULL,
        Versuche       INT           NOT NULL DEFAULT 0,
        Fehlermeldung  NVARCHAR(4000) NULL,
        AktualisiertAm DATETIME      NOT NULL DEFAULT GETDATE()
    );
    PRINT 'ETL_Fkt_Arbeitsliste wurde neu angelegt.';
END
ELSE
    PRINT 'ETL_Fkt_Arbeitsliste bereits vorhanden.';"

        SqlAusfuehren(connStr, sql, "ETL_Fkt_Arbeitsliste sicherstellen")
        Log("ETL_Fkt_Arbeitsliste: geprueft/angelegt ")

    End Sub

    ' -----------------------------------------------------------------------
    ' FehlerHistorieSicherstellen - Ensures dbo.ETL_Fkt_FehlerHistorie
    ' exists.
    ' -----------------------------------------------------------------------
    Private Sub FehlerHistorieSicherstellen(connStr As String)

        Dim sql As String =
"IF NOT EXISTS (
    SELECT 1 FROM sys.tables
    WHERE  name      = 'ETL_Fkt_FehlerHistorie'
    AND    schema_id = SCHEMA_ID('dbo')
)
BEGIN
    CREATE TABLE dbo.ETL_Fkt_FehlerHistorie
    (
        ID                 INT            IDENTITY(1,1) PRIMARY KEY,
        Verfahren          VARCHAR(200)   NULL,
        Fehlerbeschreibung NVARCHAR(4000) NOT NULL,
        Fehlerzeit         DATETIME       NOT NULL DEFAULT GETDATE()
    );
    PRINT 'ETL_Fkt_FehlerHistorie wurde neu angelegt.';
END
ELSE
    PRINT 'ETL_Fkt_FehlerHistorie bereits vorhanden.';"

        SqlAusfuehren(connStr, sql, "ETL_Fkt_FehlerHistorie sicherstellen")
        Log("ETL_Fkt_FehlerHistorie: geprueft/angelegt ")

    End Sub

    ' -----------------------------------------------------------------------
    ' MasterKeyPruefen - Ensures the database master key exists.
    ' -----------------------------------------------------------------------
    Private Sub MasterKeyPruefen(connStr As String)

        Dim sqlPruefen As String =
"SELECT CASE
    WHEN EXISTS (
        SELECT 1 FROM sys.symmetric_keys
        WHERE  name = '##MS_DatabaseMasterKey##'
    ) THEN 1 ELSE 0 END"

        Dim vorhanden As Boolean =
            Convert.ToInt32(SqlSkalarAusfuehren(connStr, sqlPruefen, "Master Key prüfen")) = 1

        If vorhanden Then
            Log("Master Key: bereits vorhanden uebersprungen ")
            Return
        End If

        Log("Master Key: nicht vorhanden wird angelegt")
        Dim sqlErstellen As String =
            "CREATE MASTER KEY ENCRYPTION BY PASSWORD = '" & _credKennwort & "';"
        SqlAusfuehren(connStr, sqlErstellen, "Master Key anlegen")
        Log("Master Key: erfolgreich angelegt ")

    End Sub

    ' -----------------------------------------------------------------------
    ' CredentialPruefen - Ensures the database scoped credential exists.
    ' -----------------------------------------------------------------------
    Private Sub CredentialPruefen(connStr As String)

        Dim sqlPruefen As String =
"SELECT COUNT(*) FROM sys.database_scoped_credentials
 WHERE  name = '" & _credName & "'"

        Dim vorhanden As Boolean =
            Convert.ToInt32(SqlSkalarAusfuehren(connStr, sqlPruefen, "Credential prüfen")) > 0

        If vorhanden Then
            Log("Credential [" & _credName & "]: bereits vorhanden uebersprungen ")
            Return
        End If

        Log("Credential [" & _credName & "]: nicht vorhanden wird angelegt")
        Dim sqlErstellen As String =
"CREATE DATABASE SCOPED CREDENTIAL [" & _credName & "]
 WITH IDENTITY = '" & _credBenutzer & "',
      SECRET   = '" & _credKennwort & "';"
        SqlAusfuehren(connStr, sqlErstellen, "Credential anlegen")
        Log("Credential [" & _credName & "]: erfolgreich angelegt ")

    End Sub

    ' -----------------------------------------------------------------------
    ' ExtDataSourcePruefen - Ensures the external data source exists.
    ' -----------------------------------------------------------------------
    Private Sub ExtDataSourcePruefen(connStr As String)

        Dim sqlPruefen As String =
"SELECT COUNT(*) FROM sys.external_data_sources
 WHERE  name = '" & _extSourceName & "'"

        Dim vorhanden As Boolean =
            Convert.ToInt32(SqlSkalarAusfuehren(connStr, sqlPruefen, "Data Source prüfen")) > 0

        If vorhanden Then
            Log("External Data Source [" & _extSourceName & "]: bereits vorhanden uebersprungen ")
            Return
        End If

        Log("External Data Source [" & _extSourceName & "]: nicht vorhanden wird angelegt")
        Dim sqlErstellen As String =
"CREATE EXTERNAL DATA SOURCE [" & _extSourceName & "]
 WITH (
     LOCATION   = N'" & _extSourceLocation & "',
     CREDENTIAL = [" & _credName & "]
 );"
        SqlAusfuehren(connStr, sqlErstellen, "Data Source anlegen")
        Log("External Data Source [" & _extSourceName & "]: erfolgreich angelegt ")

    End Sub

    ' -----------------------------------------------------------------------
    ' SchemaPruefen - Ensures the ext schema exists.
    ' -----------------------------------------------------------------------
    Private Sub SchemaPruefen(connStr As String)

        Dim sqlPruefen As String =
"SELECT COUNT(*) FROM sys.schemas
 WHERE  name = '" & _extTabSchema & "'"

        Dim vorhanden As Boolean =
            Convert.ToInt32(SqlSkalarAusfuehren(connStr, sqlPruefen, "Schema prüfen")) > 0

        If vorhanden Then
            Log("Schema [" & _extTabSchema & "]: bereits vorhanden uebersprungen ")
            Return
        End If

        Log("Schema [" & _extTabSchema & "]: nicht vorhanden wird angelegt")
        SqlAusfuehren(connStr, "CREATE SCHEMA [" & _extTabSchema & "];", "Schema anlegen")
        Log("Schema [" & _extTabSchema & "]: erfolgreich angelegt ")

    End Sub

    ' -----------------------------------------------------------------------
    ' ExtDDLTabellePruefen - Ensures the external DDL table exists.
    ' -----------------------------------------------------------------------
    Private Sub ExtDDLTabellePruefen(connStr As String)

        Dim vollName As String = _extTabSchema & "." & _extTabDDLName

        Dim sqlPruefen As String =
"SELECT COUNT(*) FROM sys.external_tables
 WHERE  schema_id = SCHEMA_ID('" & _extTabSchema & "')
 AND    name      = '" & _extTabDDLName & "'"

        Dim vorhanden As Boolean =
            Convert.ToInt32(SqlSkalarAusfuehren(connStr, sqlPruefen, "DDL-Tabelle prüfen")) > 0

        If vorhanden Then
            Log("Externe DDL-Tabelle [" & vollName & "]: bereits vorhanden uebersprungen ")
            Return
        End If

        Log("Externe DDL-Tabelle [" & vollName & "]: nicht vorhanden wird angelegt")

        Dim sqlErstellen As String =
"CREATE EXTERNAL TABLE " & vollName & "
(
    THMNAME    NVARCHAR(128) COLLATE Latin1_General_100_CS_AS_SC_UTF8,
    TABNAME    NVARCHAR(128) COLLATE Latin1_General_100_CS_AS_SC_UTF8,
    COLNAME    NVARCHAR(128) COLLATE Latin1_General_100_CS_AS_SC_UTF8,
    COLNO      SMALLINT,
    IS_NULLABLE TINYINT,
    COLLENGTH  SMALLINT,
    PRECISION  SMALLINT,
    SCALE      SMALLINT,
    TYPNAME    NVARCHAR(128) COLLATE Latin1_General_100_CS_AS_SC_UTF8
)
WITH (
    DATA_SOURCE = [" & _extSourceName & "],
    LOCATION    = '" & _extTabDDLLocation & "'
);"

        SqlAusfuehren(connStr, sqlErstellen, "DDL-Tabelle anlegen")
        Log("Externe DDL-Tabelle [" & vollName & "]: erfolgreich angelegt ")

    End Sub

    ' -----------------------------------------------------------------------
    ' SqlAusfuehren - Executes a non-query SQL statement with retry; logs
    ' warning and the full SQL statement on failure.
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

    ' -----------------------------------------------------------------------
    ' SqlSkalarAusfuehren - Executes a scalar SQL query with retry.
    ' -----------------------------------------------------------------------
    Private Function SqlSkalarAusfuehren(connStr As String,
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
    Private Sub Log(nachricht As String)
        Dim fireAgain As Boolean = False
        Dts.Events.FireInformation(0, SKRIPT_NAME, nachricht, "", 0, fireAgain)
    End Sub

    ' -----------------------------------------------------------------------
    ' LogFehler - Writes an error message to the SSIS log (FireError).
    ' -----------------------------------------------------------------------
    Private Sub LogFehler(nachricht As String)
        Dts.Events.FireError(0, SKRIPT_NAME, nachricht, "", 0)
    End Sub

    ' -----------------------------------------------------------------------
    ' ScriptResults - SSIS task result codes.
    ' -----------------------------------------------------------------------
    Public Enum ScriptResults
        Success = DTSExecResult.Success
        Failure = DTSExecResult.Failure
    End Enum

End Class
