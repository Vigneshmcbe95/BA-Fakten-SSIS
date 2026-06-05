Option Explicit On
Option Strict On

Imports System
Imports System.Data.SqlClient
Imports Microsoft.SqlServer.Dts.Runtime

' =============================================================================
' PAKET  : Fakten Laden
' SKRIPT : SCR_02_Vorbereitungen
' ZWECK  : 1. ETL_Fkt_Arbeitsliste sicherstellen
'          2. ETL_Fkt_FehlerHistorie sicherstellen
'          3. PolyBase Master Key prÃ¼fen / anlegen
'          4. PolyBase Credential prÃ¼fen / anlegen
'          5. PolyBase External Data Source prÃ¼fen / anlegen
'          6. ext Schema prÃ¼fen / anlegen
'          7. Externe DDL-Tabelle (vm_ddl_sql_server) prÃ¼fen / anlegen
' =============================================================================
<Microsoft.SqlServer.Dts.Tasks.ScriptTask.SSISScriptTaskEntryPointAttribute()>
<CLSCompliant(False)>
Partial Public Class ScriptMain
    Inherits Microsoft.SqlServer.Dts.Tasks.ScriptTask.VSTARTScriptObjectModelBase

    ' -------------------------------------------------------------------------
    ' Konstanten
    ' -------------------------------------------------------------------------
    Private Const SKRIPT_NAME As String = "SCR_01_Vorbereitungen"
    Private Const CONN_NAME As String = "Verbindung"
    Private Const MAX_VERSUCHE As Integer = 10
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

    ' -------------------------------------------------------------------------
    ' Einstiegspunkt
    ' -------------------------------------------------------------------------
    Public Sub Main()

        Log("ââââââââââââââââââââââââââââââââââââââââââââââââââââââââ")
        Log("SCR_01_Vorbereitungen â Start")
        Log("Zeitpunkt: " & DateTime.Now.ToString("dd.MM.yyyy HH:mm:ss"))
        Log("ââââââââââââââââââââââââââââââââââââââââââââââââââââââââ")

        Try
            ' Variablen laden
            VariablenLaden()

            ' Pflichtfelder prÃ¼fen
            If Not PflichtfelderPruefen() Then
                Dts.TaskResult = ScriptResults.Failure
                Return
            End If

            Dim connStr As String = HoleVerbindungszeichenfolge()

            ' -- Schritt 1: ETL_Fkt_Arbeitsliste sicherstellen -------------------
            Log("ââ Schritt 1: ETL_Fkt_Arbeitsliste sicherstellen")
            ArbeitslisteSicherstellen(connStr)

            ' -- Schritt 2: ETL_Fkt_FehlerHistorie sicherstellen -----------------
            Log("ââ Schritt 2: ETL_Fkt_FehlerHistorie sicherstellen")
            FehlerHistorieSicherstellen(connStr)

            ' -- Schritt 3: PolyBase Master Key ------------------------------
            Log("ââ Schritt 3: PolyBase Master Key prÃ¼fen")
            MasterKeyPruefen(connStr)

            ' -- Schritt 4: PolyBase Credential ------------------------------
            Log("ââ Schritt 4: PolyBase Credential prÃ¼fen")
            CredentialPruefen(connStr)

            ' -- Schritt 5: External Data Source -----------------------------
            Log("ââ Schritt 5: External Data Source prÃ¼fen")
            ExtDataSourcePruefen(connStr)

            ' -- Schritt 6: ext Schema ---------------------------------------
            Log("ââ Schritt 6: Schema [" & _extTabSchema & "] prÃ¼fen")
            SchemaPruefen(connStr)

            ' -- Schritt 7: Externe DDL-Tabelle ------------------------------
            Log("ââ Schritt 7: Externe DDL-Tabelle [" & _extTabSchema & "." & _extTabDDLName & "] prÃ¼fen")
            ExtDDLTabellePruefen(connStr)

            ' -- Schritt 8: Lokale DBO-Kopie der DDL-Tabelle erstellen -------
            Log("ââ Schritt 8: Lokale dbo-Kopie [dbo." & _extTabDDLName & "] erstellen")
            DboKopieErstellen(connStr)

            Log("ââââââââââââââââââââââââââââââââââââââââââââââââââââââââ")
            Log("SCR_01_Vorbereitungen erfolgreich abgeschlossen.")
            Log("ââââââââââââââââââââââââââââââââââââââââââââââââââââââââ")

            Dts.TaskResult = ScriptResults.Success

        Catch ex As Exception
            LogFehler("Kritischer Fehler in " & SKRIPT_NAME & ": " & ex.Message)
            Dts.TaskResult = ScriptResults.Failure
        End Try

    End Sub

    ' =========================================================================
    ' Schritt 8: Lokale dbo-Kopie der ext DDL-Tabelle erstellen
    '            SELECT INTO dbo.<extTabDDLName> FROM ext.<extTabDDLName>
    '            Wird bei jedem Lauf neu erstellt (DROP + SELECT INTO)
    '            â dient als Basis fÃ¼r where_klausel-Erkennung in SCR04
    '            â Index auf TABNAME + COLNAME fÃ¼r schnellen Lookup in SCR04
    ' =========================================================================
    Private Sub DboKopieErstellen(connStr As String)

        Dim zielTabelle As String = "dbo." & _extTabDDLName
        Dim quellTabelle As String = _extTabSchema & "." & _extTabDDLName

        ' Ziel-Tabelle lÃ¶schen falls vorhanden
        Dim sqlDrop As String =
            "IF OBJECT_ID(N'" & zielTabelle & "', N'U') IS NOT NULL " &
            "    DROP TABLE " & zielTabelle & ";"
        SqlAusfuehren(connStr, sqlDrop, "dbo-Kopie lÃ¶schen")
        Log("  Alte dbo-Kopie gelÃ¶scht (falls vorhanden)")

        ' Neu befÃ¼llen per SELECT INTO â COLNAME lowercase fÃ¼r Vergleich in SCR04
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
        '   â Non-Clustered Index â optimal fÃ¼r den Lookup in SCR04:
        '     WHERE TABNAME = @t AND COLNAME IN ('mow_id','monid')
        '   â TABNAME als fÃ¼hrende Spalte (Equality-PrÃ¤dikat)
        '   â COLNAME als zweite Spalte (IN-PrÃ¤dikat)
        Dim idxName As String = "IX_" & _extTabDDLName & "_TABNAME_COLNAME"
        Dim sqlIndex As String =
            "CREATE NONCLUSTERED INDEX [" & idxName & "] " &
            "ON " & zielTabelle & " (TABNAME, COLNAME);"
        SqlAusfuehren(connStr, sqlIndex, "Index erstellen")
        Log("  Index erstellt: " & idxName & " (TABNAME, COLNAME)")

    End Sub

    ' =========================================================================
    ' Variablen aus SSIS laden
    ' =========================================================================
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

        Log("ââââââââââââââââââââââââââââââââââââââââââââââââââââââââ")
        Log("Server              : " & _server)
        Log("Datenbank           : " & _datenbank)
        Log("Credential-Name     : " & _credName)
        Log("Ext. Source Name    : " & _extSourceName)
        Log("Ext. Source Loc     : " & _extSourceLocation)
        Log("Ext. DDL Tabelle    : " & _extTabSchema & "." & _extTabDDLName)
        Log("Ext. DDL Location   : " & _extTabDDLLocation)
        Log("Steuerlisten-Tabelle: dbo." & _steuerlistenTabelle)
        Log("ââââââââââââââââââââââââââââââââââââââââââââââââââââââââ")

    End Sub

    ' =========================================================================
    ' Pflichtfelder prÃ¼fen
    ' =========================================================================
    Private Function PflichtfelderPruefen() As Boolean

        Dim fehlend As New System.Text.StringBuilder()

        If String.IsNullOrEmpty(_server) Then fehlend.AppendLine("  â BA::Server")
        If String.IsNullOrEmpty(_datenbank) Then fehlend.AppendLine("  â BA::Datenbank")
        If String.IsNullOrEmpty(_credBenutzer) Then fehlend.AppendLine("  â BA::CredBenutzername")
        If String.IsNullOrEmpty(_credKennwort) Then fehlend.AppendLine("  â BA::CredKennwort")
        If String.IsNullOrEmpty(_extSourceName) Then fehlend.AppendLine("  â BA::ExtSourceName")
        If String.IsNullOrEmpty(_extSourceLocation) Then fehlend.AppendLine("  â BA::ExtSourceLocation")
        If String.IsNullOrEmpty(_extTabSchema) Then fehlend.AppendLine("  â BA::ExtTableSchema")
        If String.IsNullOrEmpty(_extTabDDLName) Then fehlend.AppendLine("  â BA::ExtTableName")
        If String.IsNullOrEmpty(_extTabDDLLocation) Then fehlend.AppendLine("  â BA::ExtTableLocation")

        If fehlend.Length > 0 Then
            LogFehler("Pflichtfelder fehlen:" & Environment.NewLine & fehlend.ToString())
            Return False
        End If

        Log("Pflichtfelder-PrÃ¼fung: alle Variablen vorhanden ")
        Return True

    End Function

    ' =========================================================================
    ' ETL_Fkt_Arbeitsliste sicherstellen
    ' =========================================================================
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
        Log("ETL_Fkt_Arbeitsliste: geprÃ¼ft/angelegt ")

    End Sub

    ' =========================================================================
    ' ETL_Fkt_FehlerHistorie sicherstellen
    ' =========================================================================
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
        Log("ETL_Fkt_FehlerHistorie: geprÃ¼ft/angelegt ")

    End Sub

    ' =========================================================================
    ' PolyBase Master Key prÃ¼fen / anlegen
    ' =========================================================================
    Private Sub MasterKeyPruefen(connStr As String)

        Dim sqlPruefen As String =
"SELECT CASE
    WHEN EXISTS (
        SELECT 1 FROM sys.symmetric_keys
        WHERE  name = '##MS_DatabaseMasterKey##'
    ) THEN 1 ELSE 0 END"

        Dim vorhanden As Boolean =
            Convert.ToInt32(SqlSkalarAusfuehren(connStr, sqlPruefen, "Master Key prÃ¼fen")) = 1

        If vorhanden Then
            Log("Master Key: bereits vorhanden â Ã¼bersprungen ")
            Return
        End If

        Log("Master Key: nicht vorhanden â wird angelegt")
        Dim sqlErstellen As String =
            "CREATE MASTER KEY ENCRYPTION BY PASSWORD = '" & _credKennwort & "';"
        SqlAusfuehren(connStr, sqlErstellen, "Master Key anlegen")
        Log("Master Key: erfolgreich angelegt ")

    End Sub

    ' =========================================================================
    ' PolyBase Credential prÃ¼fen / anlegen
    ' =========================================================================
    Private Sub CredentialPruefen(connStr As String)

        Dim sqlPruefen As String =
"SELECT COUNT(*) FROM sys.database_scoped_credentials
 WHERE  name = '" & _credName & "'"

        Dim vorhanden As Boolean =
            Convert.ToInt32(SqlSkalarAusfuehren(connStr, sqlPruefen, "Credential prÃ¼fen")) > 0

        If vorhanden Then
            Log("Credential [" & _credName & "]: bereits vorhanden â Ã¼bersprungen ")
            Return
        End If

        Log("Credential [" & _credName & "]: nicht vorhanden â wird angelegt")
        Dim sqlErstellen As String =
"CREATE DATABASE SCOPED CREDENTIAL [" & _credName & "]
 WITH IDENTITY = '" & _credBenutzer & "',
      SECRET   = '" & _credKennwort & "';"
        SqlAusfuehren(connStr, sqlErstellen, "Credential anlegen")
        Log("Credential [" & _credName & "]: erfolgreich angelegt ")

    End Sub

    ' =========================================================================
    ' External Data Source prÃ¼fen / anlegen
    ' =========================================================================
    Private Sub ExtDataSourcePruefen(connStr As String)

        Dim sqlPruefen As String =
"SELECT COUNT(*) FROM sys.external_data_sources
 WHERE  name = '" & _extSourceName & "'"

        Dim vorhanden As Boolean =
            Convert.ToInt32(SqlSkalarAusfuehren(connStr, sqlPruefen, "Data Source prÃ¼fen")) > 0

        If vorhanden Then
            Log("External Data Source [" & _extSourceName & "]: bereits vorhanden â Ã¼bersprungen ")
            Return
        End If

        Log("External Data Source [" & _extSourceName & "]: nicht vorhanden â wird angelegt")
        Dim sqlErstellen As String =
"CREATE EXTERNAL DATA SOURCE [" & _extSourceName & "]
 WITH (
     LOCATION   = N'" & _extSourceLocation & "',
     CREDENTIAL = [" & _credName & "]
 );"
        SqlAusfuehren(connStr, sqlErstellen, "Data Source anlegen")
        Log("External Data Source [" & _extSourceName & "]: erfolgreich angelegt ")

    End Sub

    ' =========================================================================
    ' ext Schema prÃ¼fen / anlegen
    ' =========================================================================
    Private Sub SchemaPruefen(connStr As String)

        Dim sqlPruefen As String =
"SELECT COUNT(*) FROM sys.schemas
 WHERE  name = '" & _extTabSchema & "'"

        Dim vorhanden As Boolean =
            Convert.ToInt32(SqlSkalarAusfuehren(connStr, sqlPruefen, "Schema prÃ¼fen")) > 0

        If vorhanden Then
            Log("Schema [" & _extTabSchema & "]: bereits vorhanden â Ã¼bersprungen ")
            Return
        End If

        Log("Schema [" & _extTabSchema & "]: nicht vorhanden â wird angelegt")
        SqlAusfuehren(connStr, "CREATE SCHEMA [" & _extTabSchema & "];", "Schema anlegen")
        Log("Schema [" & _extTabSchema & "]: erfolgreich angelegt ")

    End Sub

    ' =========================================================================
    ' Externe DDL-Tabelle (vm_ddl_sql_server) prÃ¼fen / anlegen
    ' =========================================================================
    Private Sub ExtDDLTabellePruefen(connStr As String)

        Dim vollName As String = _extTabSchema & "." & _extTabDDLName

        Dim sqlPruefen As String =
"SELECT COUNT(*) FROM sys.external_tables
 WHERE  schema_id = SCHEMA_ID('" & _extTabSchema & "')
 AND    name      = '" & _extTabDDLName & "'"

        Dim vorhanden As Boolean =
            Convert.ToInt32(SqlSkalarAusfuehren(connStr, sqlPruefen, "DDL-Tabelle prÃ¼fen")) > 0

        If vorhanden Then
            Log("Externe DDL-Tabelle [" & vollName & "]: bereits vorhanden â Ã¼bersprungen ")
            Return
        End If

        Log("Externe DDL-Tabelle [" & vollName & "]: nicht vorhanden â wird angelegt")

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

    ' =========================================================================
    ' SQL-Helfer: NonQuery mit Retry â gibt betroffene Zeilen zurÃ¼ck
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
    ' SQL-Helfer: Scalar mit Retry â gibt einzelnen Wert zurÃ¼ck
    ' =========================================================================
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
