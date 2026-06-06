Option Explicit On
Option Strict On

Imports System
Imports System.Data.SqlClient
Imports Microsoft.SqlServer.Dts.Runtime

' =============================================================================
'  Skript       : SCR_PaketStart
'  Paket        : Fakten Laden (SSIS)
'  Zweck        : Initialisiert den Paketlauf: stellt ETL_Fakt_LaufHistorie
'                 sicher, schliesst verwaiste Laeufe ab und legt einen neuen
'                 Lauf an (setzt BA::RunID).
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
    Private Const SKRIPT_NAME As String = "SCR_PaketStart"
    Private Const CONN_NAME As String = "Verbindung"
    Private Const MAX_VERSUCHE As Integer = 3
    Private Const WARTE_SEK As Integer = 30

    ' -----------------------------------------------------------------------
    ' Main - Einstiegspunkt - steuert den Ablauf des Skripts.
    ' -----------------------------------------------------------------------
    Public Sub Main()

        Log("Fakten Laden - Paketstart")

        Try
            Dim connStr As String = HoleVerbindungszeichenfolge()

            ' -- Schritt 1: ETL_Fakt_LaufHistorie sicherstellen -------------------
            Log("Schritt 1: ETL_Fakt_LaufHistorie sicherstellen")
            LaufHistorieSicherstellen(connStr)

            ' -- Schritt 2: Verwaiste Läufe abschliessen ---------------------
            Log("Schritt 2: Verwaiste Laeufe (LAUFEND) auf ABGEBROCHEN setzen")
            VerwaistelaeufeAbschliessen(connStr)

            ' -- Schritt 3: Fehlgeschlagene Verfahren zurücksetzen -----------
            Log("Schritt 3: ETL_Fkt_ArbeitslisteFEHLER AUSSTEHEND")
            FehlerZuruecksetzen(connStr)

            ' -- Schritt 4: Neuen Lauf anlegen -------------------------------
            Log("Schritt 4: Neuen Lauf anlegen")
            Dim runID As Integer = NeuenLaufAnlegen(connStr)
            Dts.Variables("BA::RunID").Value = runID

            Log("Paketstart erfolgreich abgeschlossen.")

            Dts.TaskResult = ScriptResults.Success

        Catch ex As Exception
            LogFehler("Kritischer Fehler in " & SKRIPT_NAME & ": " & ex.Message)
            Dts.TaskResult = ScriptResults.Failure
        End Try

    End Sub

    ' -----------------------------------------------------------------------
    ' LaufHistorieSicherstellen - Stellt sicher, dass
    ' dbo.ETL_Fakt_LaufHistorie existiert.
    ' -----------------------------------------------------------------------
    Private Sub LaufHistorieSicherstellen(connStr As String)

        Dim sql As String =
"IF NOT EXISTS (
    SELECT 1 FROM sys.tables
    WHERE  name      = 'ETL_Fakt_LaufHistorie'
    AND    schema_id = SCHEMA_ID('dbo')
)
BEGIN
    CREATE TABLE dbo.ETL_Fakt_LaufHistorie
    (
        ID                          INT           IDENTITY(1,1) PRIMARY KEY,
        RunStatus                   VARCHAR(20)   NOT NULL DEFAULT 'LAUFEND',
        PaketStartzeit              DATETIME      NOT NULL DEFAULT GETDATE(),
        PaketEndzeit                DATETIME      NULL,
        Paketfehlerzeit             DATETIME      NULL,
        PaketLaufgestamzeitSekunden INT           NULL,
        Hostname                    VARCHAR(200)  NULL,
        PackageName                 VARCHAR(200)  NULL,
        GesamtAnzahl                INT           NULL,
        ErfolgreichAnzahl           INT           NULL,
        FehlerAnzahl                INT           NULL,
        AnzahlTemplateErstellt      INT           NULL,
        AnzahlExtTabelle            INT           NULL,
        AnzahlFaktentabelle         INT           NULL,
        AnzahlPartitionsgrenzen     INT           NULL,
        AnzahlStagingErstellt       INT           NULL,
        AnzahlDatenGeladen          INT           NULL,
        AnzahlIndexInOut            INT           NULL,
        AnzahlKomprimierung         INT           NULL,
        AnzahlNcciOut               INT           NULL,
        AnzahlPartitionstausch      INT           NULL
    );
    PRINT 'ETL_Fakt_LaufHistorie wurde neu angelegt.';
END
ELSE
    PRINT 'ETL_Fakt_LaufHistorie bereits vorhanden.';"

        SqlAusfuehren(connStr, sql, "ETL_Fakt_LaufHistorie sicherstellen")
        Log("ETL_Fakt_LaufHistorie: geprueft/angelegt.")

    End Sub

    ' -----------------------------------------------------------------------
    ' VerwaistelaeufeAbschliessen - Schliesst verwaiste Laeufe ab (LAUFEND
    ' -> ABGEBROCHEN).
    ' -----------------------------------------------------------------------
    Private Sub VerwaistelaeufeAbschliessen(connStr As String)

        Dim sql As String =
"UPDATE dbo.ETL_Fakt_LaufHistorie
SET    RunStatus       = 'ABGEBROCHEN',
       Paketfehlerzeit = GETDATE()
WHERE  RunStatus = 'LAUFEND';"

        Dim betroffene As Integer = SqlAusfuehren(connStr, sql, "Verwaiste Läufe abschliessen")
        Log("Verwaiste Laeufe auf ABGEBROCHEN gesetzt: " & betroffene.ToString())

    End Sub

    ' -----------------------------------------------------------------------
    ' FehlerZuruecksetzen - Setzt FEHLER-Zeilen fuer den neuen Lauf auf
    ' AUSSTEHEND zurueck.
    ' -----------------------------------------------------------------------
    Private Sub FehlerZuruecksetzen(connStr As String)

        Dim sql As String =
"IF EXISTS (
    SELECT 1 FROM sys.tables
    WHERE  name      = 'ETL_Arbeitsliste'
    AND    schema_id = SCHEMA_ID('dbo')
)
BEGIN
    UPDATE dbo.ETL_Arbeitsliste
    SET    Status        = 'AUSSTEHEND',
           Fehlermeldung = NULL,
           AktualisiertAm = GETDATE()
    WHERE  Status = 'FEHLER';
END;"

        Dim betroffene As Integer = SqlAusfuehren(connStr, sql, "FEHLER → AUSSTEHEND")
        Log("Verfahren auf AUSSTEHEND zurueckgesetzt: " & betroffene.ToString())

    End Sub

    ' -----------------------------------------------------------------------
    ' NeuenLaufAnlegen - Legt einen neuen Lauf an und liefert die RunID
    ' zurueck.
    ' -----------------------------------------------------------------------
    Private Function NeuenLaufAnlegen(connStr As String) As Integer

        Dim sql As String =
"INSERT INTO dbo.ETL_Fakt_LaufHistorie
       (RunStatus, PaketStartzeit, Hostname, PackageName)
VALUES ('LAUFEND', GETDATE(), HOST_NAME(), APP_NAME());
SELECT SCOPE_IDENTITY();"

        Dim result As Object = SqlSkalarAusfuehren(connStr, sql, "Neuen Lauf anlegen")
        If result Is Nothing OrElse result Is DBNull.Value Then
            Throw New Exception("RunID konnte nicht ermittelt werden.")
        End If

        Return Convert.ToInt32(result)

    End Function

    ' -----------------------------------------------------------------------
    ' SqlAusfuehren - Fuehrt eine SQL-Anweisung (Non-Query) mit
    ' Wiederholung aus; protokolliert Warnung und vollstaendiges
    ' SQL-Statement bei Fehlern.
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
                    beschreibung, versuch, MAX_VERSUCHE, ex.Message))
                If versuch = 1 Then Log("SQL Statement [" & beschreibung & "]: " & sql)
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
    ' SqlSkalarAusfuehren - Fuehrt eine skalare SQL-Abfrage mit
    ' Wiederholung aus.
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
    Private Sub Log(nachricht As String)
        Dim fireAgain As Boolean = False
        Dts.Events.FireInformation(0, SKRIPT_NAME, nachricht, "", 0, fireAgain)
    End Sub

    ' -----------------------------------------------------------------------
    ' LogFehler - Schreibt eine Fehlermeldung in das SSIS-Protokoll
    ' (FireError).
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
