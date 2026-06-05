Option Explicit On
Option Strict On

Imports System
Imports System.Data.SqlClient
Imports Microsoft.SqlServer.Dts.Runtime

' =============================================================================
' PAKET  : Fakten Laden
' SKRIPT : SCR_PaketStart
' ZWECK  : 1. ETL_Fakt_LaufHistorie sicherstellen (anlegen falls nicht vorhanden)
'          2. Verwaiste LÃ¤ufe (Status = LAUFEND) â ABGEBROCHEN setzen
'          3. ETL_Arbeitsliste: FEHLER â AUSSTEHEND zurÃ¼cksetzen (Retry)
'          4. Neuen Lauf anlegen â BA::RunID setzen
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
    Private Const MAX_VERSUCHE As Integer = 10
    Private Const WARTE_SEK As Integer = 30

    ' -------------------------------------------------------------------------
    ' Einstiegspunkt
    ' -------------------------------------------------------------------------
    Public Sub Main()

        Log("ââââââââââââââââââââââââââââââââââââââââââââââââââââââââ")
        Log("Fakten Laden â Paketstart")
        Log("Zeitpunkt: " & DateTime.Now.ToString("dd.MM.yyyy HH:mm:ss"))
        Log("ââââââââââââââââââââââââââââââââââââââââââââââââââââââââ")

        Try
            Dim connStr As String = HoleVerbindungszeichenfolge()

            ' -- Schritt 1: ETL_Fakt_LaufHistorie sicherstellen -------------------
            Log("ââ Schritt 1: ETL_Fakt_LaufHistorie sicherstellen")
            LaufHistorieSicherstellen(connStr)

            ' -- Schritt 2: Verwaiste LÃ¤ufe abschliessen ---------------------
            Log("ââ Schritt 2: Verwaiste LÃ¤ufe (LAUFEND) auf ABGEBROCHEN setzen")
            VerwaistelaeufeAbschliessen(connStr)

            ' -- Schritt 3: Fehlgeschlagene Verfahren zurÃ¼cksetzen -----------
            Log("ââ Schritt 3: ETL_Fkt_ArbeitslisteFEHLER â AUSSTEHEND")
            FehlerZuruecksetzen(connStr)

            ' -- Schritt 4: Neuen Lauf anlegen -------------------------------
            Log("ââ Schritt 4: Neuen Lauf anlegen")
            Dim runID As Integer = NeuenLaufAnlegen(connStr)
            Dts.Variables("BA::RunID").Value = runID
            Log("RunID gesetzt: " & runID.ToString())

            Log("ââââââââââââââââââââââââââââââââââââââââââââââââââââââââ")
            Log("Paketstart erfolgreich abgeschlossen.")
            Log("ââââââââââââââââââââââââââââââââââââââââââââââââââââââââ")

            Dts.TaskResult = ScriptResults.Success

        Catch ex As Exception
            LogFehler("Kritischer Fehler in " & SKRIPT_NAME & ": " & ex.Message)
            Dts.TaskResult = ScriptResults.Failure
        End Try

    End Sub

    ' =========================================================================
    ' ETL_Fakt_LaufHistorie anlegen falls nicht vorhanden
    ' =========================================================================
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
        Log("ETL_Fakt_LaufHistorie: geprÃ¼ft/angelegt.")

    End Sub

    ' =========================================================================
    ' Verwaiste LÃ¤ufe (Status = LAUFEND) â ABGEBROCHEN
    ' Kein Zeitlimit â rein statusbasiert
    ' =========================================================================
    Private Sub VerwaistelaeufeAbschliessen(connStr As String)

        Dim sql As String =
"UPDATE dbo.ETL_Fakt_LaufHistorie
SET    RunStatus       = 'ABGEBROCHEN',
       Paketfehlerzeit = GETDATE()
WHERE  RunStatus = 'LAUFEND';"

        Dim betroffene As Integer = SqlAusfuehren(connStr, sql, "Verwaiste LÃ¤ufe abschliessen")
        Log("Verwaiste LÃ¤ufe auf ABGEBROCHEN gesetzt: " & betroffene.ToString())

    End Sub

    ' =========================================================================
    ' ETL_Arbeitsliste: FEHLER â AUSSTEHEND (fÃ¼r Retry-Versuche)
    ' ERFOLG-Zeilen bleiben unberÃ¼hrt (Resume-Logik)
    ' =========================================================================
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

        Dim betroffene As Integer = SqlAusfuehren(connStr, sql, "FEHLER â AUSSTEHEND")
        Log("Verfahren auf AUSSTEHEND zurÃ¼ckgesetzt: " & betroffene.ToString())

    End Sub

    ' =========================================================================
    ' Neuen Lauf in ETL_Fakt_LaufHistorie anlegen â RunID zurÃ¼ckgeben
    ' =========================================================================
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
