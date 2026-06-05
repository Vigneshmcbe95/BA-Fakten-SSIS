Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports Microsoft.SqlServer.Dts.Runtime

' =============================================================================
'  Script   : PaketEnd
'  Package  : Fakten Laden (SSIS)
'  Purpose  : Finalizes the package run: aggregates the work list status
'             counters into ETL_Fakt_LaufHistorie and logs the run summary.
'  Logging  : SSIS events only (FireInformation / FireError)
' =============================================================================
<Microsoft.SqlServer.Dts.Tasks.ScriptTask.SSISScriptTaskEntryPointAttribute()>
<CLSCompliant(False)>
Partial Public Class ScriptMain
    Inherits Microsoft.SqlServer.Dts.Tasks.ScriptTask.VSTARTScriptObjectModelBase

    Private Const SKRIPT_NAME As String = "PaketEnd"
    Private ReadOnly ConnectionName As String = "Verbindung"

    Enum ScriptResults
        Success = DTSExecResult.Success
        Failure = DTSExecResult.Failure
    End Enum

    ' -----------------------------------------------------------------------
    ' Main - Entry point - orchestrates the script flow.
    ' -----------------------------------------------------------------------
    Public Sub Main()
        Dim sqlConn As SqlConnection = Nothing
        Try
        Log("PaketEnd - Start")

            Dim runID As Integer = CInt(Dts.Variables("BA::RunID").Value)

            Dim cm As ConnectionManager = Dts.Connections(ConnectionName)
            Dim builder As New SqlConnectionStringBuilder(cm.ConnectionString)
            sqlConn = New SqlConnection(builder.ConnectionString)
            sqlConn.Open()

            Dim sql As String =
"SET NOCOUNT ON;
DECLARE @ID INT = @RunID, @Start DATETIME, @End DATETIME = GETDATE();

SELECT @Start = PaketStartzeit FROM dbo.ETL_Fakt_LaufHistorie WHERE ID = @ID;

DECLARE
    @GesamtAnzahl              INT,
    @ErfolgreichAnzahl         INT,
    @FehlerAnzahl              INT,
    @AnzahlTemplateErstellt    INT,
    @AnzahlExtTabelle          INT,
    @AnzahlFaktentabelle       INT,
    @AnzahlPartitionsgrenzen   INT,
    @AnzahlStagingErstellt     INT,
    @AnzahlDatenGeladen        INT,
    @AnzahlIndexInOut          INT,
    @AnzahlKomprimierung       INT,
    @AnzahlNcciOut             INT,
    @AnzahlPartitionstausch    INT;

SELECT
    @GesamtAnzahl            = COUNT(*),
    @ErfolgreichAnzahl       = SUM(CASE WHEN Status = 'ERFOLG'                     THEN 1 ELSE 0 END),
    @FehlerAnzahl            = SUM(CASE WHEN Status = 'FEHLER'                     THEN 1 ELSE 0 END),
    @AnzahlTemplateErstellt  = SUM(CASE WHEN Status = 'TEMPLATE_ERSTELLT'          THEN 1 ELSE 0 END),
    @AnzahlExtTabelle        = SUM(CASE WHEN Status = 'EXT_TABELLE_ERSTELLT'       THEN 1 ELSE 0 END),
    @AnzahlFaktentabelle     = SUM(CASE WHEN Status = 'FAKTENTABELLE_ERSTELLT'     THEN 1 ELSE 0 END),
    @AnzahlPartitionsgrenzen = SUM(CASE WHEN Status = 'PARTITIONSGRENZEN_ERSTELLT' THEN 1 ELSE 0 END),
    @AnzahlStagingErstellt   = SUM(CASE WHEN Status = 'STAGING_ERSTELLT'           THEN 1 ELSE 0 END),
    @AnzahlDatenGeladen      = SUM(CASE WHEN Status = 'DATEN_GELADEN'              THEN 1 ELSE 0 END),
    @AnzahlIndexInOut        = SUM(CASE WHEN Status = 'INDEX_IN_OUT'               THEN 1 ELSE 0 END),
    @AnzahlKomprimierung     = SUM(CASE WHEN Status = 'KOMPRIMIERUNG_OUT'          THEN 1 ELSE 0 END),
    @AnzahlNcciOut           = SUM(CASE WHEN Status = 'NCCI_OUT'                   THEN 1 ELSE 0 END),
    @AnzahlPartitionstausch  = SUM(CASE WHEN Status = 'PARTITIONSTAUSCH'           THEN 1 ELSE 0 END)
FROM dbo.ETL_Fkt_Arbeitsliste
WHERE RunID = @ID;

UPDATE dbo.ETL_Fakt_LaufHistorie
SET
    RunStatus                   = CASE WHEN ISNULL(@FehlerAnzahl, 0) > 0 THEN 'FEHLER' ELSE 'ERFOLG' END,
    PaketEndzeit                = @End,
    PaketLaufgestamzeitSekunden = DATEDIFF(SECOND, @Start, @End),
    GesamtAnzahl                = @GesamtAnzahl,
    ErfolgreichAnzahl           = @ErfolgreichAnzahl,
    FehlerAnzahl                = @FehlerAnzahl,
    AnzahlTemplateErstellt      = @AnzahlTemplateErstellt,
    AnzahlExtTabelle            = @AnzahlExtTabelle,
    AnzahlFaktentabelle         = @AnzahlFaktentabelle,
    AnzahlPartitionsgrenzen     = @AnzahlPartitionsgrenzen,
    AnzahlStagingErstellt       = @AnzahlStagingErstellt,
    AnzahlDatenGeladen          = @AnzahlDatenGeladen,
    AnzahlIndexInOut            = @AnzahlIndexInOut,
    AnzahlKomprimierung         = @AnzahlKomprimierung,
    AnzahlNcciOut               = @AnzahlNcciOut,
    AnzahlPartitionstausch      = @AnzahlPartitionstausch
WHERE ID = @ID;

SELECT
    ISNULL(@GesamtAnzahl, 0)        AS GesamtAnzahl,
    ISNULL(@ErfolgreichAnzahl, 0)   AS ErfolgreichAnzahl,
    ISNULL(@FehlerAnzahl, 0)        AS FehlerAnzahl,
    DATEDIFF(SECOND, @Start, @End)  AS LaufzeitSekunden;"

            Using cmd As New SqlCommand(sql, sqlConn)
                cmd.CommandTimeout = 0
                cmd.Parameters.Add("@RunID", SqlDbType.Int).Value = runID
                Using rdr As SqlDataReader = cmd.ExecuteReader()
                    If rdr.Read() Then
                        Log("LaufHistorie aktualisiert OK")
                        Log("Gesamt      : " & rdr.GetInt32(0).ToString())
                        Log("Erfolgreich : " & rdr.GetInt32(1).ToString())
                        Log("Fehler      : " & rdr.GetInt32(2).ToString())
                        Log("Laufzeit    : " & If(rdr.IsDBNull(3), "unbekannt", rdr.GetInt32(3).ToString() & " Sekunden"))
                    End If
                End Using
            End Using

            Log("PaketEnd abgeschlossen OK")
            Dts.TaskResult = ScriptResults.Success

        Catch ex As Exception
            LogFehler("PaketEnd FEHLER: " & ex.Message)
            Dts.TaskResult = ScriptResults.Failure
        Finally
            If sqlConn IsNot Nothing AndAlso sqlConn.State <> ConnectionState.Closed Then
                sqlConn.Close()
            End If
        End Try
    End Sub

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

End Class
