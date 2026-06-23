Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports Microsoft.SqlServer.Dts.Runtime

' =============================================================================
'  Skript       : PaketEnd
'  Paket        : Fakten Laden (SSIS)
'  Zweck        : Schliesst den Paketlauf ab: aggregiert die Statuszaehler
'                 der Arbeitsliste in ETL_Fakt_LaufHistorie und
'                 protokolliert die Zusammenfassung.
'  Protokoll    : Nur SSIS-Events (FireInformation / FireError)
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
    ' Main - Einstiegspunkt - steuert den Ablauf des Skripts.
    ' -----------------------------------------------------------------------
    Public Sub Main()
        Dim sqlConn As SqlConnection = Nothing
        Try
        Log("PaketEnd - Start")

            Dim runID As Integer = CInt(Dts.Variables("BA::RunID").Value)
            Dim parameterDB As String = Dts.Variables("BA::ParameterDB").Value.ToString().Trim()
            Dim parametertab As String = Dts.Variables("BA::Parametertabelle").Value.ToString().Trim()
            Dim stlTabelle As String = Dts.Variables("BA::SteuerlistenTabelle").Value.ToString().Trim()

            Dim cm As ConnectionManager = Dts.Connections(ConnectionName)
            Dim builder As New SqlConnectionStringBuilder(cm.ConnectionString)
            sqlConn = New SqlConnection(builder.ConnectionString)
            sqlConn.Open()

            Dim sql As String =
"SET NOCOUNT ON;
DECLARE @ID INT = @RunID, @Start DATETIME, @End DATETIME = GETDATE();

SELECT @Start = PaketStartzeit FROM dbo.ETL_Fakt_LaufHistorie WHERE ID = @ID;

-- Abschluss: Alle Verfahren dieses Laufs, die NICHT auf FEHLER stehen, auf ERFOLG setzen.
-- Ein Verfahren durchlaeuft alle Skripte (verarbeitet oder uebersprungen); am Paketende
-- gilt es als erfolgreich, sofern es nirgends einen Fehler gab.
UPDATE dbo.ETL_Fakt_Arbeitsliste
SET    Status = 'ERFOLG', LetzterSchritt = 'ERFOLG', AktualisiertAm = GETDATE()
WHERE  RunID = @ID AND Status <> 'FEHLER' AND Status <> 'ERFOLG';

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
FROM dbo.ETL_Fakt_Arbeitsliste
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

            ' Per-Tabelle-Zusammenfassung: was der Benutzer uebergeben hat (Filter
            ' bzw. Volllast) und was tatsaechlich in der Faktentabelle gefuellt ist.
            ZusammenfassungSchreiben(sqlConn, runID, parameterDB, parametertab, stlTabelle)

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
    ' ZusammenfassungSchreiben - Schreibt je Faktentabelle dieses Laufs eine
    ' Zusammenfassung in das SSIS-Protokoll:
    '   - Eingabe (User): partition_wert aus der Steuerliste (Filter) bzw.
    '     "VOLLLAST", wenn kein partition_wert gesetzt war.
    '   - Ergebnis: Status sowie Ist-Kennzahlen der Faktentabelle
    '     (Partitionen = DISTINCT Partitionsspalte, Zeilen, MIN, MAX).
    ' -----------------------------------------------------------------------
    Private Sub ZusammenfassungSchreiben(conn As SqlConnection, runID As Integer,
                                         parameterDB As String, parametertab As String,
                                         stlTabelle As String)

        Dim de As System.Globalization.CultureInfo = System.Globalization.CultureInfo.GetCultureInfo("de-DE")

        ' Verfahren des Laufs + Faktentabelle/Partitionsspalte (Parametertabelle)
        ' + vom Benutzer uebergebener partition_wert/Datei (Steuerliste).
        Dim liste As String =
            "SELECT a.Verfahren," &
            "       pf.Wert AS Faktentabelle," &
            "       pp.Wert AS PartCol," &
            "       a.Status," &
            "       MAX(stl.partition_wert) AS PartitionWert," &
            "       MAX(stl.FILE_NAME)      AS Datei," &
            "       MAX(a.Fehlermeldung)    AS Fehlermeldung" &
            " FROM dbo.ETL_Fakt_Arbeitsliste a" &
            " JOIN " & parameterDB & ".dbo." & parametertab & " pf ON pf.Verfahren=dbo.fn_ParamVerfahren(a.Verfahren) AND pf.Parameter='Faktentabelle'" &
            " JOIN " & parameterDB & ".dbo." & parametertab & " pp ON pp.Verfahren=dbo.fn_ParamVerfahren(a.Verfahren) AND pp.Parameter='Faktenpartitionsspalte'" &
            " LEFT JOIN dbo." & stlTabelle & " stl ON LOWER(LTRIM(RTRIM(stl.tabelle)))=LOWER(LTRIM(RTRIM(a.Verfahren)))" &
            " WHERE a.RunID=" & runID &
            " GROUP BY a.Verfahren, pf.Wert, pp.Wert, a.Status" &
            " ORDER BY pf.Wert"

        Dim rows As New List(Of String())()
        Using cmd As New SqlCommand(liste, conn)
            cmd.CommandTimeout = 0
            Using rdr As SqlDataReader = cmd.ExecuteReader()
                While rdr.Read()
                    rows.Add(New String() {
                        rdr(0).ToString().Trim(),
                        rdr(1).ToString().Trim(),
                        rdr(2).ToString().Trim(),
                        rdr(3).ToString().Trim(),
                        If(rdr.IsDBNull(4), "", rdr(4).ToString().Trim()),
                        If(rdr.IsDBNull(5), "", rdr(5).ToString().Trim()),
                        If(rdr.IsDBNull(6), "", rdr(6).ToString().Trim())})
                End While
            End Using
        End Using

        Log("")
        Log("=== ZUSAMMENFASSUNG (RunID " & runID.ToString() & ") ===")

        Dim cntErfolg As Integer = 0
        Dim cntFehler As Integer = 0

        For Each r As String() In rows
            Dim fakt As String = r(1)
            Dim partCol As String = r(2)
            If partCol.Contains("|") Then partCol = partCol.Substring(0, partCol.IndexOf("|"))
            Dim status As String = r(3)
            Dim pw As String = r(4)
            Dim datei As String = If(r(5) = "", "-", r(5))
            Dim fehlermeldung As String = r(6)

            If status = "ERFOLG" Then cntErfolg += 1
            If status = "FEHLER" Then cntFehler += 1

            Dim eingabe As String
            If String.IsNullOrEmpty(pw) Then
                eingabe = "(kein partition_wert) -> VOLLLAST (alle Partitionen)"
            Else
                eingabe = "partition_wert = " & pw & "   -> FILTER"
            End If

            ' Ist-Kennzahlen der Faktentabelle (was tatsaechlich gefuellt ist)
            Dim tabExists As Boolean = False
            Dim zeilen As Long = 0
            Dim parts As Integer = 0
            Dim mn As String = "NULL"
            Dim mx As String = "NULL"

            Dim q As String =
                "IF OBJECT_ID('dbo.[" & fakt & "]','U') IS NOT NULL" &
                "  SELECT 1 AS ex, COUNT_BIG(*) AS c, COUNT(DISTINCT [" & partCol & "]) AS p," &
                "         CONVERT(varchar(20),MIN([" & partCol & "])) AS mn," &
                "         CONVERT(varchar(20),MAX([" & partCol & "])) AS mx" &
                "  FROM dbo.[" & fakt & "]" &
                " ELSE SELECT 0 AS ex, CAST(0 AS bigint) AS c, 0 AS p," &
                "            CAST(NULL AS varchar(20)) AS mn, CAST(NULL AS varchar(20)) AS mx"

            Try
                Using cmd As New SqlCommand(q, conn)
                    cmd.CommandTimeout = 0
                    Using rdr As SqlDataReader = cmd.ExecuteReader()
                        If rdr.Read() Then
                            tabExists = Convert.ToInt32(rdr("ex")) = 1
                            If Not rdr.IsDBNull(rdr.GetOrdinal("c")) Then zeilen = Convert.ToInt64(rdr("c"))
                            If Not rdr.IsDBNull(rdr.GetOrdinal("p")) Then parts = Convert.ToInt32(rdr("p"))
                            If Not rdr.IsDBNull(rdr.GetOrdinal("mn")) Then mn = rdr("mn").ToString()
                            If Not rdr.IsDBNull(rdr.GetOrdinal("mx")) Then mx = rdr("mx").ToString()
                        End If
                    End Using
                End Using
            Catch ex As Exception
                Log("  WARNUNG: Kennzahlen fuer dbo." & fakt & " nicht lesbar: " & ex.Message)
            End Try

            Log("Tabelle: " & fakt & "   | Datei: " & datei)
            Log("  Eingabe (User) : " & eingabe)
            If tabExists Then
                Log("  Ergebnis       : Status=" & status &
                    " | Partitionen=" & parts.ToString() &
                    " | Zeilen=" & zeilen.ToString("#,##0", de) &
                    " | MIN=" & mn & " MAX=" & mx)
            Else
                Log("  Ergebnis       : Status=" & status & " | Faktentabelle nicht vorhanden")
            End If
            ' Bei Fehler die Ursache mitschreiben (z.B. fehlgeschlagene Partitionen,
            ' die der naechste Lauf neu laedt).
            If status = "FEHLER" AndAlso fehlermeldung <> "" Then
                Log("  Fehler         : " & fehlermeldung)
            End If
        Next

        Log("")
        Log("GESAMT: " & rows.Count.ToString() & " Tabellen | ERFOLG=" & cntErfolg.ToString() &
            " | FEHLER=" & cntFehler.ToString())
    End Sub

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

End Class
