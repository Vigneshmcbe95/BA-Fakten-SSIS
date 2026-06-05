Option Explicit On
Option Strict On

Imports System
Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports Microsoft.SqlServer.Dts.Runtime

' =============================================================================
'  Script   : SCR12_Staging_Erstellen
'  Package  : Fakten Laden (SSIS)
'  Purpose  : Creates one empty _out_ staging shell per partition value
'             (SELECT TOP 0 via the template column list).
'  Workflow : PARTITIONSGRENZEN_ERSTELLT -> STAGING_ERSTELLT
'  Retry    : 3 attempts per SQL statement, 30 s delay
'  Logging  : SSIS events only (FireInformation / FireError)
' =============================================================================
<Microsoft.SqlServer.Dts.Tasks.ScriptTask.SSISScriptTaskEntryPointAttribute()>
<CLSCompliant(False)>
Partial Public Class ScriptMain
    Inherits Microsoft.SqlServer.Dts.Tasks.ScriptTask.VSTARTScriptObjectModelBase

    Private Const SKRIPT_NAME As String = "SCR12_Staging_Erstellen"
    Private Const CONN_NAME As String = "Verbindung"
    Private Const MAX_VERSUCHE As Integer = 3
    Private Const WARTE_SEK As Integer = 30

    Private _runID As Integer = 0
    Private _parameterDB As String = String.Empty
    Private _parametertab As String = String.Empty
    Private _extTableSchema As String = String.Empty
    Private _datenbank As String = String.Empty

    ' -----------------------------------------------------------------------
    ' Main - Entry point - orchestrates the script flow.
    ' -----------------------------------------------------------------------
    Public Sub Main()

        Log("SCR12_Staging_Erstellen - Start")

        Try
            _runID = Convert.ToInt32(Dts.Variables("BA::RunID").Value)
            _parameterDB = Dts.Variables("BA::ParameterDB").Value.ToString().Trim()
            _parametertab = Dts.Variables("BA::Parametertabelle").Value.ToString().Trim()
            _extTableSchema = Dts.Variables("BA::ExtTableSchema").Value.ToString().Trim()
            _datenbank = Dts.Variables("BA::Datenbank").Value.ToString().Trim()


            ' ─────────────────────────────────────────────────────
            ' 1. Partitionswerte aus BA::objPartitionValues lesen
            ' ─────────────────────────────────────────────────────
            Dim partObjekt As Object = Dts.Variables("BA::objPartitionValues").Value
            If partObjekt Is Nothing Then
                Log("BA::objPartitionValues ist leer (Nothing) keine Staging-Tabellen zu erstellen")
                Dts.TaskResult = ScriptResults.Success
                Return
            End If

            Dim partArray(,) As String = CType(partObjekt, String(,))
            Dim anzahlEintraege As Integer = partArray.GetLength(0)
            Log("Partitionswerte aus SCR09: " & anzahlEintraege.ToString() & " Eintraege")

            ' Partitionswerte nach Verfahren gruppieren
            Dim verfahrenWerte As New Dictionary(Of String, List(Of PartitionsEintrag))()
            For i As Integer = 0 To anzahlEintraege - 1
                Dim verf As String = partArray(i, 0)
                Dim wert As String = partArray(i, 1)
                Dim modus As String = partArray(i, 2)
                If Not verfahrenWerte.ContainsKey(verf) Then
                    verfahrenWerte(verf) = New List(Of PartitionsEintrag)()
                End If
                verfahrenWerte(verf).Add(New PartitionsEintrag With {.Wert = wert, .Modus = modus})
            Next

            Dim connStr As String = HoleVerbindungszeichenfolge()
            Dim verfahren As List(Of VerfahrenInfo) = VerfahrenLaden(connStr)
            Log("Verfahren aus Arbeitsliste: " & verfahren.Count.ToString())

            Dim cntOK As Integer = 0
            Dim cntFehler As Integer = 0

            For Each v As VerfahrenInfo In verfahren
                Log("Verfahren: " & v.Verfahren)

                If v.LetzterSchritt = "STAGING_ERSTELLT" Then
                    Log("  bereits abgeschlossen uebersprungen OK")
                    Continue For
                End If

                ' Partitionswerte fuer dieses Verfahren suchen
                Dim meineWerte As List(Of PartitionsEintrag) = Nothing
                If Not verfahrenWerte.TryGetValue(v.Verfahren, meineWerte) OrElse meineWerte.Count = 0 Then
                    Log("  WARNUNG: Keine Partitionswerte fuer dieses Verfahren uebersprungen")
                    ProtokollSchreiben(connStr, v.Verfahren, "WARNUNG_SCR10", "Keine Partitionswerte in BA::objPartitionValues")
                    StatusSetzen(connStr, v.ID, "STAGING_ERSTELLT")
                    cntOK += 1
                    Continue For
                End If

                Try
                    StatusSetzen(connStr, v.ID, "STAGING_ERSTELLEN")

                    ' Spaltenliste aus INFORMATION_SCHEMA der Template-Tabelle
                    Dim selectList As String = HoleSpaltenlisteAusTemplate(connStr, v)

                    ' _out pro Partitionswert erstellen
                    Dim cntStaging As Integer = 0
                    For Each pe As PartitionsEintrag In meineWerte
                        Dim pvStr As String = pe.Wert
                        Dim outTabelle As String = v.Faktentabelle.ToLower() & "_out_" & pvStr

                        ' _out Tabelle per SELECT TOP 0 mit columns_dbo erstellen
                        SqlAusfuehren(connStr, "IF OBJECT_ID('dbo.[" & outTabelle & "]','U') IS NOT NULL DROP TABLE dbo.[" & outTabelle & "];", "_out loeschen")

                        SqlAusfuehren(connStr, "SELECT TOP 0 " & selectList & " INTO dbo.[" & outTabelle & "] FROM ext.[" & v.Faktentabelle.ToLower() & "];", "_out erstellen")

                        Log("  _out erstellt: " & pvStr & " | Modus: " & pe.Modus)
                        cntStaging += 1
                    Next

                    StatusSetzen(connStr, v.ID, "STAGING_ERSTELLT")
                    ProtokollSchreiben(connStr, v.Verfahren, "SCHRITT_5",
                        "Staging erstellt: " & cntStaging.ToString() & " Partition(en)" &
                        " | NEU: " & meineWerte.FindAll(Function(p) p.Modus = "NEU").Count.ToString() &
                        " | AKTUALISIERUNG: " & meineWerte.FindAll(Function(p) p.Modus = "AKTUALISIERUNG").Count.ToString())
                    cntOK += 1
                    Log("  Schritt 5 abgeschlossen OK")

                Catch ex As Exception
                    cntFehler += 1
                    FehlerSetzen(connStr, v.ID, ex.Message)
                    ProtokollSchreiben(connStr, v.Verfahren, "FEHLER_SCR10", ex.Message)
                    LogFehler("FEHLER '" & v.Verfahren & "': " & ex.Message)
                End Try
            Next

            Log("Erfolgreich: " & cntOK.ToString() & " | Fehler: " & cntFehler.ToString())
            Dts.TaskResult = If(cntFehler > 0, ScriptResults.Failure, ScriptResults.Success)

        Catch ex As Exception
            LogFehler("Kritischer Fehler: " & ex.Message)
            Dts.TaskResult = ScriptResults.Failure
        End Try

    End Sub

    ' -----------------------------------------------------------------------
    ' HoleSpaltenlisteAusTemplate - Reads the ordered column list from
    ' INFORMATION_SCHEMA of the template table.
    ' -----------------------------------------------------------------------
    Private Function HoleSpaltenlisteAusTemplate(connStr As String, v As VerfahrenInfo) As String
        Dim templateName As String = v.Faktentabelle.ToLower() & "_template"
        Log("  Lade Spaltenliste aus INFORMATION_SCHEMA: " & templateName)

        Dim sqlCols As String =
            "SELECT STRING_AGG(CAST(c.COLUMN_NAME AS nvarchar(max)), ',' + CHAR(13) + CHAR(10)) WITHIN GROUP (ORDER BY c.ORDINAL_POSITION) " &
            "FROM [" & _datenbank & "].INFORMATION_SCHEMA.COLUMNS c " &
            "WHERE c.TABLE_SCHEMA = 'dbo' AND c.TABLE_NAME = '" & templateName & "'"

        Dim versuch As Integer = 0
        While versuch < MAX_VERSUCHE
            versuch += 1
            Try
                Using conn As New SqlConnection(connStr)
                    conn.Open()
                    Using cmd As New SqlCommand(sqlCols, conn)
                        Dim r As Object = cmd.ExecuteScalar()
                        If r IsNot Nothing AndAlso r IsNot DBNull.Value Then Return r.ToString()
                    End Using
                End Using
            Catch ex As Exception
                Log(String.Format("WARNUNG [Template Spalten] Versuch {0}/{1}: {2}", versuch, MAX_VERSUCHE, ex.Message))
                If versuch < MAX_VERSUCHE Then System.Threading.Thread.Sleep(WARTE_SEK * 1000) Else Throw
            End Try
        End While

        Throw New Exception("Spaltenliste konnte nicht aus Template geladen werden: " & templateName)
    End Function

    ' -----------------------------------------------------------------------
    ' VerfahrenLaden - Loads the Verfahren to process from the work list
    ' (joined with the parameter table).
    ' -----------------------------------------------------------------------
    Private Function VerfahrenLaden(connStr As String) As List(Of VerfahrenInfo)
        Dim liste As New List(Of VerfahrenInfo)()
        Dim sql As String =
"SELECT a.ID, a.Verfahren, a.Themengebiet, a.LetzterSchritt,
        pf.Wert AS Faktentabelle, pp.Wert AS PartitionsSpalte
 FROM   dbo.ETL_Fkt_Arbeitsliste a
 JOIN   " & _parameterDB & ".dbo." & _parametertab & " pf ON pf.Verfahren=a.Verfahren AND pf.Parameter='Faktentabelle'
 JOIN   " & _parameterDB & ".dbo." & _parametertab & " pp ON pp.Verfahren=a.Verfahren AND pp.Parameter='Faktenpartitionsspalte'
 WHERE  a.Status IN ('PARTITIONSGRENZEN_ERSTELLT','STAGING_ERSTELLEN')
 AND    a.RunID = " & _runID.ToString() & " ORDER BY a.Verfahren"
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
                                Dim rohPart As String = rdr(5).ToString().Trim()
                                liste.Add(New VerfahrenInfo With {
                                    .ID = rdr.GetInt32(0), .Verfahren = rdr(1).ToString().Trim(),
                                    .Themengebiet = rdr(2).ToString().Trim(),
                                    .LetzterSchritt = If(rdr.IsDBNull(3), "", rdr(3).ToString().Trim()),
                                    .Faktentabelle = rdr(4).ToString().Trim(),
                                    .PartitionsSpalte = If(rohPart.Contains("|"), rohPart.Substring(0, rohPart.IndexOf("|")), rohPart)})
                            End While
                        End Using
                    End Using
                End Using
                Return liste
            Catch ex As Exception
                Log(String.Format("WARNUNG [Verfahren laden] Versuch {0}/{1}: {2}", versuch, MAX_VERSUCHE, ex.Message))
                If versuch < MAX_VERSUCHE Then System.Threading.Thread.Sleep(WARTE_SEK * 1000) Else Throw
            End Try
        End While
        Return liste
    End Function

    ' -----------------------------------------------------------------------
    ' StatusSetzen - Updates Status / LetzterSchritt of a work list row.
    ' -----------------------------------------------------------------------
    Private Sub StatusSetzen(connStr As String, id As Integer, status As String)
        SqlAusfuehren(connStr, "UPDATE dbo.ETL_Fkt_Arbeitsliste SET Status='" & status & "',LetzterSchritt='" & status & "',AktualisiertAm=GETDATE() WHERE ID=" & id.ToString(), "Status")
    End Sub

    ' -----------------------------------------------------------------------
    ' FehlerSetzen - Marks a work list row as FEHLER and stores the error
    ' message.
    ' -----------------------------------------------------------------------
    Private Sub FehlerSetzen(connStr As String, id As Integer, meldung As String)
        Try
            Using conn As New SqlConnection(connStr)
                conn.Open()
                Using cmd As New SqlCommand("UPDATE dbo.ETL_Fkt_Arbeitsliste SET Status='FEHLER',Fehlermeldung=@m,AktualisiertAm=GETDATE() WHERE ID=@id", conn)
                    cmd.Parameters.AddWithValue("@m", If(meldung.Length > 3900, meldung.Substring(0, 3900), meldung))
                    cmd.Parameters.AddWithValue("@id", id)
                    cmd.ExecuteNonQuery()
                End Using
            End Using
        Catch
        End Try
    End Sub

    ' -----------------------------------------------------------------------
    ' ProtokollSchreiben - Routes protocol messages to SSIS events:
    ' FEHLER_* -> FireError, everything else -> FireInformation.
    ' -----------------------------------------------------------------------
    Private Sub ProtokollSchreiben(connStr As String, verfahren As String, schritt As String, meldung As String)
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
                If versuch < MAX_VERSUCHE Then System.Threading.Thread.Sleep(WARTE_SEK * 1000)
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
                If versuch < MAX_VERSUCHE Then System.Threading.Thread.Sleep(WARTE_SEK * 1000) Else Throw
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
    ' PartitionsEintrag - Data container for one partition value entry.
    ' -----------------------------------------------------------------------
    Private Class PartitionsEintrag
        Public Property Wert As String
        Public Property Modus As String     ' "AKTUALISIERUNG" oder "NEU"
    End Class

    ' -----------------------------------------------------------------------
    ' VerfahrenInfo - Data container for one Verfahren work item.
    ' -----------------------------------------------------------------------
    Private Class VerfahrenInfo
        Public Property ID As Integer
        Public Property Verfahren As String
        Public Property Themengebiet As String
        Public Property LetzterSchritt As String
        Public Property Faktentabelle As String
        Public Property PartitionsSpalte As String
    End Class

    ' -----------------------------------------------------------------------
    ' ScriptResults - SSIS task result codes.
    ' -----------------------------------------------------------------------
    Public Enum ScriptResults
        Success = DTSExecResult.Success
        Failure = DTSExecResult.Failure
    End Enum

End Class
