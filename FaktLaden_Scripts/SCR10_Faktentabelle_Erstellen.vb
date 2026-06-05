Option Explicit On
Option Strict On

Imports System
Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports Microsoft.SqlServer.Dts.Runtime

' =============================================================================
' PAKET  : Fakten Laden
' SKRIPT : SCR10_Faktentabelle_Erstellen (v2)
' ZWECK  : Pro Verfahren:
'          IF Faktentabelle NOT EXISTS -> vollstaendige Erstellung
'            (Partitionsfunktion, Schema, Tabelle, CI/CCI/NCCI)
'          IF Faktentabelle EXISTS -> uebersprungen (kein DROP, kein DELETE)
'          Status: FAKTENTABELLE_ERSTELLEN -> FAKTENTABELLE_ERSTELLT
' =============================================================================
<Microsoft.SqlServer.Dts.Tasks.ScriptTask.SSISScriptTaskEntryPointAttribute()>
<CLSCompliant(False)>
Partial Public Class ScriptMain
    Inherits Microsoft.SqlServer.Dts.Tasks.ScriptTask.VSTARTScriptObjectModelBase

    Private Const SKRIPT_NAME As String = "SCR08_Faktentabelle_Erstellen"
    Private Const CONN_NAME As String = "Verbindung"
    Private Const MAX_VERSUCHE As Integer = 10
    Private Const WARTE_SEK As Integer = 30

    Private _runID As Integer = 0
    Private _parameterDB As String = String.Empty
    Private _parametertab As String = String.Empty
    Private _datenbank As String = String.Empty

    Public Sub Main()

        Log("════════════════════════════════════════════════════════")
        Log("SCR08_Faktentabelle_Erstellen – Start (v2: kein DROP, nur CREATE wenn nicht vorhanden)")
        Log("Zeitpunkt: " & DateTime.Now.ToString("dd.MM.yyyy HH:mm:ss"))
        Log("════════════════════════════════════════════════════════")

        Try
            _runID = Convert.ToInt32(Dts.Variables("BA::RunID").Value)
            _parameterDB = Dts.Variables("BA::ParameterDB").Value.ToString().Trim()
            _parametertab = Dts.Variables("BA::Parametertabelle").Value.ToString().Trim()
            _datenbank = Dts.Variables("BA::Datenbank").Value.ToString().Trim()

            Dim connStr As String = HoleVerbindungszeichenfolge()
            Dim verfahren As List(Of VerfahrenInfo) = VerfahrenLaden(connStr)
            Log("Verfahren: " & verfahren.Count.ToString())

            Dim cntOK As Integer = 0
            Dim cntFehler As Integer = 0

            For Each v As VerfahrenInfo In verfahren
                Log("────────────────────────────────────────────────────────")
                Log("Verfahren: " & v.Verfahren & " | Faktentabelle: " & v.Faktentabelle)
                Log("Partitionsspalte: " & v.PartitionColumn & " | IndexTyp: " & v.IndexType & " | Komprimierung: " & v.Compression)

                If v.LetzterSchritt = "FAKTENTABELLE_ERSTELLT" Then
                    Log("  -> bereits abgeschlossen -> uebersprungen")
                    Continue For
                End If

                Try
                    StatusSetzen(connStr, v.ID, "FAKTENTABELLE_ERSTELLEN")

                    ' Pruefen ob Faktentabelle bereits existiert
                    Dim tabelleExistiert As Boolean = Convert.ToInt32(SqlSkalar(connStr,
                        "SELECT CASE WHEN OBJECT_ID('dbo.[" & v.Faktentabelle & "]','U') IS NOT NULL THEN 1 ELSE 0 END",
                        "Tabelle pruefen")) = 1

                    If tabelleExistiert Then
                        ' Struktur pruefen: CCI vorhanden? Partitioniert?
                        Dim hatCCI As Boolean = Convert.ToInt32(SqlSkalar(connStr,
                            "SELECT COUNT(*) FROM sys.indexes WHERE object_id=OBJECT_ID('dbo.[" & v.Faktentabelle & "]') AND type=5",
                            "CCI pruefen")) > 0
                        Dim istPartitioniert As Boolean = Convert.ToInt32(SqlSkalar(connStr,
                            "SELECT COUNT(*) FROM sys.indexes i " &
                            "JOIN sys.partition_schemes ps ON i.data_space_id=ps.data_space_id " &
                            "WHERE i.object_id=OBJECT_ID('dbo.[" & v.Faktentabelle & "]')",
                            "Partition pruefen")) > 0
                        Dim hatNotNullSpalten As Boolean = Convert.ToInt32(SqlSkalar(connStr,
                            "SELECT COUNT(*) FROM sys.columns WHERE object_id=OBJECT_ID('dbo.[" & v.Faktentabelle & "]') AND is_nullable=0",
                            "NOT NULL pruefen")) > 0
                        Dim strukturOK As Boolean = (v.IndexType <> "CCI" OrElse hatCCI) AndAlso istPartitioniert AndAlso Not hatNotNullSpalten
                        If strukturOK Then
                            Log("  Faktentabelle dbo." & v.Faktentabelle & " existiert bereits (Struktur OK) -> uebersprungen")
                            LogSchreiben(connStr, v.Verfahren, "SCHRITT_3",
                                "Faktentabelle existiert bereits (Struktur OK): dbo." & v.Faktentabelle & " -> uebersprungen")
                        Else
                            Dim hatDaten As Boolean = Convert.ToInt32(SqlSkalar(connStr,
                                "SELECT COUNT(*) FROM (SELECT TOP 1 1 AS x FROM dbo.[" & v.Faktentabelle & "]) t",
                                "Daten pruefen")) > 0
                            If Not hatDaten Then
                                Log("  Faktentabelle dbo." & v.Faktentabelle & " falsche Struktur (leer) -> DROP + neu erstellen")
                                LogSchreiben(connStr, v.Verfahren, "SCHRITT_3",
                                    "Faktentabelle falsche Struktur (leer) -> DROP + neu erstellen: dbo." & v.Faktentabelle)
                                SqlAusfuehren(connStr, "DROP TABLE dbo.[" & v.Faktentabelle & "];", "Tabelle droppen")
                                FaktentabelleErstellen(connStr, v)
                                LogSchreiben(connStr, v.Verfahren, "SCHRITT_3",
                                    "Faktentabelle neu erstellt: dbo." & v.Faktentabelle &
                                    " | PF: PF_" & v.PartitionColumn & "_" & v.Faktentabelle &
                                    " | Index: " & v.IndexType & " | Komprimierung: " & v.Compression)
                            Else
                                Log("  WARNUNG: dbo." & v.Faktentabelle & " hat Daten aber falsche Struktur -> manuelle Korrektur noetig")
                                LogSchreiben(connStr, v.Verfahren, "SCHRITT_3",
                                    "WARNUNG: Faktentabelle hat Daten aber falsche Struktur -> uebersprungen: dbo." & v.Faktentabelle)
                            End If
                        End If
                    Else
                        ' Tabelle existiert nicht -> vollstaendig erstellen
                        Log("  Faktentabelle dbo." & v.Faktentabelle & " existiert nicht -> wird erstellt")
                        FaktentabelleErstellen(connStr, v)
                        LogSchreiben(connStr, v.Verfahren, "SCHRITT_3",
                            "Faktentabelle erstellt: dbo." & v.Faktentabelle &
                            " | PF: PF_" & v.PartitionColumn & "_" & v.Faktentabelle &
                            " | Index: " & v.IndexType & " | Komprimierung: " & v.Compression)
                    End If

                    StatusSetzen(connStr, v.ID, "FAKTENTABELLE_ERSTELLT")
                    cntOK += 1
                    Log("  -> Schritt 3 abgeschlossen")

                Catch ex As Exception
                    cntFehler += 1
                    FehlerSetzen(connStr, v.ID, ex.Message)
                    LogSchreiben(connStr, v.Verfahren, "FEHLER_SCR08", ex.Message)
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

    ' =========================================================================
    ' Faktentabelle vollstaendig erstellen (nur wenn NICHT vorhanden)
    ' =========================================================================
    Private Sub FaktentabelleErstellen(connStr As String, v As VerfahrenInfo)

        Dim pf As String = "PF_" & v.PartitionColumn & "_" & v.Faktentabelle
        Dim ps As String = "PS_" & v.PartitionColumn & "_" & v.Faktentabelle
        Dim templateTable As String = "[" & _datenbank & "].dbo." & v.Faktentabelle.ToLower() & "_template"

        Dim filegroup As String = Convert.ToString(SqlSkalar(connStr,
            "SELECT name FROM sys.filegroups WHERE is_default=1", "Filegroup"))

        ' PF/PS loeschen nur wenn nicht von anderen Tabellen benutzt
        Dim psInUse As Boolean = Convert.ToInt32(SqlSkalar(connStr,
            "SELECT COUNT(*) FROM sys.indexes i " &
            "JOIN sys.partition_schemes ps2 ON i.data_space_id=ps2.data_space_id " &
            "WHERE ps2.name='" & ps & "'",
            "PS Benutzung pruefen")) > 0
        If Not psInUse Then
            If Convert.ToInt32(SqlSkalar(connStr, "SELECT COUNT(*) FROM sys.partition_schemes WHERE name='" & ps & "'", "PS pruefen")) > 0 Then
                SqlAusfuehren(connStr, "DROP PARTITION SCHEME [" & ps & "];", "PS loeschen")
            End If
            If Convert.ToInt32(SqlSkalar(connStr, "SELECT COUNT(*) FROM sys.partition_functions WHERE name='" & pf & "'", "PF pruefen")) > 0 Then
                SqlAusfuehren(connStr, "DROP PARTITION FUNCTION [" & pf & "];", "PF loeschen")
            End If
            SqlAusfuehren(connStr, "CREATE PARTITION FUNCTION [" & pf & "](INT) AS RANGE LEFT FOR VALUES(0);", "PF erstellen")
            SqlAusfuehren(connStr, "CREATE PARTITION SCHEME [" & ps & "] AS PARTITION [" & pf & "] ALL TO ([" & filegroup & "]);", "PS erstellen")
            Log("  PF und PS erstellt: " & pf & " / " & ps)
        Else
            Log("  PF/PS bereits vorhanden und in Benutzung -> wiederverwendet: " & pf & " / " & ps)
        End If

        ' columns_dbo SELECT-Liste aus tm_polybase_struktur - identisch wie SCR12 fuer _out_-Tabellen
        Dim selectList As String = Nothing
        Dim sqlCols As String =
            "SELECT STRING_AGG(CAST(m.columns_dbo AS nvarchar(max)), ',' + CHAR(13) + CHAR(10)) WITHIN GROUP (ORDER BY m.colno) " &
            "FROM [" & _datenbank & "].INFORMATION_SCHEMA.COLUMNS c " &
            "JOIN dbo.tm_polybase_struktur m ON UPPER(LTRIM(RTRIM(m.colname))) = UPPER(LTRIM(RTRIM(c.COLUMN_NAME))) " &
            "WHERE c.TABLE_SCHEMA = 'dbo' AND c.TABLE_NAME = '" & v.Faktentabelle.ToLower() & "_template' " &
            "AND m.tabname = @tab AND m.themengebiet = @thema"

        Dim versuch As Integer = 0
        While versuch < MAX_VERSUCHE
            versuch += 1
            Try
                Using conn As New SqlConnection(connStr)
                    conn.Open()
                    Using cmd As New SqlCommand(sqlCols, conn)
                        cmd.CommandTimeout = 0
                        cmd.Parameters.AddWithValue("@tab", v.Verfahren.ToLower())
                        cmd.Parameters.AddWithValue("@thema", v.Themengebiet.ToLower())
                        Dim r As Object = cmd.ExecuteScalar()
                        If r IsNot Nothing AndAlso r IsNot DBNull.Value Then selectList = r.ToString()
                    End Using
                End Using
                Exit While
            Catch ex As Exception
                Log(String.Format("WARNUNG [Template Spalten] Versuch {0}/{1}: {2}", versuch, MAX_VERSUCHE, ex.Message))
                If versuch < MAX_VERSUCHE Then System.Threading.Thread.Sleep(WARTE_SEK * 1000) Else Throw
            End Try
        End While

        If String.IsNullOrEmpty(selectList) Then
            Throw New Exception("columns_dbo Spaltenliste fuer " & v.Faktentabelle & " konnte nicht geladen werden.")
        End If

        ' Temp-Tabelle per SELECT TOP 0 erstellen (exakt gleiche Typen + Nullability wie _out_-Tabellen)
        Dim tmpTable As String = v.Faktentabelle.ToLower() & "_STRUCTTMP_"
        SqlAusfuehren(connStr, "IF OBJECT_ID('dbo.[" & tmpTable & "]') IS NOT NULL DROP TABLE dbo.[" & tmpTable & "];", "STRUCTTMP droppen")
        SqlAusfuehren(connStr, "SELECT TOP 0 " & selectList & " INTO dbo.[" & tmpTable & "] FROM ext.[" & v.Faktentabelle.ToLower() & "];", "STRUCTTMP erstellen")

        ' CREATE TABLE DDL aus Temp-Tabelle lesen (alle Spalten NULL wie ext-Quelle)
        Dim colDDL As String = Nothing
        Dim sqlColDDL As String =
            "SELECT STRING_AGG(CAST(" &
            "    QUOTENAME(c.COLUMN_NAME) + ' ' + " &
            "    CASE " &
            "        WHEN c.DATA_TYPE IN ('varchar','nvarchar','char','nchar') " &
            "            THEN c.DATA_TYPE + '(' + CASE WHEN c.CHARACTER_MAXIMUM_LENGTH = -1 THEN 'MAX' ELSE CAST(c.CHARACTER_MAXIMUM_LENGTH AS varchar(10)) END + ')' " &
            "        WHEN c.DATA_TYPE IN ('decimal','numeric') " &
            "            THEN c.DATA_TYPE + '(' + CAST(c.NUMERIC_PRECISION AS varchar(5)) + ',' + CAST(c.NUMERIC_SCALE AS varchar(5)) + ')' " &
            "        WHEN c.DATA_TYPE IN ('datetime2','time','datetimeoffset') " &
            "            THEN c.DATA_TYPE + '(' + CAST(c.DATETIME_PRECISION AS varchar(5)) + ')' " &
            "        ELSE c.DATA_TYPE " &
            "    END + ' NULL'" &
            "AS nvarchar(max)), ',' + CHAR(13) + CHAR(10)) WITHIN GROUP (ORDER BY c.ORDINAL_POSITION) " &
            "FROM INFORMATION_SCHEMA.COLUMNS c " &
            "WHERE c.TABLE_SCHEMA = 'dbo' AND c.TABLE_NAME = '" & tmpTable & "'"

        versuch = 0
        While versuch < MAX_VERSUCHE
            versuch += 1
            Try
                Using conn As New SqlConnection(connStr)
                    conn.Open()
                    Using cmd As New SqlCommand(sqlColDDL, conn)
                        cmd.CommandTimeout = 0
                        Dim r As Object = cmd.ExecuteScalar()
                        If r IsNot Nothing AndAlso r IsNot DBNull.Value Then colDDL = r.ToString()
                    End Using
                End Using
                Exit While
            Catch ex As Exception
                Log(String.Format("WARNUNG [DDL lesen] Versuch {0}/{1}: {2}", versuch, MAX_VERSUCHE, ex.Message))
                If versuch < MAX_VERSUCHE Then System.Threading.Thread.Sleep(WARTE_SEK * 1000) Else Throw
            End Try
        End While

        SqlAusfuehren(connStr, "DROP TABLE dbo.[" & tmpTable & "];", "STRUCTTMP droppen")

        If String.IsNullOrEmpty(colDDL) Then
            Throw New Exception("DDL fuer Faktentabelle konnte nicht aus Temp-Tabelle geladen werden.")
        End If

        ' Faktentabelle auf Partitionsschema erstellen
        Dim sqlCreate As String = "CREATE TABLE dbo.[" & v.Faktentabelle.ToLower() & "] (" & colDDL & ") ON [" & ps & "]([" & v.PartitionColumn & "]);"
        SqlAusfuehren(connStr, sqlCreate, "Faktentabelle erstellen")

        SqlAusfuehren(connStr, "ALTER TABLE dbo.[" & v.Faktentabelle.ToLower() & "] SET (LOCK_ESCALATION=AUTO);", "LOCK_ESCALATION")
        Log("  Faktentabelle erstellt: dbo." & v.Faktentabelle.ToLower())

        ' CCI anlegen - erbt Partition vom Partitionsschema der Tabelle
        If v.IndexType = "CCI" Then
            SqlAusfuehren(connStr,
                "CREATE CLUSTERED COLUMNSTORE INDEX [CCI_" & v.Faktentabelle & "] ON dbo.[" & v.Faktentabelle.ToLower() & "];",
                "CCI erstellen")
            Log("  CCI angelegt (partitioniert)")
        End If

        ' NCCI
        If v.NcciFlag = "TRUE" Then
            Dim allCols As String = Convert.ToString(SqlSkalar(connStr,
                "SELECT STRING_AGG(QUOTENAME(name),',') FROM sys.columns WHERE object_id=OBJECT_ID('dbo." & v.Faktentabelle.ToLower() & "') AND is_computed=0",
                "NCCI Spalten"))
            SqlAusfuehren(connStr,
                "CREATE NONCLUSTERED COLUMNSTORE INDEX [NCCI_" & v.Faktentabelle & "] ON dbo.[" & v.Faktentabelle.ToLower() & "] (" & allCols & ");",
                "NCCI erstellen")
            Log("  NCCI angelegt")
        End If

    End Sub

    Private Function VerfahrenLaden(connStr As String) As List(Of VerfahrenInfo)
        Dim liste As New List(Of VerfahrenInfo)()
        Dim sql As String =
"SELECT a.ID, a.Verfahren, a.Themengebiet, a.LetzterSchritt," &
"       pf.Wert  AS Faktentabelle," &
"       pp.Wert  AS PartitionColumn," &
"       pc.Wert  AS Compression," &
"       pi.Wert  AS IndexType," &
"       UPPER(ISNULL(pn.Wert,'FALSE')) AS NcciFlag" &
" FROM  dbo.ETL_Fkt_Arbeitsliste a" &
" JOIN  " & _parameterDB & ".dbo." & _parametertab & " pf ON pf.Verfahren=a.Verfahren AND pf.Parameter='Faktentabelle'" &
" JOIN  " & _parameterDB & ".dbo." & _parametertab & " pp ON pp.Verfahren=a.Verfahren AND pp.Parameter='Faktenpartitionsspalte'" &
" LEFT JOIN " & _parameterDB & ".dbo." & _parametertab & " pc ON pc.Verfahren=a.Verfahren AND pc.Parameter='Faktenkomprimierung'" &
" JOIN  " & _parameterDB & ".dbo." & _parametertab & " pi ON pi.Verfahren=a.Verfahren AND pi.Parameter='FaktenClusteredIndex'" &
" LEFT JOIN " & _parameterDB & ".dbo." & _parametertab & " pn ON pn.Verfahren=a.Verfahren AND pn.Parameter='FaktenNccIndex'" &
" WHERE a.Status IN ('EXT_TABELLE_ERSTELLT','FAKTENTABELLE_ERSTELLEN')" &
" AND   a.RunID = " & _runID &
" ORDER BY a.Verfahren"
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
                                Dim rawPart As String = rdr(5).ToString().Trim()
                                Dim partCol As String = If(rawPart.Contains("|"), rawPart.Substring(0, rawPart.IndexOf("|")), rawPart)
                                liste.Add(New VerfahrenInfo With {
                                    .ID = rdr.GetInt32(0), .Verfahren = rdr(1).ToString().Trim(),
                                    .Themengebiet = rdr(2).ToString().Trim(),
                                    .LetzterSchritt = If(rdr.IsDBNull(3), "", rdr(3).ToString().Trim()),
                                    .Faktentabelle = rdr(4).ToString().Trim(),
                                    .PartitionColumn = partCol,
                                    .Compression = If(rdr.IsDBNull(6), "NONE", rdr(6).ToString().Trim().ToUpper()),
                                    .IndexType = rdr(7).ToString().Trim().ToUpper(),
                                    .NcciFlag = rdr(8).ToString().Trim()})
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

    Private Sub StatusSetzen(connStr As String, id As Integer, status As String)
        SqlAusfuehren(connStr, "UPDATE dbo.ETL_Fkt_Arbeitsliste SET Status='" & status & "',LetzterSchritt='" & status & "',AktualisiertAm=GETDATE() WHERE ID=" & id, "Status")
    End Sub

    Private Sub FehlerSetzen(connStr As String, id As Integer, msg As String)
        Try
            Using conn As New SqlConnection(connStr)
                conn.Open()
                Using cmd As New SqlCommand("UPDATE dbo.ETL_Fkt_Arbeitsliste SET Status='FEHLER',Fehlermeldung=@m,AktualisiertAm=GETDATE() WHERE ID=@id", conn)
                    cmd.Parameters.AddWithValue("@m", If(msg.Length > 3900, msg.Substring(0, 3900), msg))
                    cmd.Parameters.AddWithValue("@id", id)
                    cmd.ExecuteNonQuery()
                End Using
            End Using
        Catch
        End Try
    End Sub

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
                If versuch < MAX_VERSUCHE Then System.Threading.Thread.Sleep(WARTE_SEK * 1000)
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
                If versuch < MAX_VERSUCHE Then System.Threading.Thread.Sleep(WARTE_SEK * 1000) Else Throw
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

    Private Class VerfahrenInfo
        Public Property ID As Integer
        Public Property Verfahren As String
        Public Property Themengebiet As String
        Public Property LetzterSchritt As String
        Public Property Faktentabelle As String
        Public Property PartitionColumn As String
        Public Property Compression As String
        Public Property IndexType As String
        Public Property NcciFlag As String
    End Class

    Public Enum ScriptResults
        Success = DTSExecResult.Success
        Failure = DTSExecResult.Failure
    End Enum

End Class
