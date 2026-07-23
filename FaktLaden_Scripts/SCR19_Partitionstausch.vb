Option Explicit On
Option Strict On
Imports System
Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports Microsoft.SqlServer.Dts.Runtime

' =============================================================================
'  Skript       : SCR19_Partitionstausch
'  Paket        : Fakten Laden (SSIS)
'  Zweck        : Fuehrt den Partitionstausch je geladener Partition durch:
'                 SWITCH OUT nach _out_, CHECK-Constraint auf _in_, SWITCH
'                 IN, Aufraeumen und Endstatus.
'  Ablauf       : NCCI_OUT_ERSTELLT -> ERFOLG
'  Wiederholung : 3 Versuche je SQL-Anweisung, 30 s Wartezeit
'  Protokoll    : Nur SSIS-Events (FireInformation / FireError)
' =============================================================================
<Microsoft.SqlServer.Dts.Tasks.ScriptTask.SSISScriptTaskEntryPointAttribute()>
<CLSCompliant(False)>
Partial Public Class ScriptMain
    Inherits Microsoft.SqlServer.Dts.Tasks.ScriptTask.VSTARTScriptObjectModelBase

    Private Const SKRIPT_NAME As String = "SCR19_Partitionstausch"
    Private Const CONN_NAME As String = "Verbindung"
    Private Const MAX_VERSUCHE As Integer = 3
    Private Const WARTE_SEK As Integer = 30
    Private _runID As Integer = 0
    Private _parameterDB As String = String.Empty
    Private _parametertab As String = String.Empty

    ' -----------------------------------------------------------------------
    ' Main - Einstiegspunkt - steuert den Ablauf des Skripts.
    ' -----------------------------------------------------------------------
    Public Sub Main()
        Log("SCR19_Partitionstausch - Start")
        Try
            _runID = Convert.ToInt32(Dts.Variables("BA::RunID").Value)
            _parameterDB = Dts.Variables("BA::ParameterDB").Value.ToString().Trim()
            _parametertab = Dts.Variables("BA::Parametertabelle").Value.ToString().Trim()
            Dim connStr As String = HoleVerbindungszeichenfolge()
            ' Partitionswerte des aktuellen Laufs aus BA::objPartitionValues (gesetzt von SCR09)
            ' Nur diese Tabellen verarbeiten — kein sys.tables LIKE-Scan ueber alle Laeufe
            Dim verfahrenWerte As New Dictionary(Of String, List(Of String))()
            Dim partObjekt As Object = Dts.Variables("BA::objPartitionValues").Value
            If partObjekt IsNot Nothing Then
                Dim partArray(,) As String = CType(partObjekt, String(,))
                For i As Integer = 0 To partArray.GetLength(0) - 1
                    Dim verf As String = partArray(i, 0).Trim().ToLower()
                    Dim wert As String = partArray(i, 1).Trim()
                    If Not verfahrenWerte.ContainsKey(verf) Then verfahrenWerte(verf) = New List(Of String)()
                    verfahrenWerte(verf).Add(wert)
                Next
            End If
            Log("Partitionswerte geladen: " & verfahrenWerte.Count.ToString() & " Verfahren")
            Dim verfahren As List(Of VerfahrenInfo) = VerfahrenLaden(connStr)
            Log("Verfahren: " & verfahren.Count.ToString())
            Dim cntOK As Integer = 0
            Dim cntFehler As Integer = 0
            For Each v As VerfahrenInfo In verfahren
                Log("Verfahren: " & v.Verfahren & " | Tabelle: " & v.Faktentabelle)
                If v.LetzterSchritt = "PARTITIONSTAUSCH_ERFOLG" Then
                    Log("  bereits abgeschlossen uebersprungen OK")
                    Continue For
                End If
                Try
                    StatusSetzen(connStr, v.ID, "PARTITIONSTAUSCH")
                    Dim pf As String = "PF_" & v.PartitionColumn & "_" & v.Faktentabelle
                    ' Nur Tabellen des aktuellen Laufs — exakte Namen aus BA::objPartitionValues
                    Dim verfKey As String = v.Verfahren.Trim().ToLower()
                    Dim werteListe As List(Of String) = Nothing
                    If Not verfahrenWerte.TryGetValue(verfKey, werteListe) OrElse werteListe.Count = 0 Then
                        Log("  WARNUNG: Keine Partitionswerte in BA::objPartitionValues -> uebersprungen")
                        Continue For
                    End If
                    Log("  Partitionen: " & werteListe.Count.ToString())
                    ' Je Partition isoliert tauschen: ein Fehler bei einer Partition
                    ' bricht nicht den ganzen Lauf ab. Bei Fehler wird die _in_-Tabelle
                    ' entfernt -> der naechste Lauf laedt NUR diese Partition neu.
                    Dim partOK As Integer = 0
                    Dim partFehler As Integer = 0
                    Dim ersterFehler As String = Nothing

                    For Each pvStr As String In werteListe
                        Dim inTable  As String = v.Faktentabelle.ToLower() & "_in_"  & pvStr
                        Dim outTable As String = v.Faktentabelle.ToLower() & "_out_" & pvStr
                        Dim pnrVal As Integer = 0
                        Dim switchedOut As Boolean = False
                        Try
                            Log("  Partition: " & pvStr)
                            ' Leere Grenzpartition (keine Oracle-Daten, z.B. 202606): _in_-Tabelle
                            ' wurde von SCR13 nicht erzeugt -> Tausch ueberspringen statt Fehler.
                            ' Leere _out_-Huelle aufraeumen, damit kein Waisenobjekt bleibt.
                            If Convert.ToInt32(SqlSkalar(connStr, "SELECT CASE WHEN OBJECT_ID('dbo.[" & inTable & "]','U') IS NULL THEN 0 ELSE 1 END", "in vorhanden")) = 0 Then
                                Log("  Keine Daten geladen (leere Partition) -> Tausch uebersprungen: " & pvStr)
                                LogSchreiben(connStr, v.Verfahren, "LEER_" & pvStr, "Keine Daten in Oracle - Partitionstausch uebersprungen")
                                If Convert.ToInt32(SqlSkalar(connStr, "SELECT CASE WHEN OBJECT_ID('dbo.[" & outTable & "]','U') IS NOT NULL THEN 1 ELSE 0 END", "out vorhanden")) = 1 Then
                                    SqlAusfuehren(connStr, "DROP TABLE dbo.[" & outTable & "];", "drop leeres _out")
                                End If
                                Continue For
                            End If
                            ' Partitionsnummer per $partition-Funktion — gibt direkt die korrekte Nummer zurueck
                            Dim pnr As Object = SqlSkalar(connStr,
                                "SELECT $partition.[" & pf & "](" & pvStr & ")",
                                "Partitionsnummer")
                            If pnr Is Nothing OrElse pnr Is DBNull.Value OrElse Convert.ToInt32(pnr) = 0 Then
                                Throw New Exception("Partitionsnummer nicht gefunden fuer " & pvStr & " (Boundary fehlt - SCR11 SPLIT pruefen)")
                            End If
                            pnrVal = Convert.ToInt32(pnr)
                            Log("  Partitionsnummer: " & pnrVal.ToString())
                            ' Sicherheitsnetz: nichtclusterte (Rowstore) Indizes der Faktentabelle
                            ' (z.B. manuell angelegte wie nci_tf_bst_gb) koennen auf _out_/_in_
                            ' fehlen, wenn SCR16 sie fuer diese Partition nicht erreicht hat.
                            ' Hier direkt vor dem jeweiligen SWITCH nachpruefen/nachbauen, statt
                            ' uns auf SCR16 zu verlassen - sonst schlaegt SWITCH mit "no identical
                            ' index" fehl, obwohl die eigentliche Ursache weiter oben im Ablauf lag.
                            NonclusteredReplizieren(connStr, v.Faktentabelle, outTable)

                            ' SWITCH OUT (alte Partitionsdaten -> _out_)
                            SqlAusfuehren(connStr, "ALTER TABLE dbo.[" & v.Faktentabelle & "] SWITCH PARTITION " & pnrVal & " TO dbo.[" & outTable & "];", "SWITCH OUT")
                            switchedOut = True
                            Log("  SWITCH OUT " & outTable & " OK")
                            ' CHECK Constraint auf _in — beide Grenzen explizit (RANGE LEFT benoetigt > untere Grenze)
                            Dim ckName As String = v.PartitionColumn & "_" & pvStr & "_" & v.Faktentabelle & "_CK"
                            If Convert.ToInt32(SqlSkalar(connStr, "SELECT COUNT(*) FROM sys.check_constraints WHERE parent_object_id=OBJECT_ID('dbo." & inTable & "') AND name='" & ckName & "'", "CK pruefen")) > 0 Then
                                SqlAusfuehren(connStr, "ALTER TABLE dbo.[" & inTable & "] DROP CONSTRAINT [" & ckName & "];", "CK loeschen")
                            End If
                            ' Untere Partitionsgrenze aus PF lesen (groesster Grenzwert der kleiner ist als pvStr)
                            Dim lbObj As Object = SqlSkalar(connStr,
                                "SELECT ISNULL(MAX(CAST(sprv.value AS bigint)), -2147483648) " &
                                "FROM sys.partition_functions spf " &
                                "JOIN sys.partition_range_values sprv ON sprv.function_id=spf.function_id " &
                                "WHERE spf.name='" & pf & "' AND CAST(sprv.value AS bigint) < " & pvStr,
                                "Untere Grenze")
                            Dim lowerBound As Long = Convert.ToInt64(lbObj)
                            SqlAusfuehren(connStr,
                                "ALTER TABLE dbo.[" & inTable & "] ADD CONSTRAINT [" & ckName & "] " &
                                "CHECK([" & v.PartitionColumn & "] IS NOT NULL AND [" & v.PartitionColumn & "] > " & lowerBound.ToString() & " AND [" & v.PartitionColumn & "] <= " & pvStr & ");",
                                "CK setzen")
                            Log("  CHECK Constraint: " & ckName & " (" & v.PartitionColumn & " IS NOT NULL AND > " & lowerBound.ToString() & " AND <= " & pvStr & ") OK")
                            ' Sicherheitsnetz (siehe Kommentar bei SWITCH OUT) - auch fuer _in_
                            ' direkt vor SWITCH IN pruefen/nachbauen.
                            NonclusteredReplizieren(connStr, v.Faktentabelle, inTable)
                            ' SWITCH IN (neue Daten -> Faktenpartition)
                            SqlAusfuehren(connStr, "ALTER TABLE dbo.[" & inTable & "] SWITCH TO dbo.[" & v.Faktentabelle & "] PARTITION " & pnrVal & ";", "SWITCH IN")
                            Log("  SWITCH IN " & v.Faktentabelle & " Partition " & pnrVal.ToString() & " OK")
                            ' Cleanup (Erfolg): _out_ (alte Daten) und _in_ entfernen
                            If Convert.ToInt32(SqlSkalar(connStr, "SELECT CASE WHEN OBJECT_ID('dbo." & outTable & "','U') IS NOT NULL THEN 1 ELSE 0 END", "out pruefen")) = 1 Then
                                SqlAusfuehren(connStr, "DROP TABLE dbo.[" & outTable & "];", "drop _out")
                                Log("  _out geloescht OK")
                            End If
                            If Convert.ToInt32(SqlSkalar(connStr, "SELECT CASE WHEN OBJECT_ID('dbo." & inTable & "','U') IS NOT NULL THEN 1 ELSE 0 END", "in pruefen")) = 1 Then
                                SqlAusfuehren(connStr, "DROP TABLE dbo.[" & inTable & "];", "drop _in")
                                Log("  _in geloescht OK")
                            End If
                            LogSchreiben(connStr, v.Verfahren, "SWITCH_" & pvStr, "SWITCH IN erfolgreich " & v.Faktentabelle & " Partition " & pnrVal.ToString())
                            partOK += 1
                        Catch pex As Exception
                            partFehler += 1
                            If ersterFehler Is Nothing Then ersterFehler = "pv=" & pvStr & ": " & pex.Message
                            LogFehler("  FEHLER Partitionstausch pv=" & pvStr & ": " & pex.Message)
                            LogSchreiben(connStr, v.Verfahren, "FEHLER_SWITCH_" & pvStr, pex.Message)
                            ' Rollback: wenn schon SWITCH OUT erfolgte, alte Daten zurueck in die
                            ' (jetzt leere) Faktenpartition schalten -> KEIN Datenverlust.
                            Dim outGeleert As Boolean = Not switchedOut   ' nie ausgetauscht -> _out_ leer
                            If switchedOut AndAlso pnrVal > 0 Then
                                Try
                                    SqlAusfuehren(connStr, "ALTER TABLE dbo.[" & outTable & "] SWITCH TO dbo.[" & v.Faktentabelle & "] PARTITION " & pnrVal & ";", "ROLLBACK _out zurueck")
                                    outGeleert = True
                                    Log("  ROLLBACK: alte Daten aus " & outTable & " zurueck in Partition " & pnrVal.ToString() & " OK")
                                Catch rbx As Exception
                                    LogFehler("  WARNUNG: Rollback (_out zurueck) fehlgeschlagen fuer " & pvStr & " - alte Daten verbleiben in " & outTable & ": " & rbx.Message)
                                End Try
                            End If
                            ' _in_ IMMER entfernen -> naechster Lauf laedt NUR diese Partition neu.
                            Try
                                SqlAusfuehren(connStr, "IF OBJECT_ID('dbo.[" & inTable & "]','U') IS NOT NULL DROP TABLE dbo.[" & inTable & "];", "drop _in nach Fehler")
                            Catch
                            End Try
                            ' _out_ nur entfernen, wenn leer (sonst einzige Kopie der alten Daten -> behalten).
                            If outGeleert Then
                                Try
                                    SqlAusfuehren(connStr, "IF OBJECT_ID('dbo.[" & outTable & "]','U') IS NOT NULL DROP TABLE dbo.[" & outTable & "];", "drop _out nach Fehler")
                                Catch
                                End Try
                            End If
                        End Try
                    Next
                    ' Abschlussstatus MSSQL (immer protokollieren)
                    Dim finalMin As Object = SqlSkalar(connStr, "SELECT MIN([" & v.PartitionColumn & "]) FROM dbo.[" & v.Faktentabelle & "]", "Final MIN")
                    Dim finalMax As Object = SqlSkalar(connStr, "SELECT MAX([" & v.PartitionColumn & "]) FROM dbo.[" & v.Faktentabelle & "]", "Final MAX")
                    Dim finalCnt As Object = SqlSkalar(connStr, "SELECT COUNT_BIG(*) FROM dbo.[" & v.Faktentabelle & "]", "Final COUNT")
                    Log("  ABSCHLUSSSTATUS: " & v.Faktentabelle)
                    Log("  Getauscht OK: " & partOK.ToString() & " | Fehler: " & partFehler.ToString())
                    Log("  Zeilen: " & Convert.ToString(finalCnt))
                    Log("  MIN:    " & If(finalMin Is Nothing OrElse finalMin Is DBNull.Value, "NULL", Convert.ToString(finalMin)))
                    Log("  MAX:    " & If(finalMax Is Nothing OrElse finalMax Is DBNull.Value, "NULL", Convert.ToString(finalMax)))
                    LogSchreiben(connStr, v.Verfahren, "ABSCHLUSS",
                        "Getauscht OK: " & partOK.ToString() & " | Fehler: " & partFehler.ToString() &
                        " | Zeilen: " & Convert.ToString(finalCnt) &
                        " | MIN: " & If(finalMin Is Nothing OrElse finalMin Is DBNull.Value, "NULL", Convert.ToString(finalMin)) &
                        " | MAX: " & If(finalMax Is Nothing OrElse finalMax Is DBNull.Value, "NULL", Convert.ToString(finalMax)))

                    If partFehler > 0 Then
                        Dim fmsg As String = partFehler.ToString() & " von " & werteListe.Count.ToString() &
                            " Partition(en) Tausch fehlgeschlagen (erste: " & If(ersterFehler, "?") & "). " &
                            "Erfolgreich getauscht: " & partOK.ToString() & ". Betroffene _in_-Tabellen entfernt -> " &
                            "naechster Lauf laedt NUR die fehlgeschlagenen Partitionen neu."
                        FehlerSetzen(connStr, v.ID, fmsg)
                        LogSchreiben(connStr, v.Verfahren, "FEHLER_SCR15", fmsg)
                        LogFehler("  Verfahren '" & v.Verfahren & "': " & fmsg)
                        cntFehler += 1
                    Else
                        StatusSetzenErfolg(connStr, v.ID)
                        cntOK += 1
                        Log("  Verfahren erfolgreich abgeschlossen OK")
                    End If
                Catch ex As Exception
                    cntFehler += 1
                    FehlerSetzen(connStr, v.ID, ex.Message)
                    LogSchreiben(connStr, v.Verfahren, "FEHLER_SCR15", ex.Message)
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
    ' NonclusteredReplizieren - Liest ALLE NONCLUSTERED (Rowstore) Indizes der
    ' Faktentabelle (index_id > 1) und baut jeden fehlenden 1:1 auf der Ziel-
    ' Tabelle (_in_/_out_) nach (Name, UNIQUE, Schluessel-/INCLUDE-Spalten,
    ' Filter), BEVOR SWITCH versucht wird. Identisch zur Logik in SCR16 -
    ' hier als Sicherheitsnetz direkt vor dem SWITCH, falls SCR16 diese
    ' Tabelle nicht erreicht hat. Columnstore (NCCI) bleibt aussen vor.
    ' -----------------------------------------------------------------------
    Private Sub NonclusteredReplizieren(connStr As String, factTable As String, zielTable As String)
        Dim sql As String =
            "SELECT i.name AS idxName," &
            " 'CREATE ' + CASE WHEN i.is_unique = 1 THEN 'UNIQUE ' ELSE '' END +" &
            " 'NONCLUSTERED INDEX [' + i.name + '] ON dbo.[" & zielTable & "] (' +" &
            " STUFF((SELECT ', [' + c.name + ']' + CASE WHEN ic.is_descending_key = 1 THEN ' DESC' ELSE '' END" &
            "        FROM sys.index_columns ic JOIN sys.columns c ON c.object_id = ic.object_id AND c.column_id = ic.column_id" &
            "        WHERE ic.object_id = i.object_id AND ic.index_id = i.index_id AND ic.is_included_column = 0" &
            "        ORDER BY ic.key_ordinal FOR XML PATH(''), TYPE).value('.','nvarchar(max)'), 1, 2, '') + ')' +" &
            " ISNULL(' INCLUDE (' + STUFF((SELECT ', [' + c.name + ']'" &
            "        FROM sys.index_columns ic JOIN sys.columns c ON c.object_id = ic.object_id AND c.column_id = ic.column_id" &
            "        WHERE ic.object_id = i.object_id AND ic.index_id = i.index_id AND ic.is_included_column = 1" &
            "        ORDER BY ic.index_column_id FOR XML PATH(''), TYPE).value('.','nvarchar(max)'), 1, 2, '') + ')', '') +" &
            " ISNULL(' WHERE ' + i.filter_definition, '') + ';' AS createStmt" &
            " FROM sys.indexes i" &
            " WHERE i.object_id = OBJECT_ID('dbo.[" & factTable.ToLower() & "]')" &
            "   AND i.index_id > 1 AND i.type_desc = 'NONCLUSTERED'"

        Dim aufbau As New List(Of String())()
        Using conn As New SqlConnection(connStr)
            conn.Open()
            Using cmd As New SqlCommand(sql, conn)
                cmd.CommandTimeout = 0
                Using rdr As SqlDataReader = cmd.ExecuteReader()
                    While rdr.Read()
                        aufbau.Add(New String() {rdr(0).ToString().Trim(), rdr(1).ToString()})
                    End While
                End Using
            End Using
        End Using

        If aufbau.Count = 0 Then
            Log("    Sicherheitsnetz-Indexpruefung " & zielTable & ": keine zusaetzlichen NCI auf Faktentabelle gefunden")
            Return
        End If

        For Each idx As String() In aufbau
            Dim idxName As String = idx(0)
            Dim createStmt As String = idx(1)
            If IndexVorhanden(connStr, zielTable, idxName) Then
                Log("    Sicherheitsnetz-Indexpruefung " & zielTable & ": " & idxName & " bereits vorhanden")
            Else
                SqlAusfuehren(connStr, createStmt, "NCI " & idxName & " auf " & zielTable)
                Log("    WARNUNG: fehlenden Index " & idxName & " erst hier (SCR19) auf " & zielTable & " nachgebaut - SCR16 hat ihn nicht angelegt")
            End If
        Next
    End Sub

    ' -----------------------------------------------------------------------
    ' IndexVorhanden - Prueft, ob ein Index auf einer Tabelle existiert.
    ' -----------------------------------------------------------------------
    Private Function IndexVorhanden(connStr As String, tbl As String, idxName As String) As Boolean
        Return Convert.ToInt32(SqlSkalar(connStr, "SELECT COUNT(*) FROM sys.indexes WHERE object_id=OBJECT_ID('dbo.[" & tbl & "]') AND name='" & idxName & "'", "Index pruefen")) > 0
    End Function

    ' -----------------------------------------------------------------------
    ' VerfahrenLaden - Laedt die zu verarbeitenden Verfahren aus der
    ' Arbeitsliste (Join mit der Parametertabelle).
    ' -----------------------------------------------------------------------
    Private Function VerfahrenLaden(connStr As String) As List(Of VerfahrenInfo)
        Dim liste As New List(Of VerfahrenInfo)()
        Dim sql As String =
"SELECT a.ID, a.Verfahren, a.LetzterSchritt, pf.Wert AS Faktentabelle, pp.Wert AS PartitionColumn
 FROM   dbo.ETL_Fakt_Arbeitsliste a
 JOIN   " & _parameterDB & ".dbo." & _parametertab & " pf ON pf.Verfahren=dbo.fn_ParamVerfahren(a.Verfahren) AND pf.Parameter='Faktentabelle'
 JOIN   " & _parameterDB & ".dbo." & _parametertab & " pp ON pp.Verfahren=dbo.fn_ParamVerfahren(a.Verfahren) AND pp.Parameter='Faktenpartitionsspalte'
 WHERE  a.Status IN ('NCCI_OUT_ERSTELLT','PARTITIONSTAUSCH') AND a.RunID=" & _runID & " ORDER BY a.Verfahren"
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
                                Dim rawPart As String = rdr(4).ToString().Trim()
                                liste.Add(New VerfahrenInfo With {
                                    .ID = rdr.GetInt32(0), .Verfahren = rdr(1).ToString().Trim(),
                                    .LetzterSchritt = If(rdr.IsDBNull(2), "", rdr(2).ToString().Trim()),
                                    .Faktentabelle = rdr(3).ToString().Trim(),
                                    .PartitionColumn = If(rawPart.Contains("|"), rawPart.Substring(0, rawPart.IndexOf("|")), rawPart)})
                            End While
                        End Using
                    End Using
                End Using
                Return liste
            Catch ex As Exception
                If versuch < MAX_VERSUCHE Then System.Threading.Thread.Sleep(WARTE_SEK * 1000) Else Throw
            End Try
        End While
        Return liste
    End Function

    ' -----------------------------------------------------------------------
    ' StatusSetzen - Aktualisiert Status / LetzterSchritt einer
    ' Arbeitslisten-Zeile.
    ' -----------------------------------------------------------------------
    Private Sub StatusSetzen(connStr As String, id As Integer, status As String)
        SqlAusfuehren(connStr, "UPDATE dbo.ETL_Fakt_Arbeitsliste SET Status='" & status & "',LetzterSchritt='" & status & "',AktualisiertAm=GETDATE() WHERE ID=" & id, "Status")
    End Sub

    ' -----------------------------------------------------------------------
    ' StatusSetzenErfolg - Setzt den Endstatus ERFOLG auf einer
    ' Arbeitslisten-Zeile.
    ' -----------------------------------------------------------------------
    Private Sub StatusSetzenErfolg(connStr As String, id As Integer)
        SqlAusfuehren(connStr, "UPDATE dbo.ETL_Fakt_Arbeitsliste SET Status='ERFOLG',LetzterSchritt='ERFOLG',AktualisiertAm=GETDATE() WHERE ID=" & id, "ERFOLG")
    End Sub

    ' -----------------------------------------------------------------------
    ' FehlerSetzen - Markiert eine Arbeitslisten-Zeile als FEHLER und
    ' speichert die Fehlermeldung.
    ' -----------------------------------------------------------------------
    Private Sub FehlerSetzen(connStr As String, id As Integer, msg As String)
        Try
            Using conn As New SqlConnection(connStr)
                conn.Open()
                Using cmd As New SqlCommand("UPDATE dbo.ETL_Fakt_Arbeitsliste SET Status='FEHLER',Fehlermeldung=@m,AktualisiertAm=GETDATE() WHERE ID=@id", conn)
                    cmd.Parameters.AddWithValue("@m", If(msg.Length > 3900, msg.Substring(0, 3900), msg))
                    cmd.Parameters.AddWithValue("@id", id)
                    cmd.ExecuteNonQuery()
                End Using
            End Using
        Catch
        End Try
    End Sub

    ' -----------------------------------------------------------------------
    ' LogSchreiben - Leitet Protokollmeldungen an SSIS-Events weiter:
    ' FEHLER_* -> FireError, alles andere -> FireInformation.
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
    ' SqlAusfuehren - Fuehrt eine SQL-Anweisung (Non-Query) mit
    ' Wiederholung aus; protokolliert Warnung und vollstaendiges
    ' SQL-Statement bei Fehlern.
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
    ' SqlSkalar - Fuehrt eine skalare SQL-Abfrage mit Wiederholung aus;
    ' protokolliert Warnung und vollstaendiges SQL-Statement bei Fehlern.
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

    ' -----------------------------------------------------------------------
    ' VerfahrenInfo - Datencontainer fuer ein Verfahren der Arbeitsliste.
    ' -----------------------------------------------------------------------
    Private Class VerfahrenInfo
        Public Property ID As Integer
        Public Property Verfahren As String
        Public Property LetzterSchritt As String
        Public Property Faktentabelle As String
        Public Property PartitionColumn As String
    End Class

    ' -----------------------------------------------------------------------
    ' ScriptResults - SSIS-Task-Ergebniscodes.
    ' -----------------------------------------------------------------------
    Public Enum ScriptResults
        Success = DTSExecResult.Success
        Failure = DTSExecResult.Failure
    End Enum

End Class
