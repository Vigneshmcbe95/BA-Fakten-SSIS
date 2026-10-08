# =============================================================================
# Speicher_Ueberwachung_Nacht.ps1
# Zweck: protokolliert alle 10 Minuten den freien Platz auf E: und F: sowie die
#        Groesse der grossen Dateien unter E:\Daten in eine CSV-Datei auf C:.
#        Nur lesend - aendert nichts am Server.
# Aufruf (auf Q0003713):  powershell -ExecutionPolicy Bypass -File Speicher_Ueberwachung_Nacht.ps1
# Laeuft bis zur Endzeit (Standard: naechster Morgen 08:00) und beendet sich dann.
# =============================================================================
param(
    [int]$IntervallMinuten = 10,
    [datetime]$Ende = (Get-Date).Date.AddDays(1).AddHours(8),
    [string]$Ausgabe = "C:\Temp\Speicher_Nacht_$(Get-Date -Format yyyyMMdd).csv",
    [long]$GrosseDateiGB = 10
)

# Ausgabe bewusst auf C:, damit das Protokoll weiterlaeuft, wenn E: voll ist
New-Item -ItemType Directory -Force -Path (Split-Path $Ausgabe) | Out-Null

while ((Get-Date) -lt $Ende) {
    $zeit = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $zeilen = @()

    # 1) Freier Platz je Laufwerk
    foreach ($lw in Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='E:' OR DeviceID='F:'") {
        $zeilen += [pscustomobject]@{
            Zeit    = $zeit
            Art     = "Laufwerk"
            Name    = $lw.DeviceID
            GB      = [math]::Round($lw.Size / 1GB, 1)
            Frei_GB = [math]::Round($lw.FreeSpace / 1GB, 1)
        }
    }

    # 2) Grosse Dateien unter E:\Daten (Datenbank-, Log- und Backupdateien)
    Get-ChildItem "E:\Daten" -Recurse -File -ErrorAction SilentlyContinue |
        Where-Object { $_.Length -ge $GrosseDateiGB * 1GB } |
        ForEach-Object {
            $zeilen += [pscustomobject]@{
                Zeit    = $zeit
                Art     = "Datei"
                Name    = $_.FullName
                GB      = [math]::Round($_.Length / 1GB, 1)
                Frei_GB = $null
            }
        }

    $zeilen | Export-Csv -Path $Ausgabe -Append -NoTypeInformation -Delimiter ";" -Encoding UTF8
    Start-Sleep -Seconds ($IntervallMinuten * 60)
}
