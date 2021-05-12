$ie = New-Object -ComObject 'internetExplorer.Application' -EA Ignore -ErrorVariable global:Fehler
$ie.Visible = $false
$ie.Navigate("https://reiseauskunft.bahn.de/bin/query.exe/dn?protocol=https:")
While($ie.Busy -eq $true){Start-Sleep -s 3}
$start = $ie.Document.IHTMLDocument3_getElementById("locS0")
$start.value = Read-Host "Geben Sie den Startpunkt an" #eindeutige Station erforderlich! Beispiel: Messe/Ost (EXPO-Plaza), Hannover
$start.click()
$ziel = $ie.Document.IHTMLDocument3_getElementById("locZ0")
$ziel.value = Read-Host "Geben Sie den Zielort an" #eindeutige Station erforderlich! Beispiel: Steintor (U), Hannover
$ziel.click()
$Ankunft_Tag = $ie.Document.IHTMLDocument3_getElementById("REQ0JourneyDate") 
$Ankunft_Tag.value = "Mi, 12.05.21" #Format: DD, dd,mm,yy Beispiel: Mi, 12.05.21
$Ankunft_Zeit = $ie.Document.IHTMLDocument3_getElementById("REQ0JourneyTime") 
$Ankunft_Zeit.value = "17:51" #Format: MM:HH Beispiel: 17:51
Start-Sleep -s 1
$Such_Button = $ie.Document.IHTMLDocument3_getElementById("searchConnectionButton")
$Such_Button.click()
While($ie.Busy -eq $true){Start-Sleep -s 3}
$Ergebnis_01 = $ie.Document.IHTMLDocument3_getElementById("overview_updateC1-0")
$Ergebnis_02 = $ie.Document.IHTMLDocument3_getElementById("overview_updateC1-1")
$Ergebnis_03 = $ie.Document.IHTMLDocument3_getElementById("overview_updateC1-2")
write-host $Ergebnis_01.outerText -ForegroundColor Blue
write-host $Ergebnis_02.outerText -ForegroundColor Blue
write-host $Ergebnis_03.outerText -ForegroundColor Blue
Stop-Process -Name iexplore -EA SilentlyContinue
