#make outlook offline, use when needed
function Get-OutlookOffline {
foreach ($prof in Get-ChildItem 'HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles\'){
if (Test-Path ("$prof".Replace('HKEY_CURRENT_USER','HKCU:')+"\0a0d020000000000c000000000000046")){
$YourInput = "01,00,00,00"
$RegPath   = "$prof".Replace('HKEY_CURRENT_USER','HKCU:')+"\0a0d020000000000c000000000000046"
$AttrName  = "00030398"
$hexified = $YourInput.Split(',') | % { "0x$_"}
New-ItemProperty -Path $RegPath -Name $AttrName -PropertyType Binary -Value ([byte[]]$hexified) -Force
}}
}

#make outlook online, use when needed
function Get-OutlookOffline {
foreach ($prof in Get-ChildItem 'HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles\'){
if (Test-Path ("$prof".Replace('HKEY_CURRENT_USER','HKCU:')+"\0a0d020000000000c000000000000046")){
$YourInput = "02,00,00,00"
$RegPath   = "$prof".Replace('HKEY_CURRENT_USER','HKCU:')+"\0a0d020000000000c000000000000046"
$AttrName  = "00030398"
$hexified = $YourInput.Split(',') | % { "0x$_"}
New-ItemProperty -Path $RegPath -Name $AttrName -PropertyType Binary -Value ([byte[]]$hexified) -Force
}}
}

#create pst with appointments contacts and tasks
Write-Host "Prosze czekac... Tworzenie backupu"
Start-Sleep -s 5
$outlook = New-Object -comobject outlook.application
$namespace = $outlook.GetNameSpace("MAPI")
$calendarFolder = $namespace.GetDefaultFolder(9)
$contactsFolder = $namespace.GetDefaultFolder(10)
$TasksFolder = $namespace.GetDefaultFolder(13)
if (Test-Path 'D:\Biblioteki\Poczta\ost\'){ $PSTPath = 'D:\Biblioteki\Poczta\ost\'}
elseif (Test-Path 'D:\Biblioteki\Poczta\'){ $PSTPath = 'D:\Biblioteki\Poczta\'}
elseif (Test-Path 'D:\Biblioteki\'){ $PSTPath = 'D:\Biblioteki\'}
elseif (Test-Path 'E:\Biblioteki\Poczta\ost\'){ $PSTPath = 'E:\Biblioteki\Poczta\ost\'}
elseif (Test-Path 'E:\Biblioteki\Poczta\'){ $PSTPath = 'E:\Biblioteki\Poczta\'}
elseif (Test-Path 'E:\Biblioteki\'){ $PSTPath = 'E:\Biblioteki\'}
elseif (Test-Path 'F:\Biblioteki\Poczta\ost\'){ $PSTPath = 'F:\Biblioteki\Poczta\ost\'}
elseif (Test-Path 'F:\Biblioteki\Poczta\'){ $PSTPath = 'F:\Biblioteki\Poczta\'}
elseif (Test-Path 'F:\Biblioteki\'){ $PSTPath = 'F:\Biblioteki\'}
elseif (Test-Path 'C:\Windows\'){ $PSTPath = 'C:\'}
elseif (Test-Path 'D:\Windows\'){ $PSTPath = 'D:\'}
elseif (Test-Path 'E:\Windows\'){ $PSTPath = 'E:\'}
$PSTName = 'CalendarContactToDoBackupFromExchange.pst'
$FullPST = $PSTPath + $PSTName
$namespace.AddStore("$FullPST")
$pstFolder = $namespace.Session.Folders.GetLast()
$calendarFolder.CopyTo($pstFolder) | Out-Null
$contactsFolder.CopyTo($pstFolder) | Out-Null
$TasksFolder.CopyTo($pstFolder) | Out-Null
Start-Sleep -s 30
Write-Host "Stworzono plik pst na " $FullPST -ForegroundColor Black -BackgroundColor Green

$outlookprocess = Get-Process OUTLOOK
[void]$outlookprocess.CloseMainWindow()
Write-Host "Prosze czekac..."
!$outlookprocess.WaitForExit(30000) | Out-Null
if (!$outlookprocess.HasExited) {
	$outlookprocess.Kill()
}
Start-Sleep -s 5

#create new outlook profile
if (Test-Path 'HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles'){
New-Item -Path "HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles" -Name 'M365Profile' -Force | Out-Null
New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Outlook" -Name 'DefaultProfile' -PropertyType String -Value M365Profile -Force | Out-Null
New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Outlook" -Name 'ForceOSTPath' -PropertyType ExpandString -Value $PSTPath -Force | Out-Null
}

Start-Sleep -s 5
Start-Process outlook.exe

#wait for user action
$confirmation = "nie"
while (-not (($confirmation -eq "tak") -or ($confirmation -eq "Tak") -or ($confirmation -eq "TAK") -or ($confirmation -eq "t")))
{
	$confirmation = Read-Host "Czy zalogowales juz sie na swoje konto Office 365? [tak/nie]"	
	if (-not (($confirmation -eq "tak") -or ($confirmation -eq "Tak") -or ($confirmation -eq "TAK") -or ($confirmation -eq "t"))) {
		Write-Host "W takim razie czekamy"
		Start-Sleep -s 1
	}
}

#copy appointments contacts and tasks from pst file
Start-Sleep -s 5
$outlook = New-Object -comobject outlook.application
$namespace = $outlook.GetNameSpace("MAPI")
$calendarFolder = $namespace.GetDefaultFolder(9)
$contactsFolder = $namespace.GetDefaultFolder(10)
$tasksFolder = $namespace.GetDefaultFolder(13)
$PSTPath2 = $PSTPath
$PSTName = 'CalendarContactToDoBackupFromExchange.pst'
$FullPST = $PSTPath2 + $PSTName
$namespace.AddStore("$FullPST")
$pstFolder = $namespace.Session.Folders.GetLast()
foreach($pst in $pstFolder.Folders){
if($pst.DefaultMessageClass -eq 'IPM.Appointment'){$indexitems = 1; foreach($item in $pst.Items){$item.CopyTo($calendarFolder,1) | Out-Null; Write-Host $indexitems "/" $pst.Items.Count "appointments coppied"; $indexitems++}}
if($pst.DefaultMessageClass -eq 'IPM.Contact'){$indexitems = 1; foreach($item in $pst.Items){$coppieditem = $item.copy(); $coppieditem.move($contactsFolder) | Out-Null; Write-Host $indexitems "/" $pst.Items.Count "contacts coppied"; $indexitems++}}
if($pst.DefaultMessageClass -eq 'IPM.Task'){$indexitems = 1; foreach($item in $pst.Items){$coppieditem = $item.copy(); $coppieditem.move($tasksFolder) | Out-Null; Write-Host $indexitems "/" $pst.Items.Count "tasks coppied"; $indexitems++}}
}
$namespace.RemoveStore($pstFolder)
Write-Host "Odtworzono dane" -ForegroundColor Black -BackgroundColor Green

Start-Sleep -s 30
