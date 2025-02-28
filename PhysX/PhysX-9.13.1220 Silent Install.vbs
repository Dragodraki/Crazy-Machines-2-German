Option Explicit
On Error Resume Next


' Nicht als Variable definieren: wscript, ScriptFullName !!!
' Für Textbox statt "wscript.echo" den Befehl "msgbox" verwenden !!!
' Der msgbox/popup-Parameter +131072 erhöht zwar die Textauflösung, aber verhindert die nach definierter Zeit Schließen-Funktion!


Dim fso, warningbox, file, folder, folder2, wshell, wshshell, objShell, myKey1, myKey2, strPfad, Systemdrive, Username, AppData, oFSO, oFolder, sDirectoryPath, oFileCollection, oFile, sDir, iDaysOld, sExt, text, text2, fcreate, fdesc, WshSheg, WshShell1, WshShell2, WshShell3, ws, oShell, oWindowList, EXPLORERUMGEBUNGSVARIABLE, EXPLORERUMGEBUNGSVARIABLE2, oWindow, KeinExplorer, user, objFSO, RIGHTHERE
Dim objAllUsersProgramsFolder, strAllUsersProgramsPath, str, objfolder, objFolderItem, colVerbs, objVerb, objShortcut, objFileSystem, sRegFile, TEST, owindowsFolderPath, newDIR, ziel, MyComputer, strComputer, objWMIService, colOperatingSystems, objOperatingSystem, colProcess, objProcess, strProcess, strWMIQuery
Dim sProcessName, sComputer, oWmi, colProcessList, Repeater, objFile, objReg, objSystemInfo, strHostname, startexe, objstream, strArgs, target_folder, fs, TEMP, USERPROFILE, ALLUSERSPROFILE, LOCALAPPDATA, SYSTEMROOT, RIGHTHEREnoBACKSLASH, desktopPath, link, filesys, PARENT_RIGHTHERE, strRoot, strModify, shell, FULLSCRIPTNAME, SCRIPTFILENAME, key, output, gameexe, RIGHTHEREDoubleSlash, XPmode, exelostbox, KompatibilitaetValue, ScriptingHostProcessName, oReg, strKeyPath, linkname, linkname2, linkname3, linkname4, linkname5, PARENT_RIGHTHEREDoubleSlash, search, skriptfilename, iconpfad, strValueName, dwValue, AlleRunasRobEintraege, NeueRunasRobEintraege, PROGRAMDATA
dim ResetteRegistrySchluessel, NeuerStringName_1, NeuerStringWert_1, NeuerStringName_2, NeuerStringWert_2, NeuerStringName_3, NeuerStringWert_3, NeuerStringName_4, NeuerStringWert_4, NeuerStringName_5, NeuerStringWert_5, NeuerDwordName_1, NeuerDwordWert_1, NeuerDwordName_2, NeuerDwordWert_2, NeuerDwordName_3, NeuerDwordWert_3, NeuerDwordName_4, NeuerDwordWert_4, NeuerDwordName_5, NeuerDwordWert_5, Aktiviere_HKCU, Aktiviere_HKLM, strValue, Scriptplayer, DeleteAproDings, strContents, strFirstLine, strNewContents, Aktiviere_ALL_HKCU, oNetwork, NeuerStringWert_1A, NeuerStringWert_2A, NeuerStringWert_3A, NeuerDwordWert_1A, NeuerDwordWert_2A, NeuerDwordWert_3A, drive, GetDriveLetter, label, labelfilename ,mount_disc, GetDriveLetterRaw, Resolution, RIGHTHEREDoubleSlashNoBackslash, WindowsVersionName, WMI, OSs, OS, dtmConvertedDate, objOperatingSyst, value, NewStringWindowsVersion, MajorWindowsVersion, objArgsPScript, objArgsWScript, objArgsCScript, objArgs, MitParameterGestartet, ArgumentTotal, Arg, f, app, ProzessUNDTitelOffen, ProzessUNDPfadOffen, ExternalPathCheck, TitleCheck, ProzessEgal_PfadTitle_Warten, AufAlleProzesseVBSpfadWarten, TempFolderName, BenutzerAusListe, infobox, USERDOMAIN, LASTSHOWEDUSER
dim LastResolutionShowed, AllResolutionList, AndereAufloesung, Mindestbreite, Mindesthoehe, Seitenverhaeltnis, Result, Aufloesungstest, Horizontale, Vertikale, AufloesungCheckpoint, AufloesungAlternative, AufloesungAlternativeCheckpoint
dim FullPath, arr, dir, path, Counter, pReg, qReg, rReg, sReg, objRegistry, sSubkey, tSubkey, uSubkey, oSubkey, pSubkey, aSubKeys, bSubKeys, cSubKeys, dSubKeys, eSubKeys, sTmpValueName, sTmpKeyName1, sTmpKeyName2, sTmpKeyName3, values, Return, Log, logpath, logSubkey, logvalues, bArray, BinaryPositionNumberToChange, BinaryReplacement, AMDPath, ScalingProzedurNotwendig
dim ColOSs, OSType, service, AppName, Process, AppName2, strDelete, arrProcesses, arrWindowTitles, intProcesses, strProcList, TITLENAME2KILL, Informationsbox, strApplication, colItems, ProcessName, Apptitle, objApp, PID, ProcessPfad, nPos, strList, p, q, strKonto, strFileName, sKeyPath, NvidiaTest_1, NvidiaTest_2, objRegEx, strSearchString, colMatches, VMCheck, Action, Obj, Ver, Respond, MaxFrameworkVersionCheck, Regstry, RegstrySP, ProcessCalledFromCD, Win10OrHigherMode, VerzeichnisAbfrage, PureUsername, FullNameArray, objItem, StikyNotRestart, ActualUserSID, HKEY_CURRENT_USER, ACCOUNTNAME
dim InstallationFile, MSIAddition, ScriptingHostDirectory, InstallationParameter, sQuery


Set oShell = CreateObject("WScript.Shell")
Set oFSO = CreateObject("Scripting.FileSystemObject")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Const COMPUTER = "."
Const wbemFlagReturnImmediately = &h10
Const wbemFlagForwardOnly = &h20
Const ComputerName = "."


'--- ERKENNE WINDOWS-VERSION (BEI WINDOWS XP GEHE IN DEN XP-MODUS)

' Variante A - Prüfe per Versionsnummer (eigentlich sicheres Verfahren, solange OS-Nummer nicht größer als 100 Millionen ist)
Set dtmConvertedDate = CreateObject("WbemScripting.SWbemDateTime")
strComputer = "."
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set oss = objWMIService.ExecQuery ("Select * from Win32_OperatingSystem")
Set dtmConvertedDate = CreateObject("WbemScripting.SWbemDateTime")
strComputer = "."
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set oss = objWMIService.ExecQuery ("Select * from Win32_OperatingSystem")
For Each os in oss
  value = os.Version
NewStringWindowsVersion="."
MajorWindowsVersion=left(value, instr(value, NewStringWindowsVersion)-1)
'msgbox MajorWindowsVersion
MajorWindowsVersion = MajorWindowsVersion + 100000006
if MajorWindowsVersion < "100000012" then
XPmode = "Ja"
'msgbox "Windows XP"
else
end if
Next
'msgbox "Ergebnis der Variante A: " & XPmode


' Variante B - Prüfe per Namen (wenn hier Windows 7/8/10 vorliegt, überschreibe Ergebnis der Variante A)
if (instr(MyComputer, "SV") OR (instr(MyComputer, "TB"))) then
		Else	
		strComputer = "."
		Set objWMIService = GetObject("winmgmts:" _
			& "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

		Set colOperatingSystems = objWMIService.ExecQuery _
			("Select * from Win32_OperatingSystem")
'Windows 7
		For Each objOperatingSystem in colOperatingSystems
			  if instr(objOperatingSystem.Caption, "7") Then
			WindowsVersionName = "Windows 7"
			XPmode = "Nein"
			'msgbox "Win 7 detected"
			Else
			End If
		Next
'Windows 8
		For Each objOperatingSystem in colOperatingSystems
			  if instr(objOperatingSystem.Caption, "8") Then
			WindowsVersionName = "Windows 8"
			XPmode = "Nein"
			'msgbox "Win 8 detected"
			Else
			End If
		Next
'Windows 10
		For Each objOperatingSystem in colOperatingSystems
			  if instr(objOperatingSystem.Caption, "10") Then
			WindowsVersionName = "Windows 10"
			XPmode = "Nein"
			'msgbox "Win 10 detected"
			Else
			End If
			Next
'Windows XP
		For Each objOperatingSystem in colOperatingSystems
			  if instr(objOperatingSystem.Caption, "XP") Then
			XPmode = "Ja"
			WindowsVersionName = "Windows XP"
			'msgbox "Win XP detected"
			Else
			End If
		Next
end if
' msgbox "Ergebnis der Variante B: " & XPmode



'--- ERKENNE WINDOWS-VERSION (AB WINDOWS 10 GEHE IN DEN 10+-MODUS)

' Variante A - Prüfe per Versionsnummer (eigentlich sicheres Verfahren, solange OS-Nummer nicht größer als 100 Millionen ist)
Set dtmConvertedDate = CreateObject("WbemScripting.SWbemDateTime")
strComputer = "."
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set oss = objWMIService.ExecQuery ("Select * from Win32_OperatingSystem")
Set dtmConvertedDate = CreateObject("WbemScripting.SWbemDateTime")
strComputer = "."
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set oss = objWMIService.ExecQuery ("Select * from Win32_OperatingSystem")
For Each os in oss
  value = os.Version
NewStringWindowsVersion="."
MajorWindowsVersion=left(value, instr(value, NewStringWindowsVersion)-1)
'msgbox MajorWindowsVersion
MajorWindowsVersion = MajorWindowsVersion + 100000010
if MajorWindowsVersion >= "100000020" then
Win10OrHigherMode = "Ja"
'msgbox "Windows 10 oder höher"
else
end if
Next
'msgbox "Ergebnis der Variante A: " & Win10OrHigherMode

' Variante B - Prüfe per Namen (wenn hier Windows 7/8/10 vorliegt, überschreibe Ergebnis der Variante A)
if (instr(MyComputer, "SV") OR (instr(MyComputer, "TB"))) then
		Else	
		strComputer = "."
		Set objWMIService = GetObject("winmgmts:" _
			& "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

		Set colOperatingSystems = objWMIService.ExecQuery _
			("Select * from Win32_OperatingSystem")
'Windows 7
		For Each objOperatingSystem in colOperatingSystems
			  if instr(objOperatingSystem.Caption, "7") Then
			WindowsVersionName = "Windows 7"
			Win10OrHigherMode = "Nein"
			'msgbox "Win 7 detected"
			Else
			End If
		Next
'Windows 8
		For Each objOperatingSystem in colOperatingSystems
			  if instr(objOperatingSystem.Caption, "8") Then
			WindowsVersionName = "Windows 8"
			Win10OrHigherMode = "Nein"
			'msgbox "Win 8 detected"
			Else
			End If
		Next
'Windows 10
		For Each objOperatingSystem in colOperatingSystems
			  if instr(objOperatingSystem.Caption, "10") Then
			WindowsVersionName = "Windows 10"
			Win10OrHigherMode = "Ja"
			'msgbox "Win 10 detected"
			Else
			End If
			Next
'Windows 11
		For Each objOperatingSystem in colOperatingSystems
			  if instr(objOperatingSystem.Caption, "11") Then
			WindowsVersionName = "Windows 11"
			Win10OrHigherMode = "Ja"
			'msgbox "Win 10 detected"
			Else
			End If
			Next
'Windows XP
		For Each objOperatingSystem in colOperatingSystems
			  if instr(objOperatingSystem.Caption, "XP") Then
			Win10OrHigherMode = "Nein"
			WindowsVersionName = "Windows XP"
			'msgbox "Win XP detected"
			Else
			End If
		Next
end if
'msgbox "Ergebnis der Variante B: " & Win10OrHigherMode






'--- DEFINIERE UMGEBUNGSVARIABLEN

SET WshShell = CreateObject( "WScript.Shell" )
Userdomain = WshShell.Environment("Process")("COMPUTERNAME")

' Lese aktuellen Account-Namen aus (WICHTIG: Im Falle einer Namensumbenennung wären alle daraus abgeleiteten Umgebungsvariablen falsch)
Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_ComputerSystem",,48)
For Each objItem in colItems
      'nur  mal anzeigen, damit Du siehst was passiert
    'msgbox "UserName: " & objItem.UserName
      'String am Backslash teilen'
    FullNameArray=Split(objItem.UserName,"\")
      ' 2. Arrayfeld  anzeigen'
    ACCOUNTNAME = FullNameArray(1)
    ' msgbox  "Accountname is : " & ACCOUNTNAME
Next

' Lese mithilfe USERDOMAIN und ACCOUNTNAME die aktuelle Benutzer-SID aus
Set objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_UserAccount "&_
       "WHERE (Domain="""&USERDOMAIN&""" AND Name="""&ACCOUNTNAME&""")", "WQL", wbemFlagReturnImmediately + wbemFlagForwardOnly)
For Each objItem In colItems
     ActualUserSID = objItem.SID
Next
If IsNull(ActualUserSID) Or IsEmpty(ActualUserSID) Or ActualUserSID= "" Then
' msgbox "Failed"
Else
' msgbox ACCOUNTNAME & "'s SID ist: " & ActualUserSID
End if

' Lese mithilfe der aktuellen Benuzter-SID den Registrywert -ProfileImagePath- aus und extrahiere dessen Benutzernamen
strComputer = "."
Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")
strKeyPath = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\" & ActualUserSID
strValueName = "ProfileImagePath"
oReg.GetStringValue &H80000002,strKeyPath,strValueName,strValue
' msgbox strValue
PureUsername = right(strValue,instr(strreverse(strValue),"\")-1)
' msgbox PureUsername

SET WshShell = CreateObject ("WScript.Shell")
Systemdrive = WshShell.Environment("Process")("SYSTEMDRIVE")
Set objShell = CreateObject("WScript.Shell")
objShell.Environment("Process")("USERNAME") = PureUsername
SET WshShell = CreateObject( "WScript.Shell" )
Username = WshShell.Environment("Process")("USERNAME")
'SET WshShell = CreateObject( "WScript.Shell" )
Userprofile = WshShell.Environment("Process")("USERPROFILE")
Userprofile=Replace(Userprofile,"\WINDOWS\SysWOW64\config\systemprofile","\Users\" & PureUsername)
Userprofile=Replace(Userprofile,"\WINDOWS\system32\config\systemprofile","\Users\" & PureUsername)
Userprofile=Replace(Userprofile,"\WINDOWS\config\systemprofile","\Users\" & PureUsername)
Userprofile=Replace(Userprofile,"\Windows\syswow64\config","\Users")
Userprofile=Replace(Userprofile,"\Windows\system32\config","\Users")
Userprofile=Replace(Userprofile,"\Windows\config","\Users")
objShell.Environment("Process")("USERPROFILE") = CreateObject("Scripting.FileSystemObject").GetParentFolderName(USERPROFILE) & "\" & PureUsername
SET WshShell = CreateObject( "WScript.Shell" )
Userprofile = WshShell.Environment("Process")("USERPROFILE")
SET objShell = CreateObject("Shell.Application")
if LASTSHOWEDUSER = "" then
LASTSHOWEDUSER = "Administrator"
end if
if XPmode = "Ja" then
SET WshShell = CreateObject ("WScript.Shell")
AppData = objShell.Namespace(&H1a&).Self.Path
Set objShell = CreateObject("WScript.Shell")
APPDATA = Replace(APPDATA,LASTSHOWEDUSER,PureUsername)
Set objShell = CreateObject("WScript.Shell")
objShell.Environment("Process")("APPDATA") = APPDATA
SET WshShell = CreateObject( "WScript.Shell" )
AppData = WshShell.Environment("Process")("APPDATA")
else
Set objShell = CreateObject("WScript.Shell")
objShell.Environment("Process")("APPDATA") = SYSTEMDRIVE & "\Users\" & PureUsername & "\AppData\Roaming"
SET WshShell = CreateObject( "WScript.Shell" )
AppData = WshShell.Environment("Process")("APPDATA")
end if
SET WshShell = CreateObject ("WScript.Shell")
Allusersprofile = WshShell.Environment("Process")("ALLUSERSPROFILE")
SET objShell = CreateObject("Shell.Application")
ProgramData = objShell.Namespace(&H23&).Self.Path	' ab Vista mit ALLUSERSPROFILE identisch
SET objShell = CreateObject("Shell.Application")
Localappdata = objShell.Namespace(&H1c&).Self.Path	' so ist LOCALAPPDATA auch in Windows XP möglich :)
Localappdata=Replace(Localappdata,SYSTEMDRIVE & "\Windows\SysWOW64\config\systemprofile",USERPROFILE)
Localappdata=Replace(Localappdata,SYSTEMDRIVE & "\Windows\system32\config\systemprofile",USERPROFILE)
Localappdata=Replace(Localappdata,SYSTEMDRIVE & "\Windows\config\systemprofile",USERPROFILE)
if XPmode = "Ja" then
Localappdata=Replace(Localappdata,"Administrator",PureUsername)
Localappdata=Replace(Localappdata,PureUsername,PureUsername)	'erst bei Auflistung aller Konten 2."PureUsername" abändern
else
LOCALAPPDATA=CreateObject("Scripting.FileSystemObject").GetParentFolderName(APPDATA) & "\Local"
end if
Localappdata=Replace(Localappdata,"Administrator",PureUsername)
Set objShell = CreateObject("WScript.Shell")
objShell.Environment("Process")("LOCALAPPDATA") = LOCALAPPDATA
SET WshShell = CreateObject ("WScript.Shell")
Temp = WshShell.Environment("Process")("TEMP")		' TEMP-Ordner für Windows 7,8,10
Set objShell  = CreateObject("Scripting.FileSystemObject")
Temp = CreateObject("Scripting.FileSystemObject").GetParentFolderName(LOCALAPPDATA)
Temp = TEMP & "\Temp"				' TEMP-Ordner für Windows XP (wechselt automatisch bei höherem OS)
if TEMP = USERPROFILE & "\AppData\Temp" then
' msgbox "Laut Variableninhalt handelt es sich um Windows > XP und daher wird die TEMP-Umgebungsvariable angepasst."
Set objShell = CreateObject("WScript.Shell")
objShell.Environment("Process")("TEMP") = LOCALAPPDATA & "\Temp"
SET WshShell = CreateObject( "WScript.Shell" )
TEMP = WshShell.Environment("Process")("TEMP")	' TEMP-Ordner für Windows 7,8,10
else
end if
SET WshShell = CreateObject ("WScript.Shell")
Systemroot = WshShell.Environment("Process")("SYSTEMROOT")
SET WshShell = CreateObject ("WScript.Shell")
FULLSCRIPTNAME = ScriptFullName
if FULLSCRIPTNAME = "" then 
FULLSCRIPTNAME = Wscript.ScriptFullName
end if
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.GetFile(ScriptFullName)
SCRIPTFILENAME = objFSO.GetFileName(objFile)
if SCRIPTFILENAME = "" then 
SCRIPTFILENAME = Wscript.ScriptName
end if
set WshShell = Createobject("wscript.shell")
RIGHTHERE = WshShell.CurrentDirectory & "\"
sQuery = CreateObject("Scripting.FileSystemObject").GetParentFolderName(FULLSCRIPTNAME)
SET objShell = CreateObject("Shell.Application")
RIGHTHERE=Replace(RIGHTHERE, SYSTEMDRIVE & "\Windows\system32", sQuery)
if XPmode = "Ja" then
search="\" & SCRIPTFILENAME
RIGHTHERE=left (FULLSCRIPTNAME, instr(FULLSCRIPTNAME, search)-0)
end if
SET WshShell = CreateObject ("WScript.Shell")
RIGHTHEREnoBACKSLASH = Left(RIGHTHERE, Len(RIGHTHERE) -1)
Set filesys = CreateObject("Scripting.FileSystemObject")
PARENT_RIGHTHERE = CreateObject("Scripting.FileSystemObject").GetParentFolderName(RIGHTHERE)
RIGHTHEREDoubleSlash=Replace(RIGHTHERE,"\","\\")
PARENT_RIGHTHEREDoubleSlash=Replace(PARENT_RIGHTHERE,"\","\\")
SET WshShell = CreateObject ("WScript.Shell")
RIGHTHEREDoubleSlashNoBackslash = Left(RIGHTHEREDoubleSlash, Len(RIGHTHEREDoubleSlash) -2)
' LASTSHOWEDUSER darf bis zur nächsten Schleife nicht mehr geändert werden!
LASTSHOWEDUSER = USERNAME

'infobox = WshShell.Popup("USERDOMAIN: " & USERDOMAIN & vbCrLf  & "ACCOUNT: " & ACCOUNTNAME & vbCrLf  & "USERNAME: " & USERNAME & vbCrLf  & "SYSTEMDRIVE: " & SYSTEMDRIVE & vbCrLf  & "USERPROFILE: " & USERPROFILE  & vbCrLf  & "ALLUSERSPROFILE: " & ALLUSERSPROFILE & vbCrLf  & "PROGRAMDATA: " & PROGRAMDATA & vbCrLf  & "APPDATA: " & APPDATA & vbCrLf  & "LOCALAPPDATA: " & LOCALAPPDATA & vbCrLf  & "TEMP: " & TEMP & vbCrLf & "SYSTEMROOT: " & SYSTEMROOT  & vbCrLf  & "FULLSCRIPTNAME: " & FULLSCRIPTNAME & vbCrLf & "SCRIPTFILENAME: " & SCRIPTFILENAME  & vbCrLf & "RIGHTHERE: " & RIGHTHERE & vbCrLf  & "RIGHTHEREnoBACKSLASH: " & RIGHTHEREnoBACKSLASH & vbCrLf & "PARENT_RIGHTHERE: " & PARENT_RIGHTHERE & vbCrLf & "RIGHTHEREDoubleSlash: " & RIGHTHEREDoubleSlash & vbCrLf & "PARENT_RIGHTHEREDoubleSlash: " & PARENT_RIGHTHEREDoubleSlash, 0, "Alle Variablen", 0+64+4096+65536)


'--------------------------------------------------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------------------------------------------
ScriptingHostDirectory = PARENT_RIGHTHERE & "\"
MSIAddition = "msiexec.exe /i"
InstallationFile = "PhysX-9.13.1220.msi"
InstallationParameter = "/quiet /norestart /qn"
'--------------------------------------------------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------------------------------------------



' Prüfe, ob Installationsprogramm im gleichen Ordnerpfad liegt - wenn nicht, breche Skript ab...
Set fso = CreateObject("Scripting.FileSystemObject")
If (fso.FileExists(RIGHTHERE & InstallationFile)) Then
  ' msgbox("File exists!")





If CSI_IsAdmin Then output = "YES"
If NOT CSI_IsAdmin Then output = "NO"
' msgbox output	'Message-Box, ob ich Administrator-Privilegien habe oder nicht (YES=Adminrechte, NO=Userrechte)
Function CSI_IsAdmin()
  CSI_IsAdmin = False
  On Error Resume Next
  key = CreateObject("WScript.Shell").RegRead("HKEY_USERS\S-1-5-19\Environment\TEMP")
  If err.number = 0 Then CSI_IsAdmin = True
End Function

if not output = "YES" then
Set fso = CreateObject("Scripting.FileSystemObject") 
warningbox = MsgBox("Um PhysX auf dem Rechner zu installieren, benötigt das Skript Administrator-Privilegien. Wollen Sie fortfahren?",4+64+4096+256+65536+131072,InstallationFile)


if warningbox = vbYes then
Set WSHShell = CreateObject("WScript.Shell") 
strDesktop = WSHShell.SpecialFolders("Desktop") 
WSHShell.AppActivate strDesktop
WSHShell.SendKeys "{F13}"
Set Shell = CreateObject("Shell.Application")
ArgumentTotal = " " & ArgumentTotal
Shell.ShellExecute SYSTEMROOT & "\system32\wscript.exe", Chr(34) & FULLSCRIPTNAME & Chr(34) & ArgumentTotal, , "runas", 1
Set objWScriptShell = Nothing
End if

else
end if	'output muss nicht mehr YES sein (also für den Fall, dass Skript bereits per Adminrechte läuft, nehme an, dass Konfiguration ohne weitere Abfrage durchgeführt weden soll)






if warningbox = vbNo then
'Aktualisiere den Explorer durch Löschen des Icon-Cache, um Aufgabe zu beenden
Set wshshell = WScript.CreateObject ("wscript.shell")
else
if output = "YES" then

RequireAdmin





'------------------------------------

Set objShell = CreateObject("WScript.Shell")
sRegFile = RIGHTHERE & InstallationFile
objShell.Run MSIAddition & " " & Chr(34) & sRegFile & Chr(34) & " " & InstallationParameter, 1, True

'------------------------------------


else
end if	' Ab hier ist es nicht mehr wichtig, dass das Skript mit Adminrechten läuft (output muss nicht mehr = YES sein)

End if	' Ab hier werden Befehle unabhängig der Antwort aus der MessageBox ausgeführt

else
warningbox = MsgBox("Kann Setup ''" & InstallationFile & "'' nicht finden!",0+16+4096+256+65536+131072,"Installations-Datei fehlt")
end if	' egal, ob -InstallationFile- im selben Ordner wie VBS existiert oder nicht
