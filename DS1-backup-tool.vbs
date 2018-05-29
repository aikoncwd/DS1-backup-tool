'''''''''''''''''''''''''''''''''''''''
'' START CONFIG SECTION - EDIT THIS! ''
'''''''''''''''''''''''''''''''''''''''
	SavegamePath = "C:\Users\YOURUSERNAME\Documents\NBGI\DARK SOULS REMASTERED\YOUR-ID"
	SavegameFile = "DRAKS0005.sl2"

	MaxBackupFiles = 25				'Number of maximun backup savegame files. 25 is ok
	TimeBetweenBackups = 180000		'1000 = 1sec | 180000 = 180secs = 3 minutes

'''''''''''''''''''''''''''''''''''''''
'' END CONFIG SECTION - STOP EDITING ''
'''''''''''''''''''''''''''''''''''''''

Set oFSO = CreateObject("Scripting.FileSystemObject")
Set oWSH = CreateObject("WScript.Shell")
Set oWMI = GetObject("winmgmts:\\.\root\CIMV2")

' Check if the config is OK, there must be a valid savegame on the given path...
If oFSO.FileExists(SavegamePath & "\" & SavegameFile) = False Then
	MsgBox "There is a problem in your config, the file:" & vbCrLf & vbCrLf & SavegamePath & "\" & SavegameFile & vbCrLf & vbCrLf & "is missing!", vbError + vbOkOnly, "DS1 Savegame Backup Tool"
	WScript.Quit
End If

' Check if DARKSOULS is running, if not: execute it!
If isDarksoulsPlaying = False Then
	oWSH.Run "steam://rungameid/570940"
	WScript.Sleep 30000			' Wait 30 seconds until we start the backup routine
End If

counter = 1

Do
	If isDarksoulsPlaying = False Then
		WScript.Quit			' Stop the backup-tool if DARKSOULS is closed
	End If
	oFSO.CopyFile SavegamePath & "\" & SavegameFile, SavegamePath & "\" & counter & "-" & SavegameFile

	counter = counter + 1
	If counter > MaxBackupFiles Then counter = 1	
	WScript.Sleep TimeBetweenBackups	' Wait until we backup the savegame again
Loop

' Retarded routine to check if DARKSOULS is executed
Function isDarksoulsPlaying()
	Set cProcess = oWMI.ExecQuery("Select * from Win32_Process")
	isDarksoulsPlaying = False
	For Each oProcess in cProcess
		If UCase(oProcess.Name) = "DARKSOULSREMASTERED.EXE" Then isDarksoulsPlaying = True
	Next
End Function

' Credits /u/AikonCWD
' https://github.com/aikoncwd/DS1-backup-tool
' Version = 2.0
