Set fso = CreateObject("Scripting.FileSystemObject")
Set WshShell = CreateObject("WScript.Shell")

' Get the folder where this VBS file is located
scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)

' Build the full path to the batch file
batPath = scriptDir & "\start_api_service.bat"

' Run the batch file silently
WshShell.Run """" & batPath & """", 0, False