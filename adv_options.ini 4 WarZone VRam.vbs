Val 		= "VideoMemoryScale = 1.5"
FileLocation 	= ""	'"C:\Users\kyles\Documents\Call of Duty Modern Warfare\players"

if FileLocation = "" then
	msgbox ("Hello, it seems this is the first time you are using this Script." & vbnewline & vbnewline & _
	"There will be 2 values you need to change in this script. {Val} and {FileLocation}" & vbnewline & vbnewline & _
	"Val is set to {VideoMemoryScale = 1.5} please only change the numer {1.5} to your desired value." & vbnewline & vbnewline & _
	"FileLocation is the directory which your {adv_options.ini} file is in." & vbnewline & vbnewline & _
	"So navigate to the {adv_options.ini} file and Click once by the address, copy and past in this script." & vbnewline & vbnewline & _
	"To do that please right click on this file and click {Edit}" & vbnewline & vbnewline & _
	"Change the value for {FileLocation} by coping the address into the ("""") marks." & vbnewline & vbnewline & _
	"This is an example " & vbnewline & "FileLocation = ""C:\Users\kyles\Documents\Call of Duty Modern Warfare\players""" & vbnewline & vbnewline & _
	"Please do that now.")
	Wscript.Quit
end if

FName = FileLocation & "\adv_options.ini"
Set objFSO = CreateObject("Scripting.FileSystemObject")
If not (objFSO.FileExists(FName)) Then
	msgbox ("The file does not exist!" & vbnewline & vbnewline & "Incorrect directory please check value for {FileLocation} in this script.")
	Wscript.Quit
End If

Const ForReading = 1
Const ForWriting = 2
Dim objFile : Set objFile = objFSO.OpenTextFile(FName, ForReading)
For i = 1 to 2
	objFile.ReadLine
Next
LineThree = objFile.ReadLine
'Wscript.Echo LineThree
objFile.Close

Set objFile = objFSO.OpenTextFile(FName, ForReading)
Dim strText : strText = objFile.ReadAll
objFile.Close

if LineThree = Val then 
	LV = True
	Call Log
	Wscript.Quit
else
	Dim strNewText : strNewText = Replace(strText,"","")
	strNewText = Replace(strNewText,LineThree,Val)
	Set objFile = objFSO.OpenTextFile(FName, ForWriting)
	objFile.WriteLine strNewText
	objFile.Close
	LV = False
	Call Log
end if
'=============================================================================================================
sub Log
ScriptDir = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
LogName = "\Time Log.txt"
LogName = ScriptDir & LogName
DT = (Date & " " & Time)
Set objFSO = CreateObject("Scripting.FileSystemObject")
If not (objFSO.FileExists(LogName)) Then
	Const Hidden = 2
	Set objFile = objFSO.CreateTextFile(LogName, True)
	objFile.Write ".LOG" & vbCrLf
	objFile.WriteLine DT
	objFile.WriteLine "Log File Created."
	objFile.Close
	Set mapFile = objFSO.GetFile(LogName)
	mapFile.Attributes = Hidden
End If
Const ForReading = 1
Const ForWriting = 2
Filename = outFile
Set objFile = objFSO.OpenTextFile(LogName, ForReading)
strText = objFile.ReadAll
objFile.Close

strNewText = Replace(strText, "", ".LOG")
strNewText = Replace(strNewText,  "","")
Set objFile = objFSO.OpenTextFile(LogName, ForWriting)
objFile.WriteLine strNewText
objFile.WriteLine DT
if LV = True then
	objFile.WriteLine "adv_options.ini has been checked, No changes was made to line 3. {" & LineThree & "} = {" & Val & "}"
elseif LV = False then
	objFile.WriteLine "adv_options.ini has been checked, Line 3 was replaced from {" & LineThree & "} to {" & Val & "}"
end if
objFile.Close
end sub
