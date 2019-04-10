'Use VBScript to rename the Subtitles file name same as video name.
'This simple script in memory of the life I first knew computer program and tried to write it.It contained much funny memory.

Dim objWs, objFso, objCurrentFolder, array(), i, j, FileName
ReDim array(100)
Set objWs = createobject("Wscript.shell")
Set objFso = createobject("Scripting.FileSystemObject")

Set objCurrentFolder = objFso.GetFolder(objWs.CurrentDirectory)

i = 0 
j = 0

Function IsVideo(ExtensionName)
	Select Case ExtensionName
		Case "rmvb"
			IsVideo = True
		Case "mp4"
			IsVideo = True
		Case "mkv"
			IsVideo = True
		Case Else 
			IsVideo = False
	End Select
End Function

Function IsSubtitles(ExtensionName)
	Select Case ExtensionName
		Case "ass"
			IsSubtitles = True
		Case "srt"
			IsSubtitles = True
		Case Else
			IsSubtitles = False
	End Select
End Function

For Each f In objCurrentFolder.Files 
	If IsVideo(objFso.GetExtensionName(f)) = True Then
		FileName = replace(objFso.GetFileName(f),"."&objFso.GetExtensionName(f),"")
		array(i) = FileName
		i = i + 1
	End If
Next
ReDim Preserve array(i-1)

For Each f In objCurrentFolder.Files 
	If IsSubtitles(objFso.GetExtensionName(f)) = True Then
		FileName = array(j) & "." & objFso.GetExtensionName(f)
		f.Name = FileName
		j = j + 1
	End If
Next

Set objFso = Nothing
Set objWs = Nothing
Set objCurrentFolder = Nothing
