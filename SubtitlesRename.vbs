'Use VBScript to rename the Subtitles file name same as video name.
'This simple script in memory of the life I first knew computer program and tried to write it.It contained me much funny memory.

Dim objWs, objFso, objCurrentFolder, array()
ReDim array(100)
Set objWs = createobject("Wscript.shell")
Set objFso = createobject("Scripting.FileSystemObject")

Set objCurrentFolder = objFso.GetFolder(objWs.CurrentDirectory)

i = 0

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

For Each f In objCurrentFolder.Files 
	If IsVideo(objFso.GetExtensionName(f)) = True Then
		'wscript.echo i
		array(i) = f
		i = i + 1
	End If
	'wscript.echo f
Next
ReDim Preserve array(i-1)
For j=0 To UBound(array)
	wscript.echo array(j)
  Next
