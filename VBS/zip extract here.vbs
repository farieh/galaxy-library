Dim WshShell, strCurDir
Set WshShell = CreateObject("WScript.Shell")
strCurDir    = WshShell.CurrentDirectory

'The location of the zip file.
ZipFile=strCurDir + "\Visual Studio 2015.zip"
'The folder the contents should be extracted to.
ExtractTo=strCurDir

'If the extraction location does not exist create it.
Set fso = CreateObject("Scripting.FileSystemObject")
If NOT fso.FolderExists(ExtractTo) Then
   fso.CreateFolder(ExtractTo)
End If

'Extract the contants of the zip file.
set objShell = CreateObject("Shell.Application")
set FilesInZip=objShell.NameSpace(ZipFile).items
objShell.NameSpace(ExtractTo).CopyHere(FilesInZip)
Set fso = Nothing
Set objShell = Nothing

Set obj = CreateObject("Scripting.FileSystemObject") 'Calls the File System Object
obj.DeleteFile(ZipFile) 'Deletes the file throught the DeleteFile function
