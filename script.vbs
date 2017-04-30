
Set FSO = CreateObject("Scripting.FileSystemObject")
Set args = Wscript.Arguments
count =0
For Each arg In args
  count = count+1
Next

If count > 0 then
	Wscript.Echo args(0)
	folder_name = args(0)
else
	Wscript.Echo args
	folder_name = args("C:\Users\USER\Desktop")'default folder
End if

Set replace_name = ""
If count > 1 then
	Wscript.Echo args(1)
	replace_name = args(1)
else
	Wscript.Echo args
	replace_name = args(".id-B84DC38D.[mk.goro@aol.com].wallet")'default folder
End if


Wscript.Echo "Opening " + folder_name
ShowSubfolders FSO.GetFolder(folder_name), 3 


Sub ShowSubFolders(Folder, Depth)
	Set Folder = FSO.GetFolder(Folder)


	For Each File In Folder.Files
		sNewFile = File.Name
		sNewFile = Replace(sNewFile,replace_name,"")
		if (sNewFile<>File.Name) then 
			Wscript.Echo "Fixing file name of " + File.Name
			File.Move(File.ParentFolder+"\"+sNewFile)
		end if

	Next
    If Depth > 0 then
        For Each Subfolder in Folder.SubFolders
            Wscript.Echo Subfolder.Path
            ShowSubFolders Subfolder, Depth -1 
        Next
    End if
End Sub