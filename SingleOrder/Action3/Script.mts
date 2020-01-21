
Window("Orders").WinMenu("Menu").Select "File;Save As..."
Window("Orders").Dialog("Save").WinEdit("File name:").Set DataTable("FileName", dtGlobalSheet)
Window("Orders").Dialog("Save").WinButton("Save").Click

If (Dialog("Confirm Save As").Exist) Then
	Dialog("Confirm Save As").WinButton("Yes").Click
End If

Window("Orders").Close










