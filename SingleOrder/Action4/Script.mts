Dim DirPath
DirPath=Environment.Value("ResultDir")
print(DirPath +DataTable("ResultDir", dtGlobalSheet))

'Write into a file
Set FSO = CreateObject("Scripting.FileSystemObject")
Set oFile = FSO.CreateTextFile(DataTable("TxtFilePath", dtGlobalSheet),True)
' Writes a specified string to the file
oFile.WriteLine(DirPath +DataTable("ResultDir", dtGlobalSheet))
oFile.Close
