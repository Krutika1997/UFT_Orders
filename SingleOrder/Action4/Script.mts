Dim DirPath
DirPath=Environment.Value("ResultDir")
print(DirPath +"\Report\run_results.html")


'Write into a file
Set FSO = CreateObject("Scripting.FileSystemObject")
Set oFile = FSO.CreateTextFile("E:\Sample.txt",True)
' Writes a specified string to the file
oFile.WriteLine(DirPath +"\Report\run_results.html")
oFile.Close
