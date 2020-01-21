Dim objOutlook 
Dim objOutlookMsg
Dim olMailItem, ReportFilePath

' Create the Outlook object and the new mail object.
Set objOutlook = CreateObject("Outlook.Application") 
Set objOutlookMsg = objOutlook.CreateItem(olMailItem)

' Define mail recipients
objOutlookMsg.To = "pavani@matryxsoft.com" 

'Read the Generated Result path
Set FSO = CreateObject("Scripting.FileSystemObject")
Const ForReading = 1, ForWriting = 2, ForAppending = 8
'Now open file for reading
Set oFile2 = FSO.OpenTextFile("E:\Sample.txt", ForReading, True)
'AtEndOfStream - Returns true if the file pointer is at the end of a TextStream file; false if it is not
Do Until oFile2.AtEndOfStream = True
    ReportFilePath=oFile2.ReadLine
	print ReportFilePath
Loop 
oFile2.Close

' Define a file for attachment
'doc="C:\Users\pavani\AppData\Local\Temp\TempResults\Report\run_results.html"
'doc="E:\run_results.html"

' Body of the message
objOutlookMsg.Subject = "UFT Report" 
objOutlookMsg.Body = "PFA of the UFT Report"

' Add the attachment to the email
objOutlookMsg.Attachments.Add(ReportFilePath)

' Display the email
objOutlookMsg.Display

' Send the message
objOutlookMsg.Send

' Release the objects
set objOutlook = nothing
set objOutlookMsg = nothing

Window("Orders").Close
