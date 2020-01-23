Dim objOutlook 
Dim objOutlookMsg
Dim olMailItem, ReportFilePath

' Create the Outlook object and the new mail object.
Set objOutlook = CreateObject("Outlook.Application") 
Set objOutlookMsg = objOutlook.CreateItem(olMailItem)

' Define mail recipients
objOutlookMsg.To =DataTable("Email_To", dtGlobalSheet)

'Read the Generated Result path
Set FSO = CreateObject("Scripting.FileSystemObject")
Const ForReading = 1, ForWriting = 2, ForAppending = 8
'Now open file for reading
Set oFile2 = FSO.OpenTextFile(DataTable("FilePath", dtGlobalSheet), ForReading, True)

'AtEndOfStream - Returns true if the file pointer is at the end of a TextStream file; false if it is not
Do Until oFile2.AtEndOfStream = True
    ReportFilePath=oFile2.ReadLine
	print ReportFilePath
Loop 
oFile2.Close

' Body of the message
objOutlookMsg.Subject = DataTable("Email_Subject", dtGlobalSheet)
objOutlookMsg.Body = DataTable("Email_Body", dtGlobalSheet)

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
