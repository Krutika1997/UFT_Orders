Dim objOutlook 
Dim objOutlookMsg
Dim olMailItem

' Create the Outlook object and the new mail object.
Set objOutlook = CreateObject("Outlook.Application") 
Set objOutlookMsg = objOutlook.CreateItem(olMailItem)

' Define mail recipients
objOutlookMsg.To = "pavani@matryxsoft.com" 

' Define a file for attachment
doc="C:\Users\pavani\AppData\Local\Temp\TempResults\Report\run_results.html"
'doc="E:\run_results.html"

' Body of the message
objOutlookMsg.Subject = "UFT Report sample" 
objOutlookMsg.Body = "This is a test"

' Add the attachment to the email
objOutlookMsg.Attachments.Add(doc)

' Display the email
objOutlookMsg.Display

' Send the message
objOutlookMsg.Send

' Release the objects
set objOutlook = nothing
set objOutlookMsg = nothing

Window("Orders").Close
