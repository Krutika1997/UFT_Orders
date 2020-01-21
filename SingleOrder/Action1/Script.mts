
Window("Orders").WinMenu("Menu").Select "Orders;New order...	Ctrl+Ins"

'Verify the 'New Order' Dialog box is displayed.
 If NOT(Window("Orders").Dialog("New order").Exist) Then
	   Reporter.ReportEvent micFail,"Orders Dialog box","Orders Dialog box failed to appear"
 End If
 
 'This is testing Git
