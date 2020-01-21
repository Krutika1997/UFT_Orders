Dim iRowCount, i, d_StartDate, d_ExpiryDate

'Convert String to date format wssshile accessing from datatable
d_StartDate=CDate(DataTable("sStartDate", dtGlobalSheet))
d_ExpiryDate=CDate(DataTable("sExpiryDate", dtGlobalSheet))
  
Window("Orders").Dialog("New order").WinComboBox("Product :").Select DataTable("sProduct", dtGlobalSheet)
    
'Type keys to set the StartDate
'Window("Orders").Dialog("New order").WinCalendar("StartDate").SetDate d_StartDate
 Window("Orders").Dialog("New order").WinCalendar("StartDate").Click
 Window("Orders").Dialog("New order").WinCalendar("StartDate").Type(d_StartDate)
    
 Window("Orders").Dialog("New order").WinEdit("txt_Customer Name:").Set DataTable("sCustomerName", dtGlobalSheet)
 Window("Orders").Dialog("New order").WinEdit("txt_Street:").Set DataTable("sStreet", dtGlobalSheet)
 Window("Orders").Dialog("New order").WinEdit("txt_City:").Set DataTable("sCity", dtGlobalSheet)
 Window("Orders").Dialog("New order").WinEdit("State:").Set DataTable("sState", dtGlobalSheet)
 Window("Orders").Dialog("New order").WinEdit("txt_Zip:").Set DataTable("sZip", dtGlobalSheet)

 'Update the text property to the datatable value to select the radio button
 Window("Orders").Dialog("New order").WinRadioButton("rbn_Card").SetTOProperty "Text", DataTable("sCard",dtGlobalSheet) 
 Window("Orders").Dialog("New order").WinRadioButton("rbn_Card").Set

 'Type keys to set the ExpirationDate 
 Window("Orders").Dialog("New order").WinCalendar("ExpirationDate").Click
 Window("Orders").Dialog("New order").WinCalendar("ExpirationDate").Type(d_ExpiryDate)
   
 Window("Orders").Dialog("New order").WinButton("btn_OK").Click

 
 
 





 

  
  





