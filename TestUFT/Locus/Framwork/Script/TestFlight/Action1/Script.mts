iRowCount =  Datatable.getSheet("Action1 [TestFlight]").getRowCount
 'For i =  1 To iRowCount 'Step 1
  'DataTable.LocalSheet.SetCurrentRow(i)
  'TagAction=Ucase(DataTable("Tag",dtglobalsheet))
  User=trim((DataTable("Username","Action1 [TestFlight]")))
  Password=trim((DataTable("Password","Action1 [TestFlight]"))) @@ script infofile_;_ZIP::ssf16.xml_;_
  
  
'  ================= Start Tast =================
 
  
LocusWebEdit "Welcome: Mercury Tours","Welcome: Mercury Tours","userName",User
LocusWebEdit "Welcome: Mercury Tours","Welcome: Mercury Tours","password",Password
LocusClickImg "Welcome: Mercury Tours","Welcome: Mercury Tours","Sign-In"

'Browser("Welcome: Mercury Tours").Page("Welcome: Mercury Tours").WebEdit("userName").Set "princessdamon"
'Browser("Welcome: Mercury Tours").Page("Welcome: Mercury Tours").WebEdit("password").Set "POND7pond8" @@ script infofile_;_ZIP::ssf9.xml_;_
'Browser("Welcome: Mercury Tours").Page("Welcome: Mercury Tours").Image("Sign-In").Click 48,4



