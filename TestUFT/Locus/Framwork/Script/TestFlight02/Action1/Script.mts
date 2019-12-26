iRowCount =  Datatable.getSheet("Action1 [TestFlight02]").getRowCount
For i =  1 To iRowCount 'Step 1
  DataTable.LocalSheet.SetCurrentRow(i)

  FirstName=trim((DataTable("Firstname","Action1 [TestFlight02]")))
  LastName=trim((DataTable("Lastname","Action1 [TestFlight02]")))
  Phone=trim((DataTable("Phone","Action1 [TestFlight02]")))
  Email=trim((DataTable("Email","Action1 [TestFlight02]")))
  Address=trim((DataTable("Address","Action1 [TestFlight02]")))
  City=trim((DataTable("City","Action1 [TestFlight02]")))
  State=trim((DataTable("State","Action1 [TestFlight02]")))
  PostalCode=trim((DataTable("PostalCode","Action1 [TestFlight02]")))
  User=trim((DataTable("User","Action1 [TestFlight02]")))
  Password=trim((DataTable("Password","Action1 [TestFlight02]")))
  ConfirmPassword=trim((DataTable("Password","Action1 [TestFlight02]")))
  
  


LocusClickLink"Welcome: Mercury Tours","Welcome: Mercury Tours","REGISTER"
LocusWebEdit"Welcome: Mercury Tours","Register: Mercury Tours","firstName",FirstName
LocusWebEdit"Welcome: Mercury Tours","Register: Mercury Tours","lastName",LastName
LocusWebEdit"Welcome: Mercury Tours","Register: Mercury Tours","phone",Phone
LocusWebEdit"Welcome: Mercury Tours","Register: Mercury Tours","userName",Email
LocusWebEdit"Welcome: Mercury Tours","Register: Mercury Tours","address1",Address
LocusWebEdit"Welcome: Mercury Tours","Register: Mercury Tours","city",City
LocusWebEdit"Welcome: Mercury Tours","Register: Mercury Tours","state",State
LocusWebEdit"Welcome: Mercury Tours","Register: Mercury Tours","postalCode",PostalCode
'LocusSelect"Welcome: Mercury Tours","Register: Mercury Tours","country"
LocusWebEdit"Welcome: Mercury Tours","Register: Mercury Tours","email",User
LocusWebEdit"Welcome: Mercury Tours","Register: Mercury Tours","password",Password
LocusWebEdit"Welcome: Mercury Tours","Register: Mercury Tours","confirmPassword",ConfirmPassword
LocusClickImg"Welcome: Mercury Tours","Register: Mercury Tours","register"


 @@ script infofile_;_ZIP::ssf20.xml_;_
' @@ script infofile_;_ZIP::ssf38.xml_;_
' @@ script infofile_;_ZIP::ssf39.xml_;_
'Browser("Welcome: Mercury Tours").Page("Register: Mercury Tours").WebEdit("firstName").Set "pang" @@ script infofile_;_ZIP::ssf40.xml_;_
'Browser("Welcome: Mercury Tours").Page("Register: Mercury Tours").WebEdit("lastName").Set "pond" @@ script infofile_;_ZIP::ssf41.xml_;_
'Browser("Welcome: Mercury Tours").Page("Register: Mercury Tours").WebEdit("phone").Set "088888888" @@ script infofile_;_ZIP::ssf42.xml_;_
'Browser("Welcome: Mercury Tours").Page("Register: Mercury Tours").WebEdit("userName").Set "pp@gmail.com" @@ script infofile_;_ZIP::ssf43.xml_;_
'Browser("Welcome: Mercury Tours").Page("Register: Mercury Tours").WebEdit("address1").Set "118" @@ script infofile_;_ZIP::ssf44.xml_;_
'Browser("Welcome: Mercury Tours").Page("Register: Mercury Tours").WebEdit("city").Set "moo" @@ script infofile_;_ZIP::ssf45.xml_;_
'Browser("Welcome: Mercury Tours").Page("Register: Mercury Tours").WebEdit("state").Set "moo" @@ script infofile_;_ZIP::ssf46.xml_;_
'Browser("Welcome: Mercury Tours").Page("Register: Mercury Tours").WebEdit("postalCode").Set "54244" @@ script infofile_;_ZIP::ssf47.xml_;_
'Browser("Welcome: Mercury Tours").Page("Register: Mercury Tours").WebEdit("email").Set "test003" @@ script infofile_;_ZIP::ssf48.xml_;_
'Browser("Welcome: Mercury Tours").Page("Register: Mercury Tours").WebEdit("password").SetSecure "5ddf66f5a938f0b134c9b687fe32bb5e72581c94" @@ script infofile_;_ZIP::ssf49.xml_;_
'Browser("Welcome: Mercury Tours").Page("Register: Mercury Tours").WebEdit("confirmPassword").SetSecure "5ddf6703ef90e53bbcb3c27c7f8beff39d90b985" @@ script infofile_;_ZIP::ssf50.xml_;_
'Browser("Welcome: Mercury Tours").Page("Register: Mercury Tours").Image("register").Click 26,10 @@ script infofile_;_ZIP::ssf51.xml_;_
' @@ hightlight id_;_525312_;_script infofile_;_ZIP::ssf52.xml_;_
next
