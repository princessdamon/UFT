
For Iterator = 1 To Datatable.getSheet("Action1 [Testweb]").getRowCount
	DataTable.LocalSheet.SetCurrentRow(Iterator)
	UserName=trim((DataTable("username","Action1 [Testweb]")))
	Password=trim((DataTable("password","Action1 [Testweb]")))

 	LocusWebElement "Google","Advantage Shopping","menuUserSVGPath"
 	LocusWebEdit "Google","Advantage Shopping","username",UserName
 	LocusWebEdit "Google","Advantage Shopping","password",Password
 	LocusWebbutton "Google","Advantage Shopping","sign_in_btnundefined"
 	LocusClickLink "Google","Advantage Shopping","SpeakersCategory"
 	LocusClickImg "Google","Advantage Shopping","fetchImage?image_id=4700"

 	Price = Browser("Google").Page("Advantage Shopping").WebElement("$129.00 SOLD OUT").GetROProperty("innertext")
 	Price = Mid(Price,2,6)
 	Price = CInt(Price)
 	DataTable.Value("price","Action1 [Testweb]") = Price
 	
 	LocusWebbutton "Google","Advantage Shopping","save_to_cart"
 	LocusWebElement "Google","Advantage Shopping","menuCart"

 	Total = Browser("Google").Page("Advantage Shopping").WebElement("$258.00").GetROProperty("innertext")
 	Total = Mid(Total,2,6)
 	Total = CInt(Total)
 	DataTable.Value("Total","Action1 [Testweb]") = Total
 	
 	 If Total = Price Then
 		Status = "Pass"
 	 	DataTable.Value("Status","Action1 [Testweb]") = Status
 	 else
 		Status = "False"
 	 	DataTable.Value("Status","Action1 [Testweb]") = Status
 	 End If
 	
 	LocusWebElement "Google","Advantage Shopping","menuCart"
 	LocusWebElement "Google","Advantage Shopping","WebElement"
 	LocusWebElement "Google","Advantage Shopping","menuCart"
 	LocusWebElement "Google","Advantage Shopping","WebElement"
	LocusClickLink "Google","Advantage Shopping","HOME"
 	LocusWebElement "Google","Advantage Shopping","auyloosizer"
 	LocusClickLink "Google","Advantage Shopping","Link"
 	

Next



'DataTable.Export "D:\Locus\user_Report.CSV"



