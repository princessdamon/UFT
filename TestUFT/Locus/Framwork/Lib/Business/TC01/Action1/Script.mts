﻿ @@ hightlight id_;_7883356_;_script infofile_;_ZIP::ssf671.xml_;_

'-------------------------------------------------------------------------'
'-------------- TEST SCRIPT Start TC01_FXDeal_DeliveryOption--------------'
'-------------------------------------------------------------------------'
'iRowCount = Datatable.getSheet("TestStep").getRowCount
'DTRow = 1 To DataTable.GlobalSheet.GetRowCount
iRowCount =  Datatable.getSheet("TC01 [TC01]").getRowCount
	'For i =  1 To iRowCount 'Step 1
		'DataTable.LocalSheet.SetCurrentRow(i)
		'TagAction=Ucase(DataTable("Tag",dtglobalsheet))
		User=trim((DataTable("Username","TC01 [TC01]")))
		Password=trim((DataTable("Password","TC01 [TC01]")))
		Typeuser=trim((DataTable("Typeuser","TC01 [TC01]")))
		MainMenu= DataTable("MainMenu","TC01 [TC01]")
		Menuquery= DataTable("Menuquery","TC01 [TC01]")
		Amount=DataTable("Amount","TC01 [TC01]")
		Menu_Product=DataTable("Menu_Product","TC01 [TC01]")
		Menu_Product2=DataTable("Menu_Product2","TC01 [TC01]")
		Contract=DataTable("Contract","TC01 [TC01]")
		Type_Amount=DataTable("Type_Amount","TC01 [TC01]")
		Id_Number=DataTable("Id_Number","TC01 [TC01]")
		Id_BOT=DataTable("Id_BOT","TC01 [TC01]")
		Bill_Race=DataTable("Bill_Race","TC01 [TC01]")
		Set_Method=DataTable("Set_Method","TC01 [TC01]")
		DateDoc=DataTable("DateDoc","TC01 [TC01]")
		NoDoc=DataTable("NoDoc","TC01 [TC01]")
		Converdate=DataTable("Converdate","TC01 [TC01]")

'SystemUtil.Run "D:\Murex\SIT_C2\MX3.exe"
'	wait 60
'JavaWindow("MX.3 - MX_PERF_TEST(PROD)").JavaEdit("txt_user").Set "User"
''
''--------------------------Login - Using Library Function-------------------------'


TransationStart "TC01_FXDeal_DeliveryOption_01_OpenMurexAppClient" 'Start Transation
MR_OpenMurex "client.cmd" 'OpenApp
'MRCheckPage "TC01 [TC01]","MX.3 - MX_PERF_TEST(PROD)"
TransationEnd "TC01_FXDeal_DeliveryOption_01_OpenMurexAppClient"'End Transation

wait 3
TransationStart "TC01_FXDeal_DeliveryOption_02_SubmitLogin" 'Start Transation
MRWebEdit "TC01 [TC01]","MX.3 - MX_PERF_TEST(PROD)","txt_user", User
MRWebEdit "TC01 [TC01]","MX.3 - MX_PERF_TEST(PROD)","txt_password", Password
MRAppButton_Click "TC01 [TC01]","MX.3 - MX_PERF_TEST(PROD)","Sign_in_home"
'MRGettextOnJavaStaticText "MX.3 - MX_PERF_TEST(PROD)" ,"Welcome back, TCOE04!(st)"
TransationEnd "TC01_FXDeal_DeliveryOption_02_SubmitLogin"'End Transation
wait 3
TransationStart "TC01_FXDeal_DeliveryOption_03_ClickStart_iconFO_TH-C" 'Start Transation
MRAppText_Click "TC01 [TC01]","MX.3 - MX_PERF_TEST(PROD)", Typeuser
'MRAppButton_Click "TC01 [TC01]","MX.3 - MX_PERF_TEST(PROD)","okkk"
MRAppButton_Click "TC01 [TC01]","MX.3 - MX_PERF_TEST(PROD)","Start_Login"
MRGettextOnJavaStaticText "TC01 [TC01]","MX.3 - MX_PERF_TEST(PROD)" ,"FO_TH-C"
TransationEnd "TC01_FXDeal_DeliveryOption_03_ClickStart_iconFO_TH-C"'End Transation
wait 3
TransationStart "TC01_FXDeal_DeliveryOption_04_E-TradePad" 'Start Transation
MaximizePage "TC01 [TC01]","MX.3 - MX_PERF_TEST(PROD)"
MRSearchEdit "TC01 [TC01]","MX.3 - MX_PERF_TEST(PROD)","Search_query", MainMenu
MRCheckButtonInternalFrame "TC01 [TC01]","MX.3 - MX_PERF_TEST(PROD)" ,"e-Tradepad" ,"Products"
TransationEnd "TC01_FXDeal_DeliveryOption_04_E-TradePad"'End Transation
wait 3
TransationStart "TC01_FXDeal_DeliveryOption_05_SelectTabMulti" 'Start Transation @@ hightlight id_;_26372280_;_script infofile_;_ZIP::ssf516.xml_;_
MRAppButton_Click_frame "TC01 [TC01]","MX.3 - MX_PERF_TEST(PROD)_MX" ,"e-Tradepad" ,"Folders" @@ hightlight id_;_10857131_;_script infofile_;_ZIP::ssf528.xml_;_
MRSelectMulti "TC01 [TC01]","MX.3 - MX_PERF_TEST(PROD)_MX" ,"e-Tradepad" ,"GuiCellComponentTextFieldStrin"
'MRGettextOnJavaInternalframe "MX.3 - MX_PERF_TEST(PROD)" ,"e-Tradepad" ,"Pricing"
TransationEnd "TC01_FXDeal_DeliveryOption_05_SelectTabMulti"'End Transation
wait 3
NPID = JavaWindow("MX.3 - MX_PERF_TEST(PROD)_MX").GetROProperty("title")
Call WriteLog (NPID,"")

TransationStart "TC01_FXDeal_DeliveryOption_06_SelectProducts" 'Start Transation
MRSelect_Activate "TC01 [TC01]","MX.3 - MX_PERF_TEST(PROD)_MX", "e-Tradepad" , "JfcTreeDynamic" ,Menu_Product2
MRAppButton_Click_frame "TC01 [TC01]","MX.3 - MX_PERF_TEST(PROD)_MX" ,"e-Tradepad" ,"Products"
JavaWindow("MX.3 - MX_PERF_TEST(PROD)_MX").JavaInternalFrame("e-Tradepad").JavaEdit("GuiCellComponentTextFieldStrin").Set "Delivery"
Call Sendkey("{ENTER}")
Window("MUREX_SIT_C2_MX [PID:921936").DblClick 141,197
'JavaWindow("MX.3 - MX_PERF_TEST(PROD)_MX").JavaInternalFrame("e-Tradepad").JavaTree("JfcTreeDynamic").Activate "Products;FX Cash;1. Vanilla;Delivery Option"
'JavaWindow("MX.3 - MX_PERF_TEST(PROD)_MX").JavaInternalFrame("e-Tradepad").JavaTree("JfcTreeDynamic").DblClick @@ hightlight id_;_18_;_script infofile_;_ZIP::ssf659.xml_;_
'MRSelect_Activate "MX.3 - MX_PERF_TEST(PROD)_MX", "e-Tradepad" , "JfcTreeDynamic" ,"Products;FX Cash;1. Vanilla;Delivery Option" 'Menu_Product
TransationEnd "TC01_FXDeal_DeliveryOption_06_SelectProducts"'End Transation
wait 3
Flag = DataTable("Flag","TC01 [TC01]")
'iRowCount =  Datatable.getSheet("TC01 [TC01]").getRowCount
	For i =  Flag To 105	'Step 1
		'DataTable.LocalSheet.SetCurrentRow(i)
		'TagAction=Ucase(DataTable("Tag",dtglobalsheet))
		User=trim((DataTable("Username","TC01 [TC01]")))
		Password=trim((DataTable("Password","TC01 [TC01]")))
		Typeuser=trim((DataTable("Typeuser","TC01 [TC01]")))
		MainMenu= DataTable("MainMenu","TC01 [TC01]")
		Menuquery= DataTable("Menuquery","TC01 [TC01]")
		Amount=DataTable("Amount","TC01 [TC01]")
		Menu_Product=DataTable("Menu_Product","TC01 [TC01]")
		Menu_Product2=DataTable("Menu_Product2","TC01 [TC01]")
		Contract=DataTable("Contract","TC01 [TC01]")
		Type_Amount=DataTable("Type_Amount","TC01 [TC01]")
		Id_Number=DataTable("Id_Number","TC01 [TC01]")
		Id_BOT=DataTable("Id_BOT","TC01 [TC01]")
		Bill_Race=DataTable("Bill_Race","TC01 [TC01]")
		Set_Method=DataTable("Set_Method","TC01 [TC01]")
		DateDoc=DataTable("DateDoc","TC01 [TC01]")
		NoDoc=DataTable("NoDoc","TC01 [TC01]")
		Converdate=DataTable("Converdate","TC01 [TC01]")
		



TransationStart "TC01_FXDeal_DeliveryOption_07_CurrencyPair" 'Start Transation
MRSelectCurrency "TC01 [TC01]","MX.3 - MX_PERF_TEST(PROD)_MX","e-Tradepad","3 legs","#2"
MRChackobjectpage "TC01 [TC01]","MX.3 - MX_PERF_TEST(PROD)_MX" ,"e-Tradepad" ,"JfcTreeDynamic"
MRAppButton_Click_frame "TC01 [TC01]","MX.3 - MX_PERF_TEST(PROD)_MX","e-Tradepad","Contrack"
wait 3
MRWebEdit_Type "TC01 [TC01]","MX.3 - MX_PERF_TEST(PROD)_MX" , "Contract" , "MSearchTextField" , Contract ,"{ENTER}"
MRselectproduct_Contrack "TC01 [TC01]","MX.3 - MX_PERF_TEST(PROD)_MX","Contract","Searching in LABEL"
TransationEnd "TC01_FXDeal_DeliveryOption_07_CurrencyPair"'End Transation
wait 3

TransationStart "TC01_FXDeal_DeliveryOption_08_Cur1" 'Start Transation
MRGettextOnJavaInternalframe "TC01 [TC01]","MX.3 - MX_PERF_TEST(PROD)_MX" ,"e-Tradepad" ,"1 leg(st)"
MRInputAmount "TC01 [TC01]","MX.3 - MX_PERF_TEST(PROD)_MX" ,"e-Tradepad" ,"4 legs" ,"Pricing" ,Amount
MRInputConverdate "TC01 [TC01]","MX.3 - MX_PERF_TEST(PROD)_MX" ,"e-Tradepad" ,"4 legs" ,"Converdate" ,Converdate
TransationEnd "TC01_FXDeal_DeliveryOption_08_Cur1"'End Transation
wait 3

TransationStart "TC01_FXDeal_DeliveryOption_09_IntCounterparty" 'Start Transation
MRInputCounter "TC01 [TC01]","MX.3 - MX_PERF_TEST(PROD)_MX" ,"e-Tradepad" ,"4 legs" ,"ID_Number" ,Id_Number
TransationEnd "TC01_FXDeal_DeliveryOption_09_IntCounterparty" 'End Transation
wait 3

TransationStart "TC01_FXDeal_DeliveryOption_10_BOTpurposecode" 'Start Transation
MRClickGood "TC01 [TC01]","MX.3 - MX_PERF_TEST(PROD)_MX" ,"e-Tradepad" ,"4 legs" ,"Pricing:"
MRAppButton_Click_frame "TC01 [TC01]","MX.3 - MX_PERF_TEST(PROD)_MX" ,"e-Tradepad" ,"Contrack"
TransationEnd "TC01_FXDeal_DeliveryOption_10_BOTpurposecode" 'End Transation
wait 3

TransationStart "TC01_FXDeal_DeliveryOption_11_BOT_PLT" 'Start Transation
MRWebEdit_Type "TC01 [TC01]","MX.3 - MX_PERF_TEST(PROD)_MX" , "BOT_P_LT" , "MSearchTextField" ,Id_BOT ,"{ENTER}"
MRselectproduct_Contrack  "TC01 [TC01]","MX.3 - MX_PERF_TEST(PROD)_MX" , "BOT_P_LT" ,"Searching in _ID_"
TransationEnd "TC01_FXDeal_DeliveryOption_11_BOT_PLT" 'End Transation
wait 3

TransationStart "TC01_FXDeal_DeliveryOption_12_SubmitBook" 'Start Transation
MRAppButton_Click_frame "TC01 [TC01]","MX.3 - MX_PERF_TEST(PROD)_MX" ,"e-Tradepad" ,"Book"
TransationEnd "TC01_FXDeal_DeliveryOption_12_SubmitBook" 'End Transation
wait 3

TransationStart "TC01_FXDeal_DeliveryOption_13_SelectUser_fields" 'Start Transation
MRAppButton_Click_frame "TC01 [TC01]","MX.3 - MX_PERF_TEST(PROD)_MX" ,"Split Package" ,"User fields"
TransationEnd "TC01_FXDeal_DeliveryOption_13_SelectUser_fields" 'End Transation
wait 3

TransationStart "TC01_FXDeal_DeliveryOption_14_SelectLayouts" 'Start Transation
MRselectproduct_Contrack "TC01 [TC01]","MX.3 - MX_PERF_TEST(PROD)_MX" ,"Layouts" , "JfcSimpleTable"
TransationEnd "TC01_FXDeal_DeliveryOption_14_SelectLayouts" 'End Transation
wait 3

TransationStart "TC01_FXDeal_DeliveryOption_15_EditUnititled" 'Start Transation
MRWebEdit_Type "TC01 [TC01]","MX.3 - MX_PERF_TEST(PROD)_MX" ,"Untitled" ,"DMS Approval Doc No." ,NoDoc ,"{ENTER}"
MRWebEdit_Type "TC01 [TC01]","MX.3 - MX_PERF_TEST(PROD)_MX" ,"Untitled" ,"DMS Approval Doc Date" ,DateDoc ,"{ENTER}"
MRAppButton_Click_frame "TC01 [TC01]","MX.3 - MX_PERF_TEST(PROD)_MX" ,"Untitled" ,"OK_Untitled"
TransationEnd "TC01_FXDeal_DeliveryOption_15_EditUnititled" 'End Transation
wait 3



TransationStart "TC01_FXDeal_DeliveryOption_16_SplitPackage" 'Start Transation
MRAppButton_Click_frame "TC01 [TC01]","MX.3 - MX_PERF_TEST(PROD)_MX" ,"Split Package" ,"OK_Split"
TransationEnd "TC01_FXDeal_DeliveryOption_16_SplitPackage" 'End Transation
wait 3



TransationStart "TC01_FXDeal_DeliveryOption_17_Warning" 'Start Transation
MRAppButton_Click_Dinalog "TC01 [TC01]","MX.3 - MX_PERF_TEST(PROD)_MX" ,"Warning" ,"Continue"
TransationEnd "TC01_FXDeal_DeliveryOption_17_Warning"

If JavaWindow("MX.3 - MX_PERF_TEST(PROD)_MX").JavaInternalFrame("Limits Preview").JavaButton("CloseLimit").Exist(5) Then
	JavaWindow("MX.3 - MX_PERF_TEST(PROD)_MX").JavaInternalFrame("Limits Preview").JavaButton("CloseLimit").Click @@ hightlight id_;_19093000_;_script infofile_;_ZIP::ssf673.xml_;_
End If

TransationStart "TC01_FXDeal_DeliveryOption_18_Information(Critical)" 'Start Transation
MRCheckJavaDialog "TC01 [TC01]","MX.3 - MX_PERF_TEST(PROD)_MX" ,"Information" ,"OK_Information"
'MRGetcelltable "MX.3 - MX_PERF_TEST(PROD)_MX" ,"Information" ,"You have 1 notification" ,"0" ,"0"
MRAppButton_Click_Dinalog "TC01 [TC01]","MX.3 - MX_PERF_TEST(PROD)_MX" ,"Information" ,"OK_Information"
TransationEnd "TC01_FXDeal_DeliveryOption_18_Information(Critical)" 'End Transation
wait 3

Flag = Flag + 1
DataTable("Flag","TC01 [TC01]") = Flag 


Next

'-------------------------------------------------------------------------'
TransationStart "TC01_FXDeal_DeliveryOption_19_CloseMurexAppClient" 'Start Transation
MRAppButton_Click "TC01 [TC01]","MX.3 - MX_PERF_TEST(PROD)_MX","toprightbox_close"
MRAppButton_Click_Dinalog "TC01 [TC01]","MX.3 - MX_PERF_TEST(PROD)_MX","MX.3","ButtonYes"
TransationEnd "TC01_FXDeal_DeliveryOption_19_CloseMurexAppClient" 'End Transation
 @@ hightlight id_;_20309096_;_script infofile_;_ZIP::ssf643.xml_;_
 
wait 3
