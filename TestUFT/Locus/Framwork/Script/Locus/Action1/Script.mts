  ' Get date time for execute all test



FW_StartExecute

StrTag="Run"
'********************************************************************************
'	Initial Start Run
'	Create Initial FrameWork	(Define Setup Parameter & Import TestData)
FW_Initial	

If StrTag="" Then
		StrTag="Run"
	End If
	Tag=split(StrTag,",")
	TagArrayCount=Ubound(Tag)

For DTRow = 1 To DataTable.GlobalSheet.GetRowCount Step 1
	DataTable.GlobalSheet.SetCurrentRow(DTRow)
	TagAction=Ucase(DataTable("Tag",dtglobalsheet))
	TestCase=trim((DataTable("TestCase",dtglobalsheet)))
	TestName=trim((DataTable("TestName",dtglobalsheet)))
	'DataTable.GlobalSheet.SetCurrentRow(DTRow)
	For Index=0 to TagArrayCount Step 1
		If instr(1,TagAction,Ucase(trim(Tag(Index)))) <> 0 or trim(Tag(Index))="run" Then
		DataTable("Status",dtglobalsheet) = "Run"
			
			'	Define Parameter for Create Folder TestCase for Get Snapshot per Test Case
			strResultPathTestCase=strResultPathExc
			Call CreateFolder(strResultPathTestCase) ' Create Folder befroe run
			Call WriteLog("Row Execution " &Datatable.GlobalSheet.GetCurrentRow &vbnewline&"StartTime : "&now()&vbnewline&"TestCase : " &TestCase&vbnewline&"Test Name : " &TestName&vbnewline&"TestCase Result (SnapShot) : " &strResultPathTestCase,"")
			'	Call Function from DataTable Column TestCase
			
		
			 Select case TestCase
                                                   
					                                     				
					Case "TC01" 'TC01	FX Deal Delivery Option
						
						RunAction "Action1 [TestFlight]", oneIteration
						
					Case "TC02" 'TC02  Register

						RunAction "Action1 [TestFlight02]", oneIteration
					

			
					
             End select
		
			
			'	Call Function Stamp Result from Report to Datatable Column 'Desctiption' 			
			FW_Create_DT_Result

			Exit for
			Elseif Index=TagArrayCount then
			DataTable("Status",dtglobalsheet) =  "No Run" 'Create Status not run 
			Exit for
		Else
			DataTable("Status",dtglobalsheet) = TagAction &" <> "&Tag(Index) &  " No Run"
		End If
	Next
Next

'Call Fnc Export Result from Datatable
	FW_StopExecute
	FW_DT_Result_Export
	
	'FUNC_Close_All_Browser_IE
Call WriteLog("Execution Complete ["&now&"]","")
'ExportReportToDoc
FW_StopExecute()

'********************************************************************************


