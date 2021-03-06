Option Explicit

Public sub app_Init(strWpfWindow)
   	On error resume next 
	obj_WpfWindow=strWpfWindow
	'obj_Page=strPage
'WpfWindow("Micro Focus MyFlight Sample").WpfEdit("agentName").Set
   If WpfWindow(obj_WpfWindow).Exist(0) Then    				   
		If err.number <>0 Then
			PrintInfo "micFail", "WpfWindow", " Error Description: " & Err.Description			  
			Err.clear						
			Exit sub
		End If
   	Else
	   PrintInfo "micFail", "WpfWindow", " Browser not found: " & strBrowser & " - " & strPage 	    
    End If             	
End sub

'Public sub app_frame(strInValue)
'   	On error resume next 
'	Set obj_Frame = Description.Create  	
'	obj_Frame("name").value= strInValue
'    	
'    If Browser(obj_Browser).Page(obj_Page).Frame(obj_Frame).Exist(0) Then    			   
'		If err.number <>0 Then
'			PrintInfo "micFail", "Frame", " Error Description: " & Err.Description			  
'			Err.clear						
'			Exit sub
'		End If
'   	Else
'	   PrintInfo "micFail", "Frame", " Frame not found: " 	    
'    End If        	                                                                       
'End sub


Public Sub MRWebEdit(actionName,obj_WpfWindow,strWebEdit,strInValue)


	
	On error resume next 
	Call StartTime(StrStartTime)
	If strInValue <> "" Then
	'WpfWindow("Micro Focus MyFlight Sample").WpfEdit("agentName").Set
	
		With JavaWindow(obj_WpfWindow)			
   			If .JavaEdit(strWebEdit).Exist(60) Then
   			  ' .JavaEdit(strWebEdit).highlight  			   
   			   .JavaEdit(strWebEdit).Set strInValue  
				Call EndTime(StrEndTime)		   
			   '	If err.number = 0 Then
        			PrintInfo "micPass", strWebEdit,  " Error Description: " & Err.Description
        			StrStatus = strWebEdit & "Step is Pass"
        			Call CaptureScreenShot(strFileName)
        			Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation)
        			Err.clear						
					Exit sub
				'End If
   			Else
   			    PrintInfo "micFail", strWebEdit, strWebEdit & "WebEdit not found."   
   			    StrStatus = strWebEdit & "Step is False"
   			    Call EndTime(StrEndTime)
   			    Call CaptureScreenShot(strFileName)
				Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation)
				'Call Killservicejava
				Call Rerun(actionName,obj_WpfWindow)
         	End If         	
		End With			
	End If
End Sub 


'JavaWindow("MX.3 - MX_PERF_TEST(PROD)_MX").JavaEdit("Ctrl-F").Set searchManu

Public Sub MRSearchEdit(actionName,obj_WpfWindow,strWebEdit,strInValue)

	On error resume next 
	Call StartTime(StrStartTime)
	If strInValue <> "" Then
	'WpfWindow("Micro Focus MyFlight Sample").WpfEdit("agentName").Set
	
		With JavaWindow(obj_WpfWindow)			
   			If .JavaEdit(strWebEdit).Exist(60) Then
   			   .JavaEdit(strWebEdit).highlight  			   
   			   .JavaEdit(strWebEdit).Set strInValue
				Call Sendkey("{ENTER}")			   
  				Call EndTime(StrEndTime)
			   '	If err.number = 0 Then
        			PrintInfo "micPass", strWebEdit,  " Error Description: " & Err.Description
        			StrStatus = strWebEdit & "Step is Pass"
        			Call CaptureScreenShot(strFileName)
        			Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation)
        			Err.clear						
					Exit sub
				'End If
   			Else
   			    PrintInfo "micFail", strWebEdit, strWebEdit & "WebEdit not found."   
   			    StrStatus = strWebEdit & "Step is False"
   			    Call EndTime(StrEndTime)
   			    Call CaptureScreenShot(strFileName)
				Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation) 	
				'Call Killservicejava	
				Call Rerun(actionName,obj_WpfWindow)
         	End If         	
		End With			
	End If
End Sub 
'--------------------------------------------------------------------------------------------


'JavaWindow("MX.3 - MX_PERF_TEST(PROD)_MX").JavaInternalFrame("MModalInternalFrame_2").JavaEdit("Opening a view or layout").Set dateSearch
Public Sub MRFramSetEdit(actionName,obj_WpfWindow,strInternalFrame,strWebEdit,strInValue)

	On error resume next 
	Call StartTime(StrStartTime)
	If strInValue <> "" Then
	'WpfWindow("Micro Focus MyFlight Sample").WpfEdit("agentName").Set
	
		With JavaWindow(obj_WpfWindow)			
   			If .JavaInternalFrame(strInternalFrame).Exist(60) Then
   			   .JavaInternalFrame(strInternalFrame).JavaEdit(strWebEdit).highlight  			   
   			   .JavaInternalFrame(strInternalFrame).JavaEdit(strWebEdit).Set strInValue
				Call Sendkey("{ENTER}")			   
  				Call EndTime(StrEndTime)
			   '	If err.number = 0 Then
        			PrintInfo "micPass", strWebEdit,  " Error Description: " & Err.Description
        			StrStatus = strWebEdit & "Step is Pass"
        			Call CaptureScreenShot(strFileName)
        			Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation)
        			Err.clear						
					Exit sub
				'End If
   			Else
   			    PrintInfo "micFail", strWebEdit, strWebEdit & "WebEdit not found."   
   			    StrStatus = strWebEdit & "Step is False"
   			    Call EndTime(StrEndTime)
   			    Call CaptureScreenShot(strFileName)
				Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation) 	
				'Call Killservicejava	
				Call Rerun(actionName,obj_WpfWindow)

         	End If         	
		End With			
	End If
End Sub 


Public Sub DPWebList (strWebList, strInValue)    
	On error resume next 
	If strInValue <> "" Then
		With Browser(obj_Browser).Page(obj_Page)						
			If .WebList(strWebList).Exist(0) Then
				.WebList(strWebList).Select strInValue
				If err.number <>0 Then
        			PrintInfo "micFail", strInValue,  " Error Description: " & Err.Description
        			Err.clear						
					Exit sub
				End If
			Else
   			   PrintInfo "micFail", strWebList,  strWebList & " - WebList not found."
         	End If			
		End With		
	End If
End Sub

Public Sub DPWebList2 (strWebList, strInValue)    
	On error resume next 
	If strInValue <> "" Then
		With Browser(obj_Browser).Page(obj_Page).Frame(obj_Frame)						
			If .WebList("name:=" & strWebList).Exist(0) Then
				.WebList("name:=" & strWebList).Select strInValue
				If err.number <>0 Then
        			PrintInfo "micFail", strInValue,  " Error Description: " & Err.Description
        			Err.clear						
					Exit sub
				End If
			Else
   			   PrintInfo "micFail", strWebList,  strWebList & " - WebList not found."
         	End If			
		End With		
	End If
End Sub

Public Sub DPWebElement(strType,strInValue)
	On error resume next 	
		With Browser(obj_Browser).Page(obj_Page).Frame(obj_Frame)
   			If .WebElement(strType & ":=" & strInValue).Exist(0) Then
   			   .WebElement(strType & ":=" & strInValue).Click
   				If err.number <>0 Then
        			PrintInfo "micFail", strInValue,  " Error Description: " & Err.Description
        			Err.clear						
					Exit sub
				End If
   			Else
   			  	PrintInfo "micFail", strInValue , strInValue & " - WebElement not found."
         	End If
		End With
End Sub


Public Function GetWebElement2(strObjp,strInValue,strType)
	On error resume next 	
		With Browser(obj_Browser).Page(obj_Page).Frame(obj_Frame)
   			If .WebElement(strObjp & ":=" & strInValue).Exist(0) Then
   			   GetWebElement = .WebElement(strObjp & ":=" & strInValue).GetROProperty(strType)
   				If err.number <>0 Then
        			PrintInfo "micFail", strInValue,  " Error Description: " & Err.Description
        			Err.clear						
					Exit Function
				End If
   			Else
   			  	PrintInfo "micFail", strInValue , strInValue & " - WebElement not found."
         	End If         	
		End With
End Function

'--------------------------------------------------------------------------------------------

Public Sub DPWebLink_Click(obj_WpfWindow,strWebLink)
	On error resume next 
	If strWebLink <> "" Then
		With WpfWindow(obj_WpfWindow)			
   			If .WpfButton(strWebLink).Exist(0) Then
   			   .WpfButton(strWebLink).click   			
			   	If err.number <>0 Then
        			PrintInfo "micFail", "WebLink Click", strWebLink & " Error Description: " & Err.Description
        			StrStatus = strWebEdit & "Step is Pass"
        			Call CaptureScreenShot(strFileName)
        			Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation)
        			Err.clear						
					Exit sub
				End If
   			Else
   			   PrintInfo "micFail", "WebLink Click", strWebLink & " WebLink not found."   			   
         	End If         	
		End With			
	End If
End Sub

Public Sub DPWebLink_MultiDes_Click(strWebLink01,strWebLink02)
	On error resume next 
	If strWebLink01 <> "" and strWebLink02 <> "" Then
		With Browser(obj_Browser).Page(obj_Page)			
   			If .Link(strWebLink01,strWebLink02).Exist(2) Then
   			   .Link(strWebLink01,strWebLink02).click   			
			   	If err.number <>0 Then
        			PrintInfo "micFail", "WebLink Click", strWebLink01 & strWebLink02 & " Error Description: " & Err.Description
        			Err.clear						
					Exit sub
				End If
   			Else
   			   PrintInfo "micFail", "WebLink Click", strWebLink01 & strWebLink02 & " WebLink not found."   			   
         	End If         	
		End With			
	End If
End Sub

'--------------------------------------------------------------------------------------------
'	WebTable

Public Sub DPWebTable_GetCellData(strWebTable,intRow, intCol)
	On error resume next 	           	                    
	With Browser(obj_Browser).Page(obj_Page)                                                                 
		If .WebTable(strWebTable).Exist(0) Then
			strCellData = trim(.WebTable(strWebTable).GetCellData (intRow, intCol))
			If err.number <>0 Then
				PrintInfo "micFail", strWebTable,  " Error Description: " & Err.Description
				Err.clear                                                                                               
				Exit sub
			End If
		Else
	   		PrintInfo "micFail", strWebTable, strWebTable & " WebTable not found: " 	                                      
		End If                    
	End With                              
End Sub
'--------------------------------------------------------------------------------------------


Public Sub MRAppButton_Click(actionName,obj_WpfWindow,strWebEdit)
		'MRAppButton_Click(actionName,obj_WpfWindow,strWebEdit)
	On error resume next 
	

	Call StartTime(StrStartTime)
	'WpfWindow("Micro Focus MyFlight Sample").WpfButton("OK").Click

	If strWebEdit <> "" Then
		With JavaWindow(obj_WpfWindow)			
   			If .JavaButton(strWebEdit).Exist(60) Then
   			   .JavaButton(strWebEdit).click   		
        			Call EndTime(StrEndTime)
			   '	If err.number = 0 Then
        			PrintInfo "micPass", strWebEdit,  " Error Description: " & Err.Description
        			StrStatus = strWebEdit & "Step is Pass"
        			Call CaptureScreenShot(strFileName)
        			Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation)
        			Err.clear						
					Exit sub
				'End If
   			Else
   			    PrintInfo "micFail", "WebButton Click", strWebEdit & " WebButton not found." 
			    Call EndTime(StrEndTime)
   			    Call CaptureScreenShot(strFileName)
				Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation)    	
				'Call Killservicejava
				call Rerun(actionName,obj_WpfWindow)
         	End If         	
		End With			
	End If
End Sub


Public Sub MRAppText_Click(actionName,obj_WpfWindow,strWebEdit)  
	On error resume next 
	Call StartTime(StrStartTime)
	'WpfWindow("Micro Focus MyFlight Sample").WpfButton("OK").Click

	If strWebEdit <> "" Then
		With JavaWindow(obj_WpfWindow)
 

   			If .JavaStaticText(strWebEdit).Exist(60) Then
   					If strWebEdit = "FO_TH-C" Then
   					
						JavaWindow("MX.3 - MX_PERF_TEST(PROD)").JavaStaticText(strWebEdit).Click 35,36,"LEFT"
	
					ElseIf strWebEdit = "FO_MNCS"	Then
	
						JavaWindow("MX.3 - MX_PERF_TEST(PROD)").JavaStaticText(strWebEdit).Click 32,33,"LEFT"
						
					ElseIf strWebEdit = "BO_OPS"	Then
					
					JavaWindow("MX.3 - MX_PERF_TEST(PROD)").JavaStaticText("BO_OPS(st)").Click 35,34,"LEFT"
					JavaWindow("MX.3 - MX_PERF_TEST(PROD)").JavaStaticText("BO_OPS").Click 35,34,"LEFT"
					ElseIf strWebEdit = "FO_BLMD" Then
					
						JavaWindow("MX.3 - MX_PERF_TEST(PROD)").JavaStaticText("FO_BLMD(st)").Click 54,35,"LEFT"
						JavaWindow("MX.3 - MX_PERF_TEST(PROD)").JavaStaticText("FO_BLMD").Click 54,35,"LEFT"
						
					ElseIF strWebEdit = "BO_STL" Then
					
							JavaWindow("MX.3 - MX_PERF_TEST(PROD)").JavaStaticText("BO_STL").Click 48,32,"LEFT"

					ElseIF strWebEdit = "MLC_CONFIG" Then
										
							'JavaWindow("MX.3 - MX_PERF_TEST(PROD)").JavaStaticText("MLC_CONFIG(st)")c
							JavaWindow("MX.3 - MX_PERF_TEST(PROD)").JavaStaticText("MLC_CONFIG").Click 48,32,"LEFT"

					
					ElseIf strWebEdit = "MX3" Then

						JavaWindow("MX.3 - MX_PERF_TEST(PROD)").JavaStaticText("MX.3").Click	
										
					ElseIf strWebEdit = "ALL(st)" Then
						
						JavaInternalFrame("Undefined - Main").JavaButton("pickUpValues").Click
						
					ElseIF strWebEdit = "FO_TH-C" Then	
					
						'JavaWindow("MX.3 - MX_PERF_TEST(PROD)").JavaStaticText("FO_TH-C").Click
						JavaWindow("MX.3 - MX_PERF_TEST(PROD)").JavaStaticText("FO_TH-C").Click 38,8,"LEFT"
						
					ElseIF strWebEdit = "BOCTRL_MK" Then		

						'JavaWindow("MX.3 - MX_PERF_TEST(PROD)").JavaStaticText("BOCTRL_MK").Click
						JavaWindow("MX.3 - MX_PERF_TEST(PROD)").JavaStaticText("BOCTRL_MK").Click 24,29,"LEFT"
					
					ElseIf strWebEdit = "MO" Then
		
						JavaWindow("MX.3 - MX_PERF_TEST(PROD)").JavaStaticText("MO(st)").Click 45,38,"LEFT"
						
					ElseIf strWebEdit = "FO_TRDD_FX" Then
					
							JavaWindow("MX.3 - MX_PERF_TEST(PROD)").JavaStaticText("FO_TRDD_FX(st)").Click 30,32,"LEFT"


							'JavaWindow("MX.3 - MX_PERF_TEST(PROD)").JavaStaticText("MO(st)").Click 45,38,"LEFT"
					ElseIf strWebEdit = "FO_TRDD_BD" Then
					
							JavaWindow("MX.3 - MX_PERF_TEST(PROD)").JavaStaticText("FO_TRDD_BD(st)").Click 66,20,"LEFT"		
					ElseIf strWebEdit = "FO_FINS" Then
							JavaWindow("MX.3 - MX_PERF_TEST(PROD)").JavaStaticText("FO_FINS(st)").Click 66,20,"LEFT"	


				End If
   			   '.JavaStaticText(strWebEdit).click   		
        			Call EndTime(StrEndTime)
			   '	If err.number = 0 Then
        			PrintInfo "micPass", strWebEdit,  " Error Description: " & Err.Description
        			StrStatus = strWebEdit & "Step is Pass"
        			Call CaptureScreenShot(strFileName)
        			Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation)
        			Err.clear						
					Exit sub
				'End If
   			Else
   			    PrintInfo "micFail", "WebButton Click", strWebEdit & " WebButton not found." 
			    Call EndTime(StrEndTime)
   			    Call CaptureScreenShot(strFileName)
				Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation)   
				'Call Killservicejava	
				call Rerun(actionName,obj_WpfWindow)
         	End If         	
		End With			
	End If
End Sub

'	Web Button Click Define 2 Object Description
Public Sub DPWebButton_MultiDes_Click(strWebButton01,strWebButton02)
	On error resume next 
	If strWebButton01 <> "" and strWebButton02 <> "" Then
		With Browser(obj_Browser).Page(obj_Page)			
   			If .WebButton(strWebButton01,strWebButton02).Exist(0) Then
   			   .WebButton(strWebButton01,strWebButton02).click   			
			   	If err.number <>0 Then
        			PrintInfo "micFail", "WebButton Click", strWebButton01 & strWebButton02& " Error Description: " & Err.Description
        			Err.clear						
					Exit sub
				End If
   			Else
   			   PrintInfo "micFail", "WebButton Click", strWebButton01 & strWebButton02 & " WebButton not found."   			   
         	End If         	
		End With			
	End If
End Sub

'--------------------------------------------------------------------------------------------
'	Get Ro Property Description
Public Function DP_Obj_GetROProperty(strObject,strObjectValue,strProperty)
	On error resume next 
	If strObject <> "" and strProperty <> ""  Then

		If Lcase(strObject) = "link" Then
			With Browser(obj_Browser).Page(obj_Page)			
	   			If .Link(strObjectValue).Exist(0) Then
				   	If err.number <>0 Then
	        			PrintInfo "micFail", strObject, strObjectValue & " Error Description: " & Err.Description
	        			Err.clear						
						Exit Function
					End If
					str_Obj_value=.Link(strObjectValue).GetROProperty(strProperty)  
					DP_Obj_GetROProperty=str_Obj_value
					
	   			Else
	   			   PrintInfo "micFail", strObject, strObjectValue & " WebLink cannot GetProperty " &strProperty   			   
	         	End If         	
			End With		
			
		ELseIf Lcase(strObject) = "webelement" Then'or Lcase(strObject) = "element" Then

			With Browser(obj_Browser).Page(obj_Page)			
	   			If .WebElement(strObjectValue).Exist(0) Then
				   	If err.number <>0 Then
	        			PrintInfo "micFail", strObject, strObjectValue & " Error Description: " & Err.Description
	        			Err.clear						
						Exit Function
					End If
					str_Obj_value=.WebElement(strObjectValue).GetROProperty(strProperty)  
					DP_Obj_GetROProperty=str_Obj_value
					
	   			Else
	   			   PrintInfo "micFail", strObject, strObjectValue & " WebElement cannot GetProperty " &strProperty   			   
	         	End If         	
			End With		
			
		ELseIf Lcase(strObject) = "webbutton" or Lcase(strObject) = "button" Then

			With Browser(obj_Browser).Page(obj_Page)			
	   			If .WebButton(strObjectValue).Exist(0) Then
				   	If err.number <>0 Then
	        			PrintInfo "micFail", strObject, strObjectValue & " Error Description: " & Err.Description
	        			Err.clear						
						Exit Function
					End If
					str_Obj_value=.WebButton(strObjectValue).GetROProperty(strProperty)  
					DP_Obj_GetROProperty=str_Obj_value
					
	   			Else
	   			   PrintInfo "micFail", strObject, strObjectValue & " WebButton cannot GetProperty " &strProperty   			   
	         	End If         	
			End With	
			
		ELseIf Lcase(strObject) = "webedit" or Lcase(strObject) = "edit" Then

			With Browser(obj_Browser).Page(obj_Page)			
	   			If .WebEdit(strObjectValue).Exist(0) Then
				   	If err.number <>0 Then
	        			PrintInfo "micFail", strObject, strObjectValue & " Error Description: " & Err.Description
	        			Err.clear						
						Exit Function
					End If
					str_Obj_value=.WebEdit(strObjectValue).GetROProperty(strProperty)  
					DP_Obj_GetROProperty=str_Obj_value
					
	   			Else
	   			   PrintInfo "micFail", strObject, strObjectValue & " WebEdit cannot GetProperty " &strProperty   			   
	         	End If         	
			End With	
		ELseIf Lcase(strObject) = "webpage" or Lcase(strObject) = "page" Then
	   			If Browser(obj_Browser).Page(obj_Page).Exist(0) Then	    			
				   	If err.number <>0 Then
	        			PrintInfo "micFail", strObject, strObjectValue & " Error Description: " & Err.Description
	        			Err.clear						
						Exit Function
					End If
					str_Obj_value=Browser(obj_Browser).Page(obj_Page).GetROProperty(strProperty)  
					DP_Obj_GetROProperty=str_Obj_value
					
	   			Else
	   			   PrintInfo "micFail", strObject, strObjectValue & " WebPage cannot GetProperty " &strProperty   			   
	         	End If  

		ELseIf Lcase(strObject) = "image" Then' or Lcase(strObject) = "img" Then
			With Browser(obj_Browser).Page(obj_Page)			
	   			If .Image(strObjectValue).Exist(0) Then
				   	If err.number <>0 Then
	        			PrintInfo "micFail", strObject, strObjectValue & " Error Description: " & Err.Description
	        			Err.clear						
						Exit Function
					End If
					str_Obj_value=.Image(strObjectValue).GetROProperty(strProperty)  
					DP_Obj_GetROProperty=str_Obj_value
					
	   			Else
	   			   PrintInfo "micFail", strObject, strObjectValue & " Image cannot GetProperty " &strProperty   			   
	         	End If         	
			End With	

		End If

	End If
End Function


'	Get Ro Property Multi Description
Public Function DP_Obj_MultiDes_GetROProperty(strObject,strObjectValue01,strObjectValue02,strProperty)
	On error resume next 
	If strObject <> "" and strObjectValue01 <> "" and strObjectValue02 <> "" Then

		If Lcase(strObject) = "weblink" or Lcase(strObject) = "link" Then

			With Browser(obj_Browser).Page(obj_Page)			
	   			If .Link(strObjectValue01,strObjectValue02).Exist(0) Then
				   	If err.number <>0 Then
	        			PrintInfo "micFail", strObject, strObjectValue01 & strObjectValue02 & " Error Description: " & Err.Description
	        			Err.clear						
						Exit Function
					End If
					str_Obj_value=.Link(strObjectValue01,strObjectValue02).GetROProperty(strProperty)  
					DP_Obj_MultiDes_GetROProperty=str_Obj_value
					
	   			Else
	   			   PrintInfo "micFail", strObject, strObjectValue01 & strObjectValue02 & " WebLink cannot GetProperty " &strProperty   			   
	         	End If         	
			End With		
			
		ELseIf Lcase(strObject) = "webelement" or Lcase(strObject) = "element" Then

			With Browser(obj_Browser).Page(obj_Page)			
	   			If .WebElement(strObjectValue01,strObjectValue02).Exist(0) Then
				   	If err.number <>0 Then
	        			PrintInfo "micFail", strObject, strObjectValue01 & strObjectValue02 & " Error Description: " & Err.Description
	        			Err.clear						
						Exit Function
					End If
					str_Obj_value=.WebElement(strObjectValue01,strObjectValue02).GetROProperty(strProperty)  
					DP_Obj_MultiDes_GetROProperty=str_Obj_value
					
	   			Else
	   			   PrintInfo "micFail", strObject, strObjectValue01 & strObjectValue02 & " WebElement cannot GetProperty " &strProperty   			   
	         	End If         	
			End With		
			
		ELseIf Lcase(strObject) = "webbutton" or Lcase(strObject) = "button" Then

			With Browser(obj_Browser).Page(obj_Page)			
	   			If .WebButton(strObjectValue01,strObjectValue02).Exist(0) Then
				   	If err.number <>0 Then
	        			PrintInfo "micFail", strObject, strObjectValue01 & strObjectValue02 & " Error Description: " & Err.Description
	        			Err.clear						
						Exit Function
					End If
					str_Obj_value=.WebButton(strObjectValue01,strObjectValue02).GetROProperty(strProperty)  
					DP_Obj_MultiDes_GetROProperty=str_Obj_value
					
	   			Else
	   			   PrintInfo "micFail", strObject, strObjectValue01 & strObjectValue02 & " WebButton cannot GetProperty " &strProperty   			   
	         	End If         	
			End With	
			
		ELseIf Lcase(strObject) = "webedit" or Lcase(strObject) = "edit" Then

			With Browser(obj_Browser).Page(obj_Page)			
	   			If .WebEdit(strObjectValue01,strObjectValue02).Exist(0) Then
				   	If err.number <>0 Then
	        			PrintInfo "micFail", strObject, strObjectValue01 & strObjectValue02 & " Error Description: " & Err.Description
	        			Err.clear						
						Exit Function
					End If
					str_Obj_value=.WebEdit(strObjectValue01,strObjectValue02).GetROProperty(strProperty)  
					DP_Obj_MultiDes_GetROProperty=str_Obj_value
					
	   			Else
	   			   PrintInfo "micFail", strObject, strObjectValue01 & strObjectValue02 & " WebEdit cannot GetProperty " &strProperty   			   
	         	End If         	
			End With	
			
		ELseIf Lcase(strObject) = "image" or Lcase(strObject) = "img" Then

			With Browser(obj_Browser).Page(obj_Page)			
	   			If .Image(strObjectValue01,strObjectValue02).Exist(0) Then
				   	If err.number <>0 Then
	        			PrintInfo "micFail", strObject, strObjectValue01 & strObjectValue02 & " Error Description: " & Err.Description
	        			Err.clear						
						Exit Function
					End If
					str_Obj_value=.Image(strObjectValue01,strObjectValue02).GetROProperty(strProperty)  
					DP_Obj_MultiDes_GetROProperty=str_Obj_value
					
	   			Else
	   			   PrintInfo "micFail", strObject, strObjectValue01 & strObjectValue02 & " Image cannot GetProperty " &strProperty   			   
	         	End If         	
			End With	
			
		End If

	End If
End Function


'	Check Object and click
Public Sub DP_Obj_Exist_Click(strObject,strObjectValue,strWaitTime)
	On error resume next 
	If strWaitTime <> "" Then
		strWaitTime=Wait_Short
	End If
	
	If strObject <> "" and strObjectValue <> "" Then
	
		If Lcase(strObject) = "weblink" Then
			With Browser(obj_Browser).Page(obj_Page)			
	   			If .WebLink(strObjectValue).Exist(strWaitTime) Then 			
					.WebLink(strObjectValue).click 
	         	End If         	
			End With
		Elseif Lcase(strObject) = "link" Then
			With Browser(obj_Browser).Page(obj_Page)			
	   			If .Link(strObjectValue).Exist(strWaitTime) Then 			
					.Link(strObjectValue).click 
	         	End If         	
			End With	
		Elseif Lcase(strObject) = "image" Then
			With Browser(obj_Browser).Page(obj_Page)			
	   			If .Image(strObjectValue).Exist(strWaitTime) Then 			
					.Image(strObjectValue).click 
	         	End If         	
			End With
		Elseif Lcase(strObject) = "link" Then
			With Browser(obj_Browser).Page(obj_Page)			
	   			If .Link(strObjectValue).Exist(strWaitTime) Then 			
					.Link(strObjectValue).click 
	         	End If         	
			End With
		Elseif Lcase(strObject) = "webelement" or Lcase(strObject) = "element" Then
			With Browser(obj_Browser).Page(obj_Page)			
	   			If .WebElement(strObjectValue).Exist(strWaitTime) Then 			
					.WebElement(strObjectValue).click 
	         	End If         	
			End With
		Elseif Lcase(strObject) = "webbutton" or Lcase(strObject) = "button" Then
			With Browser(obj_Browser).Page(obj_Page)			
	   			If .WebButton(strObjectValue).Exist(strWaitTime) Then 			
					.WebButton(strObjectValue).click 
	         	End If         	
			End With
		Elseif Lcase(strObject) = "webcheckbox" or Lcase(strObject) = "checkbox" Then
			With Browser(obj_Browser).Page(obj_Page)			
	   			If .WebCheckBox(strObjectValue).Exist(strWaitTime) Then 			
					.WebCheckBox(strObjectValue).click 
	         	End If         	
			End With
		End If
	Else
			Print "Fail to call Function "&strObject&" "&strObjectValue
	 		Exit sub
	End If
	
	
End Sub


'--------------------------------------------------------------------------------------------
'	Check Object Exist
Public Function DP_Obj_Exist(strObject,strObjectValue,strWaitTime)
	On error resume next 
	If strWaitTime = "" Then
		strWaitTime=Wait_Short
	End If
	
	If strObject <> "" and strObjectValue <> "" Then
		
		'	Set Default Value return Object status
		DP_Obj_Exist="Exist"
		If Lcase(strObject) = "webedit" or Lcase(strObject) = "edit" Then
			
			With Browser(obj_Browser).Page(obj_Page)			
	   			If Not .WebEdit(strObjectValue).Exist(strWaitTime) Then 			
		   			If err.number <>0 Then
	        			PrintInfo "micFail", strObject&" "&strObjectValue,  "WebEdit Error Description: " & Err.Description
	        			Call Fnc_ExitIteration(strObject&" "&strObjectValue,  "WebEdit Error Description: " & Err.Description,"")
	        			DP_Obj_Exist="Not Exist"
	        			Err.clear						
						Exit Function
					End If 	
	         	End If         	
			End With

		ElseIf Lcase(strObject) = "weblink" or Lcase(strObject) = "link" Then
			With Browser(obj_Browser).Page(obj_Page)			
	   			If Not .Link(strObjectValue).Exist(strWaitTime) Then 			
		   			If err.number <>0 Then
	        			PrintInfo "micFail", strObject&" "&strObjectValue,  "WebLink Error Description: " & Err.Description
	        			Call Fnc_ExitIteration(strObject&" "&strObjectValue,  "WebLink Error Description: " & Err.Description,"")
	        			DP_Obj_Exist="Fail"
	        			Err.clear						
						Exit Function
					End If 	
	         	End If         	
			End With
			
		Elseif Lcase(strObject) = "weblist" or Lcase(strObject) = "list" Then
			With Browser(obj_Browser).Page(obj_Page)			
	   			If Not .WebList(strObjectValue).Exist(strWaitTime) Then 			
		   			If err.number <>0 Then
	        			PrintInfo "micFail", strObject&" "&strObjectValue,  "WebList Error Description: " & Err.Description
	        			Call Fnc_ExitIteration(strObject&" "&strObjectValue,  "WebList Error Description: " & Err.Description,"")
	        			Err.clear						
						Exit Function
					End If 	
	         	End If         	
			End With
			
		Elseif Lcase(strObject) = "webelement" or Lcase(strObject) = "element" Then
			With Browser(obj_Browser).Page(obj_Page)			
	   			If Not .WebElement(strObjectValue).Exist(strWaitTime) Then 			
		   			If err.number <>0 Then
	        			PrintInfo "micFail", strObject&" "&strObjectValue,  "WebElement Error Description: " & Err.Description
	        			Call Fnc_ExitIteration(strObject&" "&strObjectValue,  "WebElement Error Description: " & Err.Description,"")
	        			DP_Obj_Exist="Fail"
	        			Err.clear						
						Exit Function
					End If 	
	         	End If         	
			End With
			
		Elseif Lcase(strObject) = "webbutton" or Lcase(strObject) = "button" Then
			With Browser(obj_Browser).Page(obj_Page)			
	   			If Not .WebButton(strObjectValue).Exist(strWaitTime) Then 			
		   			If err.number <>0 Then
	        			PrintInfo "micFail", strObject&" "&strObjectValue,  "WebButton Error Description: " & Err.Description
	        			Call Fnc_ExitIteration(strObject&" "&strObjectValue,  "WebButton Error Description: " & Err.Description,"")
	        			Err.clear						
						Exit Function
					End If 	
	         	End If         	
			End With

		Elseif Lcase(strObject) = "webtable" or Lcase(strObject) = "table" Then
			With Browser(obj_Browser).Page(obj_Page)			
	   			If Not .WebTable(strObjectValue).Exist(strWaitTime) Then 			
		   			If err.number <>0 Then
	        			PrintInfo "micFail", strObject&" "&strObjectValue,  "WebTable Error Description: " & Err.Description
	        			Call Fnc_ExitIteration(strObject&" "&strObjectValue,  "WebTable Error Description: " & Err.Description,"")
	        			DP_Obj_Exist="Fail"
	        			Err.clear						
						Exit Function
					End If 	
	         	End If         	
			End With

		Elseif Lcase(strObject) = "image" or Lcase(strObject) = "img" Then
			With Browser(obj_Browser).Page(obj_Page)			
	   			If Not .Image(strObjectValue).Exist(strWaitTime) Then 			
		   			If err.number <>0 Then
	        			PrintInfo "micFail", strObject&" "&strObjectValue,  "Image Error Description: " & Err.Description
	        			Call Fnc_ExitIteration(strObject&" "&strObjectValue,  "Image Error Description: " & Err.Description,"")
	        			DP_Obj_Exist="Fail"
	        			Err.clear						
						Exit Function
					End If 	
	         	End If         	
			End With
			
		Else
			Print "Fail to call Function "&strObject&" "&strObjectValue
			DP_Obj_Exist="Fail"
	 		Exit Function
	 	End If
	End If
	

	
End Function


'	Check Object Multi Description Exist
Public Function DP_Obj_MultiDes_Exist(strObject,strObjectValue01,strObjectValue02,strWaitTime)
	On error resume next 
	If strWaitTime = "" Then
		strWaitTime=Wait_Short
	End If
	
	If strObject <> "" and strObjectValue01 <> "" and strObjectValue02 <> "" Then
		
		'	Set Default Value return Object status
		DP_Obj_Exist="Exist"
		If Lcase(strObject) = "webedit" or Lcase(strObject) = "edit" Then
			
			With Browser(obj_Browser).Page(obj_Page)			
	   			If Not .WebEdit(strObjectValue01,strObjectValue02).Exist(strWaitTime) Then 			
		   			If err.number <>0 Then
	        			PrintInfo "micFail", strObject&" "&strObjectValue01&strObjectValue02,  "WebEdit Error Description: " & Err.Description
	        			Call Fnc_ExitIteration(strObject&" "&strObjectValue01&strObjectValue02,  "WebEdit Error Description: " & Err.Description,"")
	        			DP_Obj_Exist="Not Exist"
	        			Err.clear						
						Exit Function
					End If 	
	         	End If         	
			End With

		ElseIf Lcase(strObject) = "weblink" or Lcase(strObject) = "link" Then
			With Browser(obj_Browser).Page(obj_Page)			
	   			If Not .Link(strObjectValue01,strObjectValue02).Exist(strWaitTime) Then 			
		   			If err.number <>0 Then
	        			PrintInfo "micFail", strObject&" "&strObjectValue01&strObjectValue02,  "WebLink Error Description: " & Err.Description
	        			Call Fnc_ExitIteration(strObject&" "&strObjectValue01&strObjectValue02,  "WebLink Error Description: " & Err.Description,"")
	        			DP_Obj_Exist="Fail"
	        			Err.clear						
						Exit Function
					End If 	
	         	End If         	
			End With
			
		Elseif Lcase(strObject) = "weblist" or Lcase(strObject) = "list" Then
			With Browser(obj_Browser).Page(obj_Page)			
	   			If Not .WebList(strObjectValue01,strObjectValue02).Exist(strWaitTime) Then 			
		   			If err.number <>0 Then
	        			PrintInfo "micFail", strObject&" "&strObjectValue01&strObjectValue02,  "WebList Error Description: " & Err.Description
	        			Call Fnc_ExitIteration(strObject&" "&strObjectValue01&strObjectValue02,  "WebList Error Description: " & Err.Description,"")
	        			Err.clear						
						Exit Function
					End If 	
	         	End If         	
			End With
			
		Elseif Lcase(strObject) = "webelement" or Lcase(strObject) = "element" Then
			With Browser(obj_Browser).Page(obj_Page)			
	   			If Not .WebElement(strObjectValue01,strObjectValue02).Exist(strWaitTime) Then 			
		   			If err.number <>0 Then
	        			PrintInfo "micFail", strObject&" "&strObjectValue01&strObjectValue02,  "WebElement Error Description: " & Err.Description
	        			Call Fnc_ExitIteration(strObject&" "&strObjectValue01&strObjectValue02,  "WebElement Error Description: " & Err.Description,"")
	        			DP_Obj_Exist="Fail"
	        			Err.clear						
						Exit Function
					End If 	
	         	End If         	
			End With
			
		Elseif Lcase(strObject) = "webbutton" or Lcase(strObject) = "button" Then
			With Browser(obj_Browser).Page(obj_Page)			
	   			If Not .WebButton(strObjectValue01,strObjectValue02).Exist(strWaitTime) Then 			
		   			If err.number <>0 Then
	        			PrintInfo "micFail", strObject&" "&strObjectValue01&strObjectValue02,  "WebButton Error Description: " & Err.Description
	        			Call Fnc_ExitIteration(strObject&" "&strObjectValue01&strObjectValue02,  "WebButton Error Description: " & Err.Description,"")
	        			Err.clear						
						Exit Function
					End If 	
	         	End If         	
			End With

		Elseif Lcase(strObject) = "webtable" or Lcase(strObject) = "table" Then
			With Browser(obj_Browser).Page(obj_Page)			
	   			If Not .WebTable(strObjectValue01,strObjectValue02).Exist(strWaitTime) Then 			
		   			If err.number <>0 Then
	        			PrintInfo "micFail", strObject&" "&strObjectValue01&strObjectValue02,  "WebTable Error Description: " & Err.Description
	        			Call Fnc_ExitIteration(strObject&" "&strObjectValue01&strObjectValue02,  "WebTable Error Description: " & Err.Description,"")
	        			DP_Obj_Exist="Fail"
	        			Err.clear						
						Exit Function
					End If 	
	         	End If         	
			End With

		Elseif Lcase(strObject) = "image" or Lcase(strObject) = "img" Then
			With Browser(obj_Browser).Page(obj_Page)			
	   			If Not .Image(strObjectValue01,strObjectValue02).Exist(strWaitTime) Then 			
		   			If err.number <>0 Then
	        			PrintInfo "micFail", strObject&" "&strObjectValue01&strObjectValue02,  "Image Error Description: " & Err.Description
	        			Call Fnc_ExitIteration(strObject&" "&strObjectValue01&strObjectValue02,  "Image Error Description: " & Err.Description,"")
	        			DP_Obj_Exist="Fail"
	        			Err.clear						
						Exit Function
					End If 	
	         	End If         	
			End With
			
		Else
			Print "Fail to call Function "&strObject&" "&strObjectValue
			DP_Obj_Exist="Fail"
	 		Exit Function
	 	End If
	End If
	

	
End Function


'--------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------
'	Check Object Sync
Public Function DP_Obj_Sync(strObject,strObjectValue,strWaitTime)
	On error resume next 
	If strWaitTime = "" Then
		strWaitTime=Wait_Short
	End If
	
	If strObject <> "" Then
		'	Set Default Value return Object status
		DP_Obj_Exist="Exist"
		If Lcase(strObject) = "webpage" or Lcase(strObject) = "page" Then
					
	   		If Browser(obj_Browser).Page(obj_Page).Exist(strWaitTime) Then 
				Browser(obj_Browser).Page(obj_Page).sync					
			Else
				If err.number <>0 Then
        			PrintInfo "micFail", strObject&" "&strObjectValue,  "Browser-Page Sync Error Description: " & Err.Description
        			Call Fnc_ExitIteration(strObject&" "&strObjectValue,  "Browser-Page Sync Error Description: " & Err.Description,"")
        			DP_Obj_Exist="Fail"
        			Err.clear						
					Exit Function
				End If 	
         	End If         	

		ElseIf Lcase(strObject) = "webedit" or Lcase(strObject) = "edit" Then
			
			With Browser(obj_Browser).Page(obj_Page)			
	   			If .WebEdit(strObjectValue).Exist(strWaitTime) Then 
					.WebEdit(strObjectValue).sync					
		   		Else
			   		If err.number <>0 Then
	        			PrintInfo "micFail", strObject&" "&strObjectValue,  "WebEdit Sync Error Description: " & Err.Description
	        			Call Fnc_ExitIteration(strObject&" "&strObjectValue,  "WebEdit Sync Error Description: " & Err.Description,"")
	        			DP_Obj_Exist="Fail"
	        			Err.clear						
						Exit Function
					End If 	
	         	End If         	
			End With

		ElseIf Lcase(strObject) = "weblink" or Lcase(strObject) = "link" Then
			With Browser(obj_Browser).Page(obj_Page)			
	   			If .Link(strObjectValue).Exist(strWaitTime) Then 	
					.Link(strObjectValue).sync	
				Else					
		   			If err.number <>0 Then
	        			PrintInfo "micFail", strObject&" "&strObjectValue,  "WebLink Sync Error Description: " & Err.Description
	        			Call Fnc_ExitIteration(strObject&" "&strObjectValue,  "WebLink Sync Error Description: " & Err.Description,"")
	        			DP_Obj_Exist="Fail"
	        			Err.clear						
						Exit Function
					End If 	
	         	End If         	
			End With
			
		Elseif Lcase(strObject) = "weblist" or Lcase(strObject) = "list" Then
			With Browser(obj_Browser).Page(obj_Page)			
	   			If .WebList(strObjectValue).Exist(strWaitTime) Then 	
					.WebList(strObjectValue).sync	   		
				Else					
		   			If err.number <>0 Then
	        			PrintInfo "micFail", strObject&" "&strObjectValue,  "WebList Sync Error Description: " & Err.Description
	        			Call Fnc_ExitIteration(strObject&" "&strObjectValue,  "WebList Sync Error Description: " & Err.Description,"")
	        			Err.clear						
						Exit Function
					End If 	
	         	End If         	
			End With
			
		Elseif Lcase(strObject) = "webelement" or Lcase(strObject) = "element" Then
			With Browser(obj_Browser).Page(obj_Page)			
	   			If .WebElement(strObjectValue).Exist(strWaitTime) Then 	
					.WebElement(strObjectValue).sync	   			
		   		Else
			   		If err.number <>0 Then
	        			PrintInfo "micFail", strObject&" "&strObjectValue,  "WebElement Sync Error Description: " & Err.Description
	        			Call Fnc_ExitIteration(strObject&" "&strObjectValue,  "WebElement Sync Error Description: " & Err.Description,"")
	        			DP_Obj_Exist="Fail"
	        			Err.clear						
						Exit Function
					End If 	
	         	End If         	
			End With
			
		Elseif Lcase(strObject) = "webbutton" or Lcase(strObject) = "button" Then
			With Browser(obj_Browser).Page(obj_Page)			
	   			If Not .WebButton(strObjectValue).Exist(strWaitTime) Then 	
					.WebButton(strObjectValue).sync	   			
		   			If err.number <>0 Then
	        			PrintInfo "micFail", strObject&" "&strObjectValue,  "WebButton Sync Error Description: " & Err.Description
	        			Call Fnc_ExitIteration(strObject&" "&strObjectValue,  "WebButton Sync Error Description: " & Err.Description,"")
	        			Err.clear						
						Exit Function
					End If 	
	         	End If         	
			End With

		Elseif Lcase(strObject) = "webtable" or Lcase(strObject) = "table" Then
			With Browser(obj_Browser).Page(obj_Page)			
	   			If .WebTable(strObjectValue).Exist(strWaitTime) Then 
					.WebTable(strObjectValue).sync	   			
		   		Else
			   		If err.number <>0 Then
	        			PrintInfo "micFail", strObject&" "&strObjectValue,  "WebTable Sync Error Description: " & Err.Description
	        			Call Fnc_ExitIteration(strObject&" "&strObjectValue,  "WebTable Sync Error Description: " & Err.Description,"")
	        			DP_Obj_Exist="Fail"
	        			Err.clear						
						Exit Function
					End If 	
	         	End If         	
			End With
			
		Else
			Print "Fail to call Function "&strObject&" "&strObjectValue
			DP_Obj_Exist="Fail"
	 		Exit Function
	 	End If
	End If
	
End Function

' Oject WebRadioGroup
Public Sub DPWebRadioGroup(strWebRadioGroup, strInValue)
	On error resume next 
	If strInValue <> "" Then
		With Browser(obj_Browser).Page(obj_Page)			
   			If .WebRadioGroup(strWebRadioGroup).Exist(0) Then
   			   .WebRadioGroup(strWebRadioGroup).Select strInValue   			
			   	If err.number <>0 Then
        			PrintInfo "micFail", strWebRadioGroup,  " Error Description: " & Err.Description
        			Err.clear						
					Exit sub
				End If
   			Else
   			   PrintInfo "micFail", strWebRadioGroup, strWebRadioGroup & "strWebRadioGroup not found."   			   
         	End If         	
		End With			
	End If
End Sub 



Public Sub MRAppSelect_Combo(actionName,obj_WpfWindow,strWebEdit,strInValue)
	On error resume next 
	Call StartTime(StrStartTime)
	If strInValue <> "" Then

		'WpfWindow("Micro Focus MyFlight Sample").WpfComboBox("fromCity").Select "Paris"
		With JavaWindow(obj_WpfWindow)			
   			If .JavaList(strWebEdit).Exist(0) Then
   			   .JavaList(strWebEdit).select strInValue   			
			   		Call EndTime(StrEndTime)
			   '	If err.number = 0 Then
        			PrintInfo "micPass", strWebEdit,  " Error Description: " & Err.Description
        			StrStatus = strWebEdit & "Step is Pass"
        			Call CaptureScreenShot(strFileName)
        			Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation)
        			Err.clear						
					Exit sub
   			Else
   			   	PrintInfo "micFail", strWebEdit, strWebEdit & "WebEdit not found."   
   			    StrStatus = strWebEdit & "Step is False"
   			    Call EndTime(StrEndTime)
   			    Call CaptureScreenShot(strFileName)
				Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation) 	
				call Rerun(actionName,obj_WpfWindow)
         	End If         	
		End With			
	End If
End Sub 


'Public Sub Getparamiter(strWebEdit,StrStatus,Myfile)
'
'	On error resume next 
'	iRowCount = Datatable.getSheet("TestStep").getRowCount
' 	Datatable.getSheet("TestStep").setCurrentRow(iRowCount+1)
' 	DataTable.Value("No","TestStep") = strWebEdit
' 	DataTable.Value("LogStep","TestStep") = StrStatus
' 	DataTable.Value("CapScreen","TestStep") = Myfile
' 	DataTable.Value("StartTime","TestStep") = StartTime
' 	DataTable.Value("EndTime","TestStep") = EndTime
'End Sub  


Public Sub StartTime(StrStartTime)

'StrStartTime = Time()
t = Timer
temp = Int(t)

Milliseconds = Int((t-temp) * 1000)

Seconds = temp mod 60
temp    = Int(temp/60)
Minutes = temp mod 60
Hours   = Int(temp/60)

'Print "Hours is "&" " & Hours &"| Minutes is "&" " & Minutes &"| Seconds is "&" " & Seconds &"| Milliseconds is "&" " & Milliseconds

StrStartTime = String(2 - Len(Hours), "0") & Hours & ":"
StrStartTime = StrStartTime & String(2 - Len(Minutes), "0") & Minutes & ":"
StrStartTime = StrStartTime & String(2 - Len(Seconds), "0") & Seconds & "."
StrStartTime = StrStartTime & String(4 - Len(Milliseconds), "0") & Milliseconds
Print StrStartTime
'MsgBox startTime
End Sub  

Public Sub EndTime(StrEndTime)

'StrEndTime = Time()
'MsgBox startTime

t = Timer
temp = Int(t)

Milliseconds = Int((t-temp) * 1000)

Seconds = temp mod 60
temp    = Int(temp/60)
Minutes = temp mod 60
Hours   = Int(temp/60)

'Print "Hours is "&" " & Hours &"| Minutes is "&" " & Minutes &"| Seconds is "&" " & Seconds &"| Milliseconds is "&" " & Milliseconds

StrEndTime = String(2 - Len(Hours), "0") & Hours & ":"
StrEndTime = StrEndTime & String(2 - Len(Minutes), "0") & Minutes & ":"
StrEndTime = StrEndTime & String(2 - Len(Seconds), "0") & Seconds & "."
StrEndTime = StrEndTime & String(4 - Len(Milliseconds), "0") & Milliseconds
'Print StrEndTime

End Sub  

Public Function Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation)

	On error resume next 
	iRowCount = Datatable.getSheet("TestStep").getRowCount
 	Datatable.getSheet("TestStep").setCurrentRow(iRowCount+1)
 	DataTable.Value("No","TestStep") = strWebEdit
 	DataTable.Value("LogStep","TestStep") = StrStatus
 	DataTable.Value("CapScreen","TestStep") = Myfile
 	DataTable.Value("StartTime","TestStep") = StrStartTime
 	DataTable.Value("EndTime","TestStep") = StrEndTime
 	DataTable.Value("Transation","TestStep") = Transation
 	
End Function  

Public Sub MRGettextOnedit(actionName,obj_WpfWindow,strWebEdit,strvalue)



	On error resume next 
	Call StartTime(StrStartTime)
	If obj_WpfWindow <> "" Then	
	'WpfWindow("Micro Focus MyFlight Sample").WpfObject("John Smith").GetROProperty("text")


		With WpfWindow(obj_WpfWindow)			
   			If .JavaEdit(strvalue).Exist(0) Then
   				strWebEdit = JavaWindow(obj_WpfWindow).JavaInternalFrame(strWebEdit).JavaEdit(strvalue).GetROProperty("text")
   			   
  				Call EndTime(StrEndTime)
			   '	If err.number = 0 Then
        			PrintInfo "micPass", strWebEdit,  " Error Description: " & Err.Description
        			StrStatus = strWebEdit & "Step is Pass"
        			Call CaptureScreenShot(strFileName)
        			Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation)
        			Err.clear						
					Exit sub
				'End If
   			Else
   			    PrintInfo "micFail", strWebEdit, strWebEdit & "WebEdit not found."   
   			    StrStatus = strWebEdit & "Step is False"
   			    Call EndTime(StrEndTime)
   			    Call CaptureScreenShot(strFileName)
				Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation) 
				call Rerun(actionName,obj_WpfWindow)

         	End If         	
		End With			
	End If
End Sub 



Public Sub MRWebEdit_Select(actionName,obj_WpfWindow,obj_frame,strWebEdit,strInValue)

'JavaWindow("MX.3 - MX_PERF_TEST(PROD)_MX").JavaInternalFrame("Undefined - Main").JavaEdit("Past/Future").Set "ALL" 

	On error resume next 
	Call StartTime(StrStartTime)
	If strInValue <> "" Then
	'WpfWindow("Micro Focus MyFlight Sample").WpfEdit("agentName").Set
	
		With JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame)		
   			If .JavaEdit(strWebEdit).Exist(60) Then
   			  ' .JavaEdit(strWebEdit).highlight
   			   .JavaEdit(strWebEdit).Set strInValue 
				Call Sendkey("{ENTER}")  
  				Call EndTime(StrEndTime)
			   '	If err.number = 0 Then
        			PrintInfo "micPass", strWebEdit,  " Error Description: " & Err.Description
        			StrStatus = strWebEdit & "Step is Pass"
        			Call CaptureScreenShot(strFileName)
        			Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation)
        			Err.clear						
					Exit sub
				'End If
   			Else
   			    PrintInfo "micFail", strWebEdit, strWebEdit & "WebEdit not found."   
   			    StrStatus = strWebEdit & "Step is False"
   			    Call EndTime(StrEndTime)
   			    Call CaptureScreenShot(strFileName)
				Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation)
				call Rerun(actionName,obj_WpfWindow)

         	End If         	
		End With			
	End If
End Sub 

'JavaWindow("MX.3 - MX_PERF_TEST(PROD)_MX").JavaInternalFrame("Undefined - Main").JavaButton("pickUpValues").Click
Public Sub MRAppButton_Click_frame(actionName,obj_WpfWindow,obj_frame,strWebEdit)
	On error resume next 
	Call StartTime(StrStartTime)
	'WpfWindow("Micro Focus MyFlight Sample").WpfButton("OK").Click

	If strWebEdit <> "" Then
		With JavaWindow(obj_WpfWindow)			
   			If .JavaInternalFrame(obj_frame).JavaButton(strWebEdit).Exist(500) Then
   			   '.JavaInternalFrame(obj_frame).JavaButton(strWebEdit).highlight
   			   .JavaInternalFrame(obj_frame).JavaButton(strWebEdit).Click 
		   
        			Call EndTime(StrEndTime)
			   '	If err.number = 0 Then
        			PrintInfo "micPass", strWebEdit,  " Error Description: " & Err.Description
        			StrStatus = strWebEdit & "Step is Pass"
        			Call CaptureScreenShot(strFileName)
        			Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation)
        			Err.clear						
					Exit sub
				'End If
   			Else
   			    PrintInfo "micFail", "WebButton Click", strWebEdit & " WebButton not found." 
			    Call EndTime(StrEndTime)
   			    Call CaptureScreenShot(strFileName)
				Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation)
				call Rerun(actionName,obj_WpfWindow)

         	End If         	
		End With			
	End If
End Sub

Public Sub MRAppButton_Click_frameBook(actionName,obj_WpfWindow,obj_frame,strWebEdit)
	On error resume next 
	Call StartTime(StrStartTime)
	'WpfWindow("Micro Focus MyFlight Sample").WpfButton("OK").Click

	If strWebEdit <> "" Then
		With JavaWindow(obj_WpfWindow)			
   			If JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame).Exist(60) Then
   			   .JavaInternalFrame(obj_frame).JavaButton(strWebEdit).highlight
   			   .JavaInternalFrame(obj_frame).JavaButton(strWebEdit).Click	
			   
						
			If 	    JavaWindow("MX.3 - MX_PERF_TEST(PROD)_MX").JavaDialog("Warning").JavaButton("Continue").Exist(5) Then
					JavaWindow("MX.3 - MX_PERF_TEST(PROD)_MX").JavaDialog("Warning").JavaButton("Continue").Click
			End If   
			   
			If  JavaWindow("MX.3 - MX_PERF_TEST(PROD)_MX").JavaInternalFrame("Limits Preview").JavaButton("Book with comment").Exist(5) Then
				JavaWindow("MX.3 - MX_PERF_TEST(PROD)_MX").JavaInternalFrame("Limits Preview").JavaButton("Book with comment").Click
				JavaWindow("MX.3 - MX_PERF_TEST(PROD)_MX").JavaInternalFrame("Comment entry").JavaEdit("Comment").Set "Performentest"
				JavaWindow("MX.3 - MX_PERF_TEST(PROD)_MX").JavaInternalFrame("Comment entry").JavaButton("Submit").Click
			End If

		

			
		   
        			Call EndTime(StrEndTime)
			   '	If err.number = 0 Then
        			PrintInfo "micPass", strWebEdit,  " Error Description: " & Err.Description
        			StrStatus = strWebEdit & "Step is Pass"
        			Call CaptureScreenShot(strFileName)
        			Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation)
        			Err.clear						
					Exit sub
				'End If
   			Else
   			    PrintInfo "micFail", "WebButton Click", strWebEdit & " WebButton not found." 
			    Call EndTime(StrEndTime)
   			    Call CaptureScreenShot(strFileName)
				Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation)
				call Rerun(actionName,obj_WpfWindow)
         	End If         	
		End With			
	End If
End Sub


Public Sub MRSelect_ActivateRow(actionName,obj_WpfWindow,obj_frame,strWebEdit,strInValue)

	On error resume next 
	Call StartTime(StrStartTime)
	If strInValue <> "" Then
	'WpfWindow("Micro Focus MyFlight Sample").WpfEdit("agentName").Set
	
		With JavaWindow(obj_WpfWindow)			
   			If .JavaInternalFrame(obj_frame).Exist(60) Then
   			   '.JavaInternalFrame(obj_frame).JavaTable(strWebEdit).highlight
   			   .JavaInternalFrame(obj_frame).JavaTable(strWebEdit).ActivateRow strInValue    
  				Call EndTime(StrEndTime)
			   '	If err.number = 0 Then
        			PrintInfo "micPass", strWebEdit,  " Error Description: " & Err.Description
        			StrStatus = strWebEdit & "Step is Pass"
        			Call CaptureScreenShot(strFileName)
        			Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation)
        			Err.clear						
					Exit sub
				'End If
   			Else
   			    PrintInfo "micFail", strWebEdit, strWebEdit & "WebEdit not found."   
   			    StrStatus = strWebEdit & "Step is False"
   			    Call EndTime(StrEndTime)
   			    Call CaptureScreenShot(strFileName)
				Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation) 
				call Rerun(actionName,obj_WpfWindow)
         	End If         	
		End With			
	End If
End Sub 


Public Sub MRJavatable_Click(actionName,obj_WpfWindow,obj_frame,strWebEdit)

'JavaWindow("MX.3 - MX_PERF_TEST(PROD)_MX_Addall").JavaInternalFrame("Portfolio selection").JavaTable("JfcSimpleTable").Click 200,10,"RIGHT"
	On error resume next 
	Call StartTime(StrStartTime)
	If strWebEdit <> "" Then
	'WpfWindow("Micro Focus MyFlight Sample").WpfEdit("agentName").Set
	
		With JavaWindow(obj_WpfWindow)			
   			If .JavaInternalFrame(obj_frame).Exist(60) Then
   			   '.JavaInternalFrame(obj_frame).JavaTable(strWebEdit).highlight
   			   .JavaInternalFrame(obj_frame).JavaTable(strWebEdit).ClickCell "#0","A","RIGHT" 
	On error resume next   
  				Call EndTime(StrEndTime)
			   '	If err.number = 0 Then
        			PrintInfo "micPass", strWebEdit,  " Error Description: " & Err.Description
        			StrStatus = strWebEdit & "Step is Pass"
        			Call CaptureScreenShot(strFileName)
        			Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation)
        			Err.clear						
					Exit sub
				'End If
   			Else
   			    PrintInfo "micFail", strWebEdit, strWebEdit & "WebEdit not found."   
   			    StrStatus = strWebEdit & "Step is False"
   			    Call EndTime(StrEndTime)
   			    Call CaptureScreenShot(strFileName)
				Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation) 	
				call Rerun(actionName,obj_WpfWindow)

         	End If         	
		End With			
	End If
End Sub 

Public Sub  MRJavatable(actionName,obj_WpfWindow,obj_frame,strWebEdit)
	On error resume next 
	Call StartTime(StrStartTime)
	If strWebEdit <> "" Then
	'WpfWindow("Micro Focus MyFlight Sample").WpfEdit("agentName").Set
	
		With JavaWindow(obj_WpfWindow)			
   			If .JavaInternalFrame(obj_frame).Exist(60) Then
   			   '.JavaInternalFrame(obj_frame).JavaTable(strWebEdit).highlight
   			   .JavaInternalFrame(obj_frame).JavaTable(strWebEdit).ClickCell "#0","A","RIGHT"
	On error resume next   
  				Call EndTime(StrEndTime)
			   '	If err.number = 0 Then
        			PrintInfo "micPass", strWebEdit,  " Error Description: " & Err.Description
        			StrStatus = strWebEdit & "Step is Pass"
        			Call CaptureScreenShot(strFileName)
        			Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation)
        			Err.clear						
					Exit sub
				'End If
   			Else
   			    PrintInfo "micFail", strWebEdit, strWebEdit & "WebEdit not found."   
   			    StrStatus = strWebEdit & "Step is False"
   			    Call EndTime(StrEndTime)
   			    Call CaptureScreenShot(strFileName)
				Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation) 
				call Rerun(actionName,obj_WpfWindow)
         	End If         	
		End With			
	End If
End Sub 



Public Sub MRSelect_SelectRow(actionName,obj_WpfWindow,obj_frame,strWebEdit,strInValue)

	On error resume next 
	Call StartTime(StrStartTime)
	If strInValue <> "" Then
	'WpfWindow("Micro Focus MyFlight Sample").WpfEdit("agentName").Set
	
		With JavaWindow(obj_WpfWindow)			
   			If .JavaInternalFrame(obj_frame).Exist(60) Then
   			   '.JavaInternalFrame(obj_frame).JavaTable(strWebEdit).highlight
   			   .JavaInternalFrame(obj_frame).JavaTable(strWebEdit).SelectRow strInValue    
  				Call EndTime(StrEndTime)
			   '	If err.number = 0 Then
        			PrintInfo "micPass", strWebEdit,  " Error Description: " & Err.Description
        			StrStatus = strWebEdit & "Step is Pass"
        			Call CaptureScreenShot(strFileName)
        			Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation)
        			Err.clear						
					Exit sub
				'End If
   			Else
   			    PrintInfo "micFail", strWebEdit, strWebEdit & "WebEdit not found."   
   			    StrStatus = strWebEdit & "Step is False"
   			    Call EndTime(StrEndTime)
   			    Call CaptureScreenShot(strFileName)
				Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation) 
				call Rerun(actionName,obj_WpfWindow)

         	End If         	
		End With			
	End If
End Sub 


Public Sub MRSelect_Cell(actionName,obj_WpfWindow,obj_frame,strWebEdit,A,B)

	On error resume next 
	Call StartTime(StrStartTime)
	If strWebEdit <> "" Then
	'JavaWindow("MX.3 - MX_PERF_TEST(PROD)_MX").JavaInternalFrame("MModalInternalFrame").JavaTable("Searching in Label").SelectCell "#3","A"
	
		With JavaWindow(obj_WpfWindow)			
   			If .JavaInternalFrame(obj_frame).Exist(60) Then
   			   '.JavaInternalFrame(obj_frame).JavaTable(strWebEdit).highlight
   			   .JavaInternalFrame(obj_frame).JavaTable(strWebEdit).SelectCell A,B   
  				Call EndTime(StrEndTime)
			   '	If err.number = 0 Then
        			PrintInfo "micPass", strWebEdit,  " Error Description: " & Err.Description
        			StrStatus = strWebEdit & "Step is Pass"
        			Wait 1
        			Call CaptureScreenShot(strFileName)
        			Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation)
        			Err.clear						
					Exit sub
				'End If
   			Else
   			    PrintInfo "micFail", strWebEdit, strWebEdit & "WebEdit not found."   
   			    StrStatus = strWebEdit & "Step is False"
   			    Call EndTime(StrEndTime)
   			    Wait 1
   			    Call CaptureScreenShot(strFileName)
				Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation) 	
				'Call Killservicejava
				call Rerun(actionName,obj_WpfWindow)

         	End If         	
		End With			
	End If
End Sub 



'JavaWindow("MX.3 - MX_PERF_TEST(PROD)_MX_Addall").JavaInternalFrame("Portfolio selection").JavaTable("JfcSimpleTable").Click 200,10,"RIGHT"

Public Sub MR_JavaMenu_Select(actionName,obj_WpfWindow,obj_frame,strWebEdit)

	On error resume next 
	Call StartTime(StrStartTime)
	If strWebEdit <> "" Then
	'JavaWindow("MX.3 - MX_PERF_TEST(PROD)_MX_Addall").JavaInternalFrame("Portfolio selection").JavaMenu("Select All").Select
	
		With JavaWindow(obj_WpfWindow)			
   			If .JavaInternalFrame(obj_frame).Exist(60) Then
   			   '.JavaInternalFrame(obj_frame).JavaTable(strWebEdit).highlight
   			   .JavaInternalFrame(obj_frame).JavaMenu(strWebEdit).Select    
  				Call EndTime(StrEndTime)
			   '	If err.number = 0 Then
        			PrintInfo "micPass", strWebEdit,  " Error Description: " & Err.Description
        			StrStatus = strWebEdit & "Step is Pass"
        			Call CaptureScreenShot(strFileName)
        			Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation)
        			Err.clear						
					Exit sub
				'End If
   			Else
   			    PrintInfo "micFail", strWebEdit, strWebEdit & "WebEdit not found."   
   			    StrStatus = strWebEdit & "Step is False"
   			    Call EndTime(StrEndTime)
   			    Call CaptureScreenShot(strFileName)
				Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation) 
				'Call Killservicejava
				call Rerun(actionName,obj_WpfWindow)

         	End If         	
		End With			
	End If
End Sub 


Public Sub MRWebEdit_Type(actionName,obj_WpfWindow,obj_frame,strWebEdit,strInValue,strkey)



	On error resume next 
	Call StartTime(StrStartTime)
	If strInValue <> "" Then
	'WpfWindow("Micro Focus MyFlight Sample").WpfEdit("agentName").Set
	
		With JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame)		
   			If .JavaEdit(strWebEdit).Exist(500) Then
   			   '.JavaEdit(strWebEdit).highlight
   			  ' wait 2
   			   .JavaEdit(strWebEdit).Type strInValue
   			  
				Call Sendkey(strkey)   			   
  				Call EndTime(StrEndTime)
			   '	If err.number = 0 Then
        			PrintInfo "micPass", strWebEdit,  " Error Description: " & Err.Description
        			StrStatus = strWebEdit & "Step is Pass"
        			Call CaptureScreenShot(strFileName)
        			Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation)
        			Err.clear						
					Exit sub
				'End If
   			Else
   			    PrintInfo "micFail", strWebEdit, strWebEdit & "WebEdit not found."   
   			    StrStatus = strWebEdit & "Step is False"
   			    Call EndTime(StrEndTime)
   			    Call CaptureScreenShot(strFileName)
				Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation) 
				'Call Killservicejava
				call Rerun(actionName,obj_WpfWindow)
         	End If         	
		End With			
	End If
End Sub 






Public Sub MRSelect_leg(actionName,obj_WpfWindow,obj_frame,strWebEdit,index,Bill_Race,Set_Method)



	On error resume next 
	Call StartTime(StrStartTime)
	If index <> "" Then
	'WpfWindow("Micro Focus MyFlight Sample").WpfEdit("agentName").Set
	
		With JavaWindow(obj_WpfWindow)			
   			If .JavaInternalFrame(obj_frame).Exist(60) Then
   				.JavaInternalFrame(obj_frame).JavaTable("Pricing").ActivateCell "#24","B"
   			    .JavaInternalFrame(obj_frame).JavaEdit(index).Set Bill_Race
				.JavaInternalFrame(obj_frame).JavaTable(ID_Number).ActivateCell "#25","B"
				.JavaInternalFrame(obj_frame).JavaEdit("ID_Number").Set Set_Method + vbLf + ""
				Call Sendkey("{ENTER}")
  				Call EndTime(StrEndTime)
			   '	If err.number = 0 Then
        			PrintInfo "micPass", strWebEdit,  " Error Description: " & Err.Description
        			StrStatus = strWebEdit & "Step is Pass"
        			Call CaptureScreenShot(strFileName)
        			Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation)
        			Err.clear						
					Exit sub
				'End If
   			Else
   			    PrintInfo "micFail", strWebEdit, strWebEdit & "WebEdit not found."   
   			    StrStatus = strWebEdit & "Step is False"
   			    Call EndTime(StrEndTime)
   			    Call CaptureScreenShot(strFileName)
				Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation) 
				'Call Killservicejava
				call Rerun(actionName,obj_WpfWindow)

         	End If         	
		End With			
	End If
End Sub 


Public Sub MRselectproduct_Contrack(actionName,obj_WpfWindow,obj_frame,strWebEdit)

	On error resume next 
	Call StartTime(StrStartTime)
	If strWebEdit <> "" Then
		With JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame)		
   			If .JavaTable(strWebEdit).Exist(500) Then 
   			   JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame).JavaTable(strWebEdit).ActivateCell "#0","A" 
  				Call EndTime(StrEndTime)
			   '	If err.number = 0 Then
        			PrintInfo "micPass", strWebEdit,  " Error Description: " & Err.Description
        			StrStatus = strWebEdit & "Step is Pass"
        			Call CaptureScreenShot(strFileName)
        			Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation)
        			Err.clear						
					Exit sub
				'End If
   			Else
   			    PrintInfo "micFail", strWebEdit, strWebEdit & "WebEdit not found."   
   			    StrStatus = strWebEdit & "Step is False"
   			    Call EndTime(StrEndTime)
   			    Call CaptureScreenShot(strFileName)
				Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation) 	
				'Call Killservicejava
				call Rerun(actionName,obj_WpfWindow)

         	End If         	
		End With			
	End If
End Sub 




Public Sub MRGetTextTotable(actionName,obj_WpfWindow,obj_frame,strWebEdit,index)


	On error resume next 
	Call StartTime(StrStartTime)
	If index <> "" Then
		With JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame)		
   			If .JavaTable(strWebEdit).Exist(60) Then
   			   strnum = JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame).JavaTable(strWebEdit).Object.getColumnCount
   			   strnum = strnum - 6
   			   strname = JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame).JavaTable(strWebEdit).GetColumnName(strnum)  
   			   JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame).JavaTable(strWebEdit).ActivateCell index,"B"
  				Call EndTime(StrEndTime)
			   '	If err.number = 0 Then
        			PrintInfo "micPass", strWebEdit,  " Error Description: " & Err.Description
        			StrStatus = strWebEdit & "Step is Pass"
        			Call CaptureScreenShot(strFileName)
        			Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation)
        			Err.clear						
					Exit sub
				'End If
   			Else
   			    PrintInfo "micFail", strWebEdit, strWebEdit & "WebEdit not found."   
   			    StrStatus = strWebEdit & "Step is False"
   			    Call EndTime(StrEndTime)
   			    Call CaptureScreenShot(strFileName)
				Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation) 	
				'Call Killservicejava
				call Rerun(actionName,obj_WpfWindow)
         	End If         	
		End With			
	End If
End Sub 




Public Sub MRInputText_Amount(actionName,obj_WpfWindow,obj_frame,strWebEdit,index,Amount,Type_Amount,Id_Number)

	On error resume next 
	
	If strWebEdit <> "" Then
		With JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame)		
   			If .JavaTable(strWebEdit).Exist(60) Then
   			wait 1
'   			amountVal = 
   				Do  Until JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame).JavaTable(strWebEdit).GetCellData("#0","C") > 0 
	   			   wait 1
	   			   JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame).JavaTable(strWebEdit).ActivateCell "#0","C"
	   			   JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame).JavaEdit(index).Set Amount + vbLf + ""
				   
 			   loop
 			   Call StartTime(StrStartTime)
' 			   wait 2
			   JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame).JavaTable(strWebEdit).ActivateCell "#2","B"
	 		   JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame).JavaEdit("Pricing:").Set Type_Amount + vbLf + ""
' 			   JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame).JavaTable(strWebEdit).ActivateCell "#17","C"
'   			   JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame).JavaEdit("ID_Number").Set Id_Number + vbLf + ""
				wait 2
   			   JavaWindow("MX.3 - MX_PERF_TEST(PROD)_MX").JavaInternalFrame("e-Tradepad").JavaTable("4 legs").ActivateCell "#20","B"
			   'JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame).JavaTable(strWebEdit).ActivateCell "#20","B"
			   JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame).JavaEdit("Pricing:").Set "1050"
			  
  				Call EndTime(StrEndTime)
			   '	If err.number = 0 Then
        			PrintInfo "micPass", strWebEdit,  " Error Description: " & Err.Description
        			StrStatus = strWebEdit & "Step is Pass"
        			Call CaptureScreenShot(strFileName)
        			Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation)
        			Err.clear						
					Exit sub
				'End If
   			Else
   			    PrintInfo "micFail", strWebEdit, strWebEdit & "WebEdit not found."   
   			    StrStatus = strWebEdit & "Step is False"
   			    Call EndTime(StrEndTime)
   			    Call CaptureScreenShot(strFileName)
				Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation) 
				'Call Killservicejava
				call Rerun(actionName,obj_WpfWindow)

         	End If         	
		End With			
	End If
End Sub 



Public Sub MRInputText_Tc2(actionName,obj_WpfWindow,obj_frame,strWebEdit)


	On error resume next 
	
	If strWebEdit <> "" Then
		With JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame)		
   			If .JavaTable(strWebEdit).Exist(60) Then
   				wait 3
   				Call StartTime(StrStartTime)
'   			   JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame).JavaTable(strWebEdit).ActivateCell "#0","C"
'   			   JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame).JavaEdit(index).Set Amount + vbLf + ""
'   			   wait 2
'			   JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame).JavaTable(strWebEdit).ActivateCell "#2","B"
' 			   JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame).JavaEdit("Pricing:").Set Type_Amount + vbLf + ""

 	
 			   JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame).JavaTable(strWebEdit).ActivateCell "#0","C"
				JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame).JavaEdit("Pricing").Set "1000" + vbLf + ""
				wait 2
				JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame).JavaTable(strWebEdit).ActivateCell "#7","C"
				avaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame).JavaEdit("ID_Number").Set Counterparty '"SCG8_XXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
				wait 2
				JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame).JavaTable(strWebEdit).ActivateCell "#9","B"
				avaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame).JavaEdit("ID_Number").Set "BKTBDTD_BD_GMG"

			  
  				Call EndTime(StrEndTime)
			   '	If err.number = 0 Then
        			PrintInfo "micPass", strWebEdit,  " Error Description: " & Err.Description
        			StrStatus = strWebEdit & "Step is Pass"
        			Call CaptureScreenShot(strFileName)
        			Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation)
        			Err.clear						
					Exit sub
				'End If
   			Else
   			    PrintInfo "micFail", strWebEdit, strWebEdit & "WebEdit not found."   
   			    StrStatus = strWebEdit & "Step is False"
   			    Call EndTime(StrEndTime)
   			    Call CaptureScreenShot(strFileName)
				Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation) 
				'Call Killservicejava
				call Rerun(actionName,obj_WpfWindow)
         	End If         	
		End With			
	End If
End Sub 




Public Sub MRselect_BOT_P_LT(actionName,obj_WpfWindow,obj_frame,strWebEdit,strtable,Id_BOT)

	On error resume next 
	Call StartTime(StrStartTime)
	If strWebEdit <> "" Then
		With JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame)		
   			If .JavaTable(strWebEdit).Exist(60) Then 
   				JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame).JavaEdit(strWebEdit).Set Id_BOT
				JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame).JavaEdit(strWebEdit).Activate
   			   JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame).JavaTable(strtable).ActivateCell "#0","A" 
   			   
  				Call EndTime(StrEndTime)
			   '	If err.number = 0 Then
        			PrintInfo "micPass", strWebEdit,  " Error Description: " & Err.Description
        			StrStatus = strWebEdit & "Step is Pass"
        			Call CaptureScreenShot(strFileName)
        			Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation)
        			Err.clear						
					Exit sub
				'End If
   			Else
   			    PrintInfo "micFail", strWebEdit, strWebEdit & "WebEdit not found."   
   			    StrStatus = strWebEdit & "Step is False"
   			    Call EndTime(StrEndTime)
   			    Call CaptureScreenShot(strFileName)
				Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation) 
				'Call Killservicejava
				call Rerun(actionName,obj_WpfWindow)
         	End If         	
		End With			
	End If
End Sub 


Public Sub MR_DelectLeg(actionName,obj_WpfWindow,obj_frame,strWebEdit,strtable,Manu)

On error resume next 
	Call StartTime(StrStartTime)
	If strWebEdit <> "" Then
		With JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame)		
   			If .JavaTable(strWebEdit).Exist(60) Then 
   				JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame).JavaTable(strtable).SelectCell "#2","H"
				JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame).JavaTable(strtable).Click  966,39,"RIGHT"
				JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame).JavaMenu(Manu).Select
   			   
  				Call EndTime(StrEndTime)
			   '	If err.number = 0 Then
        			PrintInfo "micPass", strWebEdit,  " Error Description: " & Err.Description
        			StrStatus = strWebEdit & "Step is Pass"
        			Call CaptureScreenShot(strFileName)
        			Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation)
        			Err.clear						
					Exit sub
				'End If
   			Else
   			    PrintInfo "micFail", strWebEdit, strWebEdit & "WebEdit not found."   
   			    StrStatus = strWebEdit & "Step is False"
   			    Call EndTime(StrEndTime)
   			    Call CaptureScreenShot(strFileName)
				Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation) 	
				'Call Killservicejava
				call Rerun(actionName,obj_WpfWindow)
         	End If         	
		End With			
	End If


End Sub 



Public Sub MRAppButton_Click_Dinalog(actionName,obj_WpfWindow,obj_frame,strWebEdit)
	On error resume next 
	Call StartTime(StrStartTime)
	'WpfWindow("Micro Focus MyFlight Sample").WpfButton("OK").Click

	If strWebEdit <> "" Then
		With JavaWindow(obj_WpfWindow)			
   			If .JavaDialog(obj_frame).Exist(60) Then
   			   '.JavaDialog(obj_frame).JavaButton(strWebEdit).highlight
   			   Call CaptureScreenShot(strFileName)
   			   .JavaDialog(obj_frame).JavaButton(strWebEdit).click   		
        			Call EndTime(StrEndTime)
			   '	If err.number = 0 Then
        			PrintInfo "micPass", strWebEdit,  " Error Description: " & Err.Description
        			StrStatus = strWebEdit & "Step is Pass"
        			Call CaptureScreenShot(strFileName)
        			Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation)
        			Err.clear						
					Exit sub
				'End If
   			Else
   			    PrintInfo "micFail", "WebButton Click", strWebEdit & " WebButton not found." 
			    Call EndTime(StrEndTime)
   			    Call CaptureScreenShot(strFileName)
				Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation)  
				Call Killservicejava
				call Rerun(actionName,obj_WpfWindow)
         	End If         	
		End With			
	End If
End Sub




Public Sub MRSelectCurrency(actionName,obj_WpfWindow,obj_frame,strWebEdit,index)


	On error resume next 
	Call StartTime(StrStartTime)
	If index <> "" Then
		With JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame)		
   			If .JavaTable(strWebEdit).Exist(60) Then
'   			   strnum = JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame).JavaTable(strWebEdit).Object.getColumnCount
'   			   strnum = strnum - 6
'   			   strname = JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame).JavaTable(strWebEdit).GetColumnName(strnum)  
   			   JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame).JavaTable(strWebEdit).ActivateCell index,"B"
  				Call EndTime(StrEndTime)
			   '	If err.number = 0 Then
        			PrintInfo "micPass", strWebEdit,  " Error Description: " & Err.Description
        			StrStatus = strWebEdit & "Step is Pass"
        			Call CaptureScreenShot(strFileName)
        			Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation)
        			Err.clear						
					Exit sub
				'End If
   			Else
   			    PrintInfo "micFail", strWebEdit, strWebEdit & "WebEdit not found."   
   			    StrStatus = strWebEdit & "Step is False"
   			    Call EndTime(StrEndTime)
   			    Call CaptureScreenShot(strFileName)
				Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation) 	
				'Call Killservicejava
				call Rerun(actionName,obj_WpfWindow)
         	End If         	
		End With			
	End If
End Sub 


Public Sub MRInputAmount(actionName,obj_WpfWindow,obj_frame,strWebEdit,index,Amount)
  
	On error resume next 
	
	If index <> "" Then
	wait 3
		With JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame)		
   			If .JavaTable(strWebEdit).Exist(60) Then
   			wait 2
				
				Do 
				  amountVal = JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame).JavaTable(strWebEdit).GetCellData("#0","C")	
				  JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame).JavaTable(strWebEdit).ActivateCell "#0","C"
			      JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame).JavaEdit(index).Set Amount + vbLf + ""
				Loop  Until JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame).JavaTable(strWebEdit).GetCellData("#0","C") > 0
	   				Call StartTime(StrStartTime)

	
'	
'		With JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame)		
'   			If .JavaTable(strWebEdit).Exist(60) Then
'   			wait 3
'			   Call StartTime(StrStartTime)
''			    tt = JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame).JavaEdit(index).GetROProperty("value")
''			    MsgBox tt
'			   JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame).JavaTable(strWebEdit).ActivateCell "#0","C"
'			   JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame).JavaEdit(index).set Amount + vbLf + ""
			
  				Call EndTime(StrEndTime)
			   '	If err.number = 0 Then
        			PrintInfo "micPass", strWebEdit,  " Error Description: " & Err.Description
        			StrStatus = strWebEdit & "Step is Pass"
        			Call CaptureScreenShot(strFileName)
        			Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation)
        			Err.clear						
					Exit sub
				'End If
			
   			Else
   			    PrintInfo "micFail", strWebEdit, strWebEdit & "WebEdit not found."   
   			    StrStatus = strWebEdit & "Step is False"
   			    Call EndTime(StrEndTime)
   			    Call CaptureScreenShot(strFileName)
				Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation) 
				'Call Killservicejava
				call Rerun(actionName,obj_WpfWindow)
         	End If         	
		End With			
	End If

End Sub 

Public Sub MRInputConverdate(actionName,obj_WpfWindow,obj_frame,strWebEdit,index,Converdate)


	On error resume next 
	Call StartTime(StrStartTime)
	If index <> "" Then
		With JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame)		
   			If .JavaTable(strWebEdit).Exist(60) Then
			   JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame).JavaTable(strWebEdit).ActivateCell "#2","B"
			   wait 1
			   JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame).JavaEdit(index).Set Converdate + vbLf + ""
			   JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame).JavaTable(strWebEdit).ActivateCell "#2","D"
			   wait 1
			   JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame).JavaEdit("Pricing:").Set "1" + vbLf + ""


  				Call EndTime(StrEndTime)
			   '	If err.number = 0 Then
        			PrintInfo "micPass", strWebEdit,  " Error Description: " & Err.Description
        			StrStatus = strWebEdit & "Step is Pass"
        			Call CaptureScreenShot(strFileName)
        			Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation)
        			Err.clear						
					Exit sub
				'End If
   			Else
   			    PrintInfo "micFail", strWebEdit, strWebEdit & "WebEdit not found."   
   			    StrStatus = strWebEdit & "Step is False"
   			    Call EndTime(StrEndTime)
   			    Call CaptureScreenShot(strFileName)
				Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation) 
				'Call Killservicejava
				call Rerun(actionName,obj_WpfWindow)
         	End If         	
		End With			
	End If
End Sub 

Public Sub MRInputCounter(actionName,obj_WpfWindow,obj_frame,strWebEdit,index,Id_Number)


	On error resume next 
	Call StartTime(StrStartTime)
	If index <> "" Then
		With JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame)		
   			If .JavaTable(strWebEdit).Exist(60) Then
				wait 1
			   JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame).JavaTable(strWebEdit).ActivateCell "#16","C"
			   JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame).JavaEdit(index).Set Id_Number + vbLf + ""

'			   wait 1
			   JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame).JavaTable(strWebEdit).ActivateCell "#17","C"
			   JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame).JavaEdit("Pricing:").Set "Authorize person" + vbLf + ""

  				Call EndTime(StrEndTime)
			   '	If err.number = 0 Then
        			PrintInfo "micPass", strWebEdit,  " Error Description: " & Err.Description
        			StrStatus = strWebEdit & "Step is Pass"
        			Call CaptureScreenShot(strFileName)
        			Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation)
        			Err.clear						
					Exit sub
				'End If
   			Else
   			    PrintInfo "micFail", strWebEdit, strWebEdit & "WebEdit not found."   
   			    StrStatus = strWebEdit & "Step is False"
   			    Call EndTime(StrEndTime)
   			    Call CaptureScreenShot(strFileName)
				Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation) 		
				'Call Killservicejava
				call Rerun(actionName,obj_WpfWindow)
         	End If         	
		End With			
	End If
End Sub 

Public Sub MRClickGood(actionName,obj_WpfWindow,obj_frame,strWebEdit,index)


	On error resume next 
	Call StartTime(StrStartTime)
	If index <> "" Then
		With JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame)		
   			If .JavaTable(strWebEdit).Exist(60) Then
   				
				JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame).JavaTable(strWebEdit).ActivateCell "#19","B"
 				JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame).JavaEdit(index).Set "1805" + vbLf + ""
			    JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame).JavaTable(strWebEdit).ActivateCell "#20","C"

  				Call EndTime(StrEndTime)
			   '	If err.number = 0 Then
        			PrintInfo "micPass", strWebEdit,  " Error Description: " & Err.Description
        			StrStatus = strWebEdit & "Step is Pass"
        			Call CaptureScreenShot(strFileName)
        			Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation)
        			Err.clear						
					Exit sub
				'End If
   			Else
   			    PrintInfo "micFail", strWebEdit, strWebEdit & "WebEdit not found."   
   			    StrStatus = strWebEdit & "Step is False"
   			    Call EndTime(StrEndTime)
   			    Call CaptureScreenShot(strFileName)
				Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation) 	
				'Call Killservicejava
				call Rerun(actionName,obj_WpfWindow)
         	End If         	
		End With			
	End If
End Sub 


Public Sub MRSelectMulti(actionName,obj_WpfWindow,obj_frame,strWebEdit)


	On error resume next 
	Call StartTime(StrStartTime)
	If strWebEdit <> "" Then
		With JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame)		
   			If .JavaTable(strWebEdit).Exist(60) Then
					JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame).JavaEdit(index).Set "multi" + vbLf + ""
					'JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame).JavaEdit(index).Activate

  				Call EndTime(StrEndTime)
			   '	If err.number = 0 Then
        			PrintInfo "micPass", strWebEdit,  " Error Description: " & Err.Description
        			StrStatus = strWebEdit & "Step is Pass"
        			Call CaptureScreenShot(strFileName)
        			Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation)
        			Err.clear						
					Exit sub
				'End If
   			Else
   			    PrintInfo "micFail", strWebEdit, strWebEdit & "WebEdit not found."   
   			    StrStatus = strWebEdit & "Step is False"
   			    Call EndTime(StrEndTime)
   			    Call CaptureScreenShot(strFileName)
				Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation) 		
				'Call Killservicejava
				call Rerun(actionName,obj_WpfWindow)
         	End If         	
		End With			
	End If
End Sub 



Public Sub MRExitProduct(actionName,obj_WpfWindow,obj_frame,strWebEdit)


	On error resume next 
	If strWebEdit <> "" Then
	For i = 1 To 10
		With JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame)		
   			If JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame).JavaButton(strWebEdit).Exist(10) Then
					JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame).JavaButton(strWebEdit).Click
   			Else		   			
		   			exit for 		   			
         	End If 	
		End With
	Next
	Else
		call Rerun(actionName,obj_WpfWindow)
	End If
End Sub 


Public Sub MRInputCounterOnSplit(actionName,obj_WpfWindow,obj_frame,strWebEdit,Id_Number)


	On error resume next 
	
	If Id_Number <> "" Then
		With JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame)		
   			If .JavaTable(strWebEdit).Exist(60) Then
'				wait 1
'			   JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame).JavaTable(strWebEdit).ActivateCell "#16","C"
'			   JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame).JavaEdit(index).Set Id_Number + vbLf + ""
'			   JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame).JavaEdit(index).Activate
'			   wait 1
			    JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame).JavaTable(strWebEdit).ActivateCell "#9","C"
			    Call StartTime(StrStartTime)
			    JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame).JavaEdit(strWebEdit).Set Id_Number + vbLf + ""

  				Call EndTime(StrEndTime)
			   '	If err.number = 0 Then
        			PrintInfo "micPass", strWebEdit,  " Error Description: " & Err.Description
        			StrStatus = strWebEdit & "Step is Pass"
        			Call CaptureScreenShot(strFileName)
        			Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation)
        			Err.clear						
					Exit sub
				'End If
   			Else
   			    PrintInfo "micFail", strWebEdit, strWebEdit & "WebEdit not found."   
   			    StrStatus = strWebEdit & "Step is False"
   			    Call EndTime(StrEndTime)
   			    Call CaptureScreenShot(strFileName)
				Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation) 	
				'Call Killservicejava
				call Rerun(actionName,obj_WpfWindow)
         	End If         	
		End With			
	End If
End Sub 



Public Sub MRGettextOnDinalog(actionName,obj_WpfWindow,strWebEdit,strvalue)



	On error resume next 
	Call StartTime(StrStartTime)
	If obj_WpfWindow <> "" Then	
	'WpfWindow("Micro Focus MyFlight Sample").WpfObject("John Smith").GetROProperty("text")


		With WpfWindow(obj_WpfWindow)			
   			If .JavaDialog(strWebEdit).Exist(60) Then
   				strWebEdit = JavaWindow(obj_WpfWindow).JavaDialog(strWebEdit).JavaTable(strvalue).GetROProperty("text")
   			   
  				Call EndTime(StrEndTime)
			   '	If err.number = 0 Then
        			PrintInfo "micPass", strWebEdit,  " Error Description: " & Err.Description
        			StrStatus = strWebEdit & "Step is Pass"
        			Call CaptureScreenShot(strFileName)
        			Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation)
        			Err.clear						
					Exit sub
				'End If
   			Else
   			    PrintInfo "micFail", strWebEdit, strWebEdit & "WebEdit not found."   
   			    StrStatus = strWebEdit & "Step is False"
   			    Call EndTime(StrEndTime)
   			    Call CaptureScreenShot(strFileName)
				Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation) 
				'Call Killservicejava
				call Rerun(actionName,obj_WpfWindow)
         	End If         	
		End With			
	End If
End Sub 




Public Sub MRInserTier(actionName,obj_WpfWindow,strWebEdit,strvalue)



	On error resume next 
	Call StartTime(StrStartTime)
	
	'WpfWindow("Micro Focus MyFlight Sample").WpfObject("John Smith").GetROProperty("text")


		With WpfWindow(obj_WpfWindow)			
   			If .JavaInternalFrame(strWebEdit).Exist(500) Then
   			JavaWindow(obj_WpfWindow).JavaInternalFrame(strWebEdit).JavaTable(strvalue).SelectCell "#0","A"
   			JavaWindow(obj_WpfWindow).JavaInternalFrame(strWebEdit).JavaTable(strvalue).Click 74,7,"RIGHT"
   			Call Sendkey ("^9")
 

  				Call EndTime(StrEndTime)
			   '	If err.number = 0 Then
        			PrintInfo "micPass", strWebEdit,  " Error Description: " & Err.Description
        			StrStatus = strWebEdit & "Step is Pass"
        			Call CaptureScreenShot(strFileName)
        			Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation)
        			Err.clear						
					Exit sub
				'End If
   			Else
   			    PrintInfo "micFail", strWebEdit, strWebEdit & "WebEdit not found."   
   			    StrStatus = strWebEdit & "Step is False"
   			    Call EndTime(StrEndTime)
   			    Call CaptureScreenShot(strFileName)
				Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation)
				'Call Killservicejava
				call Rerun(actionName,obj_WpfWindow)
         	End If         	
		End With			

End Sub 



Public Sub MROpenApp(actionName,strWebEdit)

	
	On error resume next 
	Call StartTime(StrStartTime)
	If strWebEdit <> "" Then	
		Call Killservicejava
 		SystemUtil.Run strWebEdit

  				Call EndTime(StrEndTime)
			   '	If err.number = 0 Then
        			PrintInfo "micPass", strWebEdit,  " Error Description: " & Err.Description
        			StrStatus = strWebEdit & "Step is Pass"
        			Call CaptureScreenShot(strFileName)
        			Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation)
        			Err.clear						
					Exit sub
				'End If
   			Else
   			    PrintInfo "micFail", strWebEdit, strWebEdit & "WebEdit not found."   
   			    StrStatus = strWebEdit & "Step is False"
   			    Call EndTime(StrEndTime)
   			    Call CaptureScreenShot(strFileName)
				Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation)
				call Rerun(actionName,obj_WpfWindow)
         	End If         				

End Sub



Public Sub MRClosesApp(actionName,strWebEdit)



	On error resume next 
	Call StartTime(StrStartTime)
	If strWebEdit <> "" Then


  				Call EndTime(StrEndTime)
			   '	If err.number = 0 Then
        			PrintInfo "micPass", strWebEdit,  " Error Description: " & Err.Description
        			StrStatus = strWebEdit & "Step is Pass"
        			Call CaptureScreenShot(strFileName)
        			Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation)
        			Err.clear						
					Exit sub
				'End If
   			Else
   			    PrintInfo "micFail", strWebEdit, strWebEdit & "WebEdit not found."   
   			    StrStatus = strWebEdit & "Step is False"
   			    Call EndTime(StrEndTime)
   			    Call CaptureScreenShot(strFileName)
				Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation)
				call Rerun(actionName,obj_WpfWindow)
         	End If         				

End Sub




Public Function GetparamiterTransactionStart(StrStartTime,Transation)

	On error resume next 
	iRowCount = Datatable.getSheet("LogTimeTransaction").getRowCount
 	Datatable.getSheet("LogTimeTransaction").setCurrentRow(iRowCount+1)
 	DataTable.Value("TransactionName","LogTimeTransaction") = Transation
 	DataTable.Value("StartTimeTransactionName","LogTimeTransaction") = StrStartTime
 	'DataTable.Value("EndTimeTransactionName","LogTimeTransaction") = StrEndTime
 	
 	
End Function  

Public Function GetparamiterTransactionEnd(StrEndTime,Transation)

	On error resume next 
	iRowCount = Datatable.getSheet("LogTimeTransaction").getRowCount
 	Datatable.getSheet("LogTimeTransaction").setCurrentRow(iRowCount)
 	DataTable.Value("TransactionName","LogTimeTransaction") = Transation
 	'DataTable.Value("StartTimeTransactionName","LogTimeTransaction") = StrStartTime
 	DataTable.Value("EndTimeTransactionName","LogTimeTransaction") = StrEndTime
 	
 	
End Function 




Public Function MRCheckPage(actionName,strWebEdit)

	Call StartTime(StrStartTime)
 	With JavaWindow(strWebEdit)			
   			If JavaWindow(strWebEdit).Exist(60) Then

  				Call EndTime(StrEndTime)
			   '	If err.number = 0 Then
        			PrintInfo "micPass", strWebEdit,  " Error Description: " & Err.Description
        			StrStatus = strWebEdit & "Step is Pass"
        			Call CaptureScreenShot(strFileName)
        			Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation)
        			Err.clear						
					'Exit sub
				'End If
   			Else
   			    PrintInfo "micFail", strWebEdit, strWebEdit & "WebEdit not found."   
   			    StrStatus = strWebEdit & "Step is False"
   			    Call EndTime(StrEndTime)
   			    Call CaptureScreenShot(strFileName)
				Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation) 
				call Rerun(actionName,obj_WpfWindow)
         	End If         	
		End With	
 	
End Function 


Public Sub MRGettextOnJavaStaticText(actionName,obj_WpfWindow,strWebEdit)

	On error resume next 
	Call StartTime(StrStartTime)
	If obj_WpfWindow <> "" Then	
	'WpfWindow("Micro Focus MyFlight Sample").WpfObject("John Smith").GetROProperty("text")


		With WpfWindow(obj_WpfWindow)			
   			If .JavaStaticText(strWebEdit).Exist(60) Then
   				strWebEdit = JavaWindow(obj_WpfWindow).JavaStaticText(strWebEdit).GetROProperty("text")
   			   
  				Call EndTime(StrEndTime)
			   '	If err.number = 0 Then
        			PrintInfo "micPass", strWebEdit,  " Error Description: " & Err.Description
        			StrStatus = strWebEdit & "Step is Pass"
        			Call CaptureScreenShot(strFileName)
        			Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation)
        			Err.clear						
					Exit sub
				'End If
   			Else
   			    PrintInfo "micFail", strWebEdit, strWebEdit & "WebEdit not found."   
   			    StrStatus = strWebEdit & "Step is False"
   			    Call EndTime(StrEndTime)
   			    Call CaptureScreenShot(strFileName)
				Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation)
				call Rerun(actionName,obj_WpfWindow)
         	End If         	
		End With			
	End If
End Sub 


Public Sub MRCheckButtonInternalFrame(actionName,obj_WpfWindow,strWebEdit,strvalue)



	On error resume next 
	Call StartTime(StrStartTime)
	If obj_WpfWindow <> "" Then	
	'WpfWindow("Micro Focus MyFlight Sample").WpfObject("John Smith").GetROProperty("text")


		With WpfWindow(obj_WpfWindow)			
   			If .JavaInternalFrame(strWebEdit).JavaButton(strvalue).Exist(60) Then		   
  				Call EndTime(StrEndTime)
			   '	If err.number = 0 Then
        			PrintInfo "micPass", strWebEdit,  " Error Description: " & Err.Description
        			StrStatus = strWebEdit & "Step is Pass"
        			Call CaptureScreenShot(strFileName)
        			Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation)
        			Err.clear						
					Exit sub
				'End If
   			Else
   			    PrintInfo "micFail", strWebEdit, strWebEdit & "WebEdit not found."   
   			    StrStatus = strWebEdit & "Step is False"
   			    Call EndTime(StrEndTime)
   			    Call CaptureScreenShot(strFileName)
				Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation)
				call Rerun(actionName,obj_WpfWindow)
         	End If         	
		End With			
	End If
End Sub 



Public Sub MRGettextOnJavaInternalframe(actionName,obj_WpfWindow,strWebEdit,strvalue)

	On error resume next 
	Call StartTime(StrStartTime)
	If obj_WpfWindow <> "" Then	
	'WpfWindow("Micro Focus MyFlight Sample").WpfObject("John Smith").GetROProperty("text")


		With WpfWindow(obj_WpfWindow)			
   			If .JavaStaticText(strWebEdit).Exist(60) Then
   				strWebEdit = JavaWindow(obj_WpfWindow).JavaInternalFrame(strWebEdit).JavaStaticText(strvalue).GetROProperty("text")
   			   
  				Call EndTime(StrEndTime)
			   '	If err.number = 0 Then
        			PrintInfo "micPass", strWebEdit,  " Error Description: " & Err.Description
        			StrStatus = strWebEdit & "Step is Pass"
        			Call CaptureScreenShot(strFileName)
        			Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation)
        			Err.clear						
					Exit sub
				'End If
   			Else
   			    PrintInfo "micFail", strWebEdit, strWebEdit & "WebEdit not found."   
   			    StrStatus = strWebEdit & "Step is False"
   			    Call EndTime(StrEndTime)
   			    Call CaptureScreenShot(strFileName)
				Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation) 
				call Rerun(actionName,obj_WpfWindow)
         	End If         	
		End With			
	End If
End Sub 


Public Sub MRInputText_AmountPorata(actionName,obj_WpfWindow,obj_frame,strWebEdit,index,Amount,Type_Amount,Id_Number)


	On error resume next 
	If strWebEdit <> "" Then
		With JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame)		
   			If .JavaTable(strWebEdit).Exist(60) Then
   			
   			wait 1
'			amountVal = JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame).JavaTable(strWebEdit).GetCellData("#0","C")
			
			Do  
			  wait 1
			  amountVal = JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame).JavaTable(strWebEdit).GetCellData("#0","C")	
			  JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame).JavaTable(strWebEdit).ActivateCell "#0","C"
		      JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame).JavaEdit(index).Set Amount + vbLf + ""
			Loop Until JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame).JavaTable(strWebEdit).GetCellData("#0","C") > 0
   				Call StartTime(StrStartTime)
'   				wait 2
'   			   JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame).JavaTable(strWebEdit).ActivateCell "#0","C"
'   			   JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame).JavaEdit(index).Set Amount + vbLf + ""
'
'   			   JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame).JavaEdit("ID_Number").Set Id_Number + vbLf + ""
'			   JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame).JavaTable(strWebEdit).ActivateCell "#17","B"
'		   		  amountVal = JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame).JavaTable(strWebEdit).GetCellData("#0","C")
'			MsgBox ppd
'			print "note" &ppd
  				Call EndTime(StrEndTime)
			   '	If err.number = 0 Then
        			PrintInfo "micPass", strWebEdit,  " Error Description: " & Err.Description
        			StrStatus = strWebEdit & "Step is Pass"
        			Call CaptureScreenShot(strFileName)
        			Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation)
        			Err.clear						
					Exit sub
				'End If
   			Else
   			    PrintInfo "micFail", strWebEdit, strWebEdit & "WebEdit not found."   
   			    StrStatus = strWebEdit & "Step is False"
   			    Call EndTime(StrEndTime)
   			    Call CaptureScreenShot(strFileName)
				Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation) 		
				'Call Killservicejava
				call Rerun(actionName,obj_WpfWindow)
         	End If         	
		End With			
	End If
End Sub 



Public Sub MRSelect_legPorata(actionName,obj_WpfWindow,obj_frame,strWebEdit,index,Bill_Race,Set_Method)



	On error resume next 
	Call StartTime(StrStartTime)
	If index <> "" Then
	'WpfWindow("Micro Focus MyFlight Sample").WpfEdit("agentName").Set
	
		With JavaWindow(obj_WpfWindow)			
   			If .JavaInternalFrame(obj_frame).Exist(500) Then
   				.JavaInternalFrame(obj_frame).JavaTable(strWebEdit).ActivateCell "#21","B"
   			    .JavaInternalFrame(obj_frame).JavaEdit(index).Set Bill_Race
				.JavaInternalFrame(obj_frame).JavaTable(strWebEdit).ActivateCell "#22","B"
				wait 1
				.JavaInternalFrame(obj_frame).JavaEdit("ID_Number").Set Set_Method + vbLf + ""
				Call Sendkey("{ENTER}")
  				Call EndTime(StrEndTime)
			   '	If err.number = 0 Then
        			PrintInfo "micPass", strWebEdit,  " Error Description: " & Err.Description
        			StrStatus = strWebEdit & "Step is Pass"
        			Call CaptureScreenShot(strFileName)
        			Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation)
        			Err.clear						
					Exit sub
				'End If
   			Else
   			    PrintInfo "micFail", strWebEdit, strWebEdit & "WebEdit not found."   
   			    StrStatus = strWebEdit & "Step is False"
   			    Call EndTime(StrEndTime)
   			    Call CaptureScreenShot(strFileName)
				Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation) 	
				'Call Killservicejava
				call Rerun(actionName,obj_WpfWindow)
         	End If         	
		End With			
	End If
End Sub 



Public Sub MRInputID_Number(actionName,obj_WpfWindow,obj_frame,strWebEdit,index,Id_Number)


	On error resume next 
	
	If Id_Number <> "" Then
		With JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame)		
   			If .JavaTable(strWebEdit).Exist(60) Then
'				wait 1
'			   JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame).JavaTable(strWebEdit).ActivateCell "#16","C"
'			   JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame).JavaEdit(index).Set Id_Number + vbLf + ""
'			   JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame).JavaEdit(index).Activate
'			   wait 1
			    JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame).JavaTable(strWebEdit).ActivateCell "#14","C"
			    Call StartTime(StrStartTime)
			    JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame).JavaEdit(index).Set Id_Number + vbLf + ""
			    wait 2
			    JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame).JavaTable(strWebEdit).ActivateCell "#18","B"
			    

  				Call EndTime(StrEndTime)
			   '	If err.number = 0 Then
        			PrintInfo "micPass", strWebEdit,  " Error Description: " & Err.Description
        			StrStatus = strWebEdit & "Step is Pass"
        			Call CaptureScreenShot(strFileName)
        			Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation)
        			Err.clear						
					Exit sub
				'End If
   			Else
   			    PrintInfo "micFail", strWebEdit, strWebEdit & "WebEdit not found."   
   			    StrStatus = strWebEdit & "Step is False"
   			    Call EndTime(StrEndTime)
   			    Call CaptureScreenShot(strFileName)
				Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation) 	
				'Call Killservicejava
				call Rerun(actionName,obj_WpfWindow)
         	End If         	
		End With			
	End If
End Sub 





Public Sub MRChackobjectpage(actionName,obj_WpfWindow,strWebEdit,strvalue)

	On error resume next 
	Call StartTime(StrStartTime)
	If obj_WpfWindow <> "" Then	
	'WpfWindow("Micro Focus MyFlight Sample").WpfObject("John Smith").GetROProperty("text")


		With JavaWindow(obj_WpfWindow)			
   			If .JavaInternalFrame(strWebEdit).JavaTree(strvalue).Exist(60) Then
   			   
  				Call EndTime(StrEndTime)
			   '	If err.number = 0 Then
        			PrintInfo "micPass", strWebEdit,  " Error Description: " & Err.Description
        			StrStatus = strWebEdit & "Step is Pass"
        			Call CaptureScreenShot(strFileName)
        			Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation)
        			Err.clear						
					Exit sub
				'End If
   			Else
   			    PrintInfo "micFail", strWebEdit, strWebEdit & "WebEdit not found."   
   			    StrStatus = strWebEdit & "Step is False"
   			    Call EndTime(StrEndTime)
   			    Call CaptureScreenShot(strFileName)
				Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation) 	
				'Call Killservicejava
				call Rerun(actionName,obj_WpfWindow)
         	End If         	
		End With			
	End If
End Sub 





Public Sub MRInputID_NumberOutright(actionName,obj_WpfWindow,obj_frame,strWebEdit,index,Id_Number)


	On error resume next 
	
	If Id_Number <> "" Then
		With JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame)		
   			If .JavaTable(strWebEdit).Exist(60) Then
'				wait 1
'			   JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame).JavaTable(strWebEdit).ActivateCell "#16","C"
'			   JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame).JavaEdit(index).Set Id_Number + vbLf + ""
'			   JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame).JavaEdit(index).Activate
'			   wait 1
			    JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame).JavaTable(strWebEdit).ActivateCell "#17","C"
			    Call StartTime(StrStartTime)
			    wait 2
			    JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame).JavaEdit(index).Set Id_Number + vbLf + ""
			    wait 2
			    JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame).JavaTable(strWebEdit).ActivateCell "#21","B"
			    

  				Call EndTime(StrEndTime)
			   '	If err.number = 0 Then
        			PrintInfo "micPass", strWebEdit,  " Error Description: " & Err.Description
        			StrStatus = strWebEdit & "Step is Pass"
        			Call CaptureScreenShot(strFileName)
        			Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation)
        			Err.clear						
					Exit sub
				'End If
   			Else
   			    PrintInfo "micFail", strWebEdit, strWebEdit & "WebEdit not found."   
   			    StrStatus = strWebEdit & "Step is False"
   			    Call EndTime(StrEndTime)
   			    Call CaptureScreenShot(strFileName)
				Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation) 
				'Call Killservicejava
				call Rerun(actionName,obj_WpfWindow)
         	End If         	
		End With			
	End If
End Sub 




Public Sub MRGettextOnJavaInternalframe(actionName,obj_WpfWindow,strWebEdit,strvalue)

	On error resume next 
	Call StartTime(StrStartTime)
	If obj_WpfWindow <> "" Then	
	'WpfWindow("Micro Focus MyFlight Sample").WpfObject("John Smith").GetROProperty("text")


		With WpfWindow(obj_WpfWindow)			
   			If .JavaStaticText(strWebEdit).Exist(60) Then
   				strWebEdit = JavaWindow(obj_WpfWindow).JavaInternalFrame(strWebEdit).JavaStaticText(strvalue).GetROProperty("text")
   			 
  				Call EndTime(StrEndTime)
			   '	If err.number = 0 Then
        			PrintInfo "micPass", strWebEdit,  " Error Description: " & Err.Description
        			StrStatus = strWebEdit & "Step is Pass"
        			Call CaptureScreenShot(strFileName)
        			Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation)
        			Err.clear						
					Exit sub
				'End If
   			Else
   			    PrintInfo "micFail", strWebEdit, strWebEdit & "WebEdit not found."   
   			    StrStatus = strWebEdit & "Step is False"
   			    Call EndTime(StrEndTime)
   			    Call CaptureScreenShot(strFileName)
				Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation) 	
				'Call Killservicejava
				call Rerun(actionName,obj_WpfWindow)
         	End If         	
		End With			
	End If
End Sub 




Public Sub MRGettextOnJavaDinalog(actionName,obj_WpfWindow,strWebEdit,strvalue)

	On error resume next 
	Call StartTime(StrStartTime)
	If obj_WpfWindow <> "" Then	
	'WpfWindow("Micro Focus MyFlight Sample").WpfObject("John Smith").GetROProperty("text")


		With WpfWindow(obj_WpfWindow)			
   			If .JavaDialog(strWebEdit).JavaStaticText(strWebEdit).Exist(60) Then
   				strWebEdit = JavaWindow(obj_WpfWindow).JavaDialog(strWebEdit).JavaStaticText(strvalue).GetROProperty("text")
   			   
  				Call EndTime(StrEndTime)
			   '	If err.number = 0 Then
        			PrintInfo "micPass", strWebEdit,  " Error Description: " & Err.Description
        			StrStatus = strWebEdit & "Step is Pass"
        			Call CaptureScreenShot(strFileName)
        			Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation)
        			Err.clear						
					Exit sub
				'End If
   			Else
   			    PrintInfo "micFail", strWebEdit, strWebEdit & "WebEdit not found."   
   			    StrStatus = strWebEdit & "Step is False"
   			    Call EndTime(StrEndTime)
   			    Call CaptureScreenShot(strFileName)
				Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation) 
				'Call Killservicejava
				call Rerun(actionName,obj_WpfWindow)
         	End If         	
		End With			
	End If
End Sub 




Public Function MRCheckPageMenu(actionName,obj_WpfWindow,strWebEdit,strvalue)

	Call StartTime(StrStartTime)
 	With JavaWindow(obj_WpfWindow)			
   			If .JavaInternalFrame(strWebEdit).JavaObject(strvalue).Exist(60) Then

  				Call EndTime(StrEndTime)
			   '	If err.number = 0 Then
        			PrintInfo "micPass", strWebEdit,  " Error Description: " & Err.Description
        			StrStatus = strWebEdit & "Step is Pass"
        			Call CaptureScreenShot(strFileName)
        			Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation)
        			Err.clear						
					'Exit sub
				'End If
   			Else
   			    PrintInfo "micFail", strWebEdit, strWebEdit & "WebEdit not found."   
   			    StrStatus = strWebEdit & "Step is False"
   			    Call EndTime(StrEndTime)
   			    Call CaptureScreenShot(strFileName)
				Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation) 	
				'Call Killservicejava
				call Rerun(actionName,obj_WpfWindow)
         	End If         	
		End With	
 	
End Function 



Public Sub MRActivedCell(actionName,obj_WpfWindow,obj_frame,strWebEdit,irow,icolum)


	On error resume next 
	 Call StartTime(StrStartTime)
	If strWebEdit <> "" Then
		With JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame)		
   			If .JavaTable(strWebEdit).Exist(60) Then

			   wait 1
			   JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame).JavaTable(strWebEdit).ActivateCell irow,icolum


  				Call EndTime(StrEndTime)
			   '	If err.number = 0 Then
        			PrintInfo "micPass", strWebEdit,  " Error Description: " & Err.Description
        			StrStatus = strWebEdit & "Step is Pass"
        			Call CaptureScreenShot(strFileName)
        			Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation)
        			Err.clear						
					Exit sub
				'End If
   			Else
   			    PrintInfo "micFail", strWebEdit, strWebEdit & "WebEdit not found."   
   			    StrStatus = strWebEdit & "Step is False"
   			    Call EndTime(StrEndTime)
   			    Call CaptureScreenShot(strFileName)
				Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation) 
				'Call Killservicejava
				call Rerun(actionName,obj_WpfWindow)
         	End If         	
		End With			
	End If
End Sub 


Public Sub MRDoubleClickCell(actionName,obj_WpfWindow,obj_frame,strWebEdit,irow,icolum)


	On error resume next 
	 Call StartTime(StrStartTime)
	If strWebEdit <> "" Then
		With JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame)		
   			If .JavaTable(strWebEdit).Exist(60) Then

			   wait 1
			   JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame).JavaTable(strWebEdit).DoubleClickCell irow,icolum


  				Call EndTime(StrEndTime)
			   '	If err.number = 0 Then
        			PrintInfo "micPass", strWebEdit,  " Error Description: " & Err.Description
        			StrStatus = strWebEdit & "Step is Pass"
        			Call CaptureScreenShot(strFileName)
        			Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation)
        			Err.clear						
					Exit sub
				'End If
   			Else
   			    PrintInfo "micFail", strWebEdit, strWebEdit & "WebEdit not found."   
   			    StrStatus = strWebEdit & "Step is False"
   			    Call EndTime(StrEndTime)
   			    Call CaptureScreenShot(strFileName)
				Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation) 
				'Call Killservicejava
				call Rerun(actionName,obj_WpfWindow)
         	End If         	
		End With			
	End If
End Sub 



Public Sub MRSettextOnCell(actionName,obj_WpfWindow,obj_frame,strWebEdit,Amount)


	On error resume next 
	
	If strWebEdit <> "" Then
		With JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame)		
   			If .JavaTable(strWebEdit).Exist(60) Then
   			wait 1
			   Call StartTime(StrStartTime)
			   JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame).JavaEdit(strWebEdit).set Amount + vbLf + ""

  				Call EndTime(StrEndTime)
			   '	If err.number = 0 Then
        			PrintInfo "micPass", strWebEdit,  " Error Description: " & Err.Description
        			StrStatus = strWebEdit & "Step is Pass"
        			Call CaptureScreenShot(strFileName)
        			Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation)
        			Err.clear						
					Exit sub
				'End If
   			Else
   			    PrintInfo "micFail", strWebEdit, strWebEdit & "WebEdit not found."   
   			    StrStatus = strWebEdit & "Step is False"
   			    Call EndTime(StrEndTime)
   			    Call CaptureScreenShot(strFileName)
				Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation) 	
				'Call Killservicejava
				call Rerun(actionName,obj_WpfWindow)
         	End If         	
		End With			
	End If
End Sub 



Public Sub MRGetcelltable(actionName,obj_WpfWindow,obj_frame,strWebEdit,irow,icol)


	On error resume next 
	
	If strWebEdit <> "" Then
		With JavaWindow(obj_WpfWindow).JavaDialog(obj_frame)		
   			If .JavaTable(strWebEdit).Exist(60) Then

			   Call StartTime(StrStartTime)

				aaa = JavaWindow(obj_WpfWindow).JavaDialog(obj_frame).JavaTable(strWebEdit).GetCellData (irow,icol)
				aaa = strWebEdit
				aaa = mid(aaa,34,5)
				DataTable.Value("Desctiption","Global") = aaa
  				Call EndTime(StrEndTime)
			   '	If err.number = 0 Then
        			PrintInfo "micPass", strWebEdit,  " Error Description: " & Err.Description
        			StrStatus = strWebEdit & "Step is Pass"
        			Call CaptureScreenShot(strFileName)
        			Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation)
        			Err.clear						
					Exit sub
				'End If
   			Else
   			    PrintInfo "micFail", strWebEdit, strWebEdit & "WebEdit not found."   
   			    StrStatus = strWebEdit & "Step is False"
   			    Call EndTime(StrEndTime)
   			    Call CaptureScreenShot(strFileName)
				Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation)
				'Call Killservicejava
				call Rerun(actionName,obj_WpfWindow)
         	End If         	
		End With			
	End If
End Sub 

'JavaWindow("MX.3 - MX_PERF_TEST(PROD)").JavaList("My MX").Select "Processing"
'JavaWindow("MX.3 - MX_PERF_TEST(PROD)").JavaList("NavigationScreenList$1").Select "#4"

Public Sub MRJavaList (actionName,strWebList,obj_Page, strInValue)    
	On error resume next 
	If strInValue = "Processing" Then
			strInValue = "#1"
			
	ElseIf strInValue = "Processing8" Then
			strInValue = "#8"	
			
	ElseIf strInValue = "Processing1"  Then
		   strInValue = "#5"	
	
	ElseIf strInValue = "ProcessingScript10"  Then
			strInValue = "#10"
	
	ElseIf strInValue = "ProcessingScript5"  Then
			strInValue = "#5"
			
	ElseIf strInValue = "ProcessingScript4"  Then
			strInValue = "#4"		
			
	ElseIf strInValue = "Processing_Script"  Then
			strInValue = "#3"
			
	ElseIf strInValue = "Script2"  Then
			strInValue = "#2"
			
	ElseIf strInValue = "Script4"  Then
			strInValue = "#4"
			
	ElseIf strInValue = "CIMRFXB012PS"  Then
			strInValue = "#1"	
			'JavaWindow("MX.3 - MX_PERF_TEST(PROD)").JavaList("Ctrl-F").Select "#1"
	ElseIf strInValue = "CI_MR_GENB010_PS"  Then
			strInValue = "#5"			
			
	End If 
	
	If strInValue <> "" Then
		With JavaWindow(strWebList).JavaList(obj_Page)						
			If JavaWindow(strWebList).JavaList(obj_Page).Exist(0) Then
			'MsgBox strInValue
				JavaWindow(strWebList).JavaList(obj_Page).Select strInValue
				If err.number <>0 Then
        			PrintInfo "micFail", strInValue,  " Error Description: " & Err.Description
        			Err.clear						
					Exit sub
				End If
			Else
   			   PrintInfo "micFail", strWebList,  strWebList & " - WebList not found."
			   call Rerun(actionName,obj_WpfWindow)
         	End If			
		End With		
	End If
End Sub



Public Sub MRPressKey(actionName,obj_WpfWindow,obj_frame,strWebEdit,index)

	On error resume next 
	Call StartTime(StrStartTime)
	If index <> "" Then
		With JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame)		
   			If .JavaTable(strWebEdit).Exist(60) Then
				JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame).JavaTable(strWebEdit).Click 321,132,"LEFT"
			    JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame).JavaTable(strWebEdit).PressKey index ,micCtrl


			    

  				Call EndTime(StrEndTime)
			   '	If err.number = 0 Then
        			PrintInfo "micPass", strWebEdit,  " Error Description: " & Err.Description
        			StrStatus = strWebEdit & "Step is Pass"
        			Call CaptureScreenShot(strFileName)
        			Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation)
        			Err.clear						
					Exit sub
				'End If
   			Else
   			    PrintInfo "micFail", strWebEdit, strWebEdit & "WebEdit not found."   
   			    StrStatus = strWebEdit & "Step is False"
   			    Call EndTime(StrEndTime)
   			    Call CaptureScreenShot(strFileName)
				Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation) 	
				'Call Killservicejava
				call Rerun(actionName,obj_WpfWindow)
         	End If         	
		End With			
	End If
End Sub 




Public Sub MRActivedCellRow(actionName,obj_WpfWindow,obj_frame,strWebEdit,irow)


	On error resume next 
	
	If strWebEdit <> "" Then
		With JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame)		
   			If .JavaTable(strWebEdit).Exist(60) Then
			   Call StartTime(StrStartTime)
			   wait 1
			   JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame).JavaTable(strWebEdit).ActivateRow irow
			   
			   If JavaWindow("MX.3 - MX_PERF_TEST(PROD)_MX").JavaInternalFrame("MX.3_2").JavaButton("Yes_update").Exist(2) Then
			   	JavaWindow("MX.3 - MX_PERF_TEST(PROD)_MX").JavaInternalFrame("MX.3_2").JavaButton("Yes_update").Click
			   End If
			   
  				Call EndTime(StrEndTime)
			   '	If err.number = 0 Then
        			PrintInfo "micPass", strWebEdit,  " Error Description: " & Err.Description
        			StrStatus = strWebEdit & "Step is Pass"
        			Call CaptureScreenShot(strFileName)
        			Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation)
        			Err.clear						
					Exit sub
				'End If
   			Else
   			    PrintInfo "micFail", strWebEdit, strWebEdit & "WebEdit not found."   
   			    StrStatus = strWebEdit & "Step is False"
   			    Call EndTime(StrEndTime)
   			    Call CaptureScreenShot(strFileName)
				Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation) 	
				'Call Killservicejava
				call Rerun(actionName,obj_WpfWindow)
         	End If         	
		End With			
	End If
End Sub 



Public Sub MRsyncInternalFrame (actionName,obj_WpfWindow,obj_Frame)

	On error resume next 
	Call StartTime(StrStartTime)
	If obj_WpfWindow <> "" Then	
	'WpfWindow("Micro Focus MyFlight Sample").WpfObject("John Smith").GetROProperty("text")


		With JavaWindow(obj_WpfWindow)			
   			If .JavaInternalFrame(obj_Frame).Exist(60) Then
   			   .JavaInternalFrame(obj_Frame).Activate
  				Call EndTime(StrEndTime)
			   '	If err.number = 0 Then
        			PrintInfo "micPass", strWebEdit,  " Error Description: " & Err.Description
        			StrStatus = strWebEdit & "Step is Pass"
        			Call CaptureScreenShot(strFileName)
        			Call Getparamiter(obj_Frame,StrStatus,Myfile,StrStartTime,StrEndTime,Transation)
        			Err.clear						
					Exit sub
				'End If
   			Else
   			    PrintInfo "micFail", strWebEdit, strWebEdit & "WebEdit not found."   
   			    StrStatus = strWebEdit & "Step is False"
   			    Call EndTime(StrEndTime)
   			    Call CaptureScreenShot(strFileName)
				Call Getparamiter(obj_Frame,StrStatus,Myfile,StrStartTime,StrEndTime,Transation) 	
				'Call Killservicejava
				call Rerun(actionName,obj_WpfWindow)
         	End If         	
		End With			
	End If
End Sub 






Public Sub MRCheckButtonJavaList(actionName,obj_WpfWindow,strWebEdit)

	On error resume next 
	Call StartTime(StrStartTime)
	If obj_WpfWindow <> "" Then	
	'JavaWindow("MX.3 - MX_PERF_TEST(PROD)").JavaList("My MX").Select


		With JavaWindow(obj_WpfWindow)			
   			If .JavaList(strWebEdit).Exist(60) Then
   				strWebEdit = JavaWindow(obj_WpfWindow).JavaList(strWebEdit)
   			   
  				Call EndTime(StrEndTime)
			   '	If err.number = 0 Then
        			PrintInfo "micPass", strWebEdit,  " Error Description: " & Err.Description
        			StrStatus = strWebEdit & "Step is Pass"
        			Call CaptureScreenShot(strFileName)
        			Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation)
        			Err.clear						
					Exit sub
				'End If
   			Else
   			    PrintInfo "micFail", strWebEdit, strWebEdit & "WebEdit not found."   
   			    StrStatus = strWebEdit & "Step is False"
   			    Call EndTime(StrEndTime)
   			    Call CaptureScreenShot(strFileName)
				Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation) 
				'Call Killservicejava	
				call Rerun(actionName,obj_WpfWindow)
         	End If         	
		End With			
	End If
End Sub 

Function MRPerformRightClick(x1,y1)
		Set ODR = CreateObject("Mercury.DeviceReplay")
		ODR.MouseClick x1,y1,RIGHT_MOUSE_BUTTON
End Function

Function MSendKeys(key)
	Set Obj = CreateObject("Wscript.Shell")
	Obj.SendKeys key,1
End Function


Public Sub MRPerformRight(actionName,obj_WpfWindow,strWebEdit)

	On error resume next 
	Call StartTime(StrStartTime)
	If obj_WpfWindow <> "" Then	
	'JavaWindow("MX.3 - MX_PERF_TEST(PROD)").JavaList("My MX").Select


		With JavaWindow(obj_WpfWindow)			
   			If .JavaList(strWebEdit).Exist(60) Then
   				strWebEdit = JavaWindow(obj_WpfWindow).JavaList(strWebEdit).MouseClick 
   			   
  				Call EndTime(StrEndTime)
			   '	If err.number = 0 Then
        			PrintInfo "micPass", strWebEdit,  " Error Description: " & Err.Description
        			StrStatus = strWebEdit & "Step is Pass"
        			Call CaptureScreenShot(strFileName)
        			Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation)
        			Err.clear						
					Exit sub
				'End If
   			Else
   			    PrintInfo "micFail", strWebEdit, strWebEdit & "WebEdit not found."   
   			    StrStatus = strWebEdit & "Step is False"
   			    Call EndTime(StrEndTime)
   			    Call CaptureScreenShot(strFileName)
				Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation) 		
				'Call Killservicejava
				call Rerun(actionName,obj_WpfWindow)
         	End If         	
		End With			
	End If
End Sub 




Public Sub MRAppClickMenuText(actionName,obj_WpfWindow,strWebEdit)
	On error resume next 
	Call StartTime(StrStartTime)
	'JavaWindow("MX.3 - MX_PERF_TEST(PROD)").JavaStaticText("MX.3").Click
	
	If strWebEdit <> "" Then
		With JavaWindow(obj_WpfWindow)			
   			If .JavaStaticText(strWebEdit).Exist(60) Then
				JavaWindow(obj_WpfWindow).JavaStaticText(strWebEdit).Click '8,18,"LEFT"

        			Call EndTime(StrEndTime)
			   '	If err.number = 0 Then
        			PrintInfo "micPass", strWebEdit,  " Error Description: " & Err.Description
        			StrStatus = strWebEdit & "Step is Pass"
        			Call CaptureScreenShot(strFileName)
        			Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation)
        			Err.clear						
					Exit sub
				'End If
   			Else
   			    PrintInfo "micFail", "WebButton Click", strWebEdit & " WebButton not found." 
			    Call EndTime(StrEndTime)
   			    Call CaptureScreenShot(strFileName)
				Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation)   
				'Call Killservicejava
				call Rerun(actionName,obj_WpfWindow)
         	End If         	
		End With			
	End If
End Sub



Public Sub MRCheckpointButton(actionName,obj_WpfWindow,obj_frame,obj_Bt,strWebEdit)


	On error resume next 
	
	If strWebEdit <> "" Then
		With JavaWindow(obj_WpfWindow).JavaDialog(obj_frame)		
   			If .JavaTable(strWebEdit).Exist(60) Then

			   Call StartTime(StrStartTime)
				Vlapp = JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame).JavaButton(obj_Bt).GetROProperty("label")
				If Vlapp = strWebEdit Then
					Call EndTime(StrEndTime)
			   '	If err.number = 0 Then
        			PrintInfo "micPass", strWebEdit,  " Error Description: " & Err.Description
        			StrStatus = strWebEdit & "Step is Pass"
        			Call CaptureScreenShot(strFileName)
        			Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation)
        			Err.clear						
					Exit sub
				'End If
					Else
						PrintInfo "micFail", strWebEdit, strWebEdit & "WebEdit not found."   
		   			    StrStatus = strWebEdit & "Step is False"
		   			    Call EndTime(StrEndTime)
		   			    Call CaptureScreenShot(strFileName)
						Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation) 	
				End If
				
				
   			Else
   			    PrintInfo "micFail", strWebEdit, strWebEdit & "WebEdit not found."   
   			    StrStatus = strWebEdit & "Step is False"
   			    Call EndTime(StrEndTime)
   			    Call CaptureScreenShot(strFileName)
				Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation) 	
				'Call Killservicejava
				call Rerun(actionName,obj_WpfWindow)
         	End If         	
		End With			
	End If
End Sub 


Public Sub MRCheckpointData(actionName,obj_WpfWindow,obj_frame,strWebEdit)

	On error resume next 

	If strWebEdit <> "" Then
		With JavaWindow(obj_WpfWindow).JavaDialog(obj_frame)		
   			If .JavaTable(strWebEdit).Exist(60) Then
			   Call StartTime(StrStartTime)
				strWebEdit = JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame).JavaTable(strWebEdit).Object.getRowCount

				If strWebEdit <> "" Then
					Call EndTime(StrEndTime)
			   '	If err.number = 0 Then
        			PrintInfo "micPass", strWebEdit,  " Error Description: " & Err.Description
        			StrStatus = strWebEdit & "Step is Pass"
        			Call CaptureScreenShot(strFileName)
        			Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation)
        			Err.clear						
					Exit sub
				'End If
					Else
						PrintInfo "micFail", strWebEdit, strWebEdit & "WebEdit not found."   
		   			    StrStatus = strWebEdit & "Step is False"
		   			    Call EndTime(StrEndTime)
		   			    Call CaptureScreenShot(strFileName)
						Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation) 
						call Rerun(actionName,obj_WpfWindow)
				End If
				
				
   			Else
   			    PrintInfo "micFail", strWebEdit, strWebEdit & "WebEdit not found."   
   			    StrStatus = strWebEdit & "Step is False"
   			    Call EndTime(StrEndTime)
   			    Call CaptureScreenShot(strFileName)
				Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation) 	
				'Call Killservicejava
				call Rerun(actionName,obj_WpfWindow)
         	End If         	
		End With			
	End If
End Sub 


Public Sub MRPressKeyObj(actionName,obj_WpfWindow,obj_frame,strWebEdit,index)

	On error resume next 
	Call StartTime(StrStartTime)
	If index <> "" Then
		With JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame)		
   			If .JavaTable(strWebEdit).Exist(60) Then
				JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame).JavaTable(strWebEdit).Click 321,132,"LEFT"
			    JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame).JavaObject(strWebEdit).PressKey index ,micCtrl


			    

  				Call EndTime(StrEndTime)
			   '	If err.number = 0 Then
        			PrintInfo "micPass", strWebEdit,  " Error Description: " & Err.Description
        			StrStatus = strWebEdit & "Step is Pass"
        			Call CaptureScreenShot(strFileName)
        			Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation)
        			Err.clear						
					Exit sub
				'End If
   			Else
   			    PrintInfo "micFail", strWebEdit, strWebEdit & "WebEdit not found."   
   			    StrStatus = strWebEdit & "Step is False"
   			    Call EndTime(StrEndTime)
   			    Call CaptureScreenShot(strFileName)
				Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation) 	
				'Call Killservicejava	
				call Rerun(actionName,obj_WpfWindow)
         	End If         	
		End With			
	End If
End Sub 


Public Sub MuselectSending(actionName,obj_WpfWindow,obj_frame,strWebEdit)

If JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame).JavaTree(strWebEdit).Exist(60) Then
	JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame).JavaTree(strWebEdit).Select "#0;#2;#2"
Else
	call Rerun(actionName,obj_WpfWindow)
End If

End Sub
Public Sub MuselectSendingMoniter(actionName,obj_WpfWindow,obj_frame,strWebEdit)

If JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame).JavaTree(strWebEdit).Exist(60) Then
	JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame).JavaTree(strWebEdit).Select "#0;#2;#1"
Else
    call Rerun(actionName,obj_WpfWindow)
End If

End Sub




Public Sub MUloopchekPackedNumberSending(actionName,PackedNumber)
If  JavaWindow("MX.3 - MX_PERF_TEST(PROD)_MX").JavaInternalFrame("MModalInternalFrame").JavaTable("Validate Confirmations").Exist(60) Then
	a = JavaWindow("MX.3 - MX_PERF_TEST(PROD)_MX").JavaInternalFrame("MModalInternalFrame").JavaTable("Validate Confirmations").GetROProperty("rows")
	b = JavaWindow("MX.3 - MX_PERF_TEST(PROD)_MX").JavaInternalFrame("MModalInternalFrame").JavaTable("Validate Confirmations").GetROProperty("cols")
		For i = 0 to a
		C = JavaWindow("MX.3 - MX_PERF_TEST(PROD)_MX").JavaInternalFrame("MModalInternalFrame").JavaTable("Validate Confirmations").GetCellData (i,8)
			If C = PackedNumber Then
			k = "#"
			g = k & i
			JavaWindow("MX.3 - MX_PERF_TEST(PROD)_MX").JavaInternalFrame("MModalInternalFrame").JavaTable("Document ID").SetCellData g ,"Selection","1"
			Exit for
			else
			
			End If
		
		Next
Else
	call Rerun(actionName,obj_WpfWindow)
End If 

End Sub 



Public Sub MRCheckInternalFramePage(actionName,obj_WpfWindow,strWebEdit,strvalue)



	On error resume next 
	Call StartTime(StrStartTime)
	If obj_WpfWindow <> "" Then	
	'WpfWindow("Micro Focus MyFlight Sample").WpfObject("John Smith").GetROProperty("text")


		With WpfWindow(obj_WpfWindow)			
   			If .JavaInternalFrame(strWebEdit).JavaObject(strvalue).Exist(60) Then		   
  				Call EndTime(StrEndTime)
			   '	If err.number = 0 Then
        			PrintInfo "micPass", strWebEdit,  " Error Description: " & Err.Description
        			StrStatus = strWebEdit & "Step is Pass"
        			Call CaptureScreenShot(strFileName)
        			Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation)
        			Err.clear						
					Exit sub
				'End If
   			Else
   			    PrintInfo "micFail", strWebEdit, strWebEdit & "WebEdit not found."   
   			    StrStatus = strWebEdit & "Step is False"
   			    Call EndTime(StrEndTime)
   			    Call CaptureScreenShot(strFileName)
				Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation) 	
				'Call Killservicejava	
				call Rerun(actionName,obj_WpfWindow)
         	End If         	
		End With			
	End If
End Sub 

Public Function MRCheckPageTable(actionName,obj_WpfWindow,strWebEdit,strvalue)

	Call StartTime(StrStartTime)
 	With JavaWindow(obj_WpfWindow)			
   			If .JavaInternalFrame(strWebEdit).JavaTable(strvalue).Exist(60) Then

  				Call EndTime(StrEndTime)
			   '	If err.number = 0 Then
        			PrintInfo "micPass", strWebEdit,  " Error Description: " & Err.Description
        			StrStatus = strWebEdit & "Step is Pass"
        			Call CaptureScreenShot(strFileName)
        			Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation)
        			Err.clear						
					'Exit sub
				'End If
   			Else
   			    PrintInfo "micFail", strWebEdit, strWebEdit & "WebEdit not found."   
   			    StrStatus = strWebEdit & "Step is False"
   			    Call EndTime(StrEndTime)
   			    Call CaptureScreenShot(strFileName)
				Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation) 	
				'Call Killservicejava
				call Rerun(actionName,obj_WpfWindow)
         	End If         	
		End With	
 	
End Function 






Public Sub MRCheckInternalFrameText(actionName,obj_WpfWindow,strWebEdit,strvalue)



	On error resume next 
	Call StartTime(StrStartTime)
	If obj_WpfWindow <> "" Then	
	'JavaWindow("MX.3 - MX_PERF_TEST(PROD)_MX").JavaInternalFrame("Check list").JavaStaticText("Processing script checking(st)").Click


		With JavaWindow(obj_WpfWindow)			
   			If .JavaInternalFrame(strWebEdit).JavaStaticText(strvalue).Exist(5000) Then		   
  				Call EndTime(StrEndTime)
			   '	If err.number = 0 Then
        			PrintInfo "micPass", strWebEdit,  " Error Description: " & Err.Description
        			StrStatus = strWebEdit & "Step is Pass"
        			Call CaptureScreenShot(strFileName)
        			Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation)
        			Err.clear						
					Exit sub
				'End If
   			Else
   			    PrintInfo "micFail", strWebEdit, strWebEdit & "WebEdit not found."   
   			    StrStatus = strWebEdit & "Step is False"
   			    Call EndTime(StrEndTime)
   			    Call CaptureScreenShot(strFileName)
				Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation) 
				'Call Killservicejava	
				call Rerun(actionName,obj_WpfWindow)
         	End If         	
		End With			
	End If
End Sub 



Public Sub MRClosesAppMurex(actionName,strWebEdit)



	On error resume next 
	Call StartTime(StrStartTime)
	If strWebEdit <> "" Then
		For i = 1 To 5
			If JavaWindow("MX.3 - MX_PERF_TEST(PROD)").JavaDialog("MX.3").JavaButton("PopupButtonYes").Exist(1) Then
			
			JavaWindow("MX.3 - MX_PERF_TEST(PROD)").JavaDialog("MX.3").JavaButton("PopupButtonYes").Click
			elseif	JavaWindow("MX.3 - MX_PERF_TEST(PROD)_MX").JavaButton("closeWhite").Exist(1) then
			JavaWindow("MX.3 - MX_PERF_TEST(PROD)_MX").JavaButton("closeWhite").Click
			else
				Exit for
		End If
		Next	
  				Call EndTime(StrEndTime)
			   '	If err.number = 0 Then
        			PrintInfo "micPass", strWebEdit,  " Error Description: " & Err.Description
        			StrStatus = strWebEdit & "Step is Pass"
        			Call CaptureScreenShot(strFileName)
        			Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation)
        			Err.clear						
					Exit sub
				'End If
   			Else
   			    PrintInfo "micFail", strWebEdit, strWebEdit & "WebEdit not found."   
   			    StrStatus = strWebEdit & "Step is False"
   			    Call EndTime(StrEndTime)
   			    Call CaptureScreenShot(strFileName)
				Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation) 	
				'Call Killservicejava
				call Rerun(actionName,obj_WpfWindow)
         	End If         				

End Sub


Public Sub MuLoopDupicateforleg(strWebEdit)
For i = 1 To strWebEdit
wait 1
JavaWindow("MX.3 - MX_PERF_TEST(PROD)_MX").JavaInternalFrame("e-Tradepad").JavaTable("5 legs").Click 103,26,"RIGHT"
wait 1
JavaWindow("MX.3 - MX_PERF_TEST(PROD)_MX3").JavaInternalFrame("e-Tradepad").JavaMenu("Duplicate leg").Select

 Next
End Sub





Public Sub MRCheckJavaDialog(actionName,obj_WpfWindow,strWebEdit,strvalue)


	On error resume next 
	Call StartTime(StrStartTime)
	If obj_WpfWindow <> "" Then	
	'JavaWindow("MX.3 - MX_PERF_TEST(PROD)_MX").JavaDialog("MX.3").JavaButton("Yes").Click

		With JavaWindow(obj_WpfWindow)			
   			If .JavaDialog(strWebEdit).JavaButton(strvalue).Exist(60) Then		   
  				Call EndTime(StrEndTime)
			   '	If err.number = 0 Then
        			PrintInfo "micPass", strWebEdit,  " Error Description: " & Err.Description
        			StrStatus = strWebEdit & "Step is Pass"
        			Call CaptureScreenShot(strFileName)
        			Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation)
        			Err.clear						
					Exit sub
				'End If
   			Else
   			    PrintInfo "micFail", strWebEdit, strWebEdit & "WebEdit not found."   
   			    StrStatus = strWebEdit & "Step is False"
   			    Call EndTime(StrEndTime)
   			    Call CaptureScreenShot(strFileName)
				Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation) 	
			'	'Call Killservicejava	
				call Rerun(actionName,obj_WpfWindow)
         	End If         	
		End With			
	End If
End Sub 




Public Sub MRCheckJavaEdit(actionName,obj_WpfWindow,strWebEdit)

	On error resume next 
	Call StartTime(StrStartTime)
	If obj_WpfWindow <> "" Then	
	'JavaWindow("MX.3 - MX_PERF_TEST(PROD)").JavaEdit("Ctrl-F").Set

		With JavaWindow(obj_WpfWindow)			
   			If .JavaEdit(strWebEdit).Exist(60) Then
   				strWebEdit = JavaWindow(obj_WpfWindow).JavaEdit(strWebEdit)
   			   
  				Call EndTime(StrEndTime)
			   '	If err.number = 0 Then
        			PrintInfo "micPass", strWebEdit,  " Error Description: " & Err.Description
        			StrStatus = strWebEdit & "Step is Pass"
        			Call CaptureScreenShot(strFileName)
        			Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation)
        			Err.clear						
					Exit sub
				'End If
   			Else
   			    PrintInfo "micFail", strWebEdit, strWebEdit & "WebEdit not found."   
   			    StrStatus = strWebEdit & "Step is False"
   			    Call EndTime(StrEndTime)
   			    Call CaptureScreenShot(strFileName)
				Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation) 
				'Call Killservicejava
				call Rerun(actionName,obj_WpfWindow)
         	End If         	
		End With			
	End If
End Sub 




Public Sub MRSelect_Activate(actionName,obj_WpfWindow,obj_frame,strWebEdit,strInValue)


	On error resume next 
	Call StartTime(StrStartTime)
	If strInValue <> "" Then
	'WpfWindow("Micro Focus MyFlight Sample").WpfEdit("agentName").Set
			Wait 3
		With JavaWindow(obj_WpfWindow)			
   			If .JavaInternalFrame(obj_frame).JavaTree(strWebEdit).Exist(60) Then
   			   .JavaInternalFrame(obj_frame).JavaTree(strWebEdit).highlight
   			   .JavaInternalFrame(obj_frame).JavaTree(strWebEdit).Activate strInValue   	   
  				Call EndTime(StrEndTime)
			   '	If err.number = 0 Then
        			PrintInfo "micPass", strWebEdit,  " Error Description: " & Err.Description
        			StrStatus = strWebEdit & "Step is Pass"
        			Call CaptureScreenShot(strFileName)
        			Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation)
        			Err.clear						
					Exit sub
				'End If
   			Else
   			    PrintInfo "micFail", strWebEdit, strWebEdit & "WebEdit not found."   
   			    StrStatus = strWebEdit & "Step is False"
   			    Call EndTime(StrEndTime)
   			    Call CaptureScreenShot(strFileName)
				Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation) 	
				'Call Killservicejava
				call Rerun(actionName,obj_WpfWindow)
         	End If         	
		End With			
	End If
End Sub 


Public Sub MRSelect(actionName,obj_WpfWindow,obj_frame,strWebEdit,strInValue)


	On error resume next 
	Call StartTime(StrStartTime)
	If strInValue <> "" Then
	'WpfWindow("Micro Focus MyFlight Sample").WpfEdit("agentName").Set
	
		With JavaWindow(obj_WpfWindow)			
   			If .JavaInternalFrame(obj_frame).JavaTree(strWebEdit).Exist(60) Then
   			   .JavaInternalFrame(obj_frame).JavaTree(strWebEdit).highlight
   			   .JavaInternalFrame(obj_frame).JavaTree(strWebEdit).Select strInValue
   			   '.JavaInternalFrame(obj_frame).JavaTree(strWebEdit).Activate strInValue   	   
  				Call EndTime(StrEndTime)
			   '	If err.number = 0 Then
        			PrintInfo "micPass", strWebEdit,  " Error Description: " & Err.Description
        			StrStatus = strWebEdit & "Step is Pass"
        			Call CaptureScreenShot(strFileName)
        			Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation)
        			Err.clear						
					Exit sub
				'End If
   			Else
   			    PrintInfo "micFail", strWebEdit, strWebEdit & "WebEdit not found."   
   			    StrStatus = strWebEdit & "Step is False"
   			    Call EndTime(StrEndTime)
   			    Call CaptureScreenShot(strFileName)
				Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation) 	
				'Call Killservicejava
				call Rerun(actionName,obj_WpfWindow)
         	End If         	
		End With			
	End If
End Sub 






Public Sub MRSearchInternalFrameEdit(actionName,obj_WpfWindow,obj_InternalFrame,strWebEdit,strInValue)

	On error resume next 
	Call StartTime(StrStartTime)
	If strInValue <> "" Then
	'JavaWindow("MX.3 - MX_PERF_TEST(PROD)_MX").JavaInternalFrame("e-Tradepad").JavaEdit("GuiCellComponentTextFieldStrin").Set
	
		With JavaWindow(obj_WpfWindow)			
   			If .JavaInternalFrame(obj_InternalFrame).Exist(60) Then
   			   .JavaEdit(strWebEdit).highlight  			   
   			   .JavaEdit(strWebEdit).Set strInValue
				Call Sendkey("{ENTER}")			   
  				Call EndTime(StrEndTime)
			   '	If err.number = 0 Then
        			PrintInfo "micPass", strWebEdit,  " Error Description: " & Err.Description
        			StrStatus = strWebEdit & "Step is Pass"
        			Call CaptureScreenShot(strFileName)
        			Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation)
        			Err.clear						
					Exit sub
				'End If
   			Else
   			    PrintInfo "micFail", strWebEdit, strWebEdit & "WebEdit not found."   
   			    StrStatus = strWebEdit & "Step is False"
   			    Call EndTime(StrEndTime)
   			    Call CaptureScreenShot(strFileName)
				Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation) 			
				'Call Killservicejava
				call Rerun(actionName,obj_WpfWindow)
         	End If         	
		End With			
	End If
End Sub 




Public Sub MRPressKeyObjtxt(actionName,obj_WpfWindow,obj_frame,strWebEdit,index)
'JavaWindow("MX.3 - MX_PERF_TEST(PROD)_MX").JavaInternalFrame("e-Tradepad").JavaStaticText("Pricing:(st)_2").PressKey "O",micCtrl

	On error resume next 
	Call StartTime(StrStartTime)
	If index <> "" Then
		With JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame)		
   			If .JavaTable(strWebEdit).Exist(60) Then
				JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame).JavaStaticText(strWebEdit).Click 321,132,"LEFT"
			    JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame).JavaStaticText(strWebEdit).PressKey index ,micCtrl		    

  				Call EndTime(StrEndTime)
			   '	If err.number = 0 Then
        			PrintInfo "micPass", strWebEdit,  " Error Description: " & Err.Description
        			StrStatus = strWebEdit & "Step is Pass"
        			Call CaptureScreenShot(strFileName)
        			Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation)
        			Err.clear						
					Exit sub
				'End If
   			Else
   			    PrintInfo "micFail", strWebEdit, strWebEdit & "WebEdit not found."   
   			    StrStatus = strWebEdit & "Step is False"
   			    Call EndTime(StrEndTime)
   			    Call CaptureScreenShot(strFileName)
				Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation) 
				'Call Killservicejava
				call Rerun(actionName,obj_WpfWindow)
         	End If         	
		End With			
	End If
End Sub 



Public Function General_CloseProcess(strProgramName)
   Dim objshell
   Set objshell=CreateObject("WScript.Shell")
   objshell.Run "TASKKILL /F /IM "& strProgramName
   Set objshell=nothing
End Function






Public Sub MRChackJavaInternalFrame(actionName,obj_WpfWindow,strWebEdit)

	On error resume next 
	Call StartTime(StrStartTime)
	If obj_WpfWindow <> "" Then	
	'JavaWindow("MX.3 - MX_PERF_TEST(PROD)_MX").JavaInternalFrame("MModalInternalFrame_3").Activate
		With JavaWindow(obj_WpfWindow)			
   			If .JavaInternalFrame(strWebEdit).Exist(60) Then
   			   
  				Call EndTime(StrEndTime)
			   '	If err.number = 0 Then
        			PrintInfo "micPass", strWebEdit,  " Error Description: " & Err.Description
        			StrStatus = strWebEdit & "Step is Pass"
        			Call CaptureScreenShot(strFileName)
        			Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation)
        			Err.clear						
					Exit sub
				'End If
   			Else
   			    PrintInfo "micFail", strWebEdit, strWebEdit & "WebEdit not found."   
   			    StrStatus = strWebEdit & "Step is False"
   			    Call EndTime(StrEndTime)
   			    Call CaptureScreenShot(strFileName)
				Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation) 	
				'Call Killservicejava
				call Rerun(actionName,obj_WpfWindow)
         	End If         	
		End With			
	End If
End Sub 




Public Sub MuSelectSettlementFlows(actionName,obj_WpfWindow,obj_frame,strWebEdit,strvalue)
If JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame).JavaTree(strWebEdit).Exist(60) Then
	JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame).JavaTree(strWebEdit).Select strvalue
Else 
	call Rerun(actionName,obj_WpfWindow)
End If
End Sub


Public Sub CheckTablefromMCL(obj_WpfWindow,obj_frame,strWebEdit,strvalue)


aa = JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame).JavaTree(strWebEdit).GetROProperty("items count")
'MsgBox aa
For i = 0 To aa
	bb = JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame).JavaTree(strWebEdit).GetItem(i)
	'MsgBox bb
	print InStr(bb,strvalue)
If InStr(bb,strvalue) <> 0 Then
    Print "Pass"
    Exit For
End If
Next
End Sub

Public Sub MRAppButton_Click2(obj_WpfWindow,strWebEdit)
		'MRAppButton_Click(actionName,obj_WpfWindow,strWebEdit)
	On error resume next 
	

	Call StartTime(StrStartTime)
	'WpfWindow("Micro Focus MyFlight Sample").WpfButton("OK").Click

	If strWebEdit <> "" Then
		With JavaWindow(obj_WpfWindow)			
   			If .JavaButton(strWebEdit).Exist(60) Then
   			   .JavaButton(strWebEdit).click   		
        			'Call EndTime(StrEndTime)
			   '	If err.number = 0 Then
        			PrintInfo "micPass", strWebEdit,  " Error Description: " & Err.Description
        			'StrStatus = strWebEdit & "Step is Pass"
        			'Call CaptureScreenShot(strFileName)
        			'Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation)
        			'Err.clear						
					'Exit sub
				'End If
   			Else
   			    PrintInfo "micFail", "WebButton Click", strWebEdit & " WebButton not found." 
			    'Call EndTime(StrEndTime)
   			    'Call CaptureScreenShot(strFileName)
				'Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation)    	
				'Call Killservicejava
         	End If         	
		End With			
	End If
End Sub


Public Sub MRAppButton_Click_Dinalog2(obj_WpfWindow,obj_frame,strWebEdit)
	On error resume next 
	Call StartTime(StrStartTime)
	'WpfWindow("Micro Focus MyFlight Sample").WpfButton("OK").Click

	If strWebEdit <> "" Then
		With JavaWindow(obj_WpfWindow)			
   			If .JavaDialog(obj_frame).Exist(60) Then
   			   '.JavaDialog(obj_frame).JavaButton(strWebEdit).highlight
   			   Call CaptureScreenShot(strFileName)
   			   .JavaDialog(obj_frame).JavaButton(strWebEdit).click   		
        			'Call EndTime(StrEndTime)
			   '	If err.number = 0 Then
        			PrintInfo "micPass", strWebEdit,  " Error Description: " & Err.Description
        			'StrStatus = strWebEdit & "Step is Pass"
        			'Call CaptureScreenShot(strFileName)
        			'Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation)
        			'Err.clear						
					'Exit sub
				'End If
   			Else
   			    PrintInfo "micFail", "WebButton Click", strWebEdit & " WebButton not found." 
			   ' Call EndTime(StrEndTime)
   			    'Call CaptureScreenShot(strFileName)
				'Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation)  
				'Call Killservicejava
				'call Rerun(actionName,obj_WpfWindow)
         	End If         	
		End With			
	End If
End Sub


Public Sub MuLoopMatcher(obj_WpfWindow,obj_frame,strWebEdit,PackedNumber)

For j = 0 To 10000


a = JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame).JavaTable(strWebEdit).GetROProperty("rows")
	'b = JavaWindow("MX.3 - MX_PERF_TEST(PROD)_MX").JavaInternalFrame("MModalInternalFrame").JavaTable("Validate Confirmations").GetROProperty("cols")
'MsgBox a
	For i = 0 To a - 1
	
		C = JavaWindow(obj_WpfWindow).JavaInternalFrame(obj_frame).JavaTable(strWebEdit).GetCellData (i,4)
		C = replace(C," ","")
		PackedNumber = replace(PackedNumber," ","")
		If C = PackedNumber Then	
		k = "Pass"		
		Exit for
		End If

	Next
	If k = "Pass" Then
		Exit for
	End If
Next	
End Sub



Public Sub MRAppText_Click_PP(actionName,obj_WpfWindow,strWebEdit)  
	On error resume next 
	Call StartTime(StrStartTime)
	'WpfWindow("Micro Focus MyFlight Sample").WpfButton("OK").Click

	If strWebEdit <> "" Then
		With JavaWindow(obj_WpfWindow)
 

   			If .JavaStaticText(strWebEdit).Exist(60) Then
   					If strWebEdit = "FO_TH-C" Then
   					
						JavaWindow("MX.3 - MX_PERF_TEST(PROD)").JavaStaticText(strWebEdit).Click 35,36,"LEFT"
	
					ElseIf strWebEdit = "FO_MNCS"	Then
	
						JavaWindow("MX.3 - MX_PERF_TEST(PROD)").JavaStaticText(strWebEdit).Click 32,33,"LEFT"
						
					ElseIf strWebEdit = "BO_OPS"	Then
					
					JavaWindow("MX.3 - MX_PERF_TEST(PROD)").JavaStaticText("BO_OPS(st)").Click 35,34,"LEFT"
					JavaWindow("MX.3 - MX_PERF_TEST(PROD)").JavaStaticText("BO_OPS").Click 35,34,"LEFT"
					ElseIf strWebEdit = "FO_BLMD" Then
					
						JavaWindow("MX.3 - MX_PERF_TEST(PROD)").JavaStaticText("FO_BLMD(st)").Click 54,35,"LEFT"
						JavaWindow("MX.3 - MX_PERF_TEST(PROD)").JavaStaticText("FO_BLMD").Click 54,35,"LEFT"
						
					ElseIF strWebEdit = "BO_STL" Then
					
							JavaWindow("MX.3 - MX_PERF_TEST(PROD)").JavaStaticText("BO_STL").Click 48,32,"LEFT"

					ElseIF strWebEdit = "MLC_CONFIG" Then
										
							'JavaWindow("MX.3 - MX_PERF_TEST(PROD)").JavaStaticText("MLC_CONFIG(st)")c
							JavaWindow("MX.3 - MX_PERF_TEST(PROD)").JavaStaticText("MLC_CONFIG").Click 48,32,"LEFT"

					
					ElseIf strWebEdit = "MX3" Then

						JavaWindow("MX.3 - MX_PERF_TEST(PROD)").JavaStaticText("MX.3").Click	
										
					ElseIf strWebEdit = "ALL(st)" Then
						
						JavaInternalFrame("Undefined - Main").JavaButton("pickUpValues").Click
						
					ElseIF strWebEdit = "FO_TH-C" Then	
					
						'JavaWindow("MX.3 - MX_PERF_TEST(PROD)").JavaStaticText("FO_TH-C").Click
						JavaWindow("MX.3 - MX_PERF_TEST(PROD)").JavaStaticText("FO_TH-C").Click 38,8,"LEFT"
						
					ElseIF strWebEdit = "BOCTRL_MK" Then		

						'JavaWindow("MX.3 - MX_PERF_TEST(PROD)").JavaStaticText("BOCTRL_MK").Click
						JavaWindow("MX.3 - MX_PERF_TEST(PROD)").JavaStaticText("BOCTRL_MK").Click 24,29,"LEFT"
					
					ElseIf strWebEdit = "MO" Then
		
						JavaWindow("MX.3 - MX_PERF_TEST(PROD)").JavaStaticText("MO(st)").Click 45,38,"LEFT"
						
					ElseIf strWebEdit = "FO_TRDD_FX" Then
					
							JavaWindow("MX.3 - MX_PERF_TEST(PROD)").JavaStaticText("FO_TRDD_FX(st)").Click 30,32,"LEFT"


							'JavaWindow("MX.3 - MX_PERF_TEST(PROD)").JavaStaticText("MO(st)").Click 45,38,"LEFT"
					ElseIf strWebEdit = "FO_TRDD_BD" Then
					
							JavaWindow("MX.3 - MX_PERF_TEST(PROD)").JavaStaticText("FO_TRDD_BD(st)").Click 66,20,"LEFT"		
					ElseIf strWebEdit = "FO_FINS" Then
							JavaWindow("MX.3 - MX_PERF_TEST(PROD)").JavaStaticText("FO_FINS(st)").Click 66,20,"LEFT"	


				End If
   			   '.JavaStaticText(strWebEdit).click   		
        			Call EndTime(StrEndTime)
			   '	If err.number = 0 Then
        			PrintInfo "micPass", strWebEdit,  " Error Description: " & Err.Description
        			StrStatus = strWebEdit & "Step is Pass"
        			Call CaptureScreenShot(strFileName)
        			Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation)
        			Err.clear						
					Exit sub
				'End If
   			Else
   			    PrintInfo "micFail", "WebButton Click", strWebEdit & " WebButton not found." 
			    Call EndTime(StrEndTime)
   			    Call CaptureScreenShot(strFileName)
				Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation)   
				'Call Killservicejava	
				call Rerun(actionName,obj_WpfWindow)
         	End If         	
		End With			
	End If
End Sub



Public Sub MR_OpenMurex(strWebEdit)

Call Killservicejava()
Set objShell = CreateObject("Wscript.Shell")

'Wscript.Echo objShell.CurrentDirectory
'
objShell.CurrentDirectory = "D:\Murex\NEW_MX_PERF"

'Wscript.Echo objShell.CurrentDirectory
objShell.run strWebEdit
'WshShell.AppActivate strWebEdit
'WScript.Sleep 100 

End Sub





Public Sub MRCheckJavaObject(actionName,obj_WpfWindow,strWebEdit,strvalue)



	On error resume next 
	Call StartTime(StrStartTime)
	If obj_WpfWindow <> "" Then	
	'WpfWindow("Micro Focus MyFlight Sample").WpfObject("John Smith").GetROProperty("text")


		With WpfWindow(obj_WpfWindow)			
   			If .JavaInternalFrame(strWebEdit).JavaObject(strvalue).Exist(60) Then		   
  				Call EndTime(StrEndTime)
			   '	If err.number = 0 Then
        			PrintInfo "micPass", strWebEdit,  " Error Description: " & Err.Description
        			StrStatus = strWebEdit & "Step is Pass"
        			Call CaptureScreenShot(strFileName)
        			Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation)
        			Err.clear						
					Exit sub
				'End If
   			Else
   			    PrintInfo "micFail", strWebEdit, strWebEdit & "WebEdit not found."   
   			    StrStatus = strWebEdit & "Step is False"
   			    Call EndTime(StrEndTime)
   			    Call CaptureScreenShot(strFileName)
				Call Getparamiter(strWebEdit,StrStatus,Myfile,StrStartTime,StrEndTime,Transation)
				call Rerun(actionName,obj_WpfWindow)
         	End If         	
		End With			
	End If
End Sub 


