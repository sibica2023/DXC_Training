'##################################################################################################################################
'Project Name : AOS 
'File Name: Create and display order
'Description:This end to end scenario is used to create and display orders in AOS application
'Developed  by/Date: Sibi C A/ 10-May-2023
'Version No:0.1
'Data File Name: Excel Sheet
'Mandatory Fields: Refer excel sheet
'Input Parameters Used:  Excel Sheet
'Output Parameters Used: Refer excel sheet
'Reviewed by/Review Date: 
'*******************************************************************************Modification history***********************************************************************************
'S.No___________________________Modified by__________________________Modified Date__________________________Reason____________________

'****************************************************************************************************************************************************************************************' 

'####################################################################################################################################

' To close all the excels sheets and browser present in the system
SystemUtil.CloseProcessByName("Excel.exe")
SystemUtil.CloseProcessByName("Chrome.exe")

' Run  QTP  in minimize mode
Set QtApp = CreateObject("QuickTest.Application") 
QtApp.WindowState = "Minimized"

'Execute Library Function file
'LoadFunctionLibrary ("C:\Users\demo\Documents\UFT One\HybridFramework\FunctionLibrary.qfl")

'Give the path of the Data file
Environment.Value("strFilePath") =  "C:\Users\demo\Documents\UFT One\DXC_Training\DataSheet\CreateAndDisplayOrders.xlsx" 

'Create an Excel Object and open the input data file
Set xlObj = CreateObject("Excel.Application") 
 xlObj.WorkBooks.Open Environment.Value("strFilePath") 
 Set xlWB = xlObj.ActiveWorkbook 
 Set xlSheet = xlWB.WorkSheets("CreateAndDisplayOrder") 

Environment.Value("AllRows") = xlSheet.UsedRange.Rows.Count

For intcurrentRow = 2 to Environment.Value("AllRows")

	Call AOS_Login (GetColValue("DT_Browser"), GetColValue("DT_Url"),GetColValue("DT_Username"), GetColValue("DT_Password"))
	Call AddProductToCart (GetColValue("DT_Quantity"))
	Call CheckoutAndRetrieveOrderNumber (GetColValue("DT_Username"), GetColValue("DT_SafePayPassword"))
	Call DisplayOrder (GetColValue("DT_OrderNumber"))
	Call AOS_Logoff ()
	
Next

xlWB.Save
xlObj.Quit

'*******************************************************************************End of Script******************************************************************************************



 'Function Name  GetColValue
  'Description  : Returns column no. based on column name
  Function GetColValue(stringCN)
			intColumnCnt=xlSheet.usedrange.Entirecolumn.count
            For i = 1 to intColumnCnt
				If (stringCN = xlSheet.Cells(1,i).value) Then
					If   xlSheet.Cells(intCurrentRow,i).value <> "" Then
						GetColValue = xlSheet.Cells(intCurrentRow,i).value
					Else
						Reporter.ReportEvent micFail,"Input Data Validation", stringCN & " Value in datasheet  is empty " 
					End If					
                    			Exit for
				End If
			Next
End Function
'--------------------------------------------------------------------------------------------------------------------------
'===================================================================================
' Function Name: SetXLVal
' Description  : To set Value to XL sheet
' Return Value : Column name, Row no and cell value
Function SetXLVal(ColumnName,RowNo,CellValue)
 intColumnCnt=xlSheet.usedrange.Entirecolumn.count
 For i = 1 to intColumnCnt
	  If (ColumnName = cstr(xlSheet.Cells(1,i).value)) Then
	   ColValue = i
	   Exit for
	  End If
 Next
 xlSheet.Cells(RowNo,ColValue)=CellValue
End Function

Function AOS_Login (varBrowser, varURL, varUsername, VarPassword)
		'Dim varBrowser, varURL, varUsername, VarPassword
		'Launch AOS URL in specified browser
		SystemUtil.Run varBrowser, varURL
		'Maximize Browser
		Browser("title:=Advantage Shopping").Maximize()
		Wait (2)
		'Login
		Browser("Advantage Shopping").Page("Advantage Shopping").Link("UserMenu").Click()
		Wait(2)
		Browser("Advantage Shopping").Page("Advantage Shopping").WebEdit("username").Set varUsername
		Browser("Advantage Shopping").Page("Advantage Shopping").WebEdit("password").SetSecure VarPassword
		Wait (3)
		Browser("Advantage Shopping").Page("Advantage Shopping").WebButton("sign_in_btnundefined").Click()
		Wait (3)
End Function

Function AOS_Logoff ()
		'Logoff from AOS
		Browser("Advantage Shopping").Page("Advantage Shopping").Link("UserMenu_2").Click
		Browser("Advantage Shopping").Page("Advantage Shopping").Link("Sign out").Highlight
		Browser("Advantage Shopping").Page("Advantage Shopping").Link("Sign out").Click
		Wait (3)
		'Close Browser
		Browser("Advantage Shopping").Close
End Function

Function AddProductToCart (varQuantity)
		Browser("Advantage Shopping").Page("Advantage Shopping").Link("TabletsCategory").Click
		Wait (2)
		Browser("Advantage Shopping").Page("Advantage Shopping").WebElement("Item_Amount").Click
		Wait (3)
		Browser("Advantage Shopping").Page("Advantage Shopping").WebElement("rabbit").Click
		Browser("Advantage Shopping").Page("Advantage Shopping").WebEdit("quantity").Set varQuantity
		Wait (2)
		Browser("Advantage Shopping").Page("Advantage Shopping").WebButton("save_to_cart").Highlight
		Browser("Advantage Shopping").Page("Advantage Shopping").WebButton("save_to_cart").Click	
End Function

Function CheckoutAndRetrieveOrderNumber (varUsername, varPassword)
		Browser("Advantage Shopping").Page("Advantage Shopping").Link("ShoppingCart").Click
		Browser("Advantage Shopping").Page("Advantage Shopping").WebButton("check_out_btn").Highlight
		Browser("Advantage Shopping").Page("Advantage Shopping").WebButton("check_out_btn").Click
		Wait (3)
		Browser("Advantage Shopping").Page("Advantage Shopping").WebButton("next_btn").Click
		Wait (2)	
		Browser("Advantage Shopping").Page("Advantage Shopping").WebEdit("safepay_username").Set varUsername
		Browser("Advantage Shopping").Page("Advantage Shopping").WebEdit("safepay_password").SetSecure varPassword
		Wait (2)
		Browser("Advantage Shopping").Page("Advantage Shopping").WebButton("pay_now_btn_SAFEPAY").Highlight
		Browser("Advantage Shopping").Page("Advantage Shopping").WebButton("pay_now_btn_SAFEPAY").Click
		Browser("Advantage Shopping").Page("Advantage Shopping").WebElement("Thank you for buying with").Check CheckPoint("Thank you for buying with Advantage")
		Browser("Advantage Shopping").Page("Advantage Shopping").WebElement("orderNumberLabel").Highlight
		VarOrderNumber = Browser("Advantage Shopping").Page("Advantage Shopping").WebElement("orderNumberLabel").GetROProperty ("innertext")
		'These codes are included to save the data into the respective sheets
		Set xlSheet = nothing
		For Iter = 1 To xlWB.Worksheets.Count
			 If xlWB.Worksheets(Iter).Name = "CreateAndDisplayOrder" Then 
				 Set xlSheet = xlWB.Worksheets(Iter)
				 setxlval "DT_OrderNumber",intCurrentRow, VarOrderNumber
				 Exit For 
		     End If 
		Next 
		Browser("Advantage Shopping").Page("Advantage Shopping").Link("HOME").Click
End Function

 Function DisplayOrder (varOrderNumber)
 		Browser("Advantage Shopping").Page("Advantage Shopping").Link("UserMenu_2").Click
 		Browser("Advantage Shopping").Page("Advantage Shopping").Link("My orders").Highlight
		Browser("Advantage Shopping").Page("Advantage Shopping").Link("My orders").Click
		Wait (2)
		'Browser("Advantage Shopping").Page("Advantage Shopping").WebElement("menuSearch").Click
		'Browser("Advantage Shopping").Page("Advantage Shopping").WebElement("menuSearch").Click
		'Wait (2)
		'Browser("Advantage Shopping").Page("Advantage Shopping").WebEdit("Search in orders").Set varOrderNumber
		'Wait (2)
		'Browser("Advantage Shopping").Page("Advantage Shopping").WebElement("Tablet").Highlight
		Browser("Advantage Shopping").Page("Advantage Shopping").Link("HOME").Click
 End Function
		
