'##################################################################################################################################
'Project Name : AOS 
'File Name: Create and display  order
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

' To close all the excels sheets present in the system
SystemUtil.CloseProcessByName("Excel.exe")

' Run  QTP  in minimize mode
Set QtApp = CreateObject("QuickTest.Application") 
QtApp.WindowState = "Minimized"

'Execute Library Function file
LoadFunctionLibrary ("C:\Users\demo\Documents\UFT One\HybridFramework\FunctionLibrary.qfl")

'Give the path of the Data file
Environment.Value("strFilePath") =  "C:\Users\demo\Documents\UFT One\DXC_Training\DataSheet\CreateAndDisplayOrders.xlsx" 

'Create an Excel Object and open the input data file
Set xlObj = CreateObject("Excel.Application") 
 xlObj.WorkBooks.Open Environment.Value("strFilePath") 
 Set xlWB = xlObj.ActiveWorkbook 
 Set xlSheet = xlWB.WorkSheets("CreateAndDisplayOrder") 

Environment.Value("AllRows") = xlSheet.UsedRange.Rows.Count

xlWB.Save
xlObj.Quit

For intcurrentRow = 2 to Environment.Value("AllRows")

Call AOS_Login (

		RunAction "Action1 [SAPLogon]", oneIteration,intcurrentRow,RunStatusLogin
			If  RunStatusLogin = "PASS" Then		
				RunAction "Action1 [VA01_CreateSalesOrder]", oneIteration, intcurrentRow,RunStatusCreateSO
			End If
			If  RunStatusCreateSO = "PASS" Then				
				RunAction "Action1 [VA03_DisplaySalesOrder]", oneIteration, intcurrentRow,RunStatusDisplay
			End If
		RunAction "Action1 [SAPLogOff]", oneIteration

Next

'*******************************************************************************End of Script******************************************************************************************

 'Function Name  GetColValue
         'Description  : Returns column no. based on column name

		 Public Function GetColValue(stringCN)
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
end Function


Function AOS_Login (Browser, URL, Username, Password)
		'Launch AOS URL in specified browser
		SystemUtil.Run Browser, URL
		'Maximize Browser
		Browser("title:=Advantage Shopping").Maximize     
		'Login
		Browser("Advantage Shopping").Page("Advantage Shopping").Link("UserMenu").Click @@ script infofile_;_ZIP::ssf11.xml_;_
		Browser("Advantage Shopping").Page("Advantage Shopping").WebEdit("username").Set Username @@ script infofile_;_ZIP::ssf12.xml_;_
		Browser("Advantage Shopping").Page("Advantage Shopping").WebEdit("password").SetSecure Password @@ script infofile_;_ZIP::ssf13.xml_;_
		Browser("Advantage Shopping").Page("Advantage Shopping").WebButton("sign_in_btnundefined").Click @@ script infofile_;_ZIP::ssf14.xml_;_
		Wait (3)
End Function


		
