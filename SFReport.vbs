'strScreenshotPath = "C:\Temp\Screenshots_SF"
strResultsPath = "C:\"
'DeleteScreenshotFolder strScreenshotPath

'#################################################################################################################
'Function Name: GetCurrentDate
'Description: This function returns the current system date in mm_dd_yyyy format for creating folder
'Parameters: No
'Created By: Satheesh Kumar Duraisamy
'Created Date: 
'Modified By: NA
'Modified Date: NA
'Modification History: NA
'Comments: NA
'#################################################################################################################
Function GetCurrentDate ()
	strMyDate = Date()
	GetCurrentDate = Cstr(Month(strMyDate))+"_"+Cstr(Day(strMyDate))+"_"+Cstr(Year(strMyDate))
End Function
'#################################################################################################################

'#################################################################################################################
'Function Name: CurrDate
'Description: This function is used for format the system date used in excel result
'Parameters: No
'Created By: Satheesh Kumar Duraisamy
'Created Date: 
'Modified By: NA
'Modified Date: NA
'Modification History: NA
'Comments: NA
'#################################################################################################################

Public Function CurrDate (TDate)
	Dim dtmTDate, dtmCurrentDate
'	dtmTDate = Replace(TDate, "/", "_")
	dtmTDate = Replace(TDate, ":", "_")
	dtmCurrentDate = Replace(dtmTDate, " ", "_")
	Environment.Value("CurrDate") = dtmCurrentDate
End Function
'#################################################################################################################

'#################################################################################################################
'Function Name: CreateResultFolder
'Description: This function is used to create datewise folder under results folder
'Parameters: No
'Created By: Satheesh Kumar Duraisamy
'Created Date:
'Modified By: NA
'Modified Date: NA
'Modification History: NA
'Comments: NA
'#################################################################################################################
Function CreateResultFolder ()
	Dim objFileSystemObject, objFolder
	Set objFileSystemObject = CreateObject("Scripting.FileSystemObject")
	strNewFolder = strResultsPath & "SFReports"
	Environment.Value("strNewFolder") = strNewFolder
	strFolderPath = strNewFolder' & "Run_Folder_" & GetCurrentDate()
	
	If (objFileSystemObject.FolderExists(strFolderPath)) Then
		CreateResultFolder = strFolderPath
	Else
		Set objFolder = objFileSystemObject.CreateFolder(strFolderPath)
		CreateResultFolder = objFolder.Path
	End IF
	strFolderPath = strNewFolder & "\Run_Folder_" & GetCurrentDate()
	Environment.Value("strFolderPath") = strFolderPath
	If (objFileSystemObject.FolderExists(strFolderPath)) Then
		CreateResultFolder = strFolderPath
	Else
		Set objFolder = objFileSystemObject.CreateFolder(strFolderPath)
		CreateResultFolder = objFolder.Path
	End IF
End Function
'#################################################################################################################

'#################################################################################################################
'Function Name: InitResultFile
'Description: This function is used for Creating excel for writing results
'Parameters: No
'Created By: Satheesh Kumar Duraisamy
'Created Date: 
'Modified By: NA
'Modified Date: NA
'Modification History: NA
'Comments: NA
'#################################################################################################################
Public Function InitResultFile ()
	Dim strScriptName, dtmCurrentDate, strFName, ExcelBook, strFilePath
	strScriptName = Environment.Value("TName")
	dtmCurrentDate = Trim(Environment.Value("CurrDate"))
	strFName = Environment.Value("strFolderPath")'Environment.Value("strFName")
	Set objExcel = CreateObject("Excel.Application")
	strFilePath = strFName &"\"& strScriptName & "--" & dtmCurrentDate &".xlsx"
	objExcel.Workbooks.Add
	objExcel.cells(1,1) = "Step No"
	objExcel.cells(1,2) = "Activities"
	objExcel.cells(1,3) = "Description"
	objExcel.cells(1,4) = "Status"
	objExcel.cells(1,5) = "ScreenShot"
	Set objRange1 = objExcel.Range("A1","E1")
	objRange1.Font.Bold = True
	objRange1.Interior.ColorIndex = 30
	ObjRange1.Font.ColorIndex = 2
	objExcel.Worksheets("sheet1").Name = "TestResult"
	objExcel.ActiveWorkbook.SaveAs strFilePath', fileformat = xlExcel8
	objExcel.Quit
End Function
'#################################################################################################################

'#################################################################################################################
'Function Name: CaptureScreen
'Description: This function capture the screenshot 
'Parameters: No
'Created By: Satheesh Kumar Duraisamy
'Created Date: 
'Modified By: NA
'Modified Date: NA
'Modification History: NA
'Comments: NA
'#################################################################################################################
Public Function CaptureScreen()
	Dim  Dfso, Df
	Set Dfso = CreateObject("Scripting.FileSystemObject")
	strCSPath = Environment.Value("strNewFolder")
	strSSPath = strCSPath &"\"&"Screenshot"
	If Not Dfso.FolderExists(strSSPath) Then
		Set Df = Dfso.CreateFolder(strSSPath)
	End If   
	ScreenShot_TempPath=strSSPath &"\" & Environment.Value("TName") & "_" & minute(Time) & "_" & second(Time) & ".png" 
	Desktop.CaptureBitmap ScreenShot_TempPath, True 
	Environment.Value("ScreenShot") = ScreenShot_TempPath
'  AttachScreenshot_ToQC (ScreenShot_TempPath)
'  Dfso.DeleteFile ScreenShot_TempPath
	Set Dfso = Nothing
End Function
'#################################################################################################################

'#################################################################################################################
'Function Name: WriteResult
'Description: This function is used write the results in excel sheet
'Parameters: strStepNo, strActivity, strDesc, strStatus, strSSCapture
'Created By: Satheesh Kumar Duraisamy
'Created Date: 
'Modified By: NA
'Modified Date: NA
'Modification History: NA
'Comments: NA
'#################################################################################################################

Public Function WriteResult (strStepNo, strActivity, strDesc, strStatus, strSSCapture)
'	Dim strScriptName, dtmCurrentDate, strFName, strStepNo, strActivity, strDesc, strStatus
	strScriptName = Environment.Value("TName")
	dtmCurrentDate = Trim(Environment.Value("CurrDate"))
	strFName = Environment.Value("strFolderPath")'Environment.Value("strFName")
	Set WriteExcel = CreateObject("Excel.Application")
	WriteExcel.Workbooks.Open strFName & "\" & strScriptName & "--" & dtmCurrentDate & ".xlsx"
	Set WriteSheet = WriteExcel.ActiveWorkbook.Worksheets ("TestResult")
		
'Write results in respective column
	intERow = WriteSheet.UsedRange.Rows.Count
	WriteExcel.Cells(intERow+1, 1) = strStepNo
	WriteExcel.Cells(intERow+1, 2) = strActivity
	WriteExcel.Cells(intERow+1, 3) = strDesc
	WriteExcel.Cells(intERow+1, 4) = strStatus

			
	If strSSCapture = "YES" OR strSSCapture = "Yes" Then
		CaptureScreen
		ExlFormula = "=HYPERLINK(" & chr(34) & Environment.Value("ScreenShot") & chr(34)  & ",""Click here for screen shot"")"
		WriteExcel.Cells(intERow+1, 5) = ExlFormula 'Environment.Value("strScreenShotClick")
'		WriteExcel.Cells(intERow+1, 5) = "Screenshot is captured and saved in UFT Log"
	Else
'		CaptureScreen
'		WriteExcel.Cells(intERow+1, 5) = "Screenshot is not captured"
	End If
'Save work sheet
	WriteExcel.ActiveWorkbook.Save
'Close work sheet
	WriteExcel.ActiveWorkbook.Close
'Close Excek application
	WriteExcel.Application.Quit
	Set WriteSheet = Nothing
	Set WriteExcel = Nothing
	
	If strStatus = "Pass" Then
		If strSSCapture = "YES" OR strSSCapture = "Yes" Then
			CaptureScreen
			Reporter.ReportEvent micPass, strStepNo & " - " & strActivity, strDesc, Environment.Value("ScreenShot") 
		Else
			Reporter.ReportEvent micPass, strStepNo & " - " & strActivity, strDesc
		End If
	Else
		CaptureScreen
		Reporter.ReportEvent micFail, strActivity, strDesc, Environment.Value("ScreenShot") 
	End If
End Function
'#################################################################################################################

'#################################################################################################################
'Function Name: DeleteScreenshotFolder
'Description: This function is used to delete the screen shot folder
'Parameters: No
'Created By: Satheesh Kumar Duraisamy
'Created Date: 
'Modified By: NA
'Modified Date: NA
'Modification History: NA
'Comments: NA
'#################################################################################################################
Public Function DeleteScreenshotFolder()
	Dim  Dfso, Df
	Set Dfso = CreateObject("Scripting.FileSystemObject")
	If Dfso.FolderExists("C:\SFReports\Screenshot") then
		Dfso.DeleteFolder("C:\SFReports\Screenshot")
	End If   
	Set Dfso = Nothing
End Function
'#################################################################################################################


