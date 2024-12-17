'Script Name: LOI_Submission_RA_LOI_And_Application_001_02
'Option Explicit
'Functions For Generating Report and Screenshot
strScriptName = Environment.Value("TestName")
Environment.Value("TName") = strScriptName
GetCurrentDate
CurrDate (Time)
CreateResultFolder
InitResultFile
CaptureScreen
'Reporter.ReportEvent micDone, "Testing", "Screenshot", Environment.Value("ScreenShot_TempPath")
'Sample Result for Excel'
'WriteResult "Step 1", "Verify Login Page is Displayed", "Login Page is loaded successfully", "Pass", "YES"

'Variable Declarations
Dim strBrowser, strURL
Dim strPcoriOnline, strFound, strVerifyPcoriOnline, strHelpInfo, strUserName, strPassword

strBrowserType = DataTable.Value("strBrowserType", dtGlobalSheet)
strURL = DataTable.Value("strURL", dtGlobalSheet)
strPcoriOnline = DataTable.Value("strPcoriOnline", dtGlobalSheet)
strUserName = DataTable.Value("strUserName", dtGlobalSheet)
strPassword = DataTable.Value("strPassword", dtGlobalSheet)

'################################################################### Step 1 #####################################################################################################
'Step 1: Navigate to "test.salesforce.com". Enter your own system admin credentials and login to Salesforce
Open_Pcori_Application strBrowserType, strURL

'################################################################### Step 2 #####################################################################################################
'Step 2: Verify on the right side of "User Name" and "Password" text present: "PCORI Online is now open for:"   
'At top of page: 'pcori' text, logo and 'Patient-Centered Outcomes Research Institute' text. 
'Below is a 'User Name' field with help text bubble upon hover: Your user name is your email address 'Password' field above a Log In button 
'with the text links: 'Forgot your password? I New User?  and Click here to visit the PCORI website

Set strBrowser = Browser("name:=Login.*").page("title:=Login.*")
If strBrowser.Exist(120) Then
	strVerifyPcoriOnline = strBrowser.WebElement("innertext:=PCORI Online is now open for.*", "visible:=True", "html tag:=SPAN").GetAllROProperties("innertext")
End If
'Browser("name:=Login.*").Sync

If Instr(1, strVerifyPcoriOnline , "PCORI Online is now open for:") > 0 Then
	strBrowser.WebElement("innertext:=PCORI Online is now open for.*", "visible:=True", "html tag:=SPAN").Highlight
	WriteResult "Step 1", "Verify Application Login", "Application is launched Successfully", "Pass", "YES"
	Reporter.ReportEvent micPass, "Verify PCORI Online is now open for:", "PCORI Online is now open for:" & " text is present on the right side of the screen", Environment.Value("ScreenShot")
Else
	WriteResult "Step 1", "Verify Application Login", "Application is launched Successfully", "Fail", "YES"
	Reporter.ReportEvent micFail, "Verify PCORI Online is now open for:", "PCORI Online is now open for:" & " text is not present on the right side of the screen", Environment.Value("ScreenShot")
End If

'Verify Logo and Text
If strBrowser.Image("file name:=PCORI-Banner.*", "name:=Image").Exist(2) Then
	strBrowser.Image("file name:=PCORI-Banner.*", "name:=Image").Highlight
	Reporter.ReportEvent micPass, "Verify Logo and Text", "Logo and Text is present at the top of the page"
Else
	Reporter.ReportEvent micFail, "Verify Logo and Text", "Logo and Text is not present at the top of the page"
End If
'Verify hover help
strFound = 0
Do 
	strBrowser.WebElement("html tag:=LIGHTNING-PRIMITIVE-ICON", "innerhtml:=<svg focusable.*").FireEvent "OnClick"
	If strBrowser.WebElement("innertext:=Your user name is your email address.*", "class:=slds-popover__body").Exist(.2) Then
		strHelpInfo = strBrowser.WebElement("innertext:=Your user name is your email address.*", "class:=slds-popover__body").GetAllROProperties ("innertext")
		strFound = 1
	End If
Loop Until strFound = 1
'User Name and Password Mandatory field Validation
strMandatory = Browser("title:=Login.*").Page("title:=Login.*").WebElement("innertext:=.*Password", "html tag:=DIV").GetROProperty("outerhtml")
If Instr(1, strMandatory, "title=""required""") > 0 Then
	WriteResult "Step 1.1", "Verify User Name field is Mandatory", "User Name field is marked as Mandatory", "Pass", "No"
Else
	WriteResult "Step 1.1", "Verify Password field is Mandatory", "Password field is marked as not Mandatory", "Fail", "No"
End If

strMandatory1 = Browser("title:=Login.*").Page("title:=Login.*").WebElement("innertext:=.*User Name.*", "html tag:=LABEL").GetROProperty("outerhtml")
If Instr(1, strMandatory1, "title=""required""") > 0 Then
	WriteResult "Step 1.2", "Verify Password field is Mandatory", "Password field is marked as Mandatory", "Pass", "No"
Else
	WriteResult "Step 1.2", "Verify Password field is Mandatory", "Password field is marked as not Mandatory", "Fail", "No"
End If

If strHelpInfo = "Your user name is your email address."  Then
	Reporter.ReportEvent micPass, "Verify Login Help Text", "Your user name is your email address. is displayed while hover the mouse"
Else
	Reporter.ReportEvent micFail, "Verify Login Help Text", "Your user name is your email address. is not displayed while hover the mouse"
End If
'Verify User Name, Password, Click here, Privacy Policy, Click here to visit the PCORI website., Forgot your password?, New User? and Log In fleids
VerifyEditBox "Step 2", strBrowser, "cmntyusrnm", "User Name", "No"
VerifyEditBox "Step 2.1", strBrowser, "cmntypswd", "Password", "No"
VerifyWebElementText "Step 2.2", strBrowser, "LABEL", ".*User Name", "No"
VerifyWebElementText "Step 2.3", strBrowser, "LABEL", ".*Password", "No"
VerifyLink  "Step 2.4", strBrowser, "Click here ", "Click here", "No"
VerifyLink  "Step 2.5", strBrowser,"Privacy Policy", "Privacy Policy", "No"
VerifyLink  "Step 2.6", strBrowser,"Click here", "Click here", "No"
VerifyText "Step 2.7", strBrowser,"to visit the PCORI website.", "Click here to visit the PCORI website.", "No"
VerifySFLButton "Step 2.8", strBrowser,"Forgot your password.*", "Forgot your password?", "No"
VerifySFLButton "Step 2.9", strBrowser,"New User.*", "New User?", "No"
'VerifyButton "Step 2.7", strBrowser,"Forgot your password?", "Forgot your password?", "No"

'################################################################### Start of Login to Application ##################################################################################
'Enter values to user name and password and click on Login Button
Set_WebEdit "Step 2.10",  strBrowser, "cmntyusrnm", strUserName, "User Name", "No"
Set_WebEditSecure "Step 2.11", strBrowser, "cmntypswd", strPassword, "Password", "No"
Click_SFLWebButton "Step 2.12", strBrowser, "Log In", "Log In", "YES"
'Browser("name:=Home.*").Sync
'Verify Login successful ADVISORY PANELS
'################################################################### End of Login ##############################################################################################
'#############################################################################################################################################################################

'############################################################################################################################################################################# @@ script infofile_;_ZIP::ssf17.xml_;_
 '################################################################### Verify Home Page and veirfy headings #######################################################################
' VerifyPageNavigation "Step 3", "name:=Home.*", "title:=Home.*", "Home Page is navigated Succesfully", "YES"
strHeadings = "Advisory Panels||Merit Review.*||Eugene Washington PCORI Engagement Award Program||PCORI Research Awards||Ambassador Program||PCORI D&I Awards||Infrastructure||Peer Review.*||Health Systems Implementation Initiative.*||Methodology Committee"
Set strBrowser = Browser("name:=Home.*").page("title:=Home.*")
If strBrowser.Exist(60) Then
	VerifyWebElementText "Step 3", strBrowser, "B", strHeadings, "No"
End If
'#############################################################################################################################################################################
'#############################################################################################################################################################################

'#############################################################################################################################################################################
'################################################################### Step 4, 5, 6, 7 #############################################################################################
'Click on Research Awards
Click_SFLWebButton "Step 4", strBrowser, "RESEARCH AWARDS", "Research Awards", "No"
Set strBrowser = Browser("name:=researchawardslandingpage.*").page("title:=researchawardslandingpage.*")
'Browser("name:=researchawardslandingpage.*").Sync

'Click on Funding Opportunities
If strBrowser.Exist(60) Then
	ClickLink "Step 5", strBrowser, "", "Funding Opportunities","Funding Opportunities", "No"
	Browser("name:=FundingOpportunities.*").Sync
End If

'Click on Funding Opportunities based on our requirement
Set strBrowser = Browser("name:=FundingOpportunities.*").page("title:=FundingOpportunities.*").Frame("title:=Explore Grant Opportunity.*")
strFOpprotunities = DataTable.Value("strFOpprotunities", dtGlobalSheet)
If strBrowser.Exist(60) Then
	ClickLink "Step 6", strBrowser, "0", strFOpprotunities, strFOpprotunities, "No"
End If

'Click on Apply button in Funding opporutnities page and verify the elements in the page
strDescription = "This Broad PCORI Funding Announcement seeks comparative clinical effectiveness research applications that address four of PCORI’s National Priorities for Health: Increase Evidence for Existing Interventions and Emerging Innovations in Health; Accelerate Progress Toward an Integrated Learning Health System; Achieve Health Equity; and Advance the Science of Dissemination, Implementation, and Health Communication. Applicants to the Cycle 1 2025 BPS PFA may select from PCORI's topic themes, as applicable, that speak to everyday health issues facing large numbers of Americans: promoting health for children, youth, and older adults; addressing violence and trauma, or substance use; improving mental and behavioral health; and clinical conditions such as cardiovascular disease, sleep disturbance, and pain management."
strInstruction1 = "PCORI funding opportunities are here: https://www.pcori.org/funding-opportunities"
strHeadings1 = "Please click on the Apply button to create your LOI for the Broad Pragmatic Studies."
Set strBrowser = Browser("name:=Sfdc Page.*").Page("title:=Sfdc Page").Frame("title:=.*PCORI Online.*")
If strBrowser.Exist(60) Then
	VerifyWebElementText "Step 7", strBrowser, "H2", "Broad Pragmatic Studies", "No"
End If
'Verify Elements in the page
'VerifyWebElementText "Step 7", strBrowser, "H2", "Broad Pragmatic Studies", "No"
strHeadings = "Description||Instructions"
VerifyWebElementText "Step7", strBrowser, "SPAN", strHeadings, "No"
VerifyWebElementText "Step7.2", strBrowser, "SPAN", strDescription, "No"
VerifyWebElementText "Step 7.3", strBrowser, "P", strHeadings1, "No"
VerifyWebElementText "Step 7.4", strBrowser, "P", strInstruction1, "No"
VerifyWebElementText "Step 7.5", strBrowser, "SPAN", "24C2 - LOI Form - BPS.*", "No"
ReadWebElementPosted "Step 7.6", strBrowser, "DIV", "Posted On.*", "No"
NavigateLinkhere "Step 8", strBrowser, "A", "https://www.pcori.org/funding-opportunities", "", "Funding Opportunities.*", "No"
Wait(2)
VerifyLink "Step 9", strBrowser, "Apply", "Apply", "No"
ClickLink "Step 10", strBrowser, "", "Apply", "Apply", "No"
'Browser("name:=Sfdc Page.*").Sync
Wait(5)
If Browser("name:=Sfdc Page.*").Page("title:=Sfdc Page").Frame("title:=.*PCORI Online.*").Link("name:=Continue", "index:=0").Exist(120) Then
	Set strBrowser = Browser("name:=Sfdc Page.*").Page("title:=Sfdc Page").Frame("title:=.*PCORI Online.*")
	ClickLink "Step 11", strBrowser,  "1", "Continue", "Continue", "No"
	Browser("name:=Sfdc Page.*").Sync
	Wait(5)
End If
'#############################################################################################################################################################################
'#############################################################################################################################################################################

'#############################################################################################################################################################################
'################################################################### Contact Information Page ###################################################################################
strBPStudies = "Satheesh Duraisamy||Merry||merry123@gmail.com||Agila Somasundaram||Abigail Keatts||Aleksandra Modrow||Agila Somasundaram||One_Smoke Test Account_DB||Washington DC||Automation"
strFieldName = "Principal Investigator \(Contact\)||Dual Principal Investigator Name||Dual Principal Investigator Email||Administrative Official||PI Designee 1||PI Designee 2||Financial Officer||Organization||Congressional District||Department"
strContactInfoText = "For instructions on using our system, click here to access our PCORI Portal User Guide.||To view additional information about the PCORI submission process, click here to view our FAQ page.||- A Principal Investigator \(PI\) and an Administrative Official \(AO\) with a valid email address are required to submit your LOI to PCORI.||- To assign a user, click the lookup icon and start to type their name. If the user does not exist in our system, they must register in PCORI Online by clicking “New User.”||- User login information from the previous PCORI Online were not migrated to the new PCORI Online.||- The AO and the PI cannot be the same individual.||- Individuals assigned at the “Contact Information” tab will have access to the LOI and Application.||- To find your Congressional District please click here.||- Fields marked with \(\*\) are required\."
strContactInfoText1 = "Click 'Save & Next' to continue to the next tab. Otherwise you could receive an error message."
strTabLinks = "Contact Information||Pre Screen Questionnaire||Resubmission||PI Information||Project Information||Project Personnel||Templates and Uploads"
strButtonNames1 = "Cancel||Review/Submit"
strButtonNames2 = "Save||Save & Next||Clear Changes"

'Verify tab Contact Information page
VerifyTabLinkGrp "Step 12", strBrowser, "A", strTabLinks, "No"

'Verify Button Names Contact Information page at the top
VerifyWebButtonGrp "Step 13", strBrowser, "INPUT", strButtonNames1, strSShot

'Navigate all links in Contact Information page
NavigateLinkhere "Step 14", strBrowser, "A", "here", "0", "PCORI-Online-Pre-Award-User-Guide.pdf", "YES"
NavigateLinkhere "Step 15", strBrowser, "A", "here", "1", "Frequently Asked Questions.*", "YES"
NavigateLinkhere "Step 16", strBrowser, "A", "PCORI Online ", "", "Login", "YES"
NavigateLinkhere "Step 17", strBrowser, "A", "here", "2", "Find Your Representative.*", "YES"

'Verify contact info text
VerifyWebElementText "Step 18", strBrowser, "P", strContactInfoText, "No"
VerifyWebElementText "Step 19", strBrowser, "DIV", strContactInfoText1, "No"
VerifyWebElementText "Step 20", strBrowser, "DIV", strFieldName, "No"

'Verify Mandatory fields (Text Boxes)
VerifyEditListMandatory "Step 20.1", strBrowser, "DIV", "", strCIPageMandatoryTextFieldID, strCIPageMandatoryTextFieldName, strCIPageMandatoryText, "No"

'Enter values in the text boxes
Set_WebEditGrp "Step 21", strBrowser, strBPStudies, strFieldName, "No"

'Verify Button Names Contact Information pagevat the bottom
VerifyWebButtonGrp "Step 22", strBrowser, "INPUT", strButtonNames2, strSShot

'Click on Save and Continue button
Click_WebButton "Step 23", strBrowser, "INPUT", "0", "Save & Next", "Save & Next", "No"
'#############################################################################################################################################################################
'#############################################################################################################################################################################

'#############################################################################################################################################################################
'################################################################### Pre Screen Questionnaire ###################################################################################
strPreScreenQ1 = strPSQText1 &"||" & strPSQText2
strPreScreenQ2 = strPSQText3 &"||" & strPSQText4 &"||" & strPSQText5 &"||" & strPSQText6 &"||" & strPSQText7 &"||" & strPSQText9 &"||" & strPSQText10
strPreScreenQ3 = strPSQText8
strRSQ1 = strRSText1 &"||" & strRSText2
strRSQ2 = strRSText3 &"||" & strRSText4


If strBrowser.WebElement("innertext:=Do any of the specific aims of your research propose:", "index:=0").Exist(120) Then
	WriteResult "Step 24", "Verify for Pre Screen Questionnaire page load", "Pre Screen Questionnaire page is loaded successfully", "Pass", strSShot
Else
	WriteResult "Step 24", "Verify for Pre Screen Questionnaire page load", "Pre Screen Questionnaire page is not loaded successfully", "Fail", "YES"
End If
'Verify all text elements in pre screen questionnaire page
VerifyWebElementText "Step 25", strBrowser, "P", strPreScreenQ1, "No"
VerifyWebElementText "Step 26", strBrowser, "DIV", strPreScreenQ2, "No"
VerifyWebElementText "Step 27", strBrowser, "B", strPreScreenQ3, "No"

'Verify for Mandatory field (WebList)
strPSQPageMandatoryListName = strPSQText4 & "||"& strPSQText5 & "||"& strPSQText6 & "||"& strPSQText7 & "||"& strPSQText9
VerifyEditListMandatory "Step 27.1", strBrowser, "DIV", "", strPSQPageMandatoryListFieldID, strPSQPageMandatoryListName, strPSQPageMandatoryLIST, "No"

'Verify values in the web list box and select the required value
strPSQPageMandatoryListName1 = strPSQText4 & "||"& strPSQText5 & "||"& strPSQText6 & "||"& strPSQText7 & "||"& strPSQText9
VerifyWebListValuesGrp "Step 27.2", strBrowser, "SELECT", "0", strPSQPageMandatoryListName1, strPSQPageDropDownValues, "No"
strPSQPageMandatoryListName = strPSQText4 & "||"& strPSQText5 & "||"& strPSQText6 & "||"& strPSQText7 & "||"& strPSQText9
VerifySelectWebListGrp "Step 28", strBrowser, "SELECT", "0", strPreScreenQ4, strPSQPageMandatoryListName, "No"

'Click on Save and Continue button
Click_WebButton "Step 29", strBrowser, "INPUT", "0", "Save & Next", "Save & Next", "No"
 Browser("name:=Sfdc Page.*").Sync
Wait(5)
'#############################################################################################################################################################################
'#############################################################################################################################################################################

'#############################################################################################################################################################################
'################################################################### Re-Submission Questionnaire ###############################################################################
'Enter all fields in Re-Submission Questionnaire page
If strBrowser.WebElement("innertext:=LOI, Application Resubmission Questions", "index:=0").Exist(60) Then
	WriteResult "Step 29.1", "Verify for Re-Submission Questionnaire page load", "Re-Submission Questionnaire page is loaded successfully", "Pass", strSShot
Else
	WriteResult "Step 29.1", "Verify for Re-Submission Questionnaire page load", "Re-Submission Questionnaire page is not loaded successfully", "Fail", "YES"
End If
VerifyWebElementText "Step 30", strBrowser, "P", strRSQ1, "No"
NavigateLinkhere "Step 31", strBrowser, "A", "here", "0", "PCORI-Online-Pre-Award-User-Guide.pdf", "No"
NavigateLinkhere "Step 32", strBrowser, "A", "here", "1", "Frequently Asked Questions.*", "No"
Wait(2)
VerifyWebElementText "Step 33", strBrowser, "DIV", strRSQ2, "No"

'Verify the list names and text box names
strRSQ3 = strRSListName1 &"||" & strRSListName2 &"||" & strRSListName3 &"||" & strRSListName4 &"||" & strRSListName5 &"||" & strRSListName6
VerifyWebElementText "Step 34", strBrowser, "DIV", strRSQ3, "No"

'Parameters for validating Mandatory List box & Text Aread fields
strRSPageMandatoryFieldName = strRSListName1 & "||" &strRSListName2 & "||" &strRSListName3 & "||" &strRSListName4 & "||" &strRSListName6
VerifyEditListMandatory "Step 34.1", strBrowser, "DIV", "", strRSPageMandatoryFieldID, strRSPageMandatoryFieldName, strRSQPageMandatoryLIST, "No"

'Verify all list boxes and select the appropriate values in Re submission page
strRSListName = strRSListName1 & "||" & strRSListName2 & "||" &strRSListName3 & "||" &strRSListName4
VerifyWebListValuesGrp "Step 34.2", strBrowser, "SELECT", "0", strRSListName, strRSQPageDropDownValues, "No"
strRSListNameSelect = strRSListName1 & "||" & strRSListName2 & "||" &strRSListName3 & "||" &strRSListName4
VerifySelectWebListGrp "Step 35", strBrowser, "SELECT", 0, strRSQ4, strRSListNameSelect, "NO"
Set_WebEditMLine "Step 36", strBrowser, "TEXTAREA", strRSApplicationNumber, 0, "LOI/Application No", strSShot ' DataSheet

'Click on Save and Continue button
Click_WebButton "Step 36.1", strBrowser, "INPUT", "0", "Save & Next", "Save & Next", "No"
'#############################################################################################################################################################################
'#############################################################################################################################################################################

'#############################################################################################################################################################################
'################################################################### PI Information Questionnaire ################################################################################
If strBrowser.WebElement("innertext:=PI Work Telephone Number", "index:=0").Exist(60) Then
	WriteResult "Step 36.2", "Verify for PI Information Questionnaire page load", "PI Information Questionnaire page is loaded successfully", "Pass", "No"
Else
	WriteResult "Step 36.2", "Verify for PI Information Questionnaire page load", "PI Information Questionnaire page is not loaded successfully", "Fail", "YES"
End If
strPIQ1 = strPIQText1 &"||" & strPIQText2 &"||" & strPIQText3 &"||" & strPIQText4 &"||" & strPIQText5 &"||" & strPIQText6 &"||" & strPIQText7 &"||" & strPIQText8 &"||" & strPIQText9 &"||" & strPIQText10 &"||" & strPIQText11 &"||" & strPIQText12 &"||" & strPIQText13 &"||" & strPIQText14
strPIQQ2 = strPIQText15 &"||" & strPIQText16
'Verify and open here link
VerifyWebElementText "Step 37", strBrowser, "P", strPIQQ2, "No"
NavigateLinkhere "Step 38", strBrowser, "A", "here", "0", "PCORI-Online-Pre-Award-User-Guide.pdf", "No"
NavigateLinkhere "Step 39", strBrowser, "A", "here", "1", "Frequently Asked Questions.*", "No"
Wait(2)
'Parameters for validating Mandatory List box & Text Aread fields
strPIQPageMandatoryFieldName =  strPIQText2 &"||" & strPIQText3 &"||" & strPIQText4 &"||" & strPIQText5 &"||" & strPIQText6 &"||" & strPIQText7 &"||" & strPIQText8 &"||" & strPIQText9 &"||" & strPIQText10 &"||" & strPIQText11 &"||" & strPIQText12 &"||" & strPIQText13 &"||" & strPIQText14
VerifyEditListMandatory "Step 39.1", strBrowser, "DIV", "", strPIQPageMandatoryFieldID, strPIQPageMandatoryFieldName, strPIQPageMandatoryLIST, "No"

'VerifyWebElementText "Step 40", strBrowser, "DIV", strRSQ2, "No"
'VerifyWebElementText "Step 41", strBrowser, "DIV", strPIQ1, "No"
Set_WebEditGroup "Step 42", strBrowser, "INPUT", "", strValuePI1, strEditNamePI1, strSShot

'Verify the list of items in the dropdown box
VerifyWebListValuesGrp "Step 42.1", strBrowser, "SELECT", "0", strListNamePI2, strPIQPageDropDownValuesPI2, "No"
VerifyWebListValuesGrp "Step 42.2", strBrowser, "SELECT", "7", strListNamePI3, strPIQPageDropDownValuesPI3, "No"
VerifyWebListValuesGrp "Step 42.3", strBrowser, "SELECT", "8", strListNamePI4, strPIQPageDropDownValuesPI4, "No"
VerifyWebListValuesGrp "Step 42.4", strBrowser, "SELECT", "9", strListNamePI5, strPIQPageDropDownValuesPI5, "No"
VerifyWebListValuesGrp "Step 42.5", strBrowser, "SELECT", "10", strListNamePI6, strPIQPageDropDownValuesPI6, "No"

'Verify and select values for all list box in the screen
VerifySelectWebListGrp "Step 43", strBrowser, "SELECT", "0", strListValuePI2, strListNamePI2, "No"
VerifySelectWebListGrp "Step 44", strBrowser, "SELECT", "7", strListValuePI3, strListNamePI3, "No"
VerifySelectWebListGrp "Step 45", strBrowser, "SELECT", "8", strListValuePI4, strListNamePI4, "No"
VerifySelectWebListGrp "Step 46", strBrowser, "SELECT", "9", strListValuePI5, strListNamePI5, "No"
VerifySelectWebListGrp "Step 47", strBrowser, "SELECT", "10", strListValuePI6, strListNamePI6, "No"

'Verify values in multilist box
VerifyValuesInMultiList "Step 47.1.1", strBrowser, "SELECT", 2, strPIQPageDropDownValuesPI7, "Interation with PCORI", "NO"
VerifyValuesInMultiList "Step 47.1.2", strBrowser, "SELECT", 5, strPIQPageDropDownValuesPI8, "Degree", "NO"
VerifyValuesInMultiList "Step 47.1.3", strBrowser, "SELECT", 12, strPIQPageDropDownValuesPI9, "Grants or Contracts", "NO"

'Select values in multi list box
VerifySelectMultiList "Step 48", strBrowser, "SELECT", 2, 0, strListValuePI7, "Interation with PCORI", "NO"
VerifySelectMultiList "Step 49", strBrowser, "SELECT", 5, 1, strListValuePI8, "Degree", "NO"
VerifySelectMultiList "Step 50", strBrowser, "SELECT", 12, 2, strListValuePI9, "Grants or Contracts", "NO"
Set_WebEditMLine "Step 51", strBrowser, "TEXTAREA", "Please describe Other degree - test", 0, "Other Organizations", "NO"

'Click on Save and Continue button
Click_WebButton "Step 51.1", strBrowser, "INPUT", "0", "Save & Next", "Save & Next", "No"
 Browser("name:=Sfdc Page.*").Sync
Wait(5)
'#############################################################################################################################################################################
'#############################################################################################################################################################################

'#############################################################################################################################################################################
'################################################################### Project Information Questionnaire ###########################################################################
'Webelement paramerer names to verify
If strBrowser.WebElement("innertext:=Project Title", "index:=0").Exist(60) Then
	WriteResult "Step 51.2", "Verify for Project Information Questionnaire page load", "Project Information Questionnaire page is loaded successfully", "Pass", "No"
Else
	WriteResult "Step 51.2", "Verify for Project Information Questionnaire page load", "Project Information Questionnaire page is not loaded successfully", "Fail", "YES"
End If
strPRIQTextV1 = strPRIQText1 &"||" & strPRIQText2 &"||" & strPRIQText3 &"||" & strPRIQText4 &"||" & strPRIQText5 &"||" & strPRIQText6 &"||" & strPRIQText7 &"||" & strPRIQText8 &"||" & strPRIQText9 &"||" & strPRIQText10 &"||" & strPRIQText11 &"||" & strPRIQText12 &"||" & strPRIQText13 &"||" & strPRIQText14 &"||" & strPRIQText15 &"||" & strPRIQText16 &"||" & strPRIQText17 &"||" & strPRIQText18 &"||" & strPRIQText19 &"||" & sstrPRIQText43 &"||" & trPRIQText20 &"||" & strPRIQText21 &"||" & strPRIQText22 &"||" & strPRIQText23 &"||" & strPRIQText24 &"||" & strPRIQText25 &"||" & strPRIQText26 &"||" & strPRIQText27 &"||" & strPRIQText28 &"||" & strPRIQText29 &"||" & strPRIQText30 &"||" & strPRIQText31 &"||" & strPRIQText44 &"||" & strPRIQText45 &"||" & strPRIQText32 &"||" & strPRIQText33 &"||" & strPRIQText34 &"||" & strPRIQText35 &"||" & strPRIQText36 &"||" & strPRIQText37 &"||" & strPRIQText38 &"||" & strPRIQText39 &"||" & strPRIQText40
strPRIQTextV2 = strPRIQText41 &"||" & strPRIQText42

'Webedit box parameters names and values
strPRIQEditTextV3 = strPRIQText1 &"||" & strPRIQText13 &"||" & strPRIQText14 &"||" & strPRIQText15 &"||" & strPRIQText17 &"||" & strPRIQText29 &"||" & strPRIQText32 &"||" & strPRIQText35
strPRIQEditValuesV3 = strPRIQIPValue1 &"||" & strPRIQIPValue2 &"||" & strPRIQIPValue3 &"||" & strPRIQIPValue4 &"||" & strPRIQIPValue5 &"||" & strPRIQIPValue6 &"||" & strPRIQIPValue7 &"||" & strPRIQIPValue8

'Weblist box paramenters names and values
strListPRIQValueInput1 = strListPRIQValue1 &"||" & strListPRIQValue2 &"||" & strListPRIQValue3 &"||" & strListPRIQValue4 &"||" & strListPRIQValue5 &"||" & strListPRIQValue6 &"||" & strListPRIQValue7 &"||" & strListPRIQValue8 &"||" & strListPRIQValue9 &"||" & strListPRIQValue10 &"||" & strListPRIQValue11 &"||" & strListPRIQValue12 &"||" & strListPRIQValue13 &"||" & strListPRIQValue14 &"||" & strListPRIQValue15
strListPRIQName1 = strPRIQText2 &"||" & strPRIQText4 &"||" & strPRIQText5 &"||" & strPRIQText6 &"||" & strPRIQText8 &"||" & strPRIQText9 &"||" & strPRIQText10 &"||" & strPRIQText11 &"||" & strPRIQText12 &"||" & strPRIQText16 &"||" & strPRIQText18 &"||" & strPRIQText43 &"||" & strPRIQText20 &"||" & strPRIQText23 &"||" & strPRIQText24
strListPRIQValueInput2 = strListPRIQValue16 &"||" & strListPRIQValue17 &"||" & strListPRIQValue18 &"||" & strListPRIQValue19'  &"||" & strListPRIQValue20 &"||" & strListPRIQValue21
strListPRIQName2 = strPRIQText30 &"||" & strPRIQText44 &"||" & strPRIQText33 &"||" & strPRIQText34' &"||" & strPRIQText36 &"||" & strPRIQText37
strListPRIQValueInput3 = strListPRIQValue20' &"||" & strListPRIQValue21 - Verify with Mohammad
strListPRIQName3 = strPRIQText36' &"||" & strPRIQText37 - Verify with Mohammad
strListPRIQValueInput4 = strListPRIQValue26
strListPRIQName4 = strPRIQText40

'Verify all text in the screen and enter values to all text boxes
VerifyWebElementText "Step 52", strBrowser, "DIV", strPRIQTextV1, "No"
VerifyWebElementText "Step 53", strBrowser, "DIV", strPRIQTextV2, "No"

'Parameters for validating Mandatory List box & Text Aread fields
strPRIQPageMandatoryFieldName = strPRIQText1 &"||" & strPRIQText2 & "||" & strPRIQText4 &"||" & strPRIQText5 &"||" & strPRIQText6 &"||" & strPRIQText8 &"||" & strPRIQText9 &"||" & strPRIQText10 &"||" & strPRIQText11 &"||" & strPRIQText12 &"||" & strPRIQText13 &"||" & strPRIQText14 &"||" & strPRIQText15 &"||" & strPRIQText16 &"||" & strPRIQText17 &"||" & strPRIQText18 &"||" & strPRIQText43 &"||" & strPRIQText20 &"||" & strPRIQText21 &"||" & strPRIQText23 &"||" & strPRIQText24 &"||" & strPRIQText25 &"||" & strPRIQText26 &"||" & strPRIQText27 &"||" & strPRIQText28 &"||" & strPRIQText29 &"||" & strPRIQText30 &"||" & strPRIQText31 &"||" & strPRIQText44 &"||" & strPRIQText45 &"||" & strPRIQText32 &"||" & strPRIQText33 &"||" & strPRIQText34 &"||" & strPRIQText35 &"||" & strPRIQText36 &"||" & strPRIQText37 &"||" & strPRIQText38 &"||" & strPRIQText39
VerifyEditListMandatory "Step 53.1", strBrowser, "DIV", "", strPRIQPageMandatoryFieldID, strPRIQPageMandatoryFieldName, strPRIQPageMandatoryLIST, "No"

'Verify all values in drop down list and multi drop down list

strPRIQPageMandatoryListName1 = strPRIQText2 &"||" & strPRIQText4 &"||" & strPRIQText5 &"||" & strPRIQText6
strPRIQPageMandatoryListNameSet2 = strPRIQText8 &"||" & strPRIQText9 &"||" & strPRIQText10
strPRIQPageMandatoryListNameSet3 = strPRIQText11 &"||" & strPRIQText12 &"||" & strPRIQText16 &"||" & strPRIQText18 &"||" & strPRIQText43' &"||" & strPRIQText23
strPRIQPageMandatoryListNameSet4 = strPRIQText23
strPRIQPageMandatoryListNameSet7 = strPRIQText30 &"||" & strPRIQText44
strPRIQPageMandatoryListNameSet8 = strPRIQText33 &"||" & strPRIQText34 &"||" & strPRIQText36
strPRIQPageMandatoryListNameSet10 = strPRIQText40

VerifyWebListValuesGrp "Step 53.2", strBrowser, "SELECT", "0", strPRIQPageMandatoryListName1, strPRIQPageDropDownValues1, "No"
VerifyWebListValuesGrp "Step 53.3", strBrowser, "SELECT", "4", strPRIQPageMandatoryListNameSet2, strPRIQPageDropDownValuesSet2, "No"
VerifyWebListValuesGrp "Step 53.4", strBrowser, "SELECT", "7", strPRIQPageMandatoryListNameSet3, strPRIQPageDropDownValuesSet3, "No"
VerifyWebListValuesGrp "Step 53.5", strBrowser, "SELECT", "13", strPRIQPageMandatoryListNameSet4, strPRIQPageDropDownValuesSet4, "No"
VerifyValuesInMultiList "Step 53.6", strBrowser, "SELECT", 16, strPRIQPageMandatoryListNameSet5, "Does your proposal focus on any of the following populations", "NO"
VerifyValuesInMultiList "Step 53.7", strBrowser, "SELECT", 19, strPRIQPageMandatoryListNameSet6, "Racial or ethnic minorities", "NO"
VerifyWebListValuesGrp "Step 53.8", strBrowser, "SELECT", "21", strPRIQPageMandatoryListNameSet7, strPRIQPageDropDownValuesSet7, "No"
VerifyWebListValuesGrp "Step 53.9", strBrowser, "SELECT", "23", strPRIQPageMandatoryListNameSet8, strPRIQPageDropDownValuesSet8, "No"
VerifyValuesInMultiList "Step 53.10", strBrowser, "SELECT", 28, strPRIQPageMandatoryListNameSet9, "Does your study proposal address a PCORI Research Priority Area", "NO"
VerifyWebListValuesGrp "Step 53.11", strBrowser, "SELECT", "30", strPRIQPageMandatoryListNameSet10, strPRIQPageDropDownValuesSet10, "No"


'Verify and Enter required values in text fields
Set_WebEditGroup "Step 54", strBrowser, "INPUT", "", strPRIQEditValuesV3, strPRIQEditTextV3, strSShot

'verify and select values for all list box
VerifySelectWebListGrp "Step 55", strBrowser, "SELECT", "0", strListPRIQValueInput1, strListPRIQName1, "NO"
VerifySelectWebListGrp "Step 56", strBrowser, "SELECT", "21", strListPRIQValueInput2, strListPRIQName2, "NO"
VerifySelectWebListGrp "Step 57", strBrowser, "SELECT", "25", strListPRIQValueInput3, strListPRIQName3, "NO"
VerifySelectWebListGrp "Step 57.1", strBrowser, "SELECT", "30", strListPRIQValueInput4, strListPRIQName4, "NO"

'Enter values for multi line text box
Set_WebEditMLine "Step 58", strBrowser, "TEXTAREA", strListPRIQValue30, 0, strPRIQText21, "NO"
Set_WebEditMLine "Step 59", strBrowser, "TEXTAREA", strListPRIQValue31, 1, strPRIQText25, "NO"
Set_WebEditMLine "Step 60", strBrowser, "TEXTAREA", strListPRIQValue32, 2, strPRIQText27, "NO"
Set_WebEditMLine "Step 61", strBrowser, "TEXTAREA", strListPRIQValue33, 3, strPRIQText31, "NO"
Set_WebEditMLine "Step 62", strBrowser, "TEXTAREA", strListPRIQValue34, 4, strPRIQText40, "NO"

'Select values for mult list box
VerifySelectMultiList "Step 63", strBrowser, "SELECT", 16, 0, strListPRIQValue27, strPRIQText26, "NO"
VerifySelectMultiList "Step 64", strBrowser, "SELECT", 19, 1, strListPRIQValue28, strPRIQText28, "NO"
VerifySelectMultiList "Step 65", strBrowser, "SELECT", 28, 2, strListPRIQValue29, strPRIQText38, "NO"

'Click on Save and Next button
Click_WebButton "Step 65.1", strBrowser, "INPUT", "0", "Save & Next", "Save & Next", "No"
'#############################################################################################################################################################################
'#############################################################################################################################################################################

'#############################################################################################################################################################################
'################################################################### Project Proposal Questionnaire ##############################################################################
If strBrowser.WebElement("innertext:=At least one key personnel entry is required.", "index:=0").Exist(60) Then
	WriteResult "Step 66", "Verify for Project Personnel Questionnaire page load", "Project Personnel Questionnaire page is loaded successfully", "Pass", "No"
Else
	WriteResult "Step 66", "Verify for Project Personnel Questionnaire page load", "Project Personnel Questionnaire page is not loaded successfully", "Fail", "YES"
End If
strPRPQTextV2 = strPRPQText1 &"||" & strPRPQText2 &"||" & strPRPQText3
VerifyWebElementText "Step 67", strBrowser, "SPAN", strPRPQTextV2, "No"
VerifyWebElementText "Step 67.1", strBrowser, "SPAN", strPRPQText4, "No"
'Verify the table and retrive column names
VerifyWebTable "Step 68", strBrowser, "TABLE", 0, "First NameLast.*", "Header", "YES", "No"
VerifyWebTable "Step 69", strBrowser, "TABLE", 1, "First NameLast.*", "Records", "", "No"
VerifySelectWebListGrp "Step 70", strBrowser, "SELECT", "0", "", "No of entries per page ", "NO"
Set_WebEditGroup "Step 71", strBrowser, "INPUT", "0", "", "Search ", strSShot
VerifyWebButtonGrp "Step 72", strBrowser, "A", "New||Next", strSShot
'VerifyWebElementText "Step 73", strBrowser, "A", "Previous||Next", "No"

'Click on New Button to add new record
Click_WebButton "Step 74", strBrowser, "A", "0", "New", "New", "ON"

'Verify web element text in create new record
'VerifyWebElementText "Step 75", strBrowser, "SPAN", strPRPQText4, "No"
strPRPNewRecordTextV1 = strPRPNewRecordText1 &"||" & strPRPNewRecordText2 &"||" & strPRPNewRecordText3
VerifyWebElementText "Step 76", strBrowser, "P", strPRPNewRecordTextV1, "No"
VerifyWebElementText "Step 77", strBrowser, "DIV", strPRPQText4, "No" ' public const
VerifyWebElementText "Step 78", strBrowser, "DIV", strPRPQText5, "No"

'Verify for mandatory and non mandatory fields
VerifyEditListMandatory "Step 79", strBrowser, "DIV", "", strPRPQPageMandatoryFieldID, strPRPQPageMandatoryFieldName, strPRPQPageMandatoryLIST, "No"

'Verify values in all drop down list and multi list box
VerifyWebListValuesGrp "Step 79.1", strBrowser, "SELECT", "0", strPRPQListBoxName7, strPRPQPageDropDownValues1, "NO"
VerifyWebListValuesGrp "Step 79.2", strBrowser, "SELECT", "4", strPRPQMultiListName9, strPRPQPageDropDownValuesSet9, "No"

'Enter values in all required fields
Set_WebEditGroup "Step 80", strBrowser, "INPUT", "0", strPRPQTextBoxValue6, strPRPQTextBoxName6, "NO"
VerifySelectWebListGrp "Step 81", strBrowser, "SELECT", "0", strPRPQListBoxValue7, strPRPQListBoxName7, "NO"
Set_WebEditMLine "Step 82", strBrowser, "TEXTAREA", strPRPQMultiLineValue8, 3, strPRPQMultiLineName8, "NO"
VerifySelectMultiList "Step 83", strBrowser, "SELECT", 4, 5, strPRPQMultiListValue9, strPRPQMultiListName9, "NO"

'Click on Yes and then click on Save button
WebElementClick "Step 84", strBrowser, "LABEL", "Yes", "NO"
'VerifyWebButtonGrp "Step 85", strBrowser, "INPUT", "Save||Cancel", "NO" Update required
Click_WebButton "Step 86", strBrowser, "INPUT", "0", "Save", "Save", "NO"

'Verify New Record is added
SearchWebTableRow "Step 87", strBrowser, "TABLE", "1", "First Name.*", strPRPQTextBoxValue8, strPRPQTextBoxValue9, "NO"
Click_WebButton "Step 88", strBrowser, "A", "0", "Next", "Next", "NO"
'#############################################################################################################################################################################
'#############################################################################################################################################################################

'#############################################################################################################################################################################
'################################################################### Templates and Uploads Questionnaire ########################################################################
'Verify all fields in Templates and UploadsTab
strTUQTextV1 =  strTUQText1 &"||" & strTUQText6
strTUQTextV2 =  strTUQText2 &"||" & strTUQText3 &"||" & strTUQText4 &"||" & strTUQText5 

VerifyWebElementText "Step 90", strBrowser, "DIV", strTUQTextV1, "No"
VerifyWebElementText "Step 91", strBrowser, "P", strTUQTextV2, "No"
VerifyWebButtonGrp "Step 92", strBrowser, "LABEL", "Choose file", "No"
VerifyWebButtonGrp "Step 93", strBrowser, "INPUT", "Upload||Save||Clear Changes", "No"
Set strBrowser = Browser("name:=Sfdc Page.*").Page("title:=Sfdc Page").Frame("title:=.*PCORI Online.*")
strBrowser.Highlight
Dim intX, intY, objDeviceReplay, intWidth, intHeight
If (Not strBrowser.WebButton("index:=3").Exist) Then
	Reporter.ReportEvent micFail, "LeftClick", "The object does not exist"		
Else	
	intX = strBrowser.WebButton("index:=3").GetROProperty("abs_x")
	intY = strBrowser.WebButton("index:=3").GetROProperty("abs_y")
	Set objDeviceReplay = CreateObject("Mercury.DeviceReplay")
	objDeviceReplay.MouseClick intX + 10, intY + 10, micLeftBtn
	Set objDeviceReplay = Nothing
End If

strPDFPath = "C:\Users\SatheeshKumarDuraisa\OneDrive - Patient-Centered Outcomes Research Institute\SFAutomation\Documents\UploadDocuments"
strPDFFileName = "Existing Report.pdf"
If Window("regexpwndtitle:= Google Chrome").Dialog("regexpwndtitle:=Open").WinEdit("attached text:=File &name:", "index:=0").Exist(10) Then
	WriteResult "Step 94", "Select File for upload", strPDFFileName & " is selected for upload", "Pass", "No"
	Window("regexpwndtitle:= Google Chrome").Dialog("regexpwndtitle:=Open").WinEdit("attached text:=File &name:", "index:=0").Type strPDFPath & "\" & strPDFFileName
	Window("regexpwndtitle:= Google Chrome").Dialog("regexpwndtitle:=Open").WinButton("regexpwndtitle:=&Open").Click
End If

Set strBrowser = Browser("name:=Sfdc Page.*").Page("title:=Sfdc Page").Frame("title:=.*PCORI Online.*")
strBrowser.Highlight
If (Not strBrowser.WebButton("index:=5").Exist) Then
	Reporter.ReportEvent micFail, "LeftClick", "The object does not exist"		
Else	
	intX = strBrowser.WebButton("index:=5").GetROProperty("abs_x")
	intY = strBrowser.WebButton("index:=5").GetROProperty("abs_y")
	Set objDeviceReplay = CreateObject("Mercury.DeviceReplay")
	objDeviceReplay.MouseClick intX + 10, intY + 10, micLeftBtn
	Set objDeviceReplay = Nothing
End If
Set strBrowser = Browser("name:=Sfdc Page.*").Page("title:=Sfdc Page").Frame("title:=.*PCORI Online.*")
If strBrowser.WebElement("innertext:=File Upload Successful", "index:=0").Exist(60) Then
	Wait(2)
	WriteResult "Step 96", "Verify Upload Status", "Upload Status is --> File Upload Successful", "Pass", "No"
End If
'Click on Save button and then click on Review/Submit button
Click_WebButton "Step 97", strBrowser, "INPUT", "0", "Save", "Save", "No"
Click_WebButton "Step 98", strBrowser, "INPUT", "0", "Review/Submit", "Review/Submit", "No"

'#############################################################################################################################################################################
'#############################################################################################################################################################################

'#############################################################################################################################################################################
'################################################################### Verifying Values in CI Section ########################################################################

strBPStudies = "Satheesh Duraisamy||Merry||merry123@gmail.com||Agila Somasundaram||Abigail Keatts||Aleksandra Modrow||Agila Somasundaram||One_Smoke Test Account_DB||Washington DC||Automation"
Set strBrowserCI = Browser("name:=Sfdc Page.*").Page("title:=Sfdc Page").Frame("title:=.*PCORI Online.*").WebTable("column names:=Principal Investigator.*", "index:=0")
VerifyValuesInContactInfoSection "Step 100", strBrowserCI, strBPStudies, "No"

'#############################################################################################################################################################################
'#############################################################################################################################################################################

'#############################################################################################################################################################################
'################################################################### Verifying Values in Pre Screen Section ########################################################################
'Verifying Pre Screen Questions values in submit page
Set strBrowserPS = Browser("name:=Sfdc Page.*").Page("title:=Sfdc Page").Frame("title:=.*PCORI Online.*").WebTable("column names:=Creation of a decision aid or.*", "index:=0")
VerifyValuesInContactInfoSection "Step 101", strBrowserPS, strPreScreenQ4, "No"

'#############################################################################################################################################################################
'#############################################################################################################################################################################

'#############################################################################################################################################################################
'################################################################### Verifying Values in Re-Submission Section ######################################################################
Set strBrowserRS = Browser("name:=Sfdc Page.*").Page("title:=Sfdc Page").Frame("title:=.*PCORI Online.*").WebTable("column names:=Have you submitted this project to PCORI before as an LOI.*", "index:=0")
VerifyValuesInContactInfoSection "Step 102", strBrowserRS, strRSApplicationVerify, "No"

'#############################################################################################################################################################################
'#############################################################################################################################################################################

'#############################################################################################################################################################################
'################################################################### Verifying Values in PI Information Section #####################################################################
Set strBrowserPIS = Browser("name:=Sfdc Page.*").Page("title:=Sfdc Page").Frame("title:=.*PCORI Online.*").WebTable("column names:=PI Work Telephone Number.*", "index:=0")
VerifyValuesInPIInformationSection "Step 103", strBrowserPIS, strListValuePI7, strListValuePI8, strListValuePI9, strPIApplicationVerify, strValuePI1, "No"

'#############################################################################################################################################################################
'#############################################################################################################################################################################

'#############################################################################################################################################################################
'################################################################### Verifying Values in Project Information Section ##################################################################
Set strBrowserPRS = Browser("name:=Sfdc Page.*").Page("title:=Sfdc Page").Frame("title:=.*PCORI Online.*").WebTable("column names:=Project Title.*", "index:=0")
strPRIQIPReviewPage = strPRIQIPValue1 & "||" & strListPRIQValue1 & "||" & strListPRIQValue2 & "||" & strListPRIQValue3 & "||" & strListPRIQValue4 & "||" & strListPRIQValue5 & "||" & strListPRIQValue6 & "||" & strListPRIQValue7 & "||" & strListPRIQValue8 & "||" & strListPRIQValue9 & "||" & strPRIQIPValue2 & "||" & strPRIQIPValue3 & "||" & strPRIQIPValue4 & "||" & strListPRIQValue10 & "||" & strPRIQIPValue5 & "||" & strListPRIQValue11 & "||" & strListPRIQValue12 & "||" & strListPRIQValue13 & "||" & strListPRIQValue30 & "||" & strListPRIQValue14 & "||" & strListPRIQValue15 & "||" & strListPRIQValue31 & "||" & "" & "||" & strListPRIQValue32 & "||" & "" & "||" & strPRIQIPValue6 & "||" & strListPRIQValue16 & "||" & strListPRIQValue33 & "||" & strListPRIQValue17 & "||" & strListPRIQValue34 & "||" & strPRIQIPValue7 & "||" & strListPRIQValue18 & "||" & strListPRIQValue19 & "||" & strPRIQIPValue8 & "||" & strListPRIQValue20 & "||" & strListPRIQValue35 & "||" & strListPRIQValue29 & "||" & strListPRIQValue26
VerifyValuesInProjectInfoSection "Step 104", strBrowser, strListPRIQValue27,strListPRIQValue28, strPRIQIPReviewPage, "No"

'Verify web element text in Project Personnel Section
strPIQTextV2 = strPRPQText1 &"||" & strPRPQText2 &"||" & strPRPQText3
VerifyWebElementText "Step 104.1", strBrowser, "SPAN", strPIQTextV2, "No"
VerifyWebElementText "Step 104.2", strBrowser, "SPAN", strPRPQText4, "No"

'Verify table for the added record
SearchWebTableRow "Step 104.4", strBrowser, "TABLE", "0", "First Name.*", strPRPQTextBoxValue8, strPRPQTextBoxValue9, "NO"
'Click on Submit Button
'Click_WebButton "Step 105", strBrowser, "INPUT", "1", "Submit", "Submit", "No"
strBrowser.WebButton("name:=Submit.*", "html tag:=INPUT", "index:=0").Click
WriteResult "Step 105", "Click on Submit button", "Submit button is clicked", "Pass", strSShot
'Click on pop up confirmation window
If strBrowser.WebElement("innertext:=Are you sure you want to submit.*", "index:=0").Exist(30) Then
	Click_WebButton "Step 106", strBrowser, "BUTTON", "0", "OK", "OK", "No"
End If
'#############################################################################################################################################################################
'#############################################################################################################################################################################

'#############################################################################################################################################################################
'################################################################### Verifying Values in LOI Table ################################################################################
'Search for the created LOI and verify the status
Set strBrowser = Browser("name:=Sfdc Page.*").Page("title:=Sfdc Page").Frame("title:=.*PCORI Online.*")
If strBrowser.WebTable("html tag:=TABLE", "innertext:=EditView.*", "index:=0").Exist(60) Then
	Wait(2)
	ClickLink "Step 107", strBrowser, "0", "LOIs.*", "LOI Section", "No"
	Browser("name:=Sfdc Page.*").Sync
End If
'Verify the table is loaded and then check for the created LOI
Wait(2)
RetriveValueWebTable "Step 108", strBrowser, "TABLE", "0", "EditView-LOI.*", strPRIQIPValue1, strProjectLead, "Yes"
'#############################################################################################################################################################################
'#############################################################################################################################################################################

'#############################################################################################################################################################################
'################################################################### Login as Internal user ################################################################################
strInternalURL = DataTable.Value("strInternalURL", DtGlobalSheet)
strBrowserType = DataTable.Value("strBrowserType", dtGlobalSheet)
strInternalUser = "satheesh@vinmarinc.com"
strInternalPassword = "675c8a8ffd48e37f5b98f401bebff9e7b8161f05b2e986fc7d0994054364f62daa8b54f2"
'Set strSFBrowser = Browser("openurl:=https://pcori--qapocdev.sandbox.my.salesforce.com", "creationtime:=1").Page("title:=Login |.*")
Set strSFBrowser = Browser("creationtime:=1").Page("title:=Login |.*")

Open_SalesForce_Application strBrowserType, strInternalURL
Wait(5)
Set_WebEdit "Step 110", strSFBrowser, "username", strInternalUser, "User Name", "Yes"
Set_WebEditSecure "Step 111", strSFBrowser, "pw", strInternalPassword, "Password", "Yes"
Click_WebButton "Step 112", strSFBrowser, "INPUT", "0", "Log In to Sandbox", "Login Button", "No"
Msgbox "Click on Approve in Salesforce Authenticator Mobile app and then click on OK in this pop up message"
'Click on LOI Tab
Set strSFBrowser = Browser("creationtime:=1").Page("title:=Home |.*")
strSFBrowser.Highlight
'ClickLink "Step 113", strSFBrowser, "0", "LOIs", "LOI Tab", strSShot
'Search LOI from Main Search
strLOINumber = Environment.Value("intLOINumber")
Click_SFLWebButton "Step 114", strSFBrowser, "Search...", "Click on Search", "No"
Set_SFLWebEdit "Step 115", strSFBrowser, "Search\.\.\.", strLOINumber, "Search", "No"
Set WshShell = CreateObject("WScript.Shell")
Set oDesc = Description.Create
oDesc("micclass").value = "SFLEdit"
oDesc("html tag").value = "INPUT" 
oDesc("placeholder").value = "Search\.\.\."
strSFBrowser.SFLEdit(oDesc).Click
WshShell.SendKeys "{ENTER}"

'Search for the LOI and activate the row
ClcikWebTableRow "Step 116", strSFBrowser, "TABLE", "0", "Item Number.*", "LOI Number", "3", strVerifyValue, "No"

'Verify LOI Status, Number, External Status, Name
Set strSFBrowser = Browser("creationtime:=1").Page("title:=.* | LOI |.*")
VerifyStatus "Step 117", strSFBrowser, "DIV", "3", "LOI Numbe.*", "LOI Number", strLOINumber, "YES"
VerifyStatus "Step 118", strSFBrowser, "DIV", "0", "LOI Status.*", "LOI Status", strLOIStatus, "N0"
VerifyStatus "Step 119", strSFBrowser, "DIV", "0", "External Status.*", "External LOI Status", strExternalStatus, "N0"
VerifyStatus "Step 120", strSFBrowser, "DIV", "3", "Name.*", "Name", strProjectLead, "N0"

'Search for Science PO Admin
Click_SFLWebButton "Step 121", strSFBrowser, "Search...", "Click on Search", "No"
Set_SFLWebEdit "Step 122", strSFBrowser, "Search\.\.\.", SciencePOAdmin, "Search", "No"
Set WshShell = CreateObject("WScript.Shell")
Set oDesc = Description.Create
oDesc("micclass").value = "SFLEdit"
oDesc("html tag").value = "INPUT" 
oDesc("placeholder").value = "Search\.\.\."
strSFBrowser.SFLEdit(oDesc).Click
WshShell.SendKeys "{ENTER}"

'#############################################################################################################################################################################
'#############################################################################################################################################################################
strScriptName = Environment.Value("TestName")
Environment.Value("TName") = strScriptName
GetCurrentDate
CurrDate (Time)
CreateResultFolder
InitResultFile
strVerifyValue = "43229"

Set strSFBrowser = Browser("creationtime:=1").Page("title:=.* | LOI |.*")

'WshShell.AppActivate ""
'Browser("Login | Salesforce").Page("Satheesh Duraisamy | LOI").WebElement("LOI StatusNew LOIEdit").Click
strSFBrowser.Highlight
Click_SFLWebButton "Step 121", strSFBrowser, "Search...", "Click on Search", "No"
Set_SFLWebEdit "Step 122", strSFBrowser, "Search\.\.\.", SciencePOAdmin, "Search", "No"
Set WshShell = CreateObject("WScript.Shell")
Set oDesc = Description.Create
oDesc("micclass").value = "SFLEdit"
oDesc("html tag").value = "INPUT" 
oDesc("placeholder").value = "Search\.\.\."
strSFBrowser.SFLEdit(oDesc).Click
WshShell.SendKeys "{ENTER}"
'CaptureScreen
'strColumnValue1 = strPRPQTextBoxValue8
'strColumnValue2 = strPRPQTextBoxValue9
Set strBrowser = Browser("name:=Sfdc Page.*").Page("title:=Sfdc Page").Frame("title:=.*PCORI Online.*")
'Verify the table is loaded and then check for the created LOI
RetriveValueWebTable "Step 108", strBrowser, "TABLE", "0", "EditView-LOI.*", strPRIQIPValue1, strProjectLead, strSShot
'#################################################################################################################
'Function Name: VerifyText
'Description: This function is used to Verify the existance of text (webelement)  
'Parameters: strStepNo, strBrowser, strObjProperty, strObjName, strSShot
'Created By: Satheesh Kumar Duraisamy
'Created Date: 
'Modified By: NA
'Modified Date: NA
'Modification History: NA
'Comments: Excel Report implemented
'#################################################################################################################
Public Function VerifyStatus (strStepNo, strBrowser, strTag, strIndex, strObjProperty, strObjName, strValueVerify, strSShot)
	
	Dim Brwser
	Set Brwser = strBrowser
    
	Set oDesc = Description.Create
	oDesc("micclass").value = "WebElement"
	oDesc("html tag").value = strTag
	oDesc("index").value = strIndex
	oDesc("innertext").value = strObjProperty

	If Brwser.WebElement(oDesc).Exist(1) Then
		Brwser.WebElement(oDesc).Click
		Brwser.WebElement(oDesc).Highlight
		strAppValue = Brwser.WebElement(oDesc).GetROProperty("innertext")
		If Instr(1, strAppValue, strValueVerify) > 0 Then
			WriteResult strStepNo, "Verify for " & strObjName & ".", strObjName & " - " & strValueVerify, "Pass", strSShot
		Else
			WriteResult strStepNo, "Verify for " & strObjName & ".", strObjName & " - " & strValueVerify, "Fail", strSShot
		End If
	Else
		WriteResult strStepNo, "Verify for " & strObjName, strObjName & " was not found", "Fail", "YES"
	End If
End Function

'Set_WebEditMLine "Step 51", strBrowser, "TEXTAREA", "Please describe Other degree - test", 0, "Other Organizations", "NO"

'Posted On 14 Nov 2024 Deadline 31 Mar 2025
'Browser("name:=Sfdc Page.*").Page("title:=Sfdc Page").Frame("title:=.*PCORI Online.*").WebButton("name:=Choose file").Highlight
'Browser("name:=Sfdc Page.*").Page("title:=Sfdc Page").Frame("title:=.*PCORI Online.*").WebButton("name:=Choose file").Click
'################################################################### Step  #####################################################################################################
'Click 'New User?' to create a new account
Set strUIAObject = UIAObject("name:=Login - Google Chrome.*").UIAObject("name:=Login")
Click_UIWebButton strUIAObject, "Log In", "Log In"

'################################################################### Step 5 #####################################################################################################
'Step5: Fill out all required fields and 
'check the "In voluntarily providing this information, you agree to abide by our website privacy policy and terms of use."  Checkbox and 
'click "Join PCORI Online"
Set strBrowser = Browser("name:=LightningSelfRegisterPage.*").page("title:=LightningSelfRegisterPage.*")
VerifyEditBox strBrowser, "cmntyfrstnm", "First Name"
strVerifyTextbox = Environment.Value("VerifyTextbox")
If strVerifyTextbox = "Success" Then
	Reporter.ReportEvent micPass, "Verify New User page is displayed", "New User page is displayed, Please enter all the required field and click on Join PCORI Online"
Else
	Reporter.ReportEvent micFail, "Verify New User page is displayed", "New User page is not displayed, execution terminated"
	ExitTest
End If

strFirstName = DataTable.Value("strFirstName", dtGlobalSheet)
strLastName = DataTable.Value("strLastName", dtGlobalSheet)
strEmail = DataTable.Value("strEmail", dtGlobalSheet)
strConfirmEmail = DataTable.Value("strConfirmEmail", dtGlobalSheet)
strNewPassword = DataTable.Value("strNewPassword", dtGlobalSheet)
strConfirmNewPassword = DataTable.Value("strConfirmNewPassword", dtGlobalSheet)

'Enver Values for all mandatory fields FirstName, LastName, Email, Confirm Email, Password, Confirm Password
Set_WebEdit strBrowser, "cmntyfrstnm", strFirstName, "First Name"
Set_WebEdit strBrowser, "cmntylastnm", strLastName, "Last Name"
Set_WebEdit strBrowser, "cmntyemail", strEmail, "EMail ID"
Set_WebEdit strBrowser, "cmntycnfrmemail", strConfirmEmail, "Confirm Email ID"
SetSecure_WebEdit strBrowser, "cmntypswd", strNewPassword, "New Password"
SetSecure_WebEdit strBrowser, "cmntycnfrmpswd", strConfirmNewPassword, "Confirm New Password"

'Select Voluntarily Check box
Click_CheckBoxElement strBrowser, "slds-checkbox_faux", "voluntarily check box"

'Click on I am not Robot Text Box
Set strBrowser = Browser("name:=LightningSelfRegisterPage.*").page("title:=LightningSelfRegisterPage.*").Frame("title:=reCAPTCHA", "html tag:=IFRAME", "name:=a.*")
Click_CheckBox strBrowser, "I'm not a robot", "I am not a robot"

'Click on Join PCORI Online button
'################################################################### Step 6 #####################################################################################################
'Step 6: Verify below "Contact Information" text present:  
'"Welcome to the PCORI Online. Please provide some basic information about yourself before proceeding to the Online homepage."
strApplyingFor = DataTable.Value("strApplyingFor", dtGlobalSheet)
strSalutation = DataTable.Value("strSalutation", dtGlobalSheet)
strGender = DataTable.Value("strGender", dtGlobalSheet)
StrHispanicOrLatino = DataTable.Value("StrHispanicOrLatino", dtGlobalSheet)
strRace = DataTable.Value("strRace", dtGlobalSheet)
strYearOfBirth = DataTable.Value("strYearOfBirth", dtGlobalSheet)
strCommunities = DataTable.Value("strCommunities", dtGlobalSheet)
strInvolvedPCORI = DataTable.Value("strInvolvedPCORI", dtGlobalSheet)
strFederalEmployee = DataTable.Value("strFederalEmployee", dtGlobalSheet)
strPositionTitle = DataTable.Value("strPositionTitle", dtGlobalSheet)
strDepartment = DataTable.Value("strDepartment", dtGlobalSheet)
strEmployerName = DataTable.Value("strEmployerName", dtGlobalSheet)
strEmployerFound = DataTable.Value("strEmployerFound", dtGlobalSheet)

strPhone = DataTable.Value("strPhone", dtGlobalSheet)
strStreet = DataTable.Value("strStreet", dtGlobalSheet)
strCity = DataTable.Value("strCity", dtGlobalSheet)
srtStateProvince = DataTable.Value("srtStateProvince", dtGlobalSheet)
strCountry = DataTable.Value("strCountry", dtGlobalSheet)
srtZipPostalCode = DataTable.Value("srtZipPostalCode", dtGlobalSheet)

Set strBrowser = Browser("name:=ContactInformationPage.*").Page("title:=ContactInformationPage")

Click_RadioBtnElement strBrowser, strApplyingFor, strApplyingFor
'Mailing Address
Set_WebEdit strBrowser, "XXX-XXX-XXXX", strPhone, "Phone"
Set_WebEdit strBrowser, "Street", strStreet, "Street"
Set_WebEdit strBrowser, "City", strCity, "City"
Select_WebList strBrowser, "USA;.*", strCountry, "Country"
Select_WebList strBrowser, "Alaska;.*", srtStateProvince, "State OR Province"
Set_WebEdit strBrowser, "Zip Code", srtZipPostalCode, "Postal Code"
Click_RadioBtnElement strBrowser, strSalutation, strSalutation
'Demographic Information
Click_RadioBtnElement strBrowser, strGender, strGender
Click_RadioBtnElement strBrowser, StrHispanicOrLatino, "Hispanic or Latino"
Click_RadioBtnElement strBrowser, strRace, strRace
Select_WebList strBrowser, "1900;.*", strYearOfBirth, "Date of Year"
Click_RadioBtnElement strBrowser, strCommunities, strCommunities
Click_CheckBox strBrowser, strInvolvedPCORI, strInvolvedPCORI
'Employer Information
Click_RadioBtnElement strBrowser, strFederalEmployee, "Federal Employee"
'Add Position, title and department, Employee lookup Search
Set_WebEdit strBrowser, "Position/Title", strPositionTitle, "Position"
Set_WebEdit strBrowser, "Department", strDepartment, "Department"
Set_WebEdit strBrowser, "Type to Search", strEmployerName, "Lookup Employee Name Search"

Select_WebRadioGroup strBrowser, "input-8", "Employee LookUp"



'################################################################## Step 12 Completed ###################################################################################3333
'''strScriptName = Environment.Value("TestName")
'''Environment.Value("TName") = strScriptName
'''GetCurrentDate
'''CurrDate (Time)
'''CreateResultFolder
'''InitResultFile
'''CaptureScreen
'''strBPStudies = "Salon||Merry||merry123@gmail.com||Agila Somasundaram||Abigail Keatts||Aleksandra Modrow||Anjana||One_Smoke Test Account_DB||Washington DC||Automation"
'''strFieldName = "Principal Investigator \(Contact\)||Dual Principal Investigator Name||Dual Principal Investigator Email||Administrative Official||PI Designee 1||PI Designee 2||Financial Officer||Organization||Congressional District||Department"
'''strContactInfoText = "For instructions on using our system, click here to access our PCORI Portal User Guide.||To view additional information about the PCORI submission process, click here to view our FAQ page.||- A Principal Investigator \(PI\) and an Administrative Official \(AO\) with a valid email address are required to submit your LOI to PCORI.||- To assign a user, click the lookup icon and start to type their name. If the user does not exist in our system, they must register in PCORI Online by clicking “New User.”||- User login information from the previous PCORI Online were not migrated to the new PCORI Online.||- The AO and the PI cannot be the same individual.||- Individuals assigned at the “Contact Information” tab will have access to the LOI and Application.||- To find your Congressional District please click here.||- Fields marked with \(*\) are required."
'''strContactInfoText1 = "Click 'Save & Next' to continue to the next tab. Otherwise you could receive an error message."
''''Verify Text boxes
'''Set strBrowser = Browser("name:=Sfdc Page.*").Page("title:=Sfdc Page")
''''Verify contact info text
'''VerifyWebElementText "Step 12", strBrowser, "P", strContactInfoText, "No"
'''VerifyWebElementText "Step 13", strBrowser, "DIV", strContactInfoText1, "No"
'''VerifyWebElementText "Step 14", strBrowser, "DIV", strFieldName, "No"
''''Enter values in the text boxes
'''Set_WebEditGrp "Step 15", strBrowser, strBPStudies, "YES"

'''Function Set_WebEditGrp (strStepNo, strBrowser, strValue, strSShot)
'''
'''	Set oDesc = Description.Create()
'''	oDesc("micclass").value = "WebEdit"
'''	oDesc("visible").value = true 
'''	Set obj = strBrowser.ChildObjects(oDesc)
'''	
'''	If Instr(1, strBPStudies, "||") > 0 Then
'''		strBPStudies = Split(strBPStudies, "||")
'''		strBPSCount = UBound(strBPStudies)
'''	End If
'''
'''	for i = 0 To strBPSCount
''''		obj(i).Set strBPStudies(i)
'''		If obj(i).Exist(1) Then
'''			obj(i).Highlight
'''			obj(i).Set strBPStudies(i)
'''			WriteResult strStepNo & "." & i, "Enter " & strBPStudies(i) & " in " & strObjName(i) & " field", strValue & " was entered in " & strObjName(i) & " field", "Pass", strSShot
''''			Reporter.ReportEvent micPass, "Enter " & strBPStudies(i) & " in " & strFieldName(i) &" field", strBPStudies(i) & " is entered in the text field"
'''		Else
''''			Reporter.ReportEvent micFail, "Enter " & strBPStudies(i) & " in text field", strBPStudies(i) & " isnot  entered in the text field"
'''			WriteResult strStepNo & "." & i, "Enter " & strValue & " in " & strObjName(i) & " field",strObjName(i) & " field was NOT found", "Fail", "YES"
'''		End If
'''	Next
'''End Function @@ script infofile_;_ZIP::ssf15.xml_;_
 
 
Set Brwser = strBrowser
'strObjProperty = "Eugene Washington PCORI Engagement Award Program"
'Set oDesc = Description.Create
'oDesc("micclass").value = "WebElement"
'oDesc("innertext").value = strObjProperty
'oDesc("html tag").value = "B" 

'strBPStudies = "Salon||Merry||merry123@gmail.com||Agila Somasundaram||Abigail Keatts||Aleksandra Modrow||Anjana||One_Smoke Test Account_DB||Washington DC||Automation"
'Set strBrowser = Browser("name:=Sfdc Page.*").Page("title:=Sfdc Page")


	

Wait(1)


'###########################################################################################################################################


 @@ script infofile_;_ZIP::ssf20.xml_;_
