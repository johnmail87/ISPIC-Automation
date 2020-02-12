'#####################################################################################

   ' Name: ISPIC_Create_Opportinity
   'Purpose : 
   'Created Date: 
   'Author: 
   'Last modified Date:
 
 '#######################################################################################

 'On Error Resume Next

'Login to ISPIC application

Proc_ISPIC_Login()

If Browser("BRDashboard").Page("PGDashboard").WebElement("lblModules").Exist(10) Then
	Browser("BRDashboard").Page("PGDashboard").WebElement("lblModules").Click
End  If

If Browser("BRDashboard").Page("PGDashboard").Link("lnkClien/Account").Exist(5) Then
	Browser("BRDashboard").Page("PGDashboard").Link("lnkClien/Account").Click
End  If

strAccountName = Trim(ReadSingleCloumn("AccountName"))

If strAccountName = "" Then
	RunAction "All Accounts [ISPIC_COMP_Accounts]", oneIteration, "NEWNOTVERIFY"

Else
	If Browser("BRClientModule").Page("PGClientModule").WebEdit("txtSearch").Exist(10) Then
		Browser("BRClientModule").Page("PGClientModule").WebEdit("txtSearch").Set strAccountName
		Browser("BRClientModule").Page("PGClientModule").WebEdit("txtSearch").DoubleClick
		Browser("BRClientModule").Page("PGClientModule").WebTable("tblAccounts").WaitProperty "rows",2,30
		If Browser("BRClientModule").Page("PGClientModule").WebTable("tblAccounts").Exist(20) Then
			Wait 8
			SET strObjWebElement = 	Browser("BRClientModule").Page("PGClientModule").WebTable("tblAccounts").ChildItem(2,7,"Image",0)
			checkAccountName =  Trim(Browser("BRClientModule").Page("PGClientModule").WebTable("tblAccounts").GetCellData(2,3))	
	 		strObjWebElement.Click
			Datatable("AccountName") = checkAccountName 
			strAccountName = checkAccountName
			Wait 5
	 		Browser("BRClientModule").Sync
			
		End If

	End If

End If

strOpportunityName = "UFT"&GenerateRandomString("4")&GenerateRandomNumber("4")

strOpportunityNo = STAGE_Identify_and_Requirements_Definition(strOpportunityName)

strCurrency = Datatable("ContractCurrency")
strContractPeriod = Datatable("ContractPeriod")

boolTechnicalDesign = Datatable("EnableTechnicalDesign")
strProductName = Datatable("ProductName")
STAGE_Qualification strCurrency,strContractPeriod,strAccountName,strOpportunityNo,boolTechnicalDesign,strProductName

STAGE_Proposal(strOpportunityNo)
STAGE_Proposal_Internal_Approval(strOpportunityNo)
STAGE_Proposal_Evaluation(strOpportunityNo)
STAGE_Finalist(strOpportunityNo)
STAGE_Deal_Won_Vetting(strOpportunityNo)
STAGE_Deal_Won_CreateWorkOrder strOpportunityNo,strAccountName 
Configure_Service()
Project_Management strOpportunityNo,strAccountName,strProductName


Proc_ISPIC_Logout()


If Err.Number <> 0 Then
	Reporter.ReportEvent micFail, "Opportuity Creation Failure Occur","Error Description : " & Err.description
	err.clear
	ExitTest 
End If





















  










































