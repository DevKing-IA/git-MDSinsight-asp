<!--#include file="../inc/SubsAndFuncs.asp"-->
<!--#include file="../inc/mail.asp"-->
<!--#include file="../inc/InSightFuncs.asp"-->
<!--#include file="../inc/InSightFuncs_Prospecting.asp"-->
<%

txtSuffix = Request.Form("txtSuffix")
txtFirstName = Request.Form("txtFirstName")
txtLastName = Request.Form("txtLastName")
txtTitle = Request.Form("txtTitle")
txtCompanyName = Request.Form("txtCompanyName")
txtAddressLine1 = Request.Form("txtAddressLine1")
txtAddressLine2 = Request.Form("txtAddressLine2")
txtCity = Request.Form("txtCity")
txtState = Request.Form("txtState")
txtZipCode = Request.Form("txtZipCode")
txtCountry = Request.Form("txtCountry")
txtEmailAddress = Request.Form("txtEmailAddress")
txtWebsiteURL = Request.Form("txtWebsiteURL")
txtPhoneNumber = Request.Form("txtPhoneNumber")
txtPhoneNumberExt = Request.Form("txtPhoneNumberExt")
txtCellPhoneNumber = Request.Form("txtCellPhoneNumber")
txtFaxNumber = Request.Form("txtFaxNumber")
txtIndustry = Request.Form("txtIndustry")

radStage = Request.Form("radStage")
txtStageNotes = Request.Form("txtStageNotes")

txtOwner = Request.Form("selProspectOwner")
chkDoNotEmailNewOwner = Request.Form("chkDoNotEmailNewOwner")

If (chkDoNotEmailNewOwner <> "" AND chkDoNotEmailNewOwner = "on") Then 
	chkDoNotEmailNewOwner = 1 
	sendEmailFlag = 0
Else 
	chkDoNotEmailNewOwner = 0
	sendEmailFlag = 1
End If

If (cint(Session("UserNo")) <> cint(selNewProspectOwner)) Then
	If sendEmailFlag <> True Then
		UserNoForCalendarUpdate = selNewProspectOwner
	Else ' Send email option is turned on
		UserNoForCalendarUpdate = 0	
	End If	
Else
	UserNoForCalendarUpdate = Session("UserNO")
End If


txtComments = Request.Form("txtComments")
txtCurrentOffering = Request.Form("txtCurrentOffering") 

txtNextActivity = Request.Form("selProspectNextActivity")
txtNextActivityDueDate = Request.Form("txtNextActivityDueDate")
txtNextActivityNotes = Request.Form("txtNextActivityNotes")

'************************************************************************************************************
'************************************************************************************************************
'************************************************************************************************************
txtMeetingLocation = Replace(Request.Form("txtMeetingLocation"),"'","''")
selAppointmentDuration = Request.Form("selAppointmentDuration")
selMeetingDuration = Request.Form("selMeetingDuration")
ProspectNewActivity = GetActivityByNum(txtNextActivity)
ProspectApptOrMeeting = GetActivityApptOrMeetingByNum(txtNextActivity)
'************************************************************************************************************
'************************************************************************************************************
'************************************************************************************************************

txtProjectedGPSpend = Request.Form("txtProjectedGPSpend")
txtNumEmployees = Request.Form("txtNumEmployees")
txtNumPantries = Request.Form("txtNumPantries")
txtLeaseExpirationDate= Request.Form("txtLeaseExpirationDate")
txtContractExpirationDate= Request.Form("txtContractExpirationDate")
txtPrimaryCompetitor = Request.Form("txtPrimaryCompetitor")
txtTelemarketerUserNo = Request.Form("txtTelemarketerUserNo")
txtLeadSource = Request.Form("txtLeadSource")

chkBottledWater = Request.Form("chkBottledWater")
chkFilteredWater = Request.Form("chkFilteredWater")
chkOCS = Request.Form("chkOCS")
chkOCS_Supply = Request.Form("chkOCS_Supply")
chkOfficeSupplies = Request.Form("chkOfficeSupplies")
chkVending = Request.Form("chkVending")
chkMicroMarket = Request.Form("chkMicroMarket")
chkPantry = Request.Form("chkPantry")

txtFormerCustomerNumber = Request.Form("txtFormerCustomerNumber")
txtFormerCustomerCancelDate = Request.Form("txtFormerCustomerCancelDate")

'*******************************************************************************************************************
'FIX ANY ENTRIES THAT MAY CONTAIN SINGLE QUOTES FOR SQL INSERT
'*******************************************************************************************************************

txtFirstName = Replace(txtFirstName,"'","''")
txtLastName = Replace(txtLastName,"'","''")
txtTitle = Replace(txtTitle,"'","''")
txtCompanyName = Replace(txtCompanyName,"'","''")
txtAddressLine1 = Replace(txtAddressLine1,"'","''")
txtAddressLine2 = Replace(txtAddressLine2,"'","''")
txtCity = Replace(txtCity,"'","''")
txtEmailAddress = Replace(txtEmailAddress,"'","''")
txtIndustry = Replace(txtIndustry,"'","''")
txtStageNotes = Replace(txtStageNotes,"'","''")
txtComments = Replace(txtComments,"'","''")
txtNextActivityNotes = Replace(txtNextActivityNotes,"'","''")
txtMeetingLocation = Replace(txtMeetingLocation,"'","''")
txtState = Replace(txtState,"'","''")
txtCountry = Replace(txtCountry,"'","''")
txtWebsiteURL = Replace(txtWebsiteURL,"'","''")
txtFormerCustomerNumber = Replace(txtFormerCustomerNumber,"'","''")
txtFormerCustomerCancelDate = Replace(txtFormerCustomerCancelDate,"'","''")
txtLeaseExpirationDate = Replace(txtLeaseExpirationDate,"'","''")
txtContractExpirationDate = Replace(txtContractExpirationDate,"'","''")
txtCurrentOffering = Replace(txtCurrentOffering,"'","''")
'*******************************************************************************************************************
'SET DEFAULT VALUES FOR ANY NON REQUIRED FIELDS LEFT BLANK DURING THE ADD PROCESS
'*******************************************************************************************************************

If txtTitle = "" Then txtTitle = 0
If txtIndustry = "" Then txtIndustry = 0
If txtTelemarketerUserNo = "" Then txtTelemarketerUserNo = 0
If txtLeadSource = "" Then txtLeadSource = 0
If txtNumEmployees = "" Then txtNumEmployees = 0
If txtProjectedGPSpend = "" Then txtProjectedGPSpend = 0
If txtNumPantries = "" Then txtNumPantries = 1
If txtPrimaryCompetitor = "" Then txtPrimaryCompetitor = 0


If (chkBottledWater <> "" AND chkBottledWater = "on") Then chkBottledWater = 1 Else chkBottledWater = 0
If (chkFilteredWater <> "" AND chkFilteredWater = "on") Then chkFilteredWater = 1 Else chkFilteredWater = 0
If (chkOCS <> "" AND chkOCS = "on") Then chkOCS = 1 Else chkOCS = 0
If (chkOCS_Supply <> "" AND chkOCS_Supply = "on") Then chkOCS_Supply = 1 Else chkOCS_Supply = 0
If (chkOfficeSupplies <> "" AND chkOfficeSupplies = "on") Then chkOfficeSupplies = 1 Else chkOfficeSupplies = 0
If (chkVending <> "" AND chkVending = "on") Then chkVending = 1 Else chkVending = 0
If (chkMicroMarket <> "" AND chkMicroMarket = "on") Then chkMicroMarket = 1 Else chkMicroMarket = 0
If (chkPantry <> "" AND chkPantry = "on") Then chkPantry = 1 Else chkPantry = 0

'*******************************************************************************************************************


'***************************************************************************************************************************************************************
'First make entry into PR_Prospects in order to return Prospect Record Identifier
'***************************************************************************************************************************************************************

SQLProspect = "INSERT INTO PR_Prospects (Company, Street, City, [State], PostalCode, Country, Website, LeadSourceNumber, IndustryNumber, EmployeeRangeNumber, "
SQLProspect = SQLProspect & "OwnerUserNo, CreatedByUserNo, TelemarketerUserNo, Floor_Suite_Room__c, ProjectedGPSpend, NumberOfPantries, FormerCustNum, CancelDate, "
SQLProspect = SQLProspect & "LeaseExpirationDate, ContractExpirationDate, Comments, CurrentOffering, Pool, LastVerifiedDate)"
SQLProspect = SQLProspect &  " VALUES (" 
SQLProspect = SQLProspect & "'"  & txtCompanyName & "','"  & txtAddressLine1 & "','"  & txtCity & "','"  & txtState & "','"  & txtZipCode & "','"  & txtCountry & "',"
SQLProspect = SQLProspect & "'"  & txtWebsiteURL & "',"  & txtLeadSource & ","  & txtIndustry & ","  & txtNumEmployees & "," & Session("UserNo") & "," & Session("UserNo") & ","
SQLProspect = SQLProspect & txtTelemarketerUserNo & ",'"  & txtAddressLine2 & "',"  & txtProjectedGPSpend & ","  & txtNumPantries & ",'"  & txtFormerCustomerNumber & "','"  & txtFormerCustomerCancelDate & "',"
SQLProspect = SQLProspect & "'"  & txtLeaseExpirationDate & "','" & txtContractExpirationDate & "','" & txtComments & "','" & txtCurrentOffering & "', 'Live', getdate())"

'Response.write(SQLProspect)

Set cnnProspect = Server.CreateObject("ADODB.Connection")
cnnProspect.open (Session("ClientCnnString"))

Set rsProspect = Server.CreateObject("ADODB.Recordset")
rsProspect.CursorLocation = 3 
Set rsProspect = cnnProspect.Execute(SQLProspect)


Description = GetUserDisplayNameByUserNo(Session("Userno")) & " added a new prospect with the company name " & txtCompanyName 
CreateAuditLogEntry GetTerm("Prospecting") & " New prospect added",GetTerm("Prospecting") & " new prospect added","Minor",0,Description



'***************************************************************************************************************************************************************
'Now Get Prospect Internal Record Identifier from PR_Prospects
'***************************************************************************************************************************************************************
SQLProspect = "SELECT TOP 1 * FROM PR_Prospects WHERE CreatedByUserNo = " & Session("UserNo") & " ORDER BY RecordCreationDateTime DESC"

rsProspect.CursorLocation = 3 
Set rsProspect = cnnProspect.Execute(SQLProspect)

If Not rsProspect.EOF Then

	ProspectIntRecID = rsProspect("InternalRecordIdentifier")

	Description = "Prospect record created"
	Record_PR_Activity ProspectIntRecID, Description, Session("UserNo")

	Set rsOtherUpdates = Server.CreateObject("ADODB.Recordset")
	rsOtherUpdates.CursorLocation = 3 
	
	'***************************
	'Create Prospect Contact
	'***************************
	
	SQLOtherUpdates = "INSERT INTO PR_ProspectContacts (ProspectIntRecID, Suffix, FirstName, LastName, ContactTitleNumber, "
	SQLOtherUpdates = SQLOtherUpdates & " Email, Phone, PhoneExt, Cell, Fax, PrimaryContact) VALUES ( "
	SQLOtherUpdates = SQLOtherUpdates & ProspectIntRecID & ",'"  & txtSuffix & "','"  & txtFirstName & "','"  & txtLastName & "', "
	SQLOtherUpdates = SQLOtherUpdates & txtTitle & ",'" & txtEmailAddress & "','" & txtPhoneNumber & "','" & txtPhoneNumberExt & "','" & txtCellPhoneNumber & "',"
	SQLOtherUpdates = SQLOtherUpdates & "'" & txtFaxNumber & "',1)"
	
	'Response.write(SQLOtherUpdates)

	Set rsOtherUpdates = cnnProspect.Execute(SQLOtherUpdates)

	Description = txtFirstName & " " & txtLastName & " was set as the primary contact upon creation of the new prospect record."
	Record_PR_Activity ProspectIntRecID, Description, Session("UserNo")

	
	'***************************
	'Set Primary Competitor
	'***************************
	
	If txtPrimaryCompetitor <> 0 Then
		
		SQLOtherUpdates = "INSERT INTO PR_ProspectCompetitors (ProspectRecID, CompetitorRecID, PrimaryCompetitor, BottledWater, FilteredWater, "
		SQLOtherUpdates = SQLOtherUpdates & " OCS, OCS_Supply, Vending, MicroMarket, Pantry, OfficeSupplies) VALUES ("
		SQLOtherUpdates = SQLOtherUpdates & ProspectIntRecID & "," & txtPrimaryCompetitor & ",1, " & chkBottledWater & "," & chkFilteredWater & ", "
		SQLOtherUpdates = SQLOtherUpdates & chkOCS & "," & chkOCS_Supply & "," & chkVending & ","
		SQLOtherUpdates = SQLOtherUpdates & chkMicroMarket & "," & chkPantry & "," & chkOfficeSupplies & ")"
	
		Set rsOtherUpdates = cnnProspect.Execute(SQLOtherUpdates)
	
		CompetitorName = GetCompetitorByNum(txtPrimaryCompetitor)
		
		Description = CompetitorName & " was set as the primary competitor upon creation of the new prospect record."
		Record_PR_Activity ProspectIntRecID, Description, Session("UserNo")
		
	End If

	'***************************
	'Set Next Activity
	'***************************

	If ProspectApptOrMeeting <> "" Then
	
		If ProspectApptOrMeeting ="Appointment" Then
		
			Duration = cint(selAppointmentDuration)
										
			SQLProspectNextActivityInsert = "INSERT INTO PR_ProspectActivities (ProspectRecID, ActivityRecID, ActivityDueDate, ActivityCreatedByUserNo, ActivityIsAppointment, ActivityIsMeeting, ActivityAppointmentDuration, Notes) "
			SQLProspectNextActivityInsert = SQLProspectNextActivityInsert & " VALUES (" & ProspectIntRecID & ", " & txtNextActivity & ",'" & txtNextActivityDueDate & "'," & Session("UserNo") & ",1,0," & Duration & ",'" & txtNextActivityNotes & "') "
						
		ElseIf ProspectApptOrMeeting ="Meeting" Then
		
			Duration = cint(selMeetingDuration)
									
			SQLProspectNextActivityInsert = "INSERT INTO PR_ProspectActivities (ProspectRecID, ActivityRecID, ActivityDueDate, ActivityCreatedByUserNo, ActivityIsAppointment, ActivityIsMeeting, ActivityMeetingDuration, ActivityMeetingLocation, Notes) "
			SQLProspectNextActivityInsert = SQLProspectNextActivityInsert & " VALUES (" & ProspectIntRecID & ", " & txtNextActivity & ",'" & txtNextActivityDueDate & "'," & Session("UserNo") & ",0,1," & Duration & ",'" & txtMeetingLocation & "','" & txtNextActivityNotes & "') "
								
		Else
						
			SQLProspectNextActivityInsert = "INSERT INTO PR_ProspectActivities (ProspectRecID, ActivityRecID, ActivityDueDate, ActivityCreatedByUserNo, ActivityIsAppointment, ActivityIsMeeting, Notes) "
			SQLProspectNextActivityInsert = SQLProspectNextActivityInsert & " VALUES (" & ProspectIntRecID & ", " & txtNextActivity & ",'" & txtNextActivityDueDate & "'," & Session("UserNo") & ",0,0,'" & txtNextActivityNotes & "') "		
		
		End If	
	Else
							
		SQLProspectNextActivityInsert = "INSERT INTO PR_ProspectActivities (ProspectRecID, ActivityRecID, ActivityDueDate, ActivityCreatedByUserNo, ActivityIsAppointment, ActivityIsMeeting, Notes) "
		SQLProspectNextActivityInsert = SQLProspectNextActivityInsert & " VALUES (" & ProspectIntRecID & ", " & txtNextActivity & ",'" & txtNextActivityDueDate & "'," & Session("UserNo") & ",0,0,'" & txtNextActivityNotes & "') "		

	End If
	

	Set rsOtherUpdates = cnnProspect.Execute(SQLProspectNextActivityInsert)

	NextActivity = GetActivityByNum(txtNextActivity)
	
	If txtNextActivityNotes <> "" Then
		Description = NextActivity & " was set as the next activity, with a due date of " & txtNextActivityDueDate & ", upon creation of the new prospect record, with the following notes: " & txtNextActivityNotes & "."
	Else
		Description = NextActivity & " was set as the next activity, with a due date of " & txtNextActivityDueDate & ", upon creation of the new prospect record."
	End If
	
	Record_PR_Activity ProspectIntRecID, Description, Session("UserNo")


	'***************************
	'Set Initial Stage
	'***************************

	SQLOtherUpdates = "INSERT INTO PR_ProspectStages (ProspectRecID, StageRecID, StageChangedByUserNo, Notes) VALUES ("
	SQLOtherUpdates = SQLOtherUpdates & ProspectIntRecID & "," & radStage & "," & Session("UserNo") & ",'" & txtStageNotes & "')"

	Set rsOtherUpdates = cnnProspect.Execute(SQLOtherUpdates)

	intitalStage = GetStageByNum(radStage)
	
	If txtStageNotes <> "" Then
		Description = intitalStage  & " was set as the stage upon creation of the new prospect record, with the following notes: " & txtStageNotes & "."
	Else
		Description = intitalStage  & " was set as the stage upon creation of the new prospect record."
	End If
	
	Record_PR_Activity ProspectIntRecID, Description, Session("UserNo")


	'***************************
	'Set Prospect Creation Note
	'***************************

	SQLOtherUpdates = "INSERT INTO PR_ProspectNotes (ProspectIntRecID, EnteredByUserNo, Note, Sticky, NoteTypeNumber) VALUES ("
	SQLOtherUpdates = SQLOtherUpdates & ProspectIntRecID & "," & Session("UserNo") & ",'Prospect Record Created.',1,4)"

	Set rsOtherUpdates = cnnProspect.Execute(SQLOtherUpdates)
	
	Description = "<strong><em>Prospect Record Created</em></strong> was set as the initial note upon creation of the new prospect record."
	Record_PR_Activity ProspectIntRecID, Description, Session("UserNo")


End If


'*****************************************************************************************************************************************
'If new prospect owner is not set to the current user,  check to see if we need to
'email the propspective prospect owner for approval before adding them as the owner
'Also check to see if we have to make entries in the users Outlook Calendar for the next activity
'*****************************************************************************************************************************************


dummy = SetOwner_MakeOutlookEntry_SendEmail(ProspectIntRecID,txtOwner,sendEmailFlag,"A")



Response.Redirect("viewProspectDetail.asp?i=" & ProspectIntRecID)

%>
