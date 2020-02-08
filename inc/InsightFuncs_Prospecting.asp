<%
'********************************
'List of all the functions & subs
'********************************
'Sub Record_PR_Activity (passedProspectIntRecID,passedActivity)
'Func TotalNumberOfProspects()
'Func GetPercentForStage (passedStageNumber)
'Func GetStageByNum (passedStageNumber)
'Func GetLeadSourceByNum (passedLeadSourceNumber)
'Func GetIndustryByNum (passedStageNumber)
'Func GetLeadSource2ByNum (passedLeadSourceNumber)
'Func GetReasonByNum (passedReason)
'Func GetActivityByNum (passedActivity)
'Func GetNoteTypeByNum (passedNoteType)
'Func GetEmployeeRangeByNum (passedEmployeeRange)
'Func GetCompetitorByNum (passedCompetitorNumber)
'Func GetContactTitleByNum (passedContactTitleNum)
'Func NumberOfProspectsByCompetitorNum (passedCompetitor)
'Func NumberOfProspectsByLeadSourceNum (passedLeadSource)
'Func NumberOfProspectsByIndustryNum (passedIndustry)
'Func NumberOfProspectsByStageNum (passedStage)
'Func NumberOfProspectsByLeadSource2Num (passedLeadSource)
'Func NumberOfProspectsByReasonNum (passedReason)
'Func NumberOfProspectsByActivityNum (passedActivity)
'Func NumberOfContactsByContactTitleNum (passedTitle)
'Func NumberOfProspectAndCustomerContactsByContactTitleNum (passedTitle)
'Func NumberOfProspectsByEmployeeRangeNum (passedRange)
'Func NumberOfLogItemsByProspectNumber (passedProspectNumber)
'Func NumberOfContactsByProspectNumber (passedProspectNumber) 
'Func NumberOfSocialMediaByProspectNumber (passedProspectNumber) 
'Func NumberOfCompetitorsByProspectNumber (passedProspectNumber) 
'Func GetProspectNameByNumber(passedProspectNumber)
'Func GetProspectStreetByNumber(passedProspectNumber)
'Func GetProspectCityByNumber(passedProspectNumber)
'Func GetProspectStateByNumber(passedProspectNumber)
'Func GetProspectPostalCodeByNumber(passedProspectNumber)
'Func GetProspectLeadSourceByProspectIntRecID (passedProspectNumber)
'Func GetCurrentProspectActivityByProspectNumber(passedProspectNumber)
'Func GetCurrentProspectActivityNumberByProspectNumber (passedProspectNumber)
'Func GetCurrentProspectActivityDueDateByProspectNumber (passedProspectNumber)
'Func GetLastProspectActivityNumberByProspectNumber (passedProspectNumber)
'Func GetLastProspectActivityDueDateByProspectNumber (passedProspectNumber)
'Func GetProspectGroupNameByNumber (passedGroupNum)
'Func GetPrimaryCompetitorIDByProspectNumber (passedProspectNumber)
'Func GetActivityCreatedByUserNo (passedInternalRecordIdentifier)
'Func GetProspectOwnerNoByNumber  (passedProspectNumber)
'Func GetProspectCurrentStageByProspectNumber (passedProspectNumber)
'Func GetProspectCurrentStageIntRecIDByProspectNumber (passedProspectNumber)
'Func GetProspectLastStageChangeDateByProspectNumber (passedProspectNumber)
'Func GetCRMMaxActivityDaysPermitted()
'Func GetCRMMaxActivityDaysWarning()
'Func GetStageReasonByStageIntRecID (passedStageIntRecID)
'Func GetActivityApptOrMeetingByNum(passedActivity1)

'Func TotalNumberOfPreexistingProspectsWeeklySnapshot(passedStartDate,passedEndDate,passedUserNos)
'Func TotalNumberOfCreatedProspectsWeeklySnapshot(passedStartDate,passedEndDate,passedUserNos)
'Func TotalNumberOfWonProspectsWeeklySnapshot(passedStartDate,passedEndDate,passedUserNos)
'Func TotalNumberOfLostProspectsWeeklySnapshot(passedStartDate,passedEndDate,passedUserNos)
'Func TotalNumberOfUnqualifiedProspectsWeeklySnapshot(passedStartDate,passedEndDate,passedUserNos)


'Func TotalNumberOfPreexistingAppointmentsWeeklySnapshot(passedStartDate,passedEndDate,passedUserNos)
'Func TotalNumberOfCreatedAppointmentsWeeklySnapshot(passedStartDate,passedEndDate,passedUserNos)
'Func TotalNumberOfCompletedAppointmentsWeeklySnapshot(passedStartDate,passedEndDate,passedUserNos)
'Func TotalNumberOfRescheduledAppointmentsWeeklySnapshot(passedStartDate,passedEndDate,passedUserNos)
'Func TotalNumberOfCancelledAppointmentsWeeklySnapshot(passedStartDate,passedEndDate,passedUserNos)
'Func TotalNumberOfNotUpdatedAppointmentsWeeklySnapshot(passedStartDate,passedEndDate,passedUserNos)
'Func TotalNumberOfExpiredAppointmentsWeeklySnapshot(passedStartDate,passedEndDate,passedUserNos)
'Func TotalNumberOfExpiredActivitiesWeeklySnapshot(passedStartDate,passedEndDate,passedUserNos)

'Func TotalNumberOfCreatedAppointmentsPreexistingWeeklySnapshot(passedStartDate,passedEndDate,passedUserNos)
'Func TotalNumberOfCompletedAppointmentsPreexistingWeeklySnapshot(passedStartDate,passedEndDate,passedUserNos)
'Func TotalNumberOfRescheduledAppointmentsPreexistingWeeklySnapshot(passedStartDate,passedEndDate,passedUserNos)
'Func TotalNumberOfCancelledAppointmentsPreexistingWeeklySnapshot(passedStartDate,passedEndDate,passedUserNos)

'Func TotalNumberOfCreatedProspectsByLeadSourceWeeklySnapshot(passedStartDate,passedEndDate,passedUserNos,passedLeadSourceIntRecID)
'Func TotalNumberOfWonProspectsByLeadSourceWeeklySnapshot(passedStartDate,passedEndDate,passedUserNos,passedLeadSourceIntRecID)
'Func TotalNumberOfLostProspectsByLeadSourceWeeklySnapshot(passedStartDate,passedEndDate,passedUserNos,passedLeadSourceIntRecID)
'Func TotalNumberOfUnqualifiedProspectsByLeadSourceWeeklySnapshot(passedStartDate,passedEndDate,passedUserNos,passedLeadSourceIntRecID)

'Func TotalNumberOfCreatedAppointmentsByLeadSourceWeeklySnapshot(passedStartDate,passedEndDate,passedUserNos,passedLeadSourceIntRecID)
'Func TotalNumberOfCompletedAppointmentsByLeadSourceWeeklySnapshot(passedStartDate,passedEndDate,passedUserNos,passedLeadSourceIntRecID)
'Func TotalNumberOfRescheduledAppointmentsByLeadSourceWeeklySnapshot(passedStartDate,passedEndDate,passedUserNos,passedLeadSourceIntRecID)
'Func TotalNumberOfCancelledAppointmentsByLeadSourceWeeklySnapshot(passedStartDate,passedEndDate,passedUserNos,passedLeadSourceIntRecID)
'Func TotalNumberOfNotUpdatedAppointmentsByLeadSourceWeeklySnapshot(passedStartDate,passedEndDate,passedUserNos,passedLeadSourceIntRecID)
'Func TotalNumberOfExpiredActivitiesByLeadSourceWeeklySnapshot(passedStartDate,passedEndDate,passedUserNos,passedLeadSourceIntRecID)


'************************************
'End List of all the functions & subs
'************************************




Sub Record_PR_Activity(passedProspectIntRecID,passedActivity,passedUserNo)

	'Creates an entry in PR_Audit
	
	SQLRecord_PR_Activity = "INSERT INTO PR_Audit (ProspectIntRecID,Activity,PerformedByUserNo) "
	SQLRecord_PR_Activity = SQLRecord_PR_Activity &  " VALUES (" & passedProspectIntRecID
	SQLRecord_PR_Activity = SQLRecord_PR_Activity & ",'"  & passedActivity & "'," & passedUserNo & ")"
	
'	response.write(SQLRecord_PR_Activity)
	
	Set cnnRecord_PR_Activity = Server.CreateObject("ADODB.Connection")
	cnnRecord_PR_Activity.open (Session("ClientCnnString"))

	Set rsRecord_PR_Activity = Server.CreateObject("ADODB.Recordset")
	rsRecord_PR_Activity.CursorLocation = 3 
	Set rsRecord_PR_Activity = cnnRecord_PR_Activity.Execute(SQLRecord_PR_Activity)
	set rsRecord_PR_Activity = Nothing
	
End Sub

Function TotalNumberOfProspects()


	Set cnnTotalNumberOfProspects = Server.CreateObject("ADODB.Connection")
	cnnTotalNumberOfProspects.open Session("ClientCnnString")

	resultTotalNumberOfProspects = 0
	
	SQLTotalNumberOfProspects =  "SELECT Count(InternalRecordIdentifier) As Expr1 FROM " & "zProspectFilter_" & trim(Session("Userno"))
		
			 
	Set rsTotalNumberOfProspects = Server.CreateObject("ADODB.Recordset")
	rsTotalNumberOfProspects.CursorLocation = 3 
	
	Set rsTotalNumberOfProspects = cnnTotalNumberOfProspects.Execute(SQLTotalNumberOfProspects)
			
	resultTotalNumberOfProspects = rsTotalNumberOfProspects("Expr1")
	
	rsTotalNumberOfProspects.Close
	set rsTotalNumberOfProspects= Nothing
	cnnTotalNumberOfProspects.Close	
	set cnnTotalNumberOfProspects= Nothing
	
	TotalNumberOfProspects = resultTotalNumberOfProspects
	
End Function



Function GetPercentForStage (passedStageNumber)

	resultGetPercentForStage = 0

	Set cnnGetPercentForStage = Server.CreateObject("ADODB.Connection")
	cnnGetPercentForStage.open Session("ClientCnnString")
		
	SQLGetPercentForStage = "Select * from PR_Stages Where InternalRecordIdentifier = " & passedStageNumber
	 
	Set rsGetPercentForStage = Server.CreateObject("ADODB.Recordset")
	rsGetPercentForStage.CursorLocation = 3 
	Set rsGetPercentForStage = cnnGetPercentForStage.Execute(SQLGetPercentForStage)
			
	If not rsGetPercentForStage.EOF Then resultGetPercentForStage = rsGetPercentForStage("ProbabilityPercent")
	
	rsGetPercentForStage.Close
	set rsGetPercentForStage= Nothing
	cnnGetPercentForStage.Close	
	set cnnGetPercentForStage= Nothing
	
	GetPercentForStage = resultGetPercentForStage
	
End Function

Function GetStageByNum (passedStageNumber)

	resultGetStageByNum = ""

	Set cnnGetStageByNum = Server.CreateObject("ADODB.Connection")
	cnnGetStageByNum.open Session("ClientCnnString")
		
	SQLGetStageByNum = "Select Stage from PR_Stages Where InternalRecordIdentifier = " & passedStageNumber
	
	Set rsGetStageByNum = Server.CreateObject("ADODB.Recordset")
	rsGetStageByNum.CursorLocation = 3 
	Set rsGetStageByNum = cnnGetStageByNum.Execute(SQLGetStageByNum)
			
	If not rsGetStageByNum.EOF Then resultGetStageByNum = rsGetStageByNum("Stage")
	
	rsGetStageByNum.Close
	set rsGetStageByNum= Nothing
	cnnGetStageByNum.Close	
	set cnnGetStageByNum= Nothing
	
	GetStageByNum = resultGetStageByNum
	
End Function

Function GetLeadSourceByNum (passedLeadSourceNumber)

	resultGetLeadSourceByNum = ""

	Set cnnGetLeadSourceByNum = Server.CreateObject("ADODB.Connection")
	cnnGetLeadSourceByNum.open Session("ClientCnnString")
		
	SQLGetLeadSourceByNum = "Select LeadSource from PR_LeadSources Where InternalRecordIdentifier = " & passedLeadSourceNumber

	Set rsGetLeadSourceByNum = Server.CreateObject("ADODB.Recordset")
	rsGetLeadSourceByNum.CursorLocation = 3 

	Set rsGetLeadSourceByNum = cnnGetLeadSourceByNum.Execute(SQLGetLeadSourceByNum)
			
	If not rsGetLeadSourceByNum.EOF Then resultGetLeadSourceByNum = rsGetLeadSourceByNum("LeadSource")
	
	rsGetLeadSourceByNum.Close
	set rsGetLeadSourceByNum= Nothing
	cnnGetLeadSourceByNum.Close	
	set cnnGetLeadSourceByNum= Nothing
	
	GetLeadSourceByNum = resultGetLeadSourceByNum
	
End Function


Function NumberOfProspectsByStageNum (passedStage)

	Set cnnNumberOfProspectsByStageNum = Server.CreateObject("ADODB.Connection")
	cnnNumberOfProspectsByStageNum.open Session("ClientCnnString")

	resultNumberOfProspectsByStageNum = 0
		
	SQLNumberOfProspectsByStageNum = "SELECT COUNT(*) AS Expr1 FROM PR_ProspectStages WHERE StageRecID = " & passedStage & " GROUP BY StageRecID"
	 
	Set rsNumberOfProspectsByStageNum = Server.CreateObject("ADODB.Recordset")
	rsNumberOfProspectsByStageNum.CursorLocation = 3 
	
	rsNumberOfProspectsByStageNum.Open SQLNumberOfProspectsByStageNum , cnnNumberOfProspectsByStageNum
		
	If NOT rsNumberOfProspectsByStageNum.EOF Then
		resultNumberOfProspectsByStageNum = rsNumberOfProspectsByStageNum("Expr1")
	Else
		resultNumberOfProspectsByStageNum = 0
	End If
	
	rsNumberOfProspectsByStageNum.Close
	set rsNumberOfProspectsByStageNum= Nothing
	cnnNumberOfProspectsByStageNum.Close	
	set cnnNumberOfProspectsByStageNum= Nothing
	
	NumberOfProspectsByStageNum = resultNumberOfProspectsByStageNum
	
End Function

Function NumberOfProspectsByIndustryNum (passedIndustry)

	Set cnnNumberOfProspectsByIndustryNum = Server.CreateObject("ADODB.Connection")
	cnnNumberOfProspectsByIndustryNum.open Session("ClientCnnString")

	resultNumberOfProspectsByIndustryNum = 0
		
	SQLNumberOfProspectsByIndustryNum = "Select * from PR_Prospects Where IndustryNumber = " & passedIndustry
	 
	Set rsNumberOfProspectsByIndustryNum = Server.CreateObject("ADODB.Recordset")
	rsNumberOfProspectsByIndustryNum.CursorLocation = 3 
	
	rsNumberOfProspectsByIndustryNum.Open SQLNumberOfProspectsByIndustryNum , cnnNumberOfProspectsByIndustryNum
			
	resultNumberOfProspectsByIndustryNum = rsNumberOfProspectsByIndustryNum.RecordCount
	
	rsNumberOfProspectsByIndustryNum.Close
	set rsNumberOfProspectsByIndustryNum= Nothing
	cnnNumberOfProspectsByIndustryNum.Close	
	set cnnNumberOfProspectsByIndustryNum= Nothing
	
	NumberOfProspectsByIndustryNum = resultNumberOfProspectsByIndustryNum
	
End Function

Function GetIndustryByNum (passedIndustryNumber)

	resultGetIndustryByNum = ""
	
	Set cnnGetIndustryByNum = Server.CreateObject("ADODB.Connection")
	cnnGetIndustryByNum.open Session("ClientCnnString")
		
	SQLGetIndustryByNum = "Select Industry from PR_Industries Where InternalRecordIdentifier = " & passedIndustryNumber
	 
	Set rsGetIndustryByNum = Server.CreateObject("ADODB.Recordset")
	rsGetIndustryByNum.CursorLocation = 3 
	Set rsGetIndustryByNum = cnnGetIndustryByNum.Execute(SQLGetIndustryByNum)
			
	If not rsGetIndustryByNum.EOF Then resultGetIndustryByNum = rsGetIndustryByNum("Industry")
	
	rsGetIndustryByNum.Close
	set rsGetIndustryByNum= Nothing
	cnnGetIndustryByNum.Close	
	set cnnGetIndustryByNum= Nothing
	
	GetIndustryByNum = resultGetIndustryByNum
	
End Function

Function NumberOfProspectsByLeadSourceNum  (passedLeadSource)

	Set cnnNumberOfProspectsByLeadSourceNum  = Server.CreateObject("ADODB.Connection")
	cnnNumberOfProspectsByLeadSourceNum.open Session("ClientCnnString")

	resultNumberOfProspectsByLeadSourceNum  = 0
		
	SQLNumberOfProspectsByLeadSourceNum  = "Select * from PR_Prospects Where LeadSourceNumber = " & passedLeadSource
	 
	Set rsNumberOfProspectsByLeadSourceNum  = Server.CreateObject("ADODB.Recordset")
	rsNumberOfProspectsByLeadSourceNum.CursorLocation = 3 
	
	rsNumberOfProspectsByLeadSourceNum.Open SQLNumberOfProspectsByLeadSourceNum  , cnnNumberOfProspectsByLeadSourceNum 
			
	resultNumberOfProspectsByLeadSourceNum  = rsNumberOfProspectsByLeadSourceNum.RecordCount
	
	rsNumberOfProspectsByLeadSourceNum.Close
	set rsNumberOfProspectsByLeadSourceNum = Nothing
	cnnNumberOfProspectsByLeadSourceNum.Close	
	set cnnNumberOfProspectsByLeadSourceNum = Nothing
	
	NumberOfProspectsByLeadSourceNum  = resultNumberOfProspectsByLeadSourceNum 
	
End Function

Function NumberOfProspectsByCompetitorNum  (passedCompetitor)
	
	Set cnnNumberOfProspectsByCompetitorNum  = Server.CreateObject("ADODB.Connection")
	cnnNumberOfProspectsByCompetitorNum.open Session("ClientCnnString")

	resultNumberOfProspectsByCompetitorNum  = 0
		
	SQLNumberOfProspectsByCompetitorNum  = "Select * from PR_ProspectCompetitors Where CompetitorRecID = " & passedCompetitor
	 
	Set rsNumberOfProspectsByCompetitorNum  = Server.CreateObject("ADODB.Recordset")
	rsNumberOfProspectsByCompetitorNum.CursorLocation = 3 
	
	rsNumberOfProspectsByCompetitorNum.Open SQLNumberOfProspectsByCompetitorNum  , cnnNumberOfProspectsByCompetitorNum 
			
	resultNumberOfProspectsByCompetitorNum  = rsNumberOfProspectsByCompetitorNum.RecordCount
	
	rsNumberOfProspectsByCompetitorNum.Close
	set rsNumberOfProspectsByCompetitorNum = Nothing
	cnnNumberOfProspectsByCompetitorNum.Close	
	set cnnNumberOfProspectsByCompetitorNum = Nothing
	
	NumberOfProspectsByCompetitorNum  = resultNumberOfProspectsByCompetitorNum 

	
End Function


Function NumberOfProspectsByLeadSource2Num  (passedLeadSource2)

	Set cnnNumberOfProspectsByLeadSource2Num  = Server.CreateObject("ADODB.Connection")
	cnnNumberOfProspectsByLeadSource2Num.open Session("ClientCnnString")

	resultNumberOfProspectsByLeadSource2Num  = 0
		
	SQLNumberOfProspectsByLeadSource2Num  = "Select * from PR_Prospects Where LeadSource2Number = " & passedLeadSource2
	 
	Set rsNumberOfProspectsByLeadSource2Num  = Server.CreateObject("ADODB.Recordset")
	rsNumberOfProspectsByLeadSource2Num.CursorLocation = 3 
	
	rsNumberOfProspectsByLeadSource2Num.Open SQLNumberOfProspectsByLeadSource2Num  , cnnNumberOfProspectsByLeadSource2Num 
			
	resultNumberOfProspectsByLeadSource2Num  = rsNumberOfProspectsByLeadSource2Num.RecordCount
	
	rsNumberOfProspectsByLeadSource2Num.Close
	set rsNumberOfProspectsByLeadSource2Num = Nothing
	cnnNumberOfProspectsByLeadSource2Num.Close	
	set cnnNumberOfProspectsByLeadSource2Num = Nothing
	
	NumberOfProspectsByLeadSource2Num  = resultNumberOfProspectsByLeadSource2Num 
	
End Function

Function GetLeadSource2ByNum (passedLeadSourceNumber2)

	resultGetLeadSource2ByNum = ""

	Set cnnGetLeadSource2ByNum = Server.CreateObject("ADODB.Connection")
	cnnGetLeadSource2ByNum.open Session("ClientCnnString")
		
	SQLGetLeadSource2ByNum = "Select * from PR_LeadSources2 Where InternalRecordIdentifier = " & passedLeadSourceNumber2
	 
	Set rsGetLeadSource2ByNum = Server.CreateObject("ADODB.Recordset")
	rsGetLeadSource2ByNum.CursorLocation = 3 
	Set rsGetLeadSource2ByNum = cnnGetLeadSource2ByNum.Execute(SQLGetLeadSource2ByNum)
			
	If not rsGetLeadSource2ByNum.EOF Then resultGetLeadSource2ByNum = rsGetLeadSource2ByNum("LeadSource")
	
	rsGetLeadSource2ByNum.Close
	set rsGetLeadSource2ByNum= Nothing
	cnnGetLeadSource2ByNum.Close	
	set cnnGetLeadSource2ByNum= Nothing
	
	GetLeadSource2ByNum = resultGetLeadSource2ByNum
	
End Function

Function NumberOfProspectsByReasonNum  (passedReason)

	Set cnnNumberOfProspectsByReasonNum  = Server.CreateObject("ADODB.Connection")
	cnnNumberOfProspectsByReasonNum.open Session("ClientCnnString")

	resultNumberOfProspectsByReasonNum  = 0
		
	SQLNumberOfProspectsByReasonNum  = "Select * from PR_Prospects Where InternalRecordIdentifier = " & passedReason
	 
	Set rsNumberOfProspectsByReasonNum  = Server.CreateObject("ADODB.Recordset")
	rsNumberOfProspectsByReasonNum.CursorLocation = 3 
	
	rsNumberOfProspectsByReasonNum.Open SQLNumberOfProspectsByReasonNum  , cnnNumberOfProspectsByReasonNum 
			
	resultNumberOfProspectsByReasonNum  = rsNumberOfProspectsByReasonNum.RecordCount
	
	rsNumberOfProspectsByReasonNum.Close
	set rsNumberOfProspectsByReasonNum = Nothing
	cnnNumberOfProspectsByReasonNum.Close	
	set cnnNumberOfProspectsByReasonNum = Nothing
	
	NumberOfProspectsByReasonNum  = resultNumberOfProspectsByReasonNum 
	
End Function

Function GetReasonByNum (passedReason)

	resultGetReasonByNum = 0

	Set cnnGetReasonByNum = Server.CreateObject("ADODB.Connection")
	cnnGetReasonByNum.open Session("ClientCnnString")
		
	SQLGetReasonByNum = "Select * from PR_Reasons Where InternalRecordIdentifier = " & passedReason
	 
	Set rsGetReasonByNum = Server.CreateObject("ADODB.Recordset")
	rsGetReasonByNum.CursorLocation = 3 
	Set rsGetReasonByNum = cnnGetReasonByNum.Execute(SQLGetReasonByNum)
			
	If not rsGetReasonByNum.EOF Then resultGetReasonByNum = rsGetReasonByNum("Reason")
	
	rsGetReasonByNum.Close
	set rsGetReasonByNum= Nothing
	cnnGetReasonByNum.Close	
	set cnnGetReasonByNum= Nothing
	
	GetReasonByNum = resultGetReasonByNum
	
End Function


Function GetEmployeeRangeByNum (passedEmployeeRange)

	resultGetEmployeeRangeByNum = ""

	Set cnnGetEmployeeRangeByNum = Server.CreateObject("ADODB.Connection")
	cnnGetEmployeeRangeByNum.open Session("ClientCnnString")
		
	SQLGetEmployeeRangeByNum = "Select Range from PR_EmployeeRangeTable  Where InternalRecordIdentifier = " & passedEmployeeRange	 
	Set rsGetEmployeeRangeByNum = Server.CreateObject("ADODB.Recordset")
	rsGetEmployeeRangeByNum.CursorLocation = 3 
	Set rsGetEmployeeRangeByNum = cnnGetEmployeeRangeByNum.Execute(SQLGetEmployeeRangeByNum)
			
	If not rsGetEmployeeRangeByNum.EOF Then resultGetEmployeeRangeByNum = rsGetEmployeeRangeByNum("Range")
	
	rsGetEmployeeRangeByNum.Close
	set rsGetEmployeeRangeByNum= Nothing
	cnnGetEmployeeRangeByNum.Close	
	set cnnGetEmployeeRangeByNum= Nothing
	
	GetEmployeeRangeByNum = resultGetEmployeeRangeByNum
	
End Function



Function NumberOfContactsByContactTitleNum  (passedTitleNum)

	Set cnnNumberOfContactsByContactTitleNum  = Server.CreateObject("ADODB.Connection")
	cnnNumberOfContactsByContactTitleNum.open Session("ClientCnnString")

	resultNumberOfContactsByContactTitleNum  = 0
		
	SQLNumberOfContactsByContactTitleNum  = "Select * from PR_ProspectContacts Where ContactTitleNumber = " & passedTitleNum
	 
	Set rsNumberOfContactsByContactTitleNum  = Server.CreateObject("ADODB.Recordset")
	rsNumberOfContactsByContactTitleNum.CursorLocation = 3 
	
	rsNumberOfContactsByContactTitleNum.Open SQLNumberOfContactsByContactTitleNum  , cnnNumberOfContactsByContactTitleNum 
			
	resultNumberOfContactsByContactTitleNum  = rsNumberOfContactsByContactTitleNum.RecordCount
	
	rsNumberOfContactsByContactTitleNum.Close
	set rsNumberOfContactsByContactTitleNum = Nothing
	cnnNumberOfContactsByContactTitleNum.Close	
	set cnnNumberOfContactsByContactTitleNum = Nothing
	
	NumberOfContactsByContactTitleNum  = resultNumberOfContactsByContactTitleNum 
	
End Function




Function NumberOfProspectAndCustomerContactsByContactTitleNum(passedTitleNum)

	Set cnnNumberOfProspectAndCustomerContactsByContactTitleNum = Server.CreateObject("ADODB.Connection")
	cnnNumberOfProspectAndCustomerContactsByContactTitleNum.open Session("ClientCnnString")

	resultNumberOfProspectAndCustomerContactsByContactTitleNum = 0
	NumberOfProspectContactsByContactTitleNum = 0
	NumberOfCustomerContactsByContactTitleNum = 0
		
	SQLNumberOfProspectAndCustomerContactsByContactTitleNum = "SELECT * FROM PR_ProspectContacts WHERE ContactTitleNumber = " & passedTitleNum
	Set rsNumberOfProspectAndCustomerContactsByContactTitleNum = Server.CreateObject("ADODB.Recordset")
	rsNumberOfProspectAndCustomerContactsByContactTitleNum.CursorLocation = 3 
	rsNumberOfProspectAndCustomerContactsByContactTitleNum.Open SQLNumberOfProspectAndCustomerContactsByContactTitleNum,cnnNumberOfProspectAndCustomerContactsByContactTitleNum 	
	NumberOfProspectContactsByContactTitleNum = rsNumberOfProspectAndCustomerContactsByContactTitleNum.RecordCount
	
	SQLNumberOfProspectAndCustomerContactsByContactTitleNum = "SELECT * FROM AR_CustomerContacts WHERE ContactTitleNumber = " & passedTitleNum
	Set rsNumberOfProspectAndCustomerContactsByContactTitleNum = Server.CreateObject("ADODB.Recordset")
	rsNumberOfProspectAndCustomerContactsByContactTitleNum.CursorLocation = 3 
	rsNumberOfProspectAndCustomerContactsByContactTitleNum.Open SQLNumberOfProspectAndCustomerContactsByContactTitleNum,cnnNumberOfProspectAndCustomerContactsByContactTitleNum 
	NumberOfCustomerContactsByContactTitleNum = rsNumberOfProspectAndCustomerContactsByContactTitleNum.RecordCount
	
	resultNumberOfProspectAndCustomerContactsByContactTitleNum = NumberOfProspectContactsByContactTitleNum + NumberOfCustomerContactsByContactTitleNum
	
	resultNumberOfProspectAndCustomerContactsByContactTitleNum = NumProspect
	rsNumberOfProspectAndCustomerContactsByContactTitleNum.Close
	set rsNumberOfProspectAndCustomerContactsByContactTitleNum = Nothing
	cnnNumberOfProspectAndCustomerContactsByContactTitleNum.Close	
	set cnnNumberOfProspectAndCustomerContactsByContactTitleNum = Nothing
	
	NumberOfProspectAndCustomerContactsByContactTitleNum  = resultNumberOfProspectAndCustomerContactsByContactTitleNum 
	
End Function



Function GetContactTitleByNum (passedContactTitleNum)

	resultGetContactTitleByNum = 0

	Set cnnGetContactTitleByNum = Server.CreateObject("ADODB.Connection")
	cnnGetContactTitleByNum.open Session("ClientCnnString")
		
	SQLGetContactTitleByNum = "Select * from PR_ContactTitles Where InternalRecordIdentifier = " & passedContactTitleNum
	 
	Set rsGetContactTitleByNum = Server.CreateObject("ADODB.Recordset")
	rsGetContactTitleByNum.CursorLocation = 3 
	Set rsGetContactTitleByNum = cnnGetContactTitleByNum.Execute(SQLGetContactTitleByNum)
			
	If not rsGetContactTitleByNum.EOF Then resultGetContactTitleByNum = rsGetContactTitleByNum("ContactTitle")
	
	rsGetContactTitleByNum.Close
	set rsGetContactTitleByNum= Nothing
	cnnGetContactTitleByNum.Close	
	set cnnGetContactTitleByNum= Nothing
	
	GetContactTitleByNum = resultGetContactTitleByNum
	
End Function




Function GetNoteTypeByNum (passedNoteType)

	resultGetNoteTypeByNum = 0

	Set cnnGetNoteTypeByNum = Server.CreateObject("ADODB.Connection")
	cnnGetNoteTypeByNum.open Session("ClientCnnString")
		
	SQLGetNoteTypeByNum = "Select * from PR_NoteTypes Where InternalRecordIdentifier = " & passedNoteType
	 
	Set rsGetNoteTypeByNum = Server.CreateObject("ADODB.Recordset")
	rsGetNoteTypeByNum.CursorLocation = 3 
	Set rsGetNoteTypeByNum = cnnGetNoteTypeByNum.Execute(SQLGetNoteTypeByNum)
			
	If not rsGetNoteTypeByNum.EOF Then resultGetNoteTypeByNum = rsGetNoteTypeByNum("NoteType")
	
	rsGetNoteTypeByNum.Close
	set rsGetNoteTypeByNum= Nothing
	cnnGetNoteTypeByNum.Close	
	set cnnGetNoteTypeByNum= Nothing
	
	GetNoteTypeByNum = resultGetNoteTypeByNum
	
End Function


Function GetProspectNameByNumber (passedProspectNumber)

	resultGetProspectNameByNumber = ""

	Set cnnGetProspectNameByNumber = Server.CreateObject("ADODB.Connection")
	cnnGetProspectNameByNumber.open Session("ClientCnnString")
		
	SQLGetProspectNameByNumber = "Select * from PR_Prospects Where InternalRecordIdentifier = " & passedProspectNumber
	 
	Set rsGetProspectNameByNumber = Server.CreateObject("ADODB.Recordset")
	rsGetProspectNameByNumber.CursorLocation = 3 
	Set rsGetProspectNameByNumber = cnnGetProspectNameByNumber.Execute(SQLGetProspectNameByNumber)
			
	If not rsGetProspectNameByNumber.EOF Then resultGetProspectNameByNumber = rsGetProspectNameByNumber("Company")
	
	rsGetProspectNameByNumber.Close
	set rsGetProspectNameByNumber= Nothing
	cnnGetProspectNameByNumber.Close	
	set cnnGetProspectNameByNumber= Nothing
	
	GetProspectNameByNumber = resultGetProspectNameByNumber
	
End Function


Function GetProspectStreetByNumber(passedProspectNumber)

	resultGetProspectStreetByNumber = ""

	Set cnnGetProspectStreetByNumber = Server.CreateObject("ADODB.Connection")
	cnnGetProspectStreetByNumber.open Session("ClientCnnString")
		
	SQLGetProspectStreetByNumber = "Select * from PR_Prospects Where InternalRecordIdentifier = " & passedProspectNumber
	 
	Set rsGetProspectStreetByNumber = Server.CreateObject("ADODB.Recordset")
	rsGetProspectStreetByNumber.CursorLocation = 3 
	Set rsGetProspectStreetByNumber = cnnGetProspectStreetByNumber.Execute(SQLGetProspectStreetByNumber)
			
	If not rsGetProspectStreetByNumber.EOF Then resultGetProspectStreetByNumber = rsGetProspectStreetByNumber("Street")
	
	rsGetProspectStreetByNumber.Close
	set rsGetProspectStreetByNumber= Nothing
	cnnGetProspectStreetByNumber.Close	
	set cnnGetProspectStreetByNumber= Nothing
	
	GetProspectStreetByNumber = resultGetProspectStreetByNumber
	
End Function


Function GetProspectCityByNumber(passedProspectNumber)

	resultGetProspectCityByNumber = ""

	Set cnnGetProspectCityByNumber = Server.CreateObject("ADODB.Connection")
	cnnGetProspectCityByNumber.open Session("ClientCnnString")
		
	SQLGetProspectCityByNumber = "Select * from PR_Prospects Where InternalRecordIdentifier = " & passedProspectNumber
	 
	Set rsGetProspectCityByNumber = Server.CreateObject("ADODB.Recordset")
	rsGetProspectCityByNumber.CursorLocation = 3 
	Set rsGetProspectCityByNumber = cnnGetProspectCityByNumber.Execute(SQLGetProspectCityByNumber)
			
	If not rsGetProspectCityByNumber.EOF Then resultGetProspectCityByNumber = rsGetProspectCityByNumber("City")
	
	rsGetProspectCityByNumber.Close
	set rsGetProspectCityByNumber= Nothing
	cnnGetProspectCityByNumber.Close	
	set cnnGetProspectCityByNumber= Nothing
	
	GetProspectCityByNumber = resultGetProspectCityByNumber
	
End Function


Function GetProspectStateByNumber(passedProspectNumber)

	resultGetProspectStateByNumber = ""

	Set cnnGetProspectStateByNumber = Server.CreateObject("ADODB.Connection")
	cnnGetProspectStateByNumber.open Session("ClientCnnString")
		
	SQLGetProspectStateByNumber = "Select * from PR_Prospects Where InternalRecordIdentifier = " & passedProspectNumber
	 
	Set rsGetProspectStateByNumber = Server.CreateObject("ADODB.Recordset")
	rsGetProspectStateByNumber.CursorLocation = 3 
	Set rsGetProspectStateByNumber = cnnGetProspectStateByNumber.Execute(SQLGetProspectStateByNumber)
			
	If not rsGetProspectStateByNumber.EOF Then resultGetProspectStateByNumber = rsGetProspectStateByNumber("State")
	
	rsGetProspectStateByNumber.Close
	set rsGetProspectStateByNumber= Nothing
	cnnGetProspectStateByNumber.Close	
	set cnnGetProspectStateByNumber= Nothing
	
	GetProspectStateByNumber = resultGetProspectStateByNumber
	
End Function


Function GetProspectPostalCodeByNumber(passedProspectNumber)

	resultGetProspectPostalCodeByNumber = ""

	Set cnnGetProspectPostalCodeByNumber = Server.CreateObject("ADODB.Connection")
	cnnGetProspectPostalCodeByNumber.open Session("ClientCnnString")
		
	SQLGetProspectPostalCodeByNumber = "Select * from PR_Prospects Where InternalRecordIdentifier = " & passedProspectNumber
	 
	Set rsGetProspectPostalCodeByNumber = Server.CreateObject("ADODB.Recordset")
	rsGetProspectPostalCodeByNumber.CursorLocation = 3 
	Set rsGetProspectPostalCodeByNumber = cnnGetProspectPostalCodeByNumber.Execute(SQLGetProspectPostalCodeByNumber)
			
	If not rsGetProspectPostalCodeByNumber.EOF Then resultGetProspectPostalCodeByNumber = rsGetProspectPostalCodeByNumber("PostalCode")
	
	rsGetProspectPostalCodeByNumber.Close
	set rsGetProspectPostalCodeByNumber= Nothing
	cnnGetProspectPostalCodeByNumber.Close	
	set cnnGetProspectPostalCodeByNumber= Nothing
	
	GetProspectPostalCodeByNumber = resultGetProspectPostalCodeByNumber
	
End Function



Function GetProspectGroupNameByNumber(passedGroupNum)

	resultGetProspectGroupNameByNumber = 0

	Set cnnGetProspectGroupNameByNumber = Server.CreateObject("ADODB.Connection")
	cnnGetProspectGroupNameByNumber.open Session("ClientCnnString")
		
	SQLGetProspectGroupNameByNumber = "Select * from PR_ProspectGroups Where InternalGroupNumber = " & passedGroupNum
	 
	Set rsGetProspectGroupNameByNumber = Server.CreateObject("ADODB.Recordset")
	rsGetProspectGroupNameByNumber.CursorLocation = 3 
	Set rsGetProspectGroupNameByNumber = cnnGetProspectGroupNameByNumber.Execute(SQLGetProspectGroupNameByNumber)
			
	If not rsGetProspectGroupNameByNumber.EOF Then resultGetProspectGroupNameByNumber = rsGetProspectGroupNameByNumber("GroupName")
	
	rsGetProspectGroupNameByNumber.Close
	set rsGetProspectGroupNameByNumber= Nothing
	cnnGetProspectGroupNameByNumber.Close	
	set cnnGetProspectGroupNameByNumber= Nothing
	
	GetProspectGroupNameByNumber = resultGetProspectGroupNameByNumber
	
End Function


Function NumberOfProspectsByEmployeeRangeNum(passedRange)

	Set cnnNumberOfEmployeeRangesByRangeNum  = Server.CreateObject("ADODB.Connection")
	cnnNumberOfEmployeeRangesByRangeNum.open Session("ClientCnnString")

	resultNumberOfEmployeeRangesByRangeNum  = 0
		
	SQLNumberOfEmployeeRangesByRangeNum  = "Select * from PR_Prospects Where EmployeeRangeNumber = " & passedRange
	 
	Set rsNumberOfEmployeeRangesByRangeNum  = Server.CreateObject("ADODB.Recordset")
	rsNumberOfEmployeeRangesByRangeNum.CursorLocation = 3 
	
	rsNumberOfEmployeeRangesByRangeNum.Open SQLNumberOfEmployeeRangesByRangeNum  , cnnNumberOfEmployeeRangesByRangeNum 
			
	resultNumberOfEmployeeRangesByRangeNum  = rsNumberOfEmployeeRangesByRangeNum.RecordCount
	
	rsNumberOfEmployeeRangesByRangeNum.Close
	set rsNumberOfEmployeeRangesByRangeNum = Nothing
	cnnNumberOfEmployeeRangesByRangeNum.Close	
	set cnnNumberOfEmployeeRangesByRangeNum = Nothing
	
	NumberOfProspectsByEmployeeRangeNum = resultNumberOfEmployeeRangesByRangeNum
		
End Function


Function userCanEditCRMOnTheFly(passedUserNo)

	resultuserCanEditCRMOnTheFly = false

	Set cnnuserCanEditCRMOnTheFly  = Server.CreateObject("ADODB.Connection")
	cnnuserCanEditCRMOnTheFly.open Session("ClientCnnString")

		
	SQLuserCanEditCRMOnTheFly  = "Select * from tblUsers where UserNo = " & passedUserNo
	 
	Set rsuserCanEditCRMOnTheFly  = Server.CreateObject("ADODB.Recordset")
	rsuserCanEditCRMOnTheFly.CursorLocation = 3 
	
	rsuserCanEditCRMOnTheFly.Open SQLuserCanEditCRMOnTheFly  , cnnuserCanEditCRMOnTheFly 
			
	If not rsuserCanEditCRMOnTheFly.eof then 
		If rsuserCanEditCRMOnTheFly("userEditCRMOnTheFly") = vbTrue Then resultuserCanEditCRMOnTheFly = True Else resultuserCanEditCRMOnTheFly = False
	End IF
	
	rsuserCanEditCRMOnTheFly.Close
	set rsuserCanEditCRMOnTheFly = Nothing
	cnnuserCanEditCRMOnTheFly.Close	
	set cnnuserCanEditCRMOnTheFly = Nothing
	
	userCanEditCRMOnTheFly = resultuserCanEditCRMOnTheFly 
		
End Function

Function GetActivityByNum(passedActivity1)

	resultGetActivityByNum = 0

	Set cnnGetActivityByNum = Server.CreateObject("ADODB.Connection")
	cnnGetActivityByNum.open Session("ClientCnnString")
		
	SQLGetActivityByNum = "Select Activity from PR_Activities Where InternalRecordIdentifier = " & passedActivity1
 
 
	Set rsGetActivityByNum = Server.CreateObject("ADODB.Recordset")
	rsGetActivityByNum.CursorLocation = 3 
	Set rsGetActivityByNum = cnnGetActivityByNum.Execute(SQLGetActivityByNum)
			
	If not rsGetActivityByNum.EOF Then resultGetActivityByNum = rsGetActivityByNum("Activity")
	
	rsGetActivityByNum.Close
	set rsGetActivityByNum= Nothing
	cnnGetActivityByNum.Close	
	set cnnGetActivityByNum= Nothing
	
	GetActivityByNum = resultGetActivityByNum
	
End Function


Function GetActivityApptOrMeetingByNum(passedActivity1)

	resultGetActivityApptOrMeetingByNum = ""

	Set cnnGetActivityApptOrMeetingByNum = Server.CreateObject("ADODB.Connection")
	cnnGetActivityApptOrMeetingByNum.open Session("ClientCnnString")
		
	SQLGetActivityApptOrMeetingByNum = "Select * from PR_Activities Where InternalRecordIdentifier = " & passedActivity1
 
 
	Set rsGetActivityApptOrMeetingByNum = Server.CreateObject("ADODB.Recordset")
	rsGetActivityApptOrMeetingByNum.CursorLocation = 3 
	Set rsGetActivityApptOrMeetingByNum = cnnGetActivityApptOrMeetingByNum.Execute(SQLGetActivityApptOrMeetingByNum)
			
	If not rsGetActivityApptOrMeetingByNum.EOF Then
		If rsGetActivityApptOrMeetingByNum("CreateAppointment") = 1 Then resultGetActivityApptOrMeetingByNum = "Appointment"
		If rsGetActivityApptOrMeetingByNum("CreateMeeting") = 1 Then resultGetActivityApptOrMeetingByNum = "Meeting"
	End If
	
	rsGetActivityApptOrMeetingByNum.Close
	set rsGetActivityApptOrMeetingByNum= Nothing
	cnnGetActivityApptOrMeetingByNum.Close	
	set cnnGetActivityApptOrMeetingByNum= Nothing
	
	GetActivityApptOrMeetingByNum = resultGetActivityApptOrMeetingByNum
	
End Function

Function GetCompetitorByNum (passedCompetitorNumber)

	resultGetCompetitorByNum = 0

	Set cnnGetCompetitorByNum = Server.CreateObject("ADODB.Connection")
	cnnGetCompetitorByNum.open Session("ClientCnnString")
		
	SQLGetCompetitorByNum = "Select * from PR_Competitors Where InternalRecordIdentifier = " & passedCompetitorNumber
	
	Set rsGetCompetitorByNum = Server.CreateObject("ADODB.Recordset")
	rsGetCompetitorByNum.CursorLocation = 3 
	Set rsGetCompetitorByNum = cnnGetCompetitorByNum.Execute(SQLGetCompetitorByNum)
			
	If not rsGetCompetitorByNum.EOF Then resultGetCompetitorByNum = rsGetCompetitorByNum("CompetitorName")
	
	rsGetCompetitorByNum.Close
	set rsGetCompetitorByNum= Nothing
	cnnGetCompetitorByNum.Close	
	set cnnGetCompetitorByNum= Nothing
	
	GetCompetitorByNum = resultGetCompetitorByNum 
	
End Function



Function NumberOfProspectsByActivityNum  (passedActivity)

	Set cnnNumberOfProspectsByActivityNum  = Server.CreateObject("ADODB.Connection")
	cnnNumberOfProspectsByActivityNum.open Session("ClientCnnString")

	resultNumberOfProspectsByActivityNum  = 0
		
	SQLNumberOfProspectsByActivityNum  = "Select COUNT(*) AS Expr1 FROM PR_ProspectActivities Where ActivityRecID = " & passedActivity
	 
	Set rsNumberOfProspectsByActivityNum  = Server.CreateObject("ADODB.Recordset")
	rsNumberOfProspectsByActivityNum.CursorLocation = 3 
	
	rsNumberOfProspectsByActivityNum.Open SQLNumberOfProspectsByActivityNum  , cnnNumberOfProspectsByActivityNum 
			
	If NOT rsNumberOfProspectsByActivityNum.EOF Then
		resultNumberOfProspectsByActivityNum  = rsNumberOfProspectsByActivityNum("Expr1")
	Else
		resultNumberOfProspectsByActivityNum  = 0
	End If
	
	rsNumberOfProspectsByActivityNum.Close
	set rsNumberOfProspectsByActivityNum = Nothing
	cnnNumberOfProspectsByActivityNum.Close	
	set cnnNumberOfProspectsByActivityNum = Nothing
	
	NumberOfProspectsByActivityNum  = resultNumberOfProspectsByActivityNum 
	
End Function


Function GetCurrentProspectActivityByProspectNumber(passedProspectNumber)

    resultGetCurrentProspectActivityByProspectNumber = ""

    Set cnnGetCurrentProspectActivityByProspectNumber = Server.CreateObject("ADODB.Connection")
    cnnGetCurrentProspectActivityByProspectNumber.open Session("ClientCnnString")
                    
    SQLGetCurrentProspectActivityByProspectNumber = "Select * from PR_ProspectActivities Where ProspectRecID = " & passedProspectNumber & " AND Status Is Null"

    Set rsGetCurrentProspectActivityByProspectNumber = Server.CreateObject("ADODB.Recordset")
    rsGetCurrentProspectActivityByProspectNumber.CursorLocation = 3 
    Set rsGetCurrentProspectActivityByProspectNumber = cnnGetCurrentProspectActivityByProspectNumber.Execute(SQLGetCurrentProspectActivityByProspectNumber)
                                    
    If not rsGetCurrentProspectActivityByProspectNumber.EOF Then resultGetCurrentProspectActivityByProspectNumber = GetActivityByNum(rsGetCurrentProspectActivityByProspectNumber("ActivityRecID"))
    
    rsGetCurrentProspectActivityByProspectNumber.Close
    set rsGetCurrentProspectActivityByProspectNumber= Nothing
    cnnGetCurrentProspectActivityByProspectNumber.Close            
    set cnnGetCurrentProspectActivityByProspectNumber= Nothing
    
    GetCurrentProspectActivityByProspectNumber = resultGetCurrentProspectActivityByProspectNumber
                
End Function


Function GetCurrentProspectActivityNumberByProspectNumber (passedProspectNumber)

	resultGetCurrentProspectActivityNumberByProspectNumber = ""

	Set cnnGetCurrentProspectActivityNumberByProspectNumber = Server.CreateObject("ADODB.Connection")
	cnnGetCurrentProspectActivityNumberByProspectNumber.open Session("ClientCnnString")
		
	SQLGetCurrentProspectActivityNumberByProspectNumber = "Select * from PR_ProspectActivities Where ProspectRecID = " & passedProspectNumber & " AND Status Is Null"
 
	Set rsGetCurrentProspectActivityNumberByProspectNumber = Server.CreateObject("ADODB.Recordset")
	rsGetCurrentProspectActivityNumberByProspectNumber.CursorLocation = 3 
	Set rsGetCurrentProspectActivityNumberByProspectNumber = cnnGetCurrentProspectActivityNumberByProspectNumber.Execute(SQLGetCurrentProspectActivityNumberByProspectNumber)
			 
	If not rsGetCurrentProspectActivityNumberByProspectNumber.EOF Then resultGetCurrentProspectActivityNumberByProspectNumber = rsGetCurrentProspectActivityNumberByProspectNumber("ActivityRecID")
	
	rsGetCurrentProspectActivityNumberByProspectNumber.Close
	set rsGetCurrentProspectActivityNumberByProspectNumber= Nothing
	cnnGetCurrentProspectActivityNumberByProspectNumber.Close	
	set cnnGetCurrentProspectActivityNumberByProspectNumber= Nothing
	
	GetCurrentProspectActivityNumberByProspectNumber = resultGetCurrentProspectActivityNumberByProspectNumber
	
End Function

Function GetCurrentProspectActivityDueDateByProspectNumber (passedProspectNumber)

	resultGetCurrentProspectActivityDueDateByProspectNumber = ""

	Set cnnGetCurrentProspectActivityDueDateByProspectNumber = Server.CreateObject("ADODB.Connection")
	cnnGetCurrentProspectActivityDueDateByProspectNumber.open Session("ClientCnnString")
		
	SQLGetCurrentProspectActivityDueDateByProspectNumber = "Select * from PR_ProspectActivities Where ProspectRecID = " & passedProspectNumber & " AND Status Is Null"
 
	Set rsGetCurrentProspectActivityDueDateByProspectNumber = Server.CreateObject("ADODB.Recordset")
	rsGetCurrentProspectActivityDueDateByProspectNumber.CursorLocation = 3 
	Set rsGetCurrentProspectActivityDueDateByProspectNumber = cnnGetCurrentProspectActivityDueDateByProspectNumber.Execute(SQLGetCurrentProspectActivityDueDateByProspectNumber)
			 
	If not rsGetCurrentProspectActivityDueDateByProspectNumber.EOF Then resultGetCurrentProspectActivityDueDateByProspectNumber =  rsGetCurrentProspectActivityDueDateByProspectNumber("ActivityDueDate")
	
	rsGetCurrentProspectActivityDueDateByProspectNumber.Close
	set rsGetCurrentProspectActivityDueDateByProspectNumber= Nothing
	cnnGetCurrentProspectActivityDueDateByProspectNumber.Close	
	set cnnGetCurrentProspectActivityDueDateByProspectNumber= Nothing
	
	GetCurrentProspectActivityDueDateByProspectNumber = resultGetCurrentProspectActivityDueDateByProspectNumber
	
End Function



Function GetLastProspectActivityNumberByProspectNumber (passedProspectNumber)

	resultGetLastProspectActivityNumberByProspectNumber = ""

	Set cnnGetLastProspectActivityNumberByProspectNumber = Server.CreateObject("ADODB.Connection")
	cnnGetLastProspectActivityNumberByProspectNumber.open Session("ClientCnnString")
		
	SQLGetLastProspectActivityNumberByProspectNumber = "SELECT TOP 1 * ActivityRecID from PR_ProspectActivities Where ProspectRecID = " & passedProspectNumber & " ORDER BY RecordCreationDateTime DESC"
 
	Set rsGetLastProspectActivityNumberByProspectNumber = Server.CreateObject("ADODB.Recordset")
	rsGetLastProspectActivityNumberByProspectNumber.CursorLocation = 3 
	Set rsGetLastProspectActivityNumberByProspectNumber = cnnGetLastProspectActivityNumberByProspectNumber.Execute(SQLGetLastProspectActivityNumberByProspectNumber)
			 
	If not rsGetLastProspectActivityNumberByProspectNumber.EOF Then resultGetLastProspectActivityNumberByProspectNumber = rsGetLastProspectActivityNumberByProspectNumber("ActivityRecID")
	
	rsGetLastProspectActivityNumberByProspectNumber.Close
	set rsGetLastProspectActivityNumberByProspectNumber= Nothing
	cnnGetLastProspectActivityNumberByProspectNumber.Close	
	set cnnGetLastProspectActivityNumberByProspectNumber= Nothing
	
	GetLastProspectActivityNumberByProspectNumber = resultGetLastProspectActivityNumberByProspectNumber
	
End Function


Function GetLastProspectActivityDueDateByProspectNumber (passedProspectNumber)

	resultGetLastProspectActivityDueDateByProspectNumber = ""

	Set cnnGetLastProspectActivityDueDateByProspectNumber = Server.CreateObject("ADODB.Connection")
	cnnGetLastProspectActivityDueDateByProspectNumber.open Session("ClientCnnString")
		
	SQLGetLastProspectActivityDueDateByProspectNumber = "SELECT TOP 1 * ActivityDueDate from PR_ProspectActivities Where ProspectRecID = " & passedProspectNumber & " ORDER BY RecordCreationDateTime DESC"
 
	Set rsGetLastProspectActivityDueDateByProspectNumber = Server.CreateObject("ADODB.Recordset")
	rsGetLastProspectActivityDueDateByProspectNumber.CursorLocation = 3 
	Set rsGetLastProspectActivityDueDateByProspectNumber = cnnGetLastProspectActivityDueDateByProspectNumber.Execute(SQLGetLastProspectActivityDueDateByProspectNumber)
			 
	If not rsGetLastProspectActivityDueDateByProspectNumber.EOF Then resultGetLastProspectActivityDueDateByProspectNumber =  rsGetLastProspectActivityDueDateByProspectNumber("ActivityDueDate")
	
	rsGetLastProspectActivityDueDateByProspectNumber.Close
	set rsGetLastProspectActivityDueDateByProspectNumber= Nothing
	cnnGetLastProspectActivityDueDateByProspectNumber.Close	
	set cnnGetLastProspectActivityDueDateByProspectNumber= Nothing
	
	GetLastProspectActivityDueDateByProspectNumber = resultGetLastProspectActivityDueDateByProspectNumber
	
End Function



Function NumberOfLogItemsByProspectNumber (passedProspectNumber)

	resultNumberOfLogItemsByProspectNumber = 0

	Set cnnNumberOfLogItemsByProspectNumber = Server.CreateObject("ADODB.Connection")
	cnnNumberOfLogItemsByProspectNumber.open Session("ClientCnnString")
		

	SQLNumberOfLogItemsByProspectNumber = "SELECT Count(DetailType) as NoteCount FROM "
	SQLNumberOfLogItemsByProspectNumber = SQLNumberOfLogItemsByProspectNumber & "(SELECT 'Note' AS DetailType, InternalRecordIdentifier as id, DateAndTime, NoteTypeNumber, EnteredByUserNo, Note, Sticky, 0 as StageNumber, '' AS ActivityStatus, 0 as ActivityNumber "
	SQLNumberOfLogItemsByProspectNumber = SQLNumberOfLogItemsByProspectNumber & "FROM PR_ProspectNotes WHERE ProspectIntRecID = " & passedProspectNumber & " "
	SQLNumberOfLogItemsByProspectNumber = SQLNumberOfLogItemsByProspectNumber & "UNION "
	SQLNumberOfLogItemsByProspectNumber = SQLNumberOfLogItemsByProspectNumber & "SELECT 'Activity' AS DetailType, InternalRecordIdentifier as id, RecordCreationDateTime, 0 AS Expr1, StatusChangedByUserNo, Notes AS Expr2, 0 AS Expr3, 0 as StageNumber, Status, ActivityRecID "
	SQLNumberOfLogItemsByProspectNumber = SQLNumberOfLogItemsByProspectNumber & "FROM PR_ProspectActivities WHERE ProspectRecID = " & passedProspectNumber & " "
	SQLNumberOfLogItemsByProspectNumber = SQLNumberOfLogItemsByProspectNumber & "UNION "
	SQLNumberOfLogItemsByProspectNumber = SQLNumberOfLogItemsByProspectNumber & "SELECT 'Stage Change' AS DetailType,InternalRecordIdentifier as id, RecordCreationDateTime, 0 AS Expr1, StageChangedByUserNo As UserNo, Notes AS Expr2, 0 AS Expr3, StageRecID AS StageNumber, '' AS ActivityStatus, 0 as ActivityNumber "
	SQLNumberOfLogItemsByProspectNumber = SQLNumberOfLogItemsByProspectNumber & "FROM PR_ProspectStages AS PR_ProspectStages_1 WHERE ProspectRecID = " & passedProspectNumber & " "
	SQLNumberOfLogItemsByProspectNumber = SQLNumberOfLogItemsByProspectNumber & "UNION "
	SQLNumberOfLogItemsByProspectNumber = SQLNumberOfLogItemsByProspectNumber & "SELECT 'Email' AS DetailType,InternalRecordIdentifier as id, RecordCreationDateTime, 0 AS Expr1, 0 As UserNo, '' AS Expr2, Sticky, 0 AS StageNumber, '' AS ActivityStatus, 0 as ActivityNumber   "
	SQLNumberOfLogItemsByProspectNumber = SQLNumberOfLogItemsByProspectNumber & "FROM PR_ProspectEmailLog AS PR_ProspectEmailLog1_1 WHERE ProspectRecID = " & passedProspectNumber & ") AS t1 "

 
	Set rsNumberOfLogItemsByProspectNumber = Server.CreateObject("ADODB.Recordset")
	rsNumberOfLogItemsByProspectNumber.CursorLocation = 3 
	Set rsNumberOfLogItemsByProspectNumber = cnnNumberOfLogItemsByProspectNumber.Execute(SQLNumberOfLogItemsByProspectNumber)
			 
	If not rsNumberOfLogItemsByProspectNumber.EOF Then resultNumberOfLogItemsByProspectNumber =  rsNumberOfLogItemsByProspectNumber("NoteCount")
	
	rsNumberOfLogItemsByProspectNumber.Close
	set rsNumberOfLogItemsByProspectNumber= Nothing
	cnnNumberOfLogItemsByProspectNumber.Close	
	set cnnNumberOfLogItemsByProspectNumber= Nothing
	
	NumberOfLogItemsByProspectNumber = resultNumberOfLogItemsByProspectNumber
	
End Function

Function NumberOfContactsByProspectNumber (passedProspectNumber)

	resultNumberOfContactsByProspectNumber = 0

	Set cnnNumberOfContactsByProspectNumber = Server.CreateObject("ADODB.Connection")
	cnnNumberOfContactsByProspectNumber.open Session("ClientCnnString")
		
	SQLNumberOfContactsByProspectNumber = "Select Count(*) As ContactCount from PR_ProspectContacts Where ProspectIntRecID = " & passedProspectNumber 
 
	Set rsNumberOfContactsByProspectNumber = Server.CreateObject("ADODB.Recordset")
	rsNumberOfContactsByProspectNumber.CursorLocation = 3 
	Set rsNumberOfContactsByProspectNumber = cnnNumberOfContactsByProspectNumber.Execute(SQLNumberOfContactsByProspectNumber)
			 
	If not rsNumberOfContactsByProspectNumber.EOF Then resultNumberOfContactsByProspectNumber =  rsNumberOfContactsByProspectNumber("ContactCount")
	
	rsNumberOfContactsByProspectNumber.Close
	set rsNumberOfContactsByProspectNumber= Nothing
	cnnNumberOfContactsByProspectNumber.Close	
	set cnnNumberOfContactsByProspectNumber= Nothing
	
	NumberOfContactsByProspectNumber = resultNumberOfContactsByProspectNumber
	
End Function

Function NumberOfSocialMediaByProspectNumber (passedProspectNumber)

	resultNumberOfSocialMediaByProspectNumber = 0

	Set cnnNumberOfSocialMediaByProspectNumber = Server.CreateObject("ADODB.Connection")
	cnnNumberOfSocialMediaByProspectNumber.open Session("ClientCnnString")
		
	SQLNumberOfSocialMediaByProspectNumber = "Select Count(*) As SocialCount from PR_ProspectSocialMedia Where ProspectIntRecID = " & passedProspectNumber 
 
	Set rsNumberOfSocialMediaByProspectNumber = Server.CreateObject("ADODB.Recordset")
	rsNumberOfSocialMediaByProspectNumber.CursorLocation = 3 
	Set rsNumberOfSocialMediaByProspectNumber = cnnNumberOfSocialMediaByProspectNumber.Execute(SQLNumberOfSocialMediaByProspectNumber)
			 
	If not rsNumberOfSocialMediaByProspectNumber.EOF Then resultNumberOfSocialMediaByProspectNumber =  rsNumberOfSocialMediaByProspectNumber("SocialCount")
	
	rsNumberOfSocialMediaByProspectNumber.Close
	set rsNumberOfSocialMediaByProspectNumber= Nothing
	cnnNumberOfSocialMediaByProspectNumber.Close	
	set cnnNumberOfSocialMediaByProspectNumber= Nothing
	
	NumberOfSocialMediaByProspectNumber = resultNumberOfSocialMediaByProspectNumber
	
End Function

Function NumberOfCompetitorsByProspectNumber (passedProspectNumber)

	resultNumberOfCompetitorsByProspectNumber = 0

	Set cnnNumberOfCompetitorsByProspectNumber = Server.CreateObject("ADODB.Connection")
	cnnNumberOfCompetitorsByProspectNumber.open Session("ClientCnnString")
		
	SQLNumberOfCompetitorsByProspectNumber = "Select Count(*) As CompetitorCount from PR_ProspectCompetitors Where ProspectRecID = " & passedProspectNumber 
 
	Set rsNumberOfCompetitorsByProspectNumber = Server.CreateObject("ADODB.Recordset")
	rsNumberOfCompetitorsByProspectNumber.CursorLocation = 3 
	Set rsNumberOfCompetitorsByProspectNumber = cnnNumberOfCompetitorsByProspectNumber.Execute(SQLNumberOfCompetitorsByProspectNumber)
			 
	If not rsNumberOfCompetitorsByProspectNumber.EOF Then resultNumberOfCompetitorsByProspectNumber =  rsNumberOfCompetitorsByProspectNumber("CompetitorCount")
	
	rsNumberOfCompetitorsByProspectNumber.Close
	set rsNumberOfCompetitorsByProspectNumber= Nothing
	cnnNumberOfCompetitorsByProspectNumber.Close	
	set cnnNumberOfCompetitorsByProspectNumber= Nothing
	
	NumberOfCompetitorsByProspectNumber = resultNumberOfCompetitorsByProspectNumber
	
End Function

Function GetPrimaryCompetitorIDByProspectNumber (passedProspectNumber)

	resultGetPrimaryCompetitorIDByProspectNumber = ""

	Set cnnGetPrimaryCompetitorIDByProspectNumber = Server.CreateObject("ADODB.Connection")
	cnnGetPrimaryCompetitorIDByProspectNumber.open Session("ClientCnnString")
		
	SQLGetPrimaryCompetitorIDByProspectNumber = "Select * from PR_ProspectCompetitors Where ProspectRecID = " & passedProspectNumber & " AND PrimaryCompetitor = 1"
 
	Set rsGetPrimaryCompetitorIDByProspectNumber = Server.CreateObject("ADODB.Recordset")
	rsGetPrimaryCompetitorIDByProspectNumber.CursorLocation = 3 
	Set rsGetPrimaryCompetitorIDByProspectNumber = cnnGetPrimaryCompetitorIDByProspectNumber.Execute(SQLGetPrimaryCompetitorIDByProspectNumber)
			 
	If not rsGetPrimaryCompetitorIDByProspectNumber.EOF Then resultGetPrimaryCompetitorIDByProspectNumber =  rsGetPrimaryCompetitorIDByProspectNumber("CompetitorRecId")
	
	rsGetPrimaryCompetitorIDByProspectNumber.Close
	set rsGetPrimaryCompetitorIDByProspectNumber= Nothing
	cnnGetPrimaryCompetitorIDByProspectNumber.Close	
	set cnnGetPrimaryCompetitorIDByProspectNumber= Nothing
	
	GetPrimaryCompetitorIDByProspectNumber = resultGetPrimaryCompetitorIDByProspectNumber
	
End Function

Function GetActivityCreatedByUserNo(passedInternalRecordIdentifier)

	resultGetActivityCreatedByUserNo = 0

	Set cnnGetActivityCreatedByUserNo = Server.CreateObject("ADODB.Connection")
	cnnGetActivityCreatedByUserNo.open Session("ClientCnnString")
		
	SQLGetActivityCreatedByUserNo = "Select * from PR_ProspectActivities Where InternalRecordIdentifier = " & passedInternalRecordIdentifier
 
	Set rsGetActivityCreatedByUserNo = Server.CreateObject("ADODB.Recordset")
	rsGetActivityCreatedByUserNo.CursorLocation = 3 
	Set rsGetActivityCreatedByUserNo = cnnGetActivityCreatedByUserNo.Execute(SQLGetActivityCreatedByUserNo )
			 
	If not rsGetActivityCreatedByUserNo.EOF Then resultGetActivityCreatedByUserNo = rsGetActivityCreatedByUserNo("ActivityCreatedByUserNo")
	
	rsGetActivityCreatedByUserNo.Close
	set rsGetActivityCreatedByUserNo = Nothing
	cnnGetActivityCreatedByUserNo.Close	
	set cnnGetActivityCreatedByUserNo = Nothing
	
	GetActivityCreatedByUserNo = resultGetActivityCreatedByUserNo 
	
End Function

Function NumberOfDocumentsByProspectNumber (passedProspectNumber)

	resultNumberOfDocumentsByProspectNumber = 0

	Set cnnNumberOfDocumentsByProspectNumber = Server.CreateObject("ADODB.Connection")
	cnnNumberOfDocumentsByProspectNumber.open Session("ClientCnnString")
		
	SQLNumberOfDocumentsByProspectNumber = "Select Count(*) As DocumentCount from PR_ProspectDocuments Where ProspectRecID = " & passedProspectNumber 
 
	Set rsNumberOfDocumentsByProspectNumber = Server.CreateObject("ADODB.Recordset")
	rsNumberOfDocumentsByProspectNumber.CursorLocation = 3 
	Set rsNumberOfDocumentsByProspectNumber = cnnNumberOfDocumentsByProspectNumber.Execute(SQLNumberOfDocumentsByProspectNumber)
			 
	If not rsNumberOfDocumentsByProspectNumber.EOF Then resultNumberOfDocumentsByProspectNumber =  rsNumberOfDocumentsByProspectNumber("DocumentCount")
	
	rsNumberOfDocumentsByProspectNumber.Close
	set rsNumberOfDocumentsByProspectNumber= Nothing
	cnnNumberOfDocumentsByProspectNumber.Close	
	set cnnNumberOfDocumentsByProspectNumber= Nothing
	
	NumberOfDocumentsByProspectNumber = resultNumberOfDocumentsByProspectNumber
	
End Function

Function GetProspectOwnerNoByNumber(passedProspectNumber)

    resultGetProspectOwnerNoByNumber = ""

    Set cnnGetProspectOwnerNoByNumber = Server.CreateObject("ADODB.Connection")
    cnnGetProspectOwnerNoByNumber.open Session("ClientCnnString")
                    
    SQLGetProspectOwnerNoByNumber = "Select OwnerUserNo from PR_Prospects Where InternalRecordIdentifier = " & passedProspectNumber
    
    Set rsGetProspectOwnerNoByNumber = Server.CreateObject("ADODB.Recordset")
    rsGetProspectOwnerNoByNumber.CursorLocation = 3 
    Set rsGetProspectOwnerNoByNumber = cnnGetProspectOwnerNoByNumber.Execute(SQLGetProspectOwnerNoByNumber)
                                    
    If not rsGetProspectOwnerNoByNumber.EOF Then resultGetProspectOwnerNoByNumber = rsGetProspectOwnerNoByNumber("OwnerUserNo")
    
    rsGetProspectOwnerNoByNumber.Close
    set rsGetProspectOwnerNoByNumber= Nothing
    cnnGetProspectOwnerNoByNumber.Close         
    set cnnGetProspectOwnerNoByNumber= Nothing
    
    GetProspectOwnerNoByNumber = resultGetProspectOwnerNoByNumber
                
End Function

Function GetProspectCurrentStageByProspectNumber (passedProspectNumber)

	resultGetProspectCurrentStageByProspectNumber = ""

	Set cnnGetProspectCurrentStageByProspectNumber = Server.CreateObject("ADODB.Connection")
	cnnGetProspectCurrentStageByProspectNumber.open Session("ClientCnnString")
		
	SQLGetProspectCurrentStageByProspectNumber = "Select Top 1 * from PR_ProspectStages Where ProspectRecID = " & passedProspectNumber & " ORDER BY RecordCReationDateTime DESC"
 
	Set rsGetProspectCurrentStageByProspectNumber = Server.CreateObject("ADODB.Recordset")
	rsGetProspectCurrentStageByProspectNumber.CursorLocation = 3 
	Set rsGetProspectCurrentStageByProspectNumber = cnnGetProspectCurrentStageByProspectNumber.Execute(SQLGetProspectCurrentStageByProspectNumber)
			 
	If not rsGetProspectCurrentStageByProspectNumber.EOF Then resultGetProspectCurrentStageByProspectNumber = rsGetProspectCurrentStageByProspectNumber("StageRecID")
	
	rsGetProspectCurrentStageByProspectNumber.Close
	set rsGetProspectCurrentStageByProspectNumber= Nothing
	cnnGetProspectCurrentStageByProspectNumber.Close	
	set cnnGetProspectCurrentStageByProspectNumber= Nothing
	
	GetProspectCurrentStageByProspectNumber = resultGetProspectCurrentStageByProspectNumber
	
End Function

Function GetProspectCurrentStageIntRecIDByProspectNumber (passedProspectNumber)

	resultGetProspectCurrentStageIntRecIDByProspectNumber = ""

	Set cnnGetProspectCurrentStageIntRecIDByProspectNumber = Server.CreateObject("ADODB.Connection")
	cnnGetProspectCurrentStageIntRecIDByProspectNumber.open Session("ClientCnnString")
		
	SQLGetProspectCurrentStageIntRecIDByProspectNumber = "Select Top 1 * from PR_ProspectStages Where ProspectRecID = " & passedProspectNumber & " ORDER BY RecordCReationDateTime DESC"
 
	Set rsGetProspectCurrentStageIntRecIDByProspectNumber = Server.CreateObject("ADODB.Recordset")
	rsGetProspectCurrentStageIntRecIDByProspectNumber.CursorLocation = 3 
	Set rsGetProspectCurrentStageIntRecIDByProspectNumber = cnnGetProspectCurrentStageIntRecIDByProspectNumber.Execute(SQLGetProspectCurrentStageIntRecIDByProspectNumber)
			 
	If not rsGetProspectCurrentStageIntRecIDByProspectNumber.EOF Then resultGetProspectCurrentStageIntRecIDByProspectNumber = rsGetProspectCurrentStageIntRecIDByProspectNumber("InternalRecordIdentifier")
	
	rsGetProspectCurrentStageIntRecIDByProspectNumber.Close
	set rsGetProspectCurrentStageIntRecIDByProspectNumber= Nothing
	cnnGetProspectCurrentStageIntRecIDByProspectNumber.Close	
	set cnnGetProspectCurrentStageIntRecIDByProspectNumber= Nothing
	
	GetProspectCurrentStageIntRecIDByProspectNumber = resultGetProspectCurrentStageIntRecIDByProspectNumber
	
End Function


Function GetProspectLastStageChangeDateByProspectNumber (passedProspectNumber)

	resultGetProspectLastStageChangeDateByProspectNumber = ""

	Set cnnGetProspectLastStageChangeDateByProspectNumber = Server.CreateObject("ADODB.Connection")
	cnnGetProspectLastStageChangeDateByProspectNumber.open Session("ClientCnnString")
		
	SQLGetProspectLastStageChangeDateByProspectNumber = "Select Top 1 RecordCReationDateTime from PR_ProspectStages Where ProspectRecID = " & passedProspectNumber & " ORDER BY RecordCReationDateTime DESC"
 
	Set rsGetProspectLastStageChangeDateByProspectNumber = Server.CreateObject("ADODB.Recordset")
	rsGetProspectLastStageChangeDateByProspectNumber.CursorLocation = 3 
	Set rsGetProspectLastStageChangeDateByProspectNumber = cnnGetProspectLastStageChangeDateByProspectNumber.Execute(SQLGetProspectLastStageChangeDateByProspectNumber)
			 
	If not rsGetProspectLastStageChangeDateByProspectNumber.EOF Then resultGetProspectLastStageChangeDateByProspectNumber = rsGetProspectLastStageChangeDateByProspectNumber("RecordCReationDateTime")
	
	rsGetProspectLastStageChangeDateByProspectNumber.Close
	set rsGetProspectLastStageChangeDateByProspectNumber= Nothing
	cnnGetProspectLastStageChangeDateByProspectNumber.Close	
	set cnnGetProspectLastStageChangeDateByProspectNumber= Nothing
	
	GetProspectLastStageChangeDateByProspectNumber = resultGetProspectLastStageChangeDateByProspectNumber
	
End Function


Function GetCRMMaxActivityDaysWarning()

    Set cnnGetCRMMaxActivityDaysWarning = Server.CreateObject("ADODB.Connection")
    cnnGetCRMMaxActivityDaysWarning.open Session("ClientCnnString")

    resultGetCRMMaxActivityDaysWarning = 0
                    
    SQLGetCRMMaxActivityDaysWarning = "Select * from Settings_Global"
    
    Set rsGetCRMMaxActivityDaysWarning = Server.CreateObject("ADODB.Recordset")
    rsGetCRMMaxActivityDaysWarning.CursorLocation = 3 
    Set rsGetCRMMaxActivityDaysWarning= cnnGetCRMMaxActivityDaysWarning.Execute(SQLGetCRMMaxActivityDaysWarning)
    

    If not rsGetCRMMaxActivityDaysWarning.eof then resultGetCRMMaxActivityDaysWarning = rsGetCRMMaxActivityDaysWarning("CRMMaxActivityDaysWarning")


    GetCRMMaxActivityDaysWarning = resultGetCRMMaxActivityDaysWarning 
                
End Function

Function GetCRMMaxActivityDaysPermitted()

    Set cnnGetCRMMaxActivityDaysPermitted = Server.CreateObject("ADODB.Connection")
    cnnGetCRMMaxActivityDaysPermitted.open Session("ClientCnnString")

    resultGetCRMMaxActivityDaysPermitted = 0
                    
    SQLGetCRMMaxActivityDaysPermitted = "Select * from Settings_Global"
    
    Set rsGetCRMMaxActivityDaysPermitted = Server.CreateObject("ADODB.Recordset")
    rsGetCRMMaxActivityDaysPermitted.CursorLocation = 3 
    Set rsGetCRMMaxActivityDaysPermitted= cnnGetCRMMaxActivityDaysPermitted.Execute(SQLGetCRMMaxActivityDaysPermitted)
    
    
    If not rsGetCRMMaxActivityDaysPermitted.eof then resultGetCRMMaxActivityDaysPermitted = rsGetCRMMaxActivityDaysPermitted("CRMMaxActivityDaysPermitted")


    GetCRMMaxActivityDaysPermitted = resultGetCRMMaxActivityDaysPermitted 
                
End Function


Function GetStageReasonByStageIntRecID (passedStageIntRecID)

	resultGetStageReasonByStageIntRecID = ""

	Set cnnGetStageReasonByStageIntRecID = Server.CreateObject("ADODB.Connection")
	cnnGetStageReasonByStageIntRecID.open Session("ClientCnnString")
		
	SQLGetStageReasonByStageIntRecID = "Select ReasonRecID from PR_ProspectReasons Where ProspectStagesRecID = " & passedStageIntRecID
 
	Set rsGetStageReasonByStageIntRecID = Server.CreateObject("ADODB.Recordset")
	rsGetStageReasonByStageIntRecID.CursorLocation = 3 
	Set rsGetStageReasonByStageIntRecID = cnnGetStageReasonByStageIntRecID.Execute(SQLGetStageReasonByStageIntRecID)
			 
	If not rsGetStageReasonByStageIntRecID.EOF Then resultGetStageReasonByStageIntRecID =  GetReasonByNum(rsGetStageReasonByStageIntRecID("ReasonRecID"))
	
	rsGetStageReasonByStageIntRecID.Close
	set rsGetStageReasonByStageIntRecID= Nothing
	cnnGetStageReasonByStageIntRecID.Close	
	set cnnGetStageReasonByStageIntRecID= Nothing
	
	GetStageReasonByStageIntRecID = resultGetStageReasonByStageIntRecID
	
End Function



Function TotalNumberOfPreexistingProspectsWeeklySnapshot(passedStartDate,passedEndDate,passedUserNos)

	Set cnnTotalNumberOfPreexistingProspectsWeeklySnapshot = Server.CreateObject("ADODB.Connection")
	cnnTotalNumberOfPreexistingProspectsWeeklySnapshot.open Session("ClientCnnString")

	resultTotalNumberOfPreexistingProspectsWeeklySnapshot = 0
	
	SQLProspects = "SELECT COUNT(PR_Prospects.CreatedByUserNo) AS SalesRepCount"
	SQLProspects = SQLProspects & " FROM PR_Prospects "
	SQLProspects = SQLProspects & " WHERE "
	SQLProspects = SQLProspects & " CAST(PR_Prospects.CreatedDate AS DATE) < '" & passedStartDate & "' AND PR_Prospects.Pool = 'Live' "
	SQLProspects = SQLProspects & " AND PR_Prospects.OwnerUserNo IN (" & passedUserNos & ")"

	'Response.Write(SQLProspects)
				 
	Set rsTotalNumberOfPreexistingProspectsWeeklySnapshot = Server.CreateObject("ADODB.Recordset")
	rsTotalNumberOfPreexistingProspectsWeeklySnapshot.CursorLocation = 3 
	
	Set rsTotalNumberOfPreexistingProspectsWeeklySnapshot = cnnTotalNumberOfPreexistingProspectsWeeklySnapshot.Execute(SQLProspects)
			
	resultTotalNumberOfPreexistingProspectsWeeklySnapshot = rsTotalNumberOfPreexistingProspectsWeeklySnapshot("SalesRepCount")
	
	rsTotalNumberOfPreexistingProspectsWeeklySnapshot.Close
	set rsTotalNumberOfPreexistingProspectsWeeklySnapshot= Nothing
	cnnTotalNumberOfPreexistingProspectsWeeklySnapshot.Close	
	set cnnTotalNumberOfPreexistingProspectsWeeklySnapshot= Nothing
	
	TotalNumberOfPreexistingProspectsWeeklySnapshot = resultTotalNumberOfPreexistingProspectsWeeklySnapshot
	
End Function



Function TotalNumberOfCreatedProspectsWeeklySnapshot(passedStartDate,passedEndDate,passedUserNos)

	Set cnnTotalNumberOfCreatedProspectsWeeklySnapshot = Server.CreateObject("ADODB.Connection")
	cnnTotalNumberOfCreatedProspectsWeeklySnapshot.open Session("ClientCnnString")

	resultTotalNumberOfCreatedProspectsWeeklySnapshot = 0
	
	SQLProspects = "SELECT COUNT(PR_Prospects.CreatedByUserNo) AS SalesRepCount "
	SQLProspects = SQLProspects & " FROM  PR_Prospects"
	SQLProspects = SQLProspects & " WHERE (OwnerUserNo IN (" & passedUserNos & ") AND "
	SQLProspects = SQLProspects & " (CAST(CreatedDate AS DATE) >= '" & passedStartDate & "') AND (CAST(CreatedDate AS DATE) <='" & passedEndDate & "')) "
				
	Set rsTotalNumberOfCreatedProspectsWeeklySnapshot = Server.CreateObject("ADODB.Recordset")
	rsTotalNumberOfCreatedProspectsWeeklySnapshot.CursorLocation = 3 
	
	Set rsTotalNumberOfCreatedProspectsWeeklySnapshot = cnnTotalNumberOfCreatedProspectsWeeklySnapshot.Execute(SQLProspects)
			
	resultTotalNumberOfCreatedProspectsWeeklySnapshot = rsTotalNumberOfCreatedProspectsWeeklySnapshot("SalesRepCount")
	
	rsTotalNumberOfCreatedProspectsWeeklySnapshot.Close
	set rsTotalNumberOfCreatedProspectsWeeklySnapshot= Nothing
	cnnTotalNumberOfCreatedProspectsWeeklySnapshot.Close	
	set cnnTotalNumberOfCreatedProspectsWeeklySnapshot= Nothing
	
	TotalNumberOfCreatedProspectsWeeklySnapshot = resultTotalNumberOfCreatedProspectsWeeklySnapshot
	
End Function



Function TotalNumberOfWonProspectsWeeklySnapshot(passedStartDate,passedEndDate,passedUserNos)

	Set cnnTotalNumberOfWonProspectsWeeklySnapshot = Server.CreateObject("ADODB.Connection")
	cnnTotalNumberOfWonProspectsWeeklySnapshot.open Session("ClientCnnString")
	
	Set rsTotalNumberOfWonProspectsWeeklySnapshot = Server.CreateObject("ADODB.Recordset")
	rsTotalNumberOfWonProspectsWeeklySnapshot.CursorLocation = 3 

	resultTotalNumberOfWonProspectsWeeklySnapshot = 0
			
	SQLProspectsInner = "SELECT Count(*) As ProspCount FROM PR_ProspectStages INNER JOIN PR_Prospects ON"
	SQLProspectsInner = SQLProspectsInner & " PR_Prospects.InternalRecordIdentifier = PR_ProspectStages.ProspectRecID "
	SQLProspectsInner = SQLProspectsInner & " WHERE "
	SQLProspectsInner = SQLProspectsInner & " (Cast(PR_ProspectStages.RecordCreationDateTime as Date) >= '" & passedStartDate & "')"
	SQLProspectsInner = SQLProspectsInner & " AND (Cast(PR_ProspectStages.RecordCreationDateTime as Date) <= '" & passedEndDate& "')"
	SQLProspectsInner = SQLProspectsInner & " AND  StageRecID = 2 "
	SQLProspectsInner = SQLProspectsInner & " AND  PR_Prospects.OwnerUserNo IN (" & passedUserNos & ") "
	
	'Response.Write(SQLProspectsInner & "<br>")

	Set rsTotalNumberOfWonProspectsWeeklySnapshot = cnnTotalNumberOfWonProspectsWeeklySnapshot.Execute(SQLProspectsInner)
	
	If NOT rsTotalNumberOfWonProspectsWeeklySnapshot.EOF Then resultTotalNumberOfWonProspectsWeeklySnapshot = rsTotalNumberOfWonProspectsWeeklySnapshot("ProspCount") 
			
	set rsTotalNumberOfWonProspectsWeeklySnapshot= Nothing
	cnnTotalNumberOfWonProspectsWeeklySnapshot.Close	
	set cnnTotalNumberOfWonProspectsWeeklySnapshot= Nothing
	
	TotalNumberOfWonProspectsWeeklySnapshot = resultTotalNumberOfWonProspectsWeeklySnapshot
	
End Function


Function TotalNumberOfLostProspectsWeeklySnapshot(passedStartDate,passedEndDate,passedUserNos)

	Set cnnTotalNumberOfLostProspectsWeeklySnapshot = Server.CreateObject("ADODB.Connection")
	cnnTotalNumberOfLostProspectsWeeklySnapshot.open Session("ClientCnnString")
	
	Set rsTotalNumberOfLostProspectsWeeklySnapshot = Server.CreateObject("ADODB.Recordset")
	rsTotalNumberOfLostProspectsWeeklySnapshot.CursorLocation = 3 

	resultTotalNumberOfLostProspectsWeeklySnapshot = 0
			
	SQLProspectsInner = "SELECT Count(*) As ProspCount FROM PR_ProspectStages INNER JOIN PR_Prospects ON"
	SQLProspectsInner = SQLProspectsInner & " PR_Prospects.InternalRecordIdentifier = PR_ProspectStages.ProspectRecID "
	SQLProspectsInner = SQLProspectsInner & " WHERE "
	SQLProspectsInner = SQLProspectsInner & " (CAST(PR_ProspectStages.RecordCreationDateTime AS DATE) >= '" & passedStartDate & "')"
	SQLProspectsInner = SQLProspectsInner & " AND (CAST(PR_ProspectStages.RecordCreationDateTime AS DATE) <= '" & passedEndDate& "')"
	SQLProspectsInner = SQLProspectsInner & " AND  StageRecID = 1 "
	SQLProspectsInner = SQLProspectsInner & " AND  PR_Prospects.OwnerUserNo IN (" & passedUserNos & ") "
	
	'Response.Write(SQLProspectsInner & "<br>")
			
	Set rsTotalNumberOfLostProspectsWeeklySnapshot = cnnTotalNumberOfLostProspectsWeeklySnapshot.Execute(SQLProspectsInner)
	
	If NOT rsTotalNumberOfLostProspectsWeeklySnapshot.EOF Then resultTotalNumberOfLostProspectsWeeklySnapshot = rsTotalNumberOfLostProspectsWeeklySnapshot("ProspCount") 
			
	set rsTotalNumberOfLostProspectsWeeklySnapshot= Nothing
	cnnTotalNumberOfLostProspectsWeeklySnapshot.Close	
	set cnnTotalNumberOfLostProspectsWeeklySnapshot= Nothing
	
	TotalNumberOfLostProspectsWeeklySnapshot = resultTotalNumberOfLostProspectsWeeklySnapshot
	
End Function



Function TotalNumberOfUnqualifiedProspectsWeeklySnapshot(passedStartDate,passedEndDate,passedUserNos)

	Set cnnTotalNumberOfUnqualifiedProspectsWeeklySnapshot = Server.CreateObject("ADODB.Connection")
	cnnTotalNumberOfUnqualifiedProspectsWeeklySnapshot.open Session("ClientCnnString")
	
	Set rsTotalNumberOfUnqualifiedProspectsWeeklySnapshot = Server.CreateObject("ADODB.Recordset")
	rsTotalNumberOfUnqualifiedProspectsWeeklySnapshot.CursorLocation = 3 

	resultTotalNumberOfUnqualifiedProspectsWeeklySnapshot = 0
			
	SQLProspectsInner = "SELECT Count(*) As ProspCount FROM PR_ProspectStages INNER JOIN PR_Prospects ON"
	SQLProspectsInner = SQLProspectsInner & " PR_Prospects.InternalRecordIdentifier = PR_ProspectStages.ProspectRecID "
	SQLProspectsInner = SQLProspectsInner & " WHERE "
	SQLProspectsInner = SQLProspectsInner & " (CAST(PR_ProspectStages.RecordCreationDateTime AS DATE) >= '" & passedStartDate & "')"
	SQLProspectsInner = SQLProspectsInner & " AND (CAST(PR_ProspectStages.RecordCreationDateTime AS DATE) <= '" & passedEndDate& "')"
	SQLProspectsInner = SQLProspectsInner & " AND  StageRecID = 0 "
	SQLProspectsInner = SQLProspectsInner & " AND  PR_Prospects.OwnerUserNo IN (" & passedUserNos & ") "
			
	Set rsTotalNumberOfUnqualifiedProspectsWeeklySnapshot = cnnTotalNumberOfUnqualifiedProspectsWeeklySnapshot.Execute(SQLProspectsInner)
	
	If NOT rsTotalNumberOfUnqualifiedProspectsWeeklySnapshot.EOF Then resultTotalNumberOfUnqualifiedProspectsWeeklySnapshot = rsTotalNumberOfUnqualifiedProspectsWeeklySnapshot("ProspCount") 
			
	set rsTotalNumberOfUnqualifiedProspectsWeeklySnapshot= Nothing
	cnnTotalNumberOfUnqualifiedProspectsWeeklySnapshot.Close	
	set cnnTotalNumberOfUnqualifiedProspectsWeeklySnapshot= Nothing
	
	TotalNumberOfUnqualifiedProspectsWeeklySnapshot = resultTotalNumberOfUnqualifiedProspectsWeeklySnapshot
	
End Function



Function TotalNumberOfPreexistingAppointmentsWeeklySnapshot(passedStartDate,passedEndDate,passedUserNos)

	Set cnnTotalNumberOfPreexistingAppointmentsWeeklySnapshot = Server.CreateObject("ADODB.Connection")
	cnnTotalNumberOfPreexistingAppointmentsWeeklySnapshot.open Session("ClientCnnString")

	resultTotalNumberOfPreexistingAppointmentsWeeklySnapshot = 0
	
	SQLProspects = "SELECT COUNT(PR_ProspectActivities.ActivityCreatedByUserNo) AS AppmtCount"
	SQLProspects = SQLProspects & " FROM  PR_ProspectActivities INNER JOIN "
	SQLProspects = SQLProspects & " PR_Prospects ON PR_ProspectActivities.ProspectRecID = PR_Prospects.InternalRecordIdentifier INNER JOIN" 
	SQLProspects = SQLProspects & " tblUsers ON PR_Prospects.OwnerUserNo = tblUsers.UserNo " 
	SQLProspects = SQLProspects & " WHERE tblUsers.userNo IN (" & passedUserNos & ") "
	SQLProspects = SQLProspects & " AND (PR_ProspectActivities.ActivityIsAppointment = 1 OR PR_ProspectActivities.ActivityIsMeeting = 1) "
	SQLProspects = SQLProspects & " AND PR_ProspectActivities.Status IS NULL "
	SQLProspects = SQLProspects & " AND CAST(PR_ProspectActivities.RecordCreationDateTime AS DATE) < '" & passedStartDate & "'"

	
	'Response.Write(SQLProspects & "<br><br>")

	Set rsTotalNumberOfPreexistingAppointmentsWeeklySnapshot = Server.CreateObject("ADODB.Recordset")
	rsTotalNumberOfPreexistingAppointmentsWeeklySnapshot.CursorLocation = 3 
	
	Set rsTotalNumberOfPreexistingAppointmentsWeeklySnapshot = cnnTotalNumberOfPreexistingAppointmentsWeeklySnapshot.Execute(SQLProspects)
			
	resultTotalNumberOfPreexistingAppointmentsWeeklySnapshot = rsTotalNumberOfPreexistingAppointmentsWeeklySnapshot("AppmtCount")
	
	rsTotalNumberOfPreexistingAppointmentsWeeklySnapshot.Close
	set rsTotalNumberOfPreexistingAppointmentsWeeklySnapshot= Nothing
	cnnTotalNumberOfPreexistingAppointmentsWeeklySnapshot.Close	
	set cnnTotalNumberOfPreexistingAppointmentsWeeklySnapshot= Nothing
	
	TotalNumberOfPreexistingAppointmentsWeeklySnapshot = resultTotalNumberOfPreexistingAppointmentsWeeklySnapshot
	
End Function



Function TotalNumberOfNotUpdatedAppointmentsWeeklySnapshot(passedStartDate,passedEndDate,passedUserNos)

	Set cnnTotalNumberOfNotUpdatedAppointmentsWeeklySnapshot = Server.CreateObject("ADODB.Connection")
	cnnTotalNumberOfNotUpdatedAppointmentsWeeklySnapshot.open Session("ClientCnnString")

	resultTotalNumberOfNotUpdatedAppointmentsWeeklySnapshot = 0
	
	SQLProspects = "SELECT COUNT(PR_ProspectActivities.ActivityCreatedByUserNo) AS AppmtCount"
	SQLProspects = SQLProspects & " FROM  PR_ProspectActivities INNER JOIN "
	SQLProspects = SQLProspects & " PR_Prospects ON PR_ProspectActivities.ProspectRecID = PR_Prospects.InternalRecordIdentifier " 
	SQLProspects = SQLProspects & " AND PR_Prospects.OwnerUserNo IN (" & passedUserNos & ") "
	SQLProspects = SQLProspects & " AND (PR_ProspectActivities.ActivityIsAppointment = 1 OR PR_ProspectActivities.ActivityIsMeeting = 1) "
	SQLProspects = SQLProspects & " AND PR_ProspectActivities.Status IS NULL "
	SQLProspects = SQLProspects & " AND CAST(PR_ProspectActivities.RecordCreationDateTime AS DATE) >= '" & passedStartDate & "' AND CAST(PR_ProspectActivities.RecordCreationDateTime AS DATE) <= '" & passedEndDate & "' "
		
	'Response.Write(SQLProspects & "<br><br>")

	Set rsTotalNumberOfNotUpdatedAppointmentsWeeklySnapshot = Server.CreateObject("ADODB.Recordset")
	rsTotalNumberOfNotUpdatedAppointmentsWeeklySnapshot.CursorLocation = 3 
	
	Set rsTotalNumberOfNotUpdatedAppointmentsWeeklySnapshot = cnnTotalNumberOfNotUpdatedAppointmentsWeeklySnapshot.Execute(SQLProspects)
			
	resultTotalNumberOfNotUpdatedAppointmentsWeeklySnapshot = rsTotalNumberOfNotUpdatedAppointmentsWeeklySnapshot("AppmtCount")
	
	rsTotalNumberOfNotUpdatedAppointmentsWeeklySnapshot.Close
	set rsTotalNumberOfNotUpdatedAppointmentsWeeklySnapshot= Nothing
	cnnTotalNumberOfNotUpdatedAppointmentsWeeklySnapshot.Close	
	set cnnTotalNumberOfNotUpdatedAppointmentsWeeklySnapshot= Nothing
	
	TotalNumberOfNotUpdatedAppointmentsWeeklySnapshot = resultTotalNumberOfNotUpdatedAppointmentsWeeklySnapshot
	
End Function




Function TotalNumberOfExpiredAppointmentsWeeklySnapshot(passedStartDate,passedEndDate,passedUserNos)

	Set cnnTotalNumberOfExpiredAppointmentsWeeklySnapshot = Server.CreateObject("ADODB.Connection")
	cnnTotalNumberOfExpiredAppointmentsWeeklySnapshot.open Session("ClientCnnString")

	resultTotalNumberOfExpiredAppointmentsWeeklySnapshot = 0
	
	SQLProspects = "SELECT COUNT(PR_ProspectActivities.ActivityCreatedByUserNo) AS AppmtCount"
	SQLProspects = SQLProspects & " FROM  PR_ProspectActivities INNER JOIN "
	SQLProspects = SQLProspects & " PR_Prospects ON PR_ProspectActivities.ProspectRecID = PR_Prospects.InternalRecordIdentifier " 
	SQLProspects = SQLProspects & " AND PR_Prospects.OwnerUserNo IN (" & passedUserNos & ") "
	SQLProspects = SQLProspects & " AND (PR_ProspectActivities.ActivityIsAppointment = 1 OR PR_ProspectActivities.ActivityIsMeeting = 1) "
	SQLProspects = SQLProspects & " AND PR_ProspectActivities.Status IS NULL "
	SQLProspects = SQLProspects & " AND CAST(PR_ProspectActivities.ActivityDueDate AS DATE) >= '" & passedStartDate & "' AND CAST(PR_ProspectActivities.ActivityDueDate AS DATE) <= '" & passedEndDate & "' "

		
	'Response.Write(SQLProspects & "<br><br>")

	Set rsTotalNumberOfExpiredAppointmentsWeeklySnapshot = Server.CreateObject("ADODB.Recordset")
	rsTotalNumberOfExpiredAppointmentsWeeklySnapshot.CursorLocation = 3 
	
	Set rsTotalNumberOfExpiredAppointmentsWeeklySnapshot = cnnTotalNumberOfExpiredAppointmentsWeeklySnapshot.Execute(SQLProspects)
			
	resultTotalNumberOfExpiredAppointmentsWeeklySnapshot = rsTotalNumberOfExpiredAppointmentsWeeklySnapshot("AppmtCount")
	
	rsTotalNumberOfExpiredAppointmentsWeeklySnapshot.Close
	set rsTotalNumberOfExpiredAppointmentsWeeklySnapshot= Nothing
	cnnTotalNumberOfExpiredAppointmentsWeeklySnapshot.Close	
	set cnnTotalNumberOfExpiredAppointmentsWeeklySnapshot= Nothing
	
	TotalNumberOfExpiredAppointmentsWeeklySnapshot = resultTotalNumberOfExpiredAppointmentsWeeklySnapshot
	
End Function



Function TotalNumberOfCreatedAppointmentsWeeklySnapshot(passedStartDate,passedEndDate,passedUserNos)

	Set cnnTotalNumberOfCreatedAppointmentsWeeklySnapshot = Server.CreateObject("ADODB.Connection")
	cnnTotalNumberOfCreatedAppointmentsWeeklySnapshot.open Session("ClientCnnString")

	resultTotalNumberOfCreatedAppointmentsWeeklySnapshot = 0
	
	SQLProspects = "SELECT COUNT(PR_ProspectActivities.ActivityCreatedByUserNo) AS AppmtCount"
	SQLProspects = SQLProspects & " FROM  PR_ProspectActivities INNER JOIN "
	SQLProspects = SQLProspects & " PR_Prospects ON PR_ProspectActivities.ProspectRecID = PR_Prospects.InternalRecordIdentifier "
	SQLProspects = SQLProspects & " AND PR_Prospects.OwnerUserNo IN (" & passedUserNos & ") "
	SQLProspects = SQLProspects & " AND (PR_ProspectActivities.ActivityIsAppointment = 1 OR PR_ProspectActivities.ActivityIsMeeting = 1) "
	SQLProspects = SQLProspects & " AND CAST(PR_ProspectActivities.RecordCreationDateTime AS DATE) >= '" & passedStartDate & "' AND CAST(PR_ProspectActivities.RecordCreationDateTime AS DATE) <= '" & passedEndDate & "' "
	
	'Response.Write(SQLProspects & "<br><br>")

	Set rsTotalNumberOfCreatedAppointmentsWeeklySnapshot = Server.CreateObject("ADODB.Recordset")
	rsTotalNumberOfCreatedAppointmentsWeeklySnapshot.CursorLocation = 3 
	
	Set rsTotalNumberOfCreatedAppointmentsWeeklySnapshot = cnnTotalNumberOfCreatedAppointmentsWeeklySnapshot.Execute(SQLProspects)
			
	resultTotalNumberOfCreatedAppointmentsWeeklySnapshot = rsTotalNumberOfCreatedAppointmentsWeeklySnapshot("AppmtCount")
	
	rsTotalNumberOfCreatedAppointmentsWeeklySnapshot.Close
	set rsTotalNumberOfCreatedAppointmentsWeeklySnapshot= Nothing
	cnnTotalNumberOfCreatedAppointmentsWeeklySnapshot.Close	
	set cnnTotalNumberOfCreatedAppointmentsWeeklySnapshot= Nothing
	
	TotalNumberOfCreatedAppointmentsWeeklySnapshot = resultTotalNumberOfCreatedAppointmentsWeeklySnapshot
	
End Function




Function TotalNumberOfCreatedAppointmentsPreexistingWeeklySnapshot(passedStartDate,passedEndDate,passedUserNos)

	Set cnnTotalNumberOfCreatedAppointmentsPreexistingWeeklySnapshot = Server.CreateObject("ADODB.Connection")
	cnnTotalNumberOfCreatedAppointmentsPreexistingWeeklySnapshot.open Session("ClientCnnString")

	resultTotalNumberOfCreatedAppointmentsPreexistingWeeklySnapshot = 0


	SQLProspects = "SELECT COUNT(PR_ProspectActivities.ActivityCreatedByUserNo) AS AppmtCount"
	SQLProspects = SQLProspects & " FROM  PR_ProspectActivities INNER JOIN "
	SQLProspects = SQLProspects & " PR_Prospects ON PR_ProspectActivities.ProspectRecID = PR_Prospects.InternalRecordIdentifier " 
	SQLProspects = SQLProspects & " WHERE "
	SQLProspects = SQLProspects & " (PR_ProspectActivities.ActivityIsAppointment = 1 OR PR_ProspectActivities.ActivityIsMeeting = 1) "
	SQLProspects = SQLProspects & " AND CAST(PR_ProspectActivities.RecordCreationDateTime AS DATE) >= '" & passedStartDate & "' AND CAST(PR_ProspectActivities.RecordCreationDateTime AS DATE) <= '" & passedEndDate & "' "
	SQLProspects = SQLProspects & " AND PR_Prospects.InternalRecordIdentifier IN ( "
	SQLProspects = SQLProspects & " SELECT PR_Prospects.InternalrecordIdentifier "
	SQLProspects = SQLProspects & " FROM  PR_Prospects"
	SQLProspects = SQLProspects & " WHERE (OwnerUserNo IN (" & passedUserNos & ") AND (CAST(CreatedDate AS DATE) < '" & passedStartDate & "')) "
	SQLProspects = SQLProspects & " ) "		

	'Response.Write(SQLProspects & "<br><br>")

	Set rsTotalNumberOfCreatedAppointmentsPreexistingWeeklySnapshot = Server.CreateObject("ADODB.Recordset")
	rsTotalNumberOfCreatedAppointmentsPreexistingWeeklySnapshot.CursorLocation = 3 
	
	Set rsTotalNumberOfCreatedAppointmentsPreexistingWeeklySnapshot = cnnTotalNumberOfCreatedAppointmentsPreexistingWeeklySnapshot.Execute(SQLProspects)
			
	resultTotalNumberOfCreatedAppointmentsPreexistingWeeklySnapshot = rsTotalNumberOfCreatedAppointmentsPreexistingWeeklySnapshot("AppmtCount")
	
	rsTotalNumberOfCreatedAppointmentsPreexistingWeeklySnapshot.Close
	set rsTotalNumberOfCreatedAppointmentsPreexistingWeeklySnapshot= Nothing
	cnnTotalNumberOfCreatedAppointmentsPreexistingWeeklySnapshot.Close	
	set cnnTotalNumberOfCreatedAppointmentsPreexistingWeeklySnapshot= Nothing
	
	TotalNumberOfCreatedAppointmentsPreexistingWeeklySnapshot = resultTotalNumberOfCreatedAppointmentsPreexistingWeeklySnapshot
	
End Function





Function TotalNumberOfCompletedAppointmentsWeeklySnapshot(passedStartDate,passedEndDate,passedUserNos)

	Set cnnTotalNumberOfCompletedAppointmentsWeeklySnapshot = Server.CreateObject("ADODB.Connection")
	cnnTotalNumberOfCompletedAppointmentsWeeklySnapshot.open Session("ClientCnnString")

	resultTotalNumberOfCompletedAppointmentsWeeklySnapshot = 0

	SQLProspects = "SELECT COUNT(PR_ProspectActivities.ActivityCreatedByUserNo) AS AppmtCount"
	SQLProspects = SQLProspects & " FROM  PR_ProspectActivities INNER JOIN "
	SQLProspects = SQLProspects & " PR_Prospects ON PR_ProspectActivities.ProspectRecID = PR_Prospects.InternalRecordIdentifier " 
	SQLProspects = SQLProspects & " AND PR_Prospects.OwnerUserNo IN (" & passedUserNos & ") "
	SQLProspects = SQLProspects & " AND (PR_ProspectActivities.ActivityIsAppointment = 1 OR PR_ProspectActivities.ActivityIsMeeting = 1)"
	SQLProspects = SQLProspects & " AND PR_ProspectActivities.Status='Completed' "
	SQLProspects = SQLProspects & " AND CAST(PR_ProspectActivities.StatusDateTime AS DATE) >= '" & passedStartDate & "' AND CAST(PR_ProspectActivities.StatusDateTime AS DATE) <= '" & passedEndDate & "' "
	SQLProspects = SQLProspects & " AND CAST(PR_ProspectActivities.ActivityDueDate AS DATE) >= '" & passedStartDate & "' AND CAST(PR_ProspectActivities.ActivityDueDate AS DATE) <= '" & passedEndDate & "' "
		
	'Response.Write(SQLProspects & "<br><br>")

	Set rsTotalNumberOfCompletedAppointmentsWeeklySnapshot = Server.CreateObject("ADODB.Recordset")
	rsTotalNumberOfCompletedAppointmentsWeeklySnapshot.CursorLocation = 3 
	
	Set rsTotalNumberOfCompletedAppointmentsWeeklySnapshot = cnnTotalNumberOfCompletedAppointmentsWeeklySnapshot.Execute(SQLProspects)
			
	resultTotalNumberOfCompletedAppointmentsWeeklySnapshot = rsTotalNumberOfCompletedAppointmentsWeeklySnapshot("AppmtCount")
	
	rsTotalNumberOfCompletedAppointmentsWeeklySnapshot.Close
	set rsTotalNumberOfCompletedAppointmentsWeeklySnapshot= Nothing
	cnnTotalNumberOfCompletedAppointmentsWeeklySnapshot.Close	
	set cnnTotalNumberOfCompletedAppointmentsWeeklySnapshot= Nothing
	
	TotalNumberOfCompletedAppointmentsWeeklySnapshot = resultTotalNumberOfCompletedAppointmentsWeeklySnapshot
	
End Function


Function TotalNumberOfCompletedAppointmentsPreexistingWeeklySnapshot(passedStartDate,passedEndDate,passedUserNos)

	Set cnnTotalNumberOfCompletedAppointmentsPreexistingWeeklySnapshot = Server.CreateObject("ADODB.Connection")
	cnnTotalNumberOfCompletedAppointmentsPreexistingWeeklySnapshot.open Session("ClientCnnString")

	resultTotalNumberOfCompletedAppointmentsPreexistingWeeklySnapshot = 0

	SQLProspects = "SELECT COUNT(PR_ProspectActivities.ActivityCreatedByUserNo) AS AppmtCount"
	SQLProspects = SQLProspects & " FROM  PR_ProspectActivities INNER JOIN "
	SQLProspects = SQLProspects & " PR_Prospects ON PR_ProspectActivities.ProspectRecID = PR_Prospects.InternalRecordIdentifier " 
	SQLProspects = SQLProspects & " WHERE "
	SQLProspects = SQLProspects & " (PR_ProspectActivities.ActivityIsAppointment = 1 OR PR_ProspectActivities.ActivityIsMeeting = 1) "
	SQLProspects = SQLProspects & " AND PR_ProspectActivities.Status='Completed' "
	SQLProspects = SQLProspects & " AND CAST(PR_ProspectActivities.RecordCreationDateTime AS DATE) >= '" & passedStartDate & "' AND CAST(PR_ProspectActivities.RecordCreationDateTime AS DATE) <= '" & passedEndDate & "' "

	SQLProspects = SQLProspects & " AND PR_Prospects.InternalRecordIdentifier IN ( "
	SQLProspects = SQLProspects &  "SELECT InternalRecordIdentifier  "
	SQLProspects = SQLProspects & " FROM  PR_Prospects"
	SQLProspects = SQLProspects & " WHERE (OwnerUserNo IN (" & passedUserNos & ") AND (CAST(CreatedDate AS DATE) < '" & passedStartDate & "') AND (CAST(CreatedDate AS DATE) <='" & passedEndDate & "')) "
	SQLProspects = SQLProspects & " ) "
		
	'Response.Write(SQLProspects & "<br><br>")

	Set rsTotalNumberOfCompletedAppointmentsPreexistingWeeklySnapshot = Server.CreateObject("ADODB.Recordset")
	rsTotalNumberOfCompletedAppointmentsPreexistingWeeklySnapshot.CursorLocation = 3 
	
	Set rsTotalNumberOfCompletedAppointmentsPreexistingWeeklySnapshot = cnnTotalNumberOfCompletedAppointmentsPreexistingWeeklySnapshot.Execute(SQLProspects)
			
	resultTotalNumberOfCompletedAppointmentsPreexistingWeeklySnapshot = rsTotalNumberOfCompletedAppointmentsPreexistingWeeklySnapshot("AppmtCount")
	
	rsTotalNumberOfCompletedAppointmentsPreexistingWeeklySnapshot.Close
	set rsTotalNumberOfCompletedAppointmentsPreexistingWeeklySnapshot= Nothing
	cnnTotalNumberOfCompletedAppointmentsPreexistingWeeklySnapshot.Close	
	set cnnTotalNumberOfCompletedAppointmentsPreexistingWeeklySnapshot= Nothing
	
	TotalNumberOfCompletedAppointmentsPreexistingWeeklySnapshot = resultTotalNumberOfCompletedAppointmentsPreexistingWeeklySnapshot
	
End Function



Function TotalNumberOfRescheduledAppointmentsWeeklySnapshot(passedStartDate,passedEndDate,passedUserNos)

	Set cnnTotalNumberOfRescheduledAppointmentsWeeklySnapshot = Server.CreateObject("ADODB.Connection")
	cnnTotalNumberOfRescheduledAppointmentsWeeklySnapshot.open Session("ClientCnnString")

	resultTotalNumberOfRescheduledAppointmentsWeeklySnapshot = 0
	
	SQLProspects = "SELECT COUNT(PR_ProspectActivities.ActivityCreatedByUserNo) AS AppmtCount"
	SQLProspects = SQLProspects & " FROM  PR_ProspectActivities INNER JOIN "
	SQLProspects = SQLProspects & " PR_Prospects ON PR_ProspectActivities.ProspectRecID = PR_Prospects.InternalRecordIdentifier " 
	SQLProspects = SQLProspects & " AND PR_Prospects.OwnerUserNo  IN (" & passedUserNos & ") "
	SQLProspects = SQLProspects & " AND (PR_ProspectActivities.ActivityIsAppointment = 1 OR PR_ProspectActivities.ActivityIsMeeting = 1) "
	SQLProspects = SQLProspects & " AND PR_ProspectActivities.Status='Rescheduled' "
	SQLProspects = SQLProspects & " AND CAST(PR_ProspectActivities.StatusDateTime AS DATE) >= '" & passedStartDate & "' AND CAST(PR_ProspectActivities.StatusDateTime AS DATE) <= '" & passedEndDate & "' "
	SQLProspects = SQLProspects & " AND CAST(PR_ProspectActivities.ActivityDueDate AS DATE) >= '" & passedStartDate & "' AND CAST(PR_ProspectActivities.ActivityDueDate AS DATE) <= '" & passedEndDate & "' "
		
	'Response.Write(SQLProspects & "<br><br>")

	Set rsTotalNumberOfRescheduledAppointmentsWeeklySnapshot = Server.CreateObject("ADODB.Recordset")
	rsTotalNumberOfRescheduledAppointmentsWeeklySnapshot.CursorLocation = 3 
	
	Set rsTotalNumberOfRescheduledAppointmentsWeeklySnapshot = cnnTotalNumberOfRescheduledAppointmentsWeeklySnapshot.Execute(SQLProspects)
			
	resultTotalNumberOfRescheduledAppointmentsWeeklySnapshot = rsTotalNumberOfRescheduledAppointmentsWeeklySnapshot("AppmtCount")
	
	rsTotalNumberOfRescheduledAppointmentsWeeklySnapshot.Close
	set rsTotalNumberOfRescheduledAppointmentsWeeklySnapshot= Nothing
	cnnTotalNumberOfRescheduledAppointmentsWeeklySnapshot.Close	
	set cnnTotalNumberOfRescheduledAppointmentsWeeklySnapshot= Nothing
	
	TotalNumberOfRescheduledAppointmentsWeeklySnapshot = resultTotalNumberOfRescheduledAppointmentsWeeklySnapshot
	
End Function





Function TotalNumberOfRescheduledAppointmentsPreexistingWeeklySnapshot(passedStartDate,passedEndDate,passedUserNos)

	Set cnnTotalNumberOfRescheduledAppointmentsPreexistingWeeklySnapshot = Server.CreateObject("ADODB.Connection")
	cnnTotalNumberOfRescheduledAppointmentsPreexistingWeeklySnapshot.open Session("ClientCnnString")

	resultTotalNumberOfRescheduledAppointmentsPreexistingWeeklySnapshot = 0
	
	SQLProspects = "SELECT COUNT(PR_ProspectActivities.ActivityCreatedByUserNo) AS AppmtCount"
	SQLProspects = SQLProspects & " FROM  PR_ProspectActivities INNER JOIN "
	SQLProspects = SQLProspects & " PR_Prospects ON PR_ProspectActivities.ProspectRecID = PR_Prospects.InternalRecordIdentifier INNER JOIN" 
	SQLProspects = SQLProspects & " tblUsers ON PR_Prospects.OwnerUserNo = tblUsers.UserNo " 
	SQLProspects = SQLProspects & " WHERE tblUsers.userNo IN (" & passedUserNos & ") "
	SQLProspects = SQLProspects & " AND (CAST(PR_Prospects.CreatedDate AS DATE) < '" & passedStartDate & "') "
	'SQLProspects = SQLProspects & " AND (PR_Prospects.Pool = 'Live') "
	SQLProspects = SQLProspects & " AND (PR_ProspectActivities.ActivityIsAppointment = 1 OR PR_ProspectActivities.ActivityIsMeeting = 1) "
	SQLProspects = SQLProspects & " AND PR_ProspectActivities.Status='Rescheduled' "
	SQLProspects = SQLProspects & " AND CAST(PR_ProspectActivities.RecordCreationDateTime AS DATE) >= '" & passedStartDate & "' AND CAST(PR_ProspectActivities.RecordCreationDateTime AS DATE) <= '" & passedEndDate & "' "
		
	'Response.Write(SQLProspects & "<br><br>")

	Set rsTotalNumberOfRescheduledAppointmentsPreexistingWeeklySnapshot = Server.CreateObject("ADODB.Recordset")
	rsTotalNumberOfRescheduledAppointmentsPreexistingWeeklySnapshot.CursorLocation = 3 
	
	Set rsTotalNumberOfRescheduledAppointmentsPreexistingWeeklySnapshot = cnnTotalNumberOfRescheduledAppointmentsPreexistingWeeklySnapshot.Execute(SQLProspects)
			
	resultTotalNumberOfRescheduledAppointmentsPreexistingWeeklySnapshot = rsTotalNumberOfRescheduledAppointmentsPreexistingWeeklySnapshot("AppmtCount")
	
	rsTotalNumberOfRescheduledAppointmentsPreexistingWeeklySnapshot.Close
	set rsTotalNumberOfRescheduledAppointmentsPreexistingWeeklySnapshot= Nothing
	cnnTotalNumberOfRescheduledAppointmentsPreexistingWeeklySnapshot.Close	
	set cnnTotalNumberOfRescheduledAppointmentsPreexistingWeeklySnapshot= Nothing
	
	TotalNumberOfRescheduledAppointmentsPreexistingWeeklySnapshot = resultTotalNumberOfRescheduledAppointmentsPreexistingWeeklySnapshot
	
End Function


Function TotalNumberOfCancelledAppointmentsWeeklySnapshot(passedStartDate,passedEndDate,passedUserNos)

	Set cnnTotalNumberOfCancelledAppointmentsWeeklySnapshot = Server.CreateObject("ADODB.Connection")
	cnnTotalNumberOfCancelledAppointmentsWeeklySnapshot.open Session("ClientCnnString")

	resultTotalNumberOfCancelledAppointmentsWeeklySnapshot = 0
	
	
	SQLProspects = "SELECT COUNT(PR_ProspectActivities.ActivityCreatedByUserNo) AS AppmtCount"
	SQLProspects = SQLProspects & " FROM  PR_ProspectActivities INNER JOIN "
	SQLProspects = SQLProspects & " PR_Prospects ON PR_ProspectActivities.ProspectRecID = PR_Prospects.InternalRecordIdentifier " 
	SQLProspects = SQLProspects & " AND PR_Prospects.OwnerUserNo IN (" & passedUserNos & ") "
	SQLProspects = SQLProspects & " AND (PR_ProspectActivities.ActivityIsAppointment = 1 OR PR_ProspectActivities.ActivityIsMeeting = 1) "
	SQLProspects = SQLProspects & " AND PR_ProspectActivities.Status='Cancelled' "
	SQLProspects = SQLProspects & " AND CAST(PR_ProspectActivities.StatusDateTime AS DATE) >= '" & passedStartDate & "' AND CAST(PR_ProspectActivities.StatusDateTime AS DATE) <= '" & passedEndDate & "' "
	SQLProspects = SQLProspects & " AND CAST(PR_ProspectActivities.ActivityDueDate AS DATE) >= '" & passedStartDate & "' AND CAST(PR_ProspectActivities.ActivityDueDate AS DATE) <= '" & passedEndDate & "' "

		
	'Response.Write(SQLProspects & "<br><br>")

	Set rsTotalNumberOfCancelledAppointmentsWeeklySnapshot = Server.CreateObject("ADODB.Recordset")
	rsTotalNumberOfCancelledAppointmentsWeeklySnapshot.CursorLocation = 3 
	
	Set rsTotalNumberOfCancelledAppointmentsWeeklySnapshot = cnnTotalNumberOfCancelledAppointmentsWeeklySnapshot.Execute(SQLProspects)
			
	resultTotalNumberOfCancelledAppointmentsWeeklySnapshot = rsTotalNumberOfCancelledAppointmentsWeeklySnapshot("AppmtCount")
	
	rsTotalNumberOfCancelledAppointmentsWeeklySnapshot.Close
	set rsTotalNumberOfCancelledAppointmentsWeeklySnapshot= Nothing
	cnnTotalNumberOfCancelledAppointmentsWeeklySnapshot.Close	
	set cnnTotalNumberOfCancelledAppointmentsWeeklySnapshot= Nothing
	
	TotalNumberOfCancelledAppointmentsWeeklySnapshot = resultTotalNumberOfCancelledAppointmentsWeeklySnapshot
	
End Function




Function TotalNumberOfCancelledAppointmentsPreexistingWeeklySnapshot(passedStartDate,passedEndDate,passedUserNos)

	Set cnnTotalNumberOfCancelledAppointmentsPreexistingWeeklySnapshot = Server.CreateObject("ADODB.Connection")
	cnnTotalNumberOfCancelledAppointmentsPreexistingWeeklySnapshot.open Session("ClientCnnString")

	resultTotalNumberOfCancelledAppointmentsPreexistingWeeklySnapshot = 0
	
	SQLProspects = "SELECT COUNT(PR_ProspectActivities.ActivityCreatedByUserNo) AS AppmtCount"
	SQLProspects = SQLProspects & " FROM  PR_ProspectActivities INNER JOIN "
	SQLProspects = SQLProspects & " PR_Prospects ON PR_ProspectActivities.ProspectRecID = PR_Prospects.InternalRecordIdentifier INNER JOIN" 
	SQLProspects = SQLProspects & " tblUsers ON PR_Prospects.OwnerUserNo = tblUsers.UserNo " 
	SQLProspects = SQLProspects & " WHERE tblUsers.userNo IN (" & passedUserNos & ") "
	SQLProspects = SQLProspects & " AND CAST(PR_Prospects.CreatedDate AS DATE) < '" & passedStartDate & "' "
	'SQLProspects = SQLProspects & " AND PR_Prospects.Pool = 'Live' "
	SQLProspects = SQLProspects & " AND (PR_ProspectActivities.ActivityIsAppointment = 1 OR PR_ProspectActivities.ActivityIsMeeting = 1) "
	SQLProspects = SQLProspects & " AND PR_ProspectActivities.Status='Cancelled' "
	SQLProspects = SQLProspects & " AND CAST(PR_ProspectActivities.RecordCreationDateTime AS DATE) >= '" & passedStartDate & "' AND CAST(PR_ProspectActivities.RecordCreationDateTime AS DATE) <= '" & passedEndDate & "' "
		
	'Response.Write(SQLProspects & "<br><br>")

	Set rsTotalNumberOfCancelledAppointmentsPreexistingWeeklySnapshot = Server.CreateObject("ADODB.Recordset")
	rsTotalNumberOfCancelledAppointmentsPreexistingWeeklySnapshot.CursorLocation = 3 
	
	Set rsTotalNumberOfCancelledAppointmentsPreexistingWeeklySnapshot = cnnTotalNumberOfCancelledAppointmentsPreexistingWeeklySnapshot.Execute(SQLProspects)
			
	resultTotalNumberOfCancelledAppointmentsPreexistingWeeklySnapshot = rsTotalNumberOfCancelledAppointmentsPreexistingWeeklySnapshot("AppmtCount")
	
	rsTotalNumberOfCancelledAppointmentsPreexistingWeeklySnapshot.Close
	set rsTotalNumberOfCancelledAppointmentsPreexistingWeeklySnapshot= Nothing
	cnnTotalNumberOfCancelledAppointmentsPreexistingWeeklySnapshot.Close	
	set cnnTotalNumberOfCancelledAppointmentsPreexistingWeeklySnapshot= Nothing
	
	TotalNumberOfCancelledAppointmentsPreexistingWeeklySnapshot = resultTotalNumberOfCancelledAppointmentsPreexistingWeeklySnapshot
	
End Function







Function TotalNumberOfExpiredActivitiesWeeklySnapshot(passedStartDate,passedEndDate,passedUserNos)

	Set cnnTotalNumberOfExpiredActivitiesWeeklySnapshot = Server.CreateObject("ADODB.Connection")
	cnnTotalNumberOfExpiredActivitiesWeeklySnapshot.open Session("ClientCnnString")

	resultTotalNumberOfExpiredActivitiesWeeklySnapshot = 0

	SQLProspects = "SELECT COUNT(PR_Prospects.InternalRecordIdentifier) AS ExpiredActivityCount FROM PR_ProspectActivities "
	SQLProspects = SQLProspects & " INNER JOIN PR_Prospects ON PR_Prospects.InternalRecordIdentifier = PR_ProspectActivities.ProspectRecId "
	SQLProspects = SQLProspects & " WHERE  (Status IS NULL) AND "
	SQLProspects = SQLProspects & " (CAST(ActivityDueDate AS DATE) <= '" & passedEndDate & "') AND "
	SQLProspects = SQLProspects & " (Pool = 'Live') AND "
	SQLProspects = SQLProspects & " OwnerUserNo IN (" & Trim(passedUserNos) & ") "
	
	'Response.Write(SQLProspects & "<br><br>")

	Set rsTotalNumberOfExpiredActivitiesWeeklySnapshot = Server.CreateObject("ADODB.Recordset")
	rsTotalNumberOfExpiredActivitiesWeeklySnapshot.CursorLocation = 3 
	
	Set rsTotalNumberOfExpiredActivitiesWeeklySnapshot = cnnTotalNumberOfExpiredActivitiesWeeklySnapshot.Execute(SQLProspects)
	
	resultTotalNumberOfExpiredActivitiesWeeklySnapshot = rsTotalNumberOfExpiredActivitiesWeeklySnapshot("ExpiredActivityCount") 


	rsTotalNumberOfExpiredActivitiesWeeklySnapshot.Close
	set rsTotalNumberOfExpiredActivitiesWeeklySnapshot= Nothing
	cnnTotalNumberOfExpiredActivitiesWeeklySnapshot.Close	
	set cnnTotalNumberOfExpiredActivitiesWeeklySnapshot= Nothing
	
	TotalNumberOfExpiredActivitiesWeeklySnapshot = resultTotalNumberOfExpiredActivitiesWeeklySnapshot
	
End Function




Function TotalNumberOfCreatedProspectsByLeadSourceWeeklySnapshot(passedStartDate,passedEndDate,passedUserNos,passedLeadSourceIntRecID)

	Set cnnTotalNumberOfCreatedProspectsByLeadSourceWeeklySnapshot = Server.CreateObject("ADODB.Connection")
	cnnTotalNumberOfCreatedProspectsByLeadSourceWeeklySnapshot.open Session("ClientCnnString")

	resultTotalNumberOfCreatedProspectsByLeadSourceWeeklySnapshot = 0
	

	SQLProspects = "SELECT COUNT(PR_Prospects.InternalRecordIdentifier) AS LeadSourceCount "
	SQLProspects = SQLProspects & " FROM  PR_Prospects INNER JOIN"
	SQLProspects = SQLProspects & " PR_LeadSources ON PR_LeadSources.InternalRecordIdentifier = PR_Prospects.LeadSourceNumber "
	SQLProspects = SQLProspects & " WHERE PR_Prospects.OwnerUserNo IN (" & passedUserNos & ") "
	SQLProspects = SQLProspects & " AND CAST(PR_Prospects.CreatedDate AS DATE) >= '" & passedStartDate & "' "
	SQLProspects = SQLProspects & " AND CAST(PR_Prospects.CreatedDate AS DATE) <= '" & passedEndDate & "' "
	SQLProspects = SQLProspects & " AND PR_Prospects.LeadSourceNumber  = " & passedLeadSourceIntRecID

	
	'Response.Write(SQLProspects & "<br><br>")

	Set rsTotalNumberOfCreatedProspectsByLeadSourceWeeklySnapshot = Server.CreateObject("ADODB.Recordset")
	rsTotalNumberOfCreatedProspectsByLeadSourceWeeklySnapshot.CursorLocation = 3 
	
	Set rsTotalNumberOfCreatedProspectsByLeadSourceWeeklySnapshot = cnnTotalNumberOfCreatedProspectsByLeadSourceWeeklySnapshot.Execute(SQLProspects)
			
	resultTotalNumberOfCreatedProspectsByLeadSourceWeeklySnapshot = rsTotalNumberOfCreatedProspectsByLeadSourceWeeklySnapshot("LeadSourceCount")
	
	rsTotalNumberOfCreatedProspectsByLeadSourceWeeklySnapshot.Close
	set rsTotalNumberOfCreatedProspectsByLeadSourceWeeklySnapshot= Nothing
	cnnTotalNumberOfCreatedProspectsByLeadSourceWeeklySnapshot.Close	
	set cnnTotalNumberOfCreatedProspectsByLeadSourceWeeklySnapshot= Nothing
	
	TotalNumberOfCreatedProspectsByLeadSourceWeeklySnapshot = resultTotalNumberOfCreatedProspectsByLeadSourceWeeklySnapshot
	
End Function





Function TotalNumberOfWonProspectsByLeadSourceWeeklySnapshot(passedStartDate,passedEndDate,passedUserNos,passedLeadSourceIntRecID)

	Set cnnTotalNumberOfWonProspectsByLeadSourceWeeklySnapshot = Server.CreateObject("ADODB.Connection")
	cnnTotalNumberOfWonProspectsByLeadSourceWeeklySnapshot.open Session("ClientCnnString")
	
	Set rsTotalNumberOfWonProspectsByLeadSourceWeeklySnapshot = Server.CreateObject("ADODB.Recordset")
	rsTotalNumberOfWonProspectsByLeadSourceWeeklySnapshot.CursorLocation = 3 

	resultTotalNumberOfWonProspectsByLeadSourceWeeklySnapshot = 0
			
	SQLProspectsInner = "SELECT COUNT(*) As ProspCount From ( "
	SQLProspectsInner = SQLProspectsInner & "SELECT Distinct ProspectRecID FROM PR_ProspectStages INNER JOIN PR_Prospects ON"
	SQLProspectsInner = SQLProspectsInner & " PR_Prospects.InternalRecordIdentifier = PR_ProspectStages.ProspectRecID "
	SQLProspectsInner = SQLProspectsInner & " WHERE "
	SQLProspectsInner = SQLProspectsInner & " (Cast(PR_ProspectStages.RecordCreationDateTime As Date) >= '" & passedStartDate & "')"
	SQLProspectsInner = SQLProspectsInner & " AND (Cast(PR_ProspectStages.RecordCreationDateTime As Date)<= '" & passedEndDate& "')"
	SQLProspectsInner = SQLProspectsInner & " AND  StageRecID = 2 "
	SQLProspectsInner = SQLProspectsInner & " AND  PR_Prospects.LeadSourceNumber = " & passedLeadSourceIntRecID & " "
	SQLProspectsInner = SQLProspectsInner & " ) as derivedtbl_1"	

			
	Set rsTotalNumberOfWonProspectsByLeadSourceWeeklySnapshot = cnnTotalNumberOfWonProspectsByLeadSourceWeeklySnapshot.Execute(SQLProspectsInner)
	
	If NOT rsTotalNumberOfWonProspectsByLeadSourceWeeklySnapshot.EOF Then resultTotalNumberOfWonProspectsByLeadSourceWeeklySnapshot = rsTotalNumberOfWonProspectsByLeadSourceWeeklySnapshot("ProspCount") 
			
	set rsTotalNumberOfWonProspectsByLeadSourceWeeklySnapshot= Nothing
	cnnTotalNumberOfWonProspectsByLeadSourceWeeklySnapshot.Close	
	set cnnTotalNumberOfWonProspectsByLeadSourceWeeklySnapshot= Nothing
	
	TotalNumberOfWonProspectsByLeadSourceWeeklySnapshot = resultTotalNumberOfWonProspectsByLeadSourceWeeklySnapshot
	
End Function


Function TotalNumberOfLostProspectsByLeadSourceWeeklySnapshot(passedStartDate,passedEndDate,passedUserNos,passedLeadSourceIntRecID)

	Set cnnTotalNumberOfLostProspectsByLeadSourceWeeklySnapshot = Server.CreateObject("ADODB.Connection")
	cnnTotalNumberOfLostProspectsByLeadSourceWeeklySnapshot.open Session("ClientCnnString")
	
	Set rsTotalNumberOfLostProspectsByLeadSourceWeeklySnapshot = Server.CreateObject("ADODB.Recordset")
	rsTotalNumberOfLostProspectsByLeadSourceWeeklySnapshot.CursorLocation = 3 

	resultTotalNumberOfLostProspectsByLeadSourceWeeklySnapshot = 0
			
	SQLProspectsInner = "SELECT Count(*) As ProspCount FROM PR_ProspectStages INNER JOIN PR_Prospects ON"
	SQLProspectsInner = SQLProspectsInner & " PR_Prospects.InternalRecordIdentifier = PR_ProspectStages.ProspectRecID "
	SQLProspectsInner = SQLProspectsInner & " WHERE "
	SQLProspectsInner = SQLProspectsInner & " (CAST(PR_ProspectStages.RecordCreationDateTime AS DATE) >= '" & passedStartDate & "')"
	SQLProspectsInner = SQLProspectsInner & " AND (CAST(PR_ProspectStages.RecordCreationDateTime AS DATE) <= '" & passedEndDate& "')"
	SQLProspectsInner = SQLProspectsInner & " AND  StageRecID = 1 "
	SQLProspectsInner = SQLProspectsInner & " AND  PR_Prospects.LeadSourceNumber = " & passedLeadSourceIntRecID & " "
			
	Set rsTotalNumberOfLostProspectsByLeadSourceWeeklySnapshot = cnnTotalNumberOfLostProspectsByLeadSourceWeeklySnapshot.Execute(SQLProspectsInner)
	
	If NOT rsTotalNumberOfLostProspectsByLeadSourceWeeklySnapshot.EOF Then resultTotalNumberOfLostProspectsByLeadSourceWeeklySnapshot = rsTotalNumberOfLostProspectsByLeadSourceWeeklySnapshot("ProspCount") 
			
	set rsTotalNumberOfLostProspectsByLeadSourceWeeklySnapshot= Nothing
	cnnTotalNumberOfLostProspectsByLeadSourceWeeklySnapshot.Close	
	set cnnTotalNumberOfLostProspectsByLeadSourceWeeklySnapshot= Nothing
	
	TotalNumberOfLostProspectsByLeadSourceWeeklySnapshot = resultTotalNumberOfLostProspectsByLeadSourceWeeklySnapshot
	
End Function


Function TotalNumberOfUnqualifiedProspectsByLeadSourceWeeklySnapshot(passedStartDate,passedEndDate,passedUserNos,passedLeadSourceIntRecID)

	Set cnnTotalNumberOfUnqualifiedProspectsByLeadSourceWeeklySnapshot = Server.CreateObject("ADODB.Connection")
	cnnTotalNumberOfUnqualifiedProspectsByLeadSourceWeeklySnapshot.open Session("ClientCnnString")
	
	Set rsTotalNumberOfUnqualifiedProspectsByLeadSourceWeeklySnapshot = Server.CreateObject("ADODB.Recordset")
	rsTotalNumberOfUnqualifiedProspectsByLeadSourceWeeklySnapshot.CursorLocation = 3 

	resultTotalNumberOfUnqualifiedProspectsByLeadSourceWeeklySnapshot = 0
			
	SQLProspectsInner = "SELECT Count(*) As ProspCount FROM PR_ProspectStages INNER JOIN PR_Prospects ON"
	SQLProspectsInner = SQLProspectsInner & " PR_Prospects.InternalRecordIdentifier = PR_ProspectStages.ProspectRecID "
	SQLProspectsInner = SQLProspectsInner & " WHERE "
	SQLProspectsInner = SQLProspectsInner & " (CAST(PR_ProspectStages.RecordCreationDateTime AS DATE) >= '" & passedStartDate & "')"
	SQLProspectsInner = SQLProspectsInner & " AND (CAST(PR_ProspectStages.RecordCreationDateTime AS DATE) <= '" & passedEndDate& "')"
	SQLProspectsInner = SQLProspectsInner & " AND  StageRecID = 0 "
	SQLProspectsInner = SQLProspectsInner & " AND  PR_Prospects.LeadSourceNumber = " & passedLeadSourceIntRecID & " "
			
	Set rsTotalNumberOfUnqualifiedProspectsByLeadSourceWeeklySnapshot = cnnTotalNumberOfUnqualifiedProspectsByLeadSourceWeeklySnapshot.Execute(SQLProspectsInner)
	
	If NOT rsTotalNumberOfUnqualifiedProspectsByLeadSourceWeeklySnapshot.EOF Then resultTotalNumberOfUnqualifiedProspectsByLeadSourceWeeklySnapshot = rsTotalNumberOfUnqualifiedProspectsByLeadSourceWeeklySnapshot("ProspCount") 
			
	set rsTotalNumberOfUnqualifiedProspectsByLeadSourceWeeklySnapshot= Nothing
	cnnTotalNumberOfUnqualifiedProspectsByLeadSourceWeeklySnapshot.Close	
	set cnnTotalNumberOfUnqualifiedProspectsByLeadSourceWeeklySnapshot= Nothing
	
	TotalNumberOfUnqualifiedProspectsByLeadSourceWeeklySnapshot = resultTotalNumberOfUnqualifiedProspectsByLeadSourceWeeklySnapshot
	
End Function





Function GetProspectLeadSourceByProspectIntRecID (passedProspectNumber)

	resultGetProspectLeadSourceByProspectIntRecID = ""

	Set cnnGetProspectLeadSourceByProspectIntRecID = Server.CreateObject("ADODB.Connection")
	cnnGetProspectLeadSourceByProspectIntRecID.open Session("ClientCnnString")
		
	SQLGetProspectLeadSourceByProspectIntRecID = "SELECT * FROM PR_Prospects WHERE InternalRecordIdentifier = " & passedProspectNumber
	 
	Set rsGetProspectLeadSourceByProspectIntRecID = Server.CreateObject("ADODB.Recordset")
	rsGetProspectLeadSourceByProspectIntRecID.CursorLocation = 3 
	Set rsGetProspectLeadSourceByProspectIntRecID = cnnGetProspectLeadSourceByProspectIntRecID.Execute(SQLGetProspectLeadSourceByProspectIntRecID)
			
	If not rsGetProspectLeadSourceByProspectIntRecID.EOF Then resultGetProspectLeadSourceByProspectIntRecID = rsGetProspectLeadSourceByProspectIntRecID("LeadSourceNumber")
	
	rsGetProspectLeadSourceByProspectIntRecID.Close
	set rsGetProspectLeadSourceByProspectIntRecID= Nothing
	cnnGetProspectLeadSourceByProspectIntRecID.Close	
	set cnnGetProspectLeadSourceByProspectIntRecID= Nothing
	
	GetProspectLeadSourceByProspectIntRecID = resultGetProspectLeadSourceByProspectIntRecID
	
End Function



Function TotalNumberOfCompletedAppointmentsByLeadSourceWeeklySnapshot(passedStartDate,passedEndDate,passedUserNos,passedLeadSourceIntRecID)

	Set cnnTotalNumberOfCompletedAppointmentsByLeadSourceWeeklySnapshot = Server.CreateObject("ADODB.Connection")
	cnnTotalNumberOfCompletedAppointmentsByLeadSourceWeeklySnapshot.open Session("ClientCnnString")

	resultTotalNumberOfCompletedAppointmentsByLeadSourceWeeklySnapshot = 0
	

	SQLProspects = "SELECT COUNT(PR_ProspectActivities.ActivityCreatedByUserNo) AS AppmtCount"
	SQLProspects = SQLProspects & " FROM  PR_ProspectActivities INNER JOIN "
	SQLProspects = SQLProspects & " PR_Prospects ON PR_ProspectActivities.ProspectRecID = PR_Prospects.InternalRecordIdentifier " 
	SQLProspects = SQLProspects & " WHERE "
	SQLProspects = SQLProspects & " (PR_ProspectActivities.ActivityIsAppointment = 1 OR PR_ProspectActivities.ActivityIsMeeting = 1) "
	SQLProspects = SQLProspects & " AND PR_ProspectActivities.Status='Completed' "
	SQLProspects = SQLProspects & " AND CAST(PR_ProspectActivities.StatusDateTime AS DATE) >= '" & passedStartDate & "' AND CAST(PR_ProspectActivities.StatusDateTime AS DATE) <= '" & passedEndDate & "' "
	SQLProspects = SQLProspects & " AND CAST(PR_ProspectActivities.ActivityDueDate AS DATE) >= '" & passedStartDate & "' AND CAST(PR_ProspectActivities.ActivityDueDate AS DATE) <= '" & passedEndDate & "' "
	SQLProspects = SQLProspects & " AND PR_Prospects.LeadSourceNumber = " & passedLeadSourceIntRecID
	
	SQLProspects = SQLProspects & " AND PR_Prospects.OwnerUserNo IN (" & passedUserNos & ") "

	
	'Response.Write(SQLProspects & "<br><br>")

	Set rsTotalNumberOfCompletedAppointmentsByLeadSourceWeeklySnapshot = Server.CreateObject("ADODB.Recordset")
	rsTotalNumberOfCompletedAppointmentsByLeadSourceWeeklySnapshot.CursorLocation = 3 
	
	Set rsTotalNumberOfCompletedAppointmentsByLeadSourceWeeklySnapshot = cnnTotalNumberOfCompletedAppointmentsByLeadSourceWeeklySnapshot.Execute(SQLProspects)
			
	resultTotalNumberOfCompletedAppointmentsByLeadSourceWeeklySnapshot = rsTotalNumberOfCompletedAppointmentsByLeadSourceWeeklySnapshot("AppmtCount")
	
	rsTotalNumberOfCompletedAppointmentsByLeadSourceWeeklySnapshot.Close
	set rsTotalNumberOfCompletedAppointmentsByLeadSourceWeeklySnapshot= Nothing
	cnnTotalNumberOfCompletedAppointmentsByLeadSourceWeeklySnapshot.Close	
	set cnnTotalNumberOfCompletedAppointmentsByLeadSourceWeeklySnapshot= Nothing
	
	TotalNumberOfCompletedAppointmentsByLeadSourceWeeklySnapshot = resultTotalNumberOfCompletedAppointmentsByLeadSourceWeeklySnapshot
	
End Function



Function TotalNumberOfRescheduledAppointmentsByLeadSourceWeeklySnapshot(passedStartDate,passedEndDate,passedUserNos,passedLeadSourceIntRecID)

	Set cnnTotalNumberOfRescheduledAppointmentsByLeadSourceWeeklySnapshot = Server.CreateObject("ADODB.Connection")
	cnnTotalNumberOfRescheduledAppointmentsByLeadSourceWeeklySnapshot.open Session("ClientCnnString")

	resultTotalNumberOfRescheduledAppointmentsByLeadSourceWeeklySnapshot = 0

	SQLProspects = "SELECT COUNT(PR_ProspectActivities.ActivityCreatedByUserNo) AS AppmtCount"
	SQLProspects = SQLProspects & " FROM  PR_ProspectActivities INNER JOIN "
	SQLProspects = SQLProspects & " PR_Prospects ON PR_ProspectActivities.ProspectRecID = PR_Prospects.InternalRecordIdentifier " 
	SQLProspects = SQLProspects & " WHERE "
	SQLProspects = SQLProspects & " (PR_ProspectActivities.ActivityIsAppointment = 1 OR PR_ProspectActivities.ActivityIsMeeting = 1) "
	SQLProspects = SQLProspects & " AND PR_ProspectActivities.Status='Rescheduled' "
	SQLProspects = SQLProspects & " AND CAST(PR_ProspectActivities.StatusDateTime AS DATE) >= '" & passedStartDate & "' AND CAST(PR_ProspectActivities.StatusDateTime AS DATE) <= '" & passedEndDate & "' "
	SQLProspects = SQLProspects & " AND CAST(PR_ProspectActivities.ActivityDueDate AS DATE) >= '" & passedStartDate & "' AND CAST(PR_ProspectActivities.ActivityDueDate AS DATE) <= '" & passedEndDate & "' "
	SQLProspects = SQLProspects & " AND PR_Prospects.LeadSourceNumber = " & passedLeadSourceIntRecID
	
	SQLProspects = SQLProspects & " AND PR_Prospects.OwnerUserNo IN ( " & passedUserNos & " )"
		
	'Response.Write(SQLProspects & "<br><br>")

	Set rsTotalNumberOfRescheduledAppointmentsByLeadSourceWeeklySnapshot = Server.CreateObject("ADODB.Recordset")
	rsTotalNumberOfRescheduledAppointmentsByLeadSourceWeeklySnapshot.CursorLocation = 3 
	
	Set rsTotalNumberOfRescheduledAppointmentsByLeadSourceWeeklySnapshot = cnnTotalNumberOfRescheduledAppointmentsByLeadSourceWeeklySnapshot.Execute(SQLProspects)
			
	resultTotalNumberOfRescheduledAppointmentsByLeadSourceWeeklySnapshot = rsTotalNumberOfRescheduledAppointmentsByLeadSourceWeeklySnapshot("AppmtCount")
	
	rsTotalNumberOfRescheduledAppointmentsByLeadSourceWeeklySnapshot.Close
	set rsTotalNumberOfRescheduledAppointmentsByLeadSourceWeeklySnapshot= Nothing
	cnnTotalNumberOfRescheduledAppointmentsByLeadSourceWeeklySnapshot.Close	
	set cnnTotalNumberOfRescheduledAppointmentsByLeadSourceWeeklySnapshot= Nothing
	
	TotalNumberOfRescheduledAppointmentsByLeadSourceWeeklySnapshot = resultTotalNumberOfRescheduledAppointmentsByLeadSourceWeeklySnapshot
	
End Function


Function TotalNumberOfCancelledAppointmentsByLeadSourceWeeklySnapshot(passedStartDate,passedEndDate,passedUserNos,passedLeadSourceIntRecID)

	Set cnnTotalNumberOfCancelledAppointmentsByLeadSourceWeeklySnapshot = Server.CreateObject("ADODB.Connection")
	cnnTotalNumberOfCancelledAppointmentsByLeadSourceWeeklySnapshot.open Session("ClientCnnString")

	resultTotalNumberOfCancelledAppointmentsByLeadSourceWeeklySnapshot = 0

	SQLProspects = "SELECT COUNT(PR_ProspectActivities.ActivityCreatedByUserNo) AS AppmtCount"
	SQLProspects = SQLProspects & " FROM  PR_ProspectActivities INNER JOIN "
	SQLProspects = SQLProspects & " PR_Prospects ON PR_ProspectActivities.ProspectRecID = PR_Prospects.InternalRecordIdentifier " 
	SQLProspects = SQLProspects & " WHERE "
	SQLProspects = SQLProspects & " (PR_ProspectActivities.ActivityIsAppointment = 1 OR PR_ProspectActivities.ActivityIsMeeting = 1) "
	SQLProspects = SQLProspects & " AND PR_ProspectActivities.Status='Cancelled' "
	SQLProspects = SQLProspects & " AND CAST(PR_ProspectActivities.StatusDateTime AS DATE) >= '" & passedStartDate & "' AND CAST(PR_ProspectActivities.StatusDateTime AS DATE) <= '" & passedEndDate & "' "
	SQLProspects = SQLProspects & " AND CAST(PR_ProspectActivities.ActivityDueDate AS DATE) >= '" & passedStartDate & "' AND CAST(PR_ProspectActivities.ActivityDueDate AS DATE) <= '" & passedEndDate & "' "
	SQLProspects = SQLProspects & " AND PR_Prospects.LeadSourceNumber = " & passedLeadSourceIntRecID
	
	SQLProspects = SQLProspects & " AND PR_Prospects.OwnerUserNo IN ( " & passedUserNos & " )"


	
	'Response.Write(SQLProspects & "<br><br>")

	Set rsTotalNumberOfCancelledAppointmentsByLeadSourceWeeklySnapshot = Server.CreateObject("ADODB.Recordset")
	rsTotalNumberOfCancelledAppointmentsByLeadSourceWeeklySnapshot.CursorLocation = 3 
	
	Set rsTotalNumberOfCancelledAppointmentsByLeadSourceWeeklySnapshot = cnnTotalNumberOfCancelledAppointmentsByLeadSourceWeeklySnapshot.Execute(SQLProspects)
			
	resultTotalNumberOfCancelledAppointmentsByLeadSourceWeeklySnapshot = rsTotalNumberOfCancelledAppointmentsByLeadSourceWeeklySnapshot("AppmtCount")
	
	rsTotalNumberOfCancelledAppointmentsByLeadSourceWeeklySnapshot.Close
	set rsTotalNumberOfCancelledAppointmentsByLeadSourceWeeklySnapshot= Nothing
	cnnTotalNumberOfCancelledAppointmentsByLeadSourceWeeklySnapshot.Close	
	set cnnTotalNumberOfCancelledAppointmentsByLeadSourceWeeklySnapshot= Nothing
	
	TotalNumberOfCancelledAppointmentsByLeadSourceWeeklySnapshot = resultTotalNumberOfCancelledAppointmentsByLeadSourceWeeklySnapshot
	
End Function


Function TotalNumberOfNotUpdatedAppointmentsByLeadSourceWeeklySnapshot(passedStartDate,passedEndDate,passedUserNos,passedLeadSourceIntRecID)

	Set cnnTotalNumberOfNotUpdatedAppointmentsByLeadSourceWeeklySnapshot = Server.CreateObject("ADODB.Connection")
	cnnTotalNumberOfNotUpdatedAppointmentsByLeadSourceWeeklySnapshot.open Session("ClientCnnString")

	resultTotalNumberOfNotUpdatedAppointmentsByLeadSourceWeeklySnapshot = 0

	SQLProspects = "SELECT COUNT(PR_ProspectActivities.ActivityCreatedByUserNo) AS AppmtCount"
	SQLProspects = SQLProspects & " FROM  PR_ProspectActivities INNER JOIN "
	SQLProspects = SQLProspects & " PR_Prospects ON PR_ProspectActivities.ProspectRecID = PR_Prospects.InternalRecordIdentifier " 
	SQLProspects = SQLProspects & " WHERE "
	SQLProspects = SQLProspects & " (PR_ProspectActivities.ActivityIsAppointment = 1 OR PR_ProspectActivities.ActivityIsMeeting = 1) "
	SQLProspects = SQLProspects & " AND PR_ProspectActivities.Status IS NULL "
	SQLProspects = SQLProspects & " AND CAST(PR_ProspectActivities.StatusDateTime AS DATE) >= '" & passedStartDate & "' AND CAST(PR_ProspectActivities.StatusDateTime AS DATE) <= '" & passedEndDate & "' "
	SQLProspects = SQLProspects & " AND CAST(PR_ProspectActivities.ActivityDueDate AS DATE) >= '" & passedStartDate & "' AND CAST(PR_ProspectActivities.ActivityDueDate AS DATE) <= '" & passedEndDate & "' "
	SQLProspects = SQLProspects & " AND PR_Prospects.LeadSourceNumber = " & passedLeadSourceIntRecID
	
	SQLProspects = SQLProspects & " AND PR_Prospects.OwnerUserNo IN ( " & passedUserNos & " )"  

	
	'Response.Write(SQLProspects & "<br><br>")

	Set rsTotalNumberOfNotUpdatedAppointmentsByLeadSourceWeeklySnapshot = Server.CreateObject("ADODB.Recordset")
	rsTotalNumberOfNotUpdatedAppointmentsByLeadSourceWeeklySnapshot.CursorLocation = 3 
	
	Set rsTotalNumberOfNotUpdatedAppointmentsByLeadSourceWeeklySnapshot = cnnTotalNumberOfNotUpdatedAppointmentsByLeadSourceWeeklySnapshot.Execute(SQLProspects)
			
	resultTotalNumberOfNotUpdatedAppointmentsByLeadSourceWeeklySnapshot = rsTotalNumberOfNotUpdatedAppointmentsByLeadSourceWeeklySnapshot("AppmtCount")
	
	rsTotalNumberOfNotUpdatedAppointmentsByLeadSourceWeeklySnapshot.Close
	set rsTotalNumberOfNotUpdatedAppointmentsByLeadSourceWeeklySnapshot= Nothing
	cnnTotalNumberOfNotUpdatedAppointmentsByLeadSourceWeeklySnapshot.Close	
	set cnnTotalNumberOfNotUpdatedAppointmentsByLeadSourceWeeklySnapshot= Nothing
	
	TotalNumberOfNotUpdatedAppointmentsByLeadSourceWeeklySnapshot = resultTotalNumberOfNotUpdatedAppointmentsByLeadSourceWeeklySnapshot
	
End Function


Function TotalNumberOfCreatedAppointmentsByLeadSourceWeeklySnapshot(passedStartDate,passedEndDate,passedUserNos,passedLeadSourceIntRecID)

	Set cnnTotalNumberOfCreatedAppointmentsByLeadSourceWeeklySnapshot = Server.CreateObject("ADODB.Connection")
	cnnTotalNumberOfCreatedAppointmentsByLeadSourceWeeklySnapshot.open Session("ClientCnnString")

	resultTotalNumberOfCreatedAppointmentsByLeadSourceWeeklySnapshot = 0
	

	SQLProspects = "SELECT COUNT(PR_ProspectActivities.ActivityCreatedByUserNo) AS AppmtCount"
	SQLProspects = SQLProspects & " FROM  PR_ProspectActivities INNER JOIN "
	SQLProspects = SQLProspects & " PR_Prospects ON PR_ProspectActivities.ProspectRecID = PR_Prospects.InternalRecordIdentifier " 
	SQLProspects = SQLProspects & " WHERE "
	SQLProspects = SQLProspects & " (PR_ProspectActivities.ActivityIsAppointment = 1 OR PR_ProspectActivities.ActivityIsMeeting = 1) "
	SQLProspects = SQLProspects & " AND CAST(PR_ProspectActivities.RecordCreationDateTime AS DATE) >= '" & passedStartDate & "' "
	SQLProspects = SQLProspects & " AND CAST(PR_ProspectActivities.RecordCreationDateTime AS DATE) <='" & passedEndDate & "' "
	SQLProspects = SQLProspects & " AND PR_Prospects.LeadSourceNumber = " & passedLeadSourceIntRecID
	SQLProspects = SQLProspects & " AND PR_Prospects.OwnerUserNo IN ( " &  passedUserNos & " )"
				

			
	'Response.Write(SQLProspects & "<br><br>")

	Set rsTotalNumberOfCreatedAppointmentsByLeadSourceWeeklySnapshot = Server.CreateObject("ADODB.Recordset")
	rsTotalNumberOfCreatedAppointmentsByLeadSourceWeeklySnapshot.CursorLocation = 3 
	
	Set rsTotalNumberOfCreatedAppointmentsByLeadSourceWeeklySnapshot = cnnTotalNumberOfCreatedAppointmentsByLeadSourceWeeklySnapshot.Execute(SQLProspects)
			
	resultTotalNumberOfCreatedAppointmentsByLeadSourceWeeklySnapshot = rsTotalNumberOfCreatedAppointmentsByLeadSourceWeeklySnapshot("AppmtCount")
	
	rsTotalNumberOfCreatedAppointmentsByLeadSourceWeeklySnapshot.Close
	set rsTotalNumberOfCreatedAppointmentsByLeadSourceWeeklySnapshot= Nothing
	cnnTotalNumberOfCreatedAppointmentsByLeadSourceWeeklySnapshot.Close	
	set cnnTotalNumberOfCreatedAppointmentsByLeadSourceWeeklySnapshot= Nothing
	
	TotalNumberOfCreatedAppointmentsByLeadSourceWeeklySnapshot = resultTotalNumberOfCreatedAppointmentsByLeadSourceWeeklySnapshot
	
End Function




Function TotalNumberOfExpiredActivitiesByLeadSourceWeeklySnapshot(passedStartDate,passedEndDate,passedUserNos,passedLeadSourceIntRecID)

	Set cnnTotalNumberOfExpiredActivitiesByLeadSourceWeeklySnapshot = Server.CreateObject("ADODB.Connection")
	cnnTotalNumberOfExpiredActivitiesByLeadSourceWeeklySnapshot.open Session("ClientCnnString")

	resultTotalNumberOfExpiredActivitiesByLeadSourceWeeklySnapshot = 0
		
	SQLProspects = "SELECT COUNT(PR_Prospects.InternalRecordIdentifier) AS ExpiredActivityCount FROM PR_ProspectActivities "
	SQLProspects = SQLProspects & " INNER JOIN PR_Prospects ON PR_Prospects.InternalRecordIdentifier = PR_ProspectActivities.ProspectRecId "
	SQLProspects = SQLProspects & " WHERE  (Status IS NULL) AND "
	SQLProspects = SQLProspects & " (CAST(ActivityDueDate AS DATE) <= '" & passedEndDate & "') AND "
	SQLProspects = SQLProspects & " (Pool = 'Live') AND "
	SQLProspects = SQLProspects & " OwnerUserNo IN (" & passedUserNos & ") "
	SQLProspects = SQLProspects & " AND LeadSourceNumber = " & passedLeadSourceIntRecID
	
	'Response.Write(SQLProspects & "<br><br>")

	Set rsTotalNumberOfExpiredActivitiesByLeadSourceWeeklySnapshot = Server.CreateObject("ADODB.Recordset")
	rsTotalNumberOfExpiredActivitiesByLeadSourceWeeklySnapshot.CursorLocation = 3 
	
	Set rsTotalNumberOfExpiredActivitiesByLeadSourceWeeklySnapshot = cnnTotalNumberOfExpiredActivitiesByLeadSourceWeeklySnapshot.Execute(SQLProspects)
			
	resultTotalNumberOfExpiredActivitiesByLeadSourceWeeklySnapshot = rsTotalNumberOfExpiredActivitiesByLeadSourceWeeklySnapshot("ExpiredActivityCount")
	
	rsTotalNumberOfExpiredActivitiesByLeadSourceWeeklySnapshot.Close
	set rsTotalNumberOfExpiredActivitiesByLeadSourceWeeklySnapshot= Nothing
	cnnTotalNumberOfExpiredActivitiesByLeadSourceWeeklySnapshot.Close	
	set cnnTotalNumberOfExpiredActivitiesByLeadSourceWeeklySnapshot= Nothing
	
	TotalNumberOfExpiredActivitiesByLeadSourceWeeklySnapshot = resultTotalNumberOfExpiredActivitiesByLeadSourceWeeklySnapshot
	
End Function


%><!--#include file="InsightFuncs_Prospecting_Outlook.asp"-->
