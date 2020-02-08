<%
'********************************
'List of all the functions & subs
'********************************
'Func AlertsDuringBizHoursOnly ()
'Func ServiceTicketIsDispatched(ServiceTicketNumber)
'Func ServiceTicketDispatchACKed(ServiceTicketNumber)
'Func ServiceTicketWasOnSite(ServiceTicketNumber)
'Func CalcServiceTicketElapsedMinutes(ServiceTicketNumber)
'Func GetServiceTicketProblemByTicketNumber(ServiceTicketNumber)
'Func GetServiceTicketProblemCodeDescByIntRecID(passedProblemCodeIntRecID)
'Func GetServiceTicketProblemLocationByTicketNumber(ServiceTicketNumber)
'Func GetServiceTicketSubmissionDateTimeByTicketNumber(ServiceTicketNumber)
'Func GetServiceTicketSubmissionSourceByTicketNumber(ServiceTicketNumber)
'Func GetServiceTicketTechNotesByTicketNumber(ServiceTicketNumber)
'Func AwaitingRedispatchSince_DateTime(passedTicketNumber)
'Func TimeFromStageStartUntilTargetDateTime(passedMemoNumber,passedStage,PassedTargetDateTime)
'Func FS_SignatureOptional()
'Func FS_TechCanDecline()
'Func FSDefaultNotificationMethod()
'Func NumberOfUrgentServiceTicketsByTech(passedServiceTechNum)
'Func MostRecentDispatchDeclineByTicket(passedTicketNumber)
'Func MostRecentDispatchDeclineRecordNumberByTicket(passedTicketNumber)
'Func GetNumberOfServiceTicketsInTimeRange(passedStartMins, passedEndMins) 
'Func GetNumberOfServiceTicketsInTimeRange2(RangeNum) 
'Func GetNumberOfFilterTicketsInTimeRange(passedStartMins, passedEndMins)
'Func GetNumberOfFilterChangesInTimeRange(passedStartMins, passedEndMins)

'Func GetNumberOfServiceCallsAwaitingDispatch()
'Func GetNumberOfServiceCallsAwaitingAcknowledgement()
'Func GetNumberOfServiceCallsAcknowledged()
'Func GetNumberOfServiceCallsEnRouteOnSite()
'Func GetNumberOfServiceCallsRedo()
'Func GetNumberOfServiceCallsRedoFiltersOnly()

'Func GetNumberOfServiceCallsAwaitingDispatchWithFilters()
'Func GetNumberOfServiceCallsAwaitingAcknowledgementWithFilters()
'Func GetNumberOfServiceCallsAcknowledgedWithFilters()
'Func GetNumberOfServiceCallsEnRouteOnSiteWithFilters()
'Func GetNumberOfServiceCallsRedoWithFilters()

'Func GetNumberOfFilterChangesForServiceTicket(passedServiceTicketNumber)
'Func GetNumberOfClosedServiceTicketsForTech(passedCloseDate,passedTechUserNo)
'Func GetNumberOfClosedFilterTicketsForTech(passedCloseDate,passedTechUserNo)
'Func GetNumberOfClosedFilterChangesForTech(passedCloseDate,passedTechUserNo)

'Func GetNumberOfClosedServiceTicketsForCustomerAcct(passedCloseDate,passedCustID)
'Func GetNumberOfClosedFilterTicketsForCustomerAcct(passedCloseDate,passedCustID)
'Func GetNumberOfClosedFilterChangesForCustomerAcct(passedCloseDate,passedCustID)

'Func GetNumberOfServiceCallsClosedThisWeek()
'func GetNumberOfServiceCallsClosedRolling5Days()
'Func GetNumberOfServiceCallsFilterChanges()
'Func ServiceCallElapsedMinutesOpenTicket(passedServiceTicketNumber)
'Func ServiceCallElapsedMinutesClosedTicket(passedServiceTicketNumber)
'Func CustHasServiceTicketNotes(passedServiceTicketNumber)
'Func NoteNewServiceTicketForUser(passedServiceTicketNumber)
'Sub  MarkNoteNewForUserServiceTicket(passedServiceTicketNumber)
'Func GetLastServiceTicketNotesByTicket(passedServiceTicketNumber)
'Func GetNumberOfMinutesInServiceDay()
'Func CustHasPendingFilterChange_NextDate (passedCust)
'Func NumberOfTicketsByProblemCode(passedProblemCode)
'Func NumberOfTicketsBySymptomCode(passedSymptomCode)
'Func NumberOfTicketsByResolutionCode(passedResolutionCode)
'Func NumberCustomerRecsDefinedForFilterID(passedFilterIntRecID)
'Func GetFilterDescByFilterIntRecID(passedrFilterIntRecID)
'Func GetFilterIDByIntRecID(passedFilterIntRecID)
'Func GetFilterDescByFilterIntRecID(passedFilterIntRecID)
'Func GetFilterLocationByFilterIntRecID(passedFilterIntRecID)
'Func FilterChangeSubmittedNewLogic(passedCustomerFilterIntRecID)
'Func TicketInServiceMemosFilterInfo(passedServiceTicketID)
'Func GetOpenFilterTicketsByCustID(passedCustID)
'Func CustHasPendingFilterChange(passedCustid)
'Func GetNumberOfServiceCallsOnHold()
'Func GetHOLDServiceTicketSTAGEDateTime(passedTicketNumber,passedStage)
'Func userCreateEquipmentSymptomCodesOnTheFly(passedUserNo)
'Func userCreateEquipmentResolutionCodesOnTheFly(passedUserNo)
'func userCreateEquipmentProblemCodesOnTheFly(passedUserNo)
'************************************
'End List of all the functions & subs
'************************************


Function AlertsDuringBizHoursOnly ()

	resultAlertsDuringBizHoursOnly = vbTrue 'Default to true

	Set cnn_CheckAlerts = Server.CreateObject("ADODB.Connection")
	cnn_CheckAlerts.open (Session("ClientCnnString"))
	Set rsAlertsDuringBizHoursOnly = Server.CreateObject("ADODB.Recordset")
	rsAlertsDuringBizHoursOnly.CursorLocation = 3 
	
	SQL_CheckAlerts = "SELECT * FROM Settings_EmailService "
	Set rsAlertsDuringBizHoursOnly = cnn_CheckAlerts.Execute(SQL_CheckAlerts)

	If not rsAlertsDuringBizHoursOnly.EOF Then resultAlertsDuringBizHoursOnly = rsAlertsDuringBizHoursOnly("AlertsDuringBizHoursOnly")

	Set rsAlertsDuringBizHoursOnly = Nothing
	cnn_CheckAlerts.Close
	Set cnn_CheckAlerts = Nothing

	AlertsDuringBizHoursOnly =	resultAlertsDuringBizHoursOnly 
	
End Function

Function ServiceTicketIsDispatched(passedServiceTicketNumber)

	resultServiceTicketIsDispatched = False
	
	Set cnnServiceTicketIsDispatched = Server.CreateObject("ADODB.Connection")
	cnnServiceTicketIsDispatched.open Session("ClientCnnString")
	
	SQLServiceTicketIsDispatched = "SELECT * FROM  FS_ServiceMemosDetail WHERE MemoNumber = " & passedServiceTicketNumber & " AND MemoStage = 'Dispatched'"

	Set rsServiceTicketIsDispatched = Server.CreateObject("ADODB.Recordset")
	rsServiceTicketIsDispatched.CursorLocation = 3 
	Set rsServiceTicketIsDispatched= cnnServiceTicketIsDispatched.Execute(SQLServiceTicketIsDispatched)
	
	If not rsServiceTicketIsDispatched.eof then resultServiceTicketIsDispatched = True
	
	set rsServiceTicketIsDispatched= Nothing
	set cnnServiceTicketIsDispatched= Nothing
	
	ServiceTicketIsDispatched= resultServiceTicketIsDispatched

End Function

Function ServiceTicketDispatchACKed(passedServiceTicketNumber)

	resultServiceTicketDispatchACKed = False
	
	Set cnnServiceTicketDispatchACKed = Server.CreateObject("ADODB.Connection")
	cnnServiceTicketDispatchACKed.open Session("ClientCnnString")
	
	SQLServiceTicketDispatchACKed = "SELECT * FROM  FS_ServiceMemosDetail WHERE MemoNumber = " & passedServiceTicketNumber & " AND MemoStage = 'Dispatch Acknowledged'"

	Set rsServiceTicketDispatchACKed = Server.CreateObject("ADODB.Recordset")
	rsServiceTicketDispatchACKed.CursorLocation = 3 
	Set rsServiceTicketDispatchACKed= cnnServiceTicketDispatchACKed.Execute(SQLServiceTicketDispatchACKed)
	
	If not rsServiceTicketDispatchACKed.eof then resultServiceTicketDispatchACKed = True
	
	set rsServiceTicketDispatchACKed= Nothing
	set cnnServiceTicketDispatchACKed= Nothing
	
	ServiceTicketDispatchACKed= resultServiceTicketDispatchACKed

End Function

Function ServiceTicketWasOnSite(passedServiceTicketNumber)

	resultServiceTicketWasOnSite = False
	
	Set cnnServiceTicketWasOnSite = Server.CreateObject("ADODB.Connection")
	cnnServiceTicketWasOnSite.open Session("ClientCnnString")
	
	SQLServiceTicketWasOnSite = "SELECT * FROM  FS_ServiceMemosDetail WHERE MemoNumber = " & passedServiceTicketNumber & " AND MemoStage = 'On Site'"


	Set rsServiceTicketWasOnSite = Server.CreateObject("ADODB.Recordset")
	rsServiceTicketWasOnSite.CursorLocation = 3 
	Set rsServiceTicketWasOnSite= cnnServiceTicketWasOnSite.Execute(SQLServiceTicketWasOnSite)
	
	If not rsServiceTicketWasOnSite.eof then resultServiceTicketWasOnSite = True
	
	set rsServiceTicketWasOnSite= Nothing
	set cnnServiceTicketWasOnSite= Nothing
	
	ServiceTicketWasOnSite= resultServiceTicketWasOnSite

End Function



Function CalcServiceTicketElapsedMinutes(passedServiceTicketNumber)

	resultCalcServiceTicketElapsedMinutes = 0

	Set cnnCalcServiceTicketElapsedMinutes = Server.CreateObject("ADODB.Connection")
	cnnCalcServiceTicketElapsedMinutes.open Session("ClientCnnString")

	SQLCalcServiceTicketElapsedMinutes = "SELECT * FROM  FS_ServiceMemos WHERE MemoNumber = '" & passedServiceTicketNumber & "'"
	
	Set rsCalcServiceTicketElapsedMinutes = Server.CreateObject("ADODB.Recordset")
	rsCalcServiceTicketElapsedMinutes.CursorLocation = 3 
	Set rsCalcServiceTicketElapsedMinutes = cnnCalcServiceTicketElapsedMinutes.Execute(SQLCalcServiceTicketElapsedMinutes)

	If ElapsedTimeCalcMethod() = "Actual" Then
		If rsCalcServiceTicketElapsedMinutes("CurrentStatus") = "CLOSE" or rsCalcServiceTicketElapsedMinutes("CurrentStatus") = "CANCEL" Then
			resultCalcServiceTicketElapsedMinutes = datediff("n",GetServiceTicketOpenDateTime(passedServiceTicketNumber),GetServiceTicketCloseDateTime(passedServiceTicketNumber))
		ElseIf rsCalcServiceTicketElapsedMinutes("CurrentStatus") = "OPEN" Then 
			resultCalcServiceTicketElapsedMinutes = datediff("n",GetServiceTicketOpenDateTime(passedServiceTicketNumber),Now())
		Elseif rsCalcServiceTicketElapsedMinutes("CurrentStatus") = "HOLD" Then
			resultCalcServiceTicketElapsedMinutes = datediff("n",GetServiceTicketOpenDateTime(rs.Fields("MemoNumber")),Now())
		End If
	Else
		If rsCalcServiceTicketElapsedMinutes("CurrentStatus") = "CLOSE" or rsCalcServiceTicketElapsedMinutes("CurrentStatus") = "CANCEL" Then
			resultCalcServiceTicketElapsedMinutes = ServiceCallElapsedMinutesClosedTicket(passedServiceTicketNumber)
		ElseIf rsCalcServiceTicketElapsedMinutes("CurrentStatus") = "OPEN" Then 
			resultCalcServiceTicketElapsedMinutes = ServiceCallElapsedMinutesOpenTicket(passedServiceTicketNumber)
		Elseif rsCalcServiceTicketElapsedMinutes("CurrentStatus") = "HOLD" Then
			resultCalcServiceTicketElapsedMinutes = 0
		End If
	End If

	CalcServiceTicketElapsedMinutes = resultCalcServiceTicketElapsedMinutes 
	
End Function



Function GetServiceTicketProblemByTicketNumber(passedTicketNumber)

	resultGetServiceTicketProblemByTicketNumber = ""
	
	Set cnnGetServiceTicketProblemByTicketNumber = Server.CreateObject("ADODB.Connection")
	cnnGetServiceTicketProblemByTicketNumber.open Session("ClientCnnString")

	SQLGetServiceTicketProblemByTicketNumber = "SELECT * FROM FS_ServiceMemos WHERE MemoNumber = '" & passedTicketNumber & "' AND RecordSubType = 'OPEN'"

	Set rsGetServiceTicketProblemByTicketNumber = Server.CreateObject("ADODB.Recordset")
	rsGetServiceTicketProblemByTicketNumber.CursorLocation = 3 
	Set rsGetServiceTicketProblemByTicketNumber = cnnGetServiceTicketProblemByTicketNumber.Execute(SQLGetServiceTicketProblemByTicketNumber)
	
	If not rsGetServiceTicketProblemByTicketNumber.eof then
		resultGetServiceTicketProblemByTicketNumber =  rsGetServiceTicketProblemByTicketNumber("ProblemDescription")
	End If

	set rsGetServiceTicketProblemByTicketNumber = Nothing
	set cnnGetServiceTicketProblemByTicketNumber= Nothing
	
	GetServiceTicketProblemByTicketNumber = resultGetServiceTicketProblemByTicketNumber

End Function


Function GetServiceTicketProblemCodeDescByIntRecID(passedProblemCodeIntRecID)

	resultGetServiceTicketProblemCodeDescByIntRecID = ""
	
	Set cnnGetServiceTicketProblemCodeDescByIntRecID = Server.CreateObject("ADODB.Connection")
	cnnGetServiceTicketProblemCodeDescByIntRecID.open Session("ClientCnnString")

	SQLGetServiceTicketProblemCodeDescByIntRecID = "SELECT * FROM FS_ProblemCodes WHERE InternalRecordIdentifier = '" & passedProblemCodeIntRecID & "'"

	Set rsGetServiceTicketProblemCodeDescByIntRecID = Server.CreateObject("ADODB.Recordset")
	rsGetServiceTicketProblemCodeDescByIntRecID.CursorLocation = 3 
	Set rsGetServiceTicketProblemCodeDescByIntRecID = cnnGetServiceTicketProblemCodeDescByIntRecID.Execute(SQLGetServiceTicketProblemCodeDescByIntRecID)
	
	If not rsGetServiceTicketProblemCodeDescByIntRecID.eof then
		resultGetServiceTicketProblemCodeDescByIntRecID =  rsGetServiceTicketProblemCodeDescByIntRecID("ProblemDescription")
	End If

	set rsGetServiceTicketProblemCodeDescByIntRecID = Nothing
	set cnnGetServiceTicketProblemCodeDescByIntRecID= Nothing
	
	GetServiceTicketProblemCodeDescByIntRecID = resultGetServiceTicketProblemCodeDescByIntRecID

End Function



Function GetServiceTicketProblemLocationByTicketNumber(passedTicketNumber)

	resultGetServiceTicketProblemLocationByTicketNumber = ""
	
	Set cnnGetServiceTicketProblemLocationByTicketNumber = Server.CreateObject("ADODB.Connection")
	cnnGetServiceTicketProblemLocationByTicketNumber.open Session("ClientCnnString")

	SQLGetServiceTicketProblemLocationByTicketNumber = "SELECT * FROM FS_ServiceMemos WHERE MemoNumber = '" & passedTicketNumber & "'"

	Set rsGetServiceTicketProblemLocationByTicketNumber = Server.CreateObject("ADODB.Recordset")
	rsGetServiceTicketProblemLocationByTicketNumber.CursorLocation = 3 
	Set rsGetServiceTicketProblemLocationByTicketNumber = cnnGetServiceTicketProblemLocationByTicketNumber.Execute(SQLGetServiceTicketProblemLocationByTicketNumber)
	
	If not rsGetServiceTicketProblemLocationByTicketNumber.eof then
		resultGetServiceTicketProblemLocationByTicketNumber =  rsGetServiceTicketProblemLocationByTicketNumber("ProblemLocation")
	End If

	set rsGetServiceTicketProblemLocationByTicketNumber = Nothing
	set cnnGetServiceTicketProblemLocationByTicketNumber= Nothing
	
	GetServiceTicketProblemLocationByTicketNumber = resultGetServiceTicketProblemLocationByTicketNumber

End Function

Function GetServiceTicketSubmissionDateTimeByTicketNumber(passedTicketNumber)

	resultGetServiceTicketSubmissionDateTimeByTicketNumber = ""
	
	Set cnnGetServiceTicketSubmissionDateTimeByTicketNumber = Server.CreateObject("ADODB.Connection")
	cnnGetServiceTicketSubmissionDateTimeByTicketNumber.open Session("ClientCnnString")

	SQLGetServiceTicketSubmissionDateTimeByTicketNumber = "SELECT * FROM FS_ServiceMemos WHERE MemoNumber = '" & passedTicketNumber & "'"

	Set rsGetServiceTicketSubmissionDateTimeByTicketNumber = Server.CreateObject("ADODB.Recordset")
	rsGetServiceTicketSubmissionDateTimeByTicketNumber.CursorLocation = 3 
	Set rsGetServiceTicketSubmissionDateTimeByTicketNumber = cnnGetServiceTicketSubmissionDateTimeByTicketNumber.Execute(SQLGetServiceTicketSubmissionDateTimeByTicketNumber)
	
	If not rsGetServiceTicketSubmissionDateTimeByTicketNumber.eof then
		resultGetServiceTicketSubmissionDateTimeByTicketNumber =  rsGetServiceTicketSubmissionDateTimeByTicketNumber("SubmissionDateTime")
	End If

	set rsGetServiceTicketSubmissionDateTimeByTicketNumber = Nothing
	set cnnGetServiceTicketSubmissionDateTimeByTicketNumber= Nothing
	
	GetServiceTicketSubmissionDateTimeByTicketNumber = resultGetServiceTicketSubmissionDateTimeByTicketNumber

End Function

Function GetServiceTicketSubmissionSourceByTicketNumber(passedTicketNumber)

	resultGetServiceTicketSubmissionSourceByTicketNumber = ""
	
	Set cnnGetServiceTicketSubmissionSourceByTicketNumber = Server.CreateObject("ADODB.Connection")
	cnnGetServiceTicketSubmissionSourceByTicketNumber.open Session("ClientCnnString")

	SQLGetServiceTicketSubmissionSourceByTicketNumber = "SELECT * FROM FS_ServiceMemos WHERE MemoNumber = '" & passedTicketNumber & "'"

	Set rsGetServiceTicketSubmissionSourceByTicketNumber = Server.CreateObject("ADODB.Recordset")
	rsGetServiceTicketSubmissionSourceByTicketNumber.CursorLocation = 3 
	Set rsGetServiceTicketSubmissionSourceByTicketNumber = cnnGetServiceTicketSubmissionSourceByTicketNumber.Execute(SQLGetServiceTicketSubmissionSourceByTicketNumber)
	
	If not rsGetServiceTicketSubmissionSourceByTicketNumber.eof then
		resultGetServiceTicketSubmissionSourceByTicketNumber=  rsGetServiceTicketSubmissionSourceByTicketNumber("SubmissionSource")
	End If

	set rsGetServiceTicketSubmissionSourceByTicketNumber = Nothing
	set cnnGetServiceTicketSubmissionSourceByTicketNumber= Nothing
	
	GetServiceTicketSubmissionSourceByTicketNumber = resultGetServiceTicketSubmissionSourceByTicketNumber

End Function



Function AwaitingRedispatchSince_DateTime(passedTicketNumber)

	resultAwaitingRedispatchSince_DateTime = ""
	
	Set cnnAwaitingRedispatchSince_DateTime = Server.CreateObject("ADODB.Connection")
	cnnAwaitingRedispatchSince_DateTime.open Session("ClientCnnString")

	SQLAwaitingRedispatchSince_DateTime = "Select * from FS_ServiceMemosRedispatch Where MemoNumber ='" & passedTicketNumber & "'"
	
	Set rsAwaitingRedispatchSince_DateTime = Server.CreateObject("ADODB.Recordset")
	rsAwaitingRedispatchSince_DateTime.CursorLocation = 3 
	Set rsAwaitingRedispatchSince_DateTime = cnnAwaitingRedispatchSince_DateTime.Execute(SQLAwaitingRedispatchSince_DateTime)

	If not rsAwaitingRedispatchSince_DateTime.Eof Then resultAwaitingRedispatchSince_DateTime = rsAwaitingRedispatchSince_DateTime("RecordCreationDateTime") 
		
	set rsAwaitingRedispatchSince_DateTime = Nothing
	set cnnAwaitingRedispatchSince_DateTime= Nothing
	
	AwaitingRedispatchSince_DateTime = resultAwaitingRedispatchSince_DateTime

End Function

Function TimeFromStageStartUntilTargetDateTime(passedMemoNumber,passedStage,PassedTargetDateTime)

	resultTimeFromStageStartUntilTargetDateTime = 0

	Set cnnTimeFromStageStartUntilTargetDateTime = Server.CreateObject("ADODB.Connection")
	cnnTimeFromStageStartUntilTargetDateTime.open Session("ClientCnnString")

	SQLTimeFromStageStartUntilTargetDateTime = "SELECT TOP 1 * FROM  FS_ServiceMemosDetail WHERE MemoNumber = '" & passedMemoNumber & "' AND MemoStage='" & passedStage & "' ORDER BY RecordCreatedDateTime DESC"
	
	Set rsTimeFromStageStartUntilTargetDateTime = Server.CreateObject("ADODB.Recordset")
	rsTimeFromStageStartUntilTargetDateTime.CursorLocation = 3 
	Set rsTimeFromStageStartUntilTargetDateTime = cnnTimeFromStageStartUntilTargetDateTime.Execute(SQLTimeFromStageStartUntilTargetDateTime)

	If Not rsTimeFromStageStartUntilTargetDateTime.EOF Then
		StageStarDateTime = rsTimeFromStageStartUntilTargetDateTime("RecordCreatedDateTime")
	Else
		'Received but no stage, so get the date time from the OPEN record
		SQLTimeFromStageStartUntilTargetDateTime = "SELECT * FROM  FS_ServiceMemos WHERE MemoNumber = '" & passedMemoNumber & "' AND RecordSubType='OPEN'"
		Set rsTimeFromStageStartUntilTargetDateTime = cnnTimeFromStageStartUntilTargetDateTime.Execute(SQLTimeFromStageStartUntilTargetDateTime)
		If Not rsTimeFromStageStartUntilTargetDateTime.EOF Then
			StageStarDateTime = rsTimeFromStageStartUntilTargetDateTime("RecordCreatedateTime")
		Else
			'I don't know what to do here
		End IF
	End If
	'Redispatch is different, so check there too
	If passedStage = "Redispatch" Then StageStarDateTime = AwaitingRedispatchSince_DateTime(passedMemoNumber)


	'This should only occur on open tickets
	If ElapsedTimeCalcMethod() = "Actual" Then
		resultTimeFromStageStartUntilTargetDateTime = datediff("n",StageStarDateTime,PassedTargetDateTime)
	Else
		debugmsg=0
		totalElapsedMinutes = 0
	
		'Get the normal business day start & end because it is used
		'in a lot o fplaces
		Set cnn10 = Server.CreateObject("ADODB.Connection")
		cnn10.open (Session("ClientCnnString"))
		Set rs10 = Server.CreateObject("ADODB.Recordset")
		rs10.CursorLocation = 3 
		SQL10 = "SELECT * FROM Settings_FieldService"
		Set rs10 = cnn10.Execute(SQL10 )
		NormalBizDayStartTime = rs10.fields("ServiceDayStartTime")
		NormalBizDayEndTime = rs10.fields("ServiceDayEndTime")
		MinutesInFullDay = GetNumberOfMinutesInServiceDay()
		Set rs10 = Nothing
		cnn10.close	
		Set cnn10=Nothing
				
		If datediff("d",StageStarDateTime,PassedTargetDateTime) < 1 Then
			'It was only opened today so get the number of minutes from the time
			'it was open to Now() or until th end of the day if Now() is after hours

			OpenedDate = FormatDateTime(StageStarDateTime,2)
			OpenedTime = FormatDateTime(StageStarDateTime,4)
			If NumberofWorkMinutes_DateOpened(OpenedDate,OpenedTime,NormalBizDayStartTime,NormalBizDayEndTime) <> 0 Then 'They were not closed
				If FormatDateTime(PassedTargetDateTime,4) > NormalBizDayEndTime Then ' It was closed after hours so just give them the minutes since opening that day
					totalElapsedMinutes =  NumberofWorkMinutes_DateOpened(OpenedDate,OpenedTime,NormalBizDayStartTime,NormalBizDayEndTime)
				Else
					'Otherwise, it is just the actual minutes
					totalElapsedMinutes =  datediff("n",StageStarDateTime,PassedTargetDateTime)
				End If
			Else
				'They were closed
				totalElapsedMinutes = 0
			End If
		Else
			'Get the number of minutes remaining on the day it was opened
			OpenedDate = FormatDateTime(StageStarDateTime,2)
			OpenedTime = FormatDateTime(StageStarDateTime,4)
							
			totalElapsedMinutes =  NumberofWorkMinutes_DateOpened(OpenedDate,OpenedTime,NormalBizDayStartTime,NormalBizDayEndTime)
				
			'Now see if there are more days to add on
			EndDate = FormatDateTime(PassedTargetDateTime,2)
							
			WorkingDate=DateAdd("d",1,StageStarDateTime)
			WorkingDate = FormatDateTime(WorkingDate,2)
			EndDate = FormatDateTime(EndDate,2)
							
			WorkingDate = cdate(WorkingDate)
			EndDate = cdate(EndDate)
			
						
			Do While WorkingDate < EndDate
				totalElapsedMinutes  = totalElapsedMinutes + NumberofWorkMinutes_FullDay(WorkingDate,NormalBizDayStartTime,NormalBizDayEndTime)
				WorkingDate=DateAdd("d",1,WorkingDate)
			Loop
				
			' Now add in the minutes for the day that it was closed
		
			LastDayDate = FormatDateTime(PassedTargetDateTime,2)
			LastDayTime = FormatDateTime(PassedTargetDateTime,4)

			totalElapsedMinutes = totalElapsedMinutes  + NumberofWorkMinutes_DateClosed(LastDayDate,LastDayTime,NormalBizDayStartTime,NormalBizDayEndTime ) 
		End If
		resultTimeFromStageStartUntilTargetDateTime = totalElapsedMinutes
	End If	
	
	TimeFromStageStartUntilTargetDateTime = resultTimeFromStageStartUntilTargetDateTime 
	
End Function

Function FS_SignatureOptional()

	resultFS_SignatureOptional = False
	
	Set cnnFS_SignatureOptional = Server.CreateObject("ADODB.Connection")
	cnnFS_SignatureOptional.open Session("ClientCnnString")

	SQLFS_SignatureOptional = "SELECT FS_SignatureOptional FROM Settings_Global"

	Set rsFS_SignatureOptional = Server.CreateObject("ADODB.Recordset")
	rsFS_SignatureOptional.CursorLocation = 3 
	Set rsFS_SignatureOptional = cnnFS_SignatureOptional.Execute(SQLFS_SignatureOptional)
	
	If not rsFS_SignatureOptional.eof then 
		If rsFS_SignatureOptional("FS_SignatureOptional") = 1 Then resultFS_SignatureOptional = True
	End IF	
	
	set rsFS_SignatureOptional = Nothing
	cnnFS_SignatureOptional.Close
	set cnnFS_SignatureOptional = Nothing
	
	FS_SignatureOptional = resultFS_SignatureOptional


End Function

Function NumberOfUrgentServiceTicketsByTech(passedServiceTechNum)

	resultNumberOfUrgentServiceTicketsByTech = 0 

	SQLNumberOfUrgentServiceTicketsByTech = "SELECT Distinct Memonumber From FS_ServiceMemosDetail Where UserNoOfServiceTech ='" & passedServiceTechNum & "' AND ClosedorCancelled <> 1"
	'SQLNumberOfUrgentServiceTicketsByTech = SQLNumberOfUrgentServiceTicketsByTech & " AND (MemoStage  = 'Dispatched'"
	'SQLNumberOfUrgentServiceTicketsByTech = SQLNumberOfUrgentServiceTicketsByTech & " OR MemoStage = 'Dispatch Acknowledged'"
	'SQLNumberOfUrgentServiceTicketsByTech = SQLNumberOfUrgentServiceTicketsByTech & " OR MemoStage = 'En Route'"
	'SQLNumberOfUrgentServiceTicketsByTech = SQLNumberOfUrgentServiceTicketsByTech & " OR MemoStage = 'On Site')"
	SQLNumberOfUrgentServiceTicketsByTech = SQLNumberOfUrgentServiceTicketsByTech & " AND MemoNumber Not In (Select MemoNumber from FS_ServiceMemosRedispatch) "


	Set cnnNumberOfUrgentServiceTicketsByTech = Server.CreateObject("ADODB.Connection")
	cnnNumberOfUrgentServiceTicketsByTech.open (Session("ClientCnnString"))
	Set rNumberOfUrgentServiceTicketsByTech = Server.CreateObject("ADODB.Recordset")
	rNumberOfUrgentServiceTicketsByTech.CursorLocation = 3 
	Set rNumberOfUrgentServiceTicketsByTech = cnnNumberOfUrgentServiceTicketsByTech.Execute(SQLNumberOfUrgentServiceTicketsByTech)

	If not rNumberOfUrgentServiceTicketsByTech.EOF Then
		Do While Not rNumberOfUrgentServiceTicketsByTech.EOF
			If LastTechUserNo(rNumberOfUrgentServiceTicketsByTech("MemoNumber")) = passedServiceTechNum Then ' If we are not the latest tech, it was reassigned & isnt ours anymore
				If TicketIsUrgent(rNumberOfUrgentServiceTicketsByTech("MemoNumber")) = True Then 
					resultNumberOfUrgentServiceTicketsByTech = resultNumberOfUrgentServiceTicketsByTech + 1
				End If
			End If
			rNumberOfUrgentServiceTicketsByTech.MoveNext
		Loop	
	End IF	
	cnnNumberOfUrgentServiceTicketsByTech.close
	set rNumberOfUrgentServiceTicketsByTech = nothing
	set cnnNumberOfUrgentServiceTicketsByTech= nothing	
	
	NumberOfUrgentServiceTicketsByTech = resultNumberOfUrgentServiceTicketsByTech 
	
End Function

Function FS_TechCanDecline()

	resultFS_TechCanDecline = False
	
	Set cnnFS_TechCanDecline = Server.CreateObject("ADODB.Connection")
	cnnFS_TechCanDecline.open Session("ClientCnnString")

	SQLFS_TechCanDecline = "SELECT FS_TechCanDecline FROM Settings_Global"

	Set rsFS_TechCanDecline = Server.CreateObject("ADODB.Recordset")
	rsFS_TechCanDecline.CursorLocation = 3 
	Set rsFS_TechCanDecline = cnnFS_TechCanDecline.Execute(SQLFS_TechCanDecline)
	
	If not rsFS_TechCanDecline.eof then 
		If rsFS_TechCanDecline("FS_TechCanDecline") = 1 Then resultFS_TechCanDecline = True
	End IF	
	
	set rsFS_TechCanDecline = Nothing
	cnnFS_TechCanDecline.Close
	set cnnFS_TechCanDecline = Nothing
	
	FS_TechCanDecline = resultFS_TechCanDecline


End Function

Function MostRecentDispatchDeclineByTicket(passedTicketNumber)

	resultMostRecentDispatchDeclineByTicket = ""
	
	Set cnnMostRecentDispatchDeclineByTicket = Server.CreateObject("ADODB.Connection")
	cnnMostRecentDispatchDeclineByTicket.open Session("ClientCnnString")

	SQLMostRecentDispatchDeclineByTicket = "SELECT TOP 1 UserNoOfServiceTech FROM FS_ServiceMemosDetail WHERE "
	SQLMostRecentDispatchDeclineByTicket = SQLMostRecentDispatchDeclineByTicket & " MemoStage = 'Dispatch Declined' "
	SQLMostRecentDispatchDeclineByTicket = SQLMostRecentDispatchDeclineByTicket & "	AND MemoNumber = '" & passedTicketNumber & "' ORDER BY UserNoOfServiceTech DESC"

	Set rsMostRecentDispatchDeclineByTicket = Server.CreateObject("ADODB.Recordset")
	rsMostRecentDispatchDeclineByTicket.CursorLocation = 3 
	Set rsMostRecentDispatchDeclineByTicket = cnnMostRecentDispatchDeclineByTicket.Execute(SQLMostRecentDispatchDeclineByTicket)
	
	If not rsMostRecentDispatchDeclineByTicket.eof Then resultMostRecentDispatchDeclineByTicket = rsMostRecentDispatchDeclineByTicket("UserNoOfServiceTech")

	
	set rsMostRecentDispatchDeclineByTicket = Nothing
	cnnMostRecentDispatchDeclineByTicket.Close
	set cnnMostRecentDispatchDeclineByTicket = Nothing
	
	MostRecentDispatchDeclineByTicket = resultMostRecentDispatchDeclineByTicket


End Function

Function MostRecentDispatchDeclineRecordNumberByTicket(passedTicketNumber)

	resultMostRecentDispatchDeclineRecordNumberByTicket = ""
	
	Set cnnMostRecentDispatchDeclineRecordNumberByTicket = Server.CreateObject("ADODB.Connection")
	cnnMostRecentDispatchDeclineRecordNumberByTicket.open Session("ClientCnnString")

	SQLMostRecentDispatchDeclineRecordNumberByTicket = "SELECT TOP 1 ServiceDetailRecNumber FROM FS_ServiceMemosDetail WHERE "
	SQLMostRecentDispatchDeclineRecordNumberByTicket = SQLMostRecentDispatchDeclineRecordNumberByTicket & "  MemoStage = 'Dispatch Declined' "
	SQLMostRecentDispatchDeclineRecordNumberByTicket = SQLMostRecentDispatchDeclineRecordNumberByTicket & "	AND MemoNumber = '" & passedTicketNumber & "' ORDER BY ServiceDetailRecNumber DESC"

	Set rsMostRecentDispatchDeclineRecordNumberByTicket = Server.CreateObject("ADODB.Recordset")
	rsMostRecentDispatchDeclineRecordNumberByTicket.CursorLocation = 3 
	Set rsMostRecentDispatchDeclineRecordNumberByTicket = cnnMostRecentDispatchDeclineRecordNumberByTicket.Execute(SQLMostRecentDispatchDeclineRecordNumberByTicket)
	
	If not rsMostRecentDispatchDeclineRecordNumberByTicket.eof Then resultMostRecentDispatchDeclineRecordNumberByTicket = rsMostRecentDispatchDeclineRecordNumberByTicket("ServiceDetailRecNumber")

	
	set rsMostRecentDispatchDeclineRecordNumberByTicket = Nothing
	cnnMostRecentDispatchDeclineRecordNumberByTicket.Close
	set cnnMostRecentDispatchDeclineRecordNumberByTicket = Nothing
	
	MostRecentDispatchDeclineRecordNumberByTicket = resultMostRecentDispatchDeclineRecordNumberByTicket


End Function

  
Function FSDefaultNotificationMethod()

	resultFSDefaultNotificationMethod = ""
	
	Set cnnFSDefaultNotificationMethod = Server.CreateObject("ADODB.Connection")
	cnnFSDefaultNotificationMethod.open Session("ClientCnnString")

	SQLFSDefaultNotificationMethod = "SELECT FSDefaultNotificationMethod FROM Settings_Global"

	Set rsFSDefaultNotificationMethod = Server.CreateObject("ADODB.Recordset")
	rsFSDefaultNotificationMethod.CursorLocation = 3 
	Set rsFSDefaultNotificationMethod = cnnFSDefaultNotificationMethod.Execute(SQLFSDefaultNotificationMethod)
	
	If not rsFSDefaultNotificationMethod.eof Then resultFSDefaultNotificationMethod = rsFSDefaultNotificationMethod("FSDefaultNotificationMethod")
	
	set rsFSDefaultNotificationMethod = Nothing
	cnnFSDefaultNotificationMethod.Close
	set cnnFSDefaultNotificationMethod = Nothing
	
	FSDefaultNotificationMethod = resultFSDefaultNotificationMethod

End Function


Function GetServiceTicketTechNotesByTicketNumber(passedTicketNumber)

	resultGetServiceTicketTechNotesByTicketNumber = ""
	
	Set cnnGetServiceTicketTechNotesByTicketNumber = Server.CreateObject("ADODB.Connection")
	cnnGetServiceTicketTechNotesByTicketNumber.open Session("ClientCnnString")

	SQLGetServiceTicketTechNotesByTicketNumber = "SELECT * FROM FS_ServiceMemos WHERE MemoNumber = '" & passedTicketNumber & "'"

	Set rsGetServiceTicketTechNotesByTicketNumber = Server.CreateObject("ADODB.Recordset")
	rsGetServiceTicketTechNotesByTicketNumber.CursorLocation = 3 
	Set rsGetServiceTicketTechNotesByTicketNumber = cnnGetServiceTicketTechNotesByTicketNumber.Execute(SQLGetServiceTicketTechNotesByTicketNumber)
	
	If not rsGetServiceTicketTechNotesByTicketNumber.eof then
		resultGetServiceTicketTechNotesByTicketNumber =  rsGetServiceTicketTechNotesByTicketNumber("ServiceNotesFromTech")
	End If

	set rsGetServiceTicketTechNotesByTicketNumber = Nothing
	set cnnGetServiceTicketTechNotesByTicketNumber= Nothing
	
	GetServiceTicketTechNotesByTicketNumber = resultGetServiceTicketTechNotesByTicketNumber

End Function


 Function GetNumberOfServiceTicketsInTimeRange(passedStartMins, passedEndMins)

	resultGetNumberOfServiceTicketsInTimeRange = 0
	
	Set cnnGetNumberOfServiceTicketsInTimeRange = Server.CreateObject("ADODB.Connection")
	cnnGetNumberOfServiceTicketsInTimeRange.open Session("ClientCnnString")

	If FilterChangeModuleOn() Then
		If ShowSeparateFilterChangesTabOnServiceScreen = 1 Then
			SQLGetNumberOfServiceTicketsInTimeRange = "SELECT * FROM FS_ServiceMemos WHERE CurrentStatus = 'OPEN' AND RecordSubType = 'OPEN' AND FilterChange <> 1  ORDER BY submissionDateTime DESC"
		Else
			SQLGetNumberOfServiceTicketsInTimeRange = "SELECT * FROM FS_ServiceMemos WHERE CurrentStatus = 'OPEN' AND RecordSubType = 'OPEN' ORDER BY submissionDateTime DESC"		
		End If
	Else
		SQLGetNumberOfServiceTicketsInTimeRange = "SELECT * FROM FS_ServiceMemos WHERE CurrentStatus = 'OPEN' AND RecordSubType = 'OPEN' AND FilterChange <> 1 ORDER BY submissionDateTime DESC"
	End If

	Set rsGetNumberOfServiceTicketsInTimeRange = Server.CreateObject("ADODB.Recordset")
	rsGetNumberOfServiceTicketsInTimeRange.CursorLocation = 3 
	Set rsGetNumberOfServiceTicketsInTimeRange = cnnGetNumberOfServiceTicketsInTimeRange.Execute(SQLGetNumberOfServiceTicketsInTimeRange)
	
	If not rsGetNumberOfServiceTicketsInTimeRange.eof then
		
		Do While Not rsGetNumberOfServiceTicketsInTimeRange.eof
		
			'elapsedMinutes = ServiceCallElapsedMinutes(rsGetNumberOfServiceTicketsInTimeRange("MemoNumber"))
			elapsedMinutes = ServiceCallElapsedMinutesOpenTicket(rsGetNumberOfServiceTicketsInTimeRange("MemoNumber"))
		
			If elapsedMinutes >= passedStartMins AND elapsedMinutes <= passedEndMins Then
				resultGetNumberOfServiceTicketsInTimeRange =  resultGetNumberOfServiceTicketsInTimeRange + 1
			End If
		
			rsGetNumberOfServiceTicketsInTimeRange.movenext
		Loop
	
		
	End If

	set rsGetNumberOfServiceTicketsInTimeRange = Nothing
	set cnnGetNumberOfServiceTicketsInTimeRange= Nothing
	
	GetNumberOfServiceTicketsInTimeRange = resultGetNumberOfServiceTicketsInTimeRange

End Function

 Function GetNumberOfServiceTicketsInTimeRange2(passedRangeNum)

	resultGetNumberOfServiceTicketsInTimeRange2 = 0
	
	Set cnnGetNumberOfServiceTicketsInTimeRange2 = Server.CreateObject("ADODB.Connection")
	cnnGetNumberOfServiceTicketsInTimeRange2.open Session("ClientCnnString")

	SQLGetNumberOfServiceTicketsInTimeRange2 = "SELECT Distinct MemoNumber  FROM FS_ServiceMemos WHERE CurrentStatus = 'OPEN' AND RecordSubType = 'OPEN' AND FilterChange <> 1"' ORDER BY submissionDateTime DESC"
'Response.Write(SQLGetNumberOfServiceTicketsInTimeRange2 & "<br>")
	Set rsGetNumberOfServiceTicketsInTimeRange2 = Server.CreateObject("ADODB.Recordset")
	rsGetNumberOfServiceTicketsInTimeRange2.CursorLocation = 3 
	Set rsGetNumberOfServiceTicketsInTimeRange2 = cnnGetNumberOfServiceTicketsInTimeRange2.Execute(SQLGetNumberOfServiceTicketsInTimeRange2)
	
	If not rsGetNumberOfServiceTicketsInTimeRange2.eof then
		
		Do While Not rsGetNumberOfServiceTicketsInTimeRange2.eof
		
			'elapsedMinutes = ServiceCallElapsedMinutes(rsGetNumberOfServiceTicketsInTimeRange2("MemoNumber"))
			elapsedMinutes = ServiceCallElapsedMinutesOpenTicket(rsGetNumberOfServiceTicketsInTimeRange2("MemoNumber"))
		
			Select Case passedRangeNum
				Case 1
					If ServiceCallElapsedMinutesOpenTicket(rsGetNumberOfServiceTicketsInTimeRange2.Fields("MemoNumber")) <= GetNumberOfMinutesInServiceDay() Then
						resultGetNumberOfServiceTicketsInTimeRange2 =  resultGetNumberOfServiceTicketsInTimeRange2 + 1
					End If
				Case 2
					If ServiceCallElapsedMinutesOpenTicket(rsGetNumberOfServiceTicketsInTimeRange2.Fields("MemoNumber")) > GetNumberOfMinutesInServiceDay() AND ServiceCallElapsedMinutesOpenTicket(rsGetNumberOfServiceTicketsInTimeRange2.Fields("MemoNumber")) <= GetNumberOfMinutesInServiceDay() * 2 Then
						resultGetNumberOfServiceTicketsInTimeRange2 =  resultGetNumberOfServiceTicketsInTimeRange2 + 1					
					End If
				Case 3
					If ServiceCallElapsedMinutesOpenTicket(rsGetNumberOfServiceTicketsInTimeRange2.Fields("MemoNumber")) > GetNumberOfMinutesInServiceDay() * 2 AND ServiceCallElapsedMinutesOpenTicket(rsGetNumberOfServiceTicketsInTimeRange2.Fields("MemoNumber")) <= GetNumberOfMinutesInServiceDay() * 5 Then
						resultGetNumberOfServiceTicketsInTimeRange2 =  resultGetNumberOfServiceTicketsInTimeRange2 + 1						
					End If
				Case 4
					If ServiceCallElapsedMinutesOpenTicket(rsGetNumberOfServiceTicketsInTimeRange2.Fields("MemoNumber")) > GetNumberOfMinutesInServiceDay() * 5 Then
						resultGetNumberOfServiceTicketsInTimeRange2 =  resultGetNumberOfServiceTicketsInTimeRange2 + 1						
					End If
			End Select
			
		
			rsGetNumberOfServiceTicketsInTimeRange2.movenext
		Loop
	
		
	End If

	set rsGetNumberOfServiceTicketsInTimeRange2 = Nothing
	set cnnGetNumberOfServiceTicketsInTimeRange2= Nothing
	
	GetNumberOfServiceTicketsInTimeRange2 = resultGetNumberOfServiceTicketsInTimeRange2

End Function




Function GetNumberOfFilterTicketsInTimeRange(passedStartMins, passedEndMins)

	resultGetNumberOfFilterTicketsInTimeRange = 0
	
	Set cnnGetNumberOfFilterTicketsInTimeRange = Server.CreateObject("ADODB.Connection")
	cnnGetNumberOfFilterTicketsInTimeRange.open Session("ClientCnnString")

	SQLGetNumberOfFilterTicketsInTimeRange = "SELECT * FROM FS_ServiceMemos WHERE CurrentStatus = 'OPEN' AND RecordSubType = 'OPEN' AND FilterChange = 1 ORDER BY submissionDateTime DESC"

	Set rsGetNumberOfFilterTicketsInTimeRange = Server.CreateObject("ADODB.Recordset")
	rsGetNumberOfFilterTicketsInTimeRange.CursorLocation = 3 
	Set rsGetNumberOfFilterTicketsInTimeRange = cnnGetNumberOfFilterTicketsInTimeRange.Execute(SQLGetNumberOfFilterTicketsInTimeRange)
	
	If not rsGetNumberOfFilterTicketsInTimeRange.eof then
		
		Do While Not rsGetNumberOfFilterTicketsInTimeRange.eof
		
			'elapsedMinutes = ServiceCallElapsedMinutes(rsGetNumberOfFilterTicketsInTimeRange("MemoNumber"))
			elapsedMinutes = ServiceCallElapsedMinutesOpenTicket(rsGetNumberOfFilterTicketsInTimeRange("MemoNumber"))
		
			If elapsedMinutes >= passedStartMins AND elapsedMinutes <= passedEndMins Then
				resultGetNumberOfFilterTicketsInTimeRange =  resultGetNumberOfFilterTicketsInTimeRange + 1
			End If
		
			rsGetNumberOfFilterTicketsInTimeRange.movenext
		Loop
	
		
	End If

	set rsGetNumberOfFilterTicketsInTimeRange = Nothing
	set cnnGetNumberOfFilterTicketsInTimeRange= Nothing
	
	GetNumberOfFilterTicketsInTimeRange = resultGetNumberOfFilterTicketsInTimeRange

End Function


Function GetNumberOfFilterChangesInTimeRange(passedStartMins, passedEndMins)

	resultGetNumberOfFilterChangesInTimeRange = 0
	
	Set cnnGetNumberOfFilterChangesInTimeRange = Server.CreateObject("ADODB.Connection")
	cnnGetNumberOfFilterChangesInTimeRange.open Session("ClientCnnString")

	SQLGetNumberOfFilterChangesInTimeRange = "SELECT * FROM FS_ServiceMemos WHERE CurrentStatus = 'OPEN' AND RecordSubType = 'OPEN' AND FilterChange = 1 ORDER BY submissionDateTime DESC"

	Set rsGetNumberOfFilterChangesInTimeRange = Server.CreateObject("ADODB.Recordset")
	rsGetNumberOfFilterChangesInTimeRange.CursorLocation = 3 
	Set rsGetNumberOfFilterChangesInTimeRange = cnnGetNumberOfFilterChangesInTimeRange.Execute(SQLGetNumberOfFilterChangesInTimeRange)
	
	If not rsGetNumberOfFilterChangesInTimeRange.eof then
		
		Do While Not rsGetNumberOfFilterChangesInTimeRange.eof
		
			elapsedMinutes = ServiceCallElapsedMinutesOpenTicket(rsGetNumberOfFilterChangesInTimeRange("MemoNumber"))
		
			If elapsedMinutes >= passedStartMins AND elapsedMinutes <= passedEndMins Then
				resultGetNumberOfFilterChangesInTimeRange = resultGetNumberOfFilterChangesInTimeRange + GetNumberOfFilterChangesForServiceTicket(rsGetNumberOfFilterChangesInTimeRange("MemoNumber"))
			End If
		
			rsGetNumberOfFilterChangesInTimeRange.movenext
		Loop
	
		
	End If

	set rsGetNumberOfFilterChangesInTimeRange = Nothing
	set cnnGetNumberOfFilterChangesInTimeRange= Nothing
	
	GetNumberOfFilterChangesInTimeRange = resultGetNumberOfFilterChangesInTimeRange

End Function


Function GetNumberOfServiceCallsAwaitingDispatch()

	resultNumOfServiceCallsAwaitingDispatch = 0
	
	SQL = "SELECT * FROM FS_ServiceMemos WHERE CurrentStatus = 'OPEN' AND RecordSubType = 'OPEN'  AND FilterChange <> 1  ORDER BY submissionDateTime DESC"

	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rs = cnn8.Execute(SQL)


	If not rs.EOF Then

		Do While Not rs.EOF
		
			GetServiceTicketCurrentStageVar = GetServiceTicketCurrentStage(rs.Fields("MemoNumber"))
		
			If rs.Fields("RecordSubType") <> "HOLD" AND (GetServiceTicketCurrentStageVar = "Received" OR GetServiceTicketCurrentStageVar = "Released" OR GetServiceTicketCurrentStageVar = "Declined") Then
				
				If rs.Fields("CurrentStatus") = rs.Fields("RecordSubType") Then ' Show only 1 line per memo, the most current status
	
					resultNumOfServiceCallsAwaitingDispatch = resultNumOfServiceCallsAwaitingDispatch + 1				
				
				End If
				
			
			End If 'End Awaiting Dispatch Check 
			
			rs.movenext
		loop
		
	End If

	set rs = Nothing
	cnn8.close
	set cnn8 = Nothing
	
	GetNumberOfServiceCallsAwaitingDispatch = resultNumOfServiceCallsAwaitingDispatch
	
End Function



Function GetNumberOfServiceCallsAwaitingAcknowledgement()

	Set cnnNumOfServiceCallsAwaitingAcknowledgement = Server.CreateObject("ADODB.Connection")
	cnnNumOfServiceCallsAwaitingAcknowledgement.open Session("ClientCnnString")

	resultNumOfServiceCallsAwaitingAcknowledgement = 0
	
	'Where there is a dispatch, but not a dispatch acknowledged
	
    SQLNumOfServiceCallsAwaitingAcknowledgement = "SELECT DISTINCT MemoNumber FROM FS_ServiceMemos WHERE CurrentStatus = 'OPEN' AND RecordSubType = 'OPEN' AND FilterChange <> 1 "   
	
	'Response.Write(SQLNumOfServiceCallsAwaitingAcknowledgement)
	 
	Set rsNumOfServiceCallsAwaitingAcknowledgement = Server.CreateObject("ADODB.Recordset")
	rsNumOfServiceCallsAwaitingAcknowledgement.CursorLocation = 3 
	
	rsNumOfServiceCallsAwaitingAcknowledgement.Open SQLNumOfServiceCallsAwaitingAcknowledgement , cnnNumOfServiceCallsAwaitingAcknowledgement
	
	
	If NOT rsNumOfServiceCallsAwaitingAcknowledgement.EOF Then
	
		Do While NOT rsNumOfServiceCallsAwaitingAcknowledgement.EOF
		
			GetServiceTicketCurrentStageVar = GetServiceTicketCurrentStage(rsNumOfServiceCallsAwaitingAcknowledgement.Fields("MemoNumber"))
		
			If GetServiceTicketCurrentStageVar = "Dispatched" Then
			
				resultNumOfServiceCallsAwaitingAcknowledgement = resultNumOfServiceCallsAwaitingAcknowledgement + 1
				
			End If
		
		rsNumOfServiceCallsAwaitingAcknowledgement.MoveNext
		
		Loop
		
	End If

	
	rsNumOfServiceCallsAwaitingAcknowledgement.Close
	set rsNumOfServiceCallsAwaitingAcknowledgement= Nothing
	cnnNumOfServiceCallsAwaitingAcknowledgement.Close
	set cnnNumOfServiceCallsAwaitingAcknowledgement= Nothing
	
	GetNumberOfServiceCallsAwaitingAcknowledgement = resultNumOfServiceCallsAwaitingAcknowledgement
	
End Function



Function GetNumberOfServiceCallsAcknowledged()

	Set cnnNumOfServiceCallsAcknowledged = Server.CreateObject("ADODB.Connection")
	cnnNumOfServiceCallsAcknowledged.open Session("ClientCnnString")

	resultNumOfServiceCallsAcknowledged = 0
	
    SQLNumOfServiceCallsAcknowledged = "SELECT DISTINCT MemoNumber FROM FS_ServiceMemos WHERE CurrentStatus = 'OPEN' AND RecordSubType = 'OPEN'  AND FilterChange <> 1 "    

	'Response.Write(SQLNumOfServiceCallsAcknowledged)
	 
	Set rsNumOfServiceCallsAcknowledged = Server.CreateObject("ADODB.Recordset")
	rsNumOfServiceCallsAcknowledged.CursorLocation = 3 
	
	rsNumOfServiceCallsAcknowledged.Open SQLNumOfServiceCallsAcknowledged , cnnNumOfServiceCallsAcknowledged
	
	
	If NOT rsNumOfServiceCallsAcknowledged.EOF Then
	
		Do While NOT rsNumOfServiceCallsAcknowledged.EOF
		
			GetServiceTicketCurrentStageVar = GetServiceTicketCurrentStage(rsNumOfServiceCallsAcknowledged.Fields("MemoNumber"))
		
			If GetServiceTicketCurrentStageVar = "Dispatch Acknowledged" Then
			
				resultNumOfServiceCallsAcknowledged = resultNumOfServiceCallsAcknowledged + 1
				
			End If
		
		rsNumOfServiceCallsAcknowledged.MoveNext
		
		Loop
		
	End If
		
	rsNumOfServiceCallsAcknowledged.Close
	set rsNumOfServiceCallsAcknowledged= Nothing
	cnnNumOfServiceCallsAcknowledged.Close
	set cnnNumOfServiceCallsAcknowledged= Nothing
	
	GetNumberOfServiceCallsAcknowledged = resultNumOfServiceCallsAcknowledged
	
End Function


Function GetNumberOfServiceCallsEnRouteOnSite()

	Set cnnNumOfServiceCallsEnRouteOnSite = Server.CreateObject("ADODB.Connection")
	cnnNumOfServiceCallsEnRouteOnSite.open Session("ClientCnnString")

	resultNumOfServiceCallsEnRouteOnSite = 0
	
    SQLNumOfServiceCallsEnRouteOnSite = "SELECT DISTINCT MemoNumber FROM FS_ServiceMemos WHERE CurrentStatus = 'OPEN' AND RecordSubType = 'OPEN'  AND FilterChange <> 1 "
    
	'Response.Write(SQLNumOfServiceCallsEnRouteOnSite)
	 
	Set rsNumOfServiceCallsEnRouteOnSite = Server.CreateObject("ADODB.Recordset")
	rsNumOfServiceCallsEnRouteOnSite.CursorLocation = 3 
	
	rsNumOfServiceCallsEnRouteOnSite.Open SQLNumOfServiceCallsEnRouteOnSite , cnnNumOfServiceCallsEnRouteOnSite
	
	If NOT rsNumOfServiceCallsEnRouteOnSite.EOF Then
	
		Do While NOT rsNumOfServiceCallsEnRouteOnSite.EOF
		
			GetServiceTicketCurrentStageVar = GetServiceTicketCurrentStage(rsNumOfServiceCallsEnRouteOnSite.Fields("MemoNumber"))
		
			If GetServiceTicketCurrentStageVar = "En Route" OR GetServiceTicketCurrentStageVar = "On Site" Then
			
				resultNumOfServiceCallsEnRouteOnSite = resultNumOfServiceCallsEnRouteOnSite + 1
				
			End If
		
		rsNumOfServiceCallsEnRouteOnSite.MoveNext
		
		Loop
		
	End If
	
	rsNumOfServiceCallsEnRouteOnSite.Close
	set rsNumOfServiceCallsEnRouteOnSite= Nothing
	cnnNumOfServiceCallsEnRouteOnSite.Close
	set cnnNumOfServiceCallsEnRouteOnSite= Nothing
	
	GetNumberOfServiceCallsEnRouteOnSite = resultNumOfServiceCallsEnRouteOnSite
	
End Function



Function GetNumberOfServiceCallsRedo()

	Set cnnNumOfServiceCallsRedo = Server.CreateObject("ADODB.Connection")
	cnnNumOfServiceCallsRedo.open Session("ClientCnnString")

	resultNumOfServiceCallsRedo = 0
	
	SQLNumOfServiceCallsRedo = "SELECT DISTINCT MemoNumber FROM FS_ServiceMemos WHERE CurrentStatus = 'OPEN' AND RecordSubType = 'OPEN'  AND FilterChange <> 1 "

	'Response.Write(SQLNumOfServiceCallsRedo)
	 
	Set rsNumOfServiceCallsRedo = Server.CreateObject("ADODB.Recordset")
	rsNumOfServiceCallsRedo.CursorLocation = 3 
	
	rsNumOfServiceCallsRedo.Open SQLNumOfServiceCallsRedo , cnnNumOfServiceCallsRedo
	

	If NOT rsNumOfServiceCallsRedo.EOF Then
	
		Do While NOT rsNumOfServiceCallsRedo.EOF
		
			GetServiceTicketCurrentStageVar = GetServiceTicketCurrentStage(rsNumOfServiceCallsRedo.Fields("MemoNumber"))
		
			If GetServiceTicketCurrentStageVar = "Unable To Work" OR  GetServiceTicketCurrentStageVar = "Swap" OR GetServiceTicketCurrentStageVar = "Wait for parts" OR GetServiceTicketCurrentStageVar = "Follow Up" Then
			
				resultNumOfServiceCallsRedo = resultNumOfServiceCallsRedo + 1
				
			End If
		
		rsNumOfServiceCallsRedo.MoveNext
		
		Loop
		
	End If

	
	rsNumOfServiceCallsRedo.Close
	set rsNumOfServiceCallsRedo= Nothing
	cnnNumOfServiceCallsRedo.Close
	set cnnNumOfServiceCallsRedo= Nothing
	
	GetNumberOfServiceCallsRedo = resultNumOfServiceCallsRedo
	
End Function




Function GetNumberOfServiceCallsRedoFiltersOnly()

	Set cnnNumOfServiceCallsRedoFiltersOnly = Server.CreateObject("ADODB.Connection")
	cnnNumOfServiceCallsRedoFiltersOnly.open Session("ClientCnnString")

	resultNumOfServiceCallsRedoFiltersOnly = 0
	
	SQLNumOfServiceCallsRedoFiltersOnly = "SELECT DISTINCT MemoNumber FROM FS_ServiceMemos WHERE CurrentStatus = 'OPEN' AND RecordSubType = 'OPEN'  AND FilterChange = 1 "

	'Response.Write(SQLNumOfServiceCallsRedoFiltersOnly)
	 
	Set rsNumOfServiceCallsRedoFiltersOnly = Server.CreateObject("ADODB.Recordset")
	rsNumOfServiceCallsRedoFiltersOnly.CursorLocation = 3 
	
	rsNumOfServiceCallsRedoFiltersOnly.Open SQLNumOfServiceCallsRedoFiltersOnly , cnnNumOfServiceCallsRedoFiltersOnly
	

	If NOT rsNumOfServiceCallsRedoFiltersOnly.EOF Then
	
		Do While NOT rsNumOfServiceCallsRedoFiltersOnly.EOF
		
			GetServiceTicketCurrentStageVar = GetServiceTicketCurrentStage(rsNumOfServiceCallsRedoFiltersOnly.Fields("MemoNumber"))
		
			If GetServiceTicketCurrentStageVar = "Unable To Work" OR  GetServiceTicketCurrentStageVar = "Swap" OR GetServiceTicketCurrentStageVar = "Wait for parts" OR GetServiceTicketCurrentStageVar = "Follow Up" Then
			
				resultNumOfServiceCallsRedoFiltersOnly = resultNumOfServiceCallsRedoFiltersOnly + 1
				
			End If
		
		rsNumOfServiceCallsRedoFiltersOnly.MoveNext
		
		Loop
		
	End If

	
	rsNumOfServiceCallsRedoFiltersOnly.Close
	set rsNumOfServiceCallsRedoFiltersOnly= Nothing
	cnnNumOfServiceCallsRedoFiltersOnly.Close
	set cnnNumOfServiceCallsRedoFiltersOnly= Nothing
	
	GetNumberOfServiceCallsRedoFiltersOnly = resultNumOfServiceCallsRedoFiltersOnly
	
End Function


Function GetNumberOfServiceCallsAwaitingDispatchWithFilters()

	resultNumOfServiceCallsAwaitingDispatchWithFilters = 0
	
	SQL = "SELECT * FROM FS_ServiceMemos WHERE CurrentStatus = 'OPEN' AND RecordSubType = 'OPEN' ORDER BY submissionDateTime DESC"

	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rs = cnn8.Execute(SQL)


	If not rs.EOF Then

		Do While Not rs.EOF
		
			GetServiceTicketCurrentStageVar = GetServiceTicketCurrentStage(rs.Fields("MemoNumber"))
		
			If rs.Fields("RecordSubType") <> "HOLD" AND (GetServiceTicketCurrentStageVar = "Received" OR GetServiceTicketCurrentStageVar = "Released" OR GetServiceTicketCurrentStageVar = "Declined") Then
				
				If rs.Fields("CurrentStatus") = rs.Fields("RecordSubType") Then ' Show only 1 line per memo, the most current status
	
					resultNumOfServiceCallsAwaitingDispatchWithFilters = resultNumOfServiceCallsAwaitingDispatchWithFilters + 1				
				
				End If
				
			
			End If 'End Awaiting Dispatch Check 
			
			rs.movenext
		loop
		
	End If

	set rs = Nothing
	cnn8.close
	set cnn8 = Nothing
	
	GetNumberOfServiceCallsAwaitingDispatchWithFilters = resultNumOfServiceCallsAwaitingDispatchWithFilters
	
End Function



Function GetNumberOfServiceCallsAwaitingAcknowledgementWithFilters()

	Set cnnNumOfServiceCallsAwaitingAcknowledgementWithFilters = Server.CreateObject("ADODB.Connection")
	cnnNumOfServiceCallsAwaitingAcknowledgementWithFilters.open Session("ClientCnnString")

	resultNumOfServiceCallsAwaitingAcknowledgementWithFilters = 0
	
	'Where there is a dispatch, but not a dispatch acknowledged
	
    SQLNumOfServiceCallsAwaitingAcknowledgementWithFilters = "SELECT DISTINCT MemoNumber FROM FS_ServiceMemos WHERE CurrentStatus = 'OPEN' AND RecordSubType = 'OPEN'"   
	
	'Response.Write(SQLNumOfServiceCallsAwaitingAcknowledgementWithFilters)
	 
	Set rsNumOfServiceCallsAwaitingAcknowledgementWithFilters = Server.CreateObject("ADODB.Recordset")
	rsNumOfServiceCallsAwaitingAcknowledgementWithFilters.CursorLocation = 3 
	
	rsNumOfServiceCallsAwaitingAcknowledgementWithFilters.Open SQLNumOfServiceCallsAwaitingAcknowledgementWithFilters , cnnNumOfServiceCallsAwaitingAcknowledgementWithFilters
	
	
	If NOT rsNumOfServiceCallsAwaitingAcknowledgementWithFilters.EOF Then
	
		Do While NOT rsNumOfServiceCallsAwaitingAcknowledgementWithFilters.EOF
		
			GetServiceTicketCurrentStageVar = GetServiceTicketCurrentStage(rsNumOfServiceCallsAwaitingAcknowledgementWithFilters.Fields("MemoNumber"))
		
			If GetServiceTicketCurrentStageVar = "Dispatched" Then
			
				resultNumOfServiceCallsAwaitingAcknowledgementWithFilters = resultNumOfServiceCallsAwaitingAcknowledgementWithFilters + 1
				
			End If
		
		rsNumOfServiceCallsAwaitingAcknowledgementWithFilters.MoveNext
		
		Loop
		
	End If

	
	rsNumOfServiceCallsAwaitingAcknowledgementWithFilters.Close
	set rsNumOfServiceCallsAwaitingAcknowledgementWithFilters= Nothing
	cnnNumOfServiceCallsAwaitingAcknowledgementWithFilters.Close
	set cnnNumOfServiceCallsAwaitingAcknowledgementWithFilters= Nothing
	
	GetNumberOfServiceCallsAwaitingAcknowledgementWithFilters = resultNumOfServiceCallsAwaitingAcknowledgementWithFilters
	
End Function



Function GetNumberOfServiceCallsAcknowledgedWithFilters()

	Set cnnNumOfServiceCallsAcknowledgedWithFilters = Server.CreateObject("ADODB.Connection")
	cnnNumOfServiceCallsAcknowledgedWithFilters.open Session("ClientCnnString")

	resultNumOfServiceCallsAcknowledgedWithFilters = 0
	
    SQLNumOfServiceCallsAcknowledgedWithFilters = "SELECT DISTINCT MemoNumber FROM FS_ServiceMemos WHERE CurrentStatus = 'OPEN' AND RecordSubType = 'OPEN'"    

	'Response.Write(SQLNumOfServiceCallsAcknowledgedWithFilters)
	 
	Set rsNumOfServiceCallsAcknowledgedWithFilters = Server.CreateObject("ADODB.Recordset")
	rsNumOfServiceCallsAcknowledgedWithFilters.CursorLocation = 3 
	
	rsNumOfServiceCallsAcknowledgedWithFilters.Open SQLNumOfServiceCallsAcknowledgedWithFilters , cnnNumOfServiceCallsAcknowledgedWithFilters
	
	
	If NOT rsNumOfServiceCallsAcknowledgedWithFilters.EOF Then
	
		Do While NOT rsNumOfServiceCallsAcknowledgedWithFilters.EOF
		
			GetServiceTicketCurrentStageVar = GetServiceTicketCurrentStage(rsNumOfServiceCallsAcknowledgedWithFilters.Fields("MemoNumber"))
		
			If GetServiceTicketCurrentStageVar = "Dispatch Acknowledged" Then
			
				resultNumOfServiceCallsAcknowledgedWithFilters = resultNumOfServiceCallsAcknowledgedWithFilters + 1
				
			End If
		
		rsNumOfServiceCallsAcknowledgedWithFilters.MoveNext
		
		Loop
		
	End If
		
	rsNumOfServiceCallsAcknowledgedWithFilters.Close
	set rsNumOfServiceCallsAcknowledgedWithFilters= Nothing
	cnnNumOfServiceCallsAcknowledgedWithFilters.Close
	set cnnNumOfServiceCallsAcknowledgedWithFilters= Nothing
	
	GetNumberOfServiceCallsAcknowledgedWithFilters = resultNumOfServiceCallsAcknowledgedWithFilters
	
End Function


Function GetNumberOfServiceCallsEnRouteOnSiteWithFilters()

	Set cnnNumOfServiceCallsEnRouteOnSiteWithFilters = Server.CreateObject("ADODB.Connection")
	cnnNumOfServiceCallsEnRouteOnSiteWithFilters.open Session("ClientCnnString")

	resultNumOfServiceCallsEnRouteOnSiteWithFilters = 0
	
    SQLNumOfServiceCallsEnRouteOnSiteWithFilters = "SELECT DISTINCT MemoNumber FROM FS_ServiceMemos WHERE CurrentStatus = 'OPEN' AND RecordSubType = 'OPEN'"
    
	'Response.Write(SQLNumOfServiceCallsEnRouteOnSiteWithFilters)
	 
	Set rsNumOfServiceCallsEnRouteOnSiteWithFilters = Server.CreateObject("ADODB.Recordset")
	rsNumOfServiceCallsEnRouteOnSiteWithFilters.CursorLocation = 3 
	
	rsNumOfServiceCallsEnRouteOnSiteWithFilters.Open SQLNumOfServiceCallsEnRouteOnSiteWithFilters , cnnNumOfServiceCallsEnRouteOnSiteWithFilters
	
	If NOT rsNumOfServiceCallsEnRouteOnSiteWithFilters.EOF Then
	
		Do While NOT rsNumOfServiceCallsEnRouteOnSiteWithFilters.EOF
		
			GetServiceTicketCurrentStageVar = GetServiceTicketCurrentStage(rsNumOfServiceCallsEnRouteOnSiteWithFilters.Fields("MemoNumber"))
		
			If GetServiceTicketCurrentStageVar = "En Route" OR GetServiceTicketCurrentStageVar = "On Site" Then
			
				resultNumOfServiceCallsEnRouteOnSiteWithFilters = resultNumOfServiceCallsEnRouteOnSiteWithFilters + 1
				
			End If
		
		rsNumOfServiceCallsEnRouteOnSiteWithFilters.MoveNext
		
		Loop
		
	End If
	
	rsNumOfServiceCallsEnRouteOnSiteWithFilters.Close
	set rsNumOfServiceCallsEnRouteOnSiteWithFilters= Nothing
	cnnNumOfServiceCallsEnRouteOnSiteWithFilters.Close
	set cnnNumOfServiceCallsEnRouteOnSiteWithFilters= Nothing
	
	GetNumberOfServiceCallsEnRouteOnSiteWithFilters = resultNumOfServiceCallsEnRouteOnSiteWithFilters
	
End Function




Function GetNumberOfServiceCallsRedoWithFilters()

	Set cnnNumOfServiceCallsRedoWithFilters = Server.CreateObject("ADODB.Connection")
	cnnNumOfServiceCallsRedoWithFilters.open Session("ClientCnnString")

	resultNumOfServiceCallsRedoWithFilters = 0
	
	SQLNumOfServiceCallsRedoWithFilters = "SELECT DISTINCT MemoNumber FROM FS_ServiceMemos WHERE CurrentStatus = 'OPEN' AND RecordSubType = 'OPEN'"

	'Response.Write(SQLNumOfServiceCallsRedoWithFilters)
	 
	Set rsNumOfServiceCallsRedoWithFilters = Server.CreateObject("ADODB.Recordset")
	rsNumOfServiceCallsRedoWithFilters.CursorLocation = 3 
	
	rsNumOfServiceCallsRedoWithFilters.Open SQLNumOfServiceCallsRedoWithFilters , cnnNumOfServiceCallsRedoWithFilters
	

	If NOT rsNumOfServiceCallsRedoWithFilters.EOF Then
	
		Do While NOT rsNumOfServiceCallsRedoWithFilters.EOF
		
			GetServiceTicketCurrentStageVar = GetServiceTicketCurrentStage(rsNumOfServiceCallsRedoWithFilters.Fields("MemoNumber"))
		
			If GetServiceTicketCurrentStageVar = "Unable To Work" OR  GetServiceTicketCurrentStageVar = "Swap" OR GetServiceTicketCurrentStageVar = "Wait for parts" OR GetServiceTicketCurrentStageVar = "Follow Up" Then
			
				resultNumOfServiceCallsRedoWithFilters = resultNumOfServiceCallsRedoWithFilters + 1
				
			End If
		
		rsNumOfServiceCallsRedoWithFilters.MoveNext
		
		Loop
		
	End If

	
	rsNumOfServiceCallsRedoWithFilters.Close
	set rsNumOfServiceCallsRedoWithFilters= Nothing
	cnnNumOfServiceCallsRedoWithFilters.Close
	set cnnNumOfServiceCallsRedoWithFilters= Nothing
	
	GetNumberOfServiceCallsRedoWithFilters = resultNumOfServiceCallsRedoWithFilters
	
End Function




Function GetNumberOfServiceCallsClosedThisWeek()

	'This returns all the services calls closed between Saturday-Friday of the current week
	'We want to look at RecordCreatedateTime not SubmissionDateTime
	
	'******************************************************************************
	'First Get Date of Last Saturday
	'******************************************************************************
	
	'WeekDay() returns 1 - 7 (Sunday - Saturday).
	today = WeekDay(Date())
	
	'Workout the offset then use DateAdd() to minus that number of days.
	Select Case today
	Case 1 'Sunday
	  offsetDaysSaturday = 1
	Case 2 'Monday
	  offsetDaysSaturday = 2
	Case 3 'Tuesday
	  offsetDaysSaturday = 3
	Case 4 'Wednesday
	  offsetDaysSaturday = 4
	Case 5 'Thursday
	  offsetDaysSaturday = 5
	Case 6 'Friday
	  offsetDaysSaturday = 6
	Case 7 'Saturday
	  offsetDaysSaturday = 7
	End Select
	
	lastSaturday = DateAdd("d", -offsetDaysSaturday, Date())	
	
	'******************************************************************************
	
	'******************************************************************************
	'Then Get Date of This Friday
	'******************************************************************************
		
	'Workout the offset then use DateAdd() to add that number of days.
	Select Case today
	Case 1 'Sunday
	  offsetDaysFriday = 5
	Case 2 'Monday
	  offsetDaysFriday = 4
	Case 3 'Tuesday
	  offsetDaysFriday = 3
	Case 4 'Wednesday
	  offsetDaysFriday = 1
	Case 5 'Thursday
	  offsetDaysFriday = 2
	Case 6 'Friday
	  offsetDaysFriday = 0
	Case 7 'Saturday
	  offsetDaysFriday = 6
	End Select
	
	thisFriday = DateAdd("d", offsetDaysFriday, Date())	
	
	'******************************************************************************
	

	Set cnnNumOfServiceCallsClosedThisWeek = Server.CreateObject("ADODB.Connection")
	cnnNumOfServiceCallsClosedThisWeek.open Session("ClientCnnString")
	
	resultNumOfServiceCallsClosedThisWeek = 0
	
	SQLNumOfServiceCallsClosedThisWeek = "SELECT DISTINCT MemoNumber FROM FS_ServiceMemos WHERE CurrentStatus = 'CLOSE' AND RecordSubType = 'CLOSE'"
	SQLNumOfServiceCallsClosedThisWeek = SQLNumOfServiceCallsClosedThisWeek & " AND RecordCreatedateTime >= '" & lastSaturday & "' AND RecordCreatedateTime <= '" & thisFriday & "' "
	 
	'Response.Write(SQLNumOfServiceCallsClosedThisWeek)
	 
	Set rsNumOfServiceCallsClosedThisWeek = Server.CreateObject("ADODB.Recordset")
	rsNumOfServiceCallsClosedThisWeek.CursorLocation = 3 
	
	rsNumOfServiceCallsClosedThisWeek.Open SQLNumOfServiceCallsClosedThisWeek , cnnNumOfServiceCallsClosedThisWeek
	
	resultNumOfServiceCallsClosedThisWeek = rsNumOfServiceCallsClosedThisWeek.RecordCount
	
	rsNumOfServiceCallsClosedThisWeek.Close
	set rsNumOfServiceCallsClosedThisWeek= Nothing
	cnnNumOfServiceCallsClosedThisWeek.Close
	set cnnNumOfServiceCallsClosedThisWeek= Nothing
	
	GetNumberOfServiceCallsClosedThisWeek = resultNumOfServiceCallsClosedThisWeek
	
End Function


Function GetNumberOfServiceCallsClosedRolling5Days()


	resultNumOfServiceCallsClosedRolling5Days  = 0


	'******************************************************************************
	'Obtain Dates of Rolling Last Five Work Days
	'******************************************************************************

	'*******************************************************************************
	'Obtain the first working day to start counting 5 days backwards from
	'Start with today. If today is a weekend or a closed company holiday, 
	'then subtract one day and repeat the check. Once we find a valid workday,
	'stop looping and start counting back 5 valid business days from the
	'starting date
	'*******************************************************************************
	
	firstDateOfPastFiveDaysFound = False
	firstDateOfPastFiveDays = Now()
	
	Do While firstDateOfPastFiveDaysFound = False
	
		SQL = "SELECT * FROM Settings_CompanyCalendar WHERE MonthNum='" & Month(firstDateOfPastFiveDays) & "' AND DayNum ='" & Day(firstDateOfPastFiveDays) & "' AND YearNum='" & Year(firstDateOfPastFiveDays) & "'"
	
		Set cnn9 = Server.CreateObject("ADODB.Connection")
		cnn9.open (Session("ClientCnnString"))
		Set rs9 = Server.CreateObject("ADODB.Recordset")
		rs9.CursorLocation = 3 
		Set rs9 = cnn9.Execute(SQL)
			
		If not rs9.EOF Then			
			Select Case rs9("OpenClosedCloseEarly")
			Case "Closed"
				firstDateOfPastFiveDaysFound = False
			Case "Close Early"
				firstDateOfPastFiveDaysFound = True
			Case Else
				firstDateOfPastFiveDaysFound = True
			End Select
		Else
			firstDateOfPastFiveDaysFound = True
		End If
					
		'******************************************************
		'Make sure that that the date is also not a weekend
		'If it is, set the control variable to false
		'And go back one calendar day to test as the next day
		'******************************************************
		
		If firstDateOfPastFiveDaysFound = True Then			
			If Weekday(firstDateOfPastFiveDays,vbMonday) <= 5 Then
				firstDateOfPastFiveDaysFound = True
			Else
				firstDateOfPastFiveDaysFound = False
				firstDateOfPastFiveDays = DateAdd("d", -1, firstDateOfPastFiveDays)
			End If
		Else
			firstDateOfPastFiveDays = DateAdd("d", -1, firstDateOfPastFiveDays)
		End If	
		
		
	
	Loop
	
	
	
	'*************************************************************
	'At this point we have found the ending business day
	'Now we need to loop through the days prior to this day and
	'come up with the last 5 business days starting date
	'*************************************************************
	
	'Response.Write("firstDateOfPastFiveDays : " & firstDateOfPastFiveDays & "<br>")
	lastDateOfPastFiveDays = DateAdd("d", -1, firstDateOfPastFiveDays)
	validDaysGoneBackSoFar = 1
	
	Do While validDaysGoneBackSoFar < 4
	
		lastDateOfPastFiveDaysFound = False
	
		SQL = "SELECT * FROM Settings_CompanyCalendar WHERE MonthNum='" & Month(lastDateOfPastFiveDays) & "' AND DayNum ='" & Day(lastDateOfPastFiveDays) & "' AND YearNum='" & Year(lastDateOfPastFiveDays) & "'"
	
		Set cnn9 = Server.CreateObject("ADODB.Connection")
		cnn9.open (Session("ClientCnnString"))
		Set rs9 = Server.CreateObject("ADODB.Recordset")
		rs9.CursorLocation = 3 
		Set rs9 = cnn9.Execute(SQL)
			
		If not rs9.EOF Then
			Select Case rs9("OpenClosedCloseEarly")
			Case "Closed"
				lastDateOfPastFiveDaysFound = False
			Case "Close Early"
				lastDateOfPastFiveDaysFound = True
			Case Else
				lastDateOfPastFiveDaysFound = True
			End Select
		Else
			lastDateOfPastFiveDaysFound = True
		End If
					
		'******************************************************
		'Make sure that that the date is also not a weekend
		'If it is, set the control variable to false
		'And go back one calendar day to test as the next day
		'******************************************************

		If lastDateOfPastFiveDaysFound = True Then				
			If Weekday(lastDateOfPastFiveDays,vbMonday) <= 5 Then
				lastDateOfPastFiveDaysFound = True
				validDaysGoneBackSoFar = validDaysGoneBackSoFar + 1
				lastDateOfPastFiveDays = DateAdd("d", -1, lastDateOfPastFiveDays)
			Else
				lastDateOfPastFiveDaysFound = False
				lastDateOfPastFiveDays = DateAdd("d", -1, lastDateOfPastFiveDays)
			End If
		Else
			lastDateOfPastFiveDays = DateAdd("d", -1, lastDateOfPastFiveDays)
		End If	
	
		
		'Response.Write("lastDateOfPastFiveDays : " & lastDateOfPastFiveDays & "<br>")
		'Response.Write("validDaysGoneBackSoFar : " & validDaysGoneBackSoFar & "<br>")
	
	Loop


	set rs9 = Nothing
	cnn9.close
	set cnn9 = Nothing
						

	
	Set cnnNumOfServiceCallsClosedRolling5Days = Server.CreateObject("ADODB.Connection")
	cnnNumOfServiceCallsClosedRolling5Days.open Session("ClientCnnString")
	
	resultNumOfServiceCallsClosedRolling5Days = 0
	
	SQLNumOfServiceCallsClosedRolling5Days = "SELECT DISTINCT MemoNumber FROM FS_ServiceMemos WHERE CurrentStatus = 'CLOSE' AND RecordSubType = 'CLOSE'"
	SQLNumOfServiceCallsClosedRolling5Days = SQLNumOfServiceCallsClosedRolling5Days & " AND RecordCreatedateTime >= '" & lastDateOfPastFiveDays & "' AND RecordCreatedateTime <= '" & firstDateOfPastFiveDays & "' "
	 
	'Response.Write(SQLNumOfServiceCallsClosedRolling5Days)
	 
	Set rsNumOfServiceCallsClosedRolling5Days = Server.CreateObject("ADODB.Recordset")
	rsNumOfServiceCallsClosedRolling5Days.CursorLocation = 3 
	
	rsNumOfServiceCallsClosedRolling5Days.Open SQLNumOfServiceCallsClosedRolling5Days, cnnNumOfServiceCallsClosedRolling5Days
	
	resultNumOfServiceCallsClosedRolling5Days = rsNumOfServiceCallsClosedRolling5Days.RecordCount
	
	rsNumOfServiceCallsClosedRolling5Days.Close
	set rsNumOfServiceCallsClosedRolling5Days= Nothing
	cnnNumOfServiceCallsClosedRolling5Days.Close
	set cnnNumOfServiceCallsClosedRolling5Days= Nothing
	
	GetNumberOfServiceCallsClosedRolling5Days = resultNumOfServiceCallsClosedRolling5Days
	

End Function




Function GetNumberOfServiceCallsFilterChanges()

	Set cnnNumOfServiceCallsFilterChanges = Server.CreateObject("ADODB.Connection")
	cnnNumOfServiceCallsFilterChanges.open Session("ClientCnnString")

	resultNumOfServiceCallsFilterChanges = 0
	
	SQLNumOfServiceCallsFilterChanges = "SELECT COUNT(DISTINCT MemoNumber) AS FilterChangeCount FROM FS_ServiceMemos WHERE CurrentStatus = 'OPEN' AND RecordSubType = 'OPEN' AND FilterChange = 1 "

	'Response.Write(SQLNumOfServiceCallsFilterChanges)
	 
	Set rsNumOfServiceCallsFilterChanges = Server.CreateObject("ADODB.Recordset")
	rsNumOfServiceCallsFilterChanges.CursorLocation = 3 
	
	rsNumOfServiceCallsFilterChanges.Open SQLNumOfServiceCallsFilterChanges , cnnNumOfServiceCallsFilterChanges
	
	If NOT rsNumOfServiceCallsFilterChanges.EOF Then
		resultNumOfServiceCallsFilterChanges = rsNumOfServiceCallsFilterChanges("FilterChangeCount")		
	End If

	rsNumOfServiceCallsFilterChanges.Close
	set rsNumOfServiceCallsFilterChanges= Nothing
	cnnNumOfServiceCallsFilterChanges.Close
	set cnnNumOfServiceCallsFilterChanges= Nothing
	
	GetNumberOfServiceCallsFilterChanges = resultNumOfServiceCallsFilterChanges
	
End Function



Function ServiceCallElapsedMinutesOpenTicket(passedServiceTicketNumber)
	
	totalElapsedMinutes = 0
	
	'Get the normal business day start & end because it is used in a lot of places
	Set cnn10 = Server.CreateObject("ADODB.Connection")
	cnn10.open (Session("ClientCnnString"))
	Set rs10 = Server.CreateObject("ADODB.Recordset")
	rs10.CursorLocation = 3 
	
	SQL10 = "SELECT * FROM Settings_CompanyID"
	Set rs10 = cnn10.Execute(SQL10 )
	
	NormalBizDayStartTime = rs10.fields("BusinessDayStart")
	NormalBizDayEndTime = rs10.fields("BusinessDayEnd")
	MinutesInFullDay = dateDiff("n",NormalBizDayStartTime,NormalBizDayEndTime)
	
	Set rs10 = Nothing
	cnn10.close	
	Set cnn10=Nothing

	OpenedDate = FormatDateTime(GetServiceTicketOpenDateTime(passedServiceTicketNumber),2)
	OpenedTime = FormatDateTime(GetServiceTicketOpenDateTime(passedServiceTicketNumber),4)
	OpenedDateTime = GetServiceTicketOpenDateTime(passedServiceTicketNumber)
	
	ServiceTicketStatus = GetServiceTicketStatus(passedServiceTicketNumber)
	
	If ServiceTicketStatus <> "OPEN" Then 
		ClosedDate =  FormatDateTime(GetServiceTicketCloseDateTime(passedServiceTicketNumber),2)
		ClosedTime = FormatDateTime(GetServiceTicketCloseDateTime(passedServiceTicketNumber),4)
		ClosedDateTime = GetServiceTicketCloseDateTime(passedServiceTicketNumber)
	Else
		ClosedDate =  ""
		ClosedTime = ""
		ClosedDateTime = ""
	End If
	
	'Response.Write("passedServiceTicketNumber : " & passedServiceTicketNumber & "<br>")
	'Response.Write("OpenedDate : " & OpenedDate & "<br>")
	'Response.Write("OpenedTime : " & OpenedTime & "<br>")
	'Response.Write("OpenedDateTime : " & OpenedDateTime & "<br>")
	'Response.Write("ClosedDate : " & ClosedDate & "<br>")
	'Response.Write("ClosedTime : " & ClosedTime & "<br>")
	'Response.Write("ClosedDateTime : " & ClosedDateTime & "<br>")
	'Response.Write("NormalBizDayStartTime: " & NormalBizDayStartTime & "<br>")
	'Response.Write("NormalBizDayEndTime: " & NormalBizDayEndTime & "<br>")	
	'Response.Write("ServiceTicketStatus : " & ServiceTicketStatus & "<br>")
	
	
	'**************************************************************************************************************************
	'IN THIS CONDITION, THE SERVICE TICKET WAS OPENED AND CLOSED ON DIFFERENT DAYS - SPANNING MORE THAN ONE DAY
	'NOW WE NEED TO CALCULATE
	'**************************************************************************************************************************	


	If ClosedDateTime <> "" Then
	
		'******************************************************************************************
		'IF THE TICKET IS CLOSED, CREATE AN ARRAY OF DATES THAT IT WAS OPEN, ENDING AT THE
		'CLOSING DATE
		'******************************************************************************************
		
		If datediff("d",OpenedDateTime,ClosedDateTime) >= 1 Then
		
			NumberOfElements = datediff("d",OpenedDateTime,ClosedDateTime)
			ReDim DaysArray(NumberOfElements)
			
			x = 0
			DateForEval = OpenedDateTime
			DateForEval = cDate(DateForEval)
			
			'Response.Write("xxxxxxxxxxxxx<br>")
			
			Do
				If x = 0 Then
					'******************************************************************************************
					'MAKE SURE WE INSERT THE FIRST DAY, BY NOT ADDING ZERO TO THE DATE
					'******************************************************************************************	
					DaysArray(x) = DateAdd("d",0,DateForEval) 
					DateForEval = DateAdd("d",0,DateForEval) 
				Else
					DaysArray(x) = FormatDateTime(DateAdd("d",1,DateForEval) ,2) & " " & CDate(NormalBizDayStartTime)
					DateForEval = FormatDateTime(DateAdd("d",1,DateForEval) ,2) & " " & CDate(NormalBizDayStartTime)
				End If
				
				'Response.Write("DaysArray(" & x & "): " & DaysArray(x) & "<br>")
				x = x + 1
			
			Loop While x <= NumberOfElements
			
		Else
		
		'***************************************************
		'TICKET WAS OPENED AND CLOSED TODAY
		'***************************************************

			NumberOfElements = 0
			ReDim DaysArray(1)
			DateForEval = OpenedDateTime
			DateForEval = cdate(DateForEval)

			'******************************************************************************************
			'WE JUST NEED TODAY STORED IN THE DATE ARRAY
			'******************************************************************************************					
			DaysArray(0) = DateAdd("d",0,DateForEval) 
			DateForEval = DateAdd("d",0,DateForEval) 
			
			'Response.Write("DaysArray(0): " & DaysArray(0) & "<br>")
			
		End If
		
	Else
	
	
		'******************************************************************************************
		'IF THE TICKET IS NOT CLOSED, CREATE AN ARRAY OF DATES THAT IT WAS OPEN, ENDING AT TODAY
		'AS THE END DATE
		'******************************************************************************************
	
		'***************************************************
		'TICKET OPENED MORE THAN ONE DAY
		'***************************************************
		If datediff("d",OpenedDateTime,Now()) >= 1 Then
		
			NumberOfElements = datediff("d",OpenedDateTime,Now())
			ReDim DaysArray(NumberOfElements)
			
			x = 0
			DateForEval = OpenedDateTime
			DateForEval = cdate(DateForEval)
			
			'Response.Write("xxxxxxxxxxxxx<br>")
			
			Do
				If x = 0 Then
					'******************************************************************************************
					'MAKE SURE WE INSERT THE FIRST DAY, BY NOT ADDING ZERO TO THE DATE
					'******************************************************************************************					
					DaysArray(x) = DateAdd("d",0,DateForEval) 
					DateForEval = DateAdd("d",0,DateForEval) 
				Else
					DaysArray(x) = FormatDAteTime(DateAdd("d",1,DateForEval) ,2) & " " & CDate(NormalBizDayStartTime)
					DateForEval = FormatDAteTime(DateAdd("d",1,DateForEval) ,2) & " " & CDate(NormalBizDayStartTime)
				End If
				
				'Response.Write("DaysArray(" & x & "): " & DaysArray(x) & "<br>")
				x = x + 1
			
			Loop While x <= NumberOfElements
			
		Else
		'***************************************************
		'TICKET WAS OPENED TODAY
		'***************************************************

			NumberOfElements = 0
			ReDim DaysArray(1)
			DateForEval = OpenedDateTime
			DateForEval = cdate(DateForEval)

			'******************************************************************************************
			'WE JUST NEED TODAY STORED IN THE DATE ARRAY
			'******************************************************************************************					
			DaysArray(0) = DateAdd("d",0,DateForEval) 
			DateForEval = DateAdd("d",0,DateForEval) 
			
			'Response.Write("DaysArray(0): " & DaysArray(0) & "<br>")
			
		End If
	
	End If

	
	'******************************************************************************************
	'LOOP THROUGH EACH DATE IN THE ARRAY TO CALCULATE THE HOURS TO ADD TO THE ELAPSED TIME
	'FOR THAT PARTICULAR DAY
	'******************************************************************************************	
	
	For x = 0 To NumberOfElements
	
		'Response.Write("xxxxxxxxxxxxx<br>")
	
		'*******************************************************************************************************
		'FOR EACH DATE, LOOKUP DATE IN COMPANY CALENDER TO SEE IF COMPANY WAS OPEN, CLOSED OR CLOSED EARLY
		'AND CALCULATE THE TOTAL MINUTES OPENED ON THAT DAY
		'*******************************************************************************************************
		'*******************************************************************************************************
		'If the date is the day the ticket is opened, calculate the time since the ticket was opened that day
		'If the date is the day the ticket is closed, calculate the time until the ticket was closed that day
		'Otherwise, the ticket was opened all day, calculate all the minutes the business was opened that day
		'*******************************************************************************************************
		
		'Response.Write("DaysArray(x): " & DaysArray(x) & "<br>")
	
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'~~~~~~~~~~~~~~~~~~~~~~~PROCESS TICKETS THAT ARE STILL OPEN~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		
		'******************************************************************************************
		'IF THE DATE WE ARE PROCESSING IS THE DATE THE TICKET WAS OPENED, GET THE ELAPSED TIME
		'FROM THE TICKET OPEN TIME TO THE CLOSE OF THE BUSINESS DAY IF IT IS AFTER BUSINESS HOURS
		'OTHERWISE, GET THE ELAPSED TIME FROM THE TIME THE TICKET WAS OPENED UNITL NOW
		'******************************************************************************************	
				
		If ClosedDateTime = "" Then
	
			'*****************************************************************************************************
			'TICKET WAS OPENED TODAY AND NOT CLOSED
			'*****************************************************************************************************	
			If datediff("d",DaysArray(x),OpenedDate) = 0 AND datediff("d",OpenedDate,Now()) = 0 Then
			
			
				If DateDiff("n",Time(),NormalBizDayEndTime) <= 0 Then
				
					'*****************************************************************************************************
					'IF IT IS AFTER HOURS, THE ELAPSED TIME WILL BE THE BUSINESS HOURS FOR THE COMPANY ON THAT ENTIRE DAY
					'*****************************************************************************************************	
					
					totalElapsedMinutes = totalElapsedMinutes + NumberofWorkMinutes(DaysArray(x),OpenedTime,NormalBizDayStartTime,NormalBizDayEndTime)
					'Response.Write("DateDiff : " & DateDiff("n",cDate(OpenedTime),cDate(NormalBizDayEndTime)) & "<br>")
					'Response.Write("1a: " & NumberofWorkMinutes(DaysArray(x),OpenedTime,NormalBizDayStartTime,NormalBizDayEndTime) & "<br>")
						
				Else
				
					'*****************************************************************************************************
					'IF IT IS STILL DURING THE BUSINESS DAY, THE ELAPSED TIME WILL BE THE TIME SINCE THE WORK DAY
					'STARTED FOR THE COMPANY, THEREFORE WE NEED TO PASS THE START TIME OF THE BUSINESS DAY AS THE START
					'TIME AND THE END TIME OF THE BUSINESS DAY GETS PASSED AS THE TIME NOW
					'*****************************************************************************************************	
									
					totalElapsedMinutes = totalElapsedMinutes + NumberofWorkMinutes(DaysArray(x),OpenedTime,NormalBizDayStartTime,Time())
					'Response.Write("DateDiff : " & DateDiff("n",OpenedTime,Time()) & "<br>")
					'Response.Write("1b: " & NumberofWorkMinutes(DaysArray(x),OpenedTime,NormalBizDayStartTime,Time()) & "<br>")
									
				End If
		
		
		
			'*****************************************************************************************************
			'TICKET WAS OPENED, BUT NOT TODAY AND IS NOT CLOSED
			'*****************************************************************************************************		
			Else
			
				'******************************************************************************************
				'IF WE ARE PROCESSING THE DATE THAT THE TICKET WAS NOT OPENED AND IS ALSO NOT TODAY,
				'THE ELAPSED TIME WILL BE THE BUSINESS HOURS FOR THE COMPANY ON THAT ENTIRE DAY,
				'AS CHECKED AGAINST THE COMPANY CALENDER
				'******************************************************************************************		
				
				'*****************************************************************************
				'THE DATE WE ARE PROCESSING IS NOT THE SAME AS THE DAY THE TICKET WAS OPENED
				'*****************************************************************************	
				If datediff("d",DaysArray(x),OpenedDate) <> 0 Then
				
					'********************************************
					'THE DATE WE ARE PROCESSING IS TODAY
					'********************************************	
					If datediff("d",DaysArray(x),Now()) = 0 Then
					
						
						'*****************************************************************************************************
						'IF IT IS AFTER HOURS, THE ELAPSED TIME WILL BE THE BUSINESS HOURS FOR THE COMPANY ON THAT ENTIRE DAY
						'*****************************************************************************************************	
				
						If DateDiff("n",Time(),NormalBizDayEndTime) <= 0 Then
							
							totalElapsedMinutes = totalElapsedMinutes + NumberofWorkMinutes(DaysArray(x),NormalBizDayStartTime,NormalBizDayStartTime,NormalBizDayEndTime)
							
							'Response.Write("totalElapsedMinutes : " & totalElapsedMinutes & "<br>")
							'Response.Write("DateDiff : " & DateDiff("n",NormalBizDayStartTime,NormalBizDayEndTime) & "<br>")
							'Response.Write("2a: " & NumberofWorkMinutes(DaysArray(x),NormalBizDayStartTime,NormalBizDayStartTime,NormalBizDayEndTime) & "<br>")
						
					
						'*****************************************************************************************************
						'IF IT IS STILL DURING THE BUSINESS DAY, THE ELAPSED TIME WILL BE THE TIME SINCE THE WORK DAY
						'STARTED FOR THE COMPANY, THEREFORE WE NEED TO PASS THE START TIME OF THE BUSINESS DAY AS THE START
						'TIME AND THE END TIME OF THE BUSINESS DAY GETS PASSED AS THE TIME NOW
						'*****************************************************************************************************	

						Else
							
							totalElapsedMinutes = totalElapsedMinutes + NumberofWorkMinutes(DaysArray(x),NormalBizDayStartTime,NormalBizDayStartTime,Time())
							'Response.Write("DateDiff : " & DateDiff("n",NormalBizDayStartTime,Time()) & "<br>")
							'Response.Write("2b: " & NumberofWorkMinutes(DaysArray(x),NormalBizDayStartTime,NormalBizDayStartTime,Time()) & "<br>")
											
						End If
	
					'********************************************
					'THE DATE WE ARE PROCESSING IS NOT TODAY
					'********************************************	
					Else
					
						'*****************************************************************************************************
						'THE ELAPSED TIME WILL BE THE BUSINESS HOURS FOR THE COMPANY ON THAT ENTIRE DAY
						'*****************************************************************************************************	

						totalElapsedMinutes = totalElapsedMinutes + NumberofWorkMinutes(DaysArray(x),NormalBizDayStartTime,NormalBizDayStartTime,NormalBizDayEndTime)
						
						'Response.Write("totalElapsedMinutes : " & totalElapsedMinutes & "<br>")
						'Response.Write("DateDiff : " & DateDiff("n",NormalBizDayStartTime,NormalBizDayEndTime) & "<br>")
						'Response.Write("2aa: " & NumberofWorkMinutes(DaysArray(x),NormalBizDayStartTime,NormalBizDayStartTime,NormalBizDayEndTime) & "<br>")
						
					End If
					
				'***************************************************************************************
				'THE DATE WE ARE PROCESSING IS THE SAME AS THE DAY THE TICKET WAS OPENED, BUT NOT TODAY
				'***************************************************************************************	
				Else	
								
			
					'*****************************************************************************************************
					'IF IT IS AFTER HOURS, THE ELAPSED TIME IS ZERO BECAUSE THE TICKET WAS OPENED AFTER BUSINESS CLOSE
					'*****************************************************************************************************	
			
					If DateDiff("n",OpenedTime,NormalBizDayEndTime) <= 0 Then
					
						totalElapsedMinutes = totalElapsedMinutes + 0
						'Response.Write("2c: ZERO<br>")
					
					Else
				
						'*****************************************************************************************************
						'THE ELAPSED TIME WILL BE FROM THE TIME THE TICKET WAS OPENED UNTIL THE END OF THE DAY
						'*****************************************************************************************************	
						
						totalElapsedMinutes = totalElapsedMinutes + NumberofWorkMinutes(DaysArray(x),OpenedTime,OpenedTime,NormalBizDayEndTime)
						
						'Response.Write("totalElapsedMinutes : " & totalElapsedMinutes & "<br>")
						'Response.Write("DateDiff : " & DateDiff("n",OpenedTime,NormalBizDayEndTime) & "<br>")
						'Response.Write("2d: " & NumberofWorkMinutes(DaysArray(x),OpenedTime,OpenedTime,NormalBizDayEndTime) & "<br>")
					
					End If
					
				End If
				
			
			
			
			End If
		
		End If
		
		
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'~~~~~~~~~~~~~~~~~~~~~~~PROCESS TICKETS THAT HAVE BEEN CLOSED~~~~~~~~~~~~~~~~~~~~~~~~~~
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		
		'******************************************************************************************
		'IF THE DATE WE ARE PROCESSING IS THE DATE THE TICKET WAS CLOSED, GET THE ELAPSED TIME
		'FROM THE TICKET OPEN TIME TO THE CLOSE OF THE BUSINESS DAY
		'******************************************************************************************	
		
		If ClosedDateTime <> "" Then
	
			'******************************************************************************************
			'IF WE ARE PROCESSING THE DATE THE TICKET WAS CLOSED, GET THE ELAPSED TIME
			'FROM THE START OF THE BUSINESS DAY UNTIL THE TIME THE TICKET WAS CLOSED
			'******************************************************************************************	
			If datediff("d",DaysArray(x),ClosedDate) = 0 Then
				totalElapsedMinutes = totalElapsedMinutes + NumberofWorkMinutes(DaysArray(x),ClosedTime,NormalBizDayStartTime,NormalBizDayEndTime)
				'Response.Write("DateDiff : " & DateDiff("n",cDate(ClosedTime),cDate(NormalBizDayEndTime)) & "<br>")
				'Response.Write("3a: " & NumberofWorkMinutes(DaysArray(x),ClosedTime,NormalBizDayStartTime,NormalBizDayEndTime) & "<br>")
			End If
		
			'******************************************************************************************
			'IF WE ARE PROCESSING THE DATE THAT THE TICKET WAS NOT OPENED AND NOT CLOSED,
			'THE ELAPSED TIME WILL BE THE BUSINESS HOURS FOR THE COMPANY ON THAT ENTIRE DAY,
			'AS CHECKED AGAINST THE COMPANY CALENDER
			'******************************************************************************************			
			If datediff("d",DaysArray(x),OpenedDate) <> 0 AND datediff("d",DaysArray(x),ClosedDate) <> 0 Then
				totalElapsedMinutes = totalElapsedMinutes + NumberofWorkMinutes(DaysArray(x),NormalBizDayStartTime,NormalBizDayStartTime,NormalBizDayEndTime)
				'Response.Write("DateDiff : " & DateDiff("n",cDate(NormalBizDayStartTime),cDate(NormalBizDayEndTime)) & "<br>")
				'Response.Write("3b: " & NumberofWorkMinutes(DaysArray(x),NormalBizDayStartTime,NormalBizDayStartTime,NormalBizDayEndTime) & "<br>")
			End If
					
		End If		

		'Response.Write("xxxxxxxxxxxxx<br>")

		'Response.Write("LoopVar:" & x &"<br>")
		
	Next


	ServiceCallElapsedMinutesOpenTicket = totalElapsedMinutes

End Function







Function ServiceCallElapsedMinutesClosedTicket(passedServiceTicketNumber)
	
	totalElapsedMinutes = 0	
	
	'Get the normal business day start & end because it is used in a lot of places
	Set cnn10 = Server.CreateObject("ADODB.Connection")
	cnn10.open (Session("ClientCnnString"))
	Set rs10 = Server.CreateObject("ADODB.Recordset")
	rs10.CursorLocation = 3 
	
	SQL10 = "SELECT * FROM Settings_CompanyID"
	Set rs10 = cnn10.Execute(SQL10 )
	
	NormalBizDayStartTime = rs10.fields("BusinessDayStart")
	NormalBizDayEndTime = rs10.fields("BusinessDayEnd")
	MinutesInFullDay = dateDiff("n",NormalBizDayStartTime,NormalBizDayEndTime)
	
	Set rs10 = Nothing
	cnn10.close	
	Set cnn10=Nothing

	OpenedDate = FormatDateTime(GetServiceTicketOpenDateTime(passedServiceTicketNumber),2)
	OpenedTime = FormatDateTime(GetServiceTicketOpenDateTime(passedServiceTicketNumber),4)
	OpenedDateTime = GetServiceTicketOpenDateTime(passedServiceTicketNumber)
	
	ServiceTicketStatus = GetServiceTicketStatus(passedServiceTicketNumber)
	
	ClosedDate =  FormatDateTime(GetServiceTicketCloseDateTime(passedServiceTicketNumber),2)
	ClosedTime = FormatDateTime(GetServiceTicketCloseDateTime(passedServiceTicketNumber),4)
	ClosedDateTime = GetServiceTicketCloseDateTime(passedServiceTicketNumber)
	
	'Response.Write("passedServiceTicketNumber : " & passedServiceTicketNumber & "<br>")
	'Response.Write("OpenedDate : " & OpenedDate & "<br>")
	'Response.Write("OpenedTime : " & OpenedTime & "<br>")
	'Response.Write("OpenedDateTime : " & OpenedDateTime & "<br>")
	'Response.Write("ClosedDate : " & ClosedDate & "<br>")
	'Response.Write("ClosedTime : " & ClosedTime & "<br>")
	'Response.Write("ClosedDateTime : " & ClosedDateTime & "<br>")
	'Response.Write("NormalBizDayStartTime: " & NormalBizDayStartTime & "<br>")
	'Response.Write("NormalBizDayEndTime: " & NormalBizDayEndTime & "<br>")	
	'Response.Write("ServiceTicketStatus : " & ServiceTicketStatus & "<br>")
	
	
	'**************************************************************************************************************************
	'TICKET WAS OPENED AND CLOSED ON DIFFERENT DATES
	'**************************************************************************************************************************	

	If datediff("d",OpenedDateTime,ClosedDateTime) >= 1 Then
	
		NumberOfElements = datediff("d",OpenedDateTime,ClosedDateTime)
		ReDim DaysArray(NumberOfElements)
		
		x = 0
		DateForEval = OpenedDateTime
		DateForEval = cDate(DateForEval)
		
		'Response.Write("xxxxxxxxxxxxx<br>")
		
		Do
			If x = 0 Then
				'******************************************************************************************
				'MAKE SURE WE INSERT THE FIRST DAY, BY NOT ADDING ZERO TO THE DATE
				'******************************************************************************************	
				DaysArray(x) = DateAdd("d",0,DateForEval) 
				DateForEval = DateAdd("d",0,DateForEval) 
			Else
				DaysArray(x) = FormatDateTime(DateAdd("d",1,DateForEval) ,2) & " " & CDate(NormalBizDayStartTime)
				DateForEval = FormatDateTime(DateAdd("d",1,DateForEval) ,2) & " " & CDate(NormalBizDayStartTime)
			End If
			
			'Response.Write("DaysArray(" & x & "): " & DaysArray(x) & "<br>")
			x = x + 1
		
		Loop While x <= NumberOfElements
		
	Else
		
		'***************************************************
		'TICKET WAS OPENED AND CLOSED TODAY
		'***************************************************

		NumberOfElements = 0
		ReDim DaysArray(1)
		DateForEval = OpenedDateTime
		DateForEval = cdate(DateForEval)

		'******************************************************************************************
		'WE JUST NEED TODAY STORED IN THE DATE ARRAY
		'******************************************************************************************					
		DaysArray(0) = DateAdd("d",0,DateForEval) 
		DateForEval = DateAdd("d",0,DateForEval) 
		
		'Response.Write("DaysArray(0): " & DaysArray(0) & "<br>")
		
	End If
	
	
	'******************************************************************************************
	'LOOP THROUGH EACH DATE IN THE ARRAY TO CALCULATE THE HOURS TO ADD TO THE ELAPSED TIME
	'FOR THAT PARTICULAR DAY
	'******************************************************************************************	
	
	For x = 0 To NumberOfElements
	
		'Response.Write("xxxxxxxxxxxxx<br>")
	
		'*******************************************************************************************************
		'FOR EACH DATE, LOOKUP DATE IN COMPANY CALENDER TO SEE IF COMPANY WAS OPEN, CLOSED OR CLOSED EARLY
		'AND CALCULATE THE TOTAL MINUTES OPENED ON THAT DAY
		'*******************************************************************************************************
		'*******************************************************************************************************
		'If the date is the day the ticket is opened, calculate the time since the ticket was opened that day
		'If the date is the day the ticket is closed, calculate the time until the ticket was closed that day
		'Otherwise, the ticket was opened all day, calculate all the minutes the business was opened that day
		'*******************************************************************************************************
		
		'Response.Write("DaysArray(x): " & DaysArray(x) & "<br>")
		
					
		'******************************************************************************************
		'IF WE ARE PROCESSING THE DATE THE TICKET WAS CLOSED, GET THE ELAPSED TIME
		'FROM THE START OF THE BUSINESS DAY UNTIL THE TIME THE TICKET WAS CLOSED
		'******************************************************************************************	
		If datediff("d",DaysArray(x),ClosedDate) = 0 Then
		
			'******************************************************************************************
			'CLOSED BEFORE THE BUSINESS DAY STARTED, NO ELAPSED TIME IS ADDED
			'******************************************************************************************
			If DateDiff("n",ClosedTime,NormalBizDayStartTime) > 0 Then
		
				totalElapsedMinutes = totalElapsedMinutes + 0
				'Response.Write("3a: ZERO <br>")

			Else
			
				If DateDiff("n",ClosedTime,NormalBizDayEndTime) < 0 Then
				
					'******************************************************************************************
					'CLOSED BEFORE THE BUSINESS DAY ENDED
					'******************************************************************************************
				
					totalElapsedMinutes = totalElapsedMinutes + NumberofWorkMinutes(DaysArray(x),NormalBizDayStartTime,NormalBizDayStartTime,NormalBizDayEndTime)
					'Response.Write("DateDiff : " & DateDiff("n",cDate(NormalBizDayStartTime),cDate(NormalBizDayEndTime)) & "<br>")
					'Response.Write("3b: " & NumberofWorkMinutes(DaysArray(x),NormalBizDayStartTime,NormalBizDayStartTime,NormalBizDayEndTime) & "<br>")
	
				Else
				
					If datediff("d",DaysArray(x),OpenedDate) = 0 Then
						'******************************************************************************************
						'IF WE ARE PROCESSING THE DATE THE TICKET WAS OPENED, GET THE ELAPSED TIME
						'FROM THE TIME THE TICKET WAS OPENED UNTIL THE END OF THE BUSINESS DAY
						'******************************************************************************************
						totalElapsedMinutes = totalElapsedMinutes + NumberofWorkMinutes(DaysArray(x),OpenedTime,NormalBizDayStartTime,NormalBizDayEndTime)
						'Response.Write("DateDiff : " & DateDiff("n",cDate(OpenedTime),cDate(NormalBizDayEndTime)) & "<br>")
						'Response.Write("3c: " & NumberofWorkMinutes(DaysArray(x),OpenedTime,NormalBizDayStartTime,NormalBizDayEndTime) & "<br>")
					Else
						'******************************************************************************************
						'TICKET WAS NOT OPENED TODAY, BUT TICKET WAS CLOSED TODAY
						'GET THE TIME FROM THE START OF THE BUSINESS DAY UNTIL TICKET CLOSE TIME
						'******************************************************************************************
						totalElapsedMinutes = totalElapsedMinutes + NumberofWorkMinutes(DaysArray(x),NormalBizDayStartTime,NormalBizDayStartTime,ClosedTime)
						'Response.Write("DateDiff : " & DateDiff("n",cDate(ClosedTime),cDate(ClosedTime)) & "<br>")
						'Response.Write("3cc: " & NumberofWorkMinutes(DaysArray(x),NormalBizDayStartTime,NormalBizDayStartTime,ClosedTime) & "<br>")
					End If
					
				End If
			
			End If
		End If
	
		'******************************************************************************************
		'IF WE ARE PROCESSING THE DATE THAT THE TICKET WAS NOT OPENED AND NOT CLOSED,
		'THE ELAPSED TIME WILL BE THE BUSINESS HOURS FOR THE COMPANY ON THAT ENTIRE DAY,
		'AS CHECKED AGAINST THE COMPANY CALENDER
		'******************************************************************************************			
		If datediff("d",DaysArray(x),OpenedDate) <> 0 AND datediff("d",DaysArray(x),ClosedDate) <> 0 Then
			totalElapsedMinutes = totalElapsedMinutes + NumberofWorkMinutes(DaysArray(x),NormalBizDayStartTime,NormalBizDayStartTime,NormalBizDayEndTime)
			'Response.Write("DateDiff : " & DateDiff("n",cDate(NormalBizDayStartTime),cDate(NormalBizDayEndTime)) & "<br>")
			'Response.Write("3d: " & NumberofWorkMinutes(DaysArray(x),NormalBizDayStartTime,NormalBizDayStartTime,NormalBizDayEndTime) & "<br>")
		End If

		If datediff("d",DaysArray(x),OpenedDate) = 0 AND datediff("d",DaysArray(x),ClosedDate) <> 0 Then
			totalElapsedMinutes = totalElapsedMinutes + NumberofWorkMinutes(DaysArray(x),OpenedTime,NormalBizDayStartTime,NormalBizDayEndTime)
			'Response.Write("DateDiff : " & DateDiff("n",cDate(OpenedTime),cDate(NormalBizDayEndTime)) & "<br>")
			'Response.Write("3e: " & NumberofWorkMinutes(DaysArray(x),OpenedTime,NormalBizDayStartTime,NormalBizDayEndTime) & "<br>")
		End If

		'Response.Write("xxxxxxxxxxxxx<br>")
		'Response.Write("LoopVar:" & x &"<br>")
		
	Next

	ServiceCallElapsedMinutesClosedTicket = totalElapsedMinutes 

End Function


Function CustHasServiceTicketNotes(passedServiceTicketNumber)

	Set cnnCustHasServiceTicketNotes = Server.CreateObject("ADODB.Connection")
	cnnCustHasServiceTicketNotes.open Session("ClientCnnString")

	resultCustHasServiceTicketNotes = False
		
	SQLCustHasServiceTicketNotes = "SELECT TOP 1 * FROM FS_ServiceMemosNotes "
	SQLCustHasServiceTicketNotes = SQLCustHasServiceTicketNotes & "WHERE ServiceTicketID = '" & passedServiceTicketNumber & "' "
	 
	Set rsCustHasServiceTicketNotes = Server.CreateObject("ADODB.Recordset")
	rsCustHasServiceTicketNotes.CursorLocation = 3 
	Set rsCustHasServiceTicketNotes= cnnCustHasServiceTicketNotes.Execute(SQLCustHasServiceTicketNotes)
	
	If not rsCustHasServiceTicketNotes.eof then resultCustHasServiceTicketNotes =  True
		
	rsCustHasServiceTicketNotes.Close
	set rsCustHasServiceTicketNotes= Nothing
	cnnCustHasServiceTicketNotes.Close	
	set cnnCustHasServiceTicketNotes = Nothing
	
	CustHasServiceTicketNotes = resultCustHasServiceTicketNotes 
	
End Function

Function NoteNewServiceTicketForUser(passedServiceTicketNumber)

	resultNoteNewServiceTicketForUser = False
	
	SQLNoteNewServiceTicketForUser = "SELECT * FROM FS_ServiceMemosNotesUserViewed Where ServiceTicketID ='" & passedServiceTicketNumber & "' AND UserNo = " & Session("Userno")	
	Set cnnNoteNewServiceTicketForUser = Server.CreateObject("ADODB.Connection")
	cnnNoteNewServiceTicketForUser.open (Session("ClientCnnString"))
	Set rsNoteNewServiceTicketForUser = Server.CreateObject("ADODB.Recordset")
	rsNoteNewServiceTicketForUser.CursorLocation = 3 
	Set rsNoteNewServiceTicketForUser = cnnNoteNewServiceTicketForUser.Execute(SQLNoteNewServiceTicketForUser)

	Set rsNoteCatAnal = Server.CreateObject("ADODB.Recordset")
	rsNoteCatAnal.CursorLocation = 3 

	If not rsNoteNewServiceTicketForUser.EOF Then
		'OK, so see when the last note was created, not by us
		SQLCustHasServiceTicketNotes = "SELECT TOP 1 RecordCreationDateTime FROM FS_ServiceMemosNotes "
		SQLCustHasServiceTicketNotes = SQLCustHasServiceTicketNotes & " WHERE ServiceTicketID = '" & passedServiceTicketNumber & "' "
		SQLCustHasServiceTicketNotes = SQLCustHasServiceTicketNotes & " ORDER BY RecordCreationDateTime DESC"
		
		Set rsNoteCatAnal = cnnNoteNewServiceTicketForUser.Execute(SQLCustHasServiceTicketNotes)
		If Not rsNoteCatAnal.Eof Then
			If rsNoteNewServiceTicketForUser("DateLastViewed") < rsNoteCatAnal("RecordCreationDateTime")  Then resultNoteNewServiceTicketForUser = True
		End If
	Else
		resultNoteNewServiceTicketForUser = True 'Also true if they have never seen any of them
	End If
	cnnNoteNewServiceTicketForUser.close
	set rsNoteNewServiceTicketForUser = nothing
	set rsNoteCatAnal = nothing
	set cnnNoteNewServiceTicketForUser= nothing	

	NoteNewServiceTicketForUser = resultNoteNewServiceTicketForUser

End Function

Sub MarkNoteNewForUserServiceTicket(passedServiceTicketNumber)

	SQLMarkNoteNewForUserServiceTicket = "SELECT * FROM FS_ServiceMemosNotesUserViewed Where ServiceTicketID ='" & passedServiceTicketNumber & "' AND UserNo = " & Session("Userno")	
	Set cnnMarkNoteNewForUserServiceTicket = Server.CreateObject("ADODB.Connection")
	cnnMarkNoteNewForUserServiceTicket.open (Session("ClientCnnString"))
	Set rMarkNoteNewForUserServiceTicket = Server.CreateObject("ADODB.Recordset")
	rMarkNoteNewForUserServiceTicket.CursorLocation = 3 
	Set rMarkNoteNewForUserServiceTicket = cnnMarkNoteNewForUserServiceTicket.Execute(SQLMarkNoteNewForUserServiceTicket)

	If rMarkNoteNewForUserServiceTicket.EOF Then ' Nothing there so we need to insert
		SQLMarkNoteNewForUserServiceTicket = "INSERT INTO FS_ServiceMemosNotesUserViewed (ServiceTicketID,UserNo) VALUES ('" & passedServiceTicketNumber & "',"  & Session("UserNo") & ")"
	Else
		SQLMarkNoteNewForUserServiceTicket = "UPDATE FS_ServiceMemosNotesUserViewed Set DateLastViewed = getdate() Where ServiceTicketID ='" & passedServiceTicketNumber & "' AND UserNo = " & Session("Userno")
	End If
	
	Set rMarkNoteNewForUserServiceTicket = cnnMarkNoteNewForUserServiceTicket.Execute(SQLMarkNoteNewForUserServiceTicket)
		
	cnnMarkNoteNewForUserServiceTicket.close
	set rMarkNoteNewForUserServiceTicket = nothing
	set cnnMarkNoteNewForUserServiceTicket= nothing	

End Sub

Function GetLastServiceTicketNotesByTicket(passedServiceTicketNumber)

	Set cnnGetLastServiceTicketNotesByTicket = Server.CreateObject("ADODB.Connection")
	cnnGetLastServiceTicketNotesByTicket.open Session("ClientCnnString")

	resultGetLastServiceTicketNotesByTicket = ""
		
	SQLGetLastServiceTicketNotesByTicket = "SELECT TOP 1 * FROM FS_ServiceMemosNotes "
	SQLGetLastServiceTicketNotesByTicket = SQLGetLastServiceTicketNotesByTicket & "WHERE ServiceTicketID = '" & passedServiceTicketNumber & "' ORDER BY InternalRecordIdentifier DESC"
	 
'Response.Write(SQLGetLastServiceTicketNotesByTicket)
	
	Set rsGetLastServiceTicketNotesByTicket = Server.CreateObject("ADODB.Recordset")
	rsGetLastServiceTicketNotesByTicket.CursorLocation = 3 
	Set rsGetLastServiceTicketNotesByTicket= cnnGetLastServiceTicketNotesByTicket.Execute(SQLGetLastServiceTicketNotesByTicket)
	
	If not rsGetLastServiceTicketNotesByTicket.eof then resultGetLastServiceTicketNotesByTicket =  rsGetLastServiceTicketNotesByTicket("Note")
		
	rsGetLastServiceTicketNotesByTicket.Close
	set rsGetLastServiceTicketNotesByTicket= Nothing
	cnnGetLastServiceTicketNotesByTicket.Close	
	set cnnGetLastServiceTicketNotesByTicket = Nothing
	
	GetLastServiceTicketNotesByTicket = resultGetLastServiceTicketNotesByTicket 
	
End Function



Function GetNumberOfMinutesInServiceDay()

	Set cnnNumberOfMinutesInServiceDay = Server.CreateObject("ADODB.Connection")
	cnnNumberOfMinutesInServiceDay.open Session("ClientCnnString")

	resultNumberOfMinutesInServiceDay = 0
	
    SQLNumberOfMinutesInServiceDay = "SELECT ServiceDayStartTime, ServiceDayEndTime, ServiceDayElapsedTimeCalculationMethod FROM Settings_FieldService"    

	'Response.Write(SQLNumberOfMinutesInServiceDay)
	 
	Set rsNumberOfMinutesInServiceDay = Server.CreateObject("ADODB.Recordset")
	rsNumberOfMinutesInServiceDay.CursorLocation = 3 
	
	rsNumberOfMinutesInServiceDay.Open SQLNumberOfMinutesInServiceDay, cnnNumberOfMinutesInServiceDay
	
	If NOT rsNumberOfMinutesInServiceDay.EOF Then
	
		ServiceDayStartTime = rsNumberOfMinutesInServiceDay.Fields("ServiceDayStartTime")
		ServiceDayEndTime = rsNumberOfMinutesInServiceDay.Fields("ServiceDayEndTime")
		ServiceDayElapsedTimeCalculationMethod = rsNumberOfMinutesInServiceDay.Fields("ServiceDayElapsedTimeCalculationMethod")
		
		If ServiceDayStartTime = "" OR ServiceDayStartTime = "" Then
			resultNumberOfMinutesInServiceDay = 1440
		Else
			If ServiceDayElapsedTimeCalculationMethod = "Actual" Then
				resultNumberOfMinutesInServiceDay = 1440
			Else
				resultNumberOfMinutesInServiceDay = DateDiff("n", ServiceDayStartTime, ServiceDayEndTime)
			End If
		End If
		
	End If
		
	rsNumberOfMinutesInServiceDay.Close
	set rsNumberOfMinutesInServiceDay= Nothing
	cnnNumberOfMinutesInServiceDay.Close
	set cnnNumberOfMinutesInServiceDay= Nothing
	
	GetNumberOfMinutesInServiceDay = resultNumberOfMinutesInServiceDay
	
End Function

Function FChange_NextDate(passedCustid)

	resultFChange_NextDate = ""

	'Remember, it reads the setting FieldServiceDays from tblSetting_Global to determine how many days to use in the evaluation
	Set cnnFChange_NextDate = Server.CreateObject("ADODB.Connection")
	Set rsFChange_NextDate = Server.CreateObject("ADODB.Recordset")
	
	cnnFChange_NextDate.open (Session("ClientCnnString"))
	
	SQLFChange_NextDate = "SELECT * FROM Settings_Global"
		
	rsFChange_NextDate.CursorLocation = 3 
	Set rsFChange_NextDate = cnnFChange_NextDate.Execute(SQLFChange_NextDate)
	
	If not rsFChange_NextDate.EOF Then
		FChangeDays = rsFChange_NextDate("FilterChangeDays")
	Else
		FChangeDays = 15
	End If
	
	SQLFChange_NextDate = "SELECT TOP 1 *, "
	SQLFChange_NextDate = SQLFChange_NextDate & "CASE FrequencyType WHEN 'D' THEN dateadd(d,FrequencyTime, LastChangeDateTime) "
	SQLFChange_NextDate = SQLFChange_NextDate & "WHEN 'W' THEN dateadd(d, (FrequencyTime * 7), LastChangeDateTime) "
	SQLFChange_NextDate = SQLFChange_NextDate & "WHEN 'M' THEN dateadd(d, (FrequencyTime * 28), LastChangeDateTime) END AS NextChangeDate "
	SQLFChange_NextDate = SQLFChange_NextDate & "FROM FS_CustomerFilters WHERE CustID = '" & passedCustid & "' "
	
'	SQLFChange_NextDate = SQLFChange_NextDate & "AND ("
'	SQLFChange_NextDate = SQLFChange_NextDate & " FS_CustomerFilters.FilterIntRecID NOT IN "
'	SQLFChange_NextDate = SQLFChange_NextDate & "(SELECT FilterIntRecID FROM FS_ServiceMemosFilterInfo WHERE CustID = '" & passedCustid & "' AND "
'	SQLFChange_NextDate = SQLFChange_NextDate & " ServiceTicketID IN (Select MemoNumber FROM FS_ServiceMemos WHERE CurrentStatus='OPEN')) "
'	SQLFChange_NextDate = SQLFChange_NextDate & ")"
	
	SQLFChange_NextDate = SQLFChange_NextDate & "ORDER BY NextChangeDate "
	
	'Response.Write(SQLFChange_NextDate)
	Set rsFChange_NextDate = cnnFChange_NextDate.Execute(SQLFChange_NextDate)
	
	If not rsFChange_NextDate.eof then 	resultFChange_NextDate = cdate(rsFChange_NextDate("NextChangeDate"))
	
	set rsFChange_NextDate = Nothing
	cnnFChange_NextDate.Close
	set cnnFChange_NextDate = Nothing
	
	FChange_NextDate = resultFChange_NextDate

End Function

Function NumberOfTicketsByProblemCode (passedproblemCode)

	resultNumberOfTicketsByProblemCode = ""
	
	Set cnnNumberOfTicketsByProblemCode = Server.CreateObject("ADODB.Connection")
	cnnNumberOfTicketsByProblemCode.open Session("ClientCnnString")

		
	SQLNumberOfTicketsByProblemCode = "SELECT Count(*) AS ProbCount FROM FS_ServiceMemos WHERE ProblemCode = " & passedproblemCode
	 
	Set rsNumberOfTicketsByProblemCode = Server.CreateObject("ADODB.Recordset")
	rsNumberOfTicketsByProblemCode.CursorLocation = 3 

	Set rsNumberOfTicketsByProblemCode = cnnNumberOfTicketsByProblemCode.Execute(SQLNumberOfTicketsByProblemCode)

	If NOT rsNumberOfTicketsByProblemCode.EOF Then
		resultNumberOfTicketsByProblemCode = rsNumberOfTicketsByProblemCode("probCount")
	End If
	
	set rsNumberOfTicketsByProblemCode= Nothing
	cnnNumberOfTicketsByProblemCode.Close	
	set cnnNumberOfTicketsByProblemCode= Nothing
	
	NumberOfTicketsByProblemCode = resultNumberOfTicketsByProblemCode
	
End Function


Function NumberOfTicketsBySymptomCode (passedSymptomCode)

	resultNumberOfTicketsBySymptomCode = ""
	
	Set cnnNumberOfTicketsBySymptomCode = Server.CreateObject("ADODB.Connection")
	cnnNumberOfTicketsBySymptomCode.open Session("ClientCnnString")

		
	SQLNumberOfTicketsBySymptomCode = "SELECT Count(*) AS SymptomCount FROM FS_ServiceMemos WHERE SymptomCode = " & passedSymptomCode
	 
	Set rsNumberOfTicketsBySymptomCode = Server.CreateObject("ADODB.Recordset")
	rsNumberOfTicketsBySymptomCode.CursorLocation = 3 

	Set rsNumberOfTicketsBySymptomCode = cnnNumberOfTicketsBySymptomCode.Execute(SQLNumberOfTicketsBySymptomCode)

	If NOT rsNumberOfTicketsBySymptomCode.EOF Then
		resultNumberOfTicketsBySymptomCode = rsNumberOfTicketsBySymptomCode("SymptomCount")
	End If
	
	set rsNumberOfTicketsBySymptomCode= Nothing
	cnnNumberOfTicketsBySymptomCode.Close	
	set cnnNumberOfTicketsBySymptomCode= Nothing
	
	NumberOfTicketsBySymptomCode = resultNumberOfTicketsBySymptomCode
	
End Function


Function NumberOfTicketsByResolutionCode (passedResolutionCode)

	resultNumberOfTicketsByResolutionCode = ""
	
	Set cnnNumberOfTicketsByResolutionCode = Server.CreateObject("ADODB.Connection")
	cnnNumberOfTicketsByResolutionCode.open Session("ClientCnnString")

		
	SQLNumberOfTicketsByResolutionCode = "SELECT Count(*) AS ResolutionCount FROM FS_ServiceMemosDetail WHERE ResolutionCode = " & passedResolutionCode
	 
	Set rsNumberOfTicketsByResolutionCode = Server.CreateObject("ADODB.Recordset")
	rsNumberOfTicketsByResolutionCode.CursorLocation = 3 

	Set rsNumberOfTicketsByResolutionCode = cnnNumberOfTicketsByResolutionCode.Execute(SQLNumberOfTicketsByResolutionCode)

	If NOT rsNumberOfTicketsByResolutionCode.EOF Then
		resultNumberOfTicketsByResolutionCode = rsNumberOfTicketsByResolutionCode("ResolutionCount")
	End If
	
	set rsNumberOfTicketsByResolutionCode= Nothing
	cnnNumberOfTicketsByResolutionCode.Close	
	set cnnNumberOfTicketsByResolutionCode= Nothing
	
	NumberOfTicketsByResolutionCode = resultNumberOfTicketsByResolutionCode
	
End Function



'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function NumberCustomerRecsDefinedForFilterID(passedFilterIntRecID)

	Set cnnNumCustDefinedForFilterID = Server.CreateObject("ADODB.Connection")
	cnnNumCustDefinedForFilterID.open Session("ClientCnnString")

	resultNumCustDefinedForFilterID = 0
	
	SQLNumCustDefinedForFilterID = "SELECT COUNT(*) AS FILTERCOUNT FROM FS_CustomerFilters INNER JOIN IC_Filters ON "
	SQLNumCustDefinedForFilterID = SQLNumCustDefinedForFilterID & "IC_Filters.InternalRecordIdentifier = FS_CustomerFilters.FilterIntRecID "
	SQLNumCustDefinedForFilterID = SQLNumCustDefinedForFilterID & " WHERE IC_Filters.InternalRecordIdentifier = " & passedFilterIntRecID

	 
	Set rsNumCustDefinedForFilterID = Server.CreateObject("ADODB.Recordset")
	rsNumCustDefinedForFilterID.CursorLocation = 3 
	
	rsNumCustDefinedForFilterID.Open SQLNumCustDefinedForFilterID,cnnNumCustDefinedForFilterID 
			
	resultNumCustDefinedForFilterID = rsNumCustDefinedForFilterID("FILTERCOUNT")
	
	rsNumCustDefinedForFilterID.Close
	set rsNumCustDefinedForFilterID = Nothing
	cnnNumCustDefinedForFilterID.Close	
	set cnnNumCustDefinedForFilterID = Nothing
	
	NumberCustomerRecsDefinedForFilterID = resultNumCustDefinedForFilterID
	
End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************



'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function GetFilterDescByIntRecID(passedMovementCodeIntRecID)

	Set cnnGetFilterDescByIntRecID = Server.CreateObject("ADODB.Connection")
	cnnGetFilterDescByIntRecID.open Session("ClientCnnString")

	resultGetFilterDescByIntRecID = ""
		
	SQLGetFilterDescByIntRecID = "SELECT * FROM IC_Filters WHERE InternalRecordIdentifier = " & passedMovementCodeIntRecID
	 
	Set rsGetFilterDescByIntRecID  = Server.CreateObject("ADODB.Recordset")
	rsGetFilterDescByIntRecID.CursorLocation = 3 
	
	rsGetFilterDescByIntRecID.Open SQLGetFilterDescByIntRecID,cnnGetFilterDescByIntRecID 
			
	resultGetFilterDescByIntRecID = rsGetFilterDescByIntRecID("Description")
	
	rsGetFilterDescByIntRecID.Close
	set rsGetFilterDescByIntRecID = Nothing
	cnnGetFilterDescByIntRecID.Close	
	set cnnGetFilterDescByIntRecID = Nothing
	
	GetFilterDescByIntRecID  = resultGetFilterDescByIntRecID 
	
End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************




'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function GetFilterIDByIntRecID(passedMovementCodeIntRecID)

	Set cnnGetFilterIDByIntRecID = Server.CreateObject("ADODB.Connection")
	cnnGetFilterIDByIntRecID.open Session("ClientCnnString")

	resultGetFilterIDByIntRecID = ""
		
	SQLGetFilterIDByIntRecID = "SELECT * FROM IC_Filters WHERE InternalRecordIdentifier = " & passedMovementCodeIntRecID
	 
	Set rsGetFilterIDByIntRecID  = Server.CreateObject("ADODB.Recordset")
	rsGetFilterIDByIntRecID.CursorLocation = 3 
	
	rsGetFilterIDByIntRecID.Open SQLGetFilterIDByIntRecID,cnnGetFilterIDByIntRecID 
			
	resultGetFilterIDByIntRecID = rsGetFilterIDByIntRecID("FilterID")
	
	rsGetFilterIDByIntRecID.Close
	set rsGetFilterIDByIntRecID = Nothing
	cnnGetFilterIDByIntRecID.Close	
	set cnnGetFilterIDByIntRecID = Nothing
	
	GetFilterIDByIntRecID  = resultGetFilterIDByIntRecID 
	
End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************


'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function GetFilterDescByFilterIntRecID(passedFilterIntRecID)

	Set cnnGetFilterDescByFilterIntRecID = Server.CreateObject("ADODB.Connection")
	cnnGetFilterDescByFilterIntRecID.open Session("ClientCnnString")

	resultGetFilterDescByFilterIntRecID = ""
		
	SQLGetFilterDescByFilterIntRecID = "SELECT * FROM IC_Filters WHERE InternalRecordIdentifier = " & passedFilterIntRecID
	 
	Set rsGetFilterDescByFilterIntRecID  = Server.CreateObject("ADODB.Recordset")
	rsGetFilterDescByFilterIntRecID.CursorLocation = 3 
	
	rsGetFilterDescByFilterIntRecID.Open SQLGetFilterDescByFilterIntRecID,cnnGetFilterDescByFilterIntRecID 
			
	resultGetFilterDescByFilterIntRecID = rsGetFilterDescByFilterIntRecID("Description")
	
	rsGetFilterDescByFilterIntRecID.Close
	set rsGetFilterDescByFilterIntRecID = Nothing
	cnnGetFilterDescByFilterIntRecID.Close	
	set cnnGetFilterDescByFilterIntRecID = Nothing
	
	GetFilterDescByFilterIntRecID  = resultGetFilterDescByFilterIntRecID 
	
End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************



'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function GetFilterLocationByFilterIntRecID(passedFilterIntRecID)

	Set cnnGetFilterLocationByFilterIntRecID = Server.CreateObject("ADODB.Connection")
	cnnGetFilterLocationByFilterIntRecID.open Session("ClientCnnString")

	resultGetFilterLocationByFilterIntRecID = ""
		
	SQLGetFilterLocationByFilterIntRecID = "SELECT * FROM FS_CustomerFilters WHERE FilterIntRecID = " & passedFilterIntRecID
	 
	Set rsGetFilterLocationByFilterIntRecID  = Server.CreateObject("ADODB.Recordset")
	rsGetFilterLocationByFilterIntRecID.CursorLocation = 3 
	
	rsGetFilterLocationByFilterIntRecID.Open SQLGetFilterLocationByFilterIntRecID,cnnGetFilterLocationByFilterIntRecID 
			
	resultGetFilterLocationByFilterIntRecID = rsGetFilterLocationByFilterIntRecID("Notes")
	
	rsGetFilterLocationByFilterIntRecID.Close
	set rsGetFilterLocationByFilterIntRecID = Nothing
	cnnGetFilterLocationByFilterIntRecID.Close	
	set cnnGetFilterLocationByFilterIntRecID = Nothing
	
	GetFilterLocationByFilterIntRecID  = resultGetFilterLocationByFilterIntRecID 
	
End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************


'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function FilterChangeSubmittedNewLogic(passedCustomerFilterIntRecID)

	' Looks to see if that filter for that customer is on FS_ServiceMemosFilterInfo and not completed
	
	resultFilterChangeSubmittedNewLogic = False
	
	Set cnnFilterChangeSubmittedNewLogic = Server.CreateObject("ADODB.Connection")
	cnnFilterChangeSubmittedNewLogic.open Session("ClientCnnString")
	Set rsFilterChangeSubmittedNewLogic  = Server.CreateObject("ADODB.Recordset")
		
	SQLFilterChangeSubmittedNewLogic = "SELECT * FROM FS_CustomerFilters WHERE InternalRecordIdentifier = " & passedCustomerFilterIntRecID
	 
	Set rsFilterChangeSubmittedNewLogic  = cnnFilterChangeSubmittedNewLogic.Execute(SQLFilterChangeSubmittedNewLogic)
	
	If Not rsFilterChangeSubmittedNewLogic.EOF Then
	
		CustID = rsFilterChangeSubmittedNewLogic("CustID")
	
		SQLFilterChangeSubmittedNewLogic = "SELECT DISTINCT ServiceTicketID FROM FS_ServiceMemosFilterInfo WHERE CustID = '" & CustID & "'"
		SQLFilterChangeSubmittedNewLogic =SQLFilterChangeSubmittedNewLogic & " AND ServiceTicketID IN "
		SQLFilterChangeSubmittedNewLogic =SQLFilterChangeSubmittedNewLogic & " (SELECT MemoNumber FROM FS_ServiceMemos WHERE CurrentStatus='OPEN' or CurrentStatus='HOLD') "	

		Set rsFilterChangeSubmittedNewLogic  = cnnFilterChangeSubmittedNewLogic.Execute(SQLFilterChangeSubmittedNewLogic)
		
		If Not rsFilterChangeSubmittedNewLogic.EOF Then resultFilterChangeSubmittedNewLogic = True ' Found one
		
	End If

	rsFilterChangeSubmittedNewLogic.Close
	set rsFilterChangeSubmittedNewLogic = Nothing
	cnnFilterChangeSubmittedNewLogic.Close	
	set cnnFilterChangeSubmittedNewLogic = Nothing
	
	FilterChangeSubmittedNewLogic  = resultFilterChangeSubmittedNewLogic 
	
End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************


'**************************************************************************************************************************************
'**************************************************************************************************************************************

Function GetNumberOfFilterChangesForServiceTicket(passedServiceTicketNumber)

	Set cnnNumOfFilterChangesForServiceTicket = Server.CreateObject("ADODB.Connection")
	cnnNumOfFilterChangesForServiceTicket.open Session("ClientCnnString")

	resultNumOfFilterChangesForServiceTicket = 0
	
	SQLNumOfFilterChangesForServiceTicket = "SELECT COUNT(*) AS FilterCount FROM FS_ServiceMemosFilterInfo WHERE ServiceTicketID = '" & passedServiceTicketNumber & "'"

	'Response.Write(SQLNumOfFilterChangesForServiceTicket)
	 
	Set rsNumOfFilterChangesForServiceTicket = Server.CreateObject("ADODB.Recordset")
	rsNumOfFilterChangesForServiceTicket.CursorLocation = 3 
	
	rsNumOfFilterChangesForServiceTicket.Open SQLNumOfFilterChangesForServiceTicket,cnnNumOfFilterChangesForServiceTicket

	If NOT rsNumOfFilterChangesForServiceTicket.EOF Then
		resultNumOfFilterChangesForServiceTicket = rsNumOfFilterChangesForServiceTicket.Fields("FilterCount")
	End If
	
	rsNumOfFilterChangesForServiceTicket.Close
	set rsNumOfFilterChangesForServiceTicket= Nothing
	cnnNumOfFilterChangesForServiceTicket.Close
	set cnnNumOfFilterChangesForServiceTicket= Nothing
	
	GetNumberOfFilterChangesForServiceTicket = resultNumOfFilterChangesForServiceTicket
	
End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************


'**************************************************************************************************************************************
'**************************************************************************************************************************************

Function GetNumberOfClosedFilterTicketsForTech(passedCloseDate,passedTechUserNo)

	Set cnnNumberOfClosedFilterTicketsForTech = Server.CreateObject("ADODB.Connection")
	cnnNumberOfClosedFilterTicketsForTech.open Session("ClientCnnString")

	resultNumberOfClosedFilterTicketsForTech = 0
	
	SQLNumberOfClosedFilterTicketsForTech = " SELECT COUNT(*) AS NumCallsForTech FROM FS_ServiceMemos WHERE CurrentStatus = 'CLOSE' AND RecordSubType = 'CLOSE' AND UserNoOfServiceTech = " & passedTechUserNo & " "
	SQLNumberOfClosedFilterTicketsForTech = SQLNumberOfClosedFilterTicketsForTech & " AND Month(RecordCreatedateTime) = " & Month(passedCloseDate)
	SQLNumberOfClosedFilterTicketsForTech = SQLNumberOfClosedFilterTicketsForTech & " AND Year(RecordCreatedateTime) = " & Year(passedCloseDate)
	SQLNumberOfClosedFilterTicketsForTech = SQLNumberOfClosedFilterTicketsForTech & " AND Day(RecordCreatedateTime) = " & Day(passedCloseDate) 
	SQLNumberOfClosedFilterTicketsForTech = SQLNumberOfClosedFilterTicketsForTech & " AND FilterChange = 1 "
		
	'Response.Write(SQLNumberOfClosedFilterTicketsForTech)
	 
	Set rsNumberOfClosedFilterTicketsForTech = Server.CreateObject("ADODB.Recordset")
	rsNumberOfClosedFilterTicketsForTech.CursorLocation = 3 
	
	rsNumberOfClosedFilterTicketsForTech.Open SQLNumberOfClosedFilterTicketsForTech,cnnNumberOfClosedFilterTicketsForTech

	If NOT rsNumberOfClosedFilterTicketsForTech.EOF Then
		resultNumberOfClosedFilterTicketsForTech = rsNumberOfClosedFilterTicketsForTech.Fields("NumCallsForTech")
	End If
	
	rsNumberOfClosedFilterTicketsForTech.Close
	set rsNumberOfClosedFilterTicketsForTech= Nothing
	cnnNumberOfClosedFilterTicketsForTech.Close
	set cnnNumberOfClosedFilterTicketsForTech= Nothing
	
	GetNumberOfClosedFilterTicketsForTech = resultNumberOfClosedFilterTicketsForTech
	
End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************



'**************************************************************************************************************************************
'**************************************************************************************************************************************

Function GetNumberOfClosedServiceTicketsForTech(passedCloseDate,passedTechUserNo)

	Set cnnNumberOfClosedServiceTicketsForTech = Server.CreateObject("ADODB.Connection")
	cnnNumberOfClosedServiceTicketsForTech.open Session("ClientCnnString")

	resultNumberOfClosedServiceTicketsForTech = 0
	
	SQLNumberOfClosedServiceTicketsForTech = " SELECT COUNT(*) AS NumCallsForTech FROM FS_ServiceMemos WHERE CurrentStatus = 'CLOSE' AND RecordSubType = 'CLOSE' AND UserNoOfServiceTech = " & passedTechUserNo & " "
	SQLNumberOfClosedServiceTicketsForTech = SQLNumberOfClosedServiceTicketsForTech & " AND Month(RecordCreatedateTime) = " & Month(passedCloseDate)
	SQLNumberOfClosedServiceTicketsForTech = SQLNumberOfClosedServiceTicketsForTech & " AND Year(RecordCreatedateTime) = " & Year(passedCloseDate)
	SQLNumberOfClosedServiceTicketsForTech = SQLNumberOfClosedServiceTicketsForTech & " AND Day(RecordCreatedateTime) = " & Day(passedCloseDate) 
	SQLNumberOfClosedServiceTicketsForTech = SQLNumberOfClosedServiceTicketsForTech & " AND FilterChange <> 1 "
		
	'Response.Write(SQLNumberOfClosedServiceTicketsForTech)
	 
	Set rsNumberOfClosedServiceTicketsForTech = Server.CreateObject("ADODB.Recordset")
	rsNumberOfClosedServiceTicketsForTech.CursorLocation = 3 
	
	rsNumberOfClosedServiceTicketsForTech.Open SQLNumberOfClosedServiceTicketsForTech,cnnNumberOfClosedServiceTicketsForTech

	If NOT rsNumberOfClosedServiceTicketsForTech.EOF Then
		resultNumberOfClosedServiceTicketsForTech = rsNumberOfClosedServiceTicketsForTech.Fields("NumCallsForTech")
	End If
	
	rsNumberOfClosedServiceTicketsForTech.Close
	set rsNumberOfClosedServiceTicketsForTech= Nothing
	cnnNumberOfClosedServiceTicketsForTech.Close
	set cnnNumberOfClosedServiceTicketsForTech= Nothing
	
	GetNumberOfClosedServiceTicketsForTech = resultNumberOfClosedServiceTicketsForTech
	
End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************



'**************************************************************************************************************************************
'**************************************************************************************************************************************

Function GetNumberOfClosedFilterChangesForTech(passedCloseDate,passedTechUserNo)

	Set cnnNumberOfClosedFilterChangesForTech = Server.CreateObject("ADODB.Connection")
	cnnNumberOfClosedFilterChangesForTech.open Session("ClientCnnString")

	Set cnnFSServiceMemosFilterInfo = Server.CreateObject("ADODB.Connection")
	cnnFSServiceMemosFilterInfo.open Session("ClientCnnString")

	resultNumberOfClosedFilterChangesForTech = 0
	
	SQLNumberOfClosedFilterChangesForTech = " SELECT * FROM FS_ServiceMemos WHERE CurrentStatus = 'CLOSE' AND RecordSubType = 'CLOSE' AND UserNoOfServiceTech = " & passedTechUserNo & " "
	SQLNumberOfClosedFilterChangesForTech = SQLNumberOfClosedFilterChangesForTech & " AND Month(RecordCreatedateTime) = " & Month(passedCloseDate)
	SQLNumberOfClosedFilterChangesForTech = SQLNumberOfClosedFilterChangesForTech & " AND Year(RecordCreatedateTime) = " & Year(passedCloseDate)
	SQLNumberOfClosedFilterChangesForTech = SQLNumberOfClosedFilterChangesForTech & " AND Day(RecordCreatedateTime) = " & Day(passedCloseDate) 
	SQLNumberOfClosedFilterChangesForTech = SQLNumberOfClosedFilterChangesForTech & " AND FilterChange = 1 "
		
	'Response.Write(SQLNumberOfClosedFilterChangesForTech)
	 
	Set rsNumberOfClosedFilterChangesForTech = Server.CreateObject("ADODB.Recordset")
	rsNumberOfClosedFilterChangesForTech.CursorLocation = 3 
	
	rsNumberOfClosedFilterChangesForTech.Open SQLNumberOfClosedFilterChangesForTech,cnnNumberOfClosedFilterChangesForTech

	If NOT rsNumberOfClosedFilterChangesForTech.EOF Then
	
		Do While NOT rsNumberOfClosedFilterChangesForTech.EOF
			ServiceTicketNumber = rsNumberOfClosedFilterChangesForTech.Fields("MemoNumber")
			resultNumberOfClosedFilterChangesForTech = resultNumberOfClosedFilterChangesForTech + GetNumberOfFilterChangesForServiceTicket(ServiceTicketNumber)
			rsNumberOfClosedFilterChangesForTech.MoveNext
		Loop
		
	End If
	
	rsNumberOfClosedFilterChangesForTech.Close
	set rsNumberOfClosedFilterChangesForTech= Nothing
	cnnNumberOfClosedFilterChangesForTech.Close
	set cnnNumberOfClosedFilterChangesForTech= Nothing
	
	GetNumberOfClosedFilterChangesForTech = resultNumberOfClosedFilterChangesForTech
	
End Function


'**************************************************************************************************************************************
'**************************************************************************************************************************************


'**************************************************************************************************************************************
'**************************************************************************************************************************************

Function GetNumberOfClosedFilterTicketsForCustomerAcct(passedCloseDate,passedCustID)

	Set cnnNumberOfClosedFilterTicketsForCustomerAcct = Server.CreateObject("ADODB.Connection")
	cnnNumberOfClosedFilterTicketsForCustomerAcct.open Session("ClientCnnString")

	resultNumberOfClosedFilterTicketsForCustomerAcct = 0
	
	SQLNumberOfClosedFilterTicketsForCustomerAcct = " SELECT COUNT(*) AS NumCallsForCustomerAcct FROM FS_ServiceMemos WHERE CurrentStatus = 'CLOSE' AND RecordSubType = 'CLOSE' AND AccountNumber = '" & passedCustID & "' "
	SQLNumberOfClosedFilterTicketsForCustomerAcct = SQLNumberOfClosedFilterTicketsForCustomerAcct & " AND Month(RecordCreatedateTime) = " & Month(passedCloseDate)
	SQLNumberOfClosedFilterTicketsForCustomerAcct = SQLNumberOfClosedFilterTicketsForCustomerAcct & " AND Year(RecordCreatedateTime) = " & Year(passedCloseDate)
	SQLNumberOfClosedFilterTicketsForCustomerAcct = SQLNumberOfClosedFilterTicketsForCustomerAcct & " AND Day(RecordCreatedateTime) = " & Day(passedCloseDate) 
	SQLNumberOfClosedFilterTicketsForCustomerAcct = SQLNumberOfClosedFilterTicketsForCustomerAcct & " AND FilterChange = 1 "
		
	'Response.Write(SQLNumberOfClosedFilterTicketsForCustomerAcct)
	 
	Set rsNumberOfClosedFilterTicketsForCustomerAcct = Server.CreateObject("ADODB.Recordset")
	rsNumberOfClosedFilterTicketsForCustomerAcct.CursorLocation = 3 
	
	rsNumberOfClosedFilterTicketsForCustomerAcct.Open SQLNumberOfClosedFilterTicketsForCustomerAcct,cnnNumberOfClosedFilterTicketsForCustomerAcct

	If NOT rsNumberOfClosedFilterTicketsForCustomerAcct.EOF Then
		resultNumberOfClosedFilterTicketsForCustomerAcct = rsNumberOfClosedFilterTicketsForCustomerAcct.Fields("NumCallsForCustomerAcct")
	End If
	
	rsNumberOfClosedFilterTicketsForCustomerAcct.Close
	set rsNumberOfClosedFilterTicketsForCustomerAcct= Nothing
	cnnNumberOfClosedFilterTicketsForCustomerAcct.Close
	set cnnNumberOfClosedFilterTicketsForCustomerAcct= Nothing
	
	GetNumberOfClosedFilterTicketsForCustomerAcct = resultNumberOfClosedFilterTicketsForCustomerAcct
	
End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************



'**************************************************************************************************************************************
'**************************************************************************************************************************************

Function GetNumberOfClosedServiceTicketsForCustomerAcct(passedCloseDate,passedCustID)

	Set cnnNumberOfClosedServiceTicketsForCustomerAcct = Server.CreateObject("ADODB.Connection")
	cnnNumberOfClosedServiceTicketsForCustomerAcct.open Session("ClientCnnString")

	resultNumberOfClosedServiceTicketsForCustomerAcct = 0
	
	SQLNumberOfClosedServiceTicketsForCustomerAcct = " SELECT COUNT(*) AS NumCallsForCustomerAcct FROM FS_ServiceMemos WHERE CurrentStatus = 'CLOSE' AND RecordSubType = 'CLOSE' AND AccountNumber = '" & passedCustID & "' "
	SQLNumberOfClosedServiceTicketsForCustomerAcct = SQLNumberOfClosedServiceTicketsForCustomerAcct & " AND Month(RecordCreatedateTime) = " & Month(passedCloseDate)
	SQLNumberOfClosedServiceTicketsForCustomerAcct = SQLNumberOfClosedServiceTicketsForCustomerAcct & " AND Year(RecordCreatedateTime) = " & Year(passedCloseDate)
	SQLNumberOfClosedServiceTicketsForCustomerAcct = SQLNumberOfClosedServiceTicketsForCustomerAcct & " AND Day(RecordCreatedateTime) = " & Day(passedCloseDate) 
	SQLNumberOfClosedServiceTicketsForCustomerAcct = SQLNumberOfClosedServiceTicketsForCustomerAcct & " AND FilterChange <> 1 "
		
	'Response.Write(SQLNumberOfClosedServiceTicketsForCustomerAcct)
	 
	Set rsNumberOfClosedServiceTicketsForCustomerAcct = Server.CreateObject("ADODB.Recordset")
	rsNumberOfClosedServiceTicketsForCustomerAcct.CursorLocation = 3 
	
	rsNumberOfClosedServiceTicketsForCustomerAcct.Open SQLNumberOfClosedServiceTicketsForCustomerAcct,cnnNumberOfClosedServiceTicketsForCustomerAcct

	If NOT rsNumberOfClosedServiceTicketsForCustomerAcct.EOF Then
		resultNumberOfClosedServiceTicketsForCustomerAcct = rsNumberOfClosedServiceTicketsForCustomerAcct.Fields("NumCallsForCustomerAcct")
	End If
	
	rsNumberOfClosedServiceTicketsForCustomerAcct.Close
	set rsNumberOfClosedServiceTicketsForCustomerAcct= Nothing
	cnnNumberOfClosedServiceTicketsForCustomerAcct.Close
	set cnnNumberOfClosedServiceTicketsForCustomerAcct= Nothing
	
	GetNumberOfClosedServiceTicketsForCustomerAcct = resultNumberOfClosedServiceTicketsForCustomerAcct
	
End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************



'**************************************************************************************************************************************
'**************************************************************************************************************************************

Function GetNumberOfClosedFilterChangesForCustomerAcct(passedCloseDate,passedCustID)

	Set cnnNumberOfClosedFilterChangesForCustomerAcct = Server.CreateObject("ADODB.Connection")
	cnnNumberOfClosedFilterChangesForCustomerAcct.open Session("ClientCnnString")

	Set cnnFSServiceMemosFilterInfo = Server.CreateObject("ADODB.Connection")
	cnnFSServiceMemosFilterInfo.open Session("ClientCnnString")

	resultNumberOfClosedFilterChangesForCustomerAcct = 0
	
	SQLNumberOfClosedFilterChangesForCustomerAcct = " SELECT * FROM FS_ServiceMemos WHERE CurrentStatus = 'CLOSE' AND RecordSubType = 'CLOSE' AND AccountNumber = '" & passedCustID & "' "
	SQLNumberOfClosedFilterChangesForCustomerAcct = SQLNumberOfClosedFilterChangesForCustomerAcct & " AND Month(RecordCreatedateTime) = " & Month(passedCloseDate)
	SQLNumberOfClosedFilterChangesForCustomerAcct = SQLNumberOfClosedFilterChangesForCustomerAcct & " AND Year(RecordCreatedateTime) = " & Year(passedCloseDate)
	SQLNumberOfClosedFilterChangesForCustomerAcct = SQLNumberOfClosedFilterChangesForCustomerAcct & " AND Day(RecordCreatedateTime) = " & Day(passedCloseDate) 
	SQLNumberOfClosedFilterChangesForCustomerAcct = SQLNumberOfClosedFilterChangesForCustomerAcct & " AND FilterChange = 1 "
		
	'Response.Write(SQLNumberOfClosedFilterChangesForCustomerAcct)
	 
	Set rsNumberOfClosedFilterChangesForCustomerAcct = Server.CreateObject("ADODB.Recordset")
	rsNumberOfClosedFilterChangesForCustomerAcct.CursorLocation = 3 
	
	rsNumberOfClosedFilterChangesForCustomerAcct.Open SQLNumberOfClosedFilterChangesForCustomerAcct,cnnNumberOfClosedFilterChangesForCustomerAcct

	If NOT rsNumberOfClosedFilterChangesForCustomerAcct.EOF Then
	
		Do While NOT rsNumberOfClosedFilterChangesForCustomerAcct.EOF
			ServiceTicketNumber = rsNumberOfClosedFilterChangesForCustomerAcct.Fields("MemoNumber")
			resultNumberOfClosedFilterChangesForCustomerAcct = resultNumberOfClosedFilterChangesForCustomerAcct + GetNumberOfFilterChangesForServiceTicket(ServiceTicketNumber)
			rsNumberOfClosedFilterChangesForCustomerAcct.MoveNext
		Loop
		
	End If
	
	rsNumberOfClosedFilterChangesForCustomerAcct.Close
	set rsNumberOfClosedFilterChangesForCustomerAcct= Nothing
	cnnNumberOfClosedFilterChangesForCustomerAcct.Close
	set cnnNumberOfClosedFilterChangesForCustomerAcct= Nothing
	
	GetNumberOfClosedFilterChangesForCustomerAcct = resultNumberOfClosedFilterChangesForCustomerAcct
	
End Function


'**************************************************************************************************************************************
'**************************************************************************************************************************************



Function TicketInServiceMemosFilterInfo(passedServiceTicketID)

	resultTicketInServiceMemosFilterInfo = ""
	
	Set cnnTicketInServiceMemosFilterInfo = Server.CreateObject("ADODB.Connection")
	cnnTicketInServiceMemosFilterInfo.open Session("ClientCnnString")
	Set rsTicketInServiceMemosFilterInfo  = Server.CreateObject("ADODB.Recordset")
		
	SQLTicketInServiceMemosFilterInfo = "SELECT ServiceTicketID FROM FS_ServiceMemosFilterInfo WHERE ServiceTicketID = '" & passedServiceTicketID & "' "
	
	Set rsTicketInServiceMemosFilterInfo  = cnnTicketInServiceMemosFilterInfo.Execute(SQLTicketInServiceMemosFilterInfo)
	
	If Not rsTicketInServiceMemosFilterInfo.EOF Then
		resultTicketInServiceMemosFilterInfo = True
	Else
	 	resultTicketInServiceMemosFilterInfo = False
 	End If
	
	rsTicketInServiceMemosFilterInfo.Close
	set rsTicketInServiceMemosFilterInfo = Nothing
	cnnTicketInServiceMemosFilterInfo.Close	
	set cnnTicketInServiceMemosFilterInfo = Nothing
	
	TicketInServiceMemosFilterInfo  = resultTicketInServiceMemosFilterInfo 
	
End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************


Function GetOpenFilterTicketsByCustID(passedCustID)

	' Returns a list of Non-Closed Filter Tickets for a customer separated by commas

	resultGetOpenFilterTicketsByCustID = ""
	
	Set cnnGetOpenFilterTicketsByCustID = Server.CreateObject("ADODB.Connection")
	cnnGetOpenFilterTicketsByCustID.open Session("ClientCnnString")
	Set rsGetOpenFilterTicketsByCustID  = Server.CreateObject("ADODB.Recordset")
		
	SQLGetOpenFilterTicketsByCustID = "SELECT DISTINCT ServiceTicketID FROM FS_ServiceMemosFilterInfo WHERE CustID = '" & passedCustID & "'"
	SQLGetOpenFilterTicketsByCustID =SQLGetOpenFilterTicketsByCustID & " AND ServiceTicketID IN "
	SQLGetOpenFilterTicketsByCustID =SQLGetOpenFilterTicketsByCustID & " (SELECT MemoNumber FROM FS_ServiceMemos WHERE CurrentStatus='OPEN' or CurrentStatus='HOLD') "	
	
	Set rsGetOpenFilterTicketsByCustID  = cnnGetOpenFilterTicketsByCustID.Execute(SQLGetOpenFilterTicketsByCustID)
	
	If Not rsGetOpenFilterTicketsByCustID.EOF Then
	
		Do While NOT rsGetOpenFilterTicketsByCustID.EOF
		
				 resultGetOpenFilterTicketsByCustID = resultGetOpenFilterTicketsByCustID & rsGetOpenFilterTicketsByCustID("ServiceTicketID") & ","
		
			rsGetOpenFilterTicketsByCustID.movenext
		Loop

 
 	End If
	
	rsGetOpenFilterTicketsByCustID.Close
	set rsGetOpenFilterTicketsByCustID = Nothing
	cnnGetOpenFilterTicketsByCustID.Close	
	set cnnGetOpenFilterTicketsByCustID = Nothing
	
	GetOpenFilterTicketsByCustID  = resultGetOpenFilterTicketsByCustID 
	
End Function


Function CustHasPendingFilterChange(passedCustid)

	'Remember, it reads the setting FieldServiceDays from tblSetting_Global
	'to determine how many days to use in the evaluation
	SQLCustHasPendingFilterChange = "SELECT * FROM Settings_Global"
	Set cnnCustHasPendingFilterChange = Server.CreateObject("ADODB.Connection")
	cnnCustHasPendingFilterChange.open (Session("ClientCnnString"))
	Set rsCustHasPendingFilterChange = Server.CreateObject("ADODB.Recordset")
	rsCustHasPendingFilterChange.CursorLocation = 3 
	Set rsCustHasPendingFilterChange = cnnCustHasPendingFilterChange.Execute(SQLCustHasPendingFilterChange)
	If not rsCustHasPendingFilterChange.EOF Then FilterChangeDays = rsCustHasPendingFilterChange("FilterChangeDays") Else FilterChangeDays = 15
	set rsCustHasPendingFilterChange = Nothing
	cnnCustHasPendingFilterChange.close
	set cnnCustHasPendingFilterChange = Nothing
	
	resultCustHasPendingFilterChange = False

	Set cnnCustHasPendingFilterChange = Server.CreateObject("ADODB.Connection")
	cnnCustHasPendingFilterChange.open Session("ClientCnnString")
	Set rsCustHasPendingFilterChange = Server.CreateObject("ADODB.Recordset")
	
	
		'Gets all filter next changes dates for this customer		
		
		SQLCustHasPendingFilterChange = "SELECT NextChangeDate, InternalRecordIdentifier "
		SQLCustHasPendingFilterChange = SQLCustHasPendingFilterChange & "FROM "
		SQLCustHasPendingFilterChange = SQLCustHasPendingFilterChange & "(SELECT *, "
		SQLCustHasPendingFilterChange = SQLCustHasPendingFilterChange & "CASE FrequencyType WHEN 'D' THEN dateadd(d,FrequencyTime, LastChangeDateTime) "
		SQLCustHasPendingFilterChange = SQLCustHasPendingFilterChange & "WHEN 'W' THEN dateadd(d, (FrequencyTime * 7), LastChangeDateTime) "
		SQLCustHasPendingFilterChange = SQLCustHasPendingFilterChange & "WHEN 'M' THEN dateadd(d, (FrequencyTime * 28), LastChangeDateTime) END AS NextChangeDate "
		SQLCustHasPendingFilterChange = SQLCustHasPendingFilterChange & "FROM FS_CustomerFilters WHERE CustID = '" & passedCustid & "' AND FS_CustomerFilters.InternalRecordIdentifier NOT IN "
		SQLCustHasPendingFilterChange = SQLCustHasPendingFilterChange & " (SELECT CustFilterIntRecID FROM FS_ServiceMemosFilterInfo WHERE ServiceTicketID NOT IN "
		
		SQLCustHasPendingFilterChange = SQLCustHasPendingFilterChange & " (SELECT MemoNumber FROM FS_ServiceMemos WHERE CurrentStatus='OPEN' or CurrentStatus='HOLD') "
		
		SQLCustHasPendingFilterChange = SQLCustHasPendingFilterChange & "))"
		SQLCustHasPendingFilterChange = SQLCustHasPendingFilterChange & " AS derivedtbl_1 "


'response.write(SQLCustHasPendingFilterChange)

		Set rsCustHasPendingFilterChange = cnnCustHasPendingFilterChange.Execute(SQLCustHasPendingFilterChange)
		
		If rsCustHasPendingFilterChange.EOF Then
			resultCustHasPendingFilterChange = False
		Else
			Do While Not rsCustHasPendingFilterChange.EOF
			
				'If ANY of the dates are in range, then the answer is true
				NextDate = rsCustHasPendingFilterChange("NextChangeDate")
				TodayPlusFilterDays = DateAdd("d",FilterChangeDays,Date())
				
				DaysTilChange = datediff("d",Date(),NextDate)

				If DaysTilChange <= FilterChangeDays Then
				
					If FilterChangeSubmittedNewLogic(rsCustHasPendingFilterChange("InternalRecordIdentifier")) <> True Then				
						resultCustHasPendingFilterChange = True
						Exit Do ' No need to look further, they have at least 1 within the allowable window or overdue
					End If
					
				End If
			
				rsCustHasPendingFilterChange.MoveNext
			Loop
		End If
		
	
	CustHasPendingFilterChange = resultCustHasPendingFilterChange

	set rsCustHasPendingFilterChange = Nothing
	cnnCustHasPendingFilterChange.Close
	set cnnCustHasPendingFilterChange = Nothing

End Function

Function GetNumberOfServiceCallsOnHold()

	Set cnnGetNumberOfServiceCallsOnHold = Server.CreateObject("ADODB.Connection")
	cnnGetNumberOfServiceCallsOnHold.open Session("ClientCnnString")

	resultGetNumberOfServiceCallsOnHold = 0
	
	SQLGetNumberOfServiceCallsOnHold = "SELECT Count(*) AS HoldCount FROM FS_ServiceMemos WHERE CurrentStatus = 'HOLD'"
	 
	Set rsGetNumberOfServiceCallsOnHold = Server.CreateObject("ADODB.Recordset")
	rsGetNumberOfServiceCallsOnHold.CursorLocation = 3 
	
	rsGetNumberOfServiceCallsOnHold.Open SQLGetNumberOfServiceCallsOnHold , cnnGetNumberOfServiceCallsOnHold
	
	If NOT rsGetNumberOfServiceCallsOnHold.EOF Then resultGetNumberOfServiceCallsOnHold = rsGetNumberOfServiceCallsOnHold("HoldCount")

	rsGetNumberOfServiceCallsOnHold.Close
	set rsGetNumberOfServiceCallsOnHold= Nothing
	cnnGetNumberOfServiceCallsOnHold.Close
	set cnnGetNumberOfServiceCallsOnHold= Nothing
	
	GetNumberOfServiceCallsOnHold = resultGetNumberOfServiceCallsOnHold
	
End Function

Function GetHOLDServiceTicketSTAGEDateTime(passedTicketNumber,passedStage)

	'Use only when advanced dispatch module is on

	resultGetHOLDServiceTicketSTAGEDateTime = ""
	
	Set cnnGetHOLDServiceTicketSTAGEDateTime = Server.CreateObject("ADODB.Connection")
	cnnGetHOLDServiceTicketSTAGEDateTime.open Session("ClientCnnString")

	SQLGetHOLDServiceTicketSTAGEDateTime = "SELECT * FROM FS_ServiceMemos where MemoNumber = '" & passedTicketNumber & "' AND RecordSubType = 'HOLD'"

	Set rsGetHOLDServiceTicketSTAGEDateTime = Server.CreateObject("ADODB.Recordset")
	rsGetHOLDServiceTicketSTAGEDateTime.CursorLocation = 3 
	Set rsGetHOLDServiceTicketSTAGEDateTime = cnnGetHOLDServiceTicketSTAGEDateTime.Execute(SQLGetHOLDServiceTicketSTAGEDateTime)
	
	If not rsGetHOLDServiceTicketSTAGEDateTime.eof then 
		resultGetHOLDServiceTicketSTAGEDateTime = rsGetHOLDServiceTicketSTAGEDateTime("SubmissionDateTime")
	End IF	
	
	set rsGetHOLDServiceTicketSTAGEDateTime = Nothing
	cnnGetHOLDServiceTicketSTAGEDateTime.Close
	set cnnGetHOLDServiceTicketSTAGEDateTime = Nothing
	
	GetHOLDServiceTicketSTAGEDateTime = resultGetHOLDServiceTicketSTAGEDateTime

End Function


Function userCreateEquipmentSymptomCodesOnTheFly(passedUserNo)

	resultuserCreateEquipmentSymptomCodesOnTheFly = false

	Set cnnuserCreateEquipmentSymptomCodesOnTheFly  = Server.CreateObject("ADODB.Connection")
	cnnuserCreateEquipmentSymptomCodesOnTheFly.open Session("ClientCnnString")

		
	SQLuserCreateEquipmentSymptomCodesOnTheFly  = "SELECT * FROM tblUsers WHERE UserNo = " & passedUserNo
	 
	Set rsuserCreateEquipmentSymptomCodesOnTheFly  = Server.CreateObject("ADODB.Recordset")
	rsuserCreateEquipmentSymptomCodesOnTheFly.CursorLocation = 3 
	
	rsuserCreateEquipmentSymptomCodesOnTheFly.Open SQLuserCreateEquipmentSymptomCodesOnTheFly , cnnuserCreateEquipmentSymptomCodesOnTheFly 
			
	If not rsuserCreateEquipmentSymptomCodesOnTheFly.eof then 
		If rsuserCreateEquipmentSymptomCodesOnTheFly("userCreateEquipmentSymptomCodesOnTheFly") = 1 Then resultuserCreateEquipmentSymptomCodesOnTheFly = True Else resultuserCreateEquipmentSymptomCodesOnTheFly = False
	End IF
	
	rsuserCreateEquipmentSymptomCodesOnTheFly.Close
	set rsuserCreateEquipmentSymptomCodesOnTheFly = Nothing
	cnnuserCreateEquipmentSymptomCodesOnTheFly.Close	
	set cnnuserCreateEquipmentSymptomCodesOnTheFly = Nothing
	
	userCreateEquipmentSymptomCodesOnTheFly = resultuserCreateEquipmentSymptomCodesOnTheFly 
		
End Function


Function userCreateEquipmentResolutionCodesOnTheFly(passedUserNo)

	resultuserCreateEquipmentResolutionCodesOnTheFly = false

	Set cnnuserCreateEquipmentResolutionCodesOnTheFly  = Server.CreateObject("ADODB.Connection")
	cnnuserCreateEquipmentResolutionCodesOnTheFly.open Session("ClientCnnString")

		
	SQLuserCreateEquipmentResolutionCodesOnTheFly  = "SELECT * FROM tblUsers WHERE UserNo = " & passedUserNo
	 
	Set rsuserCreateEquipmentResolutionCodesOnTheFly  = Server.CreateObject("ADODB.Recordset")
	rsuserCreateEquipmentResolutionCodesOnTheFly.CursorLocation = 3 
	
	rsuserCreateEquipmentResolutionCodesOnTheFly.Open SQLuserCreateEquipmentResolutionCodesOnTheFly , cnnuserCreateEquipmentResolutionCodesOnTheFly 
			
	If not rsuserCreateEquipmentResolutionCodesOnTheFly.eof then 
		If rsuserCreateEquipmentResolutionCodesOnTheFly("userCreateEquipmentResolutionCodesOnTheFly") = 1 Then resultuserCreateEquipmentResolutionCodesOnTheFly = True Else resultuserCreateEquipmentResolutionCodesOnTheFly = False
	End IF
	
	rsuserCreateEquipmentResolutionCodesOnTheFly.Close
	set rsuserCreateEquipmentResolutionCodesOnTheFly = Nothing
	cnnuserCreateEquipmentResolutionCodesOnTheFly.Close	
	set cnnuserCreateEquipmentResolutionCodesOnTheFly = Nothing
	
	userCreateEquipmentResolutionCodesOnTheFly = resultuserCreateEquipmentResolutionCodesOnTheFly 
		
End Function

Function userCreateEquipmentProblemCodesOnTheFly(passedUserNo)

	resultuserCreateEquipmentProblemCodesOnTheFly = false

	Set cnnuserCreateEquipmentProblemCodesOnTheFly  = Server.CreateObject("ADODB.Connection")
	cnnuserCreateEquipmentProblemCodesOnTheFly.open Session("ClientCnnString")

		
	SQLuserCreateEquipmentProblemCodesOnTheFly  = "SELECT * FROM tblUsers WHERE UserNo = " & passedUserNo
	 
	Set rsuserCreateEquipmentProblemCodesOnTheFly  = Server.CreateObject("ADODB.Recordset")
	rsuserCreateEquipmentProblemCodesOnTheFly.CursorLocation = 3 
	
	rsuserCreateEquipmentProblemCodesOnTheFly.Open SQLuserCreateEquipmentProblemCodesOnTheFly , cnnuserCreateEquipmentProblemCodesOnTheFly 
			
	If not rsuserCreateEquipmentProblemCodesOnTheFly.eof then 
		If rsuserCreateEquipmentProblemCodesOnTheFly("userCreateEquipmentProblemCodesOnTheFly") = 1 Then resultuserCreateEquipmentProblemCodesOnTheFly = True Else resultuserCreateEquipmentProblemCodesOnTheFly = False
	End IF
	
	rsuserCreateEquipmentProblemCodesOnTheFly.Close
	set rsuserCreateEquipmentProblemCodesOnTheFly = Nothing
	cnnuserCreateEquipmentProblemCodesOnTheFly.Close	
	set cnnuserCreateEquipmentProblemCodesOnTheFly = Nothing
	
	userCreateEquipmentProblemCodesOnTheFly = resultuserCreateEquipmentProblemCodesOnTheFly 
		
End Function


%>

