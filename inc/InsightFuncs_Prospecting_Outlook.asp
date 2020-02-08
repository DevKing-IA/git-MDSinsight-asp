<%
'Need these here
Dim	ExtraFirstName
Dim	ExtraLastName
Dim	ExtraEmail
Dim	ExtraPhone
Dim	ExtraPhoneExt
Dim	ExtraStreet
Dim	ExtraAddress2
Dim	ExtraCity
Dim	ExtraState
Dim	ExtraPostalCode
Dim	ExtraPrimaryCompetitorName 

'********************************
'List of all the functions & subs
'********************************

'Func SetOwner_MakeOutlookEntry_SendEmail (passedProspectID,passedNewOwnerUserNo,passedSendEmailFlag,passedPageSource)
'Func Prospect_Email_Accept (passedProspectID,passedNewOwnerUserNo)

'************************************
'End List of all the functions & subs
'************************************


Function SetOwner_MakeOutlookEntry_SendEmail (passedProspectID,passedNewOwnerUserNo,passedSendEmailFlag,passedPageSource)

	'passedPageSource: R = Recycle, E = Edit Prospect, A = Add Prospect ,O = Owner Request Email Accepted???
	

	DoNotSendEmail = False
	If passedSendEmailFlag = 0 or passedSendEmailFlag = "N" Then DoNotSendEmail = True
	
	UserNoForCalendarUpdate = 0
	

	
	'******************************************************************************************************************************************
	'******************************************************************************************************************************************
	'******************************************************************************************************************************************
	'******************************************************************************************************************************************		
	'UPDATE PROSPECT OWNERSHIP
	'******************************************************************************************************************************************
	'******************************************************************************************************************************************
	'******************************************************************************************************************************************
	'******************************************************************************************************************************************	
	
	'*****************************************************************************************************************************************
	'If new prospect owner is not set to the current user, email the propspective prospect owner for approval before adding them as the owner
	'*****************************************************************************************************************************************
	If (cint(Session("UserNo")) <> cint(passedNewOwnerUserNo)) Then
	
		If DoNotSendEmail = True Then
		
			'***************************************
			'Update owner in PR_Prospects
			'***************************************
	
			Set cnnProspectUpdateOwner = Server.CreateObject("ADODB.Connection")
			cnnProspectUpdateOwner.open Session("ClientCnnString")
			Set rsProspectUpdateOwner = Server.CreateObject("ADODB.Recordset")
			rsProspectUpdateOwner.CursorLocation = 3 

			'Get current owner
			OriginalOwnerUserNo = 0
			SQLProspectUpdateOwner = "Select OwnerUserNo FROM PR_Prospects WHERE InternalRecordIdentifier = " & passedProspectID
			Set rsProspectUpdateOwner = cnnProspectUpdateOwner.Execute(SQLProspectUpdateOwner)
			If Not rsProspectUpdateOwner.EOF Then OriginalOwnerUserNo = rsProspectUpdateOwner("OwnerUserNo")
			
			SQLProspectUpdateOwner = "UPDATE PR_Prospects Set OwnerUserNo = " & passedNewOwnerUserNo & " WHERE InternalRecordIdentifier = " & passedProspectID
			
			Set rsProspectUpdateOwner = cnnProspectUpdateOwner.Execute(SQLProspectUpdateOwner)
			
			set rsProspectUpdateOwner = Nothing
			cnnProspectUpdateOwner.Close
			set cnnProspectUpdateOwner = Nothing
			
			Select Case passedPageSource
				Case "R"
					Description = GetUserDisplayNameByUserNo(Session("UserNo")) & " changed ownership of the prospect " & GetProspectNameByNumber(passedProspectID) & " from  " & 	GetUserDisplayNameByUserNo(OriginalOwnerUserNo) 
					Description = Description & " to " &  GetUserDisplayNameByUserNo(passedNewOwnerUserNo) 
					Description = Description & " when moving the prospect from the " & GetTerm("recycle pool") & " back to the main prospect pool."
				Case "E"
					Description = GetUserDisplayNameByUserNo(Session("UserNo")) & " changed ownership of the prospect " & GetProspectNameByNumber(passedProspectID) & " from  " & GetUserDisplayNameByUserNo(OriginalOwnerUserNo) 
					Description = Description & " to " &  GetUserDisplayNameByUserNo(passedNewOwnerUserNo) 
					Description = Description & " from the Edit Prospect screen."
				Case "A"
					Description = GetUserDisplayNameByUserNo(Session("UserNo")) & " assigned ownership of the prospect " & GetProspectNameByNumber(passedProspectID) & " to "
					Description = Description & GetUserDisplayNameByUserNo(passedNewOwnerUserNo) 
					Description = Description & " while creating the prospect via the Add Prospect screen."
			End Select
			
			Description = Description & " The DO NOT SEND EMAIL option was selected so no email was sent and the new owner was assigned immediately."
			Description = Description & " If the Next Activity for this prospect is a meeting or appointment, and entry will be created in the Outlook calendar belonging to " & GetUserDisplayNameByUserNo(passedNewOwnerUserNo) 
			
			CreateAuditLogEntry GetTerm("Prospecting"),"Prospect owner changed/set","Major",0,Description
			Record_PR_Activity passedProspectID,Description,Session("UserNo")
			
			UserNoForCalendarUpdate = passedNewOwnerUserNo
			
		Else ' Send email option is turned on
		
			'***************************************
			'Update owner in PR_Prospects
			'***************************************
	
			Set cnnProspectUpdateOwner = Server.CreateObject("ADODB.Connection")
			cnnProspectUpdateOwner.open Session("ClientCnnString")
			Set rsProspectUpdateOwner = Server.CreateObject("ADODB.Recordset")
			rsProspectUpdateOwner.CursorLocation = 3 

			'Get current owner
			OriginalOwnerUserNo = 0
			SQLProspectUpdateOwner = "Select OwnerUserNo FROM PR_Prospects WHERE InternalRecordIdentifier = " & passedProspectID
			Set rsProspectUpdateOwner = cnnProspectUpdateOwner.Execute(SQLProspectUpdateOwner)
			If Not rsProspectUpdateOwner.EOF Then OriginalOwnerUserNo = rsProspectUpdateOwner("OwnerUserNo")
			
			SQLProspectUpdateOwner = "UPDATE PR_Prospects Set OwnerUserNo = " & Session("UserNo") & " WHERE InternalRecordIdentifier = " & passedProspectID
			
			Set rsProspectUpdateOwner = cnnProspectUpdateOwner.Execute(SQLProspectUpdateOwner)
			
			set rsProspectUpdateOwner = Nothing
			cnnProspectUpdateOwner.Close
			set cnnProspectUpdateOwner = Nothing
			
			Select Case passedPageSource
				Case "R"
					Description = GetUserDisplayNameByUserNo(Session("UserNo")) & " was assigned temporary ownership of the prospect " & GetProspectNameByNumber(passedProspectID) 
					Description = Description & " when moving the prospect from the " & GetTerm("recycle pool") & " back to the main prospect pool. "
				Case "E"
					Description = GetUserDisplayNameByUserNo(Session("UserNo")) & " was assigned temporary ownership of the prospect " & GetProspectNameByNumber(passedProspectID) 
					Description = Description & " while changing the prospect owner to " &  GetUserDisplayNameByUserNo(passedNewOwnerUserNo) 
					Description = Description & " from the Edit Prospect screen."
				Case "A"
					Description = GetUserDisplayNameByUserNo(Session("UserNo")) & " was assigned temporary ownership of the prospect " & GetProspectNameByNumber(passedProspectID) 
					Description = Description & " while creating the prospect via the Add Prospect screen."
			End Select
			
			Description = Description & " An email was sent to " & GetUserDisplayNameByUserNo(passedNewOwnerUserNo) & " assgining the prospect to them and requesting that they accept ownership. "
			Description = Description & " Once " & GetUserDisplayNameByUserNo(passedNewOwnerUserNo) & " clicks the ACCEPT link, ownership will be transferred from "
			Description = Description & GetUserDisplayNameByUserNo(Session("UserNo")) & " to " & GetUserDisplayNameByUserNo(passedNewOwnerUserNo) 
			Description = Description & " No Outlook calendar entries will be made until " & GetUserDisplayNameByUserNo(passedNewOwnerUserNo) & " accepts ownership via the email."
			
			CreateAuditLogEntry GetTerm("Prospecting"),"Prospect owner changed/set","Major",0,Description
			Record_PR_Activity passedProspectID,Description,Session("UserNo")
			
			UserNoForCalendarUpdate = 0
				
		End If	
		
	Else
	
			'***************************************
			'Update owner in PR_Prospects
			'***************************************
	
			Set cnnProspectUpdateOwner = Server.CreateObject("ADODB.Connection")
			cnnProspectUpdateOwner.open Session("ClientCnnString")
			Set rsProspectUpdateOwner = Server.CreateObject("ADODB.Recordset")
			rsProspectUpdateOwner.CursorLocation = 3 

			'Get current owner
			OriginalOwnerUserNo = 0
			SQLProspectUpdateOwner = "Select OwnerUserNo FROM PR_Prospects WHERE InternalRecordIdentifier = " & passedProspectID
			Set rsProspectUpdateOwner = cnnProspectUpdateOwner.Execute(SQLProspectUpdateOwner)
			If Not rsProspectUpdateOwner.EOF Then OriginalOwnerUserNo = rsProspectUpdateOwner("OwnerUserNo")
			
			SQLProspectUpdateOwner = "UPDATE PR_Prospects Set OwnerUserNo = " & Session("UserNo") & " WHERE InternalRecordIdentifier = " & passedProspectID
			
			Set rsProspectUpdateOwner = cnnProspectUpdateOwner.Execute(SQLProspectUpdateOwner)
			
			set rsProspectUpdateOwner = Nothing
			cnnProspectUpdateOwner.Close
			set cnnProspectUpdateOwner = Nothing
			
			Select Case passedPageSource
				Case "R"
					Description = GetUserDisplayNameByUserNo(Session("UserNo")) & " took ownership of the prospect " & GetProspectNameByNumber(passedProspectID) 
					Description = Description & " when moving the prospect from the " & GetTerm("recycle pool") & " back to the main prospect pool. "
				Case "E"
					Description = GetUserDisplayNameByUserNo(Session("UserNo")) & " took ownership of the prospect " & GetProspectNameByNumber(passedProspectID) 
					Description = Description & " from the Edit Prospect screen."
				Case "A"
					Description = GetUserDisplayNameByUserNo(Session("UserNo")) & " assigned themself ownership of the prospect " & GetProspectNameByNumber(passedProspectID) 
					Description = Description & " while creating the prospect via the Add Prospect screen."
			End Select
			
			Description = Description & " If the Next Activity for this prospect is a meeting or appointment, and entry will be created in the Outlook calendar belonging to " & GetUserDisplayNameByUserNo(Session("UserNo")) 
						
			CreateAuditLogEntry GetTerm("Prospecting"),"Prospect owner changed/set","Major",0,Description
			Record_PR_Activity passedProspectID,Description,Session("UserNo")
			
			UserNoForCalendarUpdate = Session("UserNO")

	End If


	'******************************************************************************************************************************************
	'******************************************************************************************************************************************
	'******************************************************************************************************************************************
	'******************************************************************************************************************************************		
	'' Now see if we need to create an appointment or meeting in the users email system
	'******************************************************************************************************************************************
	'******************************************************************************************************************************************
	'******************************************************************************************************************************************
	'******************************************************************************************************************************************
	'baseURL should alwats have a trailing /slash, just in case, handle either way
	If right(baseURL,1)="/" Then maildomain = Left(right(baseURL,len(baseURL)-7),len(right(baseURL,len(baseURL)-7))-1) Else maildomain = right(baseURL,len(baseURL)-7)


	If UserNoForCalendarUpdate <> 0 Then
	
		ProspectApptOrMeeting = GetActivityApptOrMeetingByNum(GetCurrentProspectActivityNumberByProspectNumber(passedProspectID)) 
		
		If ProspectApptOrMeeting = "Appointment" or ProspectApptOrMeeting = "Meeting" Then
		
			'OK, see if we have credentials for this user
	
			If GetUserEmailSystemIDByUserNo(UserNoForCalendarUpdate) <> "" AND GetUserEmailSystemPassByUserNo((UserNoForCalendarUpdate)) <> "" Then
	
					'OK, see if we allow access to calendar for this user
	
					If AllowUpdatesToUsersCalendar(UserNoForCalendarUpdate) = True Then
					
						TARGETURL = GetPOSTParams("EWSPostURL")
						USERNAME = GetUserEmailSystemIDByUserNo(UserNoForCalendarUpdate)
						PASSWORD = GetUserEmailSystemPassByUserNo(UserNoForCalendarUpdate)
					
						If Instr(ucase(Request.ServerVariables("SERVER_NAME")),"DEV") <> 0 Then
							USERNAME = "minsight@corpcofe.com"
							PASSWORD = "minsight"
						End If
						
						' Lookup the Activity & get all the stuff we are going to neeed
						Set cnnGetCurrentProspectActivityDueDateByProspectNumber = Server.CreateObject("ADODB.Connection")
						cnnGetCurrentProspectActivityDueDateByProspectNumber.open Session("ClientCnnString")
		
						SQLGetCurrentProspectActivityDueDateByProspectNumber = "Select * from PR_ProspectActivities Where ProspectRecID = " & passedProspectID & " AND Status Is Null"
 
						Set rsGetCurrentProspectActivityDueDateByProspectNumber = Server.CreateObject("ADODB.Recordset")
						rsGetCurrentProspectActivityDueDateByProspectNumber.CursorLocation = 3 
						Set rsGetCurrentProspectActivityDueDateByProspectNumber = cnnGetCurrentProspectActivityDueDateByProspectNumber.Execute(SQLGetCurrentProspectActivityDueDateByProspectNumber)
			 
						If not rsGetCurrentProspectActivityDueDateByProspectNumber.EOF Then
							ActivityRecID = rsGetCurrentProspectActivityDueDateByProspectNumber("ActivityRecID")
							ActivityDate =  rsGetCurrentProspectActivityDueDateByProspectNumber("ActivityDueDate")
							ApptDuration =  rsGetCurrentProspectActivityDueDateByProspectNumber("ActivityAppointmentDuration")
							MeetDuration =  rsGetCurrentProspectActivityDueDateByProspectNumber("ActivityMeetingDuration")
							MeetLocation =  rsGetCurrentProspectActivityDueDateByProspectNumber("ActivityMeetingLocation")
							ActivityNotes =  rsGetCurrentProspectActivityDueDateByProspectNumber("Notes")
						End If
	
						rsGetCurrentProspectActivityDueDateByProspectNumber.Close
						set rsGetCurrentProspectActivityDueDateByProspectNumber= Nothing
						cnnGetCurrentProspectActivityDueDateByProspectNumber.Close	
						set cnnGetCurrentProspectActivityDueDateByProspectNumber= Nothing
					
						' Convert from Coordinated Universal Time to EDT
'						ActivityDate = dateAdd("h",4,ActivityDate)
						ActivityDate = dateAdd("h",5,ActivityDate)
						
						ApptOrMeetStartDateTime  = Year(ActivityDate) & "-" 
						If Month(ActivityDate) < 10 Then
							ApptOrMeetStartDateTime  = ApptOrMeetStartDateTime  & "0" & Month(ActivityDate)
						Else
							ApptOrMeetStartDateTime  = ApptOrMeetStartDateTime  & Month(ActivityDate)
						End If
						ApptOrMeetStartDateTime  = ApptOrMeetStartDateTime  & "-"
						If Day(ActivityDate) < 10 Then
							ApptOrMeetStartDateTime  = ApptOrMeetStartDateTime  & "0" & Day(ActivityDate)
						Else
							ApptOrMeetStartDateTime  = ApptOrMeetStartDateTime  & Day(ActivityDate)
						End If
					
						ApptOrMeetStartDateTime  = ApptOrMeetStartDateTime  & "T"
						
						If Hour(ActivityDate) < 10 Then
							ApptOrMeetStartDateTime  = ApptOrMeetStartDateTime  & "0" & Hour(ActivityDate)
						Else
							ApptOrMeetStartDateTime  = ApptOrMeetStartDateTime  & Hour(ActivityDate)
						End If
						ApptOrMeetStartDateTime  = ApptOrMeetStartDateTime  & ":"
						If Minute(ActivityDate) < 10 Then
							ApptOrMeetStartDateTime  = ApptOrMeetStartDateTime  & "0" & Minute(ActivityDate)
						Else
							ApptOrMeetStartDateTime  = ApptOrMeetStartDateTime  & Minute(ActivityDate)
						End If
						ApptOrMeetStartDateTime  = ApptOrMeetStartDateTime  & ":"
						ApptOrMeetStartDateTime  = ApptOrMeetStartDateTime  & "00.000Z"
				
						If ProspectApptOrMeeting ="Appointment" Then
							Duration = ApptDuration
						Else
							Duration = MeetDuration
						End If
					
					
						If Not IsNumeric(Duration) Then Duration = 15
					
						txtProspectEditNextActivityDate2 = DateAdd("n",Duration,ActivityDate)
					
						ApptOrMeetEndDateTime  = Year(txtProspectEditNextActivityDate2) & "-" 
						If Month(txtProspectEditNextActivityDate2) < 10 Then
							ApptOrMeetEndDateTime  = ApptOrMeetEndDateTime  & "0" & Month(txtProspectEditNextActivityDate2)
						Else
							ApptOrMeetEndDateTime  = ApptOrMeetEndDateTime  & Month(txtProspectEditNextActivityDate2)
						End If
						ApptOrMeetEndDateTime  = ApptOrMeetEndDateTime  & "-"
						If Day(txtProspectEditNextActivityDate2) < 10 Then
							ApptOrMeetEndDateTime  = ApptOrMeetEndDateTime  & "0" & Day(txtProspectEditNextActivityDate2)
						Else
							ApptOrMeetEndDateTime  = ApptOrMeetEndDateTime  & Day(txtProspectEditNextActivityDate2)
						End If
						ApptOrMeetEndDateTime  = ApptOrMeetEndDateTime  & "T"
						If Hour(txtProspectEditNextActivityDate2) < 10 Then
							ApptOrMeetEndDateTime  = ApptOrMeetEndDateTime  & "0" & Hour(txtProspectEditNextActivityDate2)
						Else
							ApptOrMeetEndDateTime  = ApptOrMeetEndDateTime  & Hour(txtProspectEditNextActivityDate2)
						End If
						ApptOrMeetEndDateTime  = ApptOrMeetEndDateTime  & ":"
						If Minute(txtProspectEditNextActivityDate2) < 10 Then
							ApptOrMeetEndDateTime  = ApptOrMeetEndDateTime  & "0" & Minute(txtProspectEditNextActivityDate2)
						Else
							ApptOrMeetEndDateTime  = ApptOrMeetEndDateTime  & Minute(txtProspectEditNextActivityDate2)
						End If
						ApptOrMeetEndDateTime  = ApptOrMeetEndDateTime  & ":"
						ApptOrMeetEndDateTime  = ApptOrMeetEndDateTime  & "00.000Z"
						
						Select Case ProspectApptOrMeeting 
					
							Case "Appointment"
	
								Call GetExtraInfo (passedProspectID)
	
								reqStr = ""
								reqStr = reqStr & "<?xml version='1.0' encoding='utf-8'?>"
								reqStr = reqStr & "<soap:Envelope xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance'" 
								reqStr = reqStr & "       xmlns:m='http://schemas.microsoft.com/exchange/services/2006/messages'"
								reqStr = reqStr & "       xmlns:t='http://schemas.microsoft.com/exchange/services/2006/types'"
								reqStr = reqStr & "       xmlns:soap='http://schemas.xmlsoap.org/soap/envelope/'>"
								reqStr = reqStr & "  <soap:Header>"
								reqStr = reqStr & "    <t:RequestServerVersion Version='Exchange2007_SP1' />"
								reqStr = reqStr & "    <t:TimeZoneContext>"
								reqStr = reqStr & "      <t:TimeZoneDefinition Id='Eastern Standard Time' />"
								reqStr = reqStr & "    </t:TimeZoneContext>"
								reqStr = reqStr & "  </soap:Header>"
								reqStr = reqStr & "  <soap:Body>"
								reqStr = reqStr & "    <m:CreateItem SendMeetingInvitations='SendToNone'>"
								reqStr = reqStr & "      <m:Items>"
								reqStr = reqStr & "        <t:CalendarItem>"
								reqStr = reqStr & "          <t:Subject>" & Replace(GetProspectNameByNumber(passedProspectID),"&","&amp;") & "</t:Subject>"
								'***********************************
								'Additional details per Adam Henchel
								'***********************************			
								reqStr = reqStr & "          <t:Body BodyType='HTML'>" & "<![CDATA[" & GetActivityByNum(ActivityRecID)& "<BR> " & ActivityNotes & "<BR>"
								reqStr = reqStr & "Contact Name: " & ExtraFirstName & " " & ExtraLastName & "<BR>"
								reqStr = reqStr & "Phone: " &  ExtraPhone
								If ExtraPhoneExt<> "" Then 
									reqStr = reqStr & "Ext: " &  ExtraPhoneExt & "<BR>"
								Else
									reqStr = reqStr & "<BR>"
								End If
								reqStr = reqStr & "Email: " & ExtraEmail & "<BR>"
								reqStr = reqStr & "Address:" & "<BR>"
								reqStr = reqStr & ExtraStreet & "<BR>"
								If ExtraAddress2 <> "" Then reqStr = reqStr & ExtraAddress2 & "<BR>"
								reqStr = reqStr & ExtraCity & " , " &  ExtraState & " " & ExtraPostalCode	& "<BR>"
								reqStr = reqStr & "Primary Competitor: " & ExtraPrimaryCompetitorName  & "<BR>"
								reqStr = reqStr & "]]></t:Body>"
								'***************************************
								'END Additional details per Adam Henchel
								'***************************************	
								'reqStr = reqStr & "          <t:ReminderDueBy>" & ReminderDueByDateTime & "</t:ReminderDueBy>"
								reqStr = reqStr & "          <t:Start>" & ApptOrMeetStartDateTime & "</t:Start>"
								reqStr = reqStr & "          <t:End>" & ApptOrMeetEndDateTime & "</t:End>"
								reqStr = reqStr & "          <t:Location>" & Replace(GetProspectNameByNumber(passedProspectID),"&","&amp;") & "</t:Location>"
								reqStr = reqStr & "          <t:MeetingTimeZone TimeZoneName='Eastern Standard Time' />"
								reqStr = reqStr & "        </t:CalendarItem>"
								reqStr = reqStr & "      </m:Items>"
								reqStr = reqStr & "    </m:CreateItem>"
								reqStr = reqStr & "  </soap:Body>"
								reqStr = reqStr & "</soap:Envelope>"
	
							Case "Meeting"
									
								Call GetExtraInfo (passedProspectID)
									
								If MeetLocation = "" Then
								
									'Meeting so need to get location info, if blank
								
									Set cnntmpProspect = Server.CreateObject("ADODB.Connection")
									cnntmpProspect.open Session("ClientCnnString")
				
									SQLtmpProspect = "Select * from PR_Prospects Where InternalRecordIdentifier = " & PassedProspectID
			 
									Set rstmpProspect = Server.CreateObject("ADODB.Recordset")
									rstmpProspect.CursorLocation = 3 
									Set rstmpProspect = cnntmpProspect.Execute(SQLtmpProspect)
					
									If not rstmpProspect.EOF Then
								
										Street = rstmpProspect("Street")
										City = rstmpProspect("City")
										Floor_Suite_Room__c = rstmpProspect("Floor_Suite_Room__c")		
									
										If Street <> "" and Not IsNull(Street) Then MeetLocation = Street & Chr(13)
										If City <> "" and Not IsNull(City) Then MeetLocation = MeetLocation & City & Chr(13)
										If Floor_Suite_Room__c <> "" and Not IsNull(Floor_Suite_Room__c) Then MeetLocation = MeetLocation & Floor_Suite_Room__c
									
									End If
								
									rstmpProspect.Close
									set rstmpProspect= Nothing
									cnntmpProspect.Close	
									set cnntmpProspect= Nothing
								End If
								If MeetLocation = "" Then MeetLocation = GetProspectNameByNumber(passedProspectID)
					
								reqStr = ""
								reqStr = reqStr & "<?xml version='1.0' encoding='utf-8'?>"
								reqStr = reqStr & "<soap:Envelope xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns:m='http://schemas.microsoft.com/exchange/services/2006/messages'" 
								reqStr = reqStr & "       xmlns:t='http://schemas.microsoft.com/exchange/services/2006/types' xmlns:soap='http://schemas.xmlsoap.org/soap/envelope/'>"
								reqStr = reqStr & "  <soap:Header>"
								reqStr = reqStr & "    <t:RequestServerVersion Version='Exchange2007_SP1' />"
								reqStr = reqStr & "    <t:TimeZoneContext>"
								reqStr = reqStr & "      <t:TimeZoneDefinition Id='Eastern Standard Time' />"
								reqStr = reqStr & "    </t:TimeZoneContext>"
								reqStr = reqStr & "  </soap:Header>"
								reqStr = reqStr & "  <soap:Body>"
								reqStr = reqStr & "    <m:CreateItem SendMeetingInvitations='SendToAllAndSaveCopy'>"
								reqStr = reqStr & "      <m:Items>"
								reqStr = reqStr & "        <t:CalendarItem>"
								reqStr = reqStr & "          <t:Subject>" & Replace(GetProspectNameByNumber(passedProspectID),"&","&amp;") & "</t:Subject>"
								'***********************************
								'Additional details per Adam Henchel
								'***********************************			
								reqStr = reqStr & "          <t:Body BodyType='HTML'>" & "<![CDATA[" & GetActivityByNum(ActivityRecID) & "<BR> " & ActivityNotes & "<BR>"
								reqStr = reqStr & "Contact Name: " & ExtraFirstName & " " & ExtraLastName & "<BR>"
								reqStr = reqStr & "Phone: " &  ExtraPhone
								If ExtraPhoneExt<> "" Then 
									reqStr = reqStr & "Ext: " &  ExtraPhoneExt & "<BR>"
								Else
									reqStr = reqStr & "<BR>"
								End If
								reqStr = reqStr & "Email: " & ExtraEmail & "<BR>"
								reqStr = reqStr & "Address:" & "<BR>"
								reqStr = reqStr & ExtraStreet & "<BR>"
								If ExtraAddress2 <> "" Then reqStr = reqStr & ExtraAddress2 & "<BR>"
								reqStr = reqStr & ExtraCity & " , " &  ExtraState & " " & ExtraPostalCode	& "<BR>"
								reqStr = reqStr & "Primary Competitor: " & ExtraPrimaryCompetitorName  & "<BR>"
								reqStr = reqStr & "]]></t:Body>"
								'***************************************
								'END Additional details per Adam Henchel
								'***************************************	
								reqStr = reqStr & "          <t:ReminderMinutesBeforeStart>60</t:ReminderMinutesBeforeStart>"
								reqStr = reqStr & "          <t:Start>" & ApptOrMeetStartDateTime & "</t:Start>"
								reqStr = reqStr & "          <t:End>" & ApptOrMeetEndDateTime  & "</t:End>"
								reqStr = reqStr & "          <t:Location>" & Replace(MeetLocation,"&","&amp;") & "</t:Location>"
								reqStr = reqStr & "          <t:RequiredAttendees>"
								reqStr = reqStr & "            <t:Attendee>"
								reqStr = reqStr & "              <t:Mailbox>"
								If Instr(ucase(Request.ServerVariables("SERVER_NAME")),"DEV") <> 0 Then
									reqStr = reqStr & "                <t:EmailAddress>" & "minsight@corpcofe.com" & "</t:EmailAddress>"
								Else
									reqStr = reqStr & "                <t:EmailAddress>" & getUserEmailAddress(UserNoForCalendarUpdate) & "</t:EmailAddress>"
								End If							
								reqStr = reqStr & "              </t:Mailbox>"
								reqStr = reqStr & "            </t:Attendee>"
								reqStr = reqStr & "            <t:Attendee>"
								reqStr = reqStr & "              <t:Mailbox>"
								reqStr = reqStr & "                <t:EmailAddress>" & "rsmith@ocsaccess.com" & "</t:EmailAddress>"
								reqStr = reqStr & "              </t:Mailbox>"
								reqStr = reqStr & "            </t:Attendee>"
								reqStr = reqStr & "          </t:RequiredAttendees>"
								reqStr = reqStr & "          <t:MeetingTimeZone TimeZoneName='Eastern Standard Time' />"
								reqStr = reqStr & "        </t:CalendarItem>"
								reqStr = reqStr & "      </m:Items>"
								reqStr = reqStr & "    </m:CreateItem>"
								reqStr = reqStr & "  </soap:Body>"
								reqStr = reqStr & "</soap:Envelope>"	
	
						End Select
					
						'Perform the actual post
						set oXMLHTTP=CreateObject("MSXML2.XMLHTTP")
						set oXML=CreateObject("MSXML2.DOMDocument")
			
						' Send the request
						oXMLHTTP.Open "POST", TARGETURL, false, USERNAME, PASSWORD
						oXMLHTTP.SetRequestHeader "Content-Type", "text/xml"
						oXMLHTTP.Send reqStr 
		
						Stat = oXMLHTTP.status
		
						If Stat = "200" THEN 
					
							If Instr(oXMLHTTP.responseText,"NoError") <> 0 Then ' Success
						
								Description ="success! oXMLHTTP.status returned " & oXMLHTTP.status & " when posting via EWS to Exchange"
								Description = "oXMLHTTP.responseText:" & oXMLHTTP.responseText & "<br>"
								Description = Description & "PAGE: " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "<br>"
								Description = Description & "Posted to " & TARGETURL  & "<br>"
								Description = Description & "POSTED DATA:" & reqStr & "<br>"
								Description = Description & "SERNO:" & MUV_READ("ClientID") & "<br>"
							
								CreateINSIGHTAuditLogEntry Request.ServerVariables("SERVER_NAME"),Description,"n/a"
								
								'****************************
								'TEMPORARY CODE TO BE REMOVED
								'****************************
								emailbody="oXMLHTTP.status returned " & oXMLHTTP.status & " when posting via EWS to Exchange <br>"
								emailBody = emailBody & "oXMLHTTP.responseText:" & oXMLHTTP.responseText & "<br>"
								emailBody = emailBody & "PAGE: " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "<br>"
								emailBody = emailBody & "Posted to " & TARGETURL   & "<br>"
								emailBody = emailBody & "POSTED DATA:" & reqStr & "<br>"
								emailBody = emailBody & "SERNO:" & MUV_READ("ClientID") & "<br>"
								emailBody = emailBody & "USER:" & USERNAME & "<br>"
								emailBody = emailBody & "PWD:" & PASSWORD & "<br>"
								SendMail "mailsender@" & maildomain ,"rsmith@ocsaccess.com",MUV_READ("ClientID") & " EWS POST OK",emailBody, "Prospecting", "Post OK"
								'********************************
								'END TEMPORARY CODE TO BE REMOVED
								'********************************


							
								If ProspectApptOrMeeting ="Appointment" Then 
									Description = "An appointment" 
								Else 
									Description = "A meeting"
								End If
							
								Description = Description & " was created in the Outlook calendar for: " & GetUserDisplayNameByUserNo(UserNoForCalendarUpdate) & " for prospect " & GetProspectNameByNumber(passedProspectID) & " for the Activity " & GetActivityByNum(ActivityRecID)  & " for the date " & ActivityDate 
								
								CreateAuditLogEntry GetTerm("Prospecting") , ProspectApptOrMeeting & " created in calender","Major",0,Description
							
								If ProspectApptOrMeeting ="Appointment" Then 
									AppointmentDuration =  round(cint(ApptDuration)/60,2) & " hours"
									
									If ApptDuration mod 60 = 0 Then 
										AppointmentDuration =  round(cint(ApptDuration)/60,2) & " hour(s)"
									Else
										If ApptDuration < 60 Then
											AppointmentDuration =  ApptDuration & " minutes"
										Else
											AppointmentDuration = round(cint(ApptDuration)/60,2) & " hour(s) "
											AppointmentDuration = AppointmentDuration & ApptDuration mod 60 & " minutes"
										End If
									End If
									
									Description = "An appointment was created in the Outlook calendar for: " & GetUserDisplayNameByUserNo(UserNoForCalendarUpdate) & " for the Activity <strong><em>" & GetActivityByNum(ActivityRecID)  & "</em></strong> for, <strong><em>" & ActivityDate  & "</em></strong>, with a duration of <strong><em>" & AppointmentDuration & "</em></strong>."
								Else 
									MeetingDuration =  round(cint(MeetDuration)/60,2) & " hours"
									
									If MeetDuration mod 60 = 0 Then 
										MeetingDuration =  round(cint(MeetDuration)/60,2) & " hour(s)"
									Else
										If MeetDuration < 60 Then
											MeetingDuration =  MeetDuration & " minutes"
										Else
											MeetingDuration = round(cint(MeetDuration)/60,2) & " hour(s) "
											MeetingDuration = MeetingDuration& MeetDuration mod 60 & " minutes"
										End If
									End If

									Description = "An meeting was created in the Outlook calendar for: " & GetUserDisplayNameByUserNo(UserNoForCalendarUpdate) & " for the Activity <strong><em>" & GetActivityByNum(ActivityRecID)  & "</em></strong> for, <strong><em>" & ActivityDate  & "</em></strong>, with a duration of <strong><em>" & MeetingDuration & "</em></strong> at location: <strong><em>" & MeetLocation & "</em></strong>."	
								End If
							
								Record_PR_Activity passedProspectID,Description,Session("UserNo")
	
								Response.Write("OK")
							Else
								'FAILURE
								emailbody="oXMLHTTP.status returned " & oXMLHTTP.status & " when posting via EWS to Exchange <br>"
								emailBody = emailBody & "oXMLHTTP.responseText:" & oXMLHTTP.responseText & "<br>"
								emailBody = emailBody & "PAGE: " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "<br>"
								emailBody = emailBody & "Posted to " & TARGETURL   & "<br>"
								emailBody = emailBody & "POSTED DATA:" & reqStr & "<br>"
								emailBody = emailBody & "SERNO:" & MUV_READ("ClientID") & "<br>"
								emailBody = emailBody & "USER:" & USERNAME & "<br>"
								emailBody = emailBody & "PWD:" & PASSWORD & "<br>"
								SendMail "mailsender@" & maildomain ,"rsmith@ocsaccess.com",MUV_READ("ClientID") & " EWS POST ERROR",emailBody, "Prospecting", "Post Failure"
						
								Description = emailBody 
								CreateINSIGHTAuditLogEntry Request.ServerVariables("SERVER_NAME"),Description,"n/a"
								
								Response.Write("BAD")
							End If
						
						Else ' not a 200 from server
					
					'		'FAILURE
							emailbody="oXMLHTTP.status returned " & oXMLHTTP.status & " when posting via EWS to Exchange (Non 200 Status) <br>"
							emailBody = emailBody & "oXMLHTTP.responseText:" & oXMLHTTP.responseText & "<br>"
							emailBody = emailBody & "PAGE: " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "<br>"
							emailBody = emailBody & "Posted to " & TARGETURL   & "<br>"
							emailBody = emailBody & "POSTED DATA:" & reqStr & "<br>"
							emailBody = emailBody & "SERNO:" & MUV_READ("ClientID") & "<br>"
							emailBody = emailBody & "USER:" & USERNAME & "<br>"
							emailBody = emailBody & "PWD:" & PASSWORD & "<br>"

							SendMail "mailsender@" & maildomain ,"rsmith@ocsaccess.com",MUV_READ("ClientID") & " EWS POST ERROR",emailBody, "Prospecting", "Post Failure"
						
							Description = emailBody 
							CreateINSIGHTAuditLogEntry Request.ServerVariables("SERVER_NAME"),Description,"n/a"
			
						End If
					
						set oXMLHTTP=Nothing
						set oXML=Nothing
	
					End If
				
				End If
			
			End If
		
	End If	

	'******************************************************************************************************************************************
	'******************************************************************************************************************************************
	'******************************************************************************************************************************************
	'******************************************************************************************************************************************		
	'' 'DETERMINE IF OWNWERSHIP EMAIL NEEDS TO BE SENT
	'******************************************************************************************************************************************
	'******************************************************************************************************************************************
	'******************************************************************************************************************************************
	'******************************************************************************************************************************************
	If DoNotSendEmail <> True Then
	
		If cint(Session("UserNo")) <> cint(passedNewOwnerUserNo)  Then

			SQLClientID = "SELECT * FROM tblServerInfo where clientKey='"& MUV_READ("ClientID") &"'"
			
			Set ConnectionClientID  = Server.CreateObject("ADODB.Connection")
			Set RecordsetClientID  = Server.CreateObject("ADODB.Recordset")
			
			ConnectionClientID.Open InsightCnnString
			
			'Open the recordset object executing the SQL statement and return records
			RecordsetClientID.Open SQLClientID,ConnectionClientID,3,3
			
			'First lookup the ClientKey in tblServerInfo
			If RecordsetClientID.recordcount > 0 then
				userQuickLoginURL = RecordsetClientID.Fields("QuickLoginURL")
			End If
			
		
			Set ConnectionUsers= Server.CreateObject("ADODB.Connection")
			Set rsCurrProspectInfo2 = Server.CreateObject("ADODB.Recordset")
			ConnectionUsers.Open Session("ClientCnnString")
		
			'declare the SQL statement that will query the database
			SQL = "SELECT * FROM tblUsers WHERE userNo = " & cint(passedNewOwnerUserNo)
		
			'Open the recordset object executing the SQL statement and return records
			Set rsCurrProspectInfo2 = ConnectionUsers.Execute(SQL)
		
			'If there is no record with the entered username, close connection
			If rsCurrProspectInfo2.EOF then
				rsCurrProspectInfo2.close
				ConnectionUsers.close
				set rsCurrProspectInfo2 = nothing
				set ConnectionUsers= nothing
			Else		
				userEmail = rsCurrProspectInfo2("userEmail")
				userFirstName = rsCurrProspectInfo2("userFirstName")
				userLastName = rsCurrProspectInfo2("userLastName")
				userDisplayName = rsCurrProspectInfo2("userDisplayName")		
		
				SQLNextActivity = "SELECT * FROM PR_ProspectActivities where ProspectRecID = " & passedProspectID & " AND Status IS NULL"
				
				Set cnnNextActivity = Server.CreateObject("ADODB.Connection")
				cnnNextActivity.open (Session("ClientCnnString"))
				Set rsNextActivity = Server.CreateObject("ADODB.Recordset")
				rsNextActivity.CursorLocation = 3 
				Set rsNextActivity = cnnNextActivity.Execute(SQLNextActivity)
				
				If not rsNextActivity.EOF Then
				
				  	NextActivityRecID = rsNextActivity("ActivityRecID")
				  	NextActivity = GetActivityByNum(rsNextActivity("ActivityRecID"))
					NextActivityDueDate = FormatDateTime(rsNextActivity("ActivityDueDate"),2) & " " & FormatDateTime(rsNextActivity("ActivityDueDate"),3)
					daysOld = DateDiff("d",rsNextActivity("RecordCreationDateTime"),Now())
					daysOverdue = DateDiff("d",rsNextActivity("ActivityDueDate"),Now())
					
					ProspectApptOrMeeting = GetActivityApptOrMeetingByNum(GetCurrentProspectActivityNumberByProspectNumber(passedProspectID)) 
					
					If ProspectApptOrMeeting <> "" Then
					
						If ProspectApptOrMeeting = "Appointment" Then
						
							Duration = rsNextActivity("ActivityAppointmentDuration")
							Location = ""
			
						ElseIf ProspectApptOrMeeting = "Meeting" Then
						
							Duration = rsNextActivity("ActivityMeetingDuration")
							Location = rsNextActivity("ActivityMeetingLocation")
							
						Else
						
							Duration = ""
							Location = ""
						
						End If
					End If				
				End If
				Set rsNextActivity = Nothing
				cnnNextActivity.Close
				Set cnnNextActivity = Nothing
				
				
				SQLCurrProspectInfo = "SELECT * FROM PR_ProspectContacts WHERE ProspectIntRecID = " & passedProspectID & " AND PrimaryContact = 1"
				
				Set cnnCurrProspectInfo = Server.CreateObject("ADODB.Connection")
				cnnCurrProspectInfo.open (Session("ClientCnnString"))
				Set rsCurrProspectInfo = Server.CreateObject("ADODB.Recordset")
				rsCurrProspectInfo.CursorLocation = 3 
				Set rsCurrProspectInfo = cnnCurrProspectInfo.Execute(SQLCurrProspectInfo)
				
				If not rsCurrProspectInfo.EOF Then
				  	FirstName = rsCurrProspectInfo("FirstName")
				  	LastName = rsCurrProspectInfo("LastName") 
				End If
				Set rsCurrProspectInfo = Nothing
				cnnCurrProspectInfo.Close
				Set cnnCurrProspectInfo = Nothing
				
				
				SQLCurrProspectInfo2 = "SELECT * FROM PR_Prospects WHERE InternalRecordIdentifier = " & passedProspectID 
				
				Set cnnCurrProspectInfo2 = Server.CreateObject("ADODB.Connection")
				cnnCurrProspectInfo2.open (Session("ClientCnnString"))
				Set rsCurrProspectInfo2 = Server.CreateObject("ADODB.Recordset")
				rsCurrProspectInfo2.CursorLocation = 3 
				Set rsCurrProspectInfo2 = cnnCurrProspectInfo2.Execute(SQLCurrProspectInfo2)
				
				If not rsCurrProspectInfo2.EOF Then
					Company = rsCurrProspectInfo2("Company")
					Street = rsCurrProspectInfo2("Street")
					Address2 = rsCurrProspectInfo2("Floor_Suite_Room__c")
					City= rsCurrProspectInfo2("City")
					State= rsCurrProspectInfo2("State")
					PostalCode = rsCurrProspectInfo2("PostalCode")
					LeadSourceNumber = rsCurrProspectInfo2("LeadSourceNumber")
					LeadSource = GetLeadSourceByNum(LeadSourceNumber)				
					StageNumber = GetProspectCurrentStageByProspectNumber(passedProspectID)
					IndustryNumber = rsCurrProspectInfo2("IndustryNumber")	
					Industry = GetIndustryByNum(IndustryNumber)											
					OwnerUserNo = rsCurrProspectInfo2("OwnerUserNo")				
					CreatedDate= rsCurrProspectInfo2("CreatedDate")
					CreatedByUserNo= rsCurrProspectInfo2("CreatedByUserNo")				
					TelemarketerUserNo = rsCurrProspectInfo2("TelemarketerUserNo")
					Telemarketer = GetUserDisplayNameByUserNo(TelemarketerUserNo)
					ProjectedGPSpend= rsCurrProspectInfo2("ProjectedGPSpend")
					NumberOfPantries = rsCurrProspectInfo2("NumberOfPantries")
					EmployeeRangeNumber = rsCurrProspectInfo2("EmployeeRangeNumber")
					NumEmployees = GetEmployeeRangeByNum(EmployeeRangeNumber)
					CreatedDate = rsCurrProspectInfo2("CreatedDate")
					FormerCustNum = rsCurrProspectInfo2("FormerCustNum")
					CancelDate = rsCurrProspectInfo2("CancelDate")
					LeaseExpirationDate = rsCurrProspectInfo2("LeaseExpirationDate")	
					ContractExpirationDate = rsCurrProspectInfo2("ContractExpirationDate")
					Comments = rsCurrProspectInfo2("Comments")
					CurrentOffering = rsCurrProspectInfo2("CurrentOffering")		
				End If
				Set rsCurrProspectInfo2 = Nothing
				cnnCurrProspectInfo2.Close
				Set cnnCurrProspectInfo2 = Nothing
				
				PrimaryCompetitorID = GetPrimaryCompetitorIDByProspectNumber(passedProspectID)
			
				If PrimaryCompetitorID <> "" Then
					PrimaryCompetitorName = GetCompetitorByNum(PrimaryCompetitorID)
				Else
					PrimaryCompetitorName = "None Entered"
				End If
						
				%><!--#include file="../emails/prospecting_owner_request_email.asp"--><%
					
				SendMail "mailsender@" & maildomain,userEmail,emailSubject,emailBody, GetTerm("Prospecting"), GetTerm("Prospecting") & " Prospect Owner Request"
				
				ConnectionUsers.close	
				
			End If	
		
		End If

End If

End Function



Function Prospect_Email_Accept (passedProspectID,passedNewOwnerUserNo)


	UserNoForCalendarUpdate = passedNewOwnerUserNo
			

	'******************************************************************************************************************************************
	'******************************************************************************************************************************************
	'******************************************************************************************************************************************
	'******************************************************************************************************************************************		
	'' Now see if we need to create an appointment or meeting in the users email system
	'******************************************************************************************************************************************
	'******************************************************************************************************************************************
	'******************************************************************************************************************************************
	'******************************************************************************************************************************************
	'baseURL should alwats have a trailing /slash, just in case, handle either way
	If right(baseURL,1)="/" Then maildomain = Left(right(baseURL,len(baseURL)-7),len(right(baseURL,len(baseURL)-7))-1) Else maildomain = right(baseURL,len(baseURL)-7)


	If UserNoForCalendarUpdate <> 0 Then
	
		ProspectApptOrMeeting = GetActivityApptOrMeetingByNum(GetCurrentProspectActivityNumberByProspectNumber(passedProspectID)) 
		
		If ProspectApptOrMeeting = "Appointment" or ProspectApptOrMeeting = "Meeting" Then
		
			'OK, see if we have credentials for this user
	
			If GetUserEmailSystemIDByUserNo(UserNoForCalendarUpdate) <> "" AND GetUserEmailSystemPassByUserNo((UserNoForCalendarUpdate)) <> "" Then
	
					'OK, see if we allow access to calendar for this user
	
					If AllowUpdatesToUsersCalendar(UserNoForCalendarUpdate) = True Then
					
						TARGETURL = GetPOSTParams("EWSPostURL")
						USERNAME = GetUserEmailSystemIDByUserNo(UserNoForCalendarUpdate)
						PASSWORD = GetUserEmailSystemPassByUserNo(UserNoForCalendarUpdate)
					
						If Instr(ucase(Request.ServerVariables("SERVER_NAME")),"DEV") <> 0 Then
							USERNAME = "minsight@corpcofe.com"
							PASSWORD = "minsight"
						End If
						
						' Lookup the Activity & get all the stuff we are going to neeed
						Set cnnGetCurrentProspectActivityDueDateByProspectNumber = Server.CreateObject("ADODB.Connection")
						cnnGetCurrentProspectActivityDueDateByProspectNumber.open Session("ClientCnnString")
		
						SQLGetCurrentProspectActivityDueDateByProspectNumber = "Select * from PR_ProspectActivities Where ProspectRecID = " & passedProspectID & " AND Status Is Null"
 
						Set rsGetCurrentProspectActivityDueDateByProspectNumber = Server.CreateObject("ADODB.Recordset")
						rsGetCurrentProspectActivityDueDateByProspectNumber.CursorLocation = 3 
						Set rsGetCurrentProspectActivityDueDateByProspectNumber = cnnGetCurrentProspectActivityDueDateByProspectNumber.Execute(SQLGetCurrentProspectActivityDueDateByProspectNumber)
			 
						If not rsGetCurrentProspectActivityDueDateByProspectNumber.EOF Then
							ActivityRecID = rsGetCurrentProspectActivityDueDateByProspectNumber("ActivityRecID")
							ActivityDate =  rsGetCurrentProspectActivityDueDateByProspectNumber("ActivityDueDate")
							ApptDuration =  rsGetCurrentProspectActivityDueDateByProspectNumber("ActivityAppointmentDuration")
							MeetDuration =  rsGetCurrentProspectActivityDueDateByProspectNumber("ActivityMeetingDuration")
							MeetLocation =  rsGetCurrentProspectActivityDueDateByProspectNumber("ActivityMeetingLocation")
							ActivityNotes =  rsGetCurrentProspectActivityDueDateByProspectNumber("Notes")
						End If
	
						rsGetCurrentProspectActivityDueDateByProspectNumber.Close
						set rsGetCurrentProspectActivityDueDateByProspectNumber= Nothing
						cnnGetCurrentProspectActivityDueDateByProspectNumber.Close	
						set cnnGetCurrentProspectActivityDueDateByProspectNumber= Nothing
					
						ApptOrMeetStartDateTime  = Year(ActivityDate) & "-" 
						If Month(ActivityDate) < 10 Then
							ApptOrMeetStartDateTime  = ApptOrMeetStartDateTime  & "0" & Month(ActivityDate)
						Else
							ApptOrMeetStartDateTime  = ApptOrMeetStartDateTime  & Month(ActivityDate)
						End If
						ApptOrMeetStartDateTime  = ApptOrMeetStartDateTime  & "-"
						If Day(ActivityDate) < 10 Then
							ApptOrMeetStartDateTime  = ApptOrMeetStartDateTime  & "0" & Day(ActivityDate)
						Else
							ApptOrMeetStartDateTime  = ApptOrMeetStartDateTime  & Day(ActivityDate)
						End If
					
						ApptOrMeetStartDateTime  = ApptOrMeetStartDateTime  & "T"
						
						If Hour(ActivityDate) < 10 Then
							ApptOrMeetStartDateTime  = ApptOrMeetStartDateTime  & "0" & Hour(ActivityDate)
						Else
							ApptOrMeetStartDateTime  = ApptOrMeetStartDateTime  & Hour(ActivityDate)
						End If
						ApptOrMeetStartDateTime  = ApptOrMeetStartDateTime  & ":"
						If Minute(ActivityDate) < 10 Then
							ApptOrMeetStartDateTime  = ApptOrMeetStartDateTime  & "0" & Minute(ActivityDate)
						Else
							ApptOrMeetStartDateTime  = ApptOrMeetStartDateTime  & Minute(ActivityDate)
						End If
						ApptOrMeetStartDateTime  = ApptOrMeetStartDateTime  & ":"
						ApptOrMeetStartDateTime  = ApptOrMeetStartDateTime  & "00.000Z"
				
						If ProspectApptOrMeeting ="Appointment" Then
							Duration = ApptDuration
						Else
							Duration = MeetDuration
						End If
					
					
						If Not IsNumeric(Duration) Then Duration = 15
					
						txtProspectEditNextActivityDate2 = DateAdd("n",Duration,ActivityDate)
					
						ApptOrMeetEndDateTime  = Year(txtProspectEditNextActivityDate2) & "-" 
						If Month(txtProspectEditNextActivityDate2) < 10 Then
							ApptOrMeetEndDateTime  = ApptOrMeetEndDateTime  & "0" & Month(txtProspectEditNextActivityDate2)
						Else
							ApptOrMeetEndDateTime  = ApptOrMeetEndDateTime  & Month(txtProspectEditNextActivityDate2)
						End If
						ApptOrMeetEndDateTime  = ApptOrMeetEndDateTime  & "-"
						If Day(txtProspectEditNextActivityDate2) < 10 Then
							ApptOrMeetEndDateTime  = ApptOrMeetEndDateTime  & "0" & Day(txtProspectEditNextActivityDate2)
						Else
							ApptOrMeetEndDateTime  = ApptOrMeetEndDateTime  & Day(txtProspectEditNextActivityDate2)
						End If
						ApptOrMeetEndDateTime  = ApptOrMeetEndDateTime  & "T"
						If Hour(txtProspectEditNextActivityDate2) < 10 Then
							ApptOrMeetEndDateTime  = ApptOrMeetEndDateTime  & "0" & Hour(txtProspectEditNextActivityDate2)
						Else
							ApptOrMeetEndDateTime  = ApptOrMeetEndDateTime  & Hour(txtProspectEditNextActivityDate2)
						End If
						ApptOrMeetEndDateTime  = ApptOrMeetEndDateTime  & ":"
						If Minute(txtProspectEditNextActivityDate2) < 10 Then
							ApptOrMeetEndDateTime  = ApptOrMeetEndDateTime  & "0" & Minute(txtProspectEditNextActivityDate2)
						Else
							ApptOrMeetEndDateTime  = ApptOrMeetEndDateTime  & Minute(txtProspectEditNextActivityDate2)
						End If
						ApptOrMeetEndDateTime  = ApptOrMeetEndDateTime  & ":"
						ApptOrMeetEndDateTime  = ApptOrMeetEndDateTime  & "00.000Z"
						
						Select Case ProspectApptOrMeeting 
					
							Case "Appointment"
	
								Call GetExtraInfo (passedProspectID)
								
								reqStr = ""
								reqStr = reqStr & "<?xml version='1.0' encoding='utf-8'?>"
								reqStr = reqStr & "<soap:Envelope xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance'" 
								reqStr = reqStr & "       xmlns:m='http://schemas.microsoft.com/exchange/services/2006/messages'"
								reqStr = reqStr & "       xmlns:t='http://schemas.microsoft.com/exchange/services/2006/types'"
								reqStr = reqStr & "       xmlns:soap='http://schemas.xmlsoap.org/soap/envelope/'>"
								reqStr = reqStr & "  <soap:Header>"
								reqStr = reqStr & "    <t:RequestServerVersion Version='Exchange2007_SP1' />"
								reqStr = reqStr & "    <t:TimeZoneContext>"
								reqStr = reqStr & "      <t:TimeZoneDefinition Id='Eastern Standard Time' />"
								reqStr = reqStr & "    </t:TimeZoneContext>"
								reqStr = reqStr & "  </soap:Header>"
								reqStr = reqStr & "  <soap:Body>"
								reqStr = reqStr & "    <m:CreateItem SendMeetingInvitations='SendToNone'>"
								reqStr = reqStr & "      <m:Items>"
								reqStr = reqStr & "        <t:CalendarItem>"
								reqStr = reqStr & "          <t:Subject>" & Replace(GetProspectNameByNumber(passedProspectID),"&","&amp;") & "</t:Subject>"
								'***********************************
								'Additional details per Adam Henchel
								'***********************************			
								reqStr = reqStr & "          <t:Body BodyType='HTML'>" & "<![CDATA[" & GetActivityByNum(ActivityRecID) & "<BR> " & ActivityNotes & "<BR>"
								reqStr = reqStr & "Contact Name: " & ExtraFirstName & " " & ExtraLastName & "<BR>"
								reqStr = reqStr & "Phone: " &  ExtraPhone
								If ExtraPhoneExt<> "" Then 
									reqStr = reqStr & "Ext: " &  ExtraPhoneExt & "<BR>"
								Else
									reqStr = reqStr & "<BR>"
								End If
								reqStr = reqStr & "Email: " & ExtraEmail & "<BR>"
								reqStr = reqStr & "Address:" & "<BR>"
								reqStr = reqStr & ExtraStreet & "<BR>"
								If ExtraAddress2 <> "" Then reqStr = reqStr & ExtraAddress2 & "<BR>"
								reqStr = reqStr & ExtraCity & " , " &  ExtraState & " " & ExtraPostalCode	& "<BR>"
								reqStr = reqStr & "Primary Competitor: " & ExtraPrimaryCompetitorName  & "<BR>"
								reqStr = reqStr & "]]></t:Body>"
								'***************************************
								'END Additional details per Adam Henchel
								'***************************************
								'reqStr = reqStr & "          <t:ReminderDueBy>" & ReminderDueByDateTime & "</t:ReminderDueBy>"
								reqStr = reqStr & "          <t:Start>" & ApptOrMeetStartDateTime & "</t:Start>"
								reqStr = reqStr & "          <t:End>" & ApptOrMeetEndDateTime & "</t:End>"
								reqStr = reqStr & "          <t:Location>" & Replace(GetProspectNameByNumber(passedProspectID),"&","&amp;") & "</t:Location>"
								reqStr = reqStr & "          <t:MeetingTimeZone TimeZoneName='Eastern Standard Time' />"
								reqStr = reqStr & "        </t:CalendarItem>"
								reqStr = reqStr & "      </m:Items>"
								reqStr = reqStr & "    </m:CreateItem>"
								reqStr = reqStr & "  </soap:Body>"
								reqStr = reqStr & "</soap:Envelope>"
	
							Case "Meeting"
									
								If MeetLocation = "" Then
								
								
									'Meeting so need to get location info, if blank
								
									Set cnntmpProspect = Server.CreateObject("ADODB.Connection")
									cnntmpProspect.open Session("ClientCnnString")
				
									SQLtmpProspect = "Select * from PR_Prospects Where InternalRecordIdentifier = " & PassedProspectID
			 
									Set rstmpProspect = Server.CreateObject("ADODB.Recordset")
									rstmpProspect.CursorLocation = 3 
									Set rstmpProspect = cnntmpProspect.Execute(SQLtmpProspect)
					
									If not rstmpProspect.EOF Then
								
										Street = rstmpProspect("Street")
										City = rstmpProspect("City")
										Floor_Suite_Room__c = rstmpProspect("Floor_Suite_Room__c")		
									
										If Street <> "" and Not IsNull(Street) Then MeetLocation = Street & Chr(13)
										If City <> "" and Not IsNull(City) Then MeetLocation = MeetLocation & City & Chr(13)
										If Floor_Suite_Room__c <> "" and Not IsNull(Floor_Suite_Room__c) Then MeetLocation = MeetLocation & Floor_Suite_Room__c
									
									End If
								
									rstmpProspect.Close
									set rstmpProspect= Nothing
									cnntmpProspect.Close	
									set cnntmpProspect= Nothing
								End If
								If MeetLocation = "" Then MeetLocation = GetProspectNameByNumber(passedProspectID)
					
								Call GetExtraInfo (passedProspectID)
					
								reqStr = ""
								reqStr = reqStr & "<?xml version='1.0' encoding='utf-8'?>"
								reqStr = reqStr & "<soap:Envelope xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns:m='http://schemas.microsoft.com/exchange/services/2006/messages'" 
								reqStr = reqStr & "       xmlns:t='http://schemas.microsoft.com/exchange/services/2006/types' xmlns:soap='http://schemas.xmlsoap.org/soap/envelope/'>"
								reqStr = reqStr & "  <soap:Header>"
								reqStr = reqStr & "    <t:RequestServerVersion Version='Exchange2007_SP1' />"
								reqStr = reqStr & "    <t:TimeZoneContext>"
								reqStr = reqStr & "      <t:TimeZoneDefinition Id='Eastern Standard Time' />"
								reqStr = reqStr & "    </t:TimeZoneContext>"
								reqStr = reqStr & "  </soap:Header>"
								reqStr = reqStr & "  <soap:Body>"
								reqStr = reqStr & "    <m:CreateItem SendMeetingInvitations='SendToAllAndSaveCopy'>"
								reqStr = reqStr & "      <m:Items>"
								reqStr = reqStr & "        <t:CalendarItem>"
								reqStr = reqStr & "          <t:Subject>" & Replace(GetProspectNameByNumber(passedProspectID),"&","&amp;") & "</t:Subject>"
								'***********************************
								'Additional details per Adam Henchel
								'***********************************			
								reqStr = reqStr & "          <t:Body BodyType='HTML'>" & "<![CDATA[" & GetActivityByNum(ActivityRecID) & "<BR> " & ActivityNotes & "<BR>"
								reqStr = reqStr & "Contact Name: " & ExtraFirstName & " " & ExtraLastName & "<BR>"
								reqStr = reqStr & "Phone: " &  ExtraPhone
								If ExtraPhoneExt<> "" Then 
									reqStr = reqStr & "Ext: " &  ExtraPhoneExt & "<BR>"
								Else
									reqStr = reqStr & "<BR>"
								End If
								reqStr = reqStr & "Email: " & ExtraEmail & "<BR>"
								reqStr = reqStr & "Address:" & "<BR>"
								reqStr = reqStr & ExtraStreet & "<BR>"
								If ExtraAddress2 <> "" Then reqStr = reqStr & ExtraAddress2 & "<BR>"
								reqStr = reqStr & ExtraCity & " , " &  ExtraState & " " & ExtraPostalCode	& "<BR>"
								reqStr = reqStr & "Primary Competitor: " & ExtraPrimaryCompetitorName  & "<BR>"
								reqStr = reqStr & "]]></t:Body>"
								'***************************************
								'END Additional details per Adam Henchel
								'***************************************
								reqStr = reqStr & "          <t:ReminderMinutesBeforeStart>60</t:ReminderMinutesBeforeStart>"
								reqStr = reqStr & "          <t:Start>" & ApptOrMeetStartDateTime & "</t:Start>"
								reqStr = reqStr & "          <t:End>" & ApptOrMeetEndDateTime  & "</t:End>"
								reqStr = reqStr & "          <t:Location>" & Replace(MeetLocation,"&","&amp;") & "</t:Location>"
								reqStr = reqStr & "          <t:RequiredAttendees>"
								reqStr = reqStr & "            <t:Attendee>"
								reqStr = reqStr & "              <t:Mailbox>"
								If Instr(ucase(Request.ServerVariables("SERVER_NAME")),"DEV") <> 0 Then
									reqStr = reqStr & "                <t:EmailAddress>" & "minsight@corpcofe.com" & "</t:EmailAddress>"
								Else
									reqStr = reqStr & "                <t:EmailAddress>" & getUserEmailAddress(UserNoForCalendarUpdate) & "</t:EmailAddress>"
								End If							
								reqStr = reqStr & "              </t:Mailbox>"
								reqStr = reqStr & "            </t:Attendee>"
								reqStr = reqStr & "            <t:Attendee>"
								reqStr = reqStr & "              <t:Mailbox>"
								reqStr = reqStr & "                <t:EmailAddress>" & "rsmith@ocsaccess.com" & "</t:EmailAddress>"
								reqStr = reqStr & "              </t:Mailbox>"
								reqStr = reqStr & "            </t:Attendee>"
								reqStr = reqStr & "          </t:RequiredAttendees>"
								reqStr = reqStr & "          <t:MeetingTimeZone TimeZoneName='Eastern Standard Time' />"
								reqStr = reqStr & "        </t:CalendarItem>"
								reqStr = reqStr & "      </m:Items>"
								reqStr = reqStr & "    </m:CreateItem>"
								reqStr = reqStr & "  </soap:Body>"
								reqStr = reqStr & "</soap:Envelope>"	
	
						End Select
					
						'Perform the actual post
						set oXMLHTTP=CreateObject("MSXML2.XMLHTTP")
						set oXML=CreateObject("MSXML2.DOMDocument")
			
						' Send the request
						oXMLHTTP.Open "POST", TARGETURL, false, USERNAME, PASSWORD
						oXMLHTTP.SetRequestHeader "Content-Type", "text/xml"
						oXMLHTTP.Send reqStr 
		
						Stat = oXMLHTTP.status
		
						If Stat = "200" THEN 
					
							If Instr(oXMLHTTP.responseText,"NoError") <> 0 Then ' Success
						
								Description ="success! oXMLHTTP.status returned " & oXMLHTTP.status & " when posting via EWS to Exchange"
								Description = "oXMLHTTP.responseText:" & oXMLHTTP.responseText & "<br>"
								Description = Description & "PAGE: " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "<br>"
								Description = Description & "Posted to " & TARGETURL  & "<br>"
								Description = Description & "POSTED DATA:" & reqStr & "<br>"
								Description = Description & "SERNO:" & MUV_READ("ClientID") & "<br>"
							
								CreateINSIGHTAuditLogEntry Request.ServerVariables("SERVER_NAME"),Description,"n/a"
							
								If ProspectApptOrMeeting ="Appointment" Then 
									Description = "An appointment" 
								Else 
									Description = "A meeting"
								End If
							
								Description = Description & " was created in the Outlook calendar for: " & GetUserDisplayNameByUserNo(UserNoForCalendarUpdate) & " for prospect " & GetProspectNameByNumber(passedProspectID) & " for the Activity " & GetActivityByNum(ActivityRecID)  & " for the date " & ActivityDate 
								
								CreateAuditLogEntry GetTerm("Prospecting") , ProspectApptOrMeeting & " created in calender","Major",0,Description
							
								If ProspectApptOrMeeting ="Appointment" Then 
									AppointmentDuration =  round(cint(ApptDuration)/60,2) & " hours"
									
									If ApptDuration mod 60 = 0 Then 
										AppointmentDuration =  round(cint(ApptDuration)/60,2) & " hour(s)"
									Else
										If ApptDuration < 60 Then
											AppointmentDuration =  ApptDuration & " minutes"
										Else
											AppointmentDuration = round(cint(ApptDuration)/60,2) & " hour(s) "
											AppointmentDuration = AppointmentDuration & ApptDuration mod 60 & " minutes"
										End If
									End If
									
									Description = "An appointment was created in the Outlook calendar for: " & GetUserDisplayNameByUserNo(UserNoForCalendarUpdate) & " for the Activity <strong><em>" & GetActivityByNum(ActivityRecID)  & "</em></strong> for, <strong><em>" & ActivityDate  & "</em></strong>, with a duration of <strong><em>" & AppointmentDuration & "</em></strong>."
								Else 
									MeetingDuration =  round(cint(MeetDuration)/60,2) & " hours"
									
									If MeetDuration mod 60 = 0 Then 
										MeetingDuration =  round(cint(MeetDuration)/60,2) & " hour(s)"
									Else
										If MeetDuration < 60 Then
											MeetingDuration =  MeetDuration & " minutes"
										Else
											MeetingDuration = round(cint(MeetDuration)/60,2) & " hour(s) "
											MeetingDuration = MeetingDuration& MeetDuration mod 60 & " minutes"
										End If
									End If

									Description = "An meeting was created in the Outlook calendar for: " & GetUserDisplayNameByUserNo(UserNoForCalendarUpdate) & " for the Activity <strong><em>" & GetActivityByNum(ActivityRecID)  & "</em></strong> for, <strong><em>" & ActivityDate  & "</em></strong>, with a duration of <strong><em>" & MeetingDuration & "</em></strong> at location: <strong><em>" & MeetLocation & "</em></strong>."	
								End If
							
								Record_PR_Activity passedProspectID,Description,Session("UserNo")
	
								'Response.Write("OK")
							Else
								'FAILURE
								emailbody="oXMLHTTP.status returned " & oXMLHTTP.status & " when posting via EWS to Exchange <br>"
								emailBody = emailBody & "oXMLHTTP.responseText:" & oXMLHTTP.responseText & "<br>"
								emailBody = emailBody & "PAGE: " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "<br>"
								emailBody = emailBody & "Posted to " & TARGETURL   & "<br>"
								emailBody = emailBody & "POSTED DATA:" & reqStr & "<br>"
								emailBody = emailBody & "SERNO:" & MUV_READ("ClientID") & "<br>"
								SendMail "mailsender@" & maildomain ,"rsmith@ocsaccess.com",MUV_READ("ClientID") & " EWS POST ERROR",emailBody, "Prospecting", "Post Failure"
						
								Description = emailBody 
								CreateINSIGHTAuditLogEntry Request.ServerVariables("SERVER_NAME"),Description,"n/a"
								
								'Response.Write("BAD")
							End If
						
						Else ' not a 200 from server
					
					'		'FAILURE
							emailbody="oXMLHTTP.status returned " & oXMLHTTP.status & " when posting via EWS to Exchange (Non 200 Status) <br>"
							emailBody = emailBody & "oXMLHTTP.responseText:" & oXMLHTTP.responseText & "<br>"
							emailBody = emailBody & "PAGE: " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "<br>"
							emailBody = emailBody & "Posted to " & TARGETURL   & "<br>"
							emailBody = emailBody & "POSTED DATA:" & reqStr & "<br>"
							emailBody = emailBody & "SERNO:" & MUV_READ("ClientID") & "<br>"
							SendMail "mailsender@" & maildomain ,"rsmith@ocsaccess.com",MUV_READ("ClientID") & " EWS POST ERROR",emailBody, "Prospecting", "Post Failure"
						
							Description = emailBody 
							CreateINSIGHTAuditLogEntry Request.ServerVariables("SERVER_NAME"),Description,"n/a"
			
						End If
					
						set oXMLHTTP=Nothing
						set oXML=Nothing
	
					End If
				
				End If
			
			End If
		
	End If	


End Function

	
Sub GetExtraInfo (ExtrapassedProspectID)

ExtraFirstName = "":ExtraLastName = "":ExtraEmail = "":ExtraPhone = "" :ExtraPhoneExt = "" 
ExtraStreet = "" :ExtraAddress2 = "" :ExtraCity = "" :ExtraState = "" :ExtraPostalCode = "" :ExtraPrimaryCompetitorName = ""


	'***********************************************************				
	'This is where we get all the extra info Adam Henchel wanted
	'************************************************************
	SQLCurrProspectInfo = "SELECT * FROM PR_ProspectContacts WHERE ProspectIntRecID = " & ExtrapassedProspectID & " AND PrimaryContact = 1"

	Set cnnCurrProspectInfo = Server.CreateObject("ADODB.Connection")
	cnnCurrProspectInfo.open (Session("ClientCnnString"))
	Set rsCurrProspectInfo = Server.CreateObject("ADODB.Recordset")
	rsCurrProspectInfo.CursorLocation = 3 
	Set rsCurrProspectInfo = cnnCurrProspectInfo.Execute(SQLCurrProspectInfo)

	If not rsCurrProspectInfo.EOF Then
	  	If Not IsNull(rsCurrProspectInfo ("FirstName")) Then ExtraFirstName = Replace(rsCurrProspectInfo("FirstName"),"&","&amp;")
	  	If Not IsNull(rsCurrProspectInfo ("LastName")) Then ExtraLastName = Replace(rsCurrProspectInfo("LastName"),"&","&amp;")
	  	If Not IsNull(rsCurrProspectInfo ("Email")) Then ExtraEmail = Replace(rsCurrProspectInfo("Email"),"&","&amp;") 
	  	If Not IsNull(rsCurrProspectInfo ("Phone")) Then ExtraPhone = Replace(rsCurrProspectInfo("Phone"),"&","&amp;") 
	  	If Not IsNull(rsCurrProspectInfo ("PhoneExt")) Then ExtraPhoneExt = Replace(rsCurrProspectInfo("PhoneExt"),"&","&amp;") 
	End If
	Set rsCurrProspectInfo = Nothing
	cnnCurrProspectInfo.Close
	Set cnnCurrProspectInfo = Nothing

	SQLCurrProspectInfo2 = "SELECT * FROM PR_Prospects WHERE InternalRecordIdentifier = " & ExtrapassedProspectID
	
	Set cnnCurrProspectInfo2 = Server.CreateObject("ADODB.Connection")
	cnnCurrProspectInfo2.open (Session("ClientCnnString"))
	Set rsCurrProspectInfo2 = Server.CreateObject("ADODB.Recordset")
	rsCurrProspectInfo2.CursorLocation = 3 
	Set rsCurrProspectInfo2 = cnnCurrProspectInfo2.Execute(SQLCurrProspectInfo2)
	
	If not rsCurrProspectInfo2.EOF Then
		If Not IsNull(rsCurrProspectInfo2("Street")) Then ExtraStreet = Replace(rsCurrProspectInfo2("Street"),"&","&amp;")
		If Not IsNull(rsCurrProspectInfo2("Floor_Suite_Room__c")) Then ExtraAddress2 = Replace(rsCurrProspectInfo2("Floor_Suite_Room__c"),"&","&amp;")
		If Not IsNull(rsCurrProspectInfo2("City")) Then ExtraCity = Replace(rsCurrProspectInfo2("City"),"&","&amp;")
		If Not IsNull(rsCurrProspectInfo2("State")) Then ExtraState = Replace(rsCurrProspectInfo2("State"),"&","&amp;")
		If Not IsNull(rsCurrProspectInfo2("PostalCode")) Then ExtraPostalCode = Replace(rsCurrProspectInfo2("PostalCode"),"&","&amp;")
	End If
	Set rsCurrProspectInfo2 = Nothing
	cnnCurrProspectInfo2.Close
	Set cnnCurrProspectInfo2 = Nothing
	
	PrimaryCompetitorID = GetPrimaryCompetitorIDByProspectNumber(ExtrapassedProspectID)

	If PrimaryCompetitorID <> "" Then
		If Not IsNull(GetCompetitorByNum(PrimaryCompetitorID)) Then ExtraPrimaryCompetitorName = Replace(GetCompetitorByNum(PrimaryCompetitorID),"&","&amp;")
	Else
		ExtraPrimaryCompetitorName = "None Entered"
	End If
	'***********************************************************				
	'END This is where we get all the extra info Adam Henchel wanted
	'************************************************************
End Sub

%>
