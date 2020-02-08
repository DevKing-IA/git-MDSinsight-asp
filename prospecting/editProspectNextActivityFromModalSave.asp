<%
'Need these here
'Dim	ExtraFirstName
'Dim	ExtraLastName
'Dim	ExtraEmail
'Dim	ExtraPhone
'Dim	ExtraPhoneExt
'Dim	ExtraStreet
'Dim	ExtraAddress2
'Dim	ExtraCity
'Dim	ExtraState
'Dim	ExtraPostalCode
'Dim	ExtraPrimaryCompetitorName 

'baseURL should alwats have a trailing /slash, just in case, handle either way
If right(baseURL,1)="/" Then maildomain = Left(right(baseURL,len(baseURL)-7),len(right(baseURL,len(baseURL)-7))-1) Else maildomain = right(baseURL,len(baseURL)-7)
maildomain = Replace(UCASE(maildomain),"WWW.","")

txtInternalRecordIdentifier = Request.Form("txtInternalRecordIdentifier")

'Response.write("txtInternalRecordIdentifier : " & txtInternalRecordIdentifier)

selProspectCurrentActivity = GetCurrentProspectActivityByProspectNumber(txtInternalRecordIdentifier)
selProspectCurrentActivityStatus = Request.Form("selProspectCurrentActivityStatus")

selProspectNextActivity = Request.Form("selProspectNextActivity")
txtProspectEditNextActivityNotes = Request.Form("txtProspectEditNextActivityNotes")
If txtProspectEditNextActivityNotes <> "" Then txtProspectEditNextActivityNotes = Replace(txtProspectEditNextActivityNotes,"&","&amp;")

txtProspectEditNextActivityDate = Request.Form("txtProspectEditNextActivityDate")

txtMeetingLocation = Replace(Request.Form("txtMeetingLocation"),"'","''")
selAppointmentDuration = Request.Form("selAppointmentDuration")
selMeetingDuration = Request.Form("selMeetingDuration")

ProspectOwnerUserNo = GetProspectOwnerNoByNumber(txtInternalRecordIdentifier)
'Response.Write("<br><br>selMeetingDuration : " & selMeetingDuration & "<br><br>")
'Response.Write("txtMeetingLocation : " & txtMeetingLocation & "<br><br>")
'Response.Write("selAppointmentDuration : " & selAppointmentDuration & "<br><br>")

ProspectName = GetProspectNameByNumber(txtInternalRecordIdentifier)	
ProspectNewActivity = GetActivityByNum (selProspectNextActivity)

ProspectApptOrMeeting = GetActivityApptOrMeetingByNum(selProspectNextActivity)

If selProspectCurrentActivity <> "" AND selProspectNextActivity <> "" AND txtInternalRecordIdentifier <> "" Then
	
	'Update current activity

	Set cnnProspectNextActivityUpdate = Server.CreateObject("ADODB.Connection")
	cnnProspectNextActivityUpdate.open Session("ClientCnnString")
	
	SQLProspectNextActivityUpdate = "UPDATE PR_ProspectActivities Set Status = '" & selProspectCurrentActivityStatus & "',StatusDateTime = GetDate(), "
	SQLProspectNextActivityUpdate = SQLProspectNextActivityUpdate & " Notes = '" & txtProspectEditNextActivityNotes & "', StatusChangedByUserNo = " & Session("UserNo")
	SQLProspectNextActivityUpdate = SQLProspectNextActivityUpdate & " WHERE ProspectRecID = " & txtInternalRecordIdentifier & " AND Status IS NULL "
	
	Set rsProspectNextActivityUpdate = Server.CreateObject("ADODB.Recordset")
	rsProspectNextActivityUpdate.CursorLocation = 3 
	Set rsProspectNextActivityUpdate = cnnProspectNextActivityUpdate.Execute(SQLProspectNextActivityUpdate)
	
	
	Description = "The next activity <strong><em>" & selProspectCurrentActivity & "</em></strong> for prospect " & ProspectName  & " was changed to <strong><em>" & selProspectCurrentActivityStatus & "</em></strong> by " & GetUserDisplayNameByUserNo(Session("UserNo")) 
	CreateAuditLogEntry GetTerm("Prospecting") & " next activity changed",GetTerm("Prospecting") & " next activity changed","Major",0,Description
	
	Description = "The next activity <strong><em>" & selProspectCurrentActivity & "</em></strong> was changed to <strong><em>" & selProspectCurrentActivityStatus & "</em></strong> by " & GetUserDisplayNameByUserNo(Session("UserNo")) 
	Record_PR_Activity txtInternalRecordIdentifier,Description,Session("UserNo")
	
	
	set rsProspectNextActivityUpdate = Nothing
	cnnProspectNextActivityUpdate.Close
	set cnnProspectNextActivityUpdate = Nothing
		
	'Insert new activity
	
	Set cnnProspectNextActivityInsert = Server.CreateObject("ADODB.Connection")
	cnnProspectNextActivityInsert.open Session("ClientCnnString")
		
	If ProspectApptOrMeeting <> "" Then

		If ProspectApptOrMeeting ="Appointment" Then
		
			Duration = cint(selAppointmentDuration)
										
			SQLProspectNextActivityInsert = "INSERT INTO PR_ProspectActivities (ProspectRecID, ActivityRecID, ActivityDueDate, ActivityCreatedByUserNo, ActivityIsAppointment, ActivityIsMeeting, ActivityAppointmentDuration) "
			SQLProspectNextActivityInsert = SQLProspectNextActivityInsert & " VALUES (" & txtInternalRecordIdentifier & ", " & selProspectNextActivity & ",'" & txtProspectEditNextActivityDate & "'," & ProspectOwnerUserNo & ",1,0," & Duration & ") "
						
		ElseIf ProspectApptOrMeeting ="Meeting" Then
		
			Duration = cint(selMeetingDuration)
									
			SQLProspectNextActivityInsert = "INSERT INTO PR_ProspectActivities (ProspectRecID, ActivityRecID, ActivityDueDate, ActivityCreatedByUserNo, ActivityIsAppointment, ActivityIsMeeting, ActivityMeetingDuration, ActivityMeetingLocation) "
			SQLProspectNextActivityInsert = SQLProspectNextActivityInsert & " VALUES (" & txtInternalRecordIdentifier & ", " & selProspectNextActivity & ",'" & txtProspectEditNextActivityDate & "'," & ProspectOwnerUserNo & ",0,1," & Duration & ",'" & txtMeetingLocation & "') "
								
		Else
						
			SQLProspectNextActivityInsert = "INSERT INTO PR_ProspectActivities (ProspectRecID, ActivityRecID, ActivityDueDate, ActivityCreatedByUserNo, ActivityIsAppointment, ActivityIsMeeting) "
			SQLProspectNextActivityInsert = SQLProspectNextActivityInsert & " VALUES (" & txtInternalRecordIdentifier & ", " & selProspectNextActivity & ",'" & txtProspectEditNextActivityDate & "'," & ProspectOwnerUserNo & ",0,0) "		
		
		End If
	Else
							
		SQLProspectNextActivityInsert = "INSERT INTO PR_ProspectActivities (ProspectRecID, ActivityRecID, ActivityDueDate, ActivityCreatedByUserNo, ActivityIsAppointment, ActivityIsMeeting) "
		SQLProspectNextActivityInsert = SQLProspectNextActivityInsert & " VALUES (" & txtInternalRecordIdentifier & ", " & selProspectNextActivity & ",'" & txtProspectEditNextActivityDate & "'," & ProspectOwnerUserNo & ",0,0) "		

	End If
	

	Set rsProspectNextActivityInsert = Server.CreateObject("ADODB.Recordset")	
	rsProspectNextActivityInsert.CursorLocation = 3 
	Set rsProspectNextActivityInsert = cnnProspectNextActivityInsert.Execute(SQLProspectNextActivityInsert)	
	
	set rsProspectNextActivityInsert = Nothing
	cnnProspectNextActivityInsert.Close
	set cnnProspectNextActivityInsert = Nothing
	
							
	Description = "The next activity " & selProspectCurrentActivity & " for prospect " & ProspectName  & " was changed to " & GetActivityByNum(selProspectNextActivity)  & " by " & GetUserDisplayNameByUserNo(Session("UserNo")) & " with a due date of " & txtProspectEditNextActivityDate 
	CreateAuditLogEntry GetTerm("Prospecting") & " next activity changed",GetTerm("Prospecting") & " next activity changed","Major",0,Description

	Description = "The next activity was set to <strong><em>" & GetActivityByNum(selProspectNextActivity) & "</em></strong> by " & GetUserDisplayNameByUserNo(Session("UserNo")) & " with a due date of " & txtProspectEditNextActivityDate  
	Record_PR_Activity txtInternalRecordIdentifier,Description,Session("UserNo")
		

	
	' Now see if we need to create an appointment or meeting in the users email system

	If ProspectApptOrMeeting <> "" Then
	
		'OK, see if we have credentials for this user

		If GetUserEmailSystemIDByUserNo(ProspectOwnerUserNo) <> "" AND GetUserEmailSystemPassByUserNo((ProspectOwnerUserNo)) <> "" Then

				'OK, see if we allow access to calendar for this user

				If AllowUpdatesToUsersCalendar(ProspectOwnerUserNo) = True Then
				
				'Try setting outlook appointment
				'TARGETURL="https://mail.corpcofe.com/ews/exchange.asmx"
				'USERNAME="minsight@corpcofe.com"
				'PASSWORD="minsight"
				'TARGETURL="https://outlook.office365.com/EWS/Exchange.asmx"

				TARGETURL = GetPOSTParams("EWSPostURL")
				USERNAME = GetUserEmailSystemIDByUserNo(ProspectOwnerUserNo)
				PASSWORD = GetUserEmailSystemPassByUserNo((ProspectOwnerUserNo))

				'Adjust from Coordinated Universal Time to EST
'				txtProspectEditNextActivityDate = dateadd("h",4,txtProspectEditNextActivityDate)
				txtProspectEditNextActivityDate = dateadd("h",5,txtProspectEditNextActivityDate)

				ApptOrMeetStartDateTime  = Year(txtProspectEditNextActivityDate) & "-" 
				If Month(txtProspectEditNextActivityDate) < 10 Then
					ApptOrMeetStartDateTime  = ApptOrMeetStartDateTime  & "0" & Month(txtProspectEditNextActivityDate)
				Else
					ApptOrMeetStartDateTime  = ApptOrMeetStartDateTime  & Month(txtProspectEditNextActivityDate)
				End If
				ApptOrMeetStartDateTime  = ApptOrMeetStartDateTime  & "-"
				If Day(txtProspectEditNextActivityDate) < 10 Then
					ApptOrMeetStartDateTime  = ApptOrMeetStartDateTime  & "0" & Day(txtProspectEditNextActivityDate)
				Else
					ApptOrMeetStartDateTime  = ApptOrMeetStartDateTime  & Day(txtProspectEditNextActivityDate)
				End If
				ApptOrMeetStartDateTime  = ApptOrMeetStartDateTime  & "T"
				If Hour(txtProspectEditNextActivityDate) < 10 Then
					ApptOrMeetStartDateTime  = ApptOrMeetStartDateTime  & "0" & Hour(txtProspectEditNextActivityDate)
				Else
					ApptOrMeetStartDateTime  = ApptOrMeetStartDateTime  & Hour(txtProspectEditNextActivityDate)
				End If
				ApptOrMeetStartDateTime  = ApptOrMeetStartDateTime  & ":"
				If Minute(txtProspectEditNextActivityDate) < 10 Then
					ApptOrMeetStartDateTime  = ApptOrMeetStartDateTime  & "0" & Minute(txtProspectEditNextActivityDate)
				Else
					ApptOrMeetStartDateTime  = ApptOrMeetStartDateTime  & Minute(txtProspectEditNextActivityDate)
				End If
				ApptOrMeetStartDateTime  = ApptOrMeetStartDateTime  & ":"
				ApptOrMeetStartDateTime  = ApptOrMeetStartDateTime  & "00.000Z"
			
				If ProspectApptOrMeeting ="Appointment" Then
					'Duration = cint(GetPOSTParams("EWSDefaultApptDuration"))
					Duration = cint(selAppointmentDuration)
				Else
					'Duration = cint(GetPOSTParams("EWSDefaultMeetingDuration"))
					Duration = cint(selMeetingDuration)
				End If
				If Not IsNumeric(Duration) Then Duration = 15
				txtProspectEditNextActivityDate2 = DateAdd("n",Duration,cdate(txtProspectEditNextActivityDate))
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

						Call GetExtraInfo (txtInternalRecordIdentifier)
						
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
						reqStr = reqStr & "          <t:Subject>" & Replace(ProspectName,"&","&amp;") & "</t:Subject>"
						'***********************************
						'Additional details per Adam Henchel
						'***********************************			
						reqStr = reqStr & "          <t:Body BodyType='HTML'>" & "<![CDATA[" & ProspectNewActivity  & "<BR> " & txtProspectEditNextActivityNotes & "<BR>"
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
						reqStr = reqStr & "          <t:Location>" & Replace(ProspectName,"&","&amp;")  & "</t:Location>"
						reqStr = reqStr & "          <t:MeetingTimeZone TimeZoneName='Eastern Standard Time' />"
						reqStr = reqStr & "        </t:CalendarItem>"
						reqStr = reqStr & "      </m:Items>"
						reqStr = reqStr & "    </m:CreateItem>"
						reqStr = reqStr & "  </soap:Body>"
						reqStr = reqStr & "</soap:Envelope>"						

					Case "Meeting"

								
						If txtMeetingLocation = "" Then
							'Meeting so need to get location info, if blank
							
							Set cnntmpProspect = Server.CreateObject("ADODB.Connection")
							cnntmpProspect.open Session("ClientCnnString")
			
							SQLtmpProspect = "Select * from PR_Prospects Where InternalRecordIdentifier = " & txtInternalRecordIdentifier
		 
							Set rstmpProspect = Server.CreateObject("ADODB.Recordset")
							rstmpProspect.CursorLocation = 3 
							Set rstmpProspect = cnntmpProspect.Execute(SQLtmpProspect)
				
							If not rstmpProspect.EOF Then
							
								Street = rstmpProspect("Street")
								City = rstmpProspect("City")
								Floor_Suite_Room__c = rstmpProspect("Floor_Suite_Room__c")		
								
								If Street <> "" and Not IsNull(Street) Then txtMeetingLocation = Street & Chr(13)
								If City <> "" and Not IsNull(City) Then txtMeetingLocation = txtMeetingLocation & City & Chr(13)
								If Floor_Suite_Room__c <> "" and Not IsNull(Floor_Suite_Room__c) Then txtMeetingLocation = txtMeetingLocation & Floor_Suite_Room__c
								
								
							End If
							
							rstmpProspect.Close
							set rstmpProspect= Nothing
							cnntmpProspect.Close	
							set cnntmpProspect= Nothing
						End If
						If txtMeetingLocation = "" Then txtMeetingLocation = ProspectName 
				
						Call GetExtraInfo (txtInternalRecordIdentifier)
				
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
						reqStr = reqStr & "          <t:Subject>" & Replace(ProspectName,"&","&amp;") & "</t:Subject>"
						'***********************************
						'Additional details per Adam Henchel
						'***********************************			
						reqStr = reqStr & "          <t:Body BodyType='HTML'>" & "<![CDATA[" & ProspectNewActivity  & "<BR> " & txtProspectEditNextActivityNotes & "<BR>"
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
						reqStr = reqStr & "          <t:Location>" & Replace(txtMeetingLocation,"&","&amp;") & "</t:Location>"
						reqStr = reqStr & "          <t:RequiredAttendees>"
						reqStr = reqStr & "            <t:Attendee>"
						reqStr = reqStr & "              <t:Mailbox>"
						reqStr = reqStr & "                <t:EmailAddress>" & getUserEmailAddress(ProspectOwnerUserNo) & "</t:EmailAddress>"
						reqStr = reqStr & "              </t:Mailbox>"
						reqStr = reqStr & "            </t:Attendee>"
'						reqStr = reqStr & "            <t:Attendee>"
'						reqStr = reqStr & "              <t:Mailbox>"
'						reqStr = reqStr & "                <t:EmailAddress>rsmith@ocsaccess.com</t:EmailAddress>"
'						reqStr = reqStr & "              </t:Mailbox>"
'						reqStr = reqStr & "            </t:Attendee>"
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
				'oXMLHTTP.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
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
						
						SendMail "mailsender@" & maildomain ,"rsmith@ocsaccess.com",MUV_READ("ClientID") & " EWS POST OK",Description, "Prospecting", "Post OK"				
						
						If ProspectApptOrMeeting ="Appointment" Then 
							Description = "An appointment" 
						Else 
							Description = "A meeting"
						End If
						
						Description = Description & " was automatically created in the Outlook calendar for: " & GetUserDisplayNameByUserNo(ProspectOwnerUserNo) & " for prospect " & ProspectName  & " for the Activity " & GetActivityByNum(selProspectNextActivity)  & " for the date " & txtProspectEditNextActivityDate 
						CreateAuditLogEntry GetTerm("Prospecting") & " next activity changed",GetTerm("Prospecting") & " next activity changed","Major",0,Description
						
						
						If ProspectApptOrMeeting ="Appointment" Then 
							AppointmentDuration = selAppointmentDuration & " (" & round(cint(selAppointmentDuration)/60,2) & " hours)"
							Description = "An appointment was automatically created in the Outlook calendar for: " & GetUserDisplayNameByUserNo(ProspectOwnerUserNo) & " for the Activity <strong><em>" & GetActivityByNum(selProspectNextActivity)  & "</em></strong> on the date, <strong><em>" & txtProspectEditNextActivityDate & "</em></strong>, with a duration of <strong><em>" & AppointmentDuration & "</em></strong>."
						Else 
							MeetingDuration = selMeetingDuration & " (" & round(cint(selMeetingDuration)/60,2) & " hours)"
							Description = "An meeting was automatically created in the Outlook calendar for: " & GetUserDisplayNameByUserNo(ProspectOwnerUserNo) & " for the Activity <strong><em>" & GetActivityByNum(selProspectNextActivity)  & "</em></strong> on the date, <strong><em>" & txtProspectEditNextActivityDate & "</em></strong>, with a duration of <strong><em>" & MeetingDuration & "</em></strong> at location: <strong><em>" & txtMeetingLocation & "</em></strong>."	
						End If
						
						Record_PR_Activity txtInternalRecordIdentifier,Description,Session("UserNo")
						

						Response.Write("OK")
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
	  	If Not IsNull(rsCurrProspectInfo("FirstName")) Then ExtraFirstName = Replace(rsCurrProspectInfo("FirstName"),"&","&amp;")
	  	If Not IsNull(rsCurrProspectInfo("LastName")) Then ExtraLastName = Replace(rsCurrProspectInfo("LastName"),"&","&amp;")
	  	If Not IsNull(rsCurrProspectInfo("Email")) Then ExtraEmail = Replace(rsCurrProspectInfo("Email"),"&","&amp;") 
	  	If Not IsNull(rsCurrProspectInfo("Phone")) Then ExtraPhone = Replace(rsCurrProspectInfo("Phone"),"&","&amp;") 
	  	If Not IsNull(rsCurrProspectInfo("PhoneExt")) Then ExtraPhoneExt = Replace(rsCurrProspectInfo("PhoneExt"),"&","&amp;") 
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