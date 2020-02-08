<%
' This code is for the chart of Audit Trail Events
TotalUsers = 0

Set cnn = Server.CreateObject("ADODB.Connection")
cnn.open Session("ClientCnnString")


'************************
'Read Settings_Screens
'************************
SelectedUserDisplayNames = ""
NumberOfDays = 10 ' This is the default unless over-ridden later
SQL = "SELECT * from Settings_Screens where ScreenNumber = 1000 AND UserNo = " & Session("userNo")
Set rsInsight2 = Server.CreateObject("ADODB.Recordset")
rsInsight2.CursorLocation = 3 
Set rsInsight2= cnn.Execute(SQL)
If NOT rsInsight2.EOF Then
	SelectedUserDisplayNames = rsInsight2("ScreenSpecificData2")
	NumberOfDays = cint(rsInsight2("ScreenSpecificData1"))
	FservCloseOnly = rsInsight2("ScreenSpecificData3")
End If
Set rsInsight2= Nothing
If FservCloseOnly <> "1" Then 
	ActivityChartSubTitle = "Last " & NumberOfDays & " Days (Excluding Login / Logout)"
Else
	ActivityChartSubTitle = "Last " & NumberOfDays & " Days (Excluding Login / Logout)<br>Field Techs showing closes only"
End IF
'****************************
'End Read Settings_Screens
'****************************



SQL = "Select userDisplayName, userType from tblUsers WHERE userArchived <> 1 order by userDisplayName"
Set rsInsight1 = Server.CreateObject("ADODB.Recordset")
rsInsight1.CursorLocation = 3 
Set rsInsight1 = cnn.Execute(SQL)

If not rsInsight1.eof then 
	
	'Max 200 users
	ReDim UserAuditArray(200,2)  'First element will userDisplayName, Second element will be activity count
	UserAuditArrayIndex = 1
	Do
	
		If SelectedUserDisplayNames = "" Then 'There was nothing specified so do them all
		
			UserAuditArray(UserAuditArrayIndex,0) = rsInsight1("userDisplayName")
	
			If FservCloseOnly ="1" AND userType="Field Service" Then
				' Now get the number of events
				SQL = "Select COUNT(AuditUserDisplayName) AS Expr1 from SC_AuditLog where "
				SQL = SQL & "AuditElementOrEventName = 'Service Call Closed' AND "
				SQL = SQL & "CAST(AuditEntryDateTime as DATE) >= '" &  FormatDateTime(DateAdd("d",-(NumberOfDays-1),Now()),2)  & "' AND "
				SQL = SQL & "AuditUserDisplayName = '" & rsInsight1("userDisplayName") & "'"
			Else
				' Now get the number of events
				SQL = "Select COUNT(AuditUserDisplayName) AS Expr1 from SC_AuditLog where "
				SQL = SQL & "AuditElementOrEventName <> 'Login' AND AuditElementOrEventName <> 'Logout' AND "
				SQL = SQL & "CAST(AuditEntryDateTime as DATE) >= '" &  FormatDateTime(DateAdd("d",-(NumberOfDays-1),Now()),2)  & "' AND "
				SQL = SQL & "AuditUserDisplayName = '" & rsInsight1("userDisplayName") & "'"
			End If

			Set rsInsight2 = Server.CreateObject("ADODB.Recordset")
			rsInsight2.CursorLocation = 3 
			Set rsInsight2 = cnn.Execute(SQL)
			
			If not rsInsight2.eof Then
				UserAuditArray(UserAuditArrayIndex,1) = rsInsight2("Expr1")
			Else
				UserAuditArray(UserAuditArrayIndex,1) = 0
			End If 
		
			UserAuditArrayIndex = UserAuditArrayIndex + 1
			
		Else ' There was users specifid in the seetings table, so check this against them, also use the # days specified
		
			If Instr( SelectedUserDisplayNames,"," & rsInsight1("userDisplayName") & ",") <> 0 Then
			
				UserAuditArray(UserAuditArrayIndex,0) = rsInsight1("userDisplayName")
		
			
				If FservCloseOnly ="1" AND userType="Field Service" Then
					' Now get the number of events
					SQL = "Select COUNT(AuditUserDisplayName) AS Expr1 from SC_AuditLog where "
					SQL = SQL & "AuditElementOrEventName = 'Service Call Closed' AND "
					SQL = SQL & "CAST(AuditEntryDateTime as DATE) >= '" &  FormatDateTime(DateAdd("d",-(NumberOfDays-1),Now()),2)  & "' AND "
					SQL = SQL & "AuditUserDisplayName = '" & rsInsight1("userDisplayName") & "'"
				Else
					' Now get the number of events
					SQL = "Select COUNT(AuditUserDisplayName) AS Expr1 from SC_AuditLog where "
					SQL = SQL & "AuditElementOrEventName <> 'Login' AND AuditElementOrEventName <> 'Logout' AND "
					SQL = SQL & "CAST(AuditEntryDateTime as DATE) >= '" &  FormatDateTime(DateAdd("d",-(NumberOfDays-1),Now()),2)  & "' AND "
					SQL = SQL & "AuditUserDisplayName = '" & rsInsight1("userDisplayName") & "'"
				End If

				Set rsInsight2 = Server.CreateObject("ADODB.Recordset")
				rsInsight2.CursorLocation = 3 
				Set rsInsight2 = cnn.Execute(SQL)

				If not rsInsight2.eof Then
					UserAuditArray(UserAuditArrayIndex,1) = rsInsight2("Expr1")
				Else
					UserAuditArray(UserAuditArrayIndex,1) = 0
				End If 
			
				UserAuditArrayIndex = UserAuditArrayIndex + 1
			End If
		End If
		
		rsInsight1.movenext
	Loop While Not rsInsight1.eof
	
End IF	

MaxCount = UserAuditArrayIndex 
aspDataVar = ""

For x = 1 to MaxCount -1
	UserAuditArray(x,0) = "<a href=""" & BaseURL & "reports/AuditTrail_OneUser.asp?unam=" & Server.URLEncode(UserAuditArray(x,0)) & """>" & UserAuditArray(x,0) & "</a>"
	aspDataVar = aspDataVar & "['" & UserAuditArray(x,0) & "'," & UserAuditArray(x,1) & "],"
Next
'Strip the last comma
If Len (aspDataVar) > 1 Then aspDataVar = Left(aspDataVar,len(aspDataVar)-1)
%>
