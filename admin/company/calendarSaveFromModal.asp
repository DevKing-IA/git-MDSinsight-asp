<!--#include file="../../inc/subsandfuncs.asp"-->
<!--#include file="../../inc/InSightFuncs.asp"-->

<%

dateToEdit = Request.Form("dateToEdit")
txtBusinessDayDescription = Request.Form("txtBusinessDayDescription")
txtBusinessDayDescription = Replace(txtBusinessDayDescription,"'","''")
closeEarlyTimepicker = Request.Form("closeEarlyTimepicker")
radUpdatedDateStatus = Request.Form("radUpdatedDateStatus")

dateToEditDateFormat = FormatDateTime(dateToEdit,2)
dateToEditYear = Year(dateToEditDateFormat)
dateToEditMonth = Month(dateToEditDateFormat)
dateToEditDay = Day(dateToEditDateFormat)
dateToEditMonthName = Left(MonthName(dateToEditMonth),3)

If closeEarlyTimepicker <> "" Then
	closeEarlyTime = FormatDateTime(closeEarlyTimepicker, 4)
End If

dateAlternate=Request.Form("alterdate")
response.write (dateAlternate)
'Response.Write("dateToEdit : " & dateToEdit & "<br>")
'Response.Write("txtBusinessDayDescription : " & txtBusinessDayDescription & "<br>")
'Response.Write("closeEarlyTimepicker : " & closeEarlyTimepicker & "<br>")
'Response.Write("radUpdatedDateStatus : " & radUpdatedDateStatus & "<br>")
'Response.Write("dateToEditYear : " & dateToEditYear & "<br>")
'Response.Write("dateToEditMonth : " & dateToEditMonth & "<br>")
'Response.Write("dateToEditDay : " & dateToEditDay & "<br>")
'Response.Write("dateToEditMonthName : " & dateToEditMonthName & "<br>")
'Response.Write("closeEarlyTime : " & closeEarlyTime & "<br>")



'****************************************************************************************
'Lookup the calendar record as it exists now so we can fill in the audit trail
'****************************************************************************************

dateCurrentlyExistsInSQL = False

SQL = "SELECT * FROM Settings_CompanyCalendar WHERE YearNum = " & dateToEditYear & " AND MonthNum = " & dateToEditMonth & " AND DayNum = " & dateToEditDay
Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3 
Set rs = cnn8.Execute(SQL)

If not rs.EOF Then
	dateCurrentlyExistsInSQL = True
	Orig_MonthNam = rs("MonthNam")
	Orig_MonthNum = rs("MonthNum")
	Orig_DayNum = rs("DayNum")
	Orig_YearNum = rs("YearNum")
	Orig_OpenClosedCloseEarly = rs("OpenClosedCloseEarly")
	Orig_ClosingTime = rs("ClosingTime")
	Orig_Description = rs("Description")
Else
	dateCurrentlyExistsInSQL = False
	Orig_MonthNam = ""
	Orig_MonthNum = ""
	Orig_DayNum = ""
	Orig_YearNum = ""
	Orig_OpenClosedCloseEarly = ""
	Orig_ClosingTime = ""
	Orig_Description = ""
End If

'Response.Write("dateCurrentlyExistsInSQL : " & dateCurrentlyExistsInSQL & "<br>")



set rs = Nothing
cnn8.close
set cnn8 = Nothing

'****************************************************************************************
'End Lookup the calendar record as it exists now so we can fill in the audit trail
'****************************************************************************************

If radUpdatedDateStatus = "Open" Then

	'******************************************************************
	'The user has updated the calendar business day status to "open",
	'so we just need to remove it from the SQL table completely
	'******************************************************************
	
	SQL = "DELETE FROM Settings_CompanyCalendar WHERE YearNum = " & dateToEditYear & " AND MonthNum = " & dateToEditMonth & " AND DayNum = " & dateToEditDay
		
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rs = cnn8.Execute(SQL)
		
	set rs = Nothing
	cnn8.close
	set cnn8 = Nothing
	
	If dateCurrentlyExistsInSQL = True Then
		If Orig_OpenClosedCloseEarly = "Close Early" Then
			AuditMessage = "The company calendar date status for "  & dateToEdit & " changed from " & Orig_OpenClosedCloseEarly & " at " & Orig_ClosingTime & " (" & Orig_Description & ") to " & radUpdatedDateStatus  			
			CreateAuditLogEntry "Company Calendar Change", "Company Calendar Change", "Major", 1, AuditMessage
		ElseIf Orig_OpenClosedCloseEarly = "Closed" Then
			AuditMessage = "The company calendar date status for "  & dateToEdit & " changed from " & Orig_OpenClosedCloseEarly & " (" & Orig_Description & ") to " & radUpdatedDateStatus  			
			CreateAuditLogEntry "Company Calendar Change", "Company Calendar Change", "Major", 1, AuditMessage	 
		End If
	Else
		AuditMessage = "The company calendar date status for "  & dateToEdit & " was set to " & radUpdatedDateStatus 			
		CreateAuditLogEntry "Company Calendar Change", "Company Calendar Change", "Major", 1, AuditMessage		
	End If

	
ElseIf radUpdatedDateStatus = "Closed" Then

	'******************************************************************
	'The user has updated the calendar business day status to "closed",
	'so we need to use our variable to see if the date needs to be
	'updated or inserted into Settings_CompanyCalendar
	'******************************************************************
	If dateCurrentlyExistsInSQL = False Then
		SQL = "INSERT INTO Settings_CompanyCalendar (MonthNam, MonthNum, DayNum, YearNum, OpenClosedCloseEarly, ClosingTime, Description"
        IF LEN(dateAlternate)>0 THEN
            SQL =SQL &",AlternateDeliveryDate"
        END IF
        SQL =SQL &") "
		SQL = SQL & " VALUES ('" & dateToEditMonthName & "'," & dateToEditMonth & "," & dateToEditDay & "," & dateToEditYear & ",'Closed','','" & txtBusinessDayDescription & "'"
        IF LEN(dateAlternate)>0 THEN
            SQL =SQL &",TRY_PARSE(RTrim(LTrim('"&dateAlternate&"')) AS datetime USING 'en-US')"
        END IF
        SQL =SQL &") "
	Else
		SQL = "UPDATE Settings_CompanyCalendar SET MonthNam= '" & dateToEditMonthName & "', MonthNum = " & dateToEditMonth & ","
		SQL = SQL & " DayNum = " & dateToEditDay & ", YearNum = " & dateToEditYear & ", OpenClosedCloseEarly='Closed', ClosingTime = '', Description ='" & txtBusinessDayDescription & "'"
        IF LEN(dateAlternate)>0 THEN
            SQL = SQL &",AlternateDeliveryDate=TRY_PARSE(RTrim(LTrim('"&dateAlternate&"')) AS datetime USING 'en-US')"
        END IF
		SQL = SQL & " WHERE YearNum = " & dateToEditYear & " AND MonthNum = " & dateToEditMonth & " AND DayNum = " & dateToEditDay
	End If
	
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rs = cnn8.Execute(SQL)
		
	set rs = Nothing
	cnn8.close
	set cnn8 = Nothing
	
	If dateCurrentlyExistsInSQL = True Then
		If Orig_OpenClosedCloseEarly = "Close Early" Then
			AuditMessage = "The company calendar date status for "  & dateToEdit & " changed from " & Orig_OpenClosedCloseEarly & " at " & Orig_ClosingTime & " (" & Orig_Description & ") to " & radUpdatedDateStatus & " (" & txtBusinessDayDescription & ")"			
			CreateAuditLogEntry "Company Calendar Change", "Company Calendar Change", "Major", 1, AuditMessage
		ElseIf Orig_OpenClosedCloseEarly = "Closed" Then
			AuditMessage = "The company calendar date status for "  & dateToEdit & " changed from " & Orig_OpenClosedCloseEarly & " (" & Orig_Description & ") to " & radUpdatedDateStatus & " (" & txtBusinessDayDescription & ")"  			
			CreateAuditLogEntry "Company Calendar Change", "Company Calendar Change", "Major", 1, AuditMessage	 
		End If
	Else
		AuditMessage = "The company calendar date status for "  & dateToEdit & " was changed from open to " & radUpdatedDateStatus & " (" & txtBusinessDayDescription & ")" 			
		CreateAuditLogEntry "Company Calendar Change", "Company Calendar Change", "Major", 1, AuditMessage		
	End If
	
	
ElseIf radUpdatedDateStatus = "Close Early" Then

	'**************************************************************************
	'The user has updated the calendar business day status to "closing early",
	'so we need to use our variable to see if the date needs to be
	'updated or inserted into Settings_CompanyCalendar
	'**************************************************************************
	If dateCurrentlyExistsInSQL = False Then
		SQL = "INSERT INTO Settings_CompanyCalendar (MonthNam, MonthNum, DayNum, YearNum, OpenClosedCloseEarly, ClosingTime, Description"
        IF LEN(dateAlternate)>0 THEN
            SQL =SQL &",AlternateDeliveryDate"
        END IF
        SQL =SQL &") "
		SQL = SQL & " VALUES ('" & dateToEditMonthName & "'," & dateToEditMonth & "," & dateToEditDay & "," & dateToEditYear & ",'Close Early','" & closeEarlyTime & "','" & txtBusinessDayDescription & "'"
        IF LEN(dateAlternate)>0 THEN
            SQL =SQL &",TRY_PARSE(RTrim(LTrim('"&dateAlternate&"')) AS datetime USING 'en-US')"
        END IF
        SQL =SQL &") "
	Else
		SQL = "UPDATE Settings_CompanyCalendar SET MonthNam= '" & dateToEditMonthName & "', MonthNum = " & dateToEditMonth & ","
		SQL = SQL & " DayNum = " & dateToEditDay & ", YearNum = " & dateToEditYear & ", OpenClosedCloseEarly ='Close Early', ClosingTime='" & closeEarlyTime & "', Description ='" & txtBusinessDayDescription & "'"
        IF LEN(dateAlternate)>0 THEN
            SQL = SQL &",AlternateDeliveryDate=TRY_PARSE(RTrim(LTrim('"&dateAlternate&"')) AS datetime USING 'en-US')"
        END IF
		SQL = SQL & " WHERE YearNum = " & dateToEditYear & " AND MonthNum = " & dateToEditMonth & " AND DayNum = " & dateToEditDay
	End If
	
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rs = cnn8.Execute(SQL)
		
	set rs = Nothing
	cnn8.close
	set cnn8 = Nothing
	
	
	If dateCurrentlyExistsInSQL = True Then
		If Orig_OpenClosedCloseEarly = "Close Early" Then
			AuditMessage = "The company calendar date status for "  & dateToEdit & " changed from " & Orig_OpenClosedCloseEarly & " at " & Orig_ClosingTime & " (" & Orig_Description & ") to " & radUpdatedDateStatus & " at " & closeEarlyTime & " (" & txtBusinessDayDescription & ")"			
			CreateAuditLogEntry "Company Calendar Change", "Company Calendar Change", "Major", 1, AuditMessage
		ElseIf Orig_OpenClosedCloseEarly = "Closed" Then
			AuditMessage = "The company calendar date status for "  & dateToEdit & " changed from " & Orig_OpenClosedCloseEarly & " (" & Orig_Description & ") to " & radUpdatedDateStatus & " at " & closeEarlyTime & " (" & txtBusinessDayDescription & ")"	  			
			CreateAuditLogEntry "Company Calendar Change", "Company Calendar Change", "Major", 1, AuditMessage	 
		End If
	Else
		AuditMessage = "The company calendar date status for "  & dateToEdit & " was changed from open to " & radUpdatedDateStatus & " at " & closeEarlyTime & " (" & txtBusinessDayDescription & ")"	 			
		CreateAuditLogEntry "Company Calendar Change", "Company Calendar Change", "Major", 1, AuditMessage		
	End If
	
	
End If


'Response.write(SQL)

Response.Redirect("main.asp#calendar1")

%>