<!--#include file="../../../js/json2.asp"-->

<%

'*****************************************************************
'CALL THE HOLIDAY API AND USE JSON2.ASP TO PARSE RESPONSE
'*****************************************************************

CurrentMonthInYear = Month(Now())
CurrentYear = Year(Now())

If ClientCountry = "United States" Then	
	country = "US"
Else
	country = "CA"
End If


dim url
url = "https://holidayapi.com/v1/holidays"
holidayApiKey = "35e7eebe-5a95-4ae1-bb53-52ba1879167d"

'****************************************************************************************
 Response.Write("Begin Entries for Settings_Company Calendar Holidays<br><br>")
'****************************************************************************************

Set cnnCompanyCalendar = Server.CreateObject("ADODB.Connection")
cnnCompanyCalendar.open (Session("ClientCnnString"))
Set rsCompanyCalendar = Server.CreateObject("ADODB.Recordset")
rsCompanyCalendar.CursorLocation = 3 	


'**************************************************************************************************************
'FIRST CLEAR OUT THE ENTRIES IN THE SC_NEEDTOKNOW TABLE FOR COMPANY CALENDAR HOLIDAYS
'**************************************************************************************************************
SQL_CompanyCalendar = "DELETE FROM SC_NeedToKnow WHERE Module = 'Global Settings' AND SubModule ='Company Calendar Holidays'"
Set rsCompanyCalendar = cnnCompanyCalendar.Execute(SQL_CompanyCalendar)
'**************************************************************************************************************


for monthNum = cInt(CurrentMonthInYear) to 12

	Set HttpReq = Server.CreateObject("MSXML2.ServerXMLHTTP")
	HttpReq.open "POST", url, false
	HttpReq.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	HttpReq.Send("key=" & holidayApiKey & "&country=" & country & "&year=" & CurrentYear & "&month=" & cInt(monthNum) & "&public=true")
	'use json2 via jscript to parse the response
	Set holidayResponseJSON = JSON.parse(HttpReq.responseText)
	
	'*****************************************************************
	'STATUS CODES RETURNED FROM HOLIDAY API
	'*****************************************************************
	'	200 Success! Everything is A-OK
	'	400 Something is wrong on your end
	'	401 Unauthorized (did you remember your API key?)
	'	402 Payment required (only historical data available is free)
	'	403 Forbidden (this API is HTTPS-only)
	'	429 Rate limit exceeded
	'	500 OH NOES!!~! Something is wrong on our end
	'*****************************************************************
	
	dim key : for each key in holidayResponseJSON.holidays.keys() 
			
		holidayName = holidayResponseJSON.holidays.get(key).name
		holidayName = Replace(holidayName, "'", "''")
		holidayCalendarDate = holidayResponseJSON.holidays.get(key).date
		holidayObservedDate = holidayResponseJSON.holidays.get(key).observed 
		
		holidayCalendarDateFormatted = formatDateTime(holidayCalendarDate,2)
		holidayObservedDateFormatted = formatDateTime(holidayObservedDate,2)
		observedYearNum = Year(holidayObservedDateFormatted)
		observedMonthNum = Month(holidayObservedDateFormatted)
		observedDayNum = Day(holidayObservedDateFormatted)
		
		'Response.write( holidayName  & "<br>")
		'Response.write( observedYearNum & "<br>")
		'Response.write( observedMonthNum & "<br>")
		'Response.write( observedDayNum & "<br>")
		'Response.write( formatDateTime(holidayCalendarDate,2)  & "<br>")
		'Response.write( holidayObservedDateFormatted & "<br>")

		'**************************************************************************************************************
		'CHECK CUSTOMER FOR MISSING HOLIDAYS IN THEIR COMPANY CALENDAR
		'IF IT IS DECEMBER OF THE CURRENT CALENDAR YEAR, THEN WE NEED TO CHECK FOR THE FOLLOWING YEAR AS WELL
		'**************************************************************************************************************
		
		SQL_CompanyCalendar = "SELECT * FROM Settings_CompanyCalendar WHERE YearNum = " & observedYearNum & " AND MonthNum = " & observedMonthNum & " AND DayNum = " & observedDayNum
		Set rsCompanyCalendar = cnnCompanyCalendar.Execute(SQL_CompanyCalendar)
		
		If rsCompanyCalendar.EOF Then
		
			SCNeedToKnow_Module = "Global Settings"
			SCNeedToKnow_SubModule = "Company Calendar Holidays"
			SCNeedToKnow_SummaryDescription = "Missing Holiday in Company Calendar"
			SCNeedToKnow_DetailedDescription1 = "The company calendar is missing the holiday " & holidayName & ". This holiday is on " & holidayCalendarDateFormatted & " and is observed on " & holidayObservedDateFormatted & ". Please add it to your company calendar if appropriate."
			SCNeedToKnow_InsightStaffOnly = 0
	
			'*****************************************************************************************************************
			'Check to see if record already exists in SC_NeedToKnow
			'*****************************************************************************************************************
			
			SQL_SCNeedToKnowCheckIfExists = " SELECT * FROM SC_NeedToKnow WHERE Module = '" & SCNeedToKnow_Module & "' AND SubModule = '" & SCNeedToKnow_SubModule & "' "
			SQL_SCNeedToKnowCheckIfExists = SQL_SCNeedToKnowCheckIfExists & " AND SummaryDescription = '" & SCNeedToKnow_SummaryDescription & "' "
			SQL_SCNeedToKnowCheckIfExists = SQL_SCNeedToKnowCheckIfExists & " AND DetailedDescription1 = '" & SCNeedToKnow_DetailedDescription1 & "' "
			
			Set rsSCNeedToKnowCheckIfExists = cnnSCNeedToKnowCheckIfExists.Execute(SQL_SCNeedToKnowCheckIfExists)
			
			If rsSCNeedToKnowCheckIfExists.EOF Then
						
				SQL_SCNeedToKnow = "INSERT INTO SC_NeedToKnow (Module, SubModule, SummaryDescription, DetailedDescription1, InsightStaffOnly) VALUES "
				SQL_SCNeedToKnow = SQL_SCNeedToKnow & " ('" & SCNeedToKnow_Module & "', '" & SCNeedToKnow_SubModule & "', '" & SCNeedToKnow_SummaryDescription & "', "
				SQL_SCNeedToKnow = SQL_SCNeedToKnow & " '" & SCNeedToKnow_DetailedDescription1 & "', " & SCNeedToKnow_InsightStaffOnly & ")"
				
				Set rsSCNeedToKnow = cnnSCNeedToKnow.Execute(SQL_SCNeedToKnow)
				
				If QuietMode = False Then
					Response.Write(SCNeedToKnow_Module & " - " & SCNeedToKnow_SubModule & " - " & SCNeedToKnow_SummaryDescription & " - " & SCNeedToKnow_DetailedDescription1 & "<br>")
					Response.Write("<hr>")
				End If
			
			End If
			'*****************************************************************************************************************
					
		End If
		
	next
next


'*****************************************************************************************************************
'PROCESS HOLIDAYS FOR THE ENTIRE FOLLOWING YEAR IF THE CURRENT MONTH IS DECEMBER
'*****************************************************************************************************************
If cInt(CurrentMonthInYear) = 12 Then
	
	NextYear = Cint(Year(Now())) + 1
	
	for monthNum = 1 to 12
	
		Set HttpReq = Server.CreateObject("MSXML2.ServerXMLHTTP")
		HttpReq.open "POST", url, false
		HttpReq.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		HttpReq.Send("key=" & holidayApiKey & "&country=" & country & "&year=" & NextYear & "&month=" & cInt(monthNum) & "&public=true")
		'use json2 via jscript to parse the response
		Set holidayResponseJSON = JSON.parse(HttpReq.responseText)
		
		'*****************************************************************
		'STATUS CODES RETURNED FROM HOLIDAY API
		'*****************************************************************
		'	200 Success! Everything is A-OK
		'	400 Something is wrong on your end
		'	401 Unauthorized (did you remember your API key?)
		'	402 Payment required (only historical data available is free)
		'	403 Forbidden (this API is HTTPS-only)
		'	429 Rate limit exceeded
		'	500 OH NOES!!~! Something is wrong on our end
		'*****************************************************************
		
		dim key2 : for each key2 in holidayResponseJSON.holidays.keys() 
				
			holidayName = holidayResponseJSON.holidays.get(key2).name
			holidayName = Replace(holidayName, "'", "''")
			holidayCalendarDate = holidayResponseJSON.holidays.get(key2).date
			holidayObservedDate = holidayResponseJSON.holidays.get(key2).observed 
			
			holidayCalendarDateFormatted = formatDateTime(holidayCalendarDate,2)
			holidayObservedDateFormatted = formatDateTime(holidayObservedDate,2)
			observedYearNum = Year(holidayObservedDateFormatted)
			observedMonthNum = Month(holidayObservedDateFormatted)
			observedDayNum = Day(holidayObservedDateFormatted)
			
			'**************************************************************************************************************
			'CHECK CUSTOMER FOR MISSING HOLIDAYS IN THEIR COMPANY CALENDAR
			'IF IT IS DECEMBER OF THE CURRENT CALENDAR YEAR, THEN WE NEED TO CHECK FOR THE FOLLOWING YEAR AS WELL
			'**************************************************************************************************************
			
			SQL_CompanyCalendar = "SELECT * FROM Settings_CompanyCalendar WHERE YearNum = " & observedYearNum & " AND MonthNum = " & observedMonthNum & " AND DayNum = " & observedDayNum
			Set rsCompanyCalendar = cnnCompanyCalendar.Execute(SQL_CompanyCalendar)
			
			If rsCompanyCalendar.EOF Then
			
				SCNeedToKnow_Module = "Global Settings"
				SCNeedToKnow_SubModule = "Company Calendar Holidays"
				SCNeedToKnow_SummaryDescription = "Missing Holiday in Company Calendar"
				SCNeedToKnow_DetailedDescription1 = "The company calendar is missing the holiday " & holidayName & ". This holiday is on " & holidayCalendarDateFormatted & " and is observed on " & holidayObservedDateFormatted & ". Please add it to your company calendar if appropriate."
				SCNeedToKnow_InsightStaffOnly = 0
		
				'*****************************************************************************************************************
				'Check to see if record already exists in SC_NeedToKnow
				'*****************************************************************************************************************
				
				SQL_SCNeedToKnowCheckIfExists = " SELECT * FROM SC_NeedToKnow WHERE Module = '" & SCNeedToKnow_Module & "' AND SubModule = '" & SCNeedToKnow_SubModule & "' "
				SQL_SCNeedToKnowCheckIfExists = SQL_SCNeedToKnowCheckIfExists & " AND SummaryDescription = '" & SCNeedToKnow_SummaryDescription & "' "
				SQL_SCNeedToKnowCheckIfExists = SQL_SCNeedToKnowCheckIfExists & " AND DetailedDescription1 = '" & SCNeedToKnow_DetailedDescription1 & "' "
				
				Set rsSCNeedToKnowCheckIfExists = cnnSCNeedToKnowCheckIfExists.Execute(SQL_SCNeedToKnowCheckIfExists)
				
				If rsSCNeedToKnowCheckIfExists.EOF Then
							
					SQL_SCNeedToKnow = "INSERT INTO SC_NeedToKnow (Module, SubModule, SummaryDescription, DetailedDescription1, InsightStaffOnly) VALUES "
					SQL_SCNeedToKnow = SQL_SCNeedToKnow & " ('" & SCNeedToKnow_Module & "', '" & SCNeedToKnow_SubModule & "', '" & SCNeedToKnow_SummaryDescription & "', "
					SQL_SCNeedToKnow = SQL_SCNeedToKnow & " '" & SCNeedToKnow_DetailedDescription1 & "', " & SCNeedToKnow_InsightStaffOnly & ")"
					
					Set rsSCNeedToKnow = cnnSCNeedToKnow.Execute(SQL_SCNeedToKnow)
					
					If QuietMode = False Then
						Response.Write(SCNeedToKnow_Module & " - " & SCNeedToKnow_SubModule & " - " & SCNeedToKnow_SummaryDescription & " - " & SCNeedToKnow_DetailedDescription1 & "<br>")
						Response.Write("<hr>")
					End If
				
				End If
				'*****************************************************************************************************************
						
			End If
			
		next
	next
	

End If
'*****************************************************************************************************************

Set rsCompanyCalendar = Nothing
cnnCompanyCalendar.Close
Set cnnCompanyCalendar = Nothing

'****************************************************************************************
 Response.Write("End Entries for Settings_Company Calendar Holidays<br><br>")
'****************************************************************************************
	
							
%>