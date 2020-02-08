<!--#include file="inc/InSightFuncs.asp"-->
<!--#include file="inc/InSightFuncs_Users.asp"-->
<!--#include file="inc/settings.asp"-->

<%


'declare the variables
Dim Connection
Dim Recordset
Dim SQL



Username = Request.Form("txtUsername")	
Password = Request.Form("txtPassword")
ClientKey = Request.Form("txtClientKey")


QuickLogin = Request.Form("txtQuickLogin")
QuickClientDestination = Request.Form("txtDestinationURL")
UserNo = Request.Form("txtUserNo")

If ClientKey = "" Then
	ClientKey = Request.Form("txtClientKeyCustom")
	customLoginPageFlag = true
Else
	customLoginPageFlag = false
End If

If QuickLogin <> "" Then
	quickLoginPageFlag = true
Else
	quickLoginPageFlag = false
End If

SQL = "SELECT * FROM tblServerInfo where clientKey='"& ClientKey &"'"

Set Connection = Server.CreateObject("ADODB.Connection")
Set Recordset = Server.CreateObject("ADODB.Recordset")

Connection.Open InsightCnnString

'Open the recordset object executing the SQL statement and return records
Recordset.Open SQL,Connection,3,3

'First lookup the ClientKey in tblServerInfo
'If there is no record with the entered client key, close connection
'and go back to login with QueryString
If Recordset.recordcount <= 0 then
	Recordset.close
	Connection.close
	set Recordset=nothing
	set Connection=nothing
	
	If customLoginPageFlag = true AND (ClientKey = "1071" OR ClientKey = "1071d") Then
		If quickLoginPageFlag = true Then
			Response.Redirect("ql-CCS.asp?login=namefailed&u=" & UserNo & "&c=" & ClientKey & "&d=" & QuickClientDestination)
		Else
			Response.Redirect("default_customLoginCCS.asp?login=namefailed")
		End If
	ElseIf customLoginPageFlag = true AND ClientKey <> "" AND ClientKey <> "1071" AND ClientKey <> "1071d" Then
		If quickLoginPageFlag = true Then
			Response.Redirect("ql.asp?login=namefailed&u=" & UserNo & "&c=" & ClientKey & "&d=" & QuickClientDestination)
		Else
			Response.Redirect("default.asp?login=namefailed"& "clientID=" & ClientKey)
		End If
	Else
		If quickLoginPageFlag = true Then
			Response.Redirect("ql.asp?login=namefailed&u=" & UserNo & "&c=" & ClientKey & "&d=" & QuickClientDestination)
		Else
			Response.Redirect("default.asp?login=namefailed")
		End If
	End If
	
Else
	Session("ClientCnnString") = "Driver={SQL Server};Server=" & Recordset.Fields("dbServer")
	Session("ClientCnnString") = Session("ClientCnnString") & ";Database=" & Recordset.Fields("dbCatalog")
	Session("ClientCnnString") = Session("ClientCnnString") & ";Uid=" & Recordset.Fields("dbLogin")
	Session("ClientCnnString") = Session("ClientCnnString") & ";Pwd=" & Recordset.Fields("dbPassword") & ";"
	dummy = MUV_Write("SQL_Owner",Recordset.Fields("dbLogin"))
	Session("SQL_Owner") = Recordset.Fields("dbLogin")
	dummy = MUV_Write("ClientID",Recordset.Fields("clientkey"))
	dummy = MUV_Write("BackendSystem",Recordset.Fields("Backend"))
	dummy = MUV_Write("SERNO",Recordset.Fields("clientkey"))
	If Recordset.Fields("advancedDispatch") = 1 Then advancedDispatch = True Else advancedDispatch = False
	'**********************
	' Load Leftnav Options
	'**********************
	dummy = MUV_Write("prospectingModuleOn",Recordset.Fields("prospectingModule"))
	dummy = MUV_Write("routingModuleOn",Recordset.Fields("routingModule"))
	dummy = MUV_Write("biModuleOn",Recordset.Fields("biModule"))
	dummy = MUV_Write("custServiceOn",Recordset.Fields("ShowCustServiceMenu"))
	dummy = MUV_Write("arModuleOn",Recordset.Fields("arModule"))
	dummy = MUV_Write("nightBatchModuleOn",Recordset.Fields("nightBatchModule"))
	dummy = MUV_Write("OrderAPIModuleOn",Recordset.Fields("OrderAPIModule"))
	dummy = MUV_Write("InventoryControlModuleOn",Recordset.Fields("InventoryControlModule"))
	dummy = MUV_Write("equipmentModuleOn",Recordset.Fields("equipmentModule"))
	dummy = MUV_Write("apModuleOn",Recordset.Fields("apModule"))
	dummy = MUV_Write("serviceModuleOn",Recordset.Fields("serviceModule"))
	dummy = MUV_Write("webFulfillmentModuleOn",Recordset.Fields("webFulfillmentModule"))
	dummy = MUV_Write("invoicingModuleOn",Recordset.Fields("invoicingModule"))
	dummy = MUV_Write("quickbooksModuleOn",Recordset.Fields("quickbooksModule"))
	dummy = MUV_Write("FilterTrax",Recordset.Fields("FilterTrax"))
	dummy = MUV_Write("ShowAddEditCustomer",Recordset.Fields("ShowAddEditCustomer"))
	'**********************

	Recordset.close
	Connection.close	
End If	


Connection.Open Session("ClientCnnString")

SQL = "SELECT * FROM tblUsers where userEmail='"& Username &"' and userpassword='"& Password &"'"
'Open the recordset object executing the SQL statement and return records
Recordset.Open SQL,Connection,3,3

'If there is no record with the entered username, close connection
'and go back to login with QueryString
If Recordset.recordcount <= 0 then
	Recordset.close
	Connection.close
	set Recordset=nothing
	set Connection=nothing
	
	If customLoginPageFlag = true AND (ClientKey = "1071" OR ClientKey = "1071d") Then
		If quickLoginPageFlag = true Then
			Response.Redirect("ql-CCS.asp?login=namefailed&u=" & UserNo & "&c=" & ClientKey & "&d=" & QuickClientDestination)
		Else
			Response.Redirect("default_customLoginCCS.asp?login=namefailed")
		End If
	ElseIf customLoginPageFlag = true AND ClientKey <> "" AND ClientKey <> "1071" AND ClientKey <> "1071d" Then
		If quickLoginPageFlag = true Then
			Response.Redirect("ql.asp?login=namefailed&u=" & UserNo & "&c=" & ClientKey & "&d=" & QuickClientDestination)
		Else
			Response.Redirect("default.asp?login=namefailed"& "clientID=" & ClientKey)
		End If
	Else
		If quickLoginPageFlag = true Then
			Response.Redirect("ql.asp?login=namefailed&u=" & UserNo & "&c=" & ClientKey & "&d=" & QuickClientDestination)
		Else
			Response.Redirect("default.asp?login=namefailed")
		End If
	End If
	
	
Else
	If Recordset.Fields("userEnabled") <> True Then
		Fname = Recordset.Fields("userFirstName")
		Lname = Recordset.Fields("userLastName")
		'User not enabled
		Recordset.close
		Connection.close
		set Recordset=nothing
		set Connection=nothing
		
		If customLoginPageFlag = true AND (ClientKey = "1071" OR ClientKey = "1071d") Then
			If quickLoginPageFlag = true Then
				Response.Redirect("ql-CCS.asp?login=disabled&u=" & UserNo & "&c=" & ClientKey & "&d=" & QuickClientDestination)
			Else
				Response.Redirect("default_customLoginCCS.asp?login=disabled")
			End If
		ElseIf customLoginPageFlag = true AND ClientKey <> "" AND ClientKey <> "1071" AND ClientKey <> "1071d" Then
			If quickLoginPageFlag = true Then
				Response.Redirect("ql.asp?login=disabled&u=" & UserNo & "&c=" & ClientKey & "&d=" & QuickClientDestination)
			Else
				Response.Redirect("default.asp?login=disabled"& "clientID=" & ClientKey)
			End If
		Else
			If quickLoginPageFlag = true Then
				Response.Redirect("ql.asp?login=disabled&u=" & UserNo & "&c=" & ClientKey & "&d=" & QuickClientDestination)
			Else
				Response.Redirect("default.asp?login=disabled")
			End If
		End If
		
		
	Else
	
		
		'******************************************************************************************************************
		'******************************************************************************************************************
		'IF THE USER HAS MADE IT PAST THE VALID LOGIN AND LOGIN NOT DISABLED CHECKS, NOW CHECK TO SEE IF THERE ARE ANY
		'TIME OR HOLIDAY RESTRICTIONS ON THEIR USER ID, AND REDIRECT THEM WITH A MESSAGE, IF APPROPRIATE
		'******************************************************************************************************************
		Function ampmTime(InTime)
			
			'Response.Write("InTime: " & InTime & "<br>")
			
			if hour(InTime) < 12 then
				OutHour = hour(InTime)
				ampm = "am"
			end if
			
			if hour(InTime) = 12 then
				OutHour = hour(InTime)
				ampm = "pm"
			end if
			
			if hour(InTime) > 12 then
				OutHour = hour(InTime) - 12
				ampm = "pm"
			end if
			
			minutes = minute(intime)
			IF LEN(minutes) < 2 THEN minutes = minutes & "0"
			IF LEN(OutHour) < 2 THEN OutHour = "0" & OutHour
			ampmTime = OutHour & ":" & minutes & " " & ampm
			
		End Function	
				
		'******************************************************************************************************************
		'FIRST CHECK SC_UserRestrictedLoginSchedule
		'******************************************************************************************************************
		
		userNo = Recordset.Fields("userNo")
		
		If userNo = "" Then
			UserNo = Request.Form("txtUserNo")
		End If
		
		SQLRestrictedLoginHours = "SELECT * FROM SC_UserRestrictedLoginSchedule WHERE userNo = " & userNo
		
		Set cnnRestrictedLoginHours = Server.CreateObject("ADODB.Connection")
		cnnRestrictedLoginHours.open (Session("ClientCnnString"))
		Set rsRestrictedLoginHours = Server.CreateObject("ADODB.Recordset")
		rsRestrictedLoginHours.CursorLocation = 3 
		Set rsRestrictedLoginHours = cnnRestrictedLoginHours.Execute(SQLRestrictedLoginHours)
		
		If NOT rsRestrictedLoginHours.EOF Then
		
			currentWeekdayNum = cInt(Weekday(Now()) - 1)
			currentHour = Hour(Now())
			Session("restrictedLoginMessage") = ""
		
			Do While NOT rsRestrictedLoginHours.EOF
			
				restrictedWeekdayNum = cInt(rsRestrictedLoginHours("DayNo"))
				
				restrictedWeekdayStartTime = rsRestrictedLoginHours("StartRestrictedTime")
				restrictedWeekdayEndTime = rsRestrictedLoginHours("EndRestrictedTime")

				restrictedWeekdayStarHour = cInt(Left(restrictedWeekdayStartTime,2))
				restrictedWeekdayEndHour = cInt(Left(restrictedWeekdayEndTime,2))
				
				If currentWeekdayNum = restrictedWeekdayNum Then
				
					If currentHour >= restrictedWeekdayStarHour AND currentHour < restrictedWeekdayEndHour Then
					
						If restrictedWeekdayStartTime = "00:00" OR restrictedWeekdayEndTime = "24:00" Then
						
							If restrictedWeekdayStartTime = "00:00" AND restrictedWeekdayEndTime = "24:00" Then
								restrictedWeekdayStartTime = "midnight"
								restrictedWeekdayEndTime = "midnight"
								
							ElseIf restrictedWeekdayStartTime = "00:00" Then
								restrictedWeekdayStartTime = "midnight"
								restrictedWeekdayEndTime = ampmTime(restrictedWeekdayEndTime)
								
							ElseIf restrictedWeekdayEndTime = "24:00" Then
								restrictedWeekdayStartTime = ampmTime(restrictedWeekdayStartTime)
								restrictedWeekdayEndTime = "midnight"
								
							End If
						Else
							restrictedWeekdayStartTime = ampmTime(restrictedWeekdayStartTime)
							restrictedWeekdayEndTime = ampmTime(restrictedWeekdayEndTime)
						End If
					
						Session("restrictedLoginMessage") = "You are currently restricted from logging in from " & restrictedWeekdayStartTime & " until " & restrictedWeekdayEndTime & "."
				
						If customLoginPageFlag = true AND (ClientKey = "1071" OR ClientKey = "1071d") Then
							If quickLoginPageFlag = true Then
								Response.Redirect("ql-CCS.asp?login=hoursresctriction&u=" & UserNo & "&c=" & ClientKey & "&d=" & QuickClientDestination)
							Else
								Response.Redirect("default_customLoginCCS.asp?login=hoursresctriction")
							End If
						ElseIf customLoginPageFlag = true AND ClientKey <> "" AND ClientKey <> "1071" AND ClientKey <> "1071d" Then
							If quickLoginPageFlag = true Then
								Response.Redirect("ql.asp?login=hoursresctriction&u=" & UserNo & "&c=" & ClientKey & "&d=" & QuickClientDestination)
							Else
								Response.Redirect("default.asp?login=hoursresctriction&clientID=" & ClientKey)
							End If
						Else
							If quickLoginPageFlag = true Then
								Response.Redirect("ql.asp?login=hoursresctriction&u=" & UserNo & "&c=" & ClientKey & "&d=" & QuickClientDestination)
							Else
								Response.Redirect("default.asp?login=hoursresctriction")
							End If
						End If
					
					End If
				
				End If
			
			rsRestrictedLoginHours.MoveNext
			Loop
		End If
		cnnRestrictedLoginHours.close
		
		
		'******************************************************************************************************************
		'NOW CHECK Settings_CompanyCalendar
		'******************************************************************************************************************
		
		SQLUserDisableHolidays = "SELECT * FROM tblUsers WHERE userNo = " & userNo
		
		Set cnnUserDisableHolidays = Server.CreateObject("ADODB.Connection")
		cnnUserDisableHolidays.open (Session("ClientCnnString"))
		Set rsUserDisableHolidays = Server.CreateObject("ADODB.Recordset")
		rsUserDisableHolidays.CursorLocation = 3 
		Set rsUserDisableHolidays = cnnUserDisableHolidays.Execute(SQLUserDisableHolidays)
		
		If NOT rsUserDisableHolidays.EOF Then
		
			userLoginDisableAccessHolidays = rsUserDisableHolidays("userLoginDisableAccessHolidays")
			
			'************************************************************
			'IF THE USER IS NOT ALLOWED TO LOGIN ON HOLIDAYS/CLOSINGS
			'AS DEFINED IN THE COMPANY CALENDAR, THEN CHECK TO SEE IF
			'TODAY IS A DEFINED HOLIDAY
			'************************************************************
		
			If userLoginDisableAccessHolidays = 1 OR userLoginDisableAccessHolidays = vbTrue OR userLoginDisableAccessHolidays = True Then

			
				currentWeekdayNum = cInt(Weekday(Now()) - 1)
				currentMonthNum = cInt(Month(Now()))
				currentDayNum = cInt(Day(Now()))
				currentYearNum = cInt(Year(Now()))
				currentHour = Hour(Now())
				currentMinutes = Minute(Now())
				Session("restrictedLoginMessage") = ""

				'************************************************************
				'ONLY RETRIEVE TODAY FROM THE HOLIDAY CALENDAR
				'************************************************************
				
				SQLHolidayHours = "SELECT * FROM Settings_CompanyCalendar WHERE MonthNum = " & currentMonthNum & " AND DayNum = " & currentDayNum & " AND YearNum = " & currentYearNum
				
				Set cnnHolidayHours = Server.CreateObject("ADODB.Connection")
				cnnHolidayHours.open (Session("ClientCnnString"))
				Set rsHolidayHours = Server.CreateObject("ADODB.Recordset")
				rsHolidayHours.CursorLocation = 3 
				Set rsHolidayHours = cnnHolidayHours.Execute(SQLHolidayHours)
				
				If NOT rsHolidayHours.EOF Then
				
					'************************************************************
					'FIRST CHECK TO SEE IF TODAY IS CLOSED; IF SO NO LOGINS
					'AT ANY TIME TODAY ARE ALLOWED
					'************************************************************
					
					If rsHolidayHours("OpenClosedCloseEarly") = "Closed" Then					
					
						Session("restrictedLoginMessage") = "Today is a company holiday and logins are restricted."
						
						If customLoginPageFlag = true AND (ClientKey = "1071" OR ClientKey = "1071d") Then
							If quickLoginPageFlag = true Then
								Response.Redirect("ql-CCS.asp?login=holidayresctriction&u=" & UserNo & "&c=" & ClientKey & "&d=" & QuickClientDestination)
							Else
								Response.Redirect("default_customLoginCCS.asp?login=holidayresctriction")
							End If
						ElseIf customLoginPageFlag = true AND ClientKey <> "" AND ClientKey <> "1071" AND ClientKey <> "1071d" Then
							If quickLoginPageFlag = true Then
								Response.Redirect("ql.asp?login=holidayresctriction&u=" & UserNo & "&c=" & ClientKey & "&d=" & QuickClientDestination)
							Else
								Response.Redirect("default.asp?login=holidayresctriction&clientID=" & ClientKey)
							End If
						Else
							If quickLoginPageFlag = true Then
								Response.Redirect("ql.asp?login=holidayresctriction&u=" & UserNo & "&c=" & ClientKey & "&d=" & QuickClientDestination)
							Else
								Response.Redirect("default.asp?login=holidayresctriction")
							End If
						End If
						
						
					'****************************************************************
					'NEXT CHECK TO SEE IF TODAY IS AN EARLY CLOSING; IF SO NO LOGINS
					'AT ANY TIME TODAY ARE ALLOWED AT OR AFTER THE CLOSE EARLY
					'TIME DEFINED IN THE SQL TABLE
					'****************************************************************
						
					ElseIf rsHolidayHours("OpenClosedCloseEarly") = "Close Early" Then
					
						ClosingTime = rsHolidayHours("ClosingTime")
						
						ClosingTimeHour = cInt(Left(ClosingTime,2))
						ClosingTimeMinutes = cInt(Right(ClosingTime,2))
	

						'************************************************************
						'CHECK TO SEE IF THE USER IS LOGGING IN AFTER THE EARLY 
						'CLOSING TIME. IF THEY HAVE LOGGED IN THE SAME HOUR, THEN
						'WE HAVE TO CHECK THE MINUTES AS WELL
						'************************************************************
						
						If (currentHour > ClosingTimeHour) OR (currentHour = ClosingTimeHour AND currentMinutes >= ClosingTimeMinutes) Then
						
							ClosingTime = ampmTime(ClosingTime)

							Session("restrictedLoginMessage") = "The company is closing early today at " & ClosingTime & ". Logins are restricted after hours."
					
							If customLoginPageFlag = true AND (ClientKey = "1071" OR ClientKey = "1071d") Then
								If quickLoginPageFlag = true Then
									Response.Redirect("ql-CCS.asp?login=holidayresctriction&u=" & UserNo & "&c=" & ClientKey & "&d=" & QuickClientDestination)
								Else
									Response.Redirect("default_customLoginCCS.asp?login=holidayresctriction")
								End If
							ElseIf customLoginPageFlag = true AND ClientKey <> "" AND ClientKey <> "1071" AND ClientKey <> "1071d" Then
								If quickLoginPageFlag = true Then
									Response.Redirect("ql.asp?login=holidayresctriction&u=" & UserNo & "&c=" & ClientKey & "&d=" & QuickClientDestination)
								Else
									Response.Redirect("default.asp?login=holidayresctriction"& "clientID=" & ClientKey)
								End If
							Else
								If quickLoginPageFlag = true Then
									Response.Redirect("ql.asp?login=holidayresctriction&u=" & UserNo & "&c=" & ClientKey & "&d=" & QuickClientDestination)
								Else
									Response.Redirect("default.asp?login=holidayresctriction")
								End If
							End If

						End If

				
					End If
				End If
				
				
				cnnHolidayHours.close
					
			End If
			
		End If
		cnnUserDisableHolidays.close
			
	
			
		'******************************************************************************************************************
		'******************************************************************************************************************
	
		If Right(ClientKey,1) = "d" Then
			ClientKeyForFileNames = LEFT(ClientKey, (LEN(ClientKey)-1))
		Else
			ClientKeyForFileNames = ClientKey
		End If	

		dummy = MUV_Write("ClientKeyForFileNames",ClientKeyForFileNames)
		dummy = MUV_Write("DisplayName",Recordset.Fields("userdisplayname"))
		Session("userEmail")=Recordset.Fields("useremail")
		If Recordset.Fields("userType") = "Admin" Then Session("adminPrivelages") = True Else Session("adminPrivelages") = False
		If Recordset.Fields("userType") = "Field Service" Then Session("FieldService") = True Else Session("FieldService") = False
		Session("userNo")=Recordset.Fields("userNo")
		userNo = Recordset.Fields("userNo")
		If Recordset.Fields("LoginLandingPageURL") <> "" Then
			If NOT IsNull(Recordset.Fields("LoginLandingPageURL")) Then
				dummy = MUV_Write("LoginLandingPageURL", baseURL & Recordset.Fields("LoginLandingPageURL"))
			Else
				dummy = MUV_Write("LoginLandingPageURL", "")
			End IF
		Else
			dummy = MUV_Write("LoginLandingPageURL", "")
		End IF
				
		Recordset.Close
		SQL = "UPDATE tblUsers SET userLastLogin = '" & Now() & "' where userNo = " & userNo
		Set Recordset = Connection.Execute(SQL)

		Connection.Close
		set Recordset=nothing
		set Connection=nothing
		If Session("FieldService") = True Then
			CreateAuditLogEntry "Login","Login","Major",0, Session("userEmail") & " logged in to Field Service WebApp"
		Else
			CreateAuditLogEntry "Login","Login","Major",0, Session("userEmail") & " logged in."		
		End If
		
		' Redirects based on user types
		' Sets the page to go to in MUV_READ("LoginPage")
		' becuase it must go through it license check
		'before redirecting the to that page

			
			
		'*******************************************************************************************
		'SET DESTINATION AFTER LOGIN BASED ON USER TYPE, OR DESTINATION URL IF SPECIFIED
		'*******************************************************************************************	
		If QuickClientDestination = "" Then
		
			dummy = MUV_WRITE("LoginPage","main/")
	
			If userIsDriver(Session("userNo")) Then dummy = MUV_WRITE("LoginPage","mobile/drivers/deliveryboard/main_menu.asp")
			
			'Field Service is special because we gather the geolocation information
			If Session("FieldService") = True Then
				If advancedDispatch = False Then 
					dummy = MUV_WRITE("LoginPage","fieldservice/main_menu.asp")
				Else
					dummy = MUV_WRITE("LoginPage","fieldserviceADV/main_menu.asp")
				End If
			End If
			
		ElseIf InStr(QuickClientDestination,"editProspect-") Then
		
			destinationParamtersArray = Split(QuickClientDestination, "-")
			prospectID = destinationParamtersArray(1)
			dummy = MUV_WRITE("LoginPage","prospecting/editProspect.asp?i=" & prospectID)
		
		ElseIf InStr(QuickClientDestination,"viewProspect-") Then
		
			destinationParamtersArray = Split(QuickClientDestination, "-")
			prospectID = destinationParamtersArray(1)
			dummy = MUV_WRITE("LoginPage","prospecting/viewProspectDetail.asp?i=" & prospectID)
		
		ElseIf InStr(QuickClientDestination,"CustAnalSum_1.asp") Then
		
			destinationParamtersArray = Split(QuickClientDestination, "-")
			qlSlsmnParam = destinationParamtersArray(1)
			dummy = MUV_WRITE("LoginPage","bizintel/CustAnalSum_1.asp?" & qlSlsmnParam )
		'	response.end

			
		End If
		
		If ClientKey = "1128" Then
			If Session("UserNo") = 72 Then dummy = MUV_WRITE("LoginPage","/mobile/inventorycontrol/main_menu.asp")	
			If Session("UserNo") = 82 Then dummy = MUV_WRITE("LoginPage","/mobile/inventorycontrol/main_menu.asp")	
		End If
		
		If Session("UserNo") = 2286 Then dummy = MUV_WRITE("LoginPage","/mobile/inventorycontrol/main_menu.asp")	
		
		'Now do the lic check
		'''''Response.Write("QuickClientDestination: " & QuickClientDestination & "<br><br>")
		'''''Response.Write(MUV_READ("LoginPage"))
		'''''Response.end
		Response.Redirect ("LicenseCheck.asp")
	End IF
End If



%>

