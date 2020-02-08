<!--#include file="InSightFuncs.asp"-->
<!--#include file="InSightFuncs_Users.asp"-->
<%

'***************************************************
'List of all the AJAX functions & subs
'***************************************************
 
'Sub ClearLoginAccessForExistingUser()
'Sub ClearLoginAccessNewForUser()
'Sub UpdateLoginAccessForExistingUser()
'Sub UpdateLoginAccessForNewUser()

'***************************************************
'End List of all the AJAX functions & subs
'***************************************************


'********************************************************************************************************************************************************
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

'ALL AJAX MODAL SUBROUTINES AND FUNCTIONS BELOW THIS AREA

'********************************************************************************************************************************************************
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

action = Request("action")

Select Case action
	Case "ClearLoginAccessForExistingUser" 
		ClearLoginAccessForExistingUser()
	Case "ClearLoginAccessNewForUser"
		ClearLoginAccessNewForUser()	
	Case "UpdateLoginAccessForExistingUser"
		UpdateLoginAccessForExistingUser()
	Case "UpdateLoginAccessForNewUser"
		UpdateLoginAccessForNewUser()
    Case "updateCustomOrDefault"
        updateCustomOrDefault()
End Select

'********************************************************************************************************************************************************
'********************************************************************************************************************************************************




'********************************************************************************************************************************************************
'********************************************************************************************************************************************************


Sub UpdateLoginAccessForExistingUser()

	userNo = Request.Form("userNo") 
	jsonString = Request.Form("jsonString") 
	
	'Response.Write(jsonString)
	
	'********************************************************************
	'When a user selects new login restricted access times, we are rebuilding ALL records in SC_UserRestrictedLoginSchedule,
	'so we need to delete all existing records first
	'********************************************************************
	
	SQLDelete = "DELETE FROM SC_UserRestrictedLoginSchedule WHERE userNo = " & userNo
	
	Set cnnDelete = Server.CreateObject("ADODB.Connection")
	cnnDelete.open (Session("ClientCnnString"))
	Set rsDelete = Server.CreateObject("ADODB.Recordset")
	rsDelete.CursorLocation = 3 
	Set rsDelete = cnnDelete.Execute(SQLDelete)
	cnnDelete.close
	
	'********************************************************************
	'Prepare the jsonString for parsing by removing extraneous characters
	'********************************************************************
	'remove the opening [ in the string
	jsonString = Right(jsonString, Len(jsonString) - 1)
	
	'remove the closing ] in the string
	jsonString = Left(jsonString,len(jsonString)-1)
	
	'remove double quotes from the string
	jsonString = Replace(jsonString, """","")
	
	'********************************************************************


	'********************************************************************
	'Now build the new login records and insert them into SC_UserRestrictedLoginSchedule
	
	Set cnnInsert = Server.CreateObject("ADODB.Connection")
	cnnInsert.open (Session("ClientCnnString"))
	Set rsInsert = Server.CreateObject("ADODB.Recordset")
	rsInsert.CursorLocation = 3 

	If InStr(jsonString,"},{") Then
	
		'Multiple days have been selected with restricted login access
		
		jsonArray = Split(jsonString, "},{")
		
		for i = 0 to Ubound(jsonArray)
		
			singleDayString = Split(jsonArray(i),",")
			
			''singleDayString[0] = Contains Day Number
			''singleDayString[1] = Contains Day Number Restricted Start Time
			''singleDayString[2] = Contains Day Number Restricted End Time
			
			'remove opening bracket from day string
			singleDayString(0) = Replace(singleDayString(0), "{", "")
			'remove closing bracket from day string
			singleDayString(2) = Replace(singleDayString(2), "}", "")
			
			dayNumber = cInt(Right(singleDayString(0), 1))
			startTime = Right(singleDayString(1), 5)
			endTime = Right(singleDayString(2), 5)
			
			SQLInsert = "INSERT INTO SC_UserRestrictedLoginSchedule (UserNo, DayNo, StartRestrictedTime, EndRestrictedTime) VALUES "
			SQLInsert = SQLInsert & " (" & userNo & "," & dayNumber & ",'" & startTime & "','" & endTime & "')"
			
			Set rsInsert = cnnInsert.Execute(SQLInsert)
			
		next

	
	Else
		'Then only one day has been selected
		
		singleDayString = Split(jsonString,",")
		
		''singleDayString[0] = Contains Day Number
		''singleDayString[1] = Contains Day Number Restricted Start Time
		''singleDayString[2] = Contains Day Number Restricted End Time

		'remove opening bracket from day string
		singleDayString(0) = Replace(singleDayString(0), "{", "")
		'remove closing bracket from day string
		singleDayString(2) = Replace(singleDayString(2), "}", "")

		dayNumber = cInt(Right(singleDayString(0), 1))
		startTime = Right(singleDayString(1), 5)
		endTime = Right(singleDayString(2), 5)
		
		SQLInsert = "INSERT INTO SC_UserRestrictedLoginSchedule (UserNo, DayNo, StartRestrictedTime, EndRestrictedTime) VALUES "
		SQLInsert = SQLInsert & " (" & userNo & "," & dayNumber & ",'" & startTime & "','" & endTime & "')"
		
		Set rsInsert = cnnInsert.Execute(SQLInsert)

	End If
	
	cnnInsert.close
	
	'********************************************************************

	'day:0,start:00:00,end:24:00
	'[{"day":0,"start":"00:00","end":"24:00"}]
	'[{"day":0,"start":"00:00","end":"24:00"},{"day":6,"start":"01:00","end":"24:00"}]
	'[{"day":0,"start":"00:00","end":"24:00"},{"day":3,"start":"00:00","end":"08:00"},{"day":3,"start":"13:00","end":"19:00"},{"day":6,"start":"00:00","end":"24:00"}]

End Sub

'********************************************************************************************************************************************************
'********************************************************************************************************************************************************



'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

Sub UpdateLoginAccessForNewUser()

	jsonString = Request.Form("jsonString") 
	
	'Response.Write(jsonString)
	
	'********************************************************************
	'When a user selects new login restricted access times, we are rebuilding ALL records in SC_UserRestrictedLoginSchedule,
	'so we need to delete all existing records first
	'********************************************************************
	
	SQLDelete = "DELETE FROM SC_UserRestrictedLoginSchedule WHERE userNo = -1"
	
	Set cnnDelete = Server.CreateObject("ADODB.Connection")
	cnnDelete.open (Session("ClientCnnString"))
	Set rsDelete = Server.CreateObject("ADODB.Recordset")
	rsDelete.CursorLocation = 3 
	Set rsDelete = cnnDelete.Execute(SQLDelete)
	cnnDelete.close
	
	'********************************************************************
	'Prepare the jsonString for parsing by removing extraneous characters
	'********************************************************************
	'remove the opening [ in the string
	jsonString = Right(jsonString, Len(jsonString) - 1)
	
	'remove the closing ] in the string
	jsonString = Left(jsonString,len(jsonString)-1)
	
	'remove double quotes from the string
	jsonString = Replace(jsonString, """","")
	
	'********************************************************************


	'********************************************************************
	'Now build the new login records and insert them into SC_UserRestrictedLoginSchedule
	
	Set cnnInsert = Server.CreateObject("ADODB.Connection")
	cnnInsert.open (Session("ClientCnnString"))
	Set rsInsert = Server.CreateObject("ADODB.Recordset")
	rsInsert.CursorLocation = 3 

	If InStr(jsonString,"},{") Then
	
		'Multiple days have been selected with restricted login access
		
		jsonArray = Split(jsonString, "},{")
		
		for i = 0 to Ubound(jsonArray)
		
			singleDayString = Split(jsonArray(i),",")
			
			''singleDayString[0] = Contains Day Number
			''singleDayString[1] = Contains Day Number Restricted Start Time
			''singleDayString[2] = Contains Day Number Restricted End Time
			
			'remove opening bracket from day string
			singleDayString(0) = Replace(singleDayString(0), "{", "")
			'remove closing bracket from day string
			singleDayString(2) = Replace(singleDayString(2), "}", "")
			
			dayNumber = cInt(Right(singleDayString(0), 1))
			startTime = Right(singleDayString(1), 5)
			endTime = Right(singleDayString(2), 5)
			
			SQLInsert = "INSERT INTO SC_UserRestrictedLoginSchedule (UserNo, DayNo, StartRestrictedTime, EndRestrictedTime) VALUES "
			SQLInsert = SQLInsert & " (-1," & dayNumber & ",'" & startTime & "','" & endTime & "')"
			
			Set rsInsert = cnnInsert.Execute(SQLInsert)
			
		next

	
	Else
		'Then only one day has been selected
		
		singleDayString = Split(jsonString,",")
		
		''singleDayString[0] = Contains Day Number
		''singleDayString[1] = Contains Day Number Restricted Start Time
		''singleDayString[2] = Contains Day Number Restricted End Time

		'remove opening bracket from day string
		singleDayString(0) = Replace(singleDayString(0), "{", "")
		'remove closing bracket from day string
		singleDayString(2) = Replace(singleDayString(2), "}", "")

		dayNumber = cInt(Right(singleDayString(0), 1))
		startTime = Right(singleDayString(1), 5)
		endTime = Right(singleDayString(2), 5)
		
		SQLInsert = "INSERT INTO SC_UserRestrictedLoginSchedule (UserNo, DayNo, StartRestrictedTime, EndRestrictedTime) VALUES "
		SQLInsert = SQLInsert & " (-1," & dayNumber & ",'" & startTime & "','" & endTime & "')"
		
		Set rsInsert = cnnInsert.Execute(SQLInsert)

	End If
	
	cnnInsert.close
	
	'********************************************************************

	'day:0,start:00:00,end:24:00
	'[{"day":0,"start":"00:00","end":"24:00"}]
	'[{"day":0,"start":"00:00","end":"24:00"},{"day":6,"start":"01:00","end":"24:00"}]
	'[{"day":0,"start":"00:00","end":"24:00"},{"day":3,"start":"00:00","end":"08:00"},{"day":3,"start":"13:00","end":"19:00"},{"day":6,"start":"00:00","end":"24:00"}]

End Sub

'********************************************************************************************************************************************************
'********************************************************************************************************************************************************


'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

Sub ClearLoginAccessForExistingUser()

	userNo = Request.Form("userNo") 
	
	Set rsSaveEquivCustID = Server.CreateObject("ADODB.Recordset")
	rsSaveEquivCustID.CursorLocation = 3 
	
	If userNo <> "" Then

		SQLDelete = "DELETE FROM SC_UserRestrictedLoginSchedule WHERE UserNo = " & userNo

		Set cnnDelete = Server.CreateObject("ADODB.Connection")
		cnnDelete.open (Session("ClientCnnString"))
		Set rsDelete = Server.CreateObject("ADODB.Recordset")
		rsDelete.CursorLocation = 3 
		Set rsDelete = cnnDelete.Execute(SQLDelete)
		cnnDelete.close
		
		Response.Write("Success")
		
	Else
		Response.Write("Cannot Clear User Access Table, Invalid Data")
		
	End If

End Sub


'********************************************************************************************************************************************************
'********************************************************************************************************************************************************



'********************************************************************************************************************************************************
'********************************************************************************************************************************************************


Sub ClearLoginAccessNewForUser()
	
	Set rsSaveEquivCustID = Server.CreateObject("ADODB.Recordset")
	rsSaveEquivCustID.CursorLocation = 3 

	SQLDelete = "DELETE FROM SC_UserRestrictedLoginSchedule WHERE UserNo = -1"

	Set cnnDelete = Server.CreateObject("ADODB.Connection")
	cnnDelete.open (Session("ClientCnnString"))
	Set rsDelete = Server.CreateObject("ADODB.Recordset")
	rsDelete.CursorLocation = 3 
	Set rsDelete = cnnDelete.Execute(SQLDelete)
	cnnDelete.close
	
	Response.Write("Success")
		

End Sub


Sub updateCustomOrDefault()
    emailid = Request.Form("id") 
    emailtype = Request.Form("type") 
	
	SQLUpdate = "UPDATE SC_EmailCustomization SET customOrDefault='" & emailtype & "' WHERE InternalRecordIdentifier = " & emailid

	Set cnnUpdate = Server.CreateObject("ADODB.Connection")
	cnnUpdate.open (Session("ClientCnnString"))
	Set rsUpdate = Server.CreateObject("ADODB.Recordset")
	rsUpdate.CursorLocation = 3 
	Set rsUpdate = cnnUpdate.Execute(SQLUpdate)
	cnnUpdate.close
	
	Response.Write("Success")
		

End Sub


'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

'********************************************************************************************************************************************************

'END ALL AJAX MODAL SUBROUTINES AND FUNCTIONS

'********************************************************************************************************************************************************
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

%>