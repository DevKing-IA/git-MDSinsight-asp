<!--#include file="../../inc/subsandfuncs.asp"-->
<!--#include file="../../inc/InSightFuncs.asp"-->

<%

accountingPeriodsArray = Split(Request.Form("periodArray"),",")
	
Set rsDelete = Server.CreateObject("ADODB.Recordset")
rsDelete.CursorLocation = 3 
Set cnnDelete = Server.CreateObject("ADODB.Connection")
cnnDelete.open (Session("ClientCnnString"))


For i = 0 to uBound(accountingPeriodsArray)

	IntRecID = cInt(accountingPeriodsArray(i))
	
	Set rsDelete2 = Server.CreateObject("ADODB.Recordset")
	rsDelete2.CursorLocation = 3 

	SQLDelete2 = "SELECT * FROM Settings_AccountingPeriods WHERE InternalRecordIdentifier = " & IntRecID		
	
	Set cnnDelete2 = Server.CreateObject("ADODB.Connection")
	cnnDelete2.open (Session("ClientCnnString"))
	Set rsDelete2 = cnnDelete2.Execute(SQLDelete2)
	
	If NOT rsDelete2.EOF Then
		PeriodYear = rsDelete2("PeriodYear")
		Period = rsDelete2("Period")
		PeriodBeginDate = formatDateTime(rsDelete2("BeginDate"),2)
		PeriodEndDate = formatDateTime(rsDelete2("EndDate"),2)				
	End If
	

	SQLDelete = "DELETE FROM Settings_AccountingPeriods WHERE InternalRecordIdentifier = " & IntRecID
	Set rsDelete = cnnDelete.Execute(SQLDelete)

	Description = GetUserDisplayNameByUserNo(Session("Userno")) & " deleted period " & Period & " in " & PeriodYear & " ranging from " & PeriodBeginDate & " to " & PeriodEndDate & "."
	CreateAuditLogEntry "Company Accounting Period Deleted","Company Accounting Period Deleted","Major",0,Description

Next

Set rsDelete = Nothing
cnnDelete.Close
Set cnnDelete= Nothing

Set rsDelete2 = Nothing
cnnDelete2.Close
Set cnnDelete2 = Nothing


Response.Redirect ("main.asp#accountingperiod")

%>