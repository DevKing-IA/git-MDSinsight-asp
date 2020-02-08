<!--#include file="../../../inc/header.asp"-->

<%

	'***********************************************************
	'Get Values Of All Form Fields Posted
	'***********************************************************
	
	If Request.Form("chkImportCustomersFromQB") = "on" then ImportCustomersFromQB = 1 Else ImportCustomersFromQB = 0

	ImportCustomersUpdateOrReplace = Request.Form("selImportCustomersUpdateOrReplace")


	'***********************************************************
	'Get Original Values For Audit Trail Entries
	'***********************************************************
	
	SQL = "SELECT * FROM Settings_Quickbooks"
	
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rs = cnn8.Execute(SQL)
	
	If not rs.EOF Then
		ImportCustomersFromQB_ORIG = rs("ImportCustomersFromQB")	
		ImportCustomersUpdateOrReplace_ORIG = rs("ImportCustomersUpdateOrReplace")	
	End If
	
	
	set rs = Nothing
	cnn8.close
	set cnn8 = Nothing
	

	'***********************************************************
	'Perform Audit Trail Entry Inserts
	'***********************************************************


	If Request.Form("chkImportCustomersFromQB")="on" then ImportCustomersFromQBMsg = "On" Else ImportCustomersFromQBMsg = "Off"
	If ImportCustomersUpdateOrReplace_ORIG = "U" then ImportCustomersUpdateOrReplace_ORIGMSG = "Update" Else ImportCustomersUpdateOrReplace_ORIGMSG = "Replace"
	
	IF ImportCustomersFromQB_ORIG <> ImportCustomersFromQB  Then
		CreateAuditLogEntry "Quickbooks Integration Settings Change", "Quickbooks Integration Settings Change", "Minor", 1, "Import customers from quickbooks changed from " & ImportCustomersFromQB_ORIG & " to " & ImportCustomersFromQBMsg 
	End If
	
	If ImportCustomersUpdateOrReplace_ORIG <> ImportCustomersUpdateOrReplace Then
		CreateAuditLogEntry "Quickbooks Integration Settings Change", "Quickbooks Integration Settings Change", "Minor", 1, "Update or Replace existing customer table with quickbooks changed from " & ImportCustomersUpdateOrReplace_ORIG & " to " & ImportCustomersUpdateOrReplace 
	End If
	


	'***********************************************************
	'Update SQL with Request Form Field Data
	'***********************************************************

	SQL = "UPDATE  " & MUV_Read("SQL_Owner") & ".Settings_Quickbooks SET  "
	SQL = SQL & "ImportCustomersFromQB = " & ImportCustomersFromQB & ","
	SQL = SQL & "ImportCustomersUpdateOrReplace	= '" & ImportCustomersUpdateOrReplace & "'"
   
	Response.Write("<br><br><br><br>" & SQL)
	
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	
	Set rs = cnn8.Execute(SQL)

	set rs = Nothing
	cnn8.close
	set cnn8 = Nothing
	

	Response.Redirect("../main.asp")
%><!--#include file="../../../inc/footer-main.asp"-->