<!--#include file="../../../inc/header.asp"-->

<%

	'***********************************************************
	'Get Values Of All Form Fields Posted
	'***********************************************************
	
	'**********************************************************
	'Now fillup the Terminology vars
	'**********************************************************
	'First find out how many
	'**********************************************************
	
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rs = cnn8.Execute("SELECT COUNT(*) AS TCount FROM SC_Terminology")
	Tcount = rs("TCount")
	ReDim TermArray(TCount)
	cnn8.close
	Set rs = Nothing
	Set cnn = Nothing
	For x = 1 to TCount
		TermArray(x) = Request.Form("txtTerm" & x)
	Next
	
	
	'**********************************************************
	'Update Terminology Table as well
	'**********************************************************
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	
	SQL = "SELECT * FROM SC_Terminology ORDER BY GenericTerm"
	Set rs = cnn8.Execute(SQL)
	
	Set cnn3 = Server.CreateObject("ADODB.Connection")
	cnn3.open (Session("ClientCnnString"))
	Set rs2 = Server.CreateObject("ADODB.Recordset")
	rs2.CursorLocation = 3 

	x=1
	If not rs.eof then
		Do
			If TermArray(x) <> "" Then
				SQL2 = "Update SC_Terminology Set CustomTerm = '" & TermArray(x) & "' Where GenericTerm='" & rs("GenericTerm") & "'"
				Set rs2 = cnn3.Execute(SQL2)
			End If
			rs.movenext
			x=x+1
		Loop while not rs.eof
	End If
	set rs2 = Nothing
	cnn3.close
	set cnn3 = Nothing
	
	
	'************************************************
	' Now do the audit trail entries for Terminology
	'************************************************
	SQL = "SELECT * FROM SC_Terminology ORDER BY GenericTerm"
	Set rs = cnn8.Execute(SQL)

	x = 1

	If not rs.EOF Then
		Do
			If TermArray(x) <> rs("CustomTerm") and TermArray(x) <> "" Then
				CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Minor", 1, "Terminology setting changed. The custom name for the generic term " & rs("GenericTerm") & " changed from  " & rs("CustomTerm") & " to " & TermArray(x)
			End If
			x=x+1
			rs.movenext
		Loop until rs.eof
	End If

	set rs = Nothing
	cnn8.close
	set cnn8 = Nothing
	
	
	Response.Redirect("customize-terminology.asp")	

%>
<!--#include file="../../../inc/footer-main.asp"-->