<!--#include file="../../inc/subsandfuncs.asp"-->
<!--#include file="../../inc/InSightFuncs.asp"-->

<%
section= Request("section")
action= Request("action")
selectedvalue = Request("selectedvalue")

If IsEmpty(selectedvalue) OR IsNull(selectedvalue) OR Not IsNumeric(selectedvalue) Then
	selectedvalue = -1
Else
	selectedvalue = Clng(selectedvalue)
End If

If section = "txtTitle" Then

	SQLContactTitles = "SELECT *, InternalRecordIdentifier as id FROM PR_ContactTitles ORDER BY ContactTitle"
	Set cnnContactTitles = Server.CreateObject("ADODB.Connection")
	cnnContactTitles.open (Session("ClientCnnString"))
	Set rsContactTitles = Server.CreateObject("ADODB.Recordset")
	rsContactTitles.CursorLocation = 3 
	Set rsContactTitles = cnnContactTitles.Execute(SQLContactTitles)
	%>
	    <option value="">Select Job Title</option>
		<option value="-1" style="font-weight:bold"> -- Add a new Job Title -- </option>
	<%
	If not rsContactTitles.EOF Then
	
		Do While Not rsContactTitles.EOF
				%><option value="<%= rsContactTitles("id") %>"  <%If selectedvalue=rsContactTitles("InternalRecordIdentifier") Then Response.Write("selected") End If%>><%= rsContactTitles("ContactTitle") %></option><%
			rsContactTitles.MoveNext						
		Loop
	End If
	Set rsContactTitles = Nothing
	cnnContactTitles.Close
	Set cnnContactTitles = Nothing
									
																	
End If	

If section = "txtTitleforTab" Then

	Set cnnContactTitles = Server.CreateObject("ADODB.Connection")
	cnnContactTitles.open (Session("ClientCnnString"))
									
	SQLContactTitles = "SELECT *, InternalRecordIdentifier as id FROM PR_ContactTitles ORDER BY ContactTitle"
	
	Set rsContactTitles = Server.CreateObject("ADODB.Recordset")
	rsContactTitles.CursorLocation = 3 
	Set rsContactTitles = cnnContactTitles.Execute(SQLContactTitles)
	
	'ContactTitles = ("[")
	ContactTitles = ("[{""id"":""0"",""title"":""Select""},")
	ContactTitles  = ContactTitles & ("{""id"":""-1"",""title"":""Add a new Job Title""},")
	If not rsContactTitles.EOF Then
		sep = ""
		Do While Not rsContactTitles.EOF
				ContactTitles = ContactTitles & (sep)
				sep = ","
				ContactTitles = ContactTitles & ("{")
				ContactTitles = ContactTitles & ("""id"":""" & Replace(rsContactTitles("id"), """", "\""") & """")
				ContactTitles = ContactTitles & (",""title"":""" & Replace(rsContactTitles("ContactTitle"), """", "\""") & """")
				ContactTitles = ContactTitles & ("}")
			rsContactTitles.MoveNext						
		Loop
	End If
	ContactTitles = ContactTitles & ("]")
	Set rsContactTitles = Nothing
	
	cnnContactTitles.Close
	Set cnnContactTitles = Nothing
									
	response.Write(ContactTitles)								
																	
End If	
					
%>