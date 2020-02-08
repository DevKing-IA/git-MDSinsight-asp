<!--#include file="../inc/subsandfuncs.asp"-->
<!--#include file="../inc/InSightFuncs.asp"-->
<!--#include file="../inc/InSightFuncs_Service.asp"-->

<%
section= Request("section")
action= Request("action")
selectedvalue = Request("selectedvalue")

If IsEmpty(selectedvalue) OR IsNull(selectedvalue) OR Not IsNumeric(selectedvalue) Then
	selectedvalue = -1
Else
	selectedvalue = Clng(selectedvalue)
End If

UserCanAddEquipmentSymptomCodesOnTheFly = userCreateEquipmentSymptomCodesOnTheFly(Session("UserNo"))
UserCanAddEquipmentProblemCodesOnTheFly = userCreateEquipmentProblemCodesOnTheFly(Session("UserNo"))
UserCanAddEquipmentResolutionCodesOnTheFly = userCreateEquipmentResolutionCodesOnTheFly(Session("UserNo"))

If section = "txtEquipmentSymptom" Then

    %>
    
    <option value="">Please Select Below</option>
    <% If UserCanAddEquipmentSymptomCodesOnTheFly = vbTrue Then %>
		<option value="-1" style="font-weight:bold"> -- Add a New Symptom -- </option>
    <% End If %>

	<%
	SQL_EQUIP = "SELECT InternalRecordIdentifier,SymptomDescription FROM FS_SymptomCodes ORDER BY SymptomDescription "
	Set cnnEquip = Server.CreateObject("ADODB.Connection")
	cnnEquip.open(Session("ClientCnnString"))

	Set rsEquip = Server.CreateObject("ADODB.Recordset")
	rsEquip.CursorLocation = 3 
	Set rsEquip = cnnEquip.Execute(SQL_EQUIP)
		
	If not rsEquip.EOF Then
		Do While NOT rsEquip.EOF
		
			SymptomDescription = rsEquip("SymptomDescription")
			InternalRecordIdentifier = rsEquip("InternalRecordIdentifier")
			%><option value="<%= InternalRecordIdentifier %>*<%= SymptomDescription %>"><%= SymptomDescription %></option><%
			
			rsEquip.MoveNext
		Loop
	Else
		%><option value="none listed">SYMPTOM NOT FOUND</option><%
	End If
	
	
	set rsEquip = Nothing
	cnnEquip.close

										
End If		



If section = "txtEquipmentProblem" Then

    %>
    
    <option value="">Please Select Below</option>
    <% If UserCanAddEquipmentProblemCodesOnTheFly = vbTrue Then %>
		<option value="-1" style="font-weight:bold"> -- Add a New Problem -- </option>
    <% End If %>

	<%
	SQL_EQUIP = "SELECT InternalRecordIdentifier,ProblemDescription FROM FS_ProblemCodes ORDER BY ProblemDescription "
	Set cnnEquip = Server.CreateObject("ADODB.Connection")
	cnnEquip.open(Session("ClientCnnString"))

	Set rsEquip = Server.CreateObject("ADODB.Recordset")
	rsEquip.CursorLocation = 3 
	Set rsEquip = cnnEquip.Execute(SQL_EQUIP)
		
	If not rsEquip.EOF Then
		Do While NOT rsEquip.EOF
		
			ProblemDescription = rsEquip("ProblemDescription")
			InternalRecordIdentifier = rsEquip("InternalRecordIdentifier")
			%><option value="<%= InternalRecordIdentifier %>*<%= ProblemDescription %>"><%= ProblemDescription %></option><%
			
			rsEquip.MoveNext
		Loop
	Else
		%><option value="none listed">PROBLEM NOT FOUND</option><%
	End If
	
	
	set rsEquip = Nothing
	cnnEquip.close

										
End If		


If section = "txtEquipmentResolution" Then

    %>
    
    <option value="">Please Select Below</option>
    <% If UserCanAddEquipmentResolutionCodesOnTheFly = vbTrue Then %>
		<option value="-1" style="font-weight:bold"> -- Add a New Resolution -- </option>
    <% End If %>

	<%
	SQL_EQUIP = "SELECT InternalRecordIdentifier,ResolutionDescription FROM FS_ResolutionCodes ORDER BY ResolutionDescription"
	Set cnnEquip = Server.CreateObject("ADODB.Connection")
	cnnEquip.open(Session("ClientCnnString"))

	Set rsEquip = Server.CreateObject("ADODB.Recordset")
	rsEquip.CursorLocation = 3 
	Set rsEquip = cnnEquip.Execute(SQL_EQUIP)
		
	If not rsEquip.EOF Then
		Do While NOT rsEquip.EOF
		
			ResolutionDescription = rsEquip("ResolutionDescription")
			InternalRecordIdentifier = rsEquip("InternalRecordIdentifier")
			%><option value="<%= InternalRecordIdentifier %>*<%= ResolutionDescription %>"><%= ResolutionDescription %></option><%
			
			rsEquip.MoveNext
		Loop
	Else
		%><option value="none listed">RESOLUTION NOT FOUND</option><%
	End If
	
	
	set rsEquip = Nothing
	cnnEquip.close

										
End If		

%>