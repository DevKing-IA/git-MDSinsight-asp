<!--#include file="../inc/InSightFuncs.asp"-->
<!--#include file="../inc/InSightFuncs_Users.asp"-->
<!--#include file="../inc/InSightFuncs_Prospecting.asp"-->
<%
DIM buffer
buffer=array()

ProspectIDArray = Split(Request.Form("prospectsArray"),",")
	
Set rsExport = Server.CreateObject("ADODB.Recordset")
Set rsExport2 = Server.CreateObject("ADODB.Recordset")
rsExport.CursorLocation = 3 
Set cnnExport = Server.CreateObject("ADODB.Connection")
cnnExport.open (Session("ClientCnnString"))


For i = 0 to uBound(ProspectIDArray)

	ProspectIDNumber = cInt(ProspectIDArray(i))
	ProspectName = GetProspectNameByNumber(ProspectIDNumber)
	
	'Now get all the contacts for that prospect and create an export record for each
	SQLExport = "SELECT FirstName, LastName, Address1, Address2, City, State, PostalCode, Country "
	SQLExport = SQLExport & "FROM PR_ProspectContacts WHERE ProspectIntRecID = " & ProspectIDNumber 

	Set rsExport = cnnExport.Execute(SQLExport)
		
	If Not rsexport.EOF Then
	
		Do While NOT rsExport.EOF
		 
		 	FirstName = rsexport("FirstName")
		 	LastName = rsexport("LastName")
		 	Address1 = rsexport("Address1")
		 	Address2 = rsexport("Address2")
		 	City = rsexport("City")
		 	State = rsexport("State")
		 	PostalCode = rsexport("PostalCode")
		 	Country  = rsexport("Country")
		 	
		 	'If the address info is blank, get the address info from the main prospect record
		 	If ISNull(Address1) Then
			 	SQLExport2 = "SELECT Street, Floor_Suite_Room__c, City, State, PostalCode, Country FROM PR_Prospects WHERE InternalRecordIdentifier = " & ProspectIDNumber 
			 	Set rsExport2 = cnnExport.Execute(SQLExport2)
	
			 	Address1 = rsexport2("Street")
			 	Address2 = rsexport2("Floor_Suite_Room__c")
			 	City = rsexport2("City")
			 	State = rsexport2("State")
			 	PostalCode = rsexport2("PostalCode")
			 	Country  = rsexport2("Country")
			End If

	
			buffer=AddItem(buffer,"""" & FirstName & """," & """" & LastName & """," & """" & ProspectName & """," & """" & Address1 & """," & """" & Address2 & """," & """" & City & """," & """" & State & """," & """" & PostalCode & """," & """" & Country  & """")
			
			rsExport.MoveNext
		Loop

	End If	
		
	Description = GetUserDisplayNameByUserNo(Session("Userno")) & " exported the prospect " & ProspectName
	CreateAuditLogEntry "Prospect exported","Prospect exported","Major",0,Description

	
Next

Set rsExport = Nothing
cnnExport.Close
Set cnnExport= Nothing

'-- the filename you give it will be the one that is shown
' to the users by default when they download

strFile = "prospect_contacts"&Right("0" & Day(Now),2)&+Right("0" & Month(Now),2)&YEAR(Now)&".txt"

Response.Clear
    
' Download to user
Response.AddHeader "Content-Disposition", "attachment; filename=" & strFile
Response.AddHeader "Content-Length", LEN(JOIN(buffer,CHR(13)&CHR(10)))
Response.ContentType = "application/octet-stream"
Response.CharSet = "UTF-8"
'-- send the stream in the response
Response.BinaryWrite(JOIN(buffer,CHR(13)&CHR(10)))


'Response.Redirect ("main.asp")

Function AddItem(arr, val)
    ReDim Preserve arr(UBound(arr) + 1)
    arr(UBound(arr)) = val
    AddItem = arr
End Function


%>