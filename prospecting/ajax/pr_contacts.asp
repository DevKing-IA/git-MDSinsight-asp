<!--#include file="../../inc/InSightFuncs.asp"-->
<!--#include file="../../inc/InSightFuncs_Users.asp"-->
<!--#include file="../../inc/InSightFuncs_Prospecting.asp"-->
<% If Session("Userno") = "" Then Response.End() %>
<%
Response.ContentType = "application/json"
ProspectIntRecID = Request.QueryString("i") 
If ProspectIntRecID = "" Then Response.End()

Set cnn = Server.CreateObject("ADODB.Connection")
cnn.open (Session("ClientCnnString"))
If Request.Form("updateAction")="save" Then


	'***************************************************************************************
	'Lookup the record as it exists now so we can fillin the audit trail
	'***************************************************************************************
	SQL = "SELECT * FROM PR_ProspectContacts WHERE InternalRecordIdentifier='"&Request.Form("updateActionId")&"'"		
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rs = cnn8.Execute(SQL)
		
	If not rs.EOF Then
	
		Orig_DecisionMaker = rs("DecisionMaker") 
		Orig_PrimaryContact = rs("PrimaryContact") 
		Orig_Suffix = rs("Suffix")
		Orig_FirstName = rs("FirstName") 
		Orig_LastName = rs("LastName")  
		Orig_Notes = rs("Notes")
		Orig_ContactTitleNumber = rs("ContactTitleNumber")  
		Orig_Email = rs("Email") 
		Orig_Phone = rs("Phone") 
		Orig_PhoneExt = rs("PhoneExt")
		Orig_Cell = rs("Cell") 
		Orig_DoNotEmail = rs("DoNotEmail") 
		Orig_Fax = rs("Fax")
		Orig_Address1 = rs("Address1")
		Orig_Address2 = rs("Address2")
		Orig_City = rs("City")
		Orig_State = rs("State")
		Orig_PostalCode = rs("PostalCode")
		Orig_Country = rs("Country")
	
	End If

	If (Orig_DecisionMaker <> "" AND Orig_DecisionMaker <> 0) Then Orig_DecisionMaker = 1 Else Orig_DecisionMaker = 0
	If (Orig_PrimaryContact <> "" AND Orig_PrimaryContact <> 0) Then Orig_PrimaryContact = 1 Else Orig_PrimaryContact = 0
	If (Orig_DoNotEmail <> "" AND Orig_DoNotEmail <> 0) Then Orig_DoNotEmail = 1 Else Orig_DoNotEmail = 0

	Query = "UPDATE PR_ProspectContacts SET "
	
	If Request.Form("PrimaryContact") = 1 Then
		Query = Query & "PrimaryContact='1', "
	Else
		Query = Query & "PrimaryContact='0', "
	End If
	
	If Request.Form("DecisionMaker") = 1 Then
		Query = Query & "DecisionMaker='1', "
	Else
		Query = Query & "DecisionMaker='0', "
	End If


	Query = Query & "Suffix='"&Request.Form("Suffix")&"', "
	Query = Query & "FirstName='"&Request.Form("FirstName")&"', "
	Query = Query & "LastName='"&Request.Form("LastName")&"', "	
	Query = Query & "Notes='"&EscapeSingleQuotes(Request.Form("Notes"))&"', "
	Query = Query & "ContactTitleNumber='"&Request.Form("ContactTitleNumber")&"', "
	Query = Query & "Email='"&Request.Form("Email")&"', "
	Query = Query & "Phone='"&Request.Form("Phone")&"', "
	Query = Query & "PhoneExt='"&Request.Form("PhoneExt")&"', "
	Query = Query & "Cell='"&Request.Form("Cell")&"', "
	
	If Request.Form("DoNotEmail") = 1 Then
		Query = Query & "DoNotEmail='1', "
	Else
		Query = Query & "DoNotEmail='0', "
	End If
	
	Query = Query & "Address1='"&EscapeSingleQuotes(Request.Form("Address1"))&"', "
	Query = Query & "Address2='"&EscapeSingleQuotes(Request.Form("Address2"))&"', "
	Query = Query & "City='"&EscapeSingleQuotes(Request.Form("City"))&"', "
	Query = Query & "State='"&EscapeSingleQuotes(Request.Form("State"))&"', "
	Query = Query & "PostalCode='"&EscapeSingleQuotes(Request.Form("PostalCode"))&"', "
	Query = Query & "Country='"&EscapeSingleQuotes(Request.Form("Country"))&"', "
	
	Query = Query & "Fax='"&Request.Form("Fax")&"' "
	Query = Query & "WHERE InternalRecordIdentifier='"&Request.Form("updateActionId")&"'"
	cnn.Execute(Query)
	
	If Request.Form("PrimaryContact") = 1 Then
		Query = "UPDATE PR_ProspectContacts SET PrimaryContact= 1 WHERE InternalRecordIdentifier <>'"&Request.Form("updateActionId")&"'"
		cnn.Execute(Query)
	End If
	

	'***************************************************************************************
	'After SQL update, record entries in audit trail
	'***************************************************************************************
	
	PrimaryContact		= Request.Form("PrimaryContact")
	DecisionMaker		= Request.Form("DecisionMaker")
	Suffix				= Request.Form("Suffix")
	FirstName			= Request.Form("FirstName")
	LastName			= Request.Form("LastName")
	Notes				= Request.Form("Notes")
	ContactTitleNumber	= Request.Form("ContactTitleNumber")
	Email				= Request.Form("Email")
	Phone				= Request.Form("Phone")
	PhoneExt			= Request.Form("PhoneExt")
	Cell				= Request.Form("Cell")
	DoNotEmail			= Request.Form("DoNotEmail")
	Fax					= Request.Form("Fax")
	Address1			= Request.Form("Address1")
	Address2			= Request.Form("Address2")
	City				= Request.Form("City")
	State				= Request.Form("State")
	PostalCode			= Request.Form("PostalCode")
	Country				= Request.Form("County")

	If (PrimaryContact <> "" AND PrimaryContact <> 0) Then PrimaryContact = 1 Else PrimaryContact = 0
	If (DecisionMaker <> "" AND DecisionMaker <> 0) Then DecisionMaker = 1 Else DecisionMaker = 0
	If (DoNotEmail <> "" AND DoNotEmail <> 0) Then DoNotEmail = 1 Else DoNotEmail = 0

	ContactName = FirstName & " " & LastName
	'***********************************************************************
	'End Lookup the record as it exists now so we can fillin the audit trail
	'***********************************************************************

	Description = ""
	
	If Orig_Suffix <> Suffix Then
	
		Description =  "Contact suffix changed from " & Orig_Suffix & " to " & Suffix & " for contact " & ContactName & " for the prospect " & GetProspectNameByNumber(ProspectIntRecID)
		CreateAuditLogEntry GetTerm("Prospecting") & " contact suffix change ",GetTerm("Prospecting"),"Minor",0,Description
		
		Description = "The suffix for the contact " & ContactName & " changed from: <em><strong> " & Orig_Suffix & "</em></strong> to: <em><strong>" & Suffix & "</em></strong>"
		Record_PR_Activity ProspectIntRecID, Description, Session("UserNo")

	End If
	
	If Orig_FirstName <> FirstName Then
	
		Description =  "Contact suffix changed from " & Orig_FirstName & " to " & FirstName & " for contact " & ContactName & " for the prospect " & GetProspectNameByNumber(ProspectIntRecID)
		CreateAuditLogEntry GetTerm("Prospecting") & " contact first name change ",GetTerm("Prospecting"),"Minor",0,Description
		
		Description = "The first name for the contact " & ContactName & " changed from: <em><strong> " & Orig_FirstName & "</em></strong> to: <em><strong>" & FirstName & "</em></strong>"
		Record_PR_Activity ProspectIntRecID, Description, Session("UserNo")

	End If
	
	If Orig_LastName <> LastName Then
	
		Description =  "Contact suffix changed from " & Orig_LastName & " to " & LastName & " for contact " & ContactName & " for the prospect " & GetProspectNameByNumber(ProspectIntRecID)
		CreateAuditLogEntry GetTerm("Prospecting") & " contact last name change ",GetTerm("Prospecting"),"Minor",0,Description
		
		Description = "The last name for the contact " & ContactName & " changed from: <em><strong> " & Orig_LastName & "</em></strong> to: <em><strong>" & LastName & "</em></strong>"
		Record_PR_Activity ProspectIntRecID, Description, Session("UserNo")

	End If
	
	If Orig_ContactTitleNumber <> ContactTitleNumber Then
	
		Description =  "Contact title changed from " & GetContactTitleByNum(Orig_ContactTitleNumber) & " to " & GetContactTitleByNum(ContactTitleNumber) & " for contact " & ContactName & " for the prospect " & GetProspectNameByNumber(ProspectIntRecID)
		CreateAuditLogEntry GetTerm("Prospecting") & " contact title change ",GetTerm("Prospecting"),"Minor",0,Description
		
		Description = "The title for the contact " & ContactName & " changed from: <em><strong> " & GetContactTitleByNum(Orig_ContactTitleNumber) & "</em></strong> to: <em><strong>" & GetContactTitleByNum(ContactTitleNumber) & "</em></strong>"
		Record_PR_Activity ProspectIntRecID, Description, Session("UserNo")

	End If
	
	If Orig_Email <> Email Then
	
		Description =  "Contact email address changed from " & Orig_Email & " to " & Email & " for contact " & ContactName & " for the prospect " & GetProspectNameByNumber(ProspectIntRecID)
		CreateAuditLogEntry GetTerm("Prospecting") & " contact email address change ",GetTerm("Prospecting"),"Minor",0,Description
		
		Description = "The email address for the contact " & ContactName & " changed from: <em><strong> " & Orig_Email & "</em></strong> to: <em><strong>" & Email & "</em></strong>"
		Record_PR_Activity ProspectIntRecID, Description, Session("UserNo")

	End If
	
	If Orig_Phone <> Phone Then
	
		Description =  "Contact phone number changed from " & Orig_Phone & " to " & Phone & " for contact " & ContactName & " for the prospect " & GetProspectNameByNumber(ProspectIntRecID)
		CreateAuditLogEntry GetTerm("Prospecting") & " contact phone number change ",GetTerm("Prospecting"),"Minor",0,Description
		
		Description = "The phone number for the contact " & ContactName & " changed from: <em><strong> " & Orig_Phone & "</em></strong> to: <em><strong>" & Phone & "</em></strong>"
		Record_PR_Activity ProspectIntRecID, Description, Session("UserNo")

	End If
	
	If Orig_PhoneExt <> PhoneExt Then
	
		Description =  "Contact phone number extension changed from " & Orig_PhoneExt & " to " & PhoneExt & " for contact " & ContactName & " for the prospect " & GetProspectNameByNumber(ProspectIntRecID)
		CreateAuditLogEntry GetTerm("Prospecting") & " contact phone number extension change ",GetTerm("Prospecting"),"Minor",0,Description
		
		Description = "The phone number extension for the contact " & ContactName & " changed from: <em><strong> " & Orig_PhoneExt & "</em></strong> to: <em><strong>" & PhoneExt & "</em></strong>"
		Record_PR_Activity ProspectIntRecID, Description, Session("UserNo")

	End If
	
	If Orig_Cell <> Cell Then
	
		Description =  "Contact cell phone number changed from " & Orig_Cell & " to " & Cell & " for contact " & ContactName & " for the prospect " & GetProspectNameByNumber(ProspectIntRecID)
		CreateAuditLogEntry GetTerm("Prospecting") & " contact cell phone number change ",GetTerm("Prospecting"),"Minor",0,Description
		
		Description = "The cell phone number for the contact " & ContactName & " changed from: <em><strong> " & Orig_Cell & "</em></strong> to: <em><strong>" & Cell & "</em></strong>"
		Record_PR_Activity ProspectIntRecID, Description, Session("UserNo")

	End If

	If Orig_Fax <> Fax Then
	
		Description =  "Contact fax number changed from " & Orig_Fax & " to " & Fax & " for contact " & ContactName & " for the prospect " & GetProspectNameByNumber(ProspectIntRecID)
		CreateAuditLogEntry GetTerm("Prospecting") & " contact fax number change ",GetTerm("Prospecting"),"Minor",0,Description
		
		Description = "The fax number for the contact " & ContactName & " changed from: <em><strong> " & Orig_Fax & "</em></strong> to: <em><strong>" & Fax & "</em></strong>"
		Record_PR_Activity ProspectIntRecID, Description, Session("UserNo")

	End If

	
	If Orig_PrimaryContact <> PrimaryContact Then
	
		If (PrimaryContact = 1 OR PrimaryContact = vbTrue) Then
			Description = ContactName & " was set to be the primary contact for the prospect " & GetProspectNameByNumber(ProspectIntRecID)
			CreateAuditLogEntry GetTerm("Prospecting")& " primary contact change ",GetTerm("Prospecting"),"Minor",0,Description
			
			Description = " The primary contact was changed to " & ContactName 
			Record_PR_Activity ProspectIntRecID, Description, Session("UserNo")
		Else
			Description = ContactName & " was un-set as the primary contact for the prospect " & GetProspectNameByNumber(ProspectIntRecID)
			CreateAuditLogEntry GetTerm("Prospecting")& " primary contact change ",GetTerm("Prospecting"),"Minor",0,Description
			
			Description = ContactName & " is no longer the primary contact."
			Record_PR_Activity ProspectIntRecID, Description, Session("UserNo")
		End If
	End If
	
	
	If Orig_DecisionMaker <> DecisionMaker Then
	
		If (DecisionMaker = 1 OR DecisionMaker = vbTrue) Then
			Description = ContactName & " was marked as a decision maker for the prospect " & GetProspectNameByNumber(ProspectIntRecID)
			CreateAuditLogEntry GetTerm("Prospecting")& " decision maker change ",GetTerm("Prospecting"),"Minor",0,Description
			
			Description = ContactName & " was marked as a decision maker "
			Record_PR_Activity ProspectIntRecID, Description, Session("UserNo")
		Else
			Description = ContactName & " is no longer a decision maker for the prospect " & GetProspectNameByNumber(ProspectIntRecID)
			CreateAuditLogEntry GetTerm("Prospecting")& " decision maker change ",GetTerm("Prospecting"),"Minor",0,Description
			
			Description = ContactName & " is not longer a decision maker "
			Record_PR_Activity ProspectIntRecID, Description, Session("UserNo")
		End If
	End If

	
	If Orig_DoNotEmail <> DoNotEmail Then
	
		If (DoNotEmail = 1 OR DoNotEmail = vbTrue) Then
			Description = "Do not email was set for the contact " & ContactName & " for the prospect " & GetProspectNameByNumber(ProspectIntRecID)
			CreateAuditLogEntry GetTerm("Prospecting")& " do not email change ",GetTerm("Prospecting"),"Minor",0,Description
			
			Description = "Do not email was set for " & ContactName 
			Record_PR_Activity ProspectIntRecID, Description, Session("UserNo")
		Else
			Description = "Do not email was un-set for the contact " & ContactName & " for the prospect " & GetProspectNameByNumber(ProspectIntRecID)
			CreateAuditLogEntry GetTerm("Prospecting")& " do not email change ",GetTerm("Prospecting"),"Minor",0,Description
			
			Description = "Do not email was un-set for " & ContactName
			Record_PR_Activity ProspectIntRecID, Description, Session("UserNo")
		End If
	End If
	
	If Orig_Notes  <> Notes Then
	
		Description =  "Contact notes changed from " & Orig_Notes  & " to " & Notes & " for contact " & ContactName & " for the prospect " & GetProspectNameByNumber(ProspectIntRecID)
		CreateAuditLogEntry GetTerm("Prospecting") & " contact notes change ",GetTerm("Prospecting"),"Minor",0,Description
		
		Description = "The notes for the contact " & ContactName & " changed from: <em><strong> " & Orig_Notes  & "</em></strong> to: <em><strong>" & Notes & "</em></strong>"
		Record_PR_Activity ProspectIntRecID, Description, Session("UserNo")

	End If
	
	'address 1
	If Orig_Address1  <> Address1 Then
	
		Description =  "Contact Address1 changed from " & Orig_Address1  & " to " & Address1 & " for contact " & ContactName & " for the prospect " & GetProspectNameByNumber(ProspectIntRecID)
		CreateAuditLogEntry GetTerm("Prospecting") & " contact Address1 change ",GetTerm("Prospecting"),"Minor",0,Description
		
		Description = "The Address1 for the contact " & ContactName & " changed from: <em><strong> " & Orig_Address1 & "</em></strong> to: <em><strong>" & Address1 & "</em></strong>"
		Record_PR_Activity ProspectIntRecID, Description, Session("UserNo")

	End If
	
	If Orig_Address2  <> Address2 Then
	
		Description =  "Contact Address2 changed from " & Orig_Address2  & " to " & Address2 & " for contact " & ContactName & " for the prospect " & GetProspectNameByNumber(ProspectIntRecID)
		CreateAuditLogEntry GetTerm("Prospecting") & " contact Address2 change ",GetTerm("Prospecting"),"Minor",0,Description
		
		Description = "The Address2 for the contact " & ContactName & " changed from: <em><strong> " & Orig_Address2  & "</em></strong> to: <em><strong>" & Address2 & "</em></strong>"
		Record_PR_Activity ProspectIntRecID, Description, Session("UserNo")

	End If
	
	If Orig_City  <> City Then
	
		Description =  "Contact City changed from " & Orig_City  & " to " & City & " for contact " & ContactName & " for the prospect " & GetProspectNameByNumber(ProspectIntRecID)
		CreateAuditLogEntry GetTerm("Prospecting") & " contact City change ",GetTerm("Prospecting"),"Minor",0,Description
		
		Description = "The City for the contact " & ContactName & " changed from: <em><strong> " & Orig_City  & "</em></strong> to: <em><strong>" & City & "</em></strong>"
		Record_PR_Activity ProspectIntRecID, Description, Session("UserNo")

	End If
	
	If Orig_State  <> State Then
	
		Description =  "Contact State changed from " & Orig_State  & " to " & State & " for contact " & ContactName & " for the prospect " & GetProspectNameByNumber(ProspectIntRecID)
		CreateAuditLogEntry GetTerm("Prospecting") & " contact State change ",GetTerm("Prospecting"),"Minor",0,Description
		
		Description = "The State for the contact " & ContactName & " changed from: <em><strong> " & Orig_State  & "</em></strong> to: <em><strong>" & State & "</em></strong>"
		Record_PR_Activity ProspectIntRecID, Description, Session("UserNo")

	End If
	
	If Orig_PostalCode  <> PostalCode Then
	
		Description =  "Contact Postal Code changed from " & Orig_PostalCode  & " to " & PostalCode & " for contact " & ContactName & " for the prospect " & GetProspectNameByNumber(ProspectIntRecID)
		CreateAuditLogEntry GetTerm("Prospecting") & " contact notes change ",GetTerm("Prospecting"),"Minor",0,Description
		
		Description = "The Postal Code for the contact " & ContactName & " changed from: <em><strong> " & Orig_PostalCode  & "</em></strong> to: <em><strong>" & PostalCode & "</em></strong>"
		Record_PR_Activity ProspectIntRecID, Description, Session("UserNo")

	End If		
	
	If Orig_Country  <> Country Then
	
		Description =  "Contact Country changed from " & Orig_Country  & " to " & Country & " for contact " & ContactName & " for the prospect " & GetProspectNameByNumber(ProspectIntRecID)
		CreateAuditLogEntry GetTerm("Prospecting") & " contact Country change ",GetTerm("Prospecting"),"Minor",0,Description
		
		Description = "The Country for the contact " & ContactName & " changed from: <em><strong> " & Orig_Country & "</em></strong> to: <em><strong>" & Country & "</em></strong>"
		Record_PR_Activity ProspectIntRecID, Description, Session("UserNo")

	End If				
		
	
	
End If






If Request.Form("updateAction")="insert" Then


	If Request.Form("PrimaryContact") = 1 Then
		PContVar=1
	Else
		PContVar=0
	End If
	
	If Request.Form("DecisionMaker") = 1 Then
		DecisionMaker=1
	Else
		DecisionMaker=0
	End If

	If Request.Form("DoNotEmail") = 1 Then
		DoNotEmail=1
	Else
		DoNotEmail=0
	End If

	If Request.Form("PrimaryContact") = 1 Then
		Query = "UPDATE PR_ProspectContacts SET PrimaryContact= 0 WHERE InternalRecordIdentifier <>'"&Request.Form("updateActionId")&"'"
		cnn.Execute(Query)
	End If

	Query = "INSERT INTO PR_ProspectContacts (ProspectIntRecID, DecisionMaker, PrimaryContact , Suffix, FirstName, LastName, Notes, ContactTitleNumber, Email, Phone, PhoneExt, Cell, DoNotEmail, Fax, Address1, Address2, City, State, PostalCode, Country) "
	Query = Query & " VALUES "
	Query = Query & "(" & ProspectIntRecID & "," & DecisionMaker & "," & PContVar & ",'" & Request.Form("Suffix") & "','" & Request.Form("FirstName") & "','" & Request.Form("LastName") & "',"
	Query = Query & "'" & EscapeSingleQuotes(Request.Form("Notes")) & "'," & Request.Form("ContactTitleNumber") & ",'" & Request.Form("Email") & "','" & Request.Form("Phone") & "','" & Request.Form("PhoneExt") & "','" & Request.Form("Cell") & "', "
	Query = Query & DoNotEmail & ",'"& Request.Form("Fax") &"'"
	Query = Query & ",'"&  EscapeSingleQuotes(Request.Form("Address1")) &"'"
	Query = Query & ",'"&  EscapeSingleQuotes(Request.Form("Address2")) &"'"
	Query = Query & ",'"&  EscapeSingleQuotes(Request.Form("City")) &"'"
	Query = Query & ",'"&  EscapeSingleQuotes(Request.Form("State")) &"'"
	Query = Query & ",'"&  EscapeSingleQuotes(Request.Form("PostalCode")) &"'"
	Query = Query & ",'"&  EscapeSingleQuotes(Request.Form("Country")) &"'"
	Query = Query & ")"
	cnn.Execute(Query)
	
	ContactName = Request.Form("FirstName") & " " & Request.Form("LastName")
	
	
	If (Request.Form("PrimaryContact") = 1 OR Request.Form("PrimaryContact") = vbTrue) Then
		Description = ContactName & " was added to the contacts for prospect " & GetProspectNameByNumber(ProspectIntRecID) & " and set to be the primary contact"
		CreateAuditLogEntry GetTerm("Prospecting") & " contact added ",GetTerm("Prospecting"),"Minor",0,Description
		
		Description = ContactName & " was added as a contact for this prospect and set to be the primary contact"
		Record_PR_Activity ProspectIntRecID, Description, Session("UserNo")
		
	Else
		Description = ContactName & " was added to the prospect " & GetProspectNameByNumber(ProspectIntRecID) & " as a contact."
		CreateAuditLogEntry GetTerm("Prospecting") & " contact added ",GetTerm("Prospecting"),"Minor",0,Description
		
		Description = ContactName & " was added to this prospect as a contact."
		Record_PR_Activity ProspectIntRecID, Description, Session("UserNo")
	End If
	
End If





If Request.Form("updateAction")="delete" Then

	SQL = "SELECT * FROM PR_ProspectContacts WHERE InternalRecordIdentifier='"&Request.Form("updateActionId")&"'"		
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rs = cnn8.Execute(SQL)
		
	If not rs.EOF Then
		PrimaryContact  = rs("PrimaryContact")
		FirstName		= rs("FirstName")
		LastName		= rs("LastName")
		ContactName 	= FirstName & " " & LastName
	End If
	
	set rs = Nothing
	cnn8.close
	set cnn8 = Nothing


	If (PrimaryContact  = 1 OR PrimaryContact  = vbTrue) Then
	
		Description = "The primary contact " & ContactName & " was removed from the prospect " & GetProspectNameByNumber(ProspectIntRecID)
		CreateAuditLogEntry GetTerm("Prospecting") & " contact removed from prospect ",GetTerm("Prospecting"),"Minor",0,Description
		
		Description = "The primary contact " & ContactName & " was removed from this prospect "
		Record_PR_Activity ProspectIntRecID, Description, Session("UserNo")
	Else
	
		Description = "The contact " & ContactName & " was removed from the prospect " & GetProspectNameByNumber(ProspectIntRecID)
		CreateAuditLogEntry GetTerm("Prospecting") & " contact removed from prospect ",GetTerm("Prospecting"),"Minor",0,Description
		
		Description = "The contact " &  ContactName & " was removed from this prospect."
		Record_PR_Activity ProspectIntRecID, Description, Session("UserNo")
	End If

	Query = "DELETE FROM PR_ProspectContacts WHERE InternalRecordIdentifier ='"&Request.Form("updateActionId")&"'"
	cnn.Execute(Query)
	
End If

If Request.Form("updateAction")="Sticky-1" Then
	Query = "UPDATE PR_ProspectContacts SET PrimaryContact=1 WHERE ProspectIntRecID ='"&Request.Form("updateActionId")&"'"
	cnn.Execute(Query)
End If

If Request.Form("updateAction")="Sticky-0" Then
	Query = "UPDATE PR_ProspectContacts SET PrimaryContact=0 WHERE ProspectIntRecID ='"&Request.Form("updateActionId")&"'"
	cnn.Execute(Query)
End If


Query = "SELECT *, InternalRecordIdentifier as id FROM PR_ProspectContacts WHERE ProspectIntRecID = " & ProspectIntRecID & " ORDER BY PrimaryContact Desc, DecisionMaker Desc, LastName"
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3 
Set rs = cnn.Execute(Query)

Response.Write("[")
If not rs.EOF Then
	sep = ""
	Do While Not rs.EOF
			Response.Write(sep)
			sep = ","
			Response.Write("{")
			Response.Write("""id"":""" & EscapeQuotes(rs("id")) & """")
			If rs("PrimaryContact") = vbTrue Then
				Response.Write(",""PrimaryContact"":1")
			Else
				Response.Write(",""PrimaryContact"":0")
			End If
			If rs("DecisionMaker") = vbTrue Then
				Response.Write(",""DecisionMaker"":1")
			Else
				Response.Write(",""DecisionMaker"":0")
			End If
			Response.Write(",""Suffix"":""" & EscapeQuotes(rs("Suffix")) & """")
			Response.Write(",""FirstName"":""" & EscapeQuotes(rs("FirstName")) & """")
			Response.Write(",""LastName"":""" & EscapeQuotes(rs("LastName")) & """")
			Response.Write(",""Notes"":""" & EscapeQuotes(rs("Notes")) & """")
			Response.Write(",""ContactTitle"":""" & EscapeQuotes(GetContactTitleByNum(rs("ContactTitleNumber"))) & """")
			Response.Write(",""ContactTitleNumber"":""" & EscapeQuotes(rs("ContactTitleNumber")) & """")
			Response.Write(",""Email"":""" & EscapeQuotes(rs("Email")) & """")
			Response.Write(",""Phone"":""" & EscapeQuotes(rs("Phone")) & """")
			Response.Write(",""PhoneExt"":""" & EscapeQuotes(rs("PhoneExt")) & """")
			Response.Write(",""Cell"":""" & EscapeQuotes(rs("Cell")) & """")
			If rs("DoNotEmail") = vbTrue Then
				Response.Write(",""DoNotEmail"":1")
			Else
				Response.Write(",""DoNotEmail"":0")
			End If
			Response.Write(",""Fax"":""" & EscapeQuotes(rs("Fax")) & """")
			Response.Write(",""Address1"":""" & EscapeQuotes(rs("Address1")) & """")
			Response.Write(",""Address2"":""" & EscapeQuotes(rs("Address2")) & """")
			Response.Write(",""City"":""" & EscapeQuotes(rs("City")) & """")
			Response.Write(",""State"":""" & EscapeQuotes(rs("State")) & """")
			Response.Write(",""PostalCode"":""" & EscapeQuotes(rs("PostalCode")) & """")
			Response.Write(",""Country"":""" & EscapeQuotes(rs("Country")) & """")
			Response.Write("}")
		rs.MoveNext						
	Loop
End If
Response.Write("]")
Set rs = Nothing
cnn.Close
Set cnn = Nothing

Function EscapeQuotes(val)
	If val <> "" Then
		EscapeQuotes = Replace(val, """", "\""")
	End If
End Function
Function EscapeSingleQuotes(val)
	If val <> "" Then
		EscapeSingleQuotes = Replace(val, "'", "''")
	End If
End Function

%> 
