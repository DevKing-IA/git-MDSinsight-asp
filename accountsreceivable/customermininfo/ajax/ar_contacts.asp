<!--#include file="../../../inc/InSightFuncs.asp"-->
<!--#include file="../../../inc/InSightFuncs_Users.asp"-->
<!--#include file="../../../inc/InSightFuncs_AR_AP.asp"-->
<!--#include file="../../../inc/InSightFuncs_Prospecting.asp"-->
<% If Session("Userno") = "" Then Response.End() %>
<%
Response.ContentType = "application/json"
CustomerID = Request.QueryString("cid")
InternalRecordIdentifier = Request.QueryString("i")  
If CustomerID = "" Then Response.End()


'***************************************************************************************
'Get internal record identifier of customer from AR_Customer
'***************************************************************************************
SQL = "SELECT * FROM AR_Customer WHERE CustNum ='"& CustomerID &"'"		
Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3 
Set rs = cnn8.Execute(SQL)

If not rs.EOF Then
	CustomerIntRecID = rs("InternalRecordIdentifier")
End If


Set cnn = Server.CreateObject("ADODB.Connection")
cnn.open (Session("ClientCnnString"))
If Request.Form("updateAction")="save" Then

	'***************************************************************************************
	'Lookup the record as it exists now so we can fillin the audit trail
	'***************************************************************************************
	SQL = "SELECT * FROM AR_CustomerContacts WHERE InternalRecordIdentifier='"&Request.Form("updateActionId")&"'"		
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
	End If

	If (Orig_DecisionMaker <> "" AND Orig_DecisionMaker <> 0) Then Orig_DecisionMaker = 1 Else Orig_DecisionMaker = 0
	If (Orig_PrimaryContact <> "" AND Orig_PrimaryContact <> 0) Then Orig_PrimaryContact = 1 Else Orig_PrimaryContact = 0
	If (Orig_DoNotEmail <> "" AND Orig_DoNotEmail <> 0) Then Orig_DoNotEmail = 1 Else Orig_DoNotEmail = 0

	Query = "UPDATE AR_CustomerContacts SET "
	
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
		
	Query = Query & "Fax='"&Request.Form("Fax")&"' "
	Query = Query & "WHERE InternalRecordIdentifier='"&Request.Form("updateActionId")&"'"
	cnn.Execute(Query)
	
	If Request.Form("PrimaryContact") = 1 Then
		Query = "UPDATE AR_CustomerContacts SET PrimaryContact= 1 WHERE InternalRecordIdentifier <>'"&Request.Form("updateActionId")&"'"
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

	If (PrimaryContact <> "" AND PrimaryContact <> 0) Then PrimaryContact = 1 Else PrimaryContact = 0
	If (DecisionMaker <> "" AND DecisionMaker <> 0) Then DecisionMaker = 1 Else DecisionMaker = 0
	If (DoNotEmail <> "" AND DoNotEmail <> 0) Then DoNotEmail = 1 Else DoNotEmail = 0

	ContactName = FirstName & " " & LastName
	'***********************************************************************
	'End Lookup the record as it exists now so we can fillin the audit trail
	'***********************************************************************

	Description = ""
	
	If Orig_Suffix <> Suffix Then
		Description =  "Contact suffix changed from " & Orig_Suffix & " to " & Suffix & " for the contact " & ContactName & " assigned to the customer " & GetCustNameByCustNum(CustomerID) & ", account " & CustomerID
		CreateAuditLogEntry GetTerm("Accounts Receivable") & " contact suffix change ",GetTerm("Accounts Receivable"),"Minor",0,Description
	End If
	
	If Orig_FirstName <> FirstName Then
		Description =  "Contact suffix changed from " & Orig_FirstName & " to " & FirstName & " for the contact " & ContactName & " assigned to the customer " & GetCustNameByCustNum(CustomerID) & ", account " & CustomerID
		CreateAuditLogEntry GetTerm("Accounts Receivable") & " contact first name change ",GetTerm("Accounts Receivable"),"Minor",0,Description
	End If
	
	If Orig_LastName <> LastName Then
		Description =  "Contact suffix changed from " & Orig_LastName & " to " & LastName & " for the contact " & ContactName & " assigned to the customer " & GetCustNameByCustNum(CustomerID) & ", account " & CustomerID
		CreateAuditLogEntry GetTerm("Accounts Receivable") & " contact last name change ",GetTerm("Accounts Receivable"),"Minor",0,Description
	End If
	
	If Orig_ContactTitleNumber <> ContactTitleNumber Then
		Description =  "Contact title changed from " & GetContactTitleByNum(Orig_ContactTitleNumber) & " to " & GetContactTitleByNum(ContactTitleNumber) & " for the contact " & ContactName & " assigned to the customer " & GetCustNameByCustNum(CustomerID) & ", account " & CustomerID
		CreateAuditLogEntry GetTerm("Accounts Receivable") & " contact title change ",GetTerm("Accounts Receivable"),"Minor",0,Description
	End If
	
	If Orig_Email <> Email Then
		Description =  "Contact email address changed from " & Orig_Email & " to " & Email & " for the contact " & ContactName & " assigned to the customer " & GetCustNameByCustNum(CustomerID) & ", account " & CustomerID
		CreateAuditLogEntry GetTerm("Accounts Receivable") & " contact email address change ",GetTerm("Accounts Receivable"),"Minor",0,Description
	End If
	
	If Orig_Phone <> Phone Then
		Description =  "Contact phone number changed from " & Orig_Phone & " to " & Phone & " for the contact " & ContactName & " assigned to the customer " & GetCustNameByCustNum(CustomerID) & ", account " & CustomerID
		CreateAuditLogEntry GetTerm("Accounts Receivable") & " contact phone number change ",GetTerm("Accounts Receivable"),"Minor",0,Description
	End If
	
	If Orig_PhoneExt <> PhoneExt Then
		Description =  "Contact phone number extension changed from " & Orig_PhoneExt & " to " & PhoneExt & " for the contact " & ContactName & " assigned to the customer " & GetCustNameByCustNum(CustomerID) & ", account " & CustomerID
		CreateAuditLogEntry GetTerm("Accounts Receivable") & " contact phone number extension change ",GetTerm("Accounts Receivable"),"Minor",0,Description
	End If
	
	If Orig_Cell <> Cell Then
		Description =  "Contact cell phone number changed from " & Orig_Cell & " to " & Cell & " for the contact " & ContactName & " assigned to the customer " & GetCustNameByCustNum(CustomerID) & ", account " & CustomerID
		CreateAuditLogEntry GetTerm("Accounts Receivable") & " contact cell phone number change ",GetTerm("Accounts Receivable"),"Minor",0,Description
	End If

	If Orig_Fax <> Fax Then
		Description =  "Contact fax number changed from " & Orig_Fax & " to " & Fax & " for the contact " & ContactName & " assigned to the customer " & GetCustNameByCustNum(CustomerID) & ", account " & CustomerID
		CreateAuditLogEntry GetTerm("Accounts Receivable") & " contact fax number change ",GetTerm("Accounts Receivable"),"Minor",0,Description
	End If

	
	If Orig_PrimaryContact <> PrimaryContact Then
		If (PrimaryContact = 1 OR PrimaryContact = vbTrue) Then
			Description = ContactName & " was set to be the primary contact assigned to the customer " & GetCustNameByCustNum(CustomerID) & ", account " & CustomerID
			CreateAuditLogEntry GetTerm("Accounts Receivable")& " primary contact change ",GetTerm("Accounts Receivable"),"Minor",0,Description
		Else
			Description = ContactName & " was un-set as the primary contact assigned to the customer " & GetCustNameByCustNum(CustomerID) & ", account " & CustomerID
			CreateAuditLogEntry GetTerm("Accounts Receivable")& " primary contact change ",GetTerm("Accounts Receivable"),"Minor",0,Description
		End If
	End If
	
	
	If Orig_DecisionMaker <> DecisionMaker Then
		If (DecisionMaker = 1 OR DecisionMaker = vbTrue) Then
			Description = ContactName & " was marked as a decision maker assigned to the customer " & GetCustNameByCustNum(CustomerID) & ", account " & CustomerID
			CreateAuditLogEntry GetTerm("Accounts Receivable")& " decision maker change ",GetTerm("Accounts Receivable"),"Minor",0,Description
		Else
			Description = ContactName & " is no longer a decision maker assigned to the customer " & GetCustNameByCustNum(CustomerID) & ", account " & CustomerID
			CreateAuditLogEntry GetTerm("Accounts Receivable")& " decision maker change ",GetTerm("Accounts Receivable"),"Minor",0,Description
		End If
	End If

	
	If Orig_DoNotEmail <> DoNotEmail Then
		If (DoNotEmail = 1 OR DoNotEmail = vbTrue) Then
			Description = "Do not email was set for the contact " & ContactName & " assigned to the customer " & GetCustNameByCustNum(CustomerID) & ", account " & CustomerID
			CreateAuditLogEntry GetTerm("Accounts Receivable")& " do not email change ",GetTerm("Accounts Receivable"),"Minor",0,Description
		Else
			Description = "Do not email was un-set for the contact " & ContactName & " assigned to the customer " & GetCustNameByCustNum(CustomerID) & ", account " & CustomerID
			CreateAuditLogEntry GetTerm("Accounts Receivable")& " do not email change ",GetTerm("Accounts Receivable"),"Minor",0,Description
		End If
	End If
	
	If Orig_Notes  <> Notes Then
		Description =  "Contact notes changed from " & Orig_Notes  & " to " & Notes & " for the contact " & ContactName & " assigned to the customer " & GetCustNameByCustNum(CustomerID) & ", account " & CustomerID
		CreateAuditLogEntry GetTerm("Accounts Receivable") & " contact notes change ",GetTerm("Accounts Receivable"),"Minor",0,Description
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
		Query = "UPDATE AR_CustomerContacts SET PrimaryContact= 0 WHERE InternalRecordIdentifier <>'"&Request.Form("updateActionId")&"'"
		cnn.Execute(Query)
	End If

	Query = "INSERT INTO AR_CustomerContacts (CustomerIntRecID, DecisionMaker, PrimaryContact , Suffix, FirstName, LastName, Notes, ContactTitleNumber, Email, Phone, PhoneExt, Cell, DoNotEmail, Fax) "
	Query = Query & " VALUES "
	Query = Query & "(" & CustomerIntRecID & "," & DecisionMaker & "," & PContVar & ",'" & Request.Form("Suffix") & "','" & Request.Form("FirstName") & "','" & Request.Form("LastName") & "',"
	Query = Query & "'" & EscapeSingleQuotes(Request.Form("Notes")) & "'," & Request.Form("ContactTitleNumber") & ",'" & Request.Form("Email") & "','" & Request.Form("Phone") & "','" & Request.Form("PhoneExt") & "','" & Request.Form("Cell") & "', "
	Query = Query & DoNotEmail & ",'"& Request.Form("Fax") &"'"
	Query = Query & ")"
	cnn.Execute(Query)
	
	ContactName = Request.Form("FirstName") & " " & Request.Form("LastName")
	
	
	If (Request.Form("PrimaryContact") = 1 OR Request.Form("PrimaryContact") = vbTrue) Then
		Description = ContactName & " was added to the contacts for customer " & GetCustNameByCustNum(CustomerID) & ", account " & CustomerID & ", and set to be the primary contact"
		CreateAuditLogEntry GetTerm("Accounts Receivable") & " contact added ",GetTerm("Accounts Receivable"),"Minor",0,Description		
	Else
		Description = ContactName & " was added to the customer " & GetCustNameByCustNum(CustomerID) & ", account " & CustomerID & ", and as a contact."
		CreateAuditLogEntry GetTerm("Accounts Receivable") & " contact added ",GetTerm("Accounts Receivable"),"Minor",0,Description
	End If
	
End If





If Request.Form("updateAction")="delete" Then

	SQL = "SELECT * FROM AR_CustomerContacts WHERE InternalRecordIdentifier='"&Request.Form("updateActionId")&"'"		
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
		Description = "The primary contact " & ContactName & " was removed from the customer " & GetCustNameByCustNum(CustomerID) & ", account " & CustomerID
		CreateAuditLogEntry GetTerm("Accounts Receivable") & " contact removed from customer ",GetTerm("Accounts Receivable"),"Minor",0,Description
	Else
		Description = "The contact " & ContactName & " was removed from the customer " & GetCustNameByCustNum(CustomerID) & ", account " & CustomerID
		CreateAuditLogEntry GetTerm("Accounts Receivable") & " contact removed from customer ",GetTerm("Accounts Receivable"),"Minor",0,Description
	End If

	Query = "DELETE FROM AR_CustomerContacts WHERE InternalRecordIdentifier ='"&Request.Form("updateActionId")&"'"
	cnn.Execute(Query)
	
End If

If Request.Form("updateAction")="Sticky-1" Then
	Query = "UPDATE AR_CustomerContacts SET PrimaryContact=1 WHERE InternalRecordIdentifier ='"&Request.Form("updateActionId")&"'"
	cnn.Execute(Query)
End If

If Request.Form("updateAction")="Sticky-0" Then
	Query = "UPDATE AR_CustomerContacts SET PrimaryContact=0 WHERE InternalRecordIdentifier ='"&Request.Form("updateActionId")&"'"
	cnn.Execute(Query)
End If


Query = "SELECT *, InternalRecordIdentifier as id FROM AR_CustomerContacts WHERE CustomerIntRecID = " & CustomerIntRecID & " ORDER BY PrimaryContact Desc, DecisionMaker Desc, LastName"
Set rs = Server.CreateObject("ADODB.Recordset")
'Response.Write(Query)
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
