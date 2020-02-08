<!--#include file="../../../inc/InSightFuncs.asp"-->
<!--#include file="../../../inc/InSightFuncs_Users.asp"-->
<% If Session("Userno") = "" Then Response.End() %>
<%

Response.ContentType = "application/json"
CustomerID = Request.QueryString("cid")
InternalRecordIdentifier = Request.QueryString("i")  
If CustomerID = "" Then Response.End()



Set cnn = Server.CreateObject("ADODB.Connection")
cnn.open (Session("ClientCnnString"))

If Request.Form("updateAction")="save" Then

	'***************************************************************************************
	'Lookup the record as it exists now so we can fillin the audit trail
	'***************************************************************************************
	SQL = "SELECT * FROM AR_CustomerBillTo WHERE InternalRecordIdentifier='"&Request.Form("updateActionId")&"'"		
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rs = cnn8.Execute(SQL)
		
	If not rs.EOF Then
	
		Orig_DefaultBillTo = rs("DefaultBillTo") 
		Orig_BillToCompany = rs("BillName") 
		Orig_BillToContactFirstName = rs("ContactFirstName") 
		Orig_BillToContactLastName = rs("ContactLastName")    
		Orig_BillToAddress1 = rs("Addr1")
		Orig_BillToAddress2 = rs("Addr2")
		Orig_BillToCity = rs("City")
		Orig_BillToState = rs("State")
		Orig_BillToZip = rs("Zip")
		Orig_BillToCountry = rs("Country")
		Orig_BillToEmail = rs("Email") 
		Orig_BillToPhone = rs("Phone") 
		Orig_BillToFax = rs("Fax")		
	
	End If

	If (Orig_DefaultBillTo <> "" AND Orig_DefaultBillTo <> 0) Then Orig_DefaultBillTo = 1 Else Orig_DefaultBillTo = 0


	Query = "UPDATE AR_CustomerBillTo SET "
	
	If Request.Form("DefaultBillTo") = 1 Then
		Query = Query & "DefaultBillTo=1, "
	Else
		Query = Query & "DefaultBillTo=0, "
	End If
	
	Query = Query & "BillName='"&Request.Form("BillToCompany")&"', "
	Query = Query & "ContactFirstName='"&Request.Form("BillToContactFirstName")&"', "
	Query = Query & "ContactLastName='"&Request.Form("BillToContactLastName")&"', "	
	Query = Query & "Addr1='"&EscapeSingleQuotes(Request.Form("BillToAddress1"))&"', "
	Query = Query & "Addr2='"&EscapeSingleQuotes(Request.Form("BillToAddress2"))&"', "
	Query = Query & "City='"&EscapeSingleQuotes(Request.Form("BillToCity"))&"', "
	Query = Query & "State='"&EscapeSingleQuotes(Request.Form("BillToState"))&"', "
	Query = Query & "Zip='"&EscapeSingleQuotes(Request.Form("BillToZip"))&"', "
	Query = Query & "Country='"&EscapeSingleQuotes(Request.Form("BillToCountry"))&"', "
	Query = Query & "Email='"&Request.Form("BillToEmail")&"', "
	Query = Query & "Phone='"&Request.Form("BillToPhone")&"', "
	Query = Query & "Fax='"&Request.Form("BillToFax")&"' "
	Query = Query & "WHERE InternalRecordIdentifier='"&Request.Form("updateActionId")&"'"
	cnn.Execute(Query)
	
	'Response.Write(Query)
	
	If Request.Form("DefaultBillTo") = 1 Then
		Query = "UPDATE AR_CustomerBillTo SET DefaultBillTo = 0 WHERE InternalRecordIdentifier <>'" & Request.Form("updateActionId") & "'"
		cnn.Execute(Query)
	End If
	

	'***************************************************************************************
	'After SQL update, record entries in audit trail
	'***************************************************************************************
	
	DefaultBillTo			= Request.Form("DefaultBillTo")
	BillToCompany			= Request.Form("BillToCompany")
	BillToContactFirstName	= Request.Form("BillToContactFirstName")
	BillToContactLastName	= Request.Form("BillToContactLastName")
	BillToAddress1			= Request.Form("BillToAddress1")
	BillToAddress2			= Request.Form("BillToAddress2")
	BillToCity				= Request.Form("BillToCity")
	BillToState				= Request.Form("BillToState")
	BillToZip				= Request.Form("BillToZip")
	BillToCountry			= Request.Form("BillToCountry")
	BillToEmail				= Request.Form("BillToEmail")
	BillToPhone				= Request.Form("BillToPhone")
	BillToFax				= Request.Form("BillToFax")

	If (DefaultBillTo <> "" AND DefaultBillTo <> 0) Then DefaultBillTo = 1 Else DefaultBillTo = 0

	ContactName = BillToContactFirstName & " " & BillToContactLastName
	
	'***********************************************************************
	'End Lookup the record as it exists now so we can fillin the audit trail
	'***********************************************************************

	Description = ""
	
	
	If Orig_DefaultBillTo <> DefaultBillTo Then
		If (DefaultBillTo = 1 OR DefaultBillTo = vbTrue) Then
			Description = ContactName & " was set to be the primary bill to location for the customer account " & CustomerID & ", " & BillToCompany
			CreateAuditLogEntry GetTerm("Accounts Receivable") & " customer default billing location change ",GetTerm("Accounts Receivable"),"Minor",0,Description
		Else
			Description = ContactName & " was un-set as the primary contact for the customer account " & CustomerID
			CreateAuditLogEntry GetTerm("Accounts Receivable") & " customer default billing location change ",GetTerm("Accounts Receivable"),"Minor",0,Description
		End If
	End If
	
	If Orig_BillToCompany <> BillToCompany Then
		Description =  GetTerm("Accounts Receivable") & " customer bill to location company name changed from " & Orig_BillToCompany & " to " & BillToCompany & " for contact " & ContactName & " for the customer account " & CustomerID & ", " & BillToCompany
		CreateAuditLogEntry GetTerm("Accounts Receivable") & " customer bill to location company name change ",GetTerm("Accounts Receivable"),"Minor",0,Description
	End If

	If Orig_BillToFirstName <> BillToFirstName Then
		Description =  GetTerm("Accounts Receivable") & " customer bill to location contact first name changed from " & Orig_BillToFirstName & " to " & BillToFirstName & " for contact " & ContactName & " for the customer account " & CustomerID & ", " & BillToCompany
		CreateAuditLogEntry GetTerm("Accounts Receivable") & " customer bill to location contact first name change ",GetTerm("Accounts Receivable"),"Minor",0,Description
	End If
	
	If Orig_BillToLastName <> BillToLastName Then
		Description =  GetTerm("Accounts Receivable") & " customer bill to location contact last name changed from " & Orig_BillToLastName & " to " & BillToLastName & " for contact " & ContactName & " for the customer account " & CustomerID & ", " & BillToCompany
		CreateAuditLogEntry GetTerm("Accounts Receivable") & " customer bill to location contact last name change ",GetTerm("Accounts Receivable"),"Minor",0,Description
	End If

	If Orig_BillToAddress1 <> BillToAddress1 Then
		Description =  GetTerm("Accounts Receivable") & " customer bill to location Address1 changed from " & Orig_BillToAddress1  & " to " & BillToAddress1 & " for contact " & ContactName & " for the customer account " & CustomerID & ", " & BillToCompany
		CreateAuditLogEntry GetTerm("Accounts Receivable") & " customer bill to location Address1 change ",GetTerm("Accounts Receivable"),"Minor",0,Description
	End If
	
	If Orig_BillToAddress2 <> BillToAddress2 Then
		Description =  GetTerm("Accounts Receivable") & " customer bill to location Address2 changed from " & Orig_BillToAddress2  & " to " & BillToAddress2 & " for contact " & ContactName & " for the customer account " & CustomerID & ", " & BillToCompany
		CreateAuditLogEntry GetTerm("Accounts Receivable") & " customer bill to location Address2 change ",GetTerm("Accounts Receivable"),"Minor",0,Description
	End If
	
	If Orig_BillToCity <> BillToCity Then
		Description =  GetTerm("Accounts Receivable") & " customer bill to location City changed from " & Orig_BillToCity  & " to " & BillToCity & " for contact " & ContactName & " for the customer account " & CustomerID & ", " & BillToCompany
		CreateAuditLogEntry GetTerm("Accounts Receivable") & " customer bill to location City change ",GetTerm("Accounts Receivable"),"Minor",0,Description
	End If
	
	If Orig_BillToState <> BillToState Then
		Description =  GetTerm("Accounts Receivable") & " customer bill to location State changed from " & Orig_BillToState  & " to " & BillToState & " for contact " & ContactName & " for the customer account " & CustomerID & ", " & BillToCompany
		CreateAuditLogEntry GetTerm("Accounts Receivable") & " customer bill to location State change ",GetTerm("Accounts Receivable"),"Minor",0,Description
	End If
	
	If Orig_BillToZip <> BillToZip Then
		Description =  GetTerm("Accounts Receivable") & " customer bill to location Postal Code changed from " & Orig_BillToZip & " to " & BillToZip & " for contact " & ContactName & " for the customer account " & CustomerID & ", " & BillToCompany
		CreateAuditLogEntry GetTerm("Accounts Receivable") & " customer bill to location zip code change ",GetTerm("Accounts Receivable"),"Minor",0,Description
	End If		
	
	If Orig_BillToCountry <> BillToCountry Then
		Description =  GetTerm("Accounts Receivable") & " customer bill to location Country changed from " & Orig_BillToCountry  & " to " & BillToCountry & " for contact " & ContactName & " for the customer account " & CustomerID & ", " & BillToCompany
		CreateAuditLogEntry GetTerm("Accounts Receivable") & " customer bill to location Country change ",GetTerm("Accounts Receivable"),"Minor",0,Description
	End If				
		
	If Orig_BillToEmail <> BillToEmail Then
		Description =  GetTerm("Accounts Receivable") & " customer bill to location email address changed from " & Orig_BillToEmail & " to " & BillToEmail & " for contact " & ContactName & " for the customer account " & CustomerID & ", " & BillToCompany
		CreateAuditLogEntry GetTerm("Accounts Receivable") & " customer bill to location email address change ",GetTerm("Accounts Receivable"),"Minor",0,Description
	End If
	
	If Orig_BillToPhone <> BillToPhone Then
		Description =  GetTerm("Accounts Receivable") & " customer bill to location phone number changed from " & Orig_BillToPhone & " to " & BillToPhone & " for contact " & ContactName & " for the customer account " & CustomerID & ", " & BillToCompany
		CreateAuditLogEntry GetTerm("Accounts Receivable") & " customer bill to location phone number change ",GetTerm("Accounts Receivable"),"Minor",0,Description
	End If
	
	If Orig_BillToFax <> BillToFax Then
		Description =  GetTerm("Accounts Receivable") & " customer bill to location fax number changed from " & Orig_BillToFax & " to " & Fax & " for contact " & ContactName & " for the customer account " & CustomerID & ", " & BillToCompany
		CreateAuditLogEntry GetTerm("Accounts Receivable") & " customer bill to location fax number change ",GetTerm("Accounts Receivable"),"Minor",0,Description
	End If

	
End If






If Request.Form("updateAction")="insert" Then


	If Request.Form("DefaultBillTo") = 1 Then
		PContVar=1
	Else
		PContVar=0
	End If
	
	If Request.Form("DefaultBillTo") = 1 Then
		Query = "UPDATE AR_CustomerBillTo SET DefaultBillTo = 0 WHERE InternalRecordIdentifier <> '" & Request.Form("updateActionId") & "'"
		cnn.Execute(Query)
	End If

	Query = "INSERT INTO AR_CustomerBillTo (CustNum, DefaultBillTo, BillName, ContactFirstName, ContactLastName, Addr1, Addr2, City, State, Zip, Country, Email, Phone, Fax) "
	Query = Query & " VALUES "
	Query = Query & "('" & CustomerID & "'," & PContVar
	Query = Query & ",'" & Request.Form("BillToCompany") & "'"
	Query = Query & ",'" & Request.Form("BillToContactFirstName") & "'"
	Query = Query & ",'" & Request.Form("BillToContactLastName") & "'"
	Query = Query & ",'"&  EscapeSingleQuotes(Request.Form("BillToAddress1")) &"'"
	Query = Query & ",'"&  EscapeSingleQuotes(Request.Form("BillToAddress2")) &"'"
	Query = Query & ",'"&  EscapeSingleQuotes(Request.Form("BillToCity")) &"'"
	Query = Query & ",'"&  EscapeSingleQuotes(Request.Form("BillToState")) &"'"
	Query = Query & ",'"&  EscapeSingleQuotes(Request.Form("BillToZip")) &"'"
	Query = Query & ",'"&  EscapeSingleQuotes(Request.Form("BillToCountry")) &"'"
	Query = Query & ",'"&  EscapeSingleQuotes(Request.Form("BillToEmail")) &"'"
	Query = Query & ",'"&  EscapeSingleQuotes(Request.Form("BillToPhone")) &"'"
	Query = Query & ",'"&  EscapeSingleQuotes(Request.Form("BillToFax")) &"'"
	Query = Query & ")"
	cnn.Execute(Query)
	
	ContactName = Request.Form("FirstName") & " " & Request.Form("LastName")
	
	If (Request.Form("DefaultBillTo") = 1 OR Request.Form("DefaultBillTo") = vbTrue) Then
		Description = ContactName & " was added as the default billing location/contact for customer account " & CustomerID & ", " & BillToCompany
		CreateAuditLogEntry GetTerm("Accounts Receivable") & " billing location added ",GetTerm("Accounts Receivable"),"Minor",0,Description
	Else
		Description = ContactName & " was added as a billing location/contact for customer account " & CustomerID & ", " & BillToCompany
		CreateAuditLogEntry GetTerm("Accounts Receivable") & " billing location added ",GetTerm("Accounts Receivable"),"Minor",0,Description
	End If
	
End If





If Request.Form("updateAction")="delete" Then

	SQL = "SELECT * FROM AR_CustomerBillTo WHERE InternalRecordIdentifier='"&Request.Form("updateActionId")&"'"		
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rs = cnn8.Execute(SQL)
	
	If not rs.EOF Then
		DefaultBillTo  = rs("DefaultBillTo")
		CustomerID = rs("CustNum")
		BillToCompanyName = rs("BillName")
		ContactFirstName = rs("ContactFirstName")
		ContactLastName	= rs("ContactLastName")
		ContactName = ContactFirstName & " " & ContactLastName
	End If
	
	set rs = Nothing
	cnn8.close
	set cnn8 = Nothing


	If (DefaultBillTo  = 1 OR DefaultBillTo  = vbTrue) Then
		Description = "The default billing location/contact " & ContactName & " was removed from customer account " & CustomerID & ", " & BillToCompanyName
		CreateAuditLogEntry GetTerm("Accounts Receivable") & " customer billing location deleted",GetTerm("Accounts Receivable"),"Minor",0,Description
	Else
		Description = "The billing location/contact " & ContactName & " was removed from customer account " & CustomerID & ", " & BillToCompanyName
		CreateAuditLogEntry GetTerm("Accounts Receivable") & " customer billing location deleted",GetTerm("Accounts Receivable"),"Minor",0,Description
	End If

	Query = "DELETE FROM AR_CustomerBillTo WHERE InternalRecordIdentifier ='"&Request.Form("updateActionId")&"'"
	cnn.Execute(Query)
	
End If

If Request.Form("updateAction")="Sticky-1" Then
	Query = "UPDATE AR_CustomerBillTo SET DefaultBillTo=1 WHERE InternalRecordIdentifier ='"&Request.Form("updateActionId")&"'"
	cnn.Execute(Query)
End If

If Request.Form("updateAction")="Sticky-0" Then
	Query = "UPDATE AR_CustomerBillTo SET DefaultBillTo=0 WHERE InternalRecordIdentifier ='"&Request.Form("updateActionId")&"'"
	cnn.Execute(Query)
End If


Query = "SELECT *, InternalRecordIdentifier as id FROM AR_CustomerBillTo WHERE CustNum = '" & CustomerID & "' ORDER BY DefaultBillTo Desc, ContactLastName"
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
			Response.Write("""id"":""" & EscapeQuotes(rs("InternalRecordIdentifier")) & """")
			If rs("DefaultBillTo") = 1 Then
				Response.Write(",""DefaultBillTo"":1")
			Else
				Response.Write(",""DefaultBillTo"":0")
			End If
			Response.Write(",""BillToCustNum"":""" & EscapeQuotes(rs("CustNum")) & """")
			Response.Write(",""BillToCompany"":""" & EscapeQuotes(rs("BillName")) & """")
			Response.Write(",""BillToContactFirstName"":""" & EscapeQuotes(rs("ContactFirstName")) & """")
			Response.Write(",""BillToContactLastName"":""" & EscapeQuotes(rs("ContactLastName")) & """")
			Response.Write(",""BillToAddress1"":""" & EscapeQuotes(rs("Addr1")) & """")
			Response.Write(",""BillToAddress2"":""" & EscapeQuotes(rs("Addr2")) & """")
			Response.Write(",""BillToCity"":""" & EscapeQuotes(rs("City")) & """")
			Response.Write(",""BillToState"":""" & EscapeQuotes(rs("State")) & """")
			Response.Write(",""BillToZip"":""" & EscapeQuotes(rs("Zip")) & """")
			Response.Write(",""BillToCountry"":""" & EscapeQuotes(rs("Country")) & """")
			Response.Write(",""BillToEmail"":""" & EscapeQuotes(rs("Email")) & """")
			Response.Write(",""BillToPhone"":""" & EscapeQuotes(rs("Phone")) & """") 
			Response.Write(",""BillToFax"":""" & EscapeQuotes(rs("Fax")) & """")
			
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
