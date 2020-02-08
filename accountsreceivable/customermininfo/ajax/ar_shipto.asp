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
	SQL = "SELECT * FROM AR_CustomerShipTo WHERE InternalRecordIdentifier='"&Request.Form("updateActionId")&"'"		
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rs = cnn8.Execute(SQL)
		
	If not rs.EOF Then
	
		Orig_DefaultShipTo = rs("DefaultShipTo") 
		Orig_ShipToCompany = rs("ShipName") 
		Orig_ShipToContactFirstName = rs("ContactFirstName") 
		Orig_ShipToContactLastName = rs("ContactLastName")    
		Orig_ShipToAddress1 = rs("Addr1")
		Orig_ShipToAddress2 = rs("Addr2")
		Orig_ShipToCity = rs("City")
		Orig_ShipToState = rs("State")
		Orig_ShipToZip = rs("Zip")
		Orig_ShipToCountry = rs("Country")
		Orig_ShipToEmail = rs("Email") 
		Orig_ShipToPhone = rs("Phone") 
		Orig_ShipToFax = rs("Fax")		
	
	End If

	If (Orig_DefaultShipTo <> "" AND Orig_DefaultShipTo <> 0) Then Orig_DefaultShipTo = 1 Else Orig_DefaultShipTo = 0


	Query = "UPDATE AR_CustomerShipTo SET "
	
	If Request.Form("DefaultShipTo") = 1 Then
		Query = Query & "DefaultShipTo=1, "
	Else
		Query = Query & "DefaultShipTo=0, "
	End If
	
	Query = Query & "ShipName='"&Request.Form("ShipToCompany")&"', "
	Query = Query & "ContactFirstName='"&Request.Form("ShipToContactFirstName")&"', "
	Query = Query & "ContactLastName='"&Request.Form("ShipToContactLastName")&"', "	
	Query = Query & "Addr1='"&EscapeSingleQuotes(Request.Form("ShipToAddress1"))&"', "
	Query = Query & "Addr2='"&EscapeSingleQuotes(Request.Form("ShipToAddress2"))&"', "
	Query = Query & "City='"&EscapeSingleQuotes(Request.Form("ShipToCity"))&"', "
	Query = Query & "State='"&EscapeSingleQuotes(Request.Form("ShipToState"))&"', "
	Query = Query & "Zip='"&EscapeSingleQuotes(Request.Form("ShipToZip"))&"', "
	Query = Query & "Country='"&EscapeSingleQuotes(Request.Form("ShipToCountry"))&"', "
	Query = Query & "Email='"&Request.Form("ShipToEmail")&"', "
	Query = Query & "Phone='"&Request.Form("ShipToPhone")&"', "
	Query = Query & "Fax='"&Request.Form("ShipToFax")&"' "
	Query = Query & "WHERE InternalRecordIdentifier='"&Request.Form("updateActionId")&"'"
	cnn.Execute(Query)
	
	'Response.Write(Query)
	
	If Request.Form("DefaultShipTo") = 1 Then
		Query = "UPDATE AR_CustomerShipTo SET DefaultShipTo = 0 WHERE InternalRecordIdentifier <>'" & Request.Form("updateActionId") & "'"
		cnn.Execute(Query)
	End If
	

	'***************************************************************************************
	'After SQL update, record entries in audit trail
	'***************************************************************************************
	
	DefaultShipTo			= Request.Form("DefaultShipTo")
	ShipToCompany			= Request.Form("ShipToCompany")
	ShipToContactFirstName	= Request.Form("ShipToContactFirstName")
	ShipToContactLastName	= Request.Form("ShipToContactLastName")
	ShipToAddress1			= Request.Form("ShipToAddress1")
	ShipToAddress2			= Request.Form("ShipToAddress2")
	ShipToCity				= Request.Form("ShipToCity")
	ShipToState				= Request.Form("ShipToState")
	ShipToZip				= Request.Form("ShipToZip")
	ShipToCountry			= Request.Form("ShipToCountry")
	ShipToEmail				= Request.Form("ShipToEmail")
	ShipToPhone				= Request.Form("ShipToPhone")
	ShipToFax				= Request.Form("ShipToFax")

	If (DefaultShipTo <> "" AND DefaultShipTo <> 0) Then DefaultShipTo = 1 Else DefaultShipTo = 0

	ContactName = ShipToContactFirstName & " " & ShipToContactLastName
	
	'***********************************************************************
	'End Lookup the record as it exists now so we can fillin the audit trail
	'***********************************************************************

	Description = ""
	
	
	If Orig_DefaultShipTo <> DefaultShipTo Then
		If (DefaultShipTo = 1 OR DefaultShipTo = vbTrue) Then
			Description = ContactName & " was set to be the primary ship to location for the customer account " & CustomerID & ", " & ShipToCompany
			CreateAuditLogEntry GetTerm("Accounts Receivable") & " customer default Shipping location change ",GetTerm("Accounts Receivable"),"Minor",0,Description
		Else
			Description = ContactName & " was un-set as the primary contact for the customer account " & CustomerID
			CreateAuditLogEntry GetTerm("Accounts Receivable") & " customer default Shipping location change ",GetTerm("Accounts Receivable"),"Minor",0,Description
		End If
	End If
	
	If Orig_ShipToCompany <> ShipToCompany Then
		Description =  GetTerm("Accounts Receivable") & " customer ship to location company name changed from " & Orig_ShipToCompany & " to " & ShipToCompany & " for contact " & ContactName & " for the customer account " & CustomerID & ", " & ShipToCompany
		CreateAuditLogEntry GetTerm("Accounts Receivable") & " customer ship to location company name change ",GetTerm("Accounts Receivable"),"Minor",0,Description
	End If

	If Orig_ShipToFirstName <> ShipToFirstName Then
		Description =  GetTerm("Accounts Receivable") & " customer ship to location contact first name changed from " & Orig_ShipToFirstName & " to " & ShipToFirstName & " for contact " & ContactName & " for the customer account " & CustomerID & ", " & ShipToCompany
		CreateAuditLogEntry GetTerm("Accounts Receivable") & " customer ship to location contact first name change ",GetTerm("Accounts Receivable"),"Minor",0,Description
	End If
	
	If Orig_ShipToLastName <> ShipToLastName Then
		Description =  GetTerm("Accounts Receivable") & " customer ship to location contact last name changed from " & Orig_ShipToLastName & " to " & ShipToLastName & " for contact " & ContactName & " for the customer account " & CustomerID & ", " & ShipToCompany
		CreateAuditLogEntry GetTerm("Accounts Receivable") & " customer ship to location contact last name change ",GetTerm("Accounts Receivable"),"Minor",0,Description
	End If

	If Orig_ShipToAddress1 <> ShipToAddress1 Then
		Description =  GetTerm("Accounts Receivable") & " customer ship to location Address1 changed from " & Orig_ShipToAddress1  & " to " & ShipToAddress1 & " for contact " & ContactName & " for the customer account " & CustomerID & ", " & ShipToCompany
		CreateAuditLogEntry GetTerm("Accounts Receivable") & " customer ship to location Address1 change ",GetTerm("Accounts Receivable"),"Minor",0,Description
	End If
	
	If Orig_ShipToAddress2 <> ShipToAddress2 Then
		Description =  GetTerm("Accounts Receivable") & " customer ship to location Address2 changed from " & Orig_ShipToAddress2  & " to " & ShipToAddress2 & " for contact " & ContactName & " for the customer account " & CustomerID & ", " & ShipToCompany
		CreateAuditLogEntry GetTerm("Accounts Receivable") & " customer ship to location Address2 change ",GetTerm("Accounts Receivable"),"Minor",0,Description
	End If
	
	If Orig_ShipToCity <> ShipToCity Then
		Description =  GetTerm("Accounts Receivable") & " customer ship to location City changed from " & Orig_ShipToCity  & " to " & ShipToCity & " for contact " & ContactName & " for the customer account " & CustomerID & ", " & ShipToCompany
		CreateAuditLogEntry GetTerm("Accounts Receivable") & " customer ship to location City change ",GetTerm("Accounts Receivable"),"Minor",0,Description
	End If
	
	If Orig_ShipToState <> ShipToState Then
		Description =  GetTerm("Accounts Receivable") & " customer ship to location State changed from " & Orig_ShipToState  & " to " & ShipToState & " for contact " & ContactName & " for the customer account " & CustomerID & ", " & ShipToCompany
		CreateAuditLogEntry GetTerm("Accounts Receivable") & " customer ship to location State change ",GetTerm("Accounts Receivable"),"Minor",0,Description
	End If
	
	If Orig_ShipToZip <> ShipToZip Then
		Description =  GetTerm("Accounts Receivable") & " customer ship to location Postal Code changed from " & Orig_ShipToZip & " to " & ShipToZip & " for contact " & ContactName & " for the customer account " & CustomerID & ", " & ShipToCompany
		CreateAuditLogEntry GetTerm("Accounts Receivable") & " customer ship to location zip code change ",GetTerm("Accounts Receivable"),"Minor",0,Description
	End If		
	
	If Orig_ShipToCountry <> ShipToCountry Then
		Description =  GetTerm("Accounts Receivable") & " customer ship to location Country changed from " & Orig_ShipToCountry  & " to " & ShipToCountry & " for contact " & ContactName & " for the customer account " & CustomerID & ", " & ShipToCompany
		CreateAuditLogEntry GetTerm("Accounts Receivable") & " customer ship to location Country change ",GetTerm("Accounts Receivable"),"Minor",0,Description
	End If				
		
	If Orig_ShipToEmail <> ShipToEmail Then
		Description =  GetTerm("Accounts Receivable") & " customer ship to location email address changed from " & Orig_ShipToEmail & " to " & ShipToEmail & " for contact " & ContactName & " for the customer account " & CustomerID & ", " & ShipToCompany
		CreateAuditLogEntry GetTerm("Accounts Receivable") & " customer ship to location email address change ",GetTerm("Accounts Receivable"),"Minor",0,Description
	End If
	
	If Orig_ShipToPhone <> ShipToPhone Then
		Description =  GetTerm("Accounts Receivable") & " customer ship to location phone number changed from " & Orig_ShipToPhone & " to " & ShipToPhone & " for contact " & ContactName & " for the customer account " & CustomerID & ", " & ShipToCompany
		CreateAuditLogEntry GetTerm("Accounts Receivable") & " customer ship to location phone number change ",GetTerm("Accounts Receivable"),"Minor",0,Description
	End If
	
	If Orig_ShipToFax <> ShipToFax Then
		Description =  GetTerm("Accounts Receivable") & " customer ship to location fax number changed from " & Orig_ShipToFax & " to " & Fax & " for contact " & ContactName & " for the customer account " & CustomerID & ", " & ShipToCompany
		CreateAuditLogEntry GetTerm("Accounts Receivable") & " customer ship to location fax number change ",GetTerm("Accounts Receivable"),"Minor",0,Description
	End If

	
End If






If Request.Form("updateAction")="insert" Then


	If Request.Form("DefaultShipTo") = 1 Then
		PContVar=1
	Else
		PContVar=0
	End If
	
	If Request.Form("DefaultShipTo") = 1 Then
		Query = "UPDATE AR_CustomerShipTo SET DefaultShipTo = 0 WHERE InternalRecordIdentifier <> '" & Request.Form("updateActionId") & "'"
		cnn.Execute(Query)
	End If

	Query = "INSERT INTO AR_CustomerShipTo (CustNum, DefaultShipTo, ShipName, ContactFirstName, ContactLastName, Addr1, Addr2, City, State, Zip, Country, Email, Phone, Fax) "
	Query = Query & " VALUES "
	Query = Query & "('" & CustomerID & "'," & PContVar
	Query = Query & ",'" & Request.Form("ShipToCompany") & "'"
	Query = Query & ",'" & Request.Form("ShipToContactFirstName") & "'"
	Query = Query & ",'" & Request.Form("ShipToContactLastName") & "'"
	Query = Query & ",'"&  EscapeSingleQuotes(Request.Form("ShipToAddress1")) &"'"
	Query = Query & ",'"&  EscapeSingleQuotes(Request.Form("ShipToAddress2")) &"'"
	Query = Query & ",'"&  EscapeSingleQuotes(Request.Form("ShipToCity")) &"'"
	Query = Query & ",'"&  EscapeSingleQuotes(Request.Form("ShipToState")) &"'"
	Query = Query & ",'"&  EscapeSingleQuotes(Request.Form("ShipToZip")) &"'"
	Query = Query & ",'"&  EscapeSingleQuotes(Request.Form("ShipToCountry")) &"'"
	Query = Query & ",'"&  EscapeSingleQuotes(Request.Form("ShipToEmail")) &"'"
	Query = Query & ",'"&  EscapeSingleQuotes(Request.Form("ShipToPhone")) &"'"
	Query = Query & ",'"&  EscapeSingleQuotes(Request.Form("ShipToFax")) &"'"
	Query = Query & ")"
	cnn.Execute(Query)
	
	ContactName = Request.Form("FirstName") & " " & Request.Form("LastName")
	
	If (Request.Form("DefaultShipTo") = 1 OR Request.Form("DefaultShipTo") = vbTrue) Then
		Description = ContactName & " was added as the default Shipping location/contact for customer account " & CustomerID & ", " & ShipToCompany
		CreateAuditLogEntry GetTerm("Accounts Receivable") & " Shipping location added ",GetTerm("Accounts Receivable"),"Minor",0,Description
	Else
		Description = ContactName & " was added as a Shipping location/contact for customer account " & CustomerID & ", " & ShipToCompany
		CreateAuditLogEntry GetTerm("Accounts Receivable") & " Shipping location added ",GetTerm("Accounts Receivable"),"Minor",0,Description
	End If
	
End If





If Request.Form("updateAction")="delete" Then

	SQL = "SELECT * FROM AR_CustomerShipTo WHERE InternalRecordIdentifier='"&Request.Form("updateActionId")&"'"		
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rs = cnn8.Execute(SQL)
	
	If not rs.EOF Then
		DefaultShipTo  = rs("DefaultShipTo")
		CustomerID = rs("CustNum")
		ShipToCompanyName = rs("ShipName")
		ContactFirstName = rs("ContactFirstName")
		ContactLastName	= rs("ContactLastName")
		ContactName = ContactFirstName & " " & ContactLastName
	End If
	
	set rs = Nothing
	cnn8.close
	set cnn8 = Nothing


	If (DefaultShipTo  = 1 OR DefaultShipTo  = vbTrue) Then
		Description = "The default Shipping location/contact " & ContactName & " was removed from customer account " & CustomerID & ", " & ShipToCompanyName
		CreateAuditLogEntry GetTerm("Accounts Receivable") & " customer Shipping location deleted",GetTerm("Accounts Receivable"),"Minor",0,Description
	Else
		Description = "The Shipping location/contact " & ContactName & " was removed from customer account " & CustomerID & ", " & ShipToCompanyName
		CreateAuditLogEntry GetTerm("Accounts Receivable") & " customer Shipping location deleted",GetTerm("Accounts Receivable"),"Minor",0,Description
	End If

	Query = "DELETE FROM AR_CustomerShipTo WHERE InternalRecordIdentifier ='"&Request.Form("updateActionId")&"'"
	cnn.Execute(Query)
	
End If

If Request.Form("updateAction")="Sticky-1" Then
	Query = "UPDATE AR_CustomerShipTo SET DefaultShipTo=1 WHERE InternalRecordIdentifier ='"&Request.Form("updateActionId")&"'"
	cnn.Execute(Query)
End If

If Request.Form("updateAction")="Sticky-0" Then
	Query = "UPDATE AR_CustomerShipTo SET DefaultShipTo=0 WHERE InternalRecordIdentifier ='"&Request.Form("updateActionId")&"'"
	cnn.Execute(Query)
End If


Query = "SELECT *, InternalRecordIdentifier as id FROM AR_CustomerShipTo WHERE CustNum = '" & CustomerID & "' ORDER BY DefaultShipTo Desc, ContactLastName"
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
			If rs("DefaultShipTo") = 1 Then
				Response.Write(",""DefaultShipTo"":1")
			Else
				Response.Write(",""DefaultShipTo"":0")
			End If
			Response.Write(",""ShipToCustNum"":""" & EscapeQuotes(rs("CustNum")) & """")
			Response.Write(",""ShipToCompany"":""" & EscapeQuotes(rs("ShipName")) & """")
			Response.Write(",""ShipToContactFirstName"":""" & EscapeQuotes(rs("ContactFirstName")) & """")
			Response.Write(",""ShipToContactLastName"":""" & EscapeQuotes(rs("ContactLastName")) & """")
			Response.Write(",""ShipToAddress1"":""" & EscapeQuotes(rs("Addr1")) & """")
			Response.Write(",""ShipToAddress2"":""" & EscapeQuotes(rs("Addr2")) & """")
			Response.Write(",""ShipToCity"":""" & EscapeQuotes(rs("City")) & """")
			Response.Write(",""ShipToState"":""" & EscapeQuotes(rs("State")) & """")
			Response.Write(",""ShipToZip"":""" & EscapeQuotes(rs("Zip")) & """")
			Response.Write(",""ShipToCountry"":""" & EscapeQuotes(rs("Country")) & """")
			Response.Write(",""ShipToEmail"":""" & EscapeQuotes(rs("Email")) & """")
			Response.Write(",""ShipToPhone"":""" & EscapeQuotes(rs("Phone")) & """") 
			Response.Write(",""ShipToFax"":""" & EscapeQuotes(rs("Fax")) & """")
			
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
