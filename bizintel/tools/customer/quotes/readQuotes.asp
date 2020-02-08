<!--#include file="../../../../inc/header.asp"-->
<!--#include file="../../../../inc/mail.asp"-->
<%
If Request.QueryString("custID") = "" Then Response.Redirect ("reports.asp") Else custID = Request.QueryString("custID")

'baseURL should alwats have a trailing /slash, just in case, handle either way
If right(baseURL,1)="/" Then maildomain = Left(right(baseURL,len(baseURL)-7),len(right(baseURL,len(baseURL)-7))-1) Else maildomain = right(baseURL,len(baseURL)-7)


'***********************************
'Post the read quotes query to UNIX
'***********************************
data = "<DATASTREAM>"
data = data & "<IDENTITY>Pm8316wyc011</IDENTITY>"
data = data & "<MODE>" & GetPOSTParams("Mode") & "</MODE>"
data = data & "<RECORD_TYPE>QUOTES</RECORD_TYPE>"
data = data & "<RECORD_SUBTYPE>READQUOTE</RECORD_SUBTYPE>"
data = data & "<SERNO>" & GetPOSTParams("Serno") & "</SERNO>"
data = data & "<ACCOUNT_OR_CHAIN>" & "A" & "</ACCOUNT_OR_CHAIN>"
data = data & "<ACCOUNT_NUM>" & custID & "</ACCOUNT_NUM>"
data = data & "<ITEM_QUOTES>Y</ITEM_QUOTES>" ' 1 True, 0 False
data = data & "<CAT_DISCOUNTS>N</CAT_DISCOUNTS>" 
data = data & "</DATASTREAM>"

Description = "Post to " & GetPOSTParams("CUSTOMERURL1")
CreateINSIGHTAuditLogEntry Request.ServerVariables("SERVER_NAME"),Description,GetPOSTParams("Mode")

Description = "data:" & data 
CreateINSIGHTAuditLogEntry Request.ServerVariables("SERVER_NAME"),Description,GetPOSTParams("Mode")

Set httpRequest = Server.CreateObject("MSXML2.ServerXMLHTTP")

httpRequest.Open "POST", GetPOSTParams("CUSTOMERURL1"), False
httpRequest.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"

Err.Clear
On Error Resume Next

httpRequest.Send data

If (Err.Number <> 0 ) Then
	emailbody="httpRequest.status returned " & httpRequest.status & " when posting <RECORD_TYPE>READQUOTES and <RECORD_SUBTYPE>ACCOUNT"& "<br>"
	emailBody = "httpRequest.responseText:" & httpRequest.responseText & "<br>"
	emailBody = emailBody & "PAGE: " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "<br>"
	emailBody = emailBody & "Posted to " & GetPOSTParams("CUSTOMERURL1") & "<br>"
	emailBody = emailBody & "POSTED DATA:" & data & "<br>"
	emailBody = emailBody & "SERNO:" & MUV_READ("ClientID") & "<br>"
	SendMail "mailsender@" & maildomain ,"projects@metroplexdata.com",MUV_READ("ClientID") & " READQUOTES POST ERROR",emailBody, "Price Tool", "Post Failure"

	Description = emailBody 
	CreateINSIGHTAuditLogEntry Request.ServerVariables("SERVER_NAME"),Description,GetPOSTParams("Mode")

	'Go to main with a Querystring, indicationg a read error
	'Doesn't matter what the value is. Any querystring being present
	'indicates a read error
	Response.Redirect("reports.asp?s=0")
End If

On Error Goto 0

	
IF httpRequest.status = 200 THEN 

	Set xmlDoc = Server.CreateObject("Msxml2.DOMDocument.6.0")
	xmlDoc.async = "false"
	xmlDoc.setProperty "ServerHTTPRequest", true
	resltOK = xmlDoc.Loadxml (httpRequest.responseText)

	If resltOK = True Then 'If Instr(httpRequest.responseText,"success") <> 0 Then ' Success
	
	
		'Do SQL prep work to get the tabke ready
		Set cnnQuotedItemsTmp  = Server.CreateObject("ADODB.Connection")
		cnnQuotedItemsTmp.open (Session("ClientCnnString"))
		Set rsQuotedItemsTmp = Server.CreateObject("ADODB.Recordset")
		rsQuotedItemsTmp.CursorLocation = 3 

		On Error Resume Next ' In caase the table isn't there
		SQLQuotedItemsTmp = "DROP TABLE zPRC_AccountQuotedItems_" & trim(Session("Userno")) 
		Set rsQuotedItemsTmp = cnnQuotedItemsTmp.Execute(SQLQuotedItemsTmp)
		On Error Goto 0

		SQLQuotedItemsTmp = "CREATE TABLE zPRC_AccountQuotedItems_" & trim(Session("Userno")) & " ( "
		SQLQuotedItemsTmp =	SQLQuotedItemsTmp & "			[InternalRecordIdentifier] [int] IDENTITY(1,1) NOT NULL, "
		SQLQuotedItemsTmp =	SQLQuotedItemsTmp & "			[RecordCreationDateTime] [datetime] NULL CONSTRAINT [DF_zPRC_AccountQuotedItems_" & trim(Session("Userno")) & "_RecordCreationDateTime]  DEFAULT (getdate()), "
		SQLQuotedItemsTmp =	SQLQuotedItemsTmp & "			[prodSKU] [varchar](255) NULL, "
		SQLQuotedItemsTmp =	SQLQuotedItemsTmp & "			[Description] [varchar](255) NULL, "
		SQLQuotedItemsTmp =	SQLQuotedItemsTmp & "			[Category] [varchar](255) NULL, "
		SQLQuotedItemsTmp =	SQLQuotedItemsTmp & "			[QuoteType] [varchar](255) NULL, "
		SQLQuotedItemsTmp =	SQLQuotedItemsTmp & "			[SuggestedQty] [varchar](255) NULL, "
		SQLQuotedItemsTmp =	SQLQuotedItemsTmp & "			[Price] [money] NULL, "
		SQLQuotedItemsTmp =	SQLQuotedItemsTmp & "			[ListPrice] [money] NULL, "
		SQLQuotedItemsTmp =	SQLQuotedItemsTmp & "			[ListFlag] [varchar](255) NULL, "
		SQLQuotedItemsTmp =	SQLQuotedItemsTmp & "			[Cost] [money] NULL, "
		SQLQuotedItemsTmp =	SQLQuotedItemsTmp & "			[DateQuoted] [datetime] NULL, "
		SQLQuotedItemsTmp =	SQLQuotedItemsTmp & "			[ExpireDate] [datetime] NULL, "
		SQLQuotedItemsTmp =	SQLQuotedItemsTmp & "			[DeleteFlag] [int] NULL, "
		SQLQuotedItemsTmp =	SQLQuotedItemsTmp & "			[NewPrice] [money] NULL, "
		SQLQuotedItemsTmp =	SQLQuotedItemsTmp & "			[NewGPPercent] [float] NULL, "
		SQLQuotedItemsTmp =	SQLQuotedItemsTmp & "			[NewExpireDate] [datetime] NULL, "
		SQLQuotedItemsTmp =	SQLQuotedItemsTmp & "			[QuotedToChainOrAccount] [varchar](255) NULL, "
		SQLQuotedItemsTmp =	SQLQuotedItemsTmp & "			[AutoGenerated] [bit] NULL "
		SQLQuotedItemsTmp =	SQLQuotedItemsTmp & "		) ON [PRIMARY] "
		Set rsQuotedItemsTmp = cnnQuotedItemsTmp.Execute(SQLQuotedItemsTmp)

		
		' Now being inserting stuff from XML data

		Set nodes = xmlDoc.selectNodes("//*")

		'Set IDnum = xmlDoc.SelectSingleNode("DATASTREAM/IDENTITY")
		'Response.Write(IDnum.text & "<br><br><br>")

		Set AccountData = xmlDoc.selectNodes("//DATASTREAM/ACCOUNT_QUOTE/*" )           
		AccountDataCount = AccountData.length
		
		
		For Each entry in AccountData 
		
			If entry.tagName = "SKU" Then
				'Response.Write("SKU:" & entry.text & ":XX<br>")   
				SKU = entry.text
			ElseIf entry.tagName = "DESCRIPTION" Then
				'Response.Write("DESCRIPTION:" & entry.text & ":XX<br>")   
				DESCRIPTION = entry.text
				DESCRIPTION = Replace(DESCRIPTION,"'","''")
			ElseIf entry.tagName = "CATEGORY" Then
				'Response.Write("CATEGORY:" & entry.text & ":XX<br>")   
				CATEGORY = entry.text
			ElseIf entry.tagName = "UOM" Then
				'Response.Write("UOM:" & entry.text & "<br>")   
				UOM = entry.text
			ElseIf entry.tagName = "SUGGESTED_QTY" Then
				'Response.Write("SUGGESTED_QTY:" & entry.text & "<br>")   
				SUGGESTED_QTY = entry.text
			ElseIf entry.tagName = "PRICE" Then
				'Response.Write("PRICE:" & entry.text & "<br>") 
				PRICE = entry.text  
			ElseIf entry.tagName = "LIST_PRICE" Then
				'Response.Write("LIST_PRICE:" & entry.text & "<br>") 
				LIST_PRICE = entry.text  
			ElseIf entry.tagName = "LIST_FLAG" Then
				'Response.Write("LIST_FLAG:" & entry.text & "<br>")   
				LIST_FLAG = entry.text
			ElseIf entry.tagName = "COST" Then
				'Response.Write("COST:" & entry.text & "<br>")   
				COST = entry.text
			ElseIf entry.tagName = "DATE_QUOTED" Then
				'Response.Write("DATE_QUOTED:" & entry.text & "<br>")  
				DATE_QUOTED = entry.text 
			ElseIf entry.tagName = "DATE_EXPIRES" Then
				'Response.Write("DATE_EXPIRES:" & entry.text & "<br>")  
				DATE_EXPIRES = entry.text 
			ElseIf entry.tagName = "CHAIN_OR_ACCOUNT" Then
				'Response.Write("CHAIN_OR_ACCOUNT:" & entry.text & "<br>") 
				CHAIN_OR_ACCOUNT = entry.text  
				
				SQLQuotedItemsTmp = "INSERT INTO zPRC_AccountQuotedItems_" & trim(Session("Userno")) & " ("
				SQLQuotedItemsTmp = SQLQuotedItemsTmp & " prodSKU, [Description], Category, QuoteType, SuggestedQty, Price, ListPrice, ListFlag, "
				SQLQuotedItemsTmp = SQLQuotedItemsTmp & " Cost, DateQuoted, "
				If DATE_EXPIRES <> "" Then
					SQLQuotedItemsTmp = SQLQuotedItemsTmp & "[ExpireDate], "
				End If
				SQLQuotedItemsTmp = SQLQuotedItemsTmp & "DeleteFlag ,QuotedToChainOrAccount, AutoGenerated)"
				SQLQuotedItemsTmp = SQLQuotedItemsTmp & " VALUES ("
				SQLQuotedItemsTmp = SQLQuotedItemsTmp & "'" & SKU & "', "
				SQLQuotedItemsTmp = SQLQuotedItemsTmp & "'" & DESCRIPTION & "', "
				SQLQuotedItemsTmp = SQLQuotedItemsTmp & "'" & CATEGORY & "', "
				SQLQuotedItemsTmp = SQLQuotedItemsTmp & "'" & UOM & "', "
				If SUGGESTED_QTY = "" Then SUGGESTED_QTY = "0"
				SQLQuotedItemsTmp = SQLQuotedItemsTmp & "'" & SUGGESTED_QTY & "', "			
				SQLQuotedItemsTmp = SQLQuotedItemsTmp & PRICE & ", "						
				SQLQuotedItemsTmp = SQLQuotedItemsTmp & LIST_PRICE & ", "						
				SQLQuotedItemsTmp = SQLQuotedItemsTmp & "'" & LIST_FLAG & "', "			
				SQLQuotedItemsTmp = SQLQuotedItemsTmp & COST & ", "						
				SQLQuotedItemsTmp = SQLQuotedItemsTmp & "'" & DATE_QUOTED & "', "	
				If DATE_EXPIRES <> "" Then
					SQLQuotedItemsTmp = SQLQuotedItemsTmp & "'" & DATE_EXPIRES & "', "	
				End If
				SQLQuotedItemsTmp = SQLQuotedItemsTmp & "0, "	
				SQLQuotedItemsTmp = SQLQuotedItemsTmp & "'" & CHAIN_OR_ACCOUNT & "', "
				SQLQuotedItemsTmp = SQLQuotedItemsTmp & "0 )"	
				
				Response.Write("<br><br><br>"&SQLQuotedItemsTmp&"<br><br><br>")
				
				Set rsQuotedItemsTmp = cnnQuotedItemsTmp.Execute(SQLQuotedItemsTmp)
				
				SKU = "" : DESCRIPTION = "" : CATEGORY = "" : UOM = "" : SUGGESTED_QTY = "" : PRICE = "" : LIST_PRICE ="" : LIST_FLAG = "" : COST = "" : DATE_QUOTED = "" : DATE_EXPIRES = "" : CHAIN_OR_ACCOUNT = ""
				
			End If
				
		Next
'response.end
		Description ="success! httpRequest.status returned " & httpRequest.status & " when posting <RECORD_TYPE>READQUOTES and <RECORD_SUBTYPE>ACCOUNT"& "<br>"
		Description = "httpRequest.responseText:" & httpRequest.responseText & "<br>"
		Description = Description & "PAGE: " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "<br>"
		Description = Description & "Posted to " & GetPOSTParams("CUSTOMERURL1") & "<br>"
		Description = Description & "POSTED DATA:" & data & "<br>"
		Description = Description & "SERNO:" & MUV_READ("ClientID") & "<br>"
		
		CreateINSIGHTAuditLogEntry Request.ServerVariables("SERVER_NAME"),Description,GetPOSTParams("Mode")
		
		'*****************************************
		'Write the info returned from the Post into
		'the temp files before going to next page
		'******************************************
		tmpvar = httpRequest.responseText
		tmpvar=ucase(tmpvar)
		tmpvar = Replace(tmpvar,"SUCCESS","")
		tmpvar = trim(tmpvar)

		Response.Redirect("quotedItemsTool.asp?custID=" & custID)

	Else
		'FAILURE
		emailbody="httpRequest.status returned " & httpRequest.status & " when posting <RECORD_TYPE>READQUOTES and <RECORD_SUBTYPE>ACCOUNT"& "<br>"
		emailBody = "httpRequest.responseText:" & httpRequest.responseText & "<br>"
		emailBody = emailBody & "PAGE: " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "<br>"
		emailBody = emailBody & "Posted to " & GetPOSTParams("CUSTOMERURL1") & "<br>"
		emailBody = emailBody & "POSTED DATA:" & data & "<br>"
		emailBody = emailBody & "SERNO:" & MUV_READ("ClientID") & "<br>"
		SendMail "mailsender@" & maildomain ,"insight@ocsaccess.com",MUV_READ("ClientID") & " READQUOTES POST ERROR",emailBody, "Price Tool", "Post Failure"
	
		Description = emailBody 
		CreateINSIGHTAuditLogEntry Request.ServerVariables("SERVER_NAME"),Description,GetPOSTParams("Mode")
		
		'Go to main with a Querystring, indicationg a read error
		'Doesn't matter what the value is. Any querystring being present
		'indicates a read error
		Response.Redirect("reports.asp?s=0")
	End If
End If

%><!--#include file="../../../../inc/footer-main.asp"-->