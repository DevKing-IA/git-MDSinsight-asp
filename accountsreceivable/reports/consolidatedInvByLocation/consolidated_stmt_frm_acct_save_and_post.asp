<!--#include file="../../../inc/subsandfuncs.asp"-->
<!--#include file="../../../inc/Insightfuncs.asp"-->
<!--#include file="../../../inc/Insightfuncs_API.asp"-->
<!--#include file="../../../inc/mail.asp"-->
<%
	
	'baseURL should always have a trailing /slash, just in case, handle either way
	If right(baseURL,1)="/" Then maildomain = Left(right(baseURL,len(baseURL)-7),len(right(baseURL,len(baseURL)-7))-1) Else maildomain = right(baseURL,len(baseURL)-7)
	sURL = Request.ServerVariables("SERVER_NAME")

	Account = Request.Form("c")
	EndDate = Request.Form("e")
	EndDate = Replace(EndDate, "~","/")
	StartDate = Request.Form("s")
	StartDate = Replace(StartDate, "~","/")
	DueDateDays = Request.Form("ddd")
	DueDateSingleDate = Request.Form("dds")
	
	If Request.Form("z") = "T" then
		SkipZeroDollar = True
	Else
		SkipZeroDollar = False
	End If
	If Request.Form("lz") = "T" then
		SkipLessThenZero = True
	Else
		SkipLessThanZero = False
	End If
	If Request.Form("lzl") = "T" then
		SkipLessThanZeroLines = True
	Else
		SkipLessThanZeroLines = False
	End If
	
	IncludedType = Request.Form("ty")
	
	UserNo = Session("UserNo")
	Username = GetUserDisplayNameByUserNo(UserNo)
	ClientKey = MUV_Read("ClientID")
	SERNO = GetPOSTParams("SERNO")	
	
	'**************************************************************************************************
	'CODE TO SAVE CONSOLIDATED INVOICE TO CLIENTFILES\CLIENTID\CUSTOMER\ACCOUNTSRECEIVABLE
	'**************************************************************************************************
	
	'Now change the name of the file
	Orig_Name =  "\clientfiles\" & Left(MUV_Read("ClientID"),4) & "\z_pdfs\ConsolidatedStatement_Account_" & Trim(Account) & "_" & Trim(Account) & Trim(Replace(EndDate,"/","")) & ".pdf"
	New_Name =  "\clientfiles\" & Left(MUV_Read("ClientID"),4) & "\customer\accountsreceivable\ConsolidatedStatement_Account_" & Trim(Account) & "_" & Trim(Account) & Trim(Replace(EndDate,"/","")) & ".pdf"
	
	'Response.Write("Orig_Name " & Orig_Name & "<br>")
	'Response.Write("New_Name " & New_Name & "<br>")
	
	Dim fso
	
	Set fso = CreateObject("Scripting.FileSystemObject")
	
	'Kill it first in case an old one is there
	On error resume next
	fso.DeleteFile Server.MapPath(New_Name)
	On error goto 0
	
	fso.CopyFile Server.MapPath(Orig_Name), Server.MapPath(New_Name)
	
	Set fso = Nothing
	
	'**************************************************************************************************
	'END CODE TO SAVE CONSOLIDATED INVOICE TO CLIENTFILES\CLIENTID\CUSTOMER\ACCOUNTSRECEIVABLE
	'**************************************************************************************************


	'**************************************************************************************************
	'CODE TO CREATE AND POST XML TO METROPLEX
	'**************************************************************************************************

		data = ""					
		
		SQLInvoices = "SELECT * FROM " & MUV_Read("SQL_Owner") & ".InvoiceHistory WHERE CustNum = '" & Account & "' AND "
		SQLInvoices = SQLInvoices & "IvsDate >= '" & StartDate & "' AND IvsDate <= '" & EndDate & "' "
		
		If SkipZeroDollar = True Then SQLInvoices = SQLInvoices & "AND IvsTotalAmt <> 0 "
		If SkipLessThanZero = True Then SQLInvoices = SQLInvoices & "AND IvsTotalAmt > 0 "
		If IncludedType <> "" Then SQLInvoices = SQLInvoices & "AND CHARINDEX(IvsType,'" & IncludedType & "') <> 0 "

		SQLInvoices = SQLInvoices & " ORDER BY IvsNum"

		Set cnnInvoices = Server.CreateObject("ADODB.Connection")
		cnnInvoices.open (Session("ClientCnnString"))
		Set rsInvoices = Server.CreateObject("ADODB.Recordset")
		rsInvoices.CursorLocation = 3 
		
		Set rsInvoices = cnnInvoices.Execute(SQLInvoices)

		If not rsInvoices.Eof Then
		
			TotalAmt = 0
			SubtotalAmt = 0
			TotalSalesTax = 0
			InvoiceCount = 0
			
			Do While not rsInvoices.Eof
				TotalAmt = TotalAmt + rsInvoices("IvsTotalAmt")
				TotalSalesTax = TotalSalesTax + rsInvoices("IvsSalesTax")
				SubtotalAmt = SubtotalAmt + (rsInvoices("IvsTotalAmt") - rsInvoices("IvsSalesTax") - rsInvoices("IvsDepositChg"))
				InvoiceCount = InvoiceCount + 1
				rsInvoices.movenext
			Loop
			
		End If
		
		Set rsInvoices = Nothing
		cnnInvoices.Close
		Set cnnInvoices = Nothing
		
			
		xmlData = "<DATASTREAM>"
		xmlData = xmlData & "<IDENTITY>Pm8316wyc011</IDENTITY>"
		
		xmlData = xmlData & "<MODE>" & GetPOSTParams("REPOSTSUMINVMODE") & "</MODE>"
		
		xmlData = xmlData & "<RECORD_TYPE>SUMMARY_INVOICE</RECORD_TYPE>"
	
		xmlData = xmlData & "<RECORD_SUBTYPE>UPSERT</RECORD_SUBTYPE>"

		xmlData = xmlData & "<SERNO>" & GetPOSTParams("SERNO") & "</SERNO>"
		
		
		xmlData = xmlData & "<INVOICE>"
		
		xmlData = xmlData & "<INVOICE_HEADER>"
		xmlData = xmlData & "<SUM_INVOICE_ID>" & Trim(Account) & Trim(Replace(EndDate,"/","")) & "</SUM_INVOICE_ID>"
		xmlData = xmlData & "<SUM_INVOICE_DATE>" & FormatDateTime(Now(),2) & "</SUM_INVOICE_DATE>" ' Date the pdf is generated
		
		If DueDateSingleDate <> "" Then 
		 	InvoiceDueDate = FormatDateTime(DueDateSingleDate,2)
		Else
			InvoiceDueDate = FormatDateTime(DateAdd("d",DueDateDays,EndDate),2)
		End If

		xmlData = xmlData & "<SUM_INVOICE_AGE_DATE>" & InvoiceDueDate & "</SUM_INVOICE_AGE_DATE>" 'Due date, however it is arrived at
		
		xmlData = xmlData & "<CUST_ID>" & Trim(Account) & "</CUST_ID>"
		xmlData = xmlData & "<NUM_INVOICES>" & InvoiceCount & "</NUM_INVOICES>"
		xmlData = xmlData & "<SUB_TOTAL>" & FormatCurrency(Round(SubtotalAmt,2),2) & "</SUB_TOTAL>"
		xmlData = xmlData & "<TOTAL_TAX>" & FormatCurrency(Round(TotalSalesTax,2),2) & "</TOTAL_TAX>"
		xmlData = xmlData & "<GRAND_TOTAL>" & FormatCurrency(Round(TotalAmt,2),2) & "</GRAND_TOTAL>"
		xmlData = xmlData & "</INVOICE_HEADER>"
		
		
		
		xmlData = xmlData & "<INVOICE_DETAILS>"


		SQLInvoicesSingle = "SELECT * FROM " & MUV_Read("SQL_Owner") & ".InvoiceHistory WHERE CustNum = '" & Account & "' AND "
		SQLInvoicesSingle = SQLInvoicesSingle & "IvsDate >= '" & StartDate & "' AND IvsDate <= '" & EndDate & "' "
		
		If SkipZeroDollar = True Then SQLInvoicesSingle = SQLInvoicesSingle & "AND IvsTotalAmt <> 0 "
		If SkipLessThanZero = True Then SQLInvoicesSingle = SQLInvoicesSingle & "AND IvsTotalAmt > 0 "
		If IncludedType <> "" Then SQLInvoicesSingle = SQLInvoicesSingle & "AND CHARINDEX(IvsType,'" & IncludedType & "') <> 0 "

		SQLInvoicesSingle = SQLInvoicesSingle & " ORDER BY IvsNum"

		Set cnnInvoicesSingle = Server.CreateObject("ADODB.Connection")
		cnnInvoicesSingle.open (Session("ClientCnnString"))
		Set rsInvoicesSingle = Server.CreateObject("ADODB.Recordset")
		rsInvoicesSingle.CursorLocation = 3 
		
		Set rsInvoicesSingle = cnnInvoicesSingle.Execute(SQLInvoicesSingle)

		If not rsInvoicesSingle.Eof Then

			DetailNum = 0
			
			Do While not rsInvoicesSingle.Eof

				DetailNum = DetailNum + 1
				
				xmlData = xmlData & "<INVOICE_DETAIL_LINE>"
				xmlData = xmlData & "<DETAIL_NUM>" & DetailNum & "</DETAIL_NUM>"
				xmlData = xmlData & "<INVOICE_ID>" & rsInvoicesSingle("IvsNum") & "</INVOICE_ID>"
				xmlData = xmlData & "<CUST_ID>" & Trim(Account) & "</CUST_ID>"
				xmlData = xmlData & "<INVOICE_TOTAL>" & FormatCurrency(Round(rsInvoicesSingle("IvsTotalAmt"),2),2) & "</INVOICE_TOTAL>"
				xmlData = xmlData & "</INVOICE_DETAIL_LINE>"
				
				rsInvoicesSingle.movenext
			Loop
		End If
		
		Set rsInvoicesSingle = Nothing
		cnnInvoicesSingle.Close
		Set cnnInvoicesSingle = Nothing
		
		
		xmlData = xmlData & "</INVOICE_DETAILS>"
		xmlData = xmlData & "</INVOICE>"
		xmlData = xmlData & "</DATASTREAM>"
		
		xmlDataForDisp = xmlData 
		xmlDataForDisp = Replace(xmlDataForDisp,"     ","")
		xmlDataForDisp = Replace(xmlDataForDisp,"    <","<")	
		xmlDataForDisp = Replace(xmlDataForDisp,"   <","<")	
		xmlDataForDisp = Replace(xmlDataForDisp,"  <","<")	
		xmlDataForDisp = Replace(xmlDataForDisp," <","<")	
		xmlDataForDisp = Replace(xmlDataForDisp,"<","[")
		xmlDataForDisp = Replace(xmlDataForDisp ,">","]")
		xmlDataForDisp = Replace(xmlDataForDisp ,"][","]<br>[")
		xmlDataForDisp = Replace(xmlDataForDisp ,"[","</b>[")
		xmlDataForDisp = Replace(xmlDataForDisp ,"]","]<b>")
		
		data = xmlData
		

		Description = "Post to " & GetPOSTParams("ServiceMemoURL1")
		
		CreateINSIGHTAuditLogEntry sURL,Description,GetPOSTParams("REPOSTSUMINVMODE") 
		
		Description = "data:" & data 
		CreateINSIGHTAuditLogEntry sURL,Description,GetPOSTParams("REPOSTSUMINVMODE") 
		
		Set httpRequest = Server.CreateObject("MSXML2.ServerXMLHTTP")
		httpRequest.Open "POST", GetAPIRepostSumInvURL(), False
	'	httpRequest.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		httpRequest.SetRequestHeader "Content-Type", "text/xml"
		
		xmlData = Replace(xmlData,"&","&amp;")
		xmlData = Replace(xmlData,chr(34),"")			
		httpRequest.Send xmlData
		
		Response.Write("API Response:" & httpRequest.responseText & "<br><br><br>")

		Identity = "Pm8316wyc011"
		
		If (Err.Number <> 0 ) Then
			emailbody="httpRequest.status returned " & httpRequest.status & " when posting <RECORD_TYPE>SUMMARY_INVOICE and <RECORD_SUBTYPE>UPSERT"& "<br><br>"
			emailBody = "httpRequest.responseText:" & httpRequest.responseText & "<br><br>"
			emailBody = emailBody & "PAGE: " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "<br>"
			emailBody = emailBody & "Posted to " & GetAPIRepostSumInvURL() & "<br><br>"
			emailBody = emailBody & "POSTED DATA:<br>" & xmlDataForDisp & "<br><br>"
			emailBody = emailBody & "SERNO: " & SERNO & "<br>"
			SendMail "mailsender@" & maildomain ,"rsmith@ocsaccess.com",SERNO & " Post Error Consolidated Invoice",emailBody, "Consolidated Invoice", "Consolidated Invoice"
		
			Description = emailBody 
			
			Call Write_API_AuditLog_Entry(Identity ,emailBody,GetPOSTParams("REPOSTSUMINVMODE"),"consolidated_stmt_frm_acct_save_and_post.asp")
			
		End If

		If httpRequest.status = 200 THEN 
		
			If Instr(httpRequest.responseText,"success") <> 0 Then ' Success
		
				emailbody="httpRequest.status returned " & httpRequest.status & " when posting <RECORD_TYPE>SUMMARY_INVOICE and <RECORD_SUBTYPE>UPSERT<"& "<br><br>"
				emailBody = emailBody & "httpRequest.responseText:" & httpRequest.responseText & "<br><br>"
				emailBody = emailBody & "PAGE: " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "<br>"
				emailBody = emailBody & "Posted to " & GetAPIRepostSumInvURL() & "<br><br>"
				emailBody = emailBody & "POSTED DATA:<br>" & xmlDataForDisp & "<br><br>"
				emailBody = emailBody & "SERNO: " & SERNO & "<br>"
				SendMail "mailsender@" & maildomain ,"rsmith@ocsaccess.com", SERNO & " Good RePost Consolidated Invoice",emailBody, "Consolidated Invoice", "Consolidated Invoice"
				
				Call Write_API_AuditLog_Entry(Identity ,emailBody,GetPOSTParams("REPOSTSUMINVMODE"),"consolidated_stmt_frm_acct_save_and_post.asp")
				
			Else
				'FAILURE
				emailbody="httpRequest.status returned " & httpRequest.status & " when posting <RECORD_TYPE>SUMMARY_INVOICE and <RECORD_SUBTYPE>UPSERT<"& "<br><br>"
				emailBody = emailBody & "httpRequest.responseText:" & httpRequest.responseText & "<br><br>"
				emailBody = emailBody & "PAGE: " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "<br>"
				emailBody = emailBody & "Posted to " & GetAPIRepostSumInvURL() & "<br><br>"
				emailBody = emailBody & "POSTED DATA:<br>" & xmlDataForDisp & "<br><br>"
				emailBody = emailBody & "SERNO: " & SERNO & "<br>"
				SendMail "mailsender@" & maildomain ,"rsmith@ocsaccess.com",SERNO & " RePost Error Consolidated Invoice",emailBody, "Consolidated Invoice", "Consolidated Invoice"
			
				Call Write_API_AuditLog_Entry(Identity ,emailBody ,GetPOSTParams("REPOSTSUMINVMODE"),"consolidated_stmt_frm_acct_save_and_post.asp")
				
			End If
			
		Else
		
				'FAILURE
				emailbody="NON 200 Response - httpRequest.status returned " & httpRequest.status & " when posting <RECORD_TYPE>SUMMARY_INVOICE and <RECORD_SUBTYPE>UPSERT<"& "<br><br>"
				emailBody = emailBody & "httpRequest.responseText:" & httpRequest.responseText & "<br><br>"
				emailBody = emailBody & "PAGE: " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "<br>"
				emailBody = emailBody & "Posted to " & GetAPIRepostSumInvURL() & "<br><br>"
				emailBody = emailBody & "POSTED DATA:<br>" & xmlDataForDisp & "<br><br>"
				emailBody = emailBody & "SERNO: " & SERNO & "<br>"
				SendMail "mailsender@" & maildomain ,"rsmith@ocsaccess.com",SERNO & " RePost Error Consolidated Invoice",emailBody, "Consolidated Invoice", "Consolidated Invoice"
			
				Call Write_API_AuditLog_Entry(Identity ,emailBody,GetPOSTParams("REPOSTSUMINVMODE"),"consolidated_stmt_frm_acct_save_and_post.asp")
	
		End If

		Set httpRequest = Nothing
	
			
		'Response.Write("<br>XX" &  data & "XX<br>")
		'Response.Write("<br>XX" &  postResponse & "XX<br>")
		'Response.end
	
	'**************************************************************************************************
	'END CODE TO CREATE AND POST XML TO METROPLEX
	'**************************************************************************************************

%>