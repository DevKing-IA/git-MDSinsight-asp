<!--#include file="../../inc/SubsAndFuncs.asp"-->
<!--#include file="../../inc/InSightFuncs.asp"-->

<%

InternalRecordIdentifier = Request.Form("txtInternalRecordIdentifier")
LastPriceChangeDate = Request.Form("txtLastPriceChangeDate")
Orig_LastPriceChangeDate = Request.Form("txtLastPriceChangeDateOriginal")

'*******************************************************************************************************************
'FIX ANY ENTRIES THAT MAY CONTAIN SINGLE QUOTES FOR SQL INSERT
'*******************************************************************************************************************

LastPriceChangeDate = Replace(LastPriceChangeDate,"'","''")

'***************************************************************************************************************************************************************
'First update main customer table, AR_Customer
'***************************************************************************************************************************************************************

Set cnnAR_Customer = Server.CreateObject("ADODB.Connection")
cnnAR_Customer.open (Session("ClientCnnString"))
Set rsAR_Customer = Server.CreateObject("ADODB.Recordset")

SQLAR_Customer = "UPDATE AR_Customer SET LastPriceChangeDate = '" & LastPriceChangeDate & "' WHERE InternalRecordIdentifier = " & InternalRecordIdentifier

rsAR_Customer.CursorLocation = 3 
Set rsAR_Customer = cnnAR_Customer.Execute(SQLAR_Customer)

Response.Write("SQLAR_Customer: " & SQLAR_Customer & "<br><br>")

'***************************************************************************************************************************************************************
'Last Price Change Date, Just Add Audit Trail Entry
'***************************************************************************************************************************************************************

If ORIG_LastPriceChangeDate <> LastPriceChangeDate Then

	If IsNull(ORIG_LastPriceChangeDate) OR ORIG_LastPriceChangeDate ="1/1/1900" OR ORIG_LastPriceChangeDate = "" Then
		Description = "The last price change date for customer " & CompanyName & " (" & AccountNumber & ") was changed to <strong><em>" & formatDateTime(LastPriceChangeDate,2) & "</em></strong> from <strong><em>NO DATE ENTERED</em></strong> by " & GetUserDisplayNameByUserNo(Session("UserNo")) 
		CreateAuditLogEntry GetTerm("Accounts Receivable") & " customer last price change date changed",GetTerm("Accounts Receivable") & " customer last price change date changed","Major",0,Description
	Else
		If DateDiff("d",cDate(ORIG_LastPriceChangeDate),cDate(LastPriceChangeDate)) <> 0 Then
			Description = "The last price change date for customer " & CompanyName & " (" & AccountNumber & ") was changed to <strong><em>" & formatDateTime(LastPriceChangeDate,2) & "</em></strong> from <strong><em>" & formatDateTime(ORIG_LastPriceChangeDate,2) & "</em></strong> by " & GetUserDisplayNameByUserNo(Session("UserNo")) 
			CreateAuditLogEntry GetTerm("Accounts Receivable") & " customer last price change date changed",GetTerm("Accounts Receivable") & " customer last price change date changed","Major",0,Description	
		End If
	End If
	
End If



Response.Redirect("editViewCustomerDetail.asp?cid=" & AccountNumber)

%>
