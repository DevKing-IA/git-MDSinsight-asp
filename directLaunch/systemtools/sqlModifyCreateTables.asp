<!--#include file="../../inc/SubsAndFuncs.asp"-->
<!--#include file="../../inc/InsightFuncs.asp"-->
<script type="text/javascript">
	function closeme() {
		window.open('', '_parent', '');
		window.close();  }
</script>
<%
Server.ScriptTimeout = 7000
'SQL Table Creation and Modification Script
'Designed to be launched via a scheduled process (Win Task Scheduler)
'Self contained page will check the alerts db and take the appropriate actions
'Usage = "http://{xxx}.{domain}.com/directLaunch/systemtools/sqlModifyCreateTables.asp?runlevel=run_now

'The runlevel parameter is inconsequential to the operation 
'of the page. It is only used so that the page will not run
'if it is loaded via an unexpected method (spiders, etc)
If Request.QueryString("runlevel") <> "run_now" then response.end

'baseURL should alwats have a trailing /slash, just in case, handle either way
If right(baseURL,1)="/" Then maildomain = Left(right(baseURL,len(baseURL)-7),len(right(baseURL,len(baseURL)-7))-1) Else maildomain = right(baseURL,len(baseURL)-7)

'This single page loops through and handles autocompletes for all databases
SQL = "SELECT * FROM tblServerInfo WHERE Active = 1"

Set TopConnection = Server.CreateObject("ADODB.Connection")
Set TopRecordset = Server.CreateObject("ADODB.Recordset")
TopConnection.Open InsightCnnString
'Open the recordset object executing the SQL statement and return records
TopRecordset.Open SQL,TopConnection,3,3


'First lookup the ClientKey in tblServerInfo
'If there is no record with the entered client key, close connection
'and exit
If Not TopRecordset.Eof Then

	Do While Not TopRecordset.EOF

		PassPhrase = TopRecordset.Fields("directLaunchPassphrase")
		ClientKey = TopRecordset.Fields("clientkey")
		


		Response.Write("******** START Processing SQL Changes For " & ClientKey  & "************<br>")
		
		Call SetClientCnnString
		
		Session("ClientCnnString") = MUV_READ("ClientCnnString") ' Until session vars are gone, then delete this
		
		If MUV_READ("cnnStatus") = "OK" Then ' else it loops and excludes dev client keys
			
			'****************************************
			'Begin Modify, Create SQL Tables
			'****************************************
			 Response.Write("Begin Modify, Create SQL Table Structures<br>")
			 'Response.write(Server.MapPath(".") & "<br><br>")
			'******************************************

			Server.ScriptTimeout = 50000
			
			%><!--#include file="sql/sqlModifyCreateTables-API_IN_Shopify.asp"--><%
			%><!--#include file="sql/sqlModifyCreateTables-FS_ServiceMemos.asp"--><%
			%><!--#include file="sql/sqlModifyCreateTables-RT_DeliveryBoard.asp"--><%
			%><!--#include file="sql/sqlModifyCreateTables-SC_AuditLogDLaunch.asp"--><%
			%><!--#include file="sql/sqlModifyCreateTables-Settings_CompanyID.asp"--><%
			%><!--#include file="sql/sqlModifyCreateTables-Settings_Global.asp"--><%
			%><!--#include file="sql/sqlModifyCreateTables-tblUsers.asp"--><%
			%><!--#include file="sql/sqlModifyCreateTables-SC_EmailCustomization.asp"--><%
			%><!--#include file="sql/sqlDropTables.asp"--><%
			%><!--#include file="sql/sqlModifyCreateTables-Settings_BizIntel.asp"--><%
			%><!--#include file="sql/sqlModifyCreateTables-AR_Customer.asp"--><%
			%><!--#include file="sql/sqlModifyCreateTables-AR_CustomerBillInfo.asp"--><%
			%><!--#include file="sql/sqlModifyCreateTables-SC_AlertsSent.asp"--><%
			%><!--#include file="sql/sqlModifyCreateTables-API_RA_PostResults.asp"--><%
			%><!--#include file="sql/sqlModifyCreateTables-API_AuditLog.asp"--><%
			%><!--#include file="sql/sqlModifyCreateTables-API_OR_RAHeader.asp"--><%	
			%><!--#include file="sql/sqlModifyCreateTables-Equipment.asp"--><%
			%><!--#include file="sql/sqlModifyCreateTables-IC_Partners.asp"--><%
			%><!--#include file="sql/sqlModifyCreateTables-IC_ProductImages.asp"--><%	
			%><!--#include file="sql/sqlModifyCreateTables-SC_NeedToKnow.asp"--><%	
			%><!--#include file="sql/sqlModifyCreateTables-API_IC_AdjustOnHand.asp"--><%			
			%><!--#include file="sql/sqlModifyCreateTables-API_IC_PostResults.asp"--><%	
			%><!--#include file="sql/sqlModifyCreateTables-BI_CompanyLeakageByPeriod.asp"--><%	
			%><!--#include file="sql/sqlModifyCreateTables-BI_CatAnalByPeriodNotesUserViewed.asp"--><%	
			%><!--#include file="sql/sqlModifyCreateTables-AR_CustomerNotes.asp"--><%	
			%><!--#include file="sql/sqlModifyCreateTables-BI_CustCatPeriodSales2.asp"--><%
			%><!--#include file="sql/sqlModifyCreateTables-BI_MCSActions.asp"--><%
			%><!--#include file="sql/sqlUpdateFields-SC_EmailCustomization_updateEmailFileName.asp"--><%				
			%><!--#include file="sql/sqlModifyCreateTables-RT_Routes.asp"--><%
			%><!--#include file="sql/sqlModifyCreateTables-Referal.asp"--><%
			%><!--#include file="sql/sqlModifyCreateTables-CustomerType.asp"--><%
			%><!--#include file="sql/sqlModifyCreateTables-Salesman.asp"--><%
			%><!--#include file="sql/sqlModifyCreateTables-AR_CustomerInactive.asp"--><%
			%><!--#include file="sql/sqlModifyCreateTables-AR_CustomerCounts.asp"--><%
			%><!--#include file="sql/sqlModifyCreateTables-Settings_Quickbooks.asp"--><%
			%><!--#include file="sql/sqlModifyCreateTables-API_OR_ORHeader.asp"--><%
			%><!--#include file="sql/sqlModifyCreateTables-BI_MCSReasons.asp"--><%
			%><!--#include file="sql/sqlModifyCreateTables-IC_Product.asp"--><%
			%><!--#include file="sql/sqlModifyCreateTables-API_IC_ReplaceOnHand.asp"--><%
			%><!--#include file="sql/sqlModifyCreateTables-BI_MCSData.asp"--><%
			%><!--#include file="sql/sqlModifyCreateTables-Settings_NeedToKnow.asp"--><%
			%><!--#include file="sql/sqlModifyCreateTables-BI_DailySalesByTypeByClass.asp"--><%
			%><!--#include file="sql/sqlModifyCreateTables-BI_MESData.asp"--><%
			%><!--#include file="sql/sqlModifyCreateTables-BI_MESActions.asp"--><%
			%><!--#include file="sql/sqlModifyCreateTables-BI_CompanyLeakageByMonth.asp"--><%
			%><!--#include file="sql/sqlModifyCreateTables-WEB_CustProdInclude.asp"--><%
			%><!--#include file="sql/sqlModifyCreateTables-WEB_Categories.asp"--><%
			%><!--#include file="sql/sqlModifyCreateTables-WEB_CategoriesXREF.asp"--><%
			%><!--#include file="sql/sqlModifyCreateTables-WEB_Rotators.asp"--><%
			%><!--#include file="sql/sqlModifyCreateTables-WEB_HomepageProducts.asp"--><%
			%><!--#include file="sql/sqlModifyCreateTables-WEB_ProductImages.asp"--><%
			%><!--#include file="sql/sqlModifyCreateTables-WEB_ProdUnits.asp"--><%	
			%><!--#include file="sql/sqlModifyCreateTables-WEB_Users.asp"--><%
			%><!--#include file="sql/sqlModifyCreateTables-WEB_Tracking.asp"--><%			
			%><!--#include file="sql/sqlModifyCreateTables-WEB_ShoppingLists.asp"--><%		
			%><!--#include file="sql/sqlModifyCreateTables-WEB_TempUserInfo.asp"--><%
			%><!--#include file="sql/sqlModifyCreateTables-WEB_TempOrder.asp"--><%
			%><!--#include file="sql/sqlModifyCreateTables-WEB_Order.asp"--><%
			%><!--#include file="sql/sqlModifyCreateTables-WEB_OrderDetails.asp"--><%			
			%><!--#include file="sql/sqlModifyCreateTables-API_IN_InvoiceHeader.asp"--><%
			%><!--#include file="sql/sqlModifyCreateTables-AR_CustomerQuotedItems.asp"--><%
			%><!--#include file="sql/sqlModifyCreateTables-AR_CustomerCategoryDiscount.asp"--><%			
			%><!--#include file="sql/sqlModifyCreateTables-AR_CustomerBillTo.asp"--><%
			%><!--#include file="sql/sqlModifyCreateTables-AR_CustomerShipTo.asp"--><%
			%><!--#include file="sql/sqlModifyCreateTables-Settings_FieldService.asp"--><%		
			%><!--#include file="sql/sqlModifyCreateTables-FS_ServiceMemosNotes.asp"--><%			
			%><!--#include file="sql/sqlModifyCreateTables-FS_ServiceMemosNotesUserViewed.asp"--><%	
			%><!--#include file="sql/sqlModifyCreateTables-RT_DeliveryBoardHistory.asp"--><%
			%><!--#include file="sql/sqlModifyCreateTables-Settings_InventoryControl.asp"--><%
			%><!--#include file="sql/sqlModifyCreateTables-FS_ServiceMemosRedispatch.asp"--><%
			%><!--#include file="sql/sqlModifyCreateTables-FS_ProblemCodes.asp"--><%
			%><!--#include file="sql/sqlModifyCreateTables-PR_Prospects.asp"--><%
			%><!--#include file="sql/sqlModifyCreateTables-EQ_Models.asp"--><%
			%><!--#include file="sql/sqlModifyCreateTables-AR_CustomerNotesUserViewed.asp"--><%
			%><!--#include file="sql/sqlModifyCreateTables-FS_CustomerFilters.asp"--><%
			%><!--#include file="sql/sqlModifyCreateTables-PR_Competitors.asp"--><%
			%><!--#include file="sql/sqlModifyCreateTables-Settings_Prospecting.asp"--><%
			%><!--#include file="sql/sqlModifyCreateTables-PR_ProspectContactSearch.asp"--><%
			%><!--#include file="sql/sqlModifyCreateTables-SC_Alerts.asp"--><%	
			%><!--#include file="sql/sqlModifyCreateTables-Settings_CompanyCalendar.asp"--><%									
			%><!--#include file="sql/sqlModifyCreateTables-API_OR_ORDetail.asp"--><%
			%><!--#include file="sql/sqlModifyCreateTables-PR_ProspectSocialMedia.asp"--><%
			%><!--#include file="sql/sqlModifyCreateTables-PR_ProspectContacts.asp"--><%
			%><!--#include file="sql/sqlModifyCreateTables-IC_Filters.asp"--><%	
			%><!--#include file="sql/sqlModifyCreateTables-FS_Parts.asp"--><%				
			%><!--#include file="sql/sqlModifyCreateTables-BI_PostedUnpostedByCustCatMonth.asp"--><%
			%><!--#include file="sql/sqlModifyCreateTables-BI_PostedUnpostedByCustCatPeriod.asp"--><%	
			%><!--#include file="sql/sqlModifyCreateTables-AR_Regions.asp"--><%		
			%><!--#include file="sql/sqlModifyCreateTables-PR_Countries.asp"--><%	
			%><!--#include file="sql/sqlModifyCreateTables-PR_States.asp"--><%				
			%><!--#include file="sql/sqlModifyCreateTables-FS_ServiceMemosFilterInfo.asp"--><%
			%><!--#include file="sql/sqlModifyCreateTables-IN_WebFulfillment.asp"--><%
			%><!--#include file="sql/sqlModifyCreateTables-Settings_AR.asp"--><%
			%><!--#include file="sql/sqlModifyCreateTables-PR_ProspectEmailLog.asp"--><%
			%><!--#include file="sql/sqlModifyCreateTables-AR_CustomerMapping.asp"--><%
			%><!--#include file="sql/sqlModifyCreateTables-IN_InvoiceHistDetail.asp"--><%
			%><!--#include file="sql/sqlModifyCreateTables-IN_InvoiceHistHeader.asp"--><%		
			%><!--#include file="sql/sqlModifyCreateTables-AR_Terms.asp"--><%		
			%><!--#include file="sql/sqlModifyCreateTables-AR_CustomerType.asp"--><%
			%><!--#include file="sql/sqlModifyCreateTables-AR_CustomerReferral.asp"--><%
			%><!--#include file="sql/sqlModifyCreateTables-AR_Chain.asp"--><%	
			%><!--#include file="sql/sqlModifyCreateTables-IN_InvoiceHistExportedSage.asp"--><%
			%><!--#include file="sql/sqlModifyCreateTables-Settings_API.asp"--><%
			%><!--#include file="sql/sqlModifyCreateTables-SC_SchedulerLog.asp"--><%
			%><!--#include file="sql/sqlModifyCreateTables-AR_CustomerContacts.asp"--><%			
			%><!--#include file="sql/sqlModifyCreateTables-AR_PaymentMethods.asp"--><%	
			%><!--#include file="sql/sqlModifyCreateTables-AP_Vendor.asp"--><%	
			%><!--#include file="sql/sqlModifyCreateTables-USER_Teams.asp"--><%			
			%><!--#include file="sql/sqlModifyCreateTables-PR_ProspectEmailMessages.asp"--><%			
			%><!--#include file="sql/sqlModifyCreateTables-PR_ProspectEmailMessagesBody.asp"--><%			
			%><!--#include file="sql/sqlModifyCreateTables-FS_ServiceMemosAnlCustMonth.asp"--><%	
			%><!--#include file="sql/sqlModifyCreateTables-SC_NoteType.asp"--><%
			%><!--#include file="sql/sqlModifyCreateTables-Settings_AccountingPeriods.asp"--><%			
			%><!--#include file="sql/sqlModifyCreateTables-FS_SymptomCodes.asp"--><%
			%><!--#include file="sql/sqlModifyCreateTables-FS_ResolutionCodes.asp"--><%
			%><!--#include file="sql/sqlModifyCreateTables-FS_ServiceMemosDetail.asp"--><%	
			%><!--#include file="sql/sqlModifyCreateTables-AR_CustomerPOS.asp"--><%
			%><!--#include file="sql/sqlModifyCreateTables-FS_WEB_Users.asp"--><%
			%><!--#include file="sql/sqlModifyCreateTables-PR_ProspectEmailMessagesAttachments.asp"--><%
		
			'sqlModifyCreateTables-CustCatPeriodSales.asp
			
			Response.Write("End Modify, Create SQL Table Structures<br>")
			'******************************************	
			
		End If
		
		Response.Write("******** DONE Processing SQL Changes For " & ClientKey & "************<br>")
			
					
	TopRecordset.movenext
	
	Loop

	TopRecordset.Close
	Set TopRecordset = Nothing
	TopConnection.Close
	Set TopConnection = Nothing
	
End If

Response.write("<script type=""text/javascript"">closeme();</script>")	
'Response.End
'*************************
'*************************
'Subs and funcs begin here


Sub SetClientCnnString

	dummy=MUV_WRITE("cnnStatus","")

	SQL = "SELECT * FROM tblServerInfo where clientKey='"& ClientKey &"'"

	Set Connection = Server.CreateObject("ADODB.Connection")
	Set Recordset = Server.CreateObject("ADODB.Recordset")
	Connection.Open InsightCnnString
	
	'Open the recordset object executing the SQL statement and return records
	Recordset.Open SQL,Connection,3,3

	
	'First lookup the ClientKey in tblServerInfo
	'If there is no record with the entered client key, close connection
	'and exit
	If Recordset.recordcount <= 0 then
		Recordset.close
		Connection.close
		set Recordset=nothing
		set Connection=nothing
	Else
		ClientCnnString = "Driver={SQL Server};Server=" & Recordset.Fields("dbServer")
		ClientCnnString = ClientCnnString & ";Database=" & Recordset.Fields("dbCatalog")
		ClientCnnString = ClientCnnString & ";Uid=" & Recordset.Fields("dbLogin")
		ClientCnnString = ClientCnnString & ";Pwd=" & Recordset.Fields("dbPassword") & ";"
		dummy = MUV_Write("ClientCnnString",ClientCnnString)
		dummy = MUV_Write("SQL_Owner",Recordset.Fields("dbLogin"))
		dummy = MUV_Write("ClientID",Recordset.Fields("clientkey"))
		Recordset.close
		Connection.close
		dummy=MUV_WRITE("cnnStatus","OK")
	End If
End Sub



%>