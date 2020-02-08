<!--#include file="../../../inc/InsightFuncs.asp"-->
<!--#include file="../../../inc/InSightFuncs_routing.asp"-->
<%
IvsNum = Request.Form("txtIvsNum")
CustNum = Request.Form("txtCustNum")
DriverComments = Request.Form("txtdriverComments")
DriverComments = Replace(DriverComments,"'","''")
DriverComments = Server.HTMLEncode(DriverComments)

Response.Write("<br>" & DriverComments & "<BR><BR>")

If  IvsNum <> "" Then
	
	SQLDeliveryBoard = "Update RT_DeliveryBoard Set DeliveryStatus = 'No Delivery' , LastDeliveryStatusChange = getdate(),ManualNextStop = 0, ManualNextStopChanged = NULL, DeliveryInProgress = 0 "
	If DriverComments <> "" Then SQLDeliveryBoard = SQLDeliveryBoard & ", DriverComments = '" & DriverComments & "' "
	SQLDeliveryBoard = SQLDeliveryBoard & " WHERE IvsNum = " & IvsNum

	
	Set cnnDeliveryBoard = Server.CreateObject("ADODB.Connection")
	cnnDeliveryBoard.open (Session("ClientCnnString"))
	Set rsDeliveryBoard = Server.CreateObject("ADODB.Recordset")
	Set rsDeliveryBoard = cnnDeliveryBoard.Execute(SQLDeliveryBoard)
	
	'Write audit trail for delivery
	'*******************************
	Description = GetUserDisplayNameByUserNo(Session("UserNo")) & " set the delivery staus of invoice " & IvsNum  & " for customer " & CustNum & " to No Delivery at " & NOW()
	CreateAuditLogEntry "Delivery Status Changed","Delivery Status Changed","Minor",0,Description 
	
	Set rsDeliveryBoard = Nothing
	cnnDeliveryBoard.Close
	Set cnnDeliveryBoard = Nothing
	
	'Commented <!--#include file="sendalertsByInv.asp"-->
	
End If

Response.redirect("viewInvoices.asp?c=" & CustNum)
%>