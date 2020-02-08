<!--#include file="../../../inc/InsightFuncs.asp"-->
<!--#include file="../../../inc/InSightFuncs_routing.asp"-->
<%
IvsNum = Request.QueryString("i")
CustNum = Request.QueryString("c")

If  IvsNum <> "" Then
	
	SQLDeliveryBoard = "Update RT_DeliveryBoard Set DeliveryStatus = NULL , LastDeliveryStatusChange = NULL, DriverComments = NULL, ManualNextStop = 0, ManualNextStopChanged = NULL, DeliveryInProgress = 0 Where IvsNum = " & IvsNum
	
	Set cnnDeliveryBoard = Server.CreateObject("ADODB.Connection")
	cnnDeliveryBoard.open (Session("ClientCnnString"))
	Set rsDeliveryBoard = Server.CreateObject("ADODB.Recordset")
	Set rsDeliveryBoard = cnnDeliveryBoard.Execute(SQLDeliveryBoard)
	
	'Write audit trail for delivery
	'*******************************
	Description = GetUserDisplayNameByUserNo(Session("UserNo")) & " reset the delivery staus of invoice " & IvsNum  & " for customer " & CustNum & " at " & NOW()
	CreateAuditLogEntry "Delivery Status Changed","Delivery Status Changed","Minor",0,Description 
	
	Set rsDeliveryBoard = Nothing
	cnnDeliveryBoard.Close
	Set cnnDeliveryBoard = Nothing

End If

Response.redirect("viewInvoices.asp?c=" & CustNum)
%>