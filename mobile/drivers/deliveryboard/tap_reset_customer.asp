<!--#include file="../../../inc/InsightFuncs.asp"-->
<!--#include file="../../../inc/InSightFuncs_routing.asp"-->
<%
CustNum = Request.QueryString("c")
RetPage = Request.QueryString("r")
If RetPage = "" Then RetPage = "main"
'Response.Write("CustNum :" & CustNum & "<br>")
'Response.End

If  CustNum <> "" Then
	
	'Mark all invoices for this customer as delivered
	
	SQLDeliveryBoard = "Update RT_DeliveryBoard Set DeliveryStatus = NULL , LastDeliveryStatusChange = NULL, DriverComments = NULL, ManualNextStop = 0, ManualNextStopChanged = NULL, DeliveryInProgress = 0 Where CustNum = '" & CustNum & "'"
	
	Set cnnDeliveryBoard = Server.CreateObject("ADODB.Connection")
	cnnDeliveryBoard.open (Session("ClientCnnString"))
	Set rsDeliveryBoard = Server.CreateObject("ADODB.Recordset")
	Set rsDeliveryBoard = cnnDeliveryBoard.Execute(SQLDeliveryBoard)
	
	'Write audit trail for delivery
	'*******************************
	Description = GetUserDisplayNameByUserNo(Session("UserNo")) & " reset the delivery staus of all delivery invoices for customer " & CustNum & " at " & NOW()
	CreateAuditLogEntry "Delivery Status Reset","Delivery Status Reset","Minor",0,Description 
	
	Set rsDeliveryBoard = Nothing
	cnnDeliveryBoard.Close
	Set cnnDeliveryBoard = Nothing

End If

If RetPage = "main" Then Response.redirect("main.asp")
If RetPage = "stop" Then Response.redirect("viewstops.asp")
%>