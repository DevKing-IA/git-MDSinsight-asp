<!--#include file="../../../inc/InsightFuncs.asp"-->
<!--#include file="../../../inc/InSightFuncs_routing.asp"-->
<%
CustNum = Request.QueryString("c")

If  CustNum <> "" Then
	
	'Set Sequence to 0 so it'll be next stop
	
	SQLDeliveryBoard = "Update RT_DeliveryBoard Set ManualNextStop = 0, ManualNextStopChanged = NULL Where CustNum = '" & CustNum & "'"
	
	Set cnnDeliveryBoard = Server.CreateObject("ADODB.Connection")
	cnnDeliveryBoard.open (Session("ClientCnnString"))
	Set rsDeliveryBoard = Server.CreateObject("ADODB.Recordset")
	Set rsDeliveryBoard = cnnDeliveryBoard.Execute(SQLDeliveryBoard)
	
	'Write audit trail for delivery
	'*******************************
	Description = GetUserDisplayNameByUserNo(Session("UserNo")) & " manually set the Next Stop to customer " & CustNum & " at " & NOW()
	CreateAuditLogEntry "Delivery Sequence Changed","Delivery Sequence Changed","Minor",0,Description 
	
	Set rsDeliveryBoard = Nothing
	cnnDeliveryBoard.Close
	Set cnnDeliveryBoard = Nothing

End If

Response.redirect("viewstops.asp")
%>