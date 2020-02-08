<!--#include file="../../../inc/InsightFuncs.asp"-->
<!--#include file="../../../inc/InSightFuncs_routing.asp"-->
<%
CustNum = Request.Form("txtCustNum")
DriverComments = Request.Form("txtdriverComments")
DriverComments = Replace(DriverComments,"'","''")
DriverComments = Server.HTMLEncode(DriverComments)


If  CustNum <> "" Then
	
	'Mark all invoices for this customer as not delivered
	
	SQLDeliveryBoard = "Update RT_DeliveryBoard Set DeliveryStatus = 'No Delivery' , LastDeliveryStatusChange = getdate(),ManualNextStop = 0, ManualNextStopChanged = NULL, DeliveryInProgress = 0 "
	If DriverComments <> "" Then SQLDeliveryBoard = SQLDeliveryBoard & ", DriverComments = '" & DriverComments & "' "
	SQLDeliveryBoard = SQLDeliveryBoard & " WHERE CustNum = '" & CustNum & "'"

	
	Set cnnDeliveryBoard = Server.CreateObject("ADODB.Connection")
	cnnDeliveryBoard.open (Session("ClientCnnString"))
	Set rsDeliveryBoard = Server.CreateObject("ADODB.Recordset")
	Set rsDeliveryBoard = cnnDeliveryBoard.Execute(SQLDeliveryBoard)
	
	'Write audit trail for delivery
	'*******************************
	Description = GetUserDisplayNameByUserNo(Session("UserNo")) & " set the delivery staus of all delivery invoices for customer " & CustNum & " to No Delivery at " & NOW()
	CreateAuditLogEntry "Delivery Status Changed","Delivery Status Changed","Minor",0,Description 

	
	'Commented <!--#include file="sendalertsByCust.asp"-->
	

	Set rsDeliveryBoard = Nothing
	cnnDeliveryBoard.Close
	Set cnnDeliveryBoard = Nothing

End If


If AutoPromptNextStopOn() = True Then
	If cInt(GetRemainingStopsByUserNo(Session("UserNo"))) > 0 Then
 		Response.redirect("viewStops.asp")
 	Else
 		Response.redirect("main.asp")
 	End If
Else
	Response.redirect("main.asp")
End If


%>