<!--#include file="../inc/header-field-service-mobile.asp"-->
<!--#include file="../inc/InsightFuncs_Service.asp"-->
<%

SelectedMemoNumber = Request.Form("txtTicketNumber")

If SelectedMemoNumber = "" Then
	SelectedMemoNumber = Request.QueryString("t")
End If
%>

<style type="text/css">
body{
	overflow-x: hidden;
}

*, ::after, ::before {
    box-sizing: border-box;
}

.fieldservice-heading{
		line-height: 1;
		padding-top: 5px;
		padding-bottom: 5px;
	}
	
	.btn-home{
		color: #fff;
		margin-top: -2px;
		margin-left: 5px;
		float: left;
 	}
	

.input-lg::-webkit-input-placeholder, textarea::-webkit-input-placeholder {
  color: #666;
}
.input-lg:-moz-placeholder, textarea:-moz-placeholder {
  color: #666;
}
.checkboxes label{
	font-weight: normal;
	margin-right: 20px;
}
.close-service-client-output{
	text-align: left;
}
.ticket-details{
	margin-bottom: 15px;
}

.row-common{
	width: 100%;
}

.row-common table td{
	width: 50%;
} 

.row{
	margin-right: 0px !important;
	margin-left: 0px !important;
}

.box{
	border: 1px solid #dbdece;
	padding-top: 10px;
	padding-bottom: 10px;
	margin-bottom: 10px;
	font-size: 12px;
	margin-right: -1px;
}

.container-fluidP{
	padding-left: 25px !important;
	padding-right: 25px !important;
}
</style>

<h1 class="fieldservice-heading" ><a href="#" class="btn-home" onClick="history.go(-1); return false;"><i class="fa fa-arrow-left"></i>
</a>Ticket # <%=SelectedMemoNumber%></h1>


<!-- field service menu starts here !-->
<div class="container-fluid">
	<div class="row">
	<%
	
	SQL = "Select * from FS_ServiceMemos WHERE MemoNumber = '" & SelectedMemoNumber & "'"
	
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3	
	set rs = cnn8.Execute (SQL)
	If not rs.EOF then 
		SelectedCustomer = rs("AccountNumber")
	End If
	Set rs = Nothing
	cnn8.Close
	Set cnn8 = Nothing
	
	%>
	<!--#include file="commonTicketDisplaypanel.asp"-->
	<!--#include file="commonCustomerDisplaypanel.asp"-->
	
	</div>
</div>

<%
If FS_TechCanDecline() <> True Then 
	'A sneaky thing to do, if you tap to view the ticket, but you haven't ACKed yet
	'We ACK it for you. After all, you did just look at the details
	
	'See if there is a dispatch ACKed record & if not, ACK it
	SQL = "Select * from FS_ServiceMemosDetail Where MemoNumber = '" & SelectedMemoNumber  & "' AND MemoStage='Dispatch Acknowledged'"
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	Set rs = cnn8.Execute(SQL)
	If rs.Eof then ' not found, so ACK it
		SQL = "INSERT INTO FS_ServiceMemosDetail (MemoNumber, CustNum, MemoStage, "
		SQL = SQL & " SubmissionDateTime, USerNoSubmittingRecord,UserNoOfServiceTech,OriginalDispatchDateTime)"
		SQL = SQL &  " VALUES (" 
		SQL = SQL & "'"  & SelectedMemoNumber & "'"
		SQL = SQL & ",'"  & SelectedCustomer & "'"
		SQL = SQL & ",'Dispatch Acknowledged'"
		SQL = SQL & ",getdate() "
		SQL = SQL & ","  & Session("UserNo") 
		SQL = SQL & ","  & Session("UserNo")
		SQL = SQL & ", '" & TicketOriginalDispatchDateTime(SelectedMemoNumber) & "')"
		'response.write(SQL)
		Set rs = cnn8.Execute(SQL)
	
		
		'Write audit trail for dispatch
		'*******************************
		Description = GetUserDisplayNameByUserNo(Session("UserNo")) & " acknowldged dispatch notification for service ticket number " & ServiceTicketNumber & " at " & NOW()
		CreateAuditLogEntry "Service Ticket System","Dispatch Acknowledged","Minor",0,Description 
		
		Set rs = Nothing
		cnn8.Close
		Set cnn8 = Nothing
	End If
End If
%><!--#include file="../inc/footer-field-service-noTimeout.asp"-->