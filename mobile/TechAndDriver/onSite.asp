<!--#include file="inc/header-tech-and-driver.asp"-->

<%SelectedMemoNumber = Request.Form("txtTicketNumber")%>


<style type="text/css">
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

.btn{
	margin-bottom: 10px;
	font-size: 10px;
}
.button-box{
	padding:1px;
 }
 
 .buttons-fluid{
	 margin: 10px;
 }
 
.btn span{
font-size:14px;
} 
</style>

<h1 class="fieldservice-heading" ><a class="btn-home" href="main.asp" role="button"><i class="fa fa-arrow-left"></i></a> Ticket # <%=SelectedMemoNumber%></h1>

<!-- buttons start here !-->
<div class="container-fluid buttons-fluid">
	<div class="row">
		
		<!-- Close button !-->
		<div class="col-lg-6 col-md-6 col-sm-6 col-xs-6 button-box">
			<form method="post" action="CloseService_PassThru.asp" name="frmCloseTicket" id="frmCloseTicket">
				<input type='hidden' id='txtTicketNumber' name='txtTicketNumber' value='<%=SelectedMemoNumber%>'>		 
				<button type="Submit" class="btn btn-primary btn-block"><span>Close</span></button>
			</form>
		</div>
		<!-- eof Close button !-->
		
		<!-- Wait for Parts button !-->
		<div class="col-lg-6 col-md-6 col-sm-6 col-xs-6 button-box">
			<form method="post" action="waitForParts_PassThru.asp" name="frmwaitForParts" id="frmwaitForParts">
				<input type='hidden' id='txtTicketNumber' name='txtTicketNumber' value='<%=SelectedMemoNumber%>'>		 
				<button type="Submit" class="btn btn-primary btn-block"><span>Wait For Parts</span></button>
			</form>
		</div>
		<!-- eof Wait for Parts button !-->
		
		<!-- Followup button !-->
		<div class="col-lg-6 col-md-6 col-sm-6 col-xs-6 button-box">
			<form method="post" action="followup_PassThru.asp" name="frmfollowup" id="frmfollowup">
				<input type='hidden' id='txtTicketNumber' name='txtTicketNumber' value='<%=SelectedMemoNumber%>'>		 
				<button type="Submit" class="btn btn-primary btn-block"><span>Follow Up</span></button>
			</form>
		</div>
		<!-- eof Followup button !-->
		
		<!-- Swap button !-->
		<% If TicketIsFilterChange(SelectedMemoNumber) <> True AND TicketIsPMCall(SelectedMemoNumber) <> True Then 'No swap on filter change%>
		<div class="col-lg-6 col-md-6 col-sm-6 col-xs-6 button-box">
			<form method="post" action="swap_PassThru.asp" name="frmswap" id="frmswap">
				<input type='hidden' id='txtTicketNumber' name='txtTicketNumber' value='<%=SelectedMemoNumber%>'>		 
				<button type="Submit" class="btn btn-primary btn-block"><span>Swap</span></button>
			</form>
		</div>
		<%End IF%>
		<!-- eof Swap button !-->
		
		<!-- Unable to work button !-->
		<div class="col-lg-6 col-md-6 col-sm-6 col-xs-6 button-box">
			<form method="post" action="unableToWork_PassThru.asp" name="frmunableToWork" id="frmunableToWork">
				<input type='hidden' id='txtTicketNumber' name='txtTicketNumber' value='<%=SelectedMemoNumber%>'>		 
				<button type="Submit" class="btn btn-primary btn-block"><span>Unable To Work</span></button>
			</form>
		</div>
		<!-- eof Unable to work button !-->
		
	</div>
</div>
<!-- buttons end here !-->

<!-- field service menu starts here !-->
<div class="container-fluid fieldservice-container">
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

<%'A sneaky thing to do, if you tap to view the ticket, but you haven't ACKed yet
'We ACK it for you. After all, you did just look at the details

'See if there is a dispatch ACKed record & if not, ACK it
SQL = "Select * from FS_ServiceMemosDetail Where MemoNumber = '" & SelectedMemoNumber  & "' AND MemoStage='Dispatch Acknowledged'"
Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rs = Server.CreateObject("ADODB.Recordset")
Set rs = cnn8.Execute(SQL)
If rs.Eof then ' not found, so ACK it
	SQL = "INSERT INTO FS_ServiceMemosDetail (MemoNumber, CustNum, MemoStage, "
	SQL = SQL & " SubmissionDateTime, USerNoSubmittingRecord,UserNoOfServiceTech,Urgent,OriginalDispatchDateTime)"
	SQL = SQL &  " VALUES (" 
	SQL = SQL & "'"  & SelectedMemoNumber & "'"
	SQL = SQL & ",'"  & SelectedCustomer & "'"
	SQL = SQL & ",'Dispatch Acknowledged'"
	SQL = SQL & ",getdate() "
	SQL = SQL & ","  & Session("UserNo") 
	SQL = SQL & ","  & Session("UserNo")
	If TicketIsUrgent(SelectedMemoNumber) Then
		SQL = SQL & ",1" 'Urgent
	Else
		SQL = SQL & ",0" 'Not Urgent
	End If
	SQL = SQL & ", '" & TicketOriginalDispatchDateTime(SelectedMemoNumber) & "')"
	response.write(SQL)
	Set rs = cnn8.Execute(SQL)

	
	'Write audit trail for dispatch
	'*******************************
	Description = GetUserDisplayNameByUserNo(Session("UserNo")) & " acknowldged dispatch notification for service ticket number " & ServiceTicketNumber & " at " & NOW()
	CreateAuditLogEntry "Service Ticket System","Dispatch Acknowledged","Minor",0,Description 
	
	Set rs = Nothing
	cnn8.Close
	Set cnn8 = Nothing
End If

%><!--#include file="inc/footer-tech-and-driver.asp"-->