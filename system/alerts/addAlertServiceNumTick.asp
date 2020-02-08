<!--#include file="../../inc/header.asp"-->
<%
refFld = ""
alrtName = ""
If MUV_Read("refFld") <> "" Then refFld = MUV_ReadAndRemove("refFld")
If MUV_Read("alrtName") <> "" Then alrtName = MUV_ReadAndRemove("alrtName")
%>
<link rel="stylesheet" type="text/css" href="<%= BaseURL %>css/tabs.css" />


<script>

 function loadRefValues()
  {   
	var refFld = document.getElementById("selRefField").value;
	var alrtName = document.getElementById("txtAlertName").value;

		    $.ajax({
		   type:'post',
		      url:'SetFieldValuesNumTick.asp',
		          data:{refFld: refFld,alrtName: alrtName},
					success: function(msg){
						window.location = "addAlertServiceNumTick.asp";
					}
		 });
  }
</script>



<!-- password strength meter !-->

<style type="text/css">

.pass-strength h5{
	margin-top: 0px;
	color: #000;
}
.popover.primary {
    border-color:#337ab7;
}
.popover.primary>.arrow {
    border-top-color:#337ab7;
}
.popover.primary>.popover-title {
    color:#fff;
    background-color:#337ab7;
    border-color:#337ab7;
}
.popover.success {
    border-color:#d6e9c6;
}
.popover.success>.arrow {
    border-top-color:#d6e9c6;
}
.popover.success>.popover-title {
    color:#3c763d;
    background-color:#dff0d8;
    border-color:#d6e9c6;
}
.popover.info {
    border-color:#bce8f1;
}
.popover.info>.arrow {
    border-top-color:#bce8f1;
}
.popover.info>.popover-title {
    color:#31708f;
    background-color:#d9edf7;
    border-color:#bce8f1;
}
.popover.warning {
    border-color:#faebcc;
}
.popover.warning>.arrow {
    border-top-color:#faebcc;
}
.popover.warning>.popover-title {
    color:#8a6d3b;
    background-color:#fcf8e3;
    border-color:#faebcc;
}
.popover.danger {
    border-color:#ebccd1;
}
.popover.danger>.arrow {
    border-top-color:#ebccd1;
}
.popover.danger>.popover-title {
    color:#a94442;
    background-color:#f2dede;
    border-color:#ebccd1;
}

.select-line{
	margin-bottom: 15px;
}

.enable-disable{
	margin-top:20px;
}

.row-line{
	margin-bottom: 25px;
}

.table th, tr, td{
	font-weight: normal;
}

.table>thead>tr>th{
	border: 0px;
}
.table thead>tr>th,.table tbody>tr>th,.table tfoot>tr>th,.table thead>tr>td,.table tbody>tr>td,.table tfoot>tr>td{
border:0px;
}

.when-col{
	width: 10%;
}

.reference-col{
	width: 45%;
}

.has-more-col{
	width: 12%;
}

.form-control{
	min-width: 100px;
}

.textarea-box{
	min-width: 260px;
}
	</style>
<!-- eof password strength meter !-->


<h1 class="page-header"><i class="fa fa-fw fa-exclamation"></i>New Service Alert</h1>


<form method="POST" action="addAlertServiceNumTick_submit.asp" name="frmAddAlert" id="frmAddAlert">

	<!-- alert name !-->
	<div class="row row-line">
		<div class="col-xs-6 col-sm-1 col-md-1 col-lg-2">
			Alert Name <input type="text" id="txtAlertName" name="txtAlertName" value="<%= alrtName %>"  class="form-control last-run-inputs">
		</div>
	</div>


	<!-- reference line !-->
	<div class="row row-line">
		<div class="col-lg-6">

			<table class="table">
				<thead>
					<tr>
						<th class="when-col">&nbsp;</th>
						<th class="reference-col">Reference Field</th>
						<th class="reference-col">Reference Value</th>
					</tr>
				</thead>

				<tbody>
					<tr>
						<th scope="row">When</th>
							<td>
								<select class="form-control"  onchange="loadRefValues()" id="selRefField" name="selRefField">
									<%
									Response.Write("<option>Any</option>")
									If refFld = "Referral" Then
										Response.Write("<option selected value='Referral'>" & GetTerm("Referral") & "</option>")
									Else
										Response.Write("<option value='Referral'>" & GetTerm("Referral") & "</option>")
									End If
									If refFld = "Primary Salesman" Then
										Response.Write("<option selected value='Primary Salesman'>" & GetTerm("Primary Salesman") & "</option>")
									Else
										Response.Write("<option value='Primary Salesman'>" & GetTerm("Primary Salesman") & "</option>")										
									End If
									If refFld = "Secondary Salesman" Then	
										Response.Write("<option selected value='Secondary Salesman'>" & GetTerm("Secondary Salesman") & "</option>")
									Else
										Response.Write("<option value='Secondary Salesman'>" & GetTerm("Secondary Salesman") & "</option>")
									End If
									If refFld = "Customer Type" Then	
										Response.Write("<option selected value='Customer Type'>" &  GetTerm("Customer") & " Type</option>")
									Else
										Response.Write("<option value='Customer Type'>" & GetTerm("Customer") & " Type</option>")
									End If									
									%>
								</select>
							</td>
							<td>
								<select class="form-control" id="selrefVal" name="selrefVal">
									<% If refFld = "" Then %>
										<option selected>Any</option>
									<% Elseif refFld = "Referral" Then %>
										<option selected>Any</option>
								      	<% 'Get all Referral options
								      	  	SQL = "SELECT ReferalCode, Name FROM Referal"
						
											Set cnn8 = Server.CreateObject("ADODB.Connection")
											cnn8.open (Session("ClientCnnString"))
											Set rs = Server.CreateObject("ADODB.Recordset")
											rs.CursorLocation = 3 
											Set rs = cnn8.Execute(SQL)
												
											If not rs.EOF Then
												Do
													Response.Write("<option value='" & rs("ReferalCode") & "'>" & rs("ReferalCode") & " - " & rs("Name") & "</option>")
													rs.movenext
												Loop until rs.eof
											End If
											set rs = Nothing
											cnn8.close
											set cnn8 = Nothing
								      	%>
									<% Elseif refFld = "Primary Salesman" Then %>
										<option selected>Any</option>
								      	<% 'Get all Slsmn 1 options
								      	  	SQL = "SELECT DISTINCT SalesmanSequence, Salesman.Name FROM Salesman "
								      	  	SQL = SQL & "Inner Join Customer on Salesman = SalesmanSequence "
								      	  	SQL = SQL & "order by SalesmanSequence "
						
											Set cnn8 = Server.CreateObject("ADODB.Connection")
											cnn8.open (Session("ClientCnnString"))
											Set rs = Server.CreateObject("ADODB.Recordset")
											rs.CursorLocation = 3 
											Set rs = cnn8.Execute(SQL)
												
											If not rs.EOF Then
												Do
													Response.Write("<option value='" & rs("SalesmanSequence") & "'>" & rs("SalesmanSequence") & " - " & rs("Name") & "</option>")
													rs.movenext
												Loop until rs.eof
											End If
											set rs = Nothing
											cnn8.close
											set cnn8 = Nothing
								      	%>
									<% Elseif refFld = "Secondary Salesman" Then %>
										<option selected>Any</option>
								      	<% 'Get all Slsmn 2 options
								      	  	SQL = "SELECT DISTINCT SalesmanSequence, Salesman.Name FROM Salesman "
								      	  	SQL = SQL & "Inner Join Customer on SecondarySalesman = SalesmanSequence "
								      	  	SQL = SQL & "order by SalesmanSequence "
											Set cnn8 = Server.CreateObject("ADODB.Connection")
											cnn8.open (Session("ClientCnnString"))
											Set rs = Server.CreateObject("ADODB.Recordset")
											rs.CursorLocation = 3 
											Set rs = cnn8.Execute(SQL)
												
											If not rs.EOF Then
												Do
													Response.Write("<option value='" & rs("SalesmanSequence") & "'>" & rs("SalesmanSequence") & " - " & rs("Name") & "</option>")
													rs.movenext
												Loop until rs.eof
											End If
											set rs = Nothing
											cnn8.close
											set cnn8 = Nothing
								      	%>
									<% Elseif refFld = "Customer Type" Then %>
								      	<option selected>Any</option>
								      	<% 'Get all Type options
								      	  	SQL = "SELECT CustTypeSequence, Description FROM CustomerType"
						
											Set cnn8 = Server.CreateObject("ADODB.Connection")
											cnn8.open (Session("ClientCnnString"))
											Set rs = Server.CreateObject("ADODB.Recordset")
											rs.CursorLocation = 3 
											Set rs = cnn8.Execute(SQL)
												
											If not rs.EOF Then
												Do
													Response.Write("<option value='" & rs("CustTypeSequence") & "'>" & rs("CustTypeSequence") & " - " & rs("Description") & "</option>")
													rs.movenext
												Loop until rs.eof
											End If
											set rs = Nothing
											cnn8.close
											set cnn8 = Nothing
								      	%>
									<% End If%>
								</select>
							</td>
						 </tr>
					 </tbody>
				</table>
			</div>
		</div>


		<div class="row row-line">
		 	 <div class="col-lg-12">
			 	 <div class="table-responsive">
				 	 <table class="table">
					 	 <thead>
						 	 <tr>
								<th>&nbsp;</th>
								<th>&nbsp;</th>
								<th>&nbsp;</th>
								<th>&nbsp;</th>
								<th>&nbsp;</th>
								<th>&nbsp;</th>
								<th>&nbsp;</th>
						 	 </tr>
	 	 			 	 </thead>
						<tbody>
							<tr>
								<th scope="row" >has MORE than</th>
									<td>
										<select class="form-control" id="selNumTickets" name="selNumTickets">
											<%For x = 0 to 50
											Response.Write("<option>" & x & "</option>")
											Next %>
										 </select>
									</td>
									<td align="center">
										service tickets in a period of
									</td>
									<td>
										<select class="form-control" id="selNumDays" name="selNumDays">
											<%For x = 1 to 120
											Response.Write("<option>" & x & "</option>")
											Next %>
										</select>
									</td>
									<td align="center">
										days, send email to
									</td>
									<td>
										<select class="form-control" id="selSendTo" name="selSendTo">
											<option value="None">None from here</option>
											<option value="Primary Salesman"><%=GetTerm("Primary Salesman")%></option>
											<option value="Secondary Salesman"><%=GetTerm("Secondary Salesman")%></option>
											<option value="BDR"><%=GetTerm("BDR")%></option>
										</select>
									</td>
									<td align="center">
										and
									</td>
									<td>
										<textarea class="form-control textarea-box" rows="3" id="txtEmails" name="txtEmails"></textarea>
											<small><strong>Separate multiple email addresses with a semicolon</strong></small>
									</td>
								</tr>
						 </tbody>
					</table>
			 	 </div>
		 	 </div>
	 	 </div>

		<div class="row row-line">
			<div class="col-lg-6">
				<strong>Verbiage to include in alert email</strong>
			    <textarea class="form-control" rows="3" id="txtVerbiage" name="txtVerbiage"></textarea>
			</div>
		

			<!-- enabled line !-->
		   	<div class="col-xs-6 col-sm-1 col-md-1 col-lg-3 enable-disable">
		  	  <strong>Enabled</strong> <input type="checkbox" checked id="chkEnabled"  name="chkEnabled">
		    </div>
		    <!-- eof enabled line !-->
	    </div>
	    
		<div class="row row-line">
			<div class="col-lg-12 alertbutton">
				<a href="<%= BaseURL %>system/alerts/main.asp#ServiceNumTicks">
    				<button type="button" class="btn btn-default">&lsaquo; Cancel &amp; Go Back To Alert List</button>
				</a>
				<button type="submit" class="btn btn-primary"><i class="far fa-save"></i> Save</button>
		    </div>
		</div>
</form>


<!--#include file="../../inc/footer-main.asp"-->
