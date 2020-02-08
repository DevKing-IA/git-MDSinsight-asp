<!--#include file="../../inc/header.asp"-->

<% InternalAlertRecNumber = Request.QueryString("a") 
If InternalAlertRecNumber = "" Then Response.Redirect("main.asp")
%>

<link rel="stylesheet" type="text/css" href="<%= BaseURL %>css/tabs.css" />

          
<script>

 function loadRefValues()
  {   
	var refFld = document.getElementById("selRefField").value;

		    $.ajax({
		   type:'post',
		      url:'SetFieldValuesSesVar.asp',
		          data:{refFld: refFld},
					success: function(msg){
						window.location = "editAlertServiceNumTick.asp";
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
	margin-top: 20px;
}
</style>
<!-- eof password strength meter !-->


<h1 class="page-header"><i class="fa fa-fw fa-exclamation"></i>Edit Service Alert</h1>

<%
SQL = "SELECT * FROM SC_Alerts where InternalAlertRecNumber = " & InternalAlertRecNumber 

Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3 
Set rs = cnn8.Execute(SQL)
	
If not rs.EOF Then
	AlertName = rs("AlertName")
	ReferenceField = rs("ReferenceField")
	ReferenceValue = rs("ReferenceValue")
	NumberOfTickets = rs("NumberOfTickets")
	NumberOfDays = rs("NumberOfDays")
	SendAlertTo = rs("SendAlertTo")
	AdditionalEmails = rs("AdditionalEmails")
	AlertEmailVerbiage = rs("AlertEmailVerbiage")
	Enabled = rs("Enabled")
	InternalAlertRecNumber = rs("InternalAlertRecNumber")
End If
set rs = Nothing
cnn8.close
set cnn8 = Nothing

refFld = ReferenceField 
refVal = ReferenceValue 
If Enabled = True then Enabled = 1 Else Enabled =0
%>



<!-- row !-->
<div class="row">

	<div class="col-lg-12">
	
		<form method="POST" action="editAlertServiceNumTick_submit.asp" name="frmeditAlert" onsubmit="return validateAlertForm();">		    
      
			
	<!-- alert name !-->
	<div class="row row-line">
		<div class="col-xs-6 col-sm-1 col-md-1 col-lg-2">
			<strong>Alert Name</strong> <input type="text" id="txtAlertName" name="txtAlertName" value="<%=AlertName%>"  class="form-control last-run-inputs">
			<input type="hidden" id="txtInternalAlertRecNumber" name="txtInternalAlertRecNumber" value="<%=InternalAlertRecNumber %>"  class="form-control last-run-inputs">
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
									If refFld = "Any" Then
										Response.Write("<option selected >Any</option>")
									Else
										Response.Write("<option>Any</option>")
									End If
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
									<% If refFld = "Any" Then %>
										<option selected value="Any">Any</option>
									<% Elseif refFld = "Referral" Then %>
										<option selected value="<%= refVal %>"><%=GetReferralNameByCode(refVal)%></option>
									<% Elseif refFld = "Primary Salesman" Then %>
										<option selected value="<%= refVal %>"><%=GetSalesmanNameBySlsmnSequence(refVal)%></option>
									<% Elseif refFld = "Secondary Salesman" Then %>
										<option selected value="<%= refVal %>"><%=GetSalesmanNameBySlsmnSequence(refVal)%></option>
									<% Elseif refFld = "Customer Type" Then %>
										<option selected value="<%= refVal %>"><%=GetCustTypeByCode(refVal)%></option>
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
												If x = NumberOfTickets Then
													Response.Write("<option selected >" & x & "</option>")
												Else
													Response.Write("<option>" & x & "</option>")												
												End If
											Next %>
										 </select>
									</td>
									<td align="center">
										<strong>service tickets in a period of</strong>
									</td>
									<td>
										<select class="form-control" id="selNumDays" name="selNumDays">
											<%For x = 1 to 120
											If x = NumberOfDays Then
												Response.Write("<option selected >" & x & "</option>")
											Else
												Response.Write("<option>" & x & "</option>")											
											End If
											Next %>
										</select>
									</td>
									<td align="center">
										<strong>days, send email to</strong>
									</td>
									<td>
										<select class="form-control" id="selSendTo" name="selSendTo">
											<%
												If SendAlertTo ="None" Then
													Response.Write("<option selected value='None'>None from here</option>")
												Else
													Response.Write("<option value='None'>None from here</option>")											
												End If
												If SendAlertTo ="Primary Salesman" Then
													Response.Write("<option selected >" & GetTerm("Primary Salesman") &"</option>")
												Else
													Response.Write("<option>" & GetTerm("Primary Salesman") &"</option>")
												End If
												If SendAlertTo ="Secondary Salesman" Then
													Response.Write("<option selected >" & GetTerm("Secondary Salesman") &"</option>")
												Else
													Response.Write("<option>" & GetTerm("Secondary Salesman") &"</option>")
												End If
												If SendAlertTo ="BDR" Then
													Response.Write("<option selected >" & GetTerm("BDR") &"</option>")
												Else
													Response.Write("<option>" & GetTerm("BDR") &"</option>")
												End If
											%>
										</select>
									</td>
									<td align="center">
										<strong>and</strong>
									</td>
									<td>
										<textarea class="form-control textarea-box" rows="3" id="txtEmails" name="txtEmails"><%=AdditionalEmails%></textarea>
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
			    <textarea class="form-control" rows="3" id="txtVerbiage" name="txtVerbiage"><%=AlertEmailVerbiage%></textarea>
			</div>
		

			<!-- enabled line !-->
		   	<div class="col-xs-6 col-sm-1 col-md-1 col-lg-3 enable-disable">
		   		<% If Enabled = 1 Then %>
				  	  <strong>Enabled</strong> <input type="checkbox" checked id="chkEnabled"  name="chkEnabled">
				<% Else %>
				  	  <strong>Enabled</strong> <input type="checkbox" id="chkEnabled"  name="chkEnabled">				
				<%End If%>
		    </div>
		    <!-- eof enabled line !-->
	    </div>
	    <br>
		<div class="row row-line">
			<div class="col-lg-12 alertbutton">
				<a href="<%= BaseURL %>system/alerts/main.asp#ServiceNumTicks">
    				<button type="button" class="btn btn-default">&lsaquo; Cancel &amp; Go Back To Alert List</button>
				</a>
				<button type="submit" class="btn btn-primary"><i class="far fa-save"></i> Save</button>
		    </div>
		</div>
</form>
	</div>	

</div>
<!-- eof row !-->    

   
<!--#include file="../../inc/footer-main.asp"-->
