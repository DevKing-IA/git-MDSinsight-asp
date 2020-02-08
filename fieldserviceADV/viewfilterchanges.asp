<!--#include file="../inc/header-field-service-mobile.asp"-->

<%SelectedAssetNumber = Request.Form("txtAssetNumber")%>

<style>
#PleaseWaitPanel{
position: fixed;
left: 20px;
top: 50px;
width: 200px;
height: 100px;
z-index: 9999;
background-color: #fff;
opacity:1.0;
text-align:center;
}    
</style>

<div id="PleaseWaitPanel">
	<br><br>Processing, please wait...<br><br>
	<img src="../img/loading.gif"/>
</div>

<script type="text/javascript">
$(document).ready(function() {
    $("#PleaseWaitPanel").hide();
});
</script>


<script>
  function myFunction(asst)
	  {   
$('#PleaseWaitPanel').show();
		  var  assetnum=asst;
				
		   if(asst!='')
		   {
		    $.ajax({
		   type:'post',
		      url:'assigntomefilter.asp',
		          data:{txtAssetNumber: assetnum},
					success: function(msg){
						window.location = "assignToMeFilter_Done.asp";
						$('#PleaseWaitPanel').hide();
					}
		 });
		  }
	}
	
</script>

</script>

<style type="text/css">
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
</style>

<h1 class="fieldservice-heading" ><a class="btn btn-default btn-home pull-left" href="filterchanges.asp" role="button"><i class="fa fa-home"></i>
Back</a>Details</h1>

<!-- buttons start here !-->
<div class="container-fluid buttons-fluid">
	<div class="row">
		
		<!-- Assign To Me button !-->
		<div class="col-lg-6 col-md-6 col-sm-6 col-xs-6 button-box">
			<form method="post" action="AssignToMeFilter.asp" name="frmAssignToMeFilter" id="frmAssignToMeFilter">
				<a href='#'><input class="btn btn-primary btn-block" type='button' name='btnAssign' id='btnAssign' onclick="myFunction('<%=SelectedAssetNumber%>')" value="Assign To Me"></a>
			</form>
		</div>
		<!-- eof Assign To Me button !-->
	</div>
</div>
<!-- buttons end here !-->
<br>
<!-- field service menu starts here !-->
<div class="container-fluid fieldservice-container">
	<div class="row">
	<%
	
	SQL = "SELECT * FROM Assets "
	SQL = SQL & "INNER JOIN EQ_ScheduledServiceDates ON Assets.assetNumber = EQ_ScheduledServiceDates.assetNumber "
	SQL = SQL & "WHERE Assets.assetNumber = '" & SelectedAssetNumber & "'"

	
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3	
	set rs = cnn8.Execute (SQL)
	If not rs.EOF then 
		SelectedCustomer = rs("custAcctNum")
	End If
	
	%>
<style type="text/css">
	.row-common{
		border: 1px solid #dbdece;
		padding-top: 10px;
		padding-bottom: 10px;
		margin-bottom: 10px;
		font-size: 12px;
	}
</style>

<!-- row !-->
<div class="row row-common">
	<div class="col-lg-6">
		<table style="width:100%;">
			<tr><td align="right"><b>Filter Info:</b></td><td>&nbsp;&nbsp;&nbsp;</td><td align="left"><%=rs("Comment1")%>
			<tr><td align="right"><b>Date:</b></td><td>&nbsp;&nbsp;&nbsp;</td><td align="left"><%=rs("nextDate1")%></td></tr>
		</table>
	</div>
</div>
	<%Set rs = Nothing
	cnn8.Close
	Set cnn8 = Nothing%>

	<!--#include file="commonCustomerDisplaypanel.asp"-->
	
	</div>
</div>
<!--#include file="../inc/footer-field-service-noTimeout.asp"-->