<!--#include file="../../../inc/header.asp"-->

<style>

	.container {
		margin-bottom: 20px;
		margin-top: 20px;
		margin-left:0px;
		width: 100%;
	}

	.container .row {
		margin-bottom: 20px;
		margin-top: 20px;
	}
	
	#PleaseWaitPanel{
		position: fixed;
		left: 470px;
		top: 275px;
		width: 975px;
		height: 300px;
		z-index: 9999;
		background-color: #fff;
		opacity:1.0;
		text-align:center;
	}   

	.btn-huge{
	    padding: 18px 28px;
	    font-size: 22px;	    
	}	
</style>

<script>

	function showSavingChangesDiv() {
	  document.getElementById('PleaseWaitPanel').style.display = "block";
	  setTimeout(function() {
	    document.getElementById('PleaseWaitPanel').style.display = "none";
	  },1500);
	   
	}
	
</script>

<%
	SQL = "SELECT * FROM Settings_Global"
	
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rs = cnn8.Execute(SQL)
	
	If not rs.EOF Then
		EZTextID = rs("EZTextingID")
		EZTextPassword = rs("EZTextingPassword")			
	End If
				
	set rs = Nothing
	cnn8.close
	set cnn8 = Nothing
	
%>


<h1 class="page-header"><i class="fa fa-globe"></i>&nbsp;Texting Settings
	<a href="<%= BaseURL %>admin/global/main.asp"><button class="btn btn-small btn-secondary pull-right" style="margin-left:20px"><i class="fas fa-arrow-alt-left"></i>&nbsp;<i class="fas fa-globe"></i>&nbsp;GLOBAL SETTINGS MAIN</button></a>
</h1>

<form method="post" action="texting-settings-submit.asp" name="frmTexting" id="frmTexting">

<div class="container">

	<%
		Response.Write("<div id='PleaseWaitPanel' style='display:none;'>")
		Response.Write("<br><br>Saving your recent changes, please wait...<br><br>")
		Response.Write("<img src=""" & baseURL & "/img/loading.gif"" />")
		Response.Write("</div>")
		Response.Flush()
	%>

	<div class="row">

		<div class="col-lg-2">&nbsp;</div>
		
			<div class="col-lg-6">

				<div class="row">
					<div class="col-lg-3">Ez-Texting Username </div>
				   	<div class="col-lg-2">
				       	<input type='text' class="form-control" id="txtEZTextingID" name="txtEZTextingID" value="<%= EZTextID %>"> 
					</div>
				</div>
			    
		        <div class="row">
			        <div class="col-lg-3">Ez-Texting Password </div>
				   	<div class="col-lg-2">
						<input type='text' class="form-control" id="txtEZTextingPassword" name="txtEZTextingPassword"  value="<%= EZTextPassword %>">
					</div>
				</div>
				
			</div><!-- col-lg-6-->
	
	</div><!-- row -->
	
	<!-- cancel / save !-->
	<div class="row pull-right">
		<div class="col-lg-12">
			<a href="<%= BaseURL %>admin/global/main.asp"><button type="button" class="btn btn-default btn-lg btn-huge"><i class="far fa-times-circle"></i> Cancel</button></a> 
			<button type="submit" class="btn btn-primary btn-lg btn-huge" onclick="showSavingChangesDiv()"><i class="far fa-save"></i> Save Changes</button>
		</div>
	</div>
	
	
</div><!-- container -->

</form>

<!--#include file="../../../inc/footer-main.asp"-->
