<!--#include file="../../../inc/header.asp"-->

<%

	SQL = "SELECT * FROM Settings_Global"
	
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rs = cnn8.Execute(SQL)
	
	If not rs.EOF Then
		ShowOpenPopupMessage =rs("NotesScreenShowPopup")		
	End If
				
	set rs = Nothing
	cnn8.close
	set cnn8 = Nothing
	
%>

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

<h1 class="page-header"><i class="fa fa-globe"></i>&nbsp;<%=GetTerm("Customer Service")%> Settings
	<a href="<%= BaseURL %>admin/global/main.asp"><button class="btn btn-small btn-secondary pull-right" style="margin-left:20px"><i class="fas fa-arrow-alt-left"></i>&nbsp;<i class="fas fa-globe"></i>&nbsp;GLOBAL SETTINGS MAIN</button></a>
</h1>


<form method="post" action="client-care-settings-submit.asp" name="frmCustServiceSettings" id="frmCustServiceSettings">

<div class="container">
	
	<%
		Response.Write("<div id='PleaseWaitPanel' style='display:none;'>")
		Response.Write("<br><br>Saving your recent changes, please wait...<br><br>")
		Response.Write("<img src=""" & baseURL & "/img/loading.gif"" />")
		Response.Write("</div>")
		Response.Flush()
	%>

	<!-- three cols !-->
	<div class="row">
	<br>
		<div class="col-lg-2">
			<strong><%=GetTerm("Customer Service")%></strong>
		</div>
	
		<div class="col-lg-6">
	
			<!-- row with data !-->
			<div class="row row-data">
				<div class="col-lg-8">Show popup on notes screen when <%=GetTerm("customer")%> has open serivce tickets.</div>
				<div class="col-lg-2">
					<%
					If ShowOpenPopupMessage = 0 Then
						Response.Write("<input type='checkbox' class='check' id='chkShowOpenPopupMessage' name='chkShowOpenPopupMessage'")
					Else
						Response.Write("<input type='checkbox' class='check' id='chkShowOpenPopupMessage' name='chkShowOpenPopupMessage' checked")
					End If
					Response.Write(">")
					%>
				</div>
			</div>
			<!-- eof row with data !-->

		</div>
		
	</div>
	
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
