<!--#include file="../../../inc/header.asp"-->


<style type="text/css">

	label{
		font-weight: normal;
	}
	
	.row-cutoff{
		width: 100%;
	}

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
	
	.post-labels{
 		padding-top: 5px;
 	}
 	
	.schedule-info{
		margin:0px -5px 0px -5px;
	}
	
	.schedule-info [class^="col-"]{
		padding:3px;
		margin:0px;
	} 
	
	.schedule-info .form-control{
		width:80%;
	}

	.tab-colors-box{
		padding:15px;
		border:2px solid #000;
		margin:0px 0px 15px 0px;
		width:100%;
		display:block;
		float:left;
	}
	
	.tab-colors-title strong{
		width:100%;
		text-align:center;
		display:block;
	}
	
	.tab-colors-title .row{
		margin-bottom:0px;
	}
	
	.line-full{
	 	margin-bottom:20px;
	}
	
	.multi-select{
		min-height:200px;
		min-width:190px;
	}

</style>


<%
	SQL = "SELECT * FROM Settings_Quickbooks"
	
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rs = cnn8.Execute(SQL)
	
	If not rs.EOF Then
		ImportCustomersFromQB = rs("ImportCustomersFromQB")	
		ImportCustomersUpdateOrReplace = rs("ImportCustomersUpdateOrReplace")	
	End If
	
	
	set rs = Nothing
	cnn8.close
	set cnn8 = Nothing
	
%>


<h1 class="page-header"><i class="fa fa-globe"></i>&nbsp;Quickbooks Integration</h1>

<form method="post" action="quickbooks-submit.asp" name="frmQuickbooks" id="frmQuickbooks">

<div class="container">

	<!-- row with data !-->
	<div class="row">
	 	<div class="col-lg-12">
	
			<% If MUV_Read("quickbooksModuleOn")  = "Disabled" Then %>
				<div class="col-lg-6">
					<br><br>
					Please contact support if you would like to activate the Quickbooks Integration module.
				</div>
			<% ElseIf MUV_Read("quickbooksModuleOn")  = "Enabled" Then  %>

	            <div class="col-lg-4">
	                <div class="col-lg-12 tab-colors-title">
						<div class="row">
							<div class="col-lg-12" align="center">
								 <strong>Customer</strong>
							</div>
						</div>
					</div>
            		
					<div class="col-lg-12">
						<div class="tab-colors-box">
							<div class="row schedule-info" style="margin-top:0px;margin-bottom:2px;">
								<div class="col-lg-12">
									<%
									If ImportCustomersFromQB = 0 Then
									Response.Write("<input type='checkbox' class='check' id='chkImportCustomersFromQB' name='chkImportCustomersFromQB'")
									Else
									Response.Write("<input type='checkbox' class='check' id='chkImportCustomersFromQB' name='chkImportCustomersFromQB' checked")
									End If
									Response.Write("> Import customers from Quickbooks")
									%>
								</div>
							</div>
							<br>
							<div class="col-lg-12">
								When importing: Update with existing table or Replace all records in existing table<br>
							</div>
							<div class="col-lg-4"><br>
								<select class="form-control pull-left" name="selImportCustomersUpdateOrReplace">
									<option value="U" <% If ImportCustomersUpdateOrReplace = "M" Then Response.Write("selected") %>>Update</option>
									<option value="R" <% If ImportCustomersUpdateOrReplace = "R" Then Response.Write("selected") %>>Replace</option>
								</select>
							</div>
						</div>
					</div>	
				</div>
			</div>

		<!-- cancel / save !-->
		<div class="row pull-right">
			<div class="col-lg-12">
				<a href="<%= BaseURL %>admin/global/main.asp"><button type="button" class="btn btn-default">Cancel</button></a> 
				<button type="submit" class="btn btn-primary"><i class="far fa-save"></i> Save</button>
			</div>
		</div>
	
	<% End If %>
	</div><!-- row -->
</div><!-- container -->

</form>

<!--#include file="../../../inc/footer-main.asp"-->
