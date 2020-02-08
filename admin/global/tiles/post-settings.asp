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
	
	.post-labels{
 		padding-top: 5px;
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
		POST_Serno = rs("POST_Serno")
		POST_Mode = rs("POST_Mode")
		EmailForNon200Responses = rs("EmailForNon200Responses")
		POST_ServiceMemoURL1 = rs("POST_ServiceMemoURL1")
		POST_AssetLocationURL1 = rs("POST_AssetLocationURL1")
		POST_ServiceMemoURL2 = rs("POST_ServiceMemoURL2")
		POST_AssetLocationURL2 = rs("POST_AssetLocationURL2")
		InternalEmail_MailDomain = rs("InternalEmail_MailDomain")
		ShowOpenPopupMessage =	rs("NotesScreenShowPopup")
		NeverPutOnHold = rs("NeverPutOnHold")
		POST_ServiceMemoURL1ONOFF = rs("POST_ServiceMemoURL1ONOFF")		
		POST_ServiceMemoURL1_MplexFormat = rs("POST_ServiceMemoURL1_MplexFormat")			
		POST_AssetLocationURL1ONOFF = rs("POST_AssetLocationURL1ONOFF")		
		POST_ServiceMemoURL2ONOFF = rs("POST_ServiceMemoURL2ONOFF")		
		POST_AssetLocationURL2ONOFF = rs("POST_AssetLocationURL2ONOFF")		
	End If
	
	set rs = Nothing
	cnn8.close
	set cnn8 = Nothing

%>

<h1 class="page-header"><i class="fa fa-globe"></i>&nbsp;POST Settings
	<a href="<%= BaseURL %>admin/global/main.asp"><button class="btn btn-small btn-secondary pull-right" style="margin-left:20px"><i class="fas fa-arrow-alt-left"></i>&nbsp;<i class="fas fa-globe"></i>&nbsp;GLOBAL SETTINGS MAIN</button></a>
</h1>


<form method="post" action="post-settings-submit.asp" name="frmPostSettings" id="frmPostSettings">

<div class="container">

    	
	<%
		Response.Write("<div id='PleaseWaitPanel' style='display:none;'>")
		Response.Write("<br><br>Saving your recent changes, please wait...<br><br>")
		Response.Write("<img src=""" & baseURL & "/img/loading.gif"" />")
		Response.Write("</div>")
		Response.Flush()
	%>
 
     <!-- row with data !-->
	         	<div class="row row-data ">
		         	
		         	
		         	<!-- content col !-->
		         	<div class="col-lg-12">
			         	
			         	<div class="row">
		         	<!-- serial no !-->
		         	<div class="col-lg-3">
			         	<div class="row">
				         	
			         	<!-- label !-->
			         	<div class="col-lg-2">
				         	 <label  for="serialnumber" class="post-labels">SerNo</label>
			         	</div>
			         	<!-- eof label !-->
			         	
					    <!-- input !-->
					    <div class="col-lg-9">
					    <input type="text" class="form-control" name="txtSerno" id="txtSerno" value="<%= POST_Serno %>">
					    </div>
					    <!-- eof input !-->
    
 			         	</div>
		         	</div>
		         	<!-- eof serial no !-->
		         	
		         	<!-- mode !-->
		         	<div class="col-lg-3">
			         	<div class="row">
				         	
				         	<!-- label !-->
			         	<div class="col-lg-2">
				         	 <label  for="mode" class="post-labels">Mode</label>
			         	</div>
			         	<!-- eof label !-->
			         	
							 <!-- input !-->
				    <div class="col-lg-9">
				    <select class="form-control" name="selMode" id="txtMode">
					  <option <% If POST_Mode="TEST" Then Response.Write(" selected ") %> >TEST</option>
					  <option <% If POST_Mode="LIVE" Then Response.Write(" selected ") %> >LIVE</option>
					</select>
				    </div>
				    <!-- eof input !-->         	
     
		         	</div>

 		         	</div>
		         	<!-- eof mode !-->
		         	
		         	<div class="col-lg-2">
						<label  for="servicememoposturl" class="post-labels">If Response <> 200 Email:</label>
					</div>
					<div class="col-lg-4">
						<input type="text" class="form-control" name="txtEmailForNon200Responses" id="txtEmailForNon200Responses" value="<%=EmailForNon200Responses%>">
					</div>

		         	
	         	</div>
   <!-- eof row with data !-->
  
  <!-- row with data !-->
	         	<div class="row row-data ">
		         	
		         	<!-- line !-->
		         	<div class="col-lg-12">
			         	<div class="row">
				         	
  
			         	<!-- Service Memo POST URL #1!-->
						<div class="col-lg-2">
							<label  for="servicememoposturl" class="post-labels">Service Memo POST URL #1</label>
						</div>
						
						<div class="col-lg-4">
							<input type="text" class="form-control" name="txtServiceMemoURL1" id="txtServiceMemoURL1" value="<%= POST_ServiceMemoURL1 %>">
						</div>
						
						<div class="col-lg-4">
						<%
							If POST_ServiceMemoURL1ONOFF = 0 Then
								Response.Write("<input type='checkbox' class='check' id='chkPOST_ServiceMemoURL1ONOFF' name='chkPOST_ServiceMemoURL1ONOFF'")
							Else
								Response.Write("<input type='checkbox' class='check' id='chkPOST_ServiceMemoURL1ONOFF' name='chkPOST_ServiceMemoURL1ONOFF' checked")
							End If
							Response.Write(">&nbsp;Turn On")
						%>
						</div>


						<div class="col-lg-4">
						<%
							If POST_ServiceMemoURL1_MplexFormat = 0 Then
								Response.Write("<input type='checkbox' class='check' id='chkPOST_ServiceMemoURL1_MplexFormat' name='chkPOST_ServiceMemoURL1_MplexFormat'")
							Else
								Response.Write("<input type='checkbox' class='check' id='chkPOST_ServiceMemoURL1_MplexFormat' name='chkPOST_ServiceMemoURL1_MplexFormat' checked")
							End If
							Response.Write("> Use Metroplex Format")
						%>
						</div>

					    <!-- eof Service Memo POST URL #1 !-->
     
   		         	   
  			         	</div>
		         	</div>
		         	<!-- eof line !-->
		         	
		         	
		         	<!-- line !-->
		         	<div class="col-lg-12">
			         	<div class="row">
				         	
  
				         	<!-- Asset Location POST URL #1!-->
				         	<div class="col-lg-2">
					         	 <label  for="assetlocationposturl" class="post-labels">Asset Location POST URL #1</label>
				         	</div>
				         	
						    <div class="col-lg-4">
							    <input type="text" class="form-control" name="txtAssetLocationURL1" id="txtAssetLocationURL1" value="<%= POST_AssetLocationURL1 %>">
						    </div>
						    
	     					<div class="col-lg-4">
							<%
								If POST_AssetLocationURL1ONOFF = 0 Then
									Response.Write("<input type='checkbox' class='check' id='chkPOST_AssetLocationURL1ONOFF' name='chkPOST_AssetLocationURL1ONOFF'")
								Else
									Response.Write("<input type='checkbox' class='check' id='chkPOST_AssetLocationURL1ONOFF' name='chkPOST_AssetLocationURL1ONOFF' checked")
								End If
								Response.Write(">&nbsp;Turn On")
							%>
							</div>

   		         	   
  			         	</div>
		         	</div>
		         	<!-- eof Asset Location POST URL #1!-->




		         	<!-- line !-->
		         	<div class="col-lg-12">
			         	<div class="row">
				         	
  
			         	<!-- label !-->
			         	<div class="col-lg-2">
				         	 <label  for="servicememoposturl" class="post-labels">Service Memo POST URL #2</label>
			         	</div>
			         	<!-- eof label !-->
			         	
					    <!-- input !-->
					    <div class="col-lg-4">
					    <input type="text" class="form-control" name="txtServiceMemoURL2" id="txtServiceMemoURL2" value="<%= POST_ServiceMemoURL2 %>">
					    </div>
					    <!-- eof input !-->
     
   		         	   	<div class="col-lg-4">
							<%
								If POST_ServiceMemoURL2ONOFF = 0 Then
									Response.Write("<input type='checkbox' class='check' id='chkPOST_ServiceMemoURL2ONOFF' name='chkPOST_ServiceMemoURL2ONOFF'")
								Else
									Response.Write("<input type='checkbox' class='check' id='chkPOST_ServiceMemoURL2ONOFF' name='chkPOST_ServiceMemoURL2ONOFF' checked")
								End If
								Response.Write(">&nbsp;Turn On")
							%>
						</div>

  			         	</div>
		         	</div>
		         	<!-- eof line !-->
		         	
		         	
		         	<!-- line !-->
		         	<div class="col-lg-12">
			         	<div class="row">
				         	
  
			         	<!-- label !-->
			         	<div class="col-lg-2">
				         	 <label  for="assetlocationposturl" class="post-labels">Asset Location POST URL #2</label>
			         	</div>
			         	<!-- eof label !-->
			         	
					    <!-- input !-->
					    <div class="col-lg-4">
					    <input type="text" class="form-control" name="txtAssetLocationURL2" id="txtAssetLocationURL2" value="<%= POST_AssetLocationURL2 %>">
					    </div>
					    <!-- eof input !-->
					    
					    
        		       	<div class="col-lg-4">
							<%
								If POST_AssetLocationURL2ONOFF = 0 Then
									Response.Write("<input type='checkbox' class='check' id='chkPOST_AssetLocationURL2ONOFF' name='chkPOST_AssetLocationURL2ONOFF'")
								Else
									Response.Write("<input type='checkbox' class='check' id='chkPOST_AssetLocationURL2ONOFF' name='chkPOST_AssetLocationURL2ONOFF' checked")
								End If
								Response.Write(">&nbsp;Turn On")
							%>
						</div>

   		         	   
  			         	</div>
		         	</div>
		         	<!-- eof line !-->
		         	
		         	
		         	<!-- line !-->
		         	<div class="col-lg-12">
			         	<div class="row">
				         	
  
			         	<!-- label !-->
			         	<div class="col-lg-2">
				         	 <label  for="internalemailmaildomain" class="post-labels">Mail domain for internal email</label>
			         	</div>
			         	<!-- eof label !-->
			         	
					    <!-- input !-->
					    <div class="col-lg-4">
					    <input type="text" class="form-control" name="txtInternalEmail_MailDomain" id="txtInternalEmail_MailDomain" value="<%= InternalEmail_MailDomain %>">
					    </div>
					    <!-- eof input !-->
					    
 

        		       	<div class="col-lg-4">
						&nbsp;
						</div>

   		         	   
  			         	</div>
		         	</div>
		         	<!-- eof line !-->

		         	<!-- Never put on hold !-->
		         	<div class="col-lg-12">
			         	<div class="row">
  			         	<!-- label !-->
			         	<div class="col-lg-2">
				         	 <label  for="internalemailmaildomain" class="post-labels">Never put tickets on HOLD</label>
			         	</div>
			         	<!-- eof label !-->
        		       	<div class="col-lg-4">
						<%
							If NeverPutOnHold = 0 Then
								Response.Write("<input type='checkbox' class='check' id='chkNeverPutOnHold' name='chkNeverPutOnHold'")
							Else
								Response.Write("<input type='checkbox' class='check' id='chkNeverPutOnHold' name='chkNeverPutOnHold' checked")
							End If
							Response.Write(">")
						%>
						</div>
        		       	<div class="col-lg-4">
						&nbsp;
						</div>
  			         	</div>
		         	</div>
		         	<!-- eof Never put on hold !-->

		         	</div>
   <!-- eof row with data !-->
   
		         	</div>
	         	</div>
   <!-- eof content col !-->
     
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
