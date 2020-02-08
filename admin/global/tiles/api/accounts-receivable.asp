<!--#include file="../../../../inc/header.asp"-->

<style>


	.content-element{
	  margin:50px 0 0 50px;
	}
	.circles-list ol {
	  list-style-type: none;
	  margin-left: 1.25em;
	  padding-left: 2.5em;
	  counter-reset: li-counter;
	  border-left: 1px solid #3c763d;
	  position: relative; }
	
	.circles-list ol > li {
	  position: relative;
	  margin-bottom: 3.125em;
	  clear: both; }
	
	.circles-list ol > li:before {
	  position: absolute;
	  top: -0.5em;
	  font-family: "Open Sans", sans-serif;
	  font-weight: 600;
	  font-size: 1em;
	  left: -3.75em;
	  width: 2.25em;
	  height: 2.25em;
	  line-height: 2.25em;
	  text-align: center;
	  z-index: 9;
	  color: #3c763d;
	  border: 2px solid #3c763d;
	  border-radius: 50%;
	  content: counter(li-counter);
	  background-color: #DFF0D8;
	  counter-increment: li-counter; }
	  	
	.row .panel-row{
	    margin-top:40px;
	    padding: 0 10px;
	}
	
	.clickable{
	    cursor: pointer;   
	}
	
	.panel-heading span {
		margin-top: -20px;
		font-size: 15px;
	}

	.container {
		margin-bottom: 20px;
		margin-top: 20px;
		margin-left:0px;
		width: 100%;
	}

	.container .row {
		margin-bottom: 20px;
		/*margin-top: 20px;*/
	}

	.line-full{
	 	margin-bottom:20px;
	}
	
	.multi-select{
		min-height:200px;
		min-width:180px;
	}
	
	.custom-select{
		width: auto !important;
		display:inline-block;
	}

	
	.select-large{
		min-width:40% !important;
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

	$(document).ready(function() {
							
		
		$('.panel .panel-body').css('display','none');
		$('.panel-heading span.clickable').addClass('panel-collapsed');
		$('.panel-heading span.clickable').find('i').removeClass('glyphicon-chevron-up').addClass('glyphicon-chevron-down');

		$(document).on('click', '.panel-heading span.clickable', function(e){
		    var $this = $(this);
			if(!$this.hasClass('panel-collapsed')) {
				$this.parents('.panel').find('.panel-body').slideUp();
				$this.addClass('panel-collapsed');
				$this.find('i').removeClass('glyphicon-chevron-up').addClass('glyphicon-chevron-down');
			} else {
				$this.parents('.panel').find('.panel-body').slideDown();
				$this.removeClass('panel-collapsed');
				$this.find('i').removeClass('glyphicon-chevron-down').addClass('glyphicon-chevron-up');
			}
		});
		
		
 		$("#toggle").click(function(){
 		
            if(!$('.panel-heading span.clickable').hasClass('panel-collapsed')) {
				$('.panel .panel-body').css('display','none');
				$('.panel-heading span.clickable').addClass('panel-collapsed');
				$('.panel-heading span.clickable').find('i').removeClass('glyphicon-chevron-up').addClass('glyphicon-chevron-down');
            }
            else {
				$('.panel .panel-body').css('display','block');
				$('.panel-heading span.clickable').removeClass('panel-collapsed');
				$('.panel-heading span.clickable').find('i').removeClass('glyphicon-chevron-down').addClass('glyphicon-chevron-up');
            }
        });	
		
	});
</script>


<%

	SQL = "SELECT * FROM Settings_AR"
	
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rs = cnn8.Execute(SQL)
	
	If not rs.EOF Then
		POST_Serno = rs("POST_Serno")
		POST_Mode = rs("POST_Mode")
		EmailForNon200Responses = rs("EmailForNon200Responses")
		POST_CustomerURL1 = rs("POST_CustomerURL1")
		POST_CustomerURL2 = rs("POST_CustomerURL2")	
		POST_CustomerURL1ONOFF = rs("POST_CustomerURL1ONOFF")		
		POST_CustomerURL2ONOFF = rs("POST_CustomerURL2ONOFF")
	End If
	
	set rs = Nothing
	cnn8.close
	set cnn8 = Nothing

%>

<h1 class="page-header"><i class="fa fa-globe"></i>&nbsp;<%= GetTerm("Accounts Receivable") %> POST Settings 
	<button id="toggle" class="btn btn-small btn-success"><i class="fas fa-arrows-v"></i>&nbsp;EXPAND/COLLAPSE ALL SETTINGS</button>
	<a href="<%= BaseURL %>admin/global/main.asp"><button class="btn btn-small btn-secondary pull-right" style="margin-left:20px"><i class="fas fa-arrow-alt-left"></i>&nbsp;<i class="fas fa-globe"></i>&nbsp;GLOBAL SETTINGS MAIN</button></a>
	<a href="<%= BaseURL %>admin/global/tiles/api/main.asp"><button class="btn btn-small btn-secondary pull-right"><i class="fas fa-arrow-alt-left"></i>&nbsp;<i class="fa fa-external-link"></i>&nbsp;API MAIN</button></a>
</h1>

<form method="post" action="accounts-receivable-submit.asp" name="frmPostSettings" id="frmPostSettings">

	<div class="container">
	
		<%
			Response.Write("<div id='PleaseWaitPanel' style='display:none;'>")
			Response.Write("<br><br>Saving your recent changes, please wait...<br><br>")
			Response.Write("<img src=""" & baseURL & "/img/loading.gif"" />")
			Response.Write("</div>")
			Response.Flush()
		%>
	
		<div class="row">
			<h3><i class="fad fa-sliders-h"></i>&nbsp;<%= GetTerm("Accounts Receivable") %> API Master Settings</h3>
		</div>
	
	    <div class="row">
	    
			<div class="col-md-4">
				<div class="panel panel-primary">
					<div class="panel-heading">
						<h3 class="panel-title">Post Settings</h3>
						<span class="pull-right clickable panel-collapsed"><i class="glyphicon glyphicon-chevron-down"></i></span> 
					</div>
					<div class="panel-body">
					
			         	<div class="row">
				         	<!-- label !-->
				         	<div class="col-lg-2">
					         	 <label for="serialnumber" class="post-labels">SerNo</label>
				         	</div>
				         	<!-- eof label !-->
				         	
						    <!-- input !-->
						    <div class="col-lg-9">
						    <input type="text" class="form-control" name="txtSerno" id="txtSerno" value="<%= POST_Serno %>">
						    </div>
						    <!-- eof input !-->
			         	</div>
						
						<div class="row">
							<!-- label !-->
							<div class="col-lg-2">
								<label  for="mode" class="post-labels">Mode</label>
							</div>
							<!-- eof label !-->
						
							<!-- input !-->
							<div class="col-lg-5">
								<select class="form-control" name="selMode" id="txtMode">
									<option <% If POST_Mode="TEST" Then Response.Write(" selected ") %> >TEST</option>
									<option <% If POST_Mode="LIVE" Then Response.Write(" selected ") %> >LIVE</option>
								</select>
							</div>
							<!-- eof input !-->         	
						</div>
						
						
						<div class="row">
							<div class="col-lg-11">
								<p><strong>If Response <> 200 Email:</strong></p>
								<input type="text" class="form-control" name="txtEmailForNon200Responses" id="txtEmailForNon200Responses" value="<%=EmailForNon200Responses%>">
							</div>
						</div>
						<!-- eof row with data !-->
			         	
						<div class="row">

						    <div class="col-lg-10">
						    	<p><strong>Customer POST URL #1&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</strong>			    	
									<%
										If POST_CustomerURL1ONOFF = 0 Then
											Response.Write("<input type='checkbox' class='check' id='chkPOST_CustomerURL1ONOFF' name='chkPOST_CustomerURL1ONOFF'")
										Else
											Response.Write("<input type='checkbox' class='check' id='chkPOST_CustomerURL1ONOFF' name='chkPOST_CustomerURL1ONOFF' checked")
										End If
										Response.Write(">&nbsp;Turn On")
									%>
						    	</p>
							    <input type="text" class="form-control" name="txtCustomerURL1" id="txtCustomerURL1" value="<%= POST_CustomerURL1 %>">
						    </div>
						    
							<div class="col-lg-10">
							</div>
						</div>
						<!-- eof row with data !-->
			         	
					
						<div class="row">
						
						    <!-- input !-->
						    <div class="col-lg-10">
						    	<p><strong>Customer POST URL #2&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</strong> 
							    	<%
										If POST_CustomerURL2ONOFF = 0 Then
											Response.Write("<input type='checkbox' class='check' id='chkPOST_CustomerURL2ONOFF' name='chkPOST_CustomerURL2ONOFF'")
										Else
											Response.Write("<input type='checkbox' class='check' id='chkPOST_CustomerURL2ONOFF' name='chkPOST_CustomerURL2ONOFF' checked")
										End If
										Response.Write(">&nbsp;Turn On")
									%>
						    	</p>
						    	<input type="text" class="form-control" name="txtCustomerURL2" id="txtCustomerURL2" value="<%= POST_CustomerURL2 %>">
						    </div>
						    <!-- eof input !-->
						    
						</div>
						<!-- eof row with data !-->
					
					</div>
				</div>
			</div>
		</div>
		
			
		<!-- cancel / save !-->
		<div class="row pull-right">
			<div class="col-lg-12">
				<a href="<%= BaseURL %>admin/global/tiles/api/main.asp"><button type="button" class="btn btn-default btn-lg btn-huge"><i class="far fa-times-circle"></i> Cancel</button></a> 
				<button type="submit" class="btn btn-primary btn-lg btn-huge" onclick="showSavingChangesDiv()"><i class="far fa-save"></i> Save Changes</button>
			</div>
		</div>
	
	
</div><!-- container -->

</form>

<!--#include file="../../../../inc/footer-main.asp"-->
