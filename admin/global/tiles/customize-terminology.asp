<!--#include file="../../../inc/header.asp"-->

<%
	
	'**********************************************************
	'Now fillup the Terminology vars
	'**********************************************************
	'First find out how many
	'**********************************************************
	
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rs = cnn8.Execute("SELECT COUNT(*) AS TCount FROM SC_Terminology")
	Tcount = rs("TCount")
	ReDim TermArray(TCount)
	cnn8.close
	
	'**********************************************************
	'Obtain Terms From the SC_Terminology Table
	'**********************************************************

	SQL = "SELECT * FROM SC_Terminology ORDER BY GenericTerm"
	
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rs = cnn8.Execute(SQL)
	
	If not rs.EOF Then
		x=1
		Do
			TermArray(x) = rs("CustomTerm")
			rs.MoveNext
			x=x+1
		Loop While not rs.eof
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
	
 	.table-size .category{
	 	width: 35%;
	 	font-weight: normal;
 	}
	
 	.table-size .group-name{
	 	width: 40%
 	}

 	.table-size .term-name{
	 	width: 60%
 	}
 
 	.table-size .term-title{
    width: 28%;
    font-weight: 100;
    font-size: 10pt;
 	}

 	.table-size .term-category{
	 	width: 28%;
 	}
 	
 	
 	.table-size .sort-order{
	 	width: 10%;
 	}
 	
 	.table-size .display{
	 	width: 15%;
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

<h1 class="page-header"><i class="fa fa-globe"></i>&nbsp;Customize Terminology
	<a href="<%= BaseURL %>admin/global/main.asp"><button class="btn btn-small btn-secondary pull-right" style="margin-left:20px"><i class="fas fa-arrow-alt-left"></i>&nbsp;<i class="fas fa-globe"></i>&nbsp;GLOBAL SETTINGS MAIN</button></a>
</h1>


<form method="post" action="customize-terminology-submit.asp" name="frmCustomizeTerminology" id="frmCustomizeTerminology">

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
						
			<!-- col category 1 to 11 !-->
			<div class="col-lg-4">
				<div class="table-responsive">
					<table class="table table-hover table-size">
						<!-- table header !-->
						<thead>
				        <tr>
				          <th class="term-title">Module</th>
				          <th class="term-title">General Term</th>
				          <th class="term-name">Replace With</th>
				        </tr>
				      </thead>
				      <!-- eof table header !-->
				      
				      <tbody>
				      <%'This is where we get all the Terms
					     				
	     				SQL = "SELECT * from SC_Terminology order by GenericTerm"
						
						Set cnn8 = Server.CreateObject("ADODB.Connection")
						cnn8.open (Session("ClientCnnString"))
						Set rs = Server.CreateObject("ADODB.Recordset")
						rs.CursorLocation = 3 
						Set rs = cnn8.Execute("Select Count(*) as TCount from SC_Terminology")
						NumberOfTerms = rs("TCount")
						Set rs = cnn8.Execute(SQL)
										
							If not rs.EOF Then
								
								NumberEachCol = NumberOfTerms/3
								If NumberOfTerms mod 3 <> 0 Then 
									NumberEachCol = int(NumberEachCol)
									NumberEachCol = NumberEachCol + 1
								End If
								' ****** Left Side
								For x = 1 to NumberEachCol
						        	%><tr>
						        		<th scope="row" class="term-title"><%= rs("Module")%></th>
						          		<th scope="row" class="term-title"><%= rs("GenericTerm")%></th>
						          		<%
						          
					          		 	ResponseLine = "<td class='term-name'><input type='text' class='form-control input-sm' value='" & TermArray(x) & "'"
					          			ResponseLine = ResponseLine & "id='txtTerm" & x & "' name ='txtTerm" & x & "'></td>"
					          			Response.write(ResponseLine)
								
										rs.movenext
									%></tr><%
								Next
							End If%>
						</tbody>
					</table>
				</div>
			</div>
			<!-- eof col category 1 to 11 !-->	
			
			<% If not rs.EOF Then %>
			<!-- col category 12 to 22 !-->
			<div class="col-lg-4">
				
				<div class="table-responsive">
					<table class="table table-hover table-size">
					
					<!-- table header !-->
					<thead>
			        <tr>
			        	  <th class="term-title">Module</th>
				          <th class="term-title">General Term</th>
				          <th class="term-name">Replace With</th>
			        </tr>
			      </thead>
			      <!-- eof table header !-->
								<%
									EndVal = NumberEachCol*2
									For x = (NumberEachCol+1) to EndVal
						        	%><tr>
						        		<th scope="row" class="term-title"><%= rs("Module")%></th>
						          		<th scope="row" class="term-title"><%= rs("GenericTerm")%></th>
						          		<%
						          
					          		 	ResponseLine = "<td class='term-name'><input type='text' class='form-control input-sm' value='" & TermArray(x) & "'"
					          			ResponseLine = ResponseLine & "id='txtTerm" & x & "' name ='txtTerm" & x & "'></td>"
					          			Response.write(ResponseLine)
								
										rs.movenext
									%></tr><%	
								Next

								%>
						
					</table>
				</div>
				
			</div>
			<!-- eof col category 12 to 22 !-->	
			<% End If%>
			
			<!-- col category 23 to 22 !-->
			<div class="col-lg-4">
				
				<div class="table-responsive">
					<table class="table table-hover table-size">
					
					<!-- table header !-->
					<thead>
			        <tr>
			        	  <th class="term-title">Module</th>
				          <th class="term-title">General Term</th>
				          <th class="term-name">Replace With</th>
			        </tr>
			      </thead>
			      <!-- eof table header !-->
								<%
									EndVal = NumberOfTerms
									For x = ((NumberEachCol*2)+1) to EndVal
						        	%><tr>
						        		<th scope="row" class="term-title"><%= rs("Module")%></th>
						          		<th scope="row" class="term-title"><%= rs("GenericTerm")%></th>
						          		<%
						          
					          		 	ResponseLine = "<td class='term-name'><input type='text' class='form-control input-sm' value='" & TermArray(x) & "'"
					          			ResponseLine = ResponseLine & "id='txtTerm" & x & "' name ='txtTerm" & x & "'></td>"
					          			Response.write(ResponseLine)
								
										rs.movenext
									%></tr><%	
								Next

								cnn8.Close
								Set rs = Nothing
								Set cnn8 = Nothing

								
								%>
						
					</table>
				</div>
				
			</div>
			<!-- eof col category 12 to 22 !-->	
			
			
		</div>
		<!-- eof three cols !-->
			
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
