<!--#include file="../../../inc/header.asp"-->
<%
	Dim CatGroup_CategoryArray(22)
	Dim CatGroup_SortOrderArray(22)
	Dim CatGroup_GroupNameArray(22)
	Dim CatGroup_ShowOnGArray(22)
	Dim CatGroup_InternalID(22)
	
	'**********************************************************************
	SQL = "SELECT * FROM Settings_CatGroups"
	
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rs = cnn8.Execute(SQL)
	
	If rs.Eof Then 
		SQL = "Insert Into Settings_CatGroups (Category,GroupName) Select CategoryID, CategoryName from " &  MUV_Read("SQL_Owner")  & ".tblCategories order by CategoryID"
		Set rs = cnn8.Execute(SQL)	
		SQL ="Update Settings_CatGroups SET SortOrder=0,ShowOnGScreen=0"
		Set rs = cnn8.Execute(SQL)	
	End If

	'Now re-select in case we just did the first time insert
	SQL = "SELECT * FROM Settings_CatGroups order by Category"
	Set rs = cnn8.Execute(SQL)

	x = 0
	
	If not rs.EOF Then
		Do
			CatGroup_CategoryArray(x) = rs("Category")
			CatGroup_SortOrderArray(x) = rs("SortOrder")
			CatGroup_GroupNameArray(x) = rs("GroupName")
			CatGroup_ShowOnGArray(x) = rs("ShowOnGScreen")
			CatGroup_InternalID(x) = rs("GroupInternalIdentifier")
			x = x + 1
			rs.movenext
		Loop until rs.eof
	End If
	
	set rs = Nothing
	cnn8.close
	set cnn8 = Nothing
	
	
	'Also count the number of categories
	'we will use this later
	CatCount=0
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rs = cnn8.Execute("Select Count(*) as CCount from tblCategories")
	CatCount = rs("CCount")
	cnn8.close
	Set rs = Nothing
	Set cnn = Nothing
	

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

<h1 class="page-header"><i class="fa fa-globe"></i>&nbsp;Category Groupings For Period Sales Displays
	<a href="<%= BaseURL %>admin/global/main.asp"><button class="btn btn-small btn-secondary pull-right" style="margin-left:20px"><i class="fas fa-arrow-alt-left"></i>&nbsp;<i class="fas fa-globe"></i>&nbsp;GLOBAL SETTINGS MAIN</button></a>
</h1>


<form method="post" action="category-groupings-submit.asp" name="frmCategoryGroupings" id="frmCategoryGroupings">

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
			<div class="col-lg-6">
				<div class="table-responsive">
					<table class="table table-hover table-size">
					
					<!-- table header !-->
					<thead>
			        <tr>
			          <th class="category">&nbsp;</th>
			          <th class="group-name">Group Name</th>
			          <th class="sort-order">Sort Order</th>
			          <th class="display">Display On Period Sales Screen</th>
			        </tr>
			      </thead>
			      <!-- eof table header !-->
			      
			      <!-- table body !-->
			      <tbody>
        
        					<% For x = 0 to 10 %>
						        <!-- row !-->
						        <tr>
						          <th scope="row" class="category"><%= CatGroup_CategoryArray(x) %> - <%= GetCategoryByID(x) %></th>
						          <%
						          
   						           	ResponseLine = "<input type='hidden' class='form-control' value='" & x & "'"
						          	ResponseLine = ResponseLine & "id='txtCatID" & x & "' name ='txtCatID" & x & "'></td>"
						          	Response.write(ResponseLine)

						          	ResponseLine = "<td class='group-name'><input type='text' class='form-control input-sm' value='" & CatGroup_GroupNameArray(x) & "'"
						          	ResponseLine = ResponseLine & "id='txtGroupName" & x & "' name ='txtGroupName" & x & "'></td>"
						          	Response.write(ResponseLine)
						          	
						          	ResponseLine = "<td class='sort-order'><input type='text' class='form-control input-sm' value='" & CatGroup_SortOrderArray(x) & "'"
						          	ResponseLine = ResponseLine & "id='txtSortOrder" & x & "' name ='txtSortOrder" & x & "'></td>"
						          	Response.write(ResponseLine)
						          	

									ResponseLine = "<td  class='display'><input type='checkbox' "
									If CatGroup_ShowOnGArray(x)=vbTrue then 
									  	ResponseLine = ResponseLine & " checked "
							        End IF
						          	ResponseLine = ResponseLine & "id='chkGScreen" & x & "' name ='chkGScreen" & x & "'></td>"
						          	Response.write(ResponseLine)
							%>
						        </tr>
						        <!-- eof row !-->
					        <% Next %>
					       
		         </tbody>
		      <!-- eof table body !-->
						
					</table>
				</div>
			</div>
			<!-- eof col category 1 to 11 !-->	
			
			<!-- col category 12 to 22 !-->
			<div class="col-lg-6">
				
				<div class="table-responsive">
					<table class="table table-hover table-size">
						
						<!-- table header !-->
						<thead>
						<tr>
							<th class="category">&nbsp;</th>
							<th class="group-name">Group Name</th>
							<th class="sort-order">Sort Order</th>
							<th class="display">Display On Period Sales Screen</th>
						</tr>
						</thead>
						<!-- eof table header !-->
			      
			      <!-- table body !-->
			      <tbody>
        
    					<% For x = 11 to CatCount -1 'account for index of 0%>
					        <!-- row !-->
					        <tr>
					          <th scope="row" class="category"><%= CatGroup_CategoryArray(x) %> - <%= GetCategoryByID(x) %></th>
					          <%
					          
					           	ResponseLine = "<input type='hidden' class='form-control' value='" & x & "'"
					          	ResponseLine = ResponseLine & "id='txtCatID" & x & "' name ='txtCatID" & x & "'></td>"
					          	Response.write(ResponseLine)

					           	ResponseLine = "<td class='group-name'><input type='text' class='form-control input-sm' value='" & CatGroup_GroupNameArray(x) & "'"
					          	ResponseLine = ResponseLine & "id='txtGroupName" & x & "' name ='txtGroupName" & x & "'></td>"
					          	Response.write(ResponseLine)
					          	
					          	ResponseLine = "<td class='sort-order'><input type='text' class='form-control input-sm' value='" & CatGroup_SortOrderArray(x) & "'"
					          	ResponseLine = ResponseLine & "id='txtSortOrder" & x & "' name ='txtSortOrder" & x & "'></td>"
					          	Response.write(ResponseLine)
					          	

								ResponseLine = "<td  class='display'><input type='checkbox' "
								If CatGroup_ShowOnGArray(x)=vbTrue then 
								  	ResponseLine = ResponseLine & " checked "
						        End IF
					          	ResponseLine = ResponseLine & "id='chkGScreen" & x & "' name ='chkGScreen" & x & "'></td>"
					          	Response.write(ResponseLine)
						%>
					        </tr>
					        <!-- eof row !-->
				        <% Next %>
					        
	        
	         		</tbody>
	      			<!-- eof table body !-->
						
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
