<!--#include file="../../inc/header.asp"-->
<!--#include file="../../inc/jquery_table_search.asp"-->
<!--#include file="../../inc/InSightFuncs.asp"-->
<!--#include file="../../inc/InSightFuncs_Service.asp"-->

 <style type="text/css">
 	.email-table{
		width:46%;
	}
	
	table.sortable th:not(.sorttable_sorted):not(.sorttable_sorted_reverse):not(.sorttable_nosort):after { 
    content: " \25B4\25BE" 
}

.nav-tabs>li>a{
	background: #f5f5f5;
	border: 1px solid #ccc;
	color: #000;
}

.nav-tabs>li>a:hover{
	border: 1px solid #ccc;
}

.nav-tabs>li.active>a, .nav-tabs>li.active>a:focus, .nav-tabs>li.active>a:hover{
	color: #000;
	border: 1px solid #ccc;
}

.container{
	max-width:600px;
	margin:0 auto;
}

.narrow-results{
	margin:0px 0px 20px 0px;
}

#filter{
	width:40%;
}

.modal-link{
	cursor:pointer;
}

.modal-content{
	max-height:360px;
	overflow-y:auto;
}

 .modal-content .row{
	 padding-bottom:20px;
 }

 .modal-content p{
	 margin-bottom:20px;
	 white-space:normal;
 }
 </style>

<!--- eof on/off scripts !-->

<h1 class="page-header">Add / Edit Problem Codes</h1>

<div class="row">
 	<div class="col-lg-12">
	 	<p><a href="<%= BaseURL %>service/menu.asp"><button type="button" class="btn btn-primary"><i class="fas fa-arrow-left"></i>&nbsp;Back To Service Main Menu</button></a></p>
 	</div>
</div>
<br>	
	<!-- tabs start here !-->
	<div class="container">

	<div class="row">

		<div class="col-lg-12">
			 <div class="input-group narrow-results"> <span class="input-group-addon">Narrow Results</span>
    		 <input id="filter" type="text" class="form-control filter-search-width" placeholder="Type here...">
 			 <a href="addproblemCode.asp"><button type="button" class="btn btn-success pull-right">Add New Problem Code</button></a>
		</div>

		
		<div class="table-responsive">
            <table class="table table-striped table-condensed table-hover table-bordered sortable">
              <thead>
                <tr>
                  <th>Problem Description</th>
                  <th>Show On <br>Website</th>
                  <th class="sorttable_nosort">Delete</th>
                </tr>
              </thead>
               <tbody class='searchable'>
              
				<%
			
				SQL = "SELECT * FROM FS_ProblemCodes order by ProblemDescription"
		
				Set cnnproblemCodes = Server.CreateObject("ADODB.Connection")
				cnnproblemCodes.open (Session("ClientCnnString"))
				Set rsproblemCodes = Server.CreateObject("ADODB.Recordset")
				rsproblemCodes.CursorLocation = 3 
				Set rsproblemCodes = cnnproblemCodes.Execute(SQL)
		
				If not rsproblemCodes.EOF Then

					Do While Not rsproblemCodes.EOF
				
			        %>
						<!-- table line !-->
						<tr>
							<%If rsproblemCodes.Fields("InternalRecordIdentifier") = 0 Then %>
								<td><%= rsproblemCodes.Fields("ProblemDescription")%></td>
							<% Else %>
								<td><a href='editproblemCode.asp?i=<%= rsproblemCodes.Fields("InternalRecordIdentifier")%>'><%= rsproblemCodes.Fields("ProblemDescription")%></a></td>
							<% End If%>
							<td><%If rsproblemCodes("ShowOnWebsite") = 1 Then Response.Write("Yes") Else Response.Write("No")%></td>
							<%' Allow delete or display modal
							If rsproblemCodes.Fields("InternalRecordIdentifier") <> 0 Then 
								If NumberOfTicketsByProblemCode(rsproblemCodes("InternalRecordIdentifier")) = 0 Then %>
									<td><a href='deleteproblemCodeQues.asp?i=<%=rsproblemCodes.Fields("InternalRecordIdentifier")%>'><i class="fas fa-trash-alt"></i></a></td>
								<%Else%>
									<td><a data-toggle="modal" data-show="true" href='deleteproblemCodeModal.asp?i=<%=rsproblemCodes("InternalRecordIdentifier")%>' data-target="#myModal"><i class="fas fa-trash-alt"></i></a></td> 
								<%End If
							Else %>
								<td>&nbsp;</td>
							<% End If %>	
					   	</tr>
					<%
						rsproblemCodes.movenext
					loop
				End If
				set rsproblemCodes = Nothing
				cnnproblemCodes.close
				set cnnproblemCodes = Nothing
	            %>
			</tbody>
		</table>
	</div>

		</div>
 

</div>
<!-- eof row !--> 

 <!-- Modal -->
<div class="modal fade" id="myModal" tabindex="-1" role="dialog" aria-labelledby="myModalLabel" aria-hidden="true">
    <div class="modal-dialog">
        <div class="modal-content">
            
            <div class="modal-body"></div>
            
        </div>
        <!-- /.modal-content -->
    </div>
    <!-- /.modal-dialog -->
</div>
<!-- /.modal -->
<!-- modal  ends here !-->
   

<!--#include file="../../inc/footer-main.asp"-->