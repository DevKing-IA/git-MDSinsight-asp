<!--#include file="../../inc/header.asp"-->
<!--#include file="../../inc/jquery_table_search.asp"-->
<!--#include file="../../inc/InsightFuncs_BizIntel.asp"-->



 
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

<h1 class="page-header">Add / Edit MCS Reasons</h1>

<div class="row">
 	<div class="col-lg-12">
	 	<p>
 <a href="addMCSReason.asp">
    	<button type="button" class="btn btn-success">Add New MCS Reason</button>
	</a>
	 	</p>
	
 	</div>
</div>
<br>	
	<!-- tabs start here !-->
	<div class="container">

	<div class="row">
		<div class="col-lg-12">

			<div class="input-group narrow-results"> <span class="input-group-addon">Narrow Results</span>

    <input id="filter" type="text" class="form-control filter-search-width" placeholder="Type here...">
</div>
		
	<div class="table-responsive">
            <table    class="table table-striped table-condensed table-hover table-bordered sortable">
              <thead>
                <tr>
                  <th>Reason</th>
                  <th class="sorttable_nosort">Delete</th>
                </tr>
              </thead>
              <tbody class='searchable'>
              
				<%
			
				SQL = "SELECT * FROM BI_MCSReasons ORDER BY Reason"
		
				Set cnn8 = Server.CreateObject("ADODB.Connection")
				cnn8.open (Session("ClientCnnString"))
				Set rs = Server.CreateObject("ADODB.Recordset")
				rs.CursorLocation = 3 
				Set rs = cnn8.Execute(SQL)
		
				If not rs.EOF Then

					Do While Not rs.EOF
				
			        %>
						<!-- table line !-->
						<tr>
							<td><a href='editMCSReason.asp?i=<%= rs.Fields("InternalRecordIdentifier")%>'><%= rs.Fields("Reason")%></a></td>
							<%' Allow delete or display modal
							If NumberOfMCSActionsWithMCSReason(rs.Fields("InternalRecordIdentifier")) = 0 Then %>
								<td><a href='deleteMCSReasonQues.asp?i=<%=rs.Fields("InternalRecordIdentifier")%>'><i class="fas fa-trash-alt"></i></a></td>
							<%Else%>
								<td><a data-toggle="modal" data-show="true" href='deleteMCSReasonModal.asp?i=<%=rs.Fields("InternalRecordIdentifier")%>' data-target="#myModal"><i class="fas fa-trash-alt"></i></a></td> 
							<%End If%>
					   	</tr>
					<%
						rs.movenext
					loop
				End If
				set rs = Nothing
				cnn8.close
				set cnn8 = Nothing
	            %>
			</tbody>
		</table>
	</div>
 
		</div>

</div>
<!-- eof row !-->

<!-- modal  starts here !-->
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