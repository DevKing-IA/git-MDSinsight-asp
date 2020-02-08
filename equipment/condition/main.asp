<!--#include file="../../inc/header.asp"-->
<!--#include file="../../inc/jquery_table_search.asp"-->
<!--#include file="../../inc/InsightFuncs_Equipment.asp"-->


 
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
		max-width:1700px;
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

<h1 class="page-header">Add / Edit <%= GetTerm("Equipment") %> Conditions</h1>

<div class="row">
 	<div class="col-lg-12">
	 	<p>
 <a href="addCondition.asp">
    	<button type="button" class="btn btn-success">Add New Condition</button>
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
                  <th width="15%">Condition</th>
                  <th width="55%">Description</th>
                  <th>Pieces of Equipment</th>
                  <th>Customers with this Equipment</th>  
                  <th class="sorttable_nosort">Edit</th>                
                  <th class="sorttable_nosort">Delete</th>
                </tr>
              </thead>
              <tbody class='searchable'>
              
				<%
			
				SQL = "SELECT * FROM EQ_Condition ORDER BY Condition ASC"
		
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
							<td width="15%"><%= rs.Fields("Condition")%></td>
							<td width="55%"><%= rs.Fields("Description")%></td>
							<td align="center"><a href='viewAllEquipmentByCondition.asp?i=<%= rs.Fields("InternalRecordIdentifier")%>'><%= NumberEquipmentRecsDefinedForCondition(rs.Fields("InternalRecordIdentifier")) %></a></td>
							<td align="center"><a href='viewCustomerEquipmentByCondition.asp?i=<%= rs.Fields("InternalRecordIdentifier")%>'><%= NumberCustomerEquipmentRecsDefinedForCondition(rs.Fields("InternalRecordIdentifier")) %></a></td>
							<td align="center"><a href='editCondition.asp?i=<%= rs.Fields("InternalRecordIdentifier")%>'><i class="fa fa-pencil"></i></a></td>

							<%' Allow delete or display modal
							If NumberEquipmentRecsDefinedForCondition(rs.Fields("InternalRecordIdentifier")) = 0 Then %>
								<td align="center"><a href='deleteConditionQues.asp?i=<%=rs.Fields("InternalRecordIdentifier")%>'><i class="fas fa-trash-alt"></i></a></td>
							<%Else%>
								<td align="center"><a data-toggle="modal" data-show="true" href='deleteConditionModal.asp?i=<%=rs.Fields("InternalRecordIdentifier")%>' data-target="#myModal"><i class="fas fa-trash-alt"></i></a></td> 
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