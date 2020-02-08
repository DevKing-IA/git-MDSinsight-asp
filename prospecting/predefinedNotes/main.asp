<!--#include file="../../inc/header.asp"-->
<!--#include file="../../inc/jquery_table_search.asp"-->



 
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

<h1 class="page-header">Add / Edit Predefined Notes</h1>

<div class="row">
 	<div class="col-lg-12">
	 	<p>
 <a href="addPredefinedNote.asp">
    	<button type="button" class="btn btn-success">Add New Predefined Note</button>
	</a>
    <a class="btn btn-primary" href="../menu.asp" role="button"><i class="fa fa-arrow-left"></i> &nbsp;Back To Menu</a>
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
                  <th>Predefined Note</th>
                  <th class="sorttable_nosort">Delete</th>
                </tr>
              </thead>
              <tbody class='searchable'>
              
				<%
			
				SQL = "SELECT * FROM PR_PredefinedNotes order by PredefinedNote"
		
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
							<!-- <td><a href='editPredefinedNote.asp?i=<%= rs.Fields("InternalRecordIdentifier")%>'><%= rs.Fields("PredefinedNote")%></a></td> !-->
							
							<td>
								<!-- link for modal !-->
									<a  data-toggle="modal" data-target="#myModal" class="modal-link"><%= rs.Fields("PredefinedNote")%></a>
								<!-- eof link for modal !-->
 							</td>
							
							<td><a href='deletePredefinedNoteQues.asp?i=<%=rs.Fields("InternalRecordIdentifier")%>'><i class="fas fa-trash-alt"></i></a></td>
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

<!-- modal starts here !-->
<div class="modal fade" id="myModal" tabindex="-1" role="dialog" aria-labelledby="myModalLabel">
  <div class="modal-dialog" role="document">
    <div class="modal-content">
      <div class="modal-header">
        <button type="button" class="close" data-dismiss="modal" aria-label="Close"><span aria-hidden="true">&times;</span></button>
        <h4 class="modal-title" id="myModalLabel">Your Modal title</h4>
      </div>

	  <!-- modal content !-->
      <div class="modal-body">

		  <!-- label with input !-->
		  <div class="row">
 			  <div class="form-group">
    <label class="col-sm-3 control-label">Label</label>
    <div class="col-sm-9">
      <input type="text" class="form-control">
    </div>
 			  </div>
 			  </div>
 		  <!-- eof label with input !-->

		   <!-- text !-->
		  <p>
 			   Suspendisse non elit ex. In vel convallis felis. Nam tincidunt auctor fringilla. Nulla in odio ultrices, venenatis nisi eget, luctus ipsum. Ut vel tempor nibh. Duis ac euismod magna. Fusce cursus mattis massa, quis facilisis nisl. Praesent non diam nisl.
		  </p>
		   <!-- eof text !-->

		    <!-- label with dropbox !-->
			<div class="row">
 			  <div class="form-group">
    <label class="col-sm-3 control-label">Label</label>
    <div class="col-sm-9">
      <select class="form-control">
		  <option>Drop Down</option>
		  <option>Drop Down</option>
		  <option>Drop Down</option>
		  <option>Drop Down</option>
		  <option>Drop Down</option>
		  <option>Drop Down</option>
      </select>
    </div>
    </div>
 		  <!-- eof label with dropbox !-->
        
      </div>
	  <!-- eof modal content !-->
      <div class="modal-footer">
        <button type="button" class="btn btn-default" data-dismiss="modal">Close</button>
        <button type="button" class="btn btn-primary">Save changes</button>
      </div>
    </div>
  </div>
</div>
								<!-- modal ends here !-->
								

<!--#include file="../../inc/footer-main.asp"-->