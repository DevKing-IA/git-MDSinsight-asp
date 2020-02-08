<!--#include file="../../inc/header.asp"-->
<!--#include file="../../inc/jquery_table_search.asp"-->
<!--#include file="../../inc/InSightFuncs.asp"-->
<!--#include file="../../inc/InSightFuncs_Service.asp"-->
<!--#include file="../../inc/InsightFuncs_AR_AP.asp"-->

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

<h1 class="page-header">Add / Edit Terms</h1>

<div class="row">
 	<div class="col-lg-12">
	 	<p>
 <a href="addTerm.asp">
    	<button type="button" class="btn btn-success">Add New Term</button>
	</a>
	 	</p>
	
 	</div>
</div>
<br>	
	<!-- tabs start here !-->
	<div class="container">

	<div class="row">
		<div class="col-lg-12">

			 <div class="input-group narrow-results"> <span class="input-group-addon">Search</span>

    <input id="filter" type="text" class="form-control filter-search-width" placeholder="Type here...">
</div>
		
	<div class="table-responsive">
            <table    class="table table-striped table-condensed table-hover table-bordered sortable">
              <thead>
                <tr>                  
                  <th>Term Description</th>
				  <th>First Terms Percent</th>
				  <th>First Terms Period</th>
				  <th>Second Terms Period</th>
				  <th>Terms Type</th>
				  <th>Credit Card Bill</th>				  
                  <th class="sorttable_nosort">Delete</th>
                </tr>
              </thead>
               <tbody class='searchable'>
              
				<%
			
				SQL = "SELECT * FROM AR_Terms order by InternalRecordIdentifier"
		
				Set cnncust = Server.CreateObject("ADODB.Connection")
				cnncust.open (Session("ClientCnnString"))
				Set rscust = Server.CreateObject("ADODB.Recordset")
				rscust.CursorLocation = 3 
				Set rscust = cnncust.Execute(SQL)
		
				If not rscust.EOF Then

					Do While Not rscust.EOF
						TermCode = rscust.Fields("InternalRecordIdentifier")
			        %>
						<!-- table line !-->
						<tr>
							<%If rscust.Fields("InternalRecordIdentifier") = 0 Then %>
								<td><%= rscust.Fields("Description")%></td>
							<% Else %>
								<td><a href='editTerm.asp?i=<%= rscust.Fields("InternalRecordIdentifier")%>'><%= rscust.Fields("Description")%></a></td>
							<% End If%>							

							<td><%= rscust.Fields("firstTermsPercent")%></td>
							<td><%= rscust.Fields("firstTermsPeriod")%></td>
							<td><%= rscust.Fields("secondTermsPeriod")%></td>
							<td><%= rscust.Fields("TermsType")%></td>
							<td><%= rscust.Fields("CreditCardBill")%></td>
							
							<%' Allow delete or display modal
							If rscust.Fields("InternalRecordIdentifier") <> 0 Then 								
								If NumberOfCustomersWithTerm(TermCode) = 0 Then %>
									<td><a href='deleteTermQues.asp?i=<%=rscust.Fields("InternalRecordIdentifier")%>'><i class="fas fa-trash-alt"></i></a></td>
								<%Else %>
									<td><a data-toggle="modal" data-show="true" href="deleteTermModal.asp?i=<%= rscust.Fields("InternalRecordIdentifier") %>" data-target="#myModal"><i class="fas fa-trash-alt"></i></a></td>
								<% End If %>
							<%Else %>
								<td>&nbsp;</td>							
							<% End If %>
					   	</tr>
					<%
						rscust.movenext
					loop
				End If
				set rscust = Nothing
				cnncust.close
				set cnncust = Nothing
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