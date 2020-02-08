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

<h1 class="page-header">Add / Edit Symptom Codes</h1>

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
 			 <a href="addsymptomCode.asp"><button type="button" class="btn btn-success pull-right">Add New Symptom Code</button></a>
		</div>
		
		<div class="table-responsive">
            <table    class="table table-striped table-condensed table-hover table-bordered sortable">
              <thead>
                <tr>
                  <th>Symptom Description</th>
                  <th>Show On <br>Website</th>
                  <th class="sorttable_nosort">Delete</th>
                </tr>
              </thead>
               <tbody class='searchable'>
              
				<%
			
				SQL = "SELECT * FROM FS_SymptomCodes order by SymptomDescription"
		
				Set cnnsymptomCodes = Server.CreateObject("ADODB.Connection")
				cnnsymptomCodes.open (Session("ClientCnnString"))
				Set rssymptomCodes = Server.CreateObject("ADODB.Recordset")
				rssymptomCodes.CursorLocation = 3 
				Set rssymptomCodes = cnnsymptomCodes.Execute(SQL)
		
				If not rssymptomCodes.EOF Then

					Do While Not rssymptomCodes.EOF
				
			        %>
						<!-- table line !-->
						<tr>
							<%If rssymptomCodes.Fields("InternalRecordIdentifier") = 0 Then %>
								<td><%= rssymptomCodes.Fields("SymptomDescription")%></td>
							<% Else %>
								<td><a href='editsymptomCode.asp?i=<%= rssymptomCodes.Fields("InternalRecordIdentifier")%>'><%= rssymptomCodes.Fields("SymptomDescription")%></a></td>
							<% End If%>
							<td><%If rssymptomCodes("ShowOnWebsite") = 1 Then Response.Write("Yes") Else Response.Write("No")%></td>
							<%' Allow delete or display modal
							If rssymptomCodes.Fields("InternalRecordIdentifier") <> 0 Then 
								If NumberOfTicketsBySymptomCode(rssymptomCodes("InternalRecordIdentifier")) = 0 Then %>
									<td><a href='deletesymptomCodeQues.asp?i=<%=rssymptomCodes.Fields("InternalRecordIdentifier")%>'><i class="fas fa-trash-alt"></i></a></td>
								<%Else%>
									<td><a data-toggle="modal" data-show="true" href='deletesymptomCodeModal.asp?i=<%=rssymptomCodes("InternalRecordIdentifier")%>' data-target="#myModal"><i class="fas fa-trash-alt"></i></a></td> 
								<%End If
							Else %>
								<td>&nbsp;</td>
							<% End If %>	
					   	</tr>
					<%
						rssymptomCodes.movenext
					loop
				End If
				set rssymptomCodes = Nothing
				cnnsymptomCodes.close
				set cnnsymptomCodes = Nothing
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