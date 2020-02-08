<!--#include file="../../inc/header.asp"-->
<!--#include file="../../inc/jquery_table_search.asp"-->
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
	max-width:1250px;
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

<h1 class="page-header">Add / Edit Region</h1>

<div class="row">
 	<div class="col-lg-12">
	 	<p>
 <a href="addregion.asp">
    	<button type="button" class="btn btn-success">Add New Region</button>
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
                  <th class="col-xs-3" style="text-align:center;vertical-align:middle">Region</th>
                  <!--<th>Cities</th>-->
				  <th class="col-xs-5" style="text-align:center;vertical-align:middle">Zip Or Postal Codes</th>
				  <th class="col-xs-2" style="text-align:center;vertical-align:middle">States Or Provinces</th>
				  <th class="col-xs-1" style="text-align:center;vertical-align:middle">Use For Service Tickets</th>
                  <th class="col-xs-1 sorttable_nosort" style="text-align:center;vertical-align:middle">Delete</th>
                </tr>
              </thead>
               <tbody class='searchable'>
              
				<%
			
				SQL = "SELECT * FROM AR_Regions order by Region"
		
				Set cnnregions = Server.CreateObject("ADODB.Connection")
				cnnregions.open (Session("ClientCnnString"))
				Set rsregions = Server.CreateObject("ADODB.Recordset")
				rsregions.CursorLocation = 3 
				Set rsregions = cnnregions.Execute(SQL)
		
				If not rsregions.EOF Then

					Do While Not rsregions.EOF
				
			        %>
						<!-- table line !-->
						<tr>
							<%If rsregions.Fields("InternalRecordIdentifier") = 0 Then %>
								<td><%= rsregions.Fields("Region")%></td>
							<% Else %>
								<td><a href='editregion.asp?i=<%= rsregions.Fields("InternalRecordIdentifier")%>'><%= rsregions.Fields("Region")%></a></td>
							<% End If%>
							<%
							'Cities = rsregions.Fields("Cities1") & ", " & rsregions.Fields("Cities2") & ", " & rsregions.Fields("Cities3")
							
							Cities = rsregions.Fields("Cities1")
							
							If rsregions.Fields("Cities2") <> "" Then
								Cities = Cities & "," & rsregions.Fields("Cities2")
							End If
							
							If rsregions.Fields("Cities3") <> "" Then
								Cities = Cities & "," & rsregions.Fields("Cities3")
							End If

							'ZipOrPostalCodes = rsregions.Fields("ZipOrPostalCodes1") & ", " & rsregions.Fields("ZipOrPostalCodes2")
							
							ZipOrPostalCodes = rsregions.Fields("ZipOrPostalCodes1")
							
							If rsregions.Fields("ZipOrPostalCodes2") <> "" Then
								ZipOrPostalCodes = ZipOrPostalCodes & "," & rsregions.Fields("ZipOrPostalCodes2")
							End If
							%>
							
							
							<!--<td style="width:40%; word-break: break-word;"><%= Cities%></td>-->
							
							<% If rsregions("CatchAllRegionIntRecIDs") <> "" Then %>
								<td><strong>CATCH-ALL</strong>&nbsp;<%= ZipOrPostalCodes %></td>
							<% Else %>
								<td>
									<% If rsregions.Fields("InternalRecordIdentifier") <> 0 Then %>
										<button data-toggle="collapse" data-target="#ZipOrPostalCodes<%= rsregions.Fields("InternalRecordIdentifier") %>" class="btn btn-success">View Zip/Postal Codes&nbsp;<i class="fas fa-map-marked-alt"></i></button>
										
										<div id="ZipOrPostalCodes<%= rsregions.Fields("InternalRecordIdentifier") %>" class="collapse">
											<%= ZipOrPostalCodes %>
										</div>	
									<% End If %>							
								</td>
							<% End If %>

							<td><%= rsregions.Fields("StatesOrProvinces")%></td>
							
							<% If rsregions("UseRegionForServiceTickets") = 0 Then %>
								<td align="center">NO</td>
							<% Else %>
								<td align="center">YES</td>
							<% End If %>
														
							<%' Allow delete or display modal
							If rsregions.Fields("InternalRecordIdentifier") <> 0 Then 								
								%>
									<td align="center"><a href='deleteregionQues.asp?i=<%=rsregions.Fields("InternalRecordIdentifier")%>'><i class="fas fa-trash-alt"></i></a></td>
							<%Else %>
								<td>&nbsp;</td>
							<% End If %>	
					   	</tr>
					<%
						rsregions.movenext
					loop
				End If
				set rsregions = Nothing
				cnnregions.close
				set cnnregions = Nothing
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