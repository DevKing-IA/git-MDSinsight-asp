<!--#include file="../../inc/header.asp"-->
<!--#include file="../../inc/jquery_table_search.asp"-->
<!--#include file="../../inc/InsightFuncs_Routing.asp"-->

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
		width:1600px !important;
		/*margin:0 auto;*/
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
	 
	 .green{
		background-color:green;
		text-align:center;
		vertical-align:middle;
		margin-top:20px;
		color:#FFF;
	}
	
	 .red{
		background-color:red;
		text-align:center;
		vertical-align:middle;	
		margin-top:20px;	
	}
	
</style>

<!--- eof on/off scripts !-->

<h1 class="page-header">Add / Edit <%= GetTerm("Routing") %>&nbsp;<%= GetTerm("Routes") %></h1>

<div class="row">
 	<div class="col-lg-12">
	 	<p>
 <a href="<%= BaseURL %>routing/routes/addRoute.asp">
    	<button type="button" class="btn btn-success">Add New <%= GetTerm("Route") %></button>
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
            <table class="table table-striped table-condensed table-hover table-bordered sortable">
              <thead>
                <tr>
                  <th>Route</th>
                  <th>Desc</th>
                  <th>Default Driver</th>
                  <th>Show On</th>
                  <th>3rd Party Carrier</th>                  
                  <th>Monday</th>
                  <th>Tuesday</th>
                  <th>Wednesday</th>
                  <th>Thursday</th>
                  <th>Friday</th>
                  <th>Saturday</th>
                  <th>Sunday</th>
                  <th>Created On</th>
                  <th class="sorttable_nosort">Edit</th>                  
                  <th class="sorttable_nosort">Delete</th>
                </tr>
              </thead>
              <tbody class='searchable'>
              
				<%
			
				SQL = "SELECT * FROM RT_Routes ORDER BY RecordCreationDateTime Desc"
		
				Set cnn8 = Server.CreateObject("ADODB.Connection")
				cnn8.open (Session("ClientCnnString"))
				Set rs = Server.CreateObject("ADODB.Recordset")
				rs.CursorLocation = 3 
				Set rs = cnn8.Execute(SQL)
		
				If NOT rs.EOF Then

					Do While Not rs.EOF
					
						InternalRecordIdentifier = rs.Fields("InternalRecordIdentifier")
						RecordCreationDateTime = rs.Fields("RecordCreationDateTime")
						RouteID = rs.Fields("RouteID")
						RouteDescription = rs.Fields("RouteDescription")
						ShowOnDBoard = rs.Fields("ShowOnDBoard")
						ShowInWebApp = rs.Fields("ShowInWebApp")
						ShowInPlanner = rs.Fields("ShowInPlanner")
						ThirdPartyCarrier = rs.Fields("ThirdPartyCarrier")
						DefaultDriverUserNo = rs.Fields("DefaultDriverUserNo")
						DefaultDriverName = GetUserFirstAndLastNameByUserNo(DefaultDriverUserNo)
						MondayRoute = rs.Fields("Monday")
						TuesdayRoute = rs.Fields("Tuesday")
						WednesdayRoute = rs.Fields("Wednesday")
						ThursdayRoute = rs.Fields("Thursday")
						FridayRoute = rs.Fields("Friday")
						SaturdayRoute = rs.Fields("Saturday")
						SundayRoute = rs.Fields("Sunday")
					
				
			        %>
						<!-- table line !-->
						<tr>
							<td><%= RouteID %></td>
							<td><%= RouteDescription %></td>
							<td><%= DefaultDriverName %></td>
							
							<td>
								<% If ShowOnDBoard = 1 Then %>
									Delivery Board: <strong>YES</strong><br>
								<% Else %>
									Delivery Board: <strong>NO</strong><br>
								<% End If %>

								<% If ShowInPlanner = 1 Then %>
									Delivery Board Planner: <strong>YES</strong><br>
								<% Else %>
									Delivery Board Planner: <strong>NO</strong><br>
								<% End If %>
								
								<% If ShowInWebApp = 1 Then %>
									Web App: <strong>YES</strong>
								<% Else %>
									Web App: <strong>NO</strong>
								<% End If %>
							</td>
							
							<td align="center" width="5%">
								<% If ThirdPartyCarrier = 1 Then %>
									<strong>YES</strong><br>
								<% Else %>
									<strong>NO</strong><br>
								<% End If %>
							</td>
							
							<% If MondayRoute = 1 Then %>
								<td class="green" width="5%"><i class="fa fa-4x fa-check-circle-o" aria-hidden="true"></i></td>
							<% Else %>
								<td class="red" width="5%"><i class="fa fa-4x fa-times-circle" aria-hidden="true"></i></td>
							<% End If %>
							
							<% If TuesdayRoute = 1 Then %>
								<td class="green" width="5%"><i class="fa fa-4x fa-check-circle-o" aria-hidden="true"></i></td>
							<% Else %>
								<td class="red" width="5%"><i class="fa fa-4x fa-times-circle" aria-hidden="true"></i></td>
							<% End If %>
							
							<% If WednesdayRoute = 1 Then %>
								<td class="green" width="5%"><i class="fa fa-4x fa-check-circle-o" aria-hidden="true"></i></td>
							<% Else %>
								<td class="red" width="5%"><i class="fa fa-4x fa-times-circle" aria-hidden="true"></i></td>
							<% End If %>
							
							<% If ThursdayRoute = 1 Then %>
								<td class="green" width="5%"><i class="fa fa-4x fa-check-circle-o" aria-hidden="true"></i></td>
							<% Else %>
								<td class="red" width="5%"><i class="fa fa-4x fa-times-circle" aria-hidden="true"></i></td>
							<% End If %>
							
							<% If FridayRoute = 1 Then %>
								<td class="green" width="5%"><i class="fa fa-4x fa-check-circle-o" aria-hidden="true"></i></td>
							<% Else %>
								<td class="red" width="5%"><i class="fa fa-4x fa-times-circle" aria-hidden="true"></i></td>
							<% End If %>
							
							<% If SaturdayRoute = 1 Then %>
								<td class="green" width="5%"><i class="fa fa-4x fa-check-circle-o" aria-hidden="true"></i></td>
							<% Else %>
								<td class="red" width="5%"><i class="fa fa-4x fa-times-circle" aria-hidden="true"></i></td>
							<% End If %>
							
							<% If SundayRoute = 1 Then %>
								<td class="green" width="5%"><i class="fa fa-4x fa-check-circle-o" aria-hidden="true"></i></td>
							<% Else %>
								<td class="red" width="5%"><i class="fa fa-4x fa-times-circle" aria-hidden="true"></i></td>
							<% End If %>
														
							<td align="center"><%= FormatDateTime(RecordCreationDateTime,2) %></td>
							
							<td align="center"><a href='editRoute.asp?i=<%= InternalRecordIdentifier %>'><i class="fa fa-pencil"></i></a></td>			
							<td align="center"><i class="fas fa-trash-alt"></i></td>
 
					   	</tr>
					<%
						rs.movenext
					Loop
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