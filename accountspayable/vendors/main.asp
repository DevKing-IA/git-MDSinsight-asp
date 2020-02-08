<!--#include file="../../inc/header.asp"-->
<!--#include file="../../inc/jquery_table_search.asp"-->
<!--#include file="../../inc/InsightFuncs_InventoryControl.asp"-->



 
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
	max-width:1100px;
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

<h1 class="page-header"><i class="fa fa-handshake-o" aria-hidden="true"></i> Add / Edit Vendors</h1>

<div class="row">
 	<div class="col-lg-12">
	 	<p>
 <a href="addVendor.asp">
    	<button type="button" class="btn btn-success">Add New Vendor</button>
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
                  <th>Vendor</th>
                  <th>API Key</th>
                  <th>Primary Contact Info</th>
                  <th>Address</th>
                  <th class="sorttable_nosort">Delete</th>
                </tr>
              </thead>
              <tbody class='searchable'>
              
				<%
			
				SQL = "SELECT * FROM AP_Vendor ORDER BY RecordCreationDateTime Desc"
		
				Set cnn8 = Server.CreateObject("ADODB.Connection")
				cnn8.open (Session("ClientCnnString"))
				Set rs = Server.CreateObject("ADODB.Recordset")
				rs.CursorLocation = 3 
				Set rs = cnn8.Execute(SQL)
		
				If NOT rs.EOF Then

					Do While Not rs.EOF
					
						InternalRecordIdentifier = rs.Fields("InternalRecordIdentifier")
						RecordCreationDateTime = rs.Fields("RecordCreationDateTime")
						vendorAPIKey = rs.Fields("vendorAPIKey")
						vendorCompanyName = rs.Fields("vendorCompanyName")
						vendorPrimaryContactName = rs.Fields("vendorPrimaryContactName")
						vendorPrimaryContactEmail = rs.Fields("vendorPrimaryContactEmail")
						vendorTechnicalContactName = rs.Fields("vendorTechnicalContactName")
						vendorTechnicalContactEmail = rs.Fields("vendorTechnicalContactEmail")
						vendorAddress = rs.Fields("vendorAddress")
						vendorAddress2 = rs.Fields("vendorAddress2")
						vendorCity = rs.Fields("vendorCity")
						vendorState = rs.Fields("vendorState")
						vendorZip = rs.Fields("vendorZip")
						vendorPhone = rs.Fields("vendorPhone")
						vendorFax = rs.Fields("vendorFax")
						
						vendorFullAddressInfo = ""
						
						If vendorAddress <> "" Then vendorFullAddressInfo = vendorAddress 
						
						If 	vendorAddress <> "" and vendorAddress2 <> "" Then vendorFullAddressInfo = vendorFullAddressInfo & ", "	
						
						If vendorAddress2 <> "" Then
							vendorFullAddressInfo = vendorFullAddressInfo & vendorAddress2 & "<br>" & vendorCity & ", " & vendorState & "  " & vendorZip
						Else
							vendorFullAddressInfo = vendorFullAddressInfo & "<br>" & vendorCity & ", " & vendorState & "  " & vendorZip
						End If		
				
				
			        %>
						<!-- table line !-->
						<tr>
							<td>
							<a href="editVendor.asp?i=<%= InternalRecordIdentifier %>"><%= vendorCompanyName %></a></td>
							<td><a href="editVendor.asp?i=<%= InternalRecordIdentifier %>"><%= vendorAPIKey %></a></td>
							
							<% If vendorPrimaryContactName <> "" AND vendorPrimaryContactEmail <> "" Then %>
								<td><a href="editVendor.asp?i=<%= InternalRecordIdentifier %>"><%= vendorPrimaryContactName %> (<%= vendorPrimaryContactEmail %>)</a></td>	
							<% Else %>
								<td>&nbsp;</td>
							<% End If %>
							
							<td><a href="editVendor.asp?i=<%= InternalRecordIdentifier %>"><%= vendorFullAddressInfo %></a></td>

							<td><a href="deleteVendorQues.asp?i=<%= InternalRecordIdentifier %>"><i class="fas fa-trash-alt"></i></a></td>
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

								

<!--#include file="../../inc/footer-main.asp"-->