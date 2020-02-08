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
 </style>

<!--- eof on/off scripts !-->

<h1 class="page-header"><i class="fa fa-handshake-o" aria-hidden="true"></i> Add / Edit Partners</h1>

<div class="row">
 	<div class="col-lg-12">
	 	<p>
 <a href="addPartner.asp">
    	<button type="button" class="btn btn-success">Add New Partner</button>
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
                  <th>Partner</th>
                  <th>API Key</th>
                  <th>Created On</th>
                  <th>Primary Contact</th>
                  <th>Contact Info</th>
                  <th>Taxable SKU Mapping</th>
                  <th>Non Taxable SKU Mapping</th>
                  <th>Cust Account Mapping</th>
                  <th>Allow Blank Prod Desc</th>
                  <th>Allow Blank Prod UOM</th>
                  <th>Current Status</th>
                  <th class="sorttable_nosort">Disable/Enable</th>
                </tr>
              </thead>
              <tbody class='searchable'>
              
				<%
			
				SQL = "SELECT * FROM IC_Partners ORDER BY RecordCreationDateTime Desc"
		
				Set cnn8 = Server.CreateObject("ADODB.Connection")
				cnn8.open (Session("ClientCnnString"))
				Set rs = Server.CreateObject("ADODB.Recordset")
				rs.CursorLocation = 3 
				Set rs = cnn8.Execute(SQL)
		
				If NOT rs.EOF Then

					Do While Not rs.EOF
					
						InternalRecordIdentifier = rs.Fields("InternalRecordIdentifier")
						RecordCreationDateTime = rs.Fields("RecordCreationDateTime")
						partnerAPIKey = rs.Fields("partnerAPIKey")
						partnerCompanyName = rs.Fields("partnerCompanyName")
						partnerPrimaryContactName = rs.Fields("partnerPrimaryContactName")
						partnerPrimaryContactEmail = rs.Fields("partnerPrimaryContactEmail")
						partnerTechnicalContactName = rs.Fields("partnerTechnicalContactName")
						partnerTechnicalContactEmail = rs.Fields("partnerTechnicalContactEmail")
						partnerAddress = rs.Fields("partnerAddress")
						partnerAddress2 = rs.Fields("partnerAddress2")
						partnerCity = rs.Fields("partnerCity")
						partnerState = rs.Fields("partnerState")
						partnerZip = rs.Fields("partnerZip")
						partnerPhone = rs.Fields("partnerPhone")
						partnerFax = rs.Fields("partnerFax")
						partnerDisabled = rs.Fields("partnerDisabled")
						partnerUnmappedTaxableSKU = rs.Fields("partnerUnmappedTaxableSKU")
						partnerUnmappedTaxableUM = rs.Fields("partnerUnmappedTaxableUM")
						partnerUnmappedTaxablePassOriginalSKU = rs.Fields("partnerUnmappedTaxablePassOriginalSKU")
						partnerUnmappedNonTaxableSKU = rs.Fields("partnerUnmappedNonTaxableSKU")
						partnerUnmappedNonTaxableUM = rs.Fields("partnerUnmappedNonTaxableUM")
						partnerUnmappedNonTaxablePassOriginalSKU = rs.Fields("partnerUnmappedNonTaxablePassOriginalSKU")
						partnerUnmappedCustomerID = rs.Fields("partnerUnmappedCustomerID")
						partnerUnmappedPassOriginalCustomerID = rs.Fields("partnerUnmappedPassOriginalCustomerID")
						
						partnerRejectsBlankProdDescs = rs.Fields("partnerRejectsBlankProdDescs")
						partnerRejectsBlankProdUOMS = rs.Fields("partnerRejectsBlankProdUOMS")
						
						If partnerRejectsBlankProdDescs = 1 OR partnerRejectsBlankProdDescs = true Then partnerRejectsBlankProdDescs = "YES" Else partnerRejectsBlankProdDescs = "NO"
						If partnerRejectsBlankProdUOMS = 1 OR partnerRejectsBlankProdUOMS = true Then partnerRejectsBlankProdUOMS = "YES" Else partnerRejectsBlankProdUOMS = "NO"

						If partnerAddress2 = "" Then
							partnerFullAddressInfo = partnerAddress & ", " & partnerAddress2 & "<br>" & partnerCity & ", " & partnerState & "  " & partnerZip
						Else
							partnerFullAddressInfo = partnerAddress & "<br>" & partnerCity & ", " & partnerState & "  " & partnerZip
						End If					
				
			        %>
						<!-- table line !-->
						<tr>
							<td>
							<a href="editPartner.asp?i=<%= InternalRecordIdentifier %>"><%= partnerCompanyName %> (<%= NumberOfSKUsDefinedForPartner(InternalRecordIdentifier) %> SKUs)</a></td>
							<td><a href="editPartner.asp?i=<%= InternalRecordIdentifier %>"><%= partnerAPIKey %></a></td>
							<td><a href="editPartner.asp?i=<%= InternalRecordIdentifier %>"><%= RecordCreationDateTime %></a></td>
							
							<% If partnerPrimaryContactName <> "" AND partnerPrimaryContactEmail <> "" Then %>
								<td><a href="editPartner.asp?i=<%= InternalRecordIdentifier %>"><%= partnerPrimaryContactName %> (<%= partnerPrimaryContactEmail %>)</a></td>	
							<% Else %>
								<td>&nbsp;</td>
							<% End If %>
							
							<td><a href="editPartner.asp?i=<%= InternalRecordIdentifier %>"><%= partnerFullAddressInfo %></a></td>
							
							<% If partnerUnmappedTaxablePassOriginalSKU = true Then %>
								<td><a href="editPartner.asp?i=<%= InternalRecordIdentifier %>">Use Original Product Code</a></td>
							<% Else %>
								<td><a href="editPartner.asp?i=<%= InternalRecordIdentifier %>">Pass <%= partnerUnmappedTaxableSKU %> ((Unit: <%= partnerUnmappedTaxableUM %>)</a></td>
							<% End If %>
							

							<% If partnerUnmappedNonTaxablePassOriginalSKU = true Then %>
								<td><a href="editPartner.asp?i=<%= InternalRecordIdentifier %>">Use Original Product Code</a></td>
							<% Else %>
								<td><a href="editPartner.asp?i=<%= InternalRecordIdentifier %>">Pass <%= partnerUnmappedNonTaxableSKU %> (Unit: <%= partnerUnmappedNonTaxableUM %>)</a></td>
							<% End If %>
							

							<% If partnerUnmappedPassOriginalCustomerID = true Then %>
								<td><a href="editPartner.asp?i=<%= InternalRecordIdentifier %>">Use Original Customer Number</a></td>
							<% Else %>
								<td><a href="editPartner.asp?i=<%= InternalRecordIdentifier %>">Pass Account <%= partnerUnmappedCustomerID %></a></td>
							<% End If %>
							
							
							<td><a href="editPartner.asp?i=<%= InternalRecordIdentifier %>"><%= partnerRejectsBlankProdDescs %></a></td>
							<td><a href="editPartner.asp?i=<%= InternalRecordIdentifier %>"><%= partnerRejectsBlankProdUOMS %></a></td>
							
							<% If partnerDisabled = 0 Then %>
								<td>Enabled</td>
								<td><a href="disablePartnerQues.asp?i=<%= InternalRecordIdentifier %>"><i class="fa fa-toggle-off"></i></a></td>
							<% Else %>
								<td>Disabled</td>
								<td><a href="enablePartnerQues.asp?i=<%= InternalRecordIdentifier %>"><i class="fa fa-toggle-on"></i></a></td>
							<% End If %>
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