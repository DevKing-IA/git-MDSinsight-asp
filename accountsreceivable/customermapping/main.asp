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

<h1 class="page-header"><i class="fa fa-map-marker" aria-hidden="true"></i> Customer Account Mapping Tool</h1>

	<!-- tabs start here !-->
	<div class="container">

	<div class="table-responsive">
            <table    class="table table-striped table-condensed table-hover table-bordered sortable">
              <thead>
                <tr>
                  <th>Partner</th>
                  <th>API Key</th>
                  <th>Created On</th>
                  <th># Accounts Mapped</th>
                  <th class="sorttable_nosort">Manage Accounts</th>
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
						
						If partnerAddress2 = "" Then
							partnerFullAddressInfo = partnerAddress & ", " & partnerAddress2 & "<br>" & partnerCity & ", " & partnerState & "  " & partnerZip
						Else
							partnerFullAddressInfo = partnerAddress & "<br>" & partnerCity & ", " & partnerState & "  " & partnerZip
						End If					
				
			        %>
						<!-- table line !-->
						<tr>
							<td>
							<a href="selectCustomerToEditByAlphabet.asp?i=<%= InternalRecordIdentifier %>"><%= partnerCompanyName %></a></td>
							<td>
							<a href="selectCustomerToEditByAlphabet.asp?i=<%= InternalRecordIdentifier %>"><%= partnerAPIKey %></a></td>
							<td>
							<a href="selectCustomerToEditByAlphabet.asp?i=<%= InternalRecordIdentifier %>"><%= RecordCreationDateTime %></a></td>
							<td>
							<a href="selectCustomerToEditByAlphabet.asp?i=<%= InternalRecordIdentifier %>"><%= NumberOfCustAccountsDefinedForPartner(InternalRecordIdentifier) %> ACCOUNTS</a></td>	
							<td>
							<a href="selectCustomerToEditByAlphabet.asp?i=<%= InternalRecordIdentifier %>"><button type="button" class="btn btn-success">MAP ACCOUNTS</button></a></td>
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