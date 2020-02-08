<!--#include file="../../inc/header.asp"-->
<!--#include file="../../inc/jquery_table_search.asp"-->
<!--#include file="../../inc/InsightFuncs_Equipment.asp"-->

<script type="text/javascript">
$(document).ready(function() {
    $("#PleaseWaitPanel").hide();
});
</script>


<% 
	InternalRecordIdentifier = Request.QueryString("i") 
	If InternalRecordIdentifier = "" Then Response.Redirect("main.asp")
%>

 
<style type="text/css">

	.table-responsive{
		width:1800px;
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
		max-width:2000px;
		margin-left:20px;
	}
	
	.narrow-results{
		margin:0px 0px 20px 0px;
	}
	
	#filter{
		width:60%;
	}
	
	#PleaseWaitPanel{
		position: fixed;
		left: 470px;
		top: 275px;
		width: 975px;
		height: 300px;
		z-index: 9999;
		background-color: #fff;
		opacity:1.0;
		text-align:center;
	}  
	
 </style>

<!--- eof on/off scripts !-->

<%
Response.Write("<div id=""PleaseWaitPanel"">")
Response.Write("<br><br>Processing Equipment Detail Data<br><br>This may take up to a full minute, please wait...<br><br>")
Response.Write("<img src=""../../img/loading.gif"" />")
Response.Write("</div>")
Response.Flush()
%>


<h1 class="page-header">View Movement Code Details For <strong><%= GetMovementCodeByIntRecID(InternalRecordIdentifier) %> - <%= GetMovementCodeDescByIntRecID(InternalRecordIdentifier) %></strong></h1>

<div class="row">
 	<div class="col-lg-12">
	 	<p><a href="<%= BaseURL %>equipment/MovementCodes/main.asp"><button type="button" class="btn btn-success">Return To Main Movement Codes Screen</button></a></p>
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
					<th>Model</th>
					<th>Customer</th>
					<th>Acct. #</th>
					<th>Status</th>
					<th>Movement Code</th>
					<th>Serial Number</th>
					<th>Asset Tag 1</th>
					<th>Purchase Or Aquisition Type</th>
					<th>Purchased From Vendor ID</th>
					<th>Purchase PO Number</th>
					<th>Purchase Date</th>
					<th>Purchase Cost</th>
					<th>Replacement Cost</th>
					<th>Aquired Condition</th>
					<th>Current Condition</th>              
					<th>Warrenty Start Date</th>
					<th>Warrenty End Date</th>              
                  	<th class="sorttable_nosort">Comments</th>
                </tr>
              </thead>
              <tbody class='searchable'>
              
				<%
			
				SQL = "SELECT * FROM EQ_Equipment INNER JOIN EQ_Models ON EQ_Models.InternalRecordIdentifier = EQ_Equipment.ModelIntRecID "
				SQL = SQL & " INNER JOIN EQ_CustomerEquipment ON EQ_CustomerEquipment.EquipIntRecID = EQ_Equipment.InternalRecordIdentifier "
				SQL = SQL & " WHERE EQ_Equipment.MovementCodeIntRecID = " & InternalRecordIdentifier & " ORDER BY EQ_Equipment.ModelIntRecID "
		
				Set cnn8 = Server.CreateObject("ADODB.Connection")
				cnn8.open (Session("ClientCnnString"))
				Set rs = Server.CreateObject("ADODB.Recordset")
				rs.CursorLocation = 3 
				Set rs = cnn8.Execute(SQL)
		
				If NOT rs.EOF Then

					Do While Not rs.EOF
					
				      RecordSource = rs("RecordSource")
				      CustID = rs("CustID")
				      ModelIntRecID = rs("ModelIntRecID")
				      StatusCodeIntRecID = rs("StatusCodeIntRecID")
				      MovementCodeIntRecID = rs("MovementCodeIntRecID")
				      SerialNumber = rs("SerialNumber")
				      AssetTag1 = rs("AssetTag1")
				      AcquisitionCodeIntRecID = rs("AcquisitionCodeIntRecID")
				      PurchasedFromVendorID = rs("PurchasedFromVendorID")
				      PurchasedViaPONumber = rs("PurchasedViaPONumber")
				      PurchaseDate = rs("PurchaseDate")
				      PurchaseCost = rs("PurchaseCost")
				      ReplacementCost = rs("ReplacementCost")
				      AquiredConditionIntRecID = rs("AquiredConditionIntRecID")
				      CurrentConditionIntRecID = rs("CurrentConditionIntRecID")
				      WarrentyStartDate = rs("WarrentyStartDate")
				      WarrentyEndDate = rs("WarrentyEndDate")
				      Comments = rs("Comments")
			        %>
						<!-- table line !-->
						<tr>
							<td><%= GetModelNameByIntRecID(ModelIntRecID) %></td>
							
							<td><%= GetCustNameByCustNum(CustID) %></td>

							<% If InStr(CustID,"<") OR InStr(CustID,">") Then %>
								<% 
									CustID = Replace(CustID, "<", "&#60;")
									CustID = Replace(CustID, ">", "&#62;")
								%>
							<% End If %>
							
							<td><%= CustID %></td>							
							
							<% If StatusCodeIntRecID <> "" Then %>
								<td><%= GetStatusCodeNameByIntRecID(StatusCodeIntRecID) %></td>
							<% Else %>
								<td>&nbsp;</td>
							<% End If %>
							
							<% If MovementCodeIntRecID <> "" Then %>
								<td><%= GetMovementCodeDescByIntRecID(MovementCodeIntRecID) %></td>
							<% Else %>
								<td>&nbsp;</td>
							<% End If %>
							
							<td><%= SerialNumber %></td>
							<td><%= AssetTag1 %></td>

							<% If AcquisitionCodeIntRecID <> "" Then %>
								<td><%= GetAcquisitionCodeDescByIntRecID(AcquisitionCodeIntRecID) %></td>
							<% Else %>
								<td>&nbsp;</td>
							<% End If %>

							<td><%= PurchasedFromVendorID %></td>
							<td><%= PurchasedViaPONumber %></td>
							<td><%= PurchaseDate %></td>
							
							<% If PurchaseCost <> "" Then %>
								<td><%= FormatCurrency(PurchaseCost,2) %></td>
							<% Else %>
								<td>&nbsp;</td>
							<% End If %>

							<% If ReplacementCost <> "" Then %>
								<td><%= FormatCurrency(ReplacementCost,2) %></td>
							<% Else %>
								<td>&nbsp;</td>
							<% End If %>
							
							<% If AquiredConditionIntRecID <> "" Then %>
								<td><%= GetConditionNameByIntRecID(AquiredConditionIntRecID) %></td>
							<% Else %>
								<td>&nbsp;</td>
							<% End If %>
								
							<% If CurrentConditionIntRecID <> "" Then %>
								<td><%= GetConditionNameByIntRecID(CurrentConditionIntRecID) %></td>
							<% Else %>
								<td>&nbsp;</td>
							<% End If %>
							
							<td><%= WarrentyStartDate %></td>
							<td><%= WarrentyEndDate %></td>
							<td><%= Comments %></td>
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
								

<!--#include file="../../inc/footer-main.asp"-->