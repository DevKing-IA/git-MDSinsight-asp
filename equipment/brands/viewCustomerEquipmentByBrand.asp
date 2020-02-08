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


<h1 class="page-header">View Customer Equipment Brand Details For <strong><%= GetBrandNameByIntRecID(InternalRecordIdentifier) %></strong></h1>

<div class="row">
 	<div class="col-lg-12">
	 	<p><a href="<%= BaseURL %>equipment/brands/main.asp"><button type="button" class="btn btn-success">Return To Main Brands Screen</button></a></p>
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
					<th>Customer</th>
					<th>Acct. #</th>
				  	<th>Description/Type</th>
				  	<th>Status</th>
				  	<th>Movement Code</th>
				  	<th>Frequency</th>
				  	<th>Install Date</th>
				  	<th>Equip. Value</th>
				  	<th>Serial #</th>
				  	<th>Asset #</th>
                </tr>
              </thead>
              <tbody class='searchable'>
              
				<%
			
				SQLCustomerEquipment = "SELECT * FROM EQ_CustomerEquipment INNER JOIN "
				SQLCustomerEquipment = SQLCustomerEquipment & " EQ_Equipment ON EQ_Equipment.InternalRecordIdentifier = EQ_CustomerEquipment.EquipIntRecID "
				SQLCustomerEquipment = SQLCustomerEquipment & " INNER JOIN "
				SQLCustomerEquipment = SQLCustomerEquipment & " EQ_Models ON EQ_Equipment.ModelIntRecID = EQ_Models.InternalRecordIdentifier "
				SQLCustomerEquipment = SQLCustomerEquipment & " WHERE EQ_Models.BrandIntRecID = " & InternalRecordIdentifier & " ORDER BY EQ_Equipment.ModelIntRecID "
				
		
				Set cnnCustomerEquipment = Server.CreateObject("ADODB.Connection")
				cnnCustomerEquipment.open (Session("ClientCnnString"))
				Set rsCustomerEquipment = cnnCustomerEquipment.Execute(SQLCustomerEquipment)
		
				If NOT rsCustomerEquipment.EOF Then

					TotalRentalAmount = 0
					TotalPurchaseCost = 0

					Do While Not rsCustomerEquipment.EOF
					
				      	CustID = rsCustomerEquipment("CustID")
				      	RecordSource = rsCustomerEquipment("RecordSource")
						InstallDate = rsCustomerEquipment("InstallDate")
						StatusCodeIntRecID = rsCustomerEquipment("StatusCodeIntRecID")
						
						SQLEquipStatusCode = "SELECT * FROM EQ_StatusCodes WHERE InternalRecordIdentifier = " & StatusCodeIntRecID
							
						Set cnnEquipStatusCode = Server.CreateObject("ADODB.Connection")
						cnnEquipStatusCode.open (Session("ClientCnnString"))
						Set rsEquipStatusCode = cnnEquipStatusCode.Execute(SQLEquipStatusCode)
						
						If NOT rsEquipStatusCode.EOF Then
							InstallType = rsEquipStatusCode("statusBackendSystemCode")
							InstallTypeFullName = rsEquipStatusCode("statusDesc")
						Else
							InstallType = ""
							InstallTypeFullName = ""
						End If
												
						
						If InstallType = "R" then
						
							RentalFrequencyType = rsCustomerEquipment("RentalFrequencyType")
							
							Select Case RentalFrequencyType
							Case "D"
								RentalFrequencyFullName = "DAYS"
							Case "M"
								RentalFrequencyFullName = "MONTH(S)"
							Case "Y"
								RentalFrequencyFullName = "YEAR(S)"
							End Select
							
							RentalFrequencyNumber = rsCustomerEquipment("RentalFrequencyNumber")
							RentAmt = rsCustomerEquipment("RentAmt")
							
							If RentAmt <> "" Then
								TotalRentalAmount = TotalRentalAmount + RentAmt
								RentAmt = FormatCurrency(RentAmt,0)
							Else
								RentAmt = 0
								RentAmt = FormatCurrency(RentAmt,0)
							End If
							
						Else
							RentalFrequencyFullName = ""
							RentalFrequencyType = ""
							RentalFrequencyNumber = ""
							RentAmt = 0
							RentAmt = FormatCurrency(RentAmt,0)
						End If
						
						
						MovementCodeIntRecID = rsCustomerEquipment("MovementCodeIntRecID")
						
						If MovementCodeIntRecID <> "" Then
						
							SQLEquipMovementCode = "SELECT * FROM EQ_MovementCodes WHERE InternalRecordIdentifier = " & MovementCodeIntRecID
								
							Set cnnEquipMovementCode = Server.CreateObject("ADODB.Connection")
							cnnEquipMovementCode.open (Session("ClientCnnString"))
							Set rsEquipMovementCode = cnnEquipMovementCode.Execute(SQLEquipMovementCode)
							
							If NOT rsEquipMovementCode.EOF Then
								MovementCode = rsEquipMovementCode("movementCode")
								MovementCodeDesc = rsEquipMovementCode("movementCodeDesc")
							Else
								MovementCode = ""
								MovementCodeDesc = ""
							End If
							
						Else
							MovementCode = ""
							MovementCodeDesc = ""
						End If

						
						SerialNumber = rsCustomerEquipment("SerialNumber")
						PurchaseCost = rsCustomerEquipment("PurchaseCost")
						
						If PurchaseCost <> "" then
							TotalPurchaseCost = TotalPurchaseCost + PurchaseCost
							PurchaseCost = FormatCurrency(PurchaseCost,2)
						End If
						
						ModelIntRecID = rsCustomerEquipment("ModelIntRecID")
						
						If ModelIntRecID <> 0 Then
							BrandName = GetBrandNameByModelIntRecID(ModelIntRecID)
						Else
							BrandName = ""
						End If
						
						AssetTag1 = rsCustomerEquipment("AssetTag1")
						Description = "DESC NEEDED"
						Description  = GetModelNameByIntRecID(rsCustomerEquipment("ModelIntRecID"))
						
						ModelCount = GetTotalNumberOfModelsForCustomer(CustID,ModelIntRecID)
      

			        %>
						<!-- table line !-->
						<tr>
							
							<td><%= GetCustNameByCustNum(CustID) %></td>
							
							<% If InStr(CustID,"<") OR InStr(CustID,">") Then %>
								<% 
									CustID = Replace(CustID, "<", "&#60;")
									CustID = Replace(CustID, ">", "&#62;")
								%>
							<% End If %>
							
							<td><%= CustID %></td>
							<% If BrandName <> "" Then %>
								<td><%= UCASE(BrandName) %>&nbsp;<%= Description %></td>
							<% Else %>
								<td><%= Description %></td>
							<% End If %>
							<td><%= InstallTypeFullName %></td>
							<td><%= MovementCode %> - <%= MovementCodeDesc %></td>
							<% If InstallType = "R" Then %>
								<td><%= RentalFrequencyNumber %>&nbsp;<%= RentalFrequencyFullName %></td>
								<td><%= RentAmt %></td>
							<% Else %>
								<td>&nbsp;</td>
								<td><%= RentAmt %></td>								
							<% End If %>
							<td align="right"><%= InstallDate %></td>
							<% If PurchaseCost <> "" Then %>
								<td align="right"><%= FormatCurrency(PurchaseCost,0) %></td>
							<% End If %>
							<td align="center"><%= SerialNumber %></td>
							<td align="center"><%= AssetTag1 %></td>
					   	</tr>
					<%
						rsCustomerEquipment.movenext
					loop
				End If
				set rsCustomerEquipment= Nothing
				cnnCustomerEquipment.close
				set cnnCustomerEquipment = Nothing
	            %>
			</tbody>
		</table>
	</div>
		</div>

</div>
<!-- eof row !-->    
								

<!--#include file="../../inc/footer-main.asp"-->