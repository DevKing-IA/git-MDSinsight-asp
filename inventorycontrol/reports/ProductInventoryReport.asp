<%
Server.ScriptTimeout = 900000 'Default value
Dim ReportNumber 
ReportNumber = 1600
%>
<!--#include file="../../inc/header.asp"-->
<!--#include file="../../inc/InsightFuncs.asp"-->
<!--#include file="../../inc/InsightFuncs_InventoryControl.asp"-->
<!--#include file="../../inc/jquery_table_search.asp"-->

<%
CreateAuditLogEntry "Inventory Control Report","Inventory Control Report","Minor",0, MUV_Read("DisplayName") & " ran the report: Product UPC Report"

'************************
'Read Settings_Reports
'************************
SQL = "SELECT * FROM Settings_Reports WHERE ReportNumber = 1600 AND UserNo = " & Session("userNo")
Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rs = Server.CreateObject("ADODB.Recordset")
Set rs= cnn8.Execute(SQL)
If NOT rs.EOF Then
	UnitUPCData = rs("ReportSpecificData1")
	CaseUPCData = rs("ReportSpecificData2")
	InventoriedItem = rs("ReportSpecificData3")
	PickableItem = rs("ReportSpecificData4")
	ProductCategoriesForInventoryReport = rs("ReportSpecificData5")
	If IsNull(UnitUPCData) Then UnitUPCData = ""
	If IsNull(CaseUPCData) Then CaseUPCData  = ""
	If IsNull(InventoriedItem) Then InventoriedItem = ""
	If IsNull(PickableItem) Then PickableItem = ""
	If IsNull(ProductCategoriesForInventoryReport) Then ProductCategoriesForInventoryReport = ""
Else
	UnitUPCData = ""
	CaseUPCData = ""
	InventoriedItem = ""
	PickableItem = ""
	ProductCategoriesForInventoryReport = ""
End If										
'****************************
'End Read Settings_Reports
'****************************



'**************************************************************************************
'Build WHERE Clause For Unit UPC, Case UPC, Inventoried and Pickable Items
'**************************************************************************************

WHERE_CLAUSE_UNITUPC = ""
WHERE_CLAUSE_CASEUPC = ""
WHERE_CLAUSE_INVENTORIEDITEM = ""
WHERE_CLAUSE_PICKABLEITEM = ""
WHERE_CLAUSE_CATEGORY = ""

If UnitUPCData = "NOTEMPTY" Then
	WHERE_CLAUSE_UNITUPC = " OR (prodUnitUPC <> '') "
ElseIf  UnitUPCData = "EMPTY" Then
	WHERE_CLAUSE_UNITUPC = " OR (prodUnitUPC = '' OR prodUnitUPC IS NULL) "
Else
	WHERE_CLAUSE_UNITUPC = ""
End If


If CaseUPCData = "NOTEMPTY" Then
	WHERE_CLAUSE_CASEUPC = " AND (prodCaseUPC <> '') "
ElseIf  CaseUPCData = "EMPTY" Then
	WHERE_CLAUSE_CASEUPC = " AND (prodCaseUPC = '' OR prodCaseUPC IS NULL) "
Else
	WHERE_CLAUSE_CASEUPC = ""
End If
			
			
If InventoriedItem = "YES" Then
	WHERE_CLAUSE_INVENTORIEDITEM = " AND (prodInventoriedItem = 1) "
ElseIf InventoriedItem = "NO" Then
	WHERE_CLAUSE_INVENTORIEDITEM = " AND (prodInventoriedItem = 0) "
Else
	WHERE_CLAUSE_INVENTORIEDITEM = " "
End If
			
If PickableItem = "YES" Then
	WHERE_CLAUSE_PICKABLEITEM = " AND (prodPickableItem = 1) "
ElseIf InventoriedItem = "NO" Then
	WHERE_CLAUSE_PICKABLEITEM = " AND (prodPickableItem = 0) "
Else
	WHERE_CLAUSE_PICKABLEITEM = ""
End If

CategoryArray = ""
CategoryArray = Split(ProductCategoriesForInventoryReport,",")

'**************************************************************************************
'Build WHERE Clause For Product Category Array
'**************************************************************************************

For z = 0 to UBound(CategoryArray)
	If z = 0 AND UBound(CategoryArray) = 0 Then
		WHERE_CLAUSE_CATEGORY = WHERE_CLAUSE_CATEGORY & " AND (prodCategory = '" & CategoryArray(z) & "' "
	ElseIf z = 0 AND UBound(CategoryArray) > 0 Then
		WHERE_CLAUSE_CATEGORY = WHERE_CLAUSE_CATEGORY & " AND ((prodCategory = '" & CategoryArray(z) & "') "
	Else
		WHERE_CLAUSE_CATEGORY = WHERE_CLAUSE_CATEGORY & " OR (prodCategory = '" & CategoryArray(z) & "')"
	End If
Next	
	
If WHERE_CLAUSE_CATEGORY <> "" Then
	WHERE_CLAUSE_CATEGORY = WHERE_CLAUSE_CATEGORY & ")"
End If

	

	%>
	
<style>

	.bs-example-modal-lg-customize .row{
		margin-bottom: 10px;
	 	width: 100%;
		overflow: hidden;
	}
	
	.bs-example-modal-lg -customize.left-column{
		background: #eaeaea;
		padding-bottom: 1000px;
	    margin-bottom: -1000px;
	}
	
	.bs-example-modal-lg-customize .left-column h4{
		margin-top: 0px;
	}
	
	.bs-example-modal-lg-customize .right-column{
		background: #fff;
		padding-bottom: 1000px;
	    margin-bottom: -1000px;
	} 
table.sortable th:not(.sorttable_sorted):not(.sorttable_sorted_reverse):not(.sorttable_nosort):after { 
    content: " \25B4\25BE" 
    
}

table.sortable thead {
    color:#222;
    font-weight: bold;
    cursor: pointer;
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



<script type="text/javascript">

	$(document).ready(function() {
	    $("#PleaseWaitPanel").hide();
	    
	});

</script>


<%
Response.Write("<div id=""PleaseWaitPanel"">")
Response.Write("<br><br>Processing Today's Product Inventory Data<br><br>This may take up to a full minute, please wait...<br><br>")
Response.Write("<img src=""../../img/loading.gif"" />")
Response.Write("</div>")
Response.Flush()

%>

 

<h3 class="page-header"><i class="fa fa-file-text-o"></i>  for <%=FormatDateTime(Now(),1) %> </h3>

<!-- row !-->
<div class="row">


<!-- responsive tables !-->
<div class="table-responsive">
	
<div class="input-group"> 
	<span class="input-group-addon">Narrow Results</span>
	<div class="row">
		<div class="col-lg-3">
			<input id="filter" type="text" class="form-control filter-search-width" placeholder="Type here...">
		</div>
		<div class="col-lg-3">

			<!-- modal button !-->
			<button type="button" class="btn btn-primary" data-toggle="modal" data-target=".bs-example-modal-lg-customize">
			  Customize
			</button>
			<% If ProductInventoryReportTableSet(ReportNumber)Then %>
				<a href="<%= BaseURL %>inventorycontrol/reports/ProductInventoryReport_Customize_ClearValues.asp"><button type="button" class="btn btn-primary">Clear Customizations</button></a>
			<% End If %>
				<!-- eof modal button !-->
			</h3>

			<!--#include file="ProductInventoryReport_Customize.asp"-->	
			
		</div>
		
		<div class="pull-right" style="margin-right:50px;">
				<button type="button" class="btn btn-success" onclick="downloadCSV()">Export To Excel</button>
				<a href="<%= BaseURL %>inventorycontrol/reports/ProductInventoryReport_PDFLaunch.asp"><button type="button" class="btn btn-success">Generate PDF</button></a>		
		</div>

		<script type="text/javascript">
		
			function downloadCSV() {
				$.ajax
				({
					type: "POST",
					url: "../../inc/InSightFuncs_AjaxForInventoryControlModals.asp",
					data: "action=GenerateInventoryReportCSV&baseURL=" + '<%=BaseURL%>',
					cache: false,
					async: false,
					success: function (url) { window.open(url, 'Download'); }
				});
			}
		
		</script>

	</div>
</div>

<table id="tableSuperSum" class="food_planner sortable table table-striped table-condensed table-hover">
  <thead>
	<tr>
	  <th class="sorttable_numeric">SKU</th>
	  <th class="sorttable">Desc</th>
	  <th class="sorttable">Category</th> 
	  <th class="sorttable">Inventoried</th>
	  <th class="sorttable">Pickable</th>
	  <th class="sorttable">Case Pricing</th>
	  <th class="sorttable">Case Desc</th>
	  <th class="sorttable">Unit UPC</th>
	  <th class="sorttable">Case UPC</th>
	</tr>
  </thead>
              

<%		
	Response.Write("<tbody class='searchable'>")

	SQL = " SELECT prodUnitUPC, prodCaseUPC, prodSKU, prodDescription, prodCategory, "
	SQL = SQL & "prodCasePricing, prodCaseDescription, prodInventoriedItem, prodPickableItem "
	SQL = SQL & "FROM IC_Product "
	SQL = SQL & " WHERE prodSKU <> '' " & WHERE_CLAUSE_UNITUPC
	SQL = SQL & " " & WHERE_CLAUSE_CASEUPC
	SQL = SQL & " " & WHERE_CLAUSE_INVENTORIEDITEM
	SQL = SQL & " " & WHERE_CLAUSE_PICKABLEITEM
	SQL = SQL & " " & WHERE_CLAUSE_CATEGORY
	SQL = SQL & " ORDER BY prodCategory, prodSKU ASC "

	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3
	rs.Open SQL, Session("ClientCnnString")
	

	Do While Not rs.EOF

		prodSKU = rs("prodSKU")
		prodDescription = rs("prodDescription")
		prodCategoryID = rs("prodCategory")
		
		If prodCategoryID <> "" Then
			prodCategoryName = GetCategoryByID(prodCategoryID)
		End If
		
		prodInventoriedItem = rs("prodInventoriedItem")

		If prodInventoriedItem = 1 OR prodInventoriedItem = vbtrue Then
			prodInventoriedItemDisplay = "YES"
		Else
			prodInventoriedItemDisplay = "NO"																		
		End If
		
		prodPickableItem = rs("prodPickableItem")

		If prodPickableItem = 1 OR prodPickableItem = vbtrue Then
			prodPickableItemDisplay = "YES"
		Else
			prodPickableItemDisplay = "NO"																		
		End If
		
		prodCasePricing = rs("prodCasePricing")
		prodCaseDescription = rs("prodCaseDescription")
		prodUnitUPC = rs("prodUnitUPC")
		prodCaseUPC = rs("prodCaseUPC")
		
		Response.Write("<tr>")
		Response.write("<td>" & prodSKU & "</td>")
		Response.write("<td>" & prodDescription & "</td>")
		Response.Write("<td>" & prodCategoryName & "</td>")
		Response.Write("<td align='center'>" & prodInventoriedItemDisplay & "</td>")
		Response.Write("<td align='center'>" & prodPickableItemDisplay & "</td>")
		Response.Write("<td align='center'>" & prodCasePricing & "</td>")
		Response.Write("<td>" & prodCaseDescription & "</td>")
		Response.Write("<td>" & prodUnitUPC & "</td>")
		Response.Write("<td>" & prodCaseUPC & "</td>")
		Response.Write("</tr>")

		rs.movenext

	Loop

	Response.Write("</tbody>")
	Response.Write("</table>")		
	Response.Write("</div>")
	
%>


	</table>
  </div>
<!-- eof responsive tables !-->



<!-- eof row !-->

<!-- row !-->
<div class="row">

<div class="col-lg-12"><hr></div>
</div>
<!-- eof row !-->

<!-- row !-->
<div class="row">

<%		

	rs.Close	
		
%>


</div>
<!-- eof row !-->


<!--#include file="../../inc/footer-main.asp"-->