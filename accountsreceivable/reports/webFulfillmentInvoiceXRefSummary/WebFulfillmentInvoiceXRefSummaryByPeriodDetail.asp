<%
Server.ScriptTimeout = 900000 'Default value
%>
<!--#include file="../../../inc/header.asp"-->
<!--#include file="../../../inc/jquery_table_search.asp"-->
<%

PeriodIntRecID = Request.QueryString("p")
	
PeriodStartDate = GetPeriodBeginDateByIntRecID(PeriodIntRecID)
PeriodEndDate = GetPeriodEndDateByIntRecID(PeriodIntRecID)
PeriodNum = GetPeriodByIntRecID(PeriodIntRecID)
PeriodYear = GetPeriodYearByIntRecID(PeriodIntRecID)

CustClassCode = Request.QueryString("cc")
CustClassCodeName = GetCustomerClassNameByID(CustClassCode)


CreateAuditLogEntry "Report","Report","Minor",0, MUV_Read("DisplayName") & " ran the report: Web Fulfillment Invoice Cross Reference By Period Detail for Orders Placed By " & CustClassCodeName & " In Period " & PeriodNum & " of " & PeriodYear

	'**************************************************
	'Read Settings_Reports
	'**************************************************
	SQL = "SELECT * from Settings_Reports where ReportNumber = 1900 AND UserNo = " & Session("userNo")
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	Set rs= cnn8.Execute(SQL)
	If NOT rs.EOF Then
	
		UseSettings_Reports = True
		
		OCSWebOrderOrMDSInvoice = rs("ReportSpecificData1")
		DatesOrPeriods = "Dates"
		StartPeriodBeingEvaluatedCustomize = PeriodStartIntRecID
		EndPeriodBeingEvaluatedCustomize = PeriodEndIntRecID
		RangeStartDateCustomize = ""
		RangeEndDateCustomize = ""
		DefaultSelectedCustomerClassesForInvoiceReport = rs("ReportSpecificData7")
		ShowOrdersWithRemarks = rs("ReportSpecificData8")
		ShowOrdersWithoutRemarks = rs("ReportSpecificData9")
		ShowOrdersThatAreInvoiced = rs("ReportSpecificData10")
		ShowOrdersThatAreNotInvoiced = rs("ReportSpecificData11")
		ShowOrdersThatAreHidden = rs("ReportSpecificData12")
		DefaultSelectedCustomerTypesForInvoiceReport = rs("ReportSpecificData13")
		If IsNull(OCSWebOrderOrMDSInvoice) Then OCSWebOrderOrMDSInvoice = ""
		If IsNull(DatesOrPeriods) Then DatesOrPeriods = "Periods"

		
		If IsNull(DefaultSelectedCustomerClassesForInvoiceReport) Then DefaultSelectedCustomerClassesForInvoiceReport = ""
		
		If IsNull(DefaultSelectedCustomerTypesForInvoiceReport) Then DefaultSelectedCustomerTypesForInvoiceReport = ""
		
		If IsNull(ShowOrdersWithRemarks) OR ShowOrdersWithRemarks = "false" Then 
			ShowOrdersWithRemarks = 0
		ElseIf ShowOrdersWithRemarks = "true" Then
			ShowOrdersWithRemarks = 1
		End If

		If IsNull(ShowOrdersWithoutRemarks) OR ShowOrdersWithoutRemarks = "false" Then 
			ShowOrdersWithoutRemarks = 0
		ElseIf ShowOrdersWithoutRemarks = "true" Then
			ShowOrdersWithoutRemarks = 1
		End If

		If IsNull(ShowOrdersThatAreInvoiced) OR ShowOrdersThatAreInvoiced = "false" Then 
			ShowOrdersThatAreInvoiced = 0
		ElseIf ShowOrdersThatAreInvoiced = "true" Then
			ShowOrdersThatAreInvoiced = 1
		End If
		
		If IsNull(ShowOrdersThatAreNotInvoiced) OR ShowOrdersThatAreNotInvoiced = "false" Then 
			ShowOrdersThatAreNotInvoiced = 0
		ElseIf ShowOrdersThatAreNotInvoiced = "true" Then
			ShowOrdersThatAreNotInvoiced = 1
		End If

		If IsNull(ShowOrdersThatAreHidden) OR ShowOrdersThatAreHidden = "false" Then 
			ShowOrdersThatAreHidden = 0
		ElseIf ShowOrdersThatAreHidden = "true" Then
			ShowOrdersThatAreHidden = 1
		End If
		
	Else
	
		UseSettings_Reports = False
		OCSWebOrderOrMDSInvoice = ""
		DatesOrPeriods = "Periods"

		StartPeriodBeingEvaluatedCustomize = PeriodStartIntRecID
		EndPeriodBeingEvaluatedCustomize = PeriodEndIntRecID

		RangeStartDateCustomize= ""
		RangeEndDateCustomize = ""	
		DefaultSelectedCustomerClassesForInvoiceReport = ""
		DefaultSelectedCustomerTypesForInvoiceReport = ""
		
		'-------------------------------------------------------------------------
		'As per Warren, Default View Should Include Orders That Are Hidden,
		'Orders Invoiced and Un-Invoiced, Orders With and Without Remarks
		'-------------------------------------------------------------------------
				
		ShowOrdersWithRemarks = 1
		ShowOrdersWithoutRemarks = 1
		ShowOrdersThatAreInvoiced = 1
		ShowOrdersThatAreNotInvoiced = 1
		ShowOrdersThatAreHidden = 0
		'-------------------------------------------------------------------------
		'-------------------------------------------------------------------------
	End If									
	'**************************************************
	'End Read Settings_Reports
	'**************************************************
	
	
	
	'**************************************************
	'Start Build Page Header Text
	'**************************************************
	
	'----------------------------------------------------------------------
	'Check for customization by OCSAccessOrderDate (By Period)
	'If no customization, use current periods this year
	'----------------------------------------------------------------------	

	If OCSWebOrderOrMDSInvoice = "OCS" Then
	
		PageHeaderText = "Web Fulfillment and Invoice Cross Reference By Period <br>for OCS Web Orders Placed By " & CustClassCodeName & " In Period " & PeriodNum & " of " & PeriodYear
		PageHeaderText = PageHeaderText & " (" & FormatDateTime(PeriodStartDate,2) & " - " & FormatDateTime(PeriodEndDate,2) & ")&nbsp;&nbsp;"
		
	ElseIf OCSWebOrderOrMDSInvoice = "MDS" Then
	
		PageHeaderText = "Web Fulfillment and Invoice Cross Reference By Period <br>for MDS Invoices Placed By " & CustClassCodeName & " In Period " & PeriodNum & " of " & PeriodYear			
		PageHeaderText = PageHeaderText & " (" & FormatDateTime(PeriodStartDate,2) & " - " & FormatDateTime(PeriodEndDate,2) & ")&nbsp;&nbsp;"
			
	Else
		PageHeaderText = "Web Fulfillment and Invoice Cross Reference By Period <br>for All Orders Placed By " & CustClassCodeName & " In Period " & PeriodNum & " of " & PeriodYear
		PageHeaderText = PageHeaderText & " (" & FormatDateTime(PeriodStartDate,2) & " - " & FormatDateTime(PeriodEndDate,2) & ")&nbsp;&nbsp;"
		
	End If
	'**************************************************
	'End Build Page Header Text
	'**************************************************
	

	'**************************************************************************************
	'Build WHERE Clause For Orders With/Without Remarks
	'**************************************************************************************
	
	WHERE_CLAUSE_REMARKS = ""
	
	If ShowOrdersWithRemarks = "1" AND ShowOrdersWithoutRemarks = "0" Then
		WHERE_CLAUSE_REMARKS = " AND (Remarks <> '')"
	ElseIf ShowOrdersWithRemarks = "0" AND ShowOrdersWithoutRemarks = "1" Then 
		WHERE_CLAUSE_REMARKS = " AND (Remarks = '')"				
	End If
	

	'**************************************************************************************
	'Build WHERE Clause For Orders That Are/Are Not Invoiced
	'**************************************************************************************

	WHERE_CLAUSE_INVOICED = ""
	
	If ShowOrdersThatAreInvoiced = "1" AND ShowOrdersThatAreNotInvoiced = "0" Then
		WHERE_CLAUSE_INVOICED = " AND (MDSInvoiceID <> '')"
	ElseIf ShowOrdersThatAreInvoiced = "0" AND ShowOrdersThatAreNotInvoiced = "1" Then 
		WHERE_CLAUSE_INVOICED = " AND (MDSInvoiceID = '')"			
	End If
	
	'**************************************************************************************
	'Build WHERE Clause For Orders That Are Hidden
	'**************************************************************************************

	WHERE_CLAUSE_HIDDEN = ""
	
	If ShowOrdersThatAreHidden = "1" Then
		WHERE_CLAUSE_HIDDEN = " AND (DontIncludeOnReport = 1 OR DontIncludeOnReport = 0)"
	ElseIf ShowOrdersThatAreHidden = "0" Then 
		WHERE_CLAUSE_HIDDEN = " AND (DontIncludeOnReport = 0)"				
	End If
	
	'**************************************************************************************
	'If Customer Type is empty from report settings, obtain all customer
	'types from AR_Customer and CustomerType
	'**************************************************************************************

	CustomerTypeArray = ""
	CustomerTypeArray = Split(DefaultSelectedCustomerTypesForInvoiceReport,",")
	

	If UBound(CustomerTypeArray) < 0 Then
	
		CustomerTypeArrayString = ""
		
		Set cnnGetAllValidCustomerTypes = Server.CreateObject("ADODB.Connection")
		cnnGetAllValidCustomerTypes.open Session("ClientCnnString")
	
		resultGetAllValidCustomerTypes = ""
			
		SQLGetAllValidCustomerTypes = "SELECT DISTINCT(CustType) FROM AR_Customer ORDER BY CustType"
		 
		Set rsGetAllValidCustomerTypes = Server.CreateObject("ADODB.Recordset")
		rsGetAllValidCustomerTypes.CursorLocation = 3 
		Set rsGetAllValidCustomerTypes= cnnGetAllValidCustomerTypes.Execute(SQLGetAllValidCustomerTypes)
		
		If NOT rsGetAllValidCustomerTypes.EOF Then 
		
			Do While NOT rsGetAllValidCustomerTypes.EOF
				CustomerTypeArrayString = CustomerTypeArrayString & rsGetAllValidCustomerTypes("CustType") & ","
				rsGetAllValidCustomerTypes.MoveNext
			Loop
				
			If Right(CustomerTypeArrayString,1) = "," Then 
				CustomerTypeArrayString = left(CustomerTypeArrayString,Len(CustomerTypeArrayString)-1)
			End If
			
			CustomerTypeArray = Split(CustomerTypeArrayString,",")

		End If
	
		rsGetAllValidCustomerTypes.Close
		set rsGetAllValidCustomerTypes= Nothing
		cnnGetAllValidCustomerTypes.Close	
		set cnnGetAllValidCustomerTypes = Nothing	
	
	End If
	
	'**************************************************************************************
	'End Build Customer Type Array
	'**************************************************************************************


	
	
	
	%>

<script type="text/javascript">

	$(document).ready(function() {
	
	    $("#PleaseWaitPanel").hide();
	    
		$('#myWebOrdersModal').on('show.bs.modal', function(e) {
		
           	 $("#btnDeleteRemark").show();
           	 $("#btnEditRemarkSave").show();
           	 $("#btnEditRemarkClose").html('Close');

		    //get data-id attribute of the clicked order
		    var InternalRecordIdentifier = $(e.relatedTarget).data('intrecid');
		    var CustID = $(e.relatedTarget).data('custid');
		    var WebOrderID = $(e.relatedTarget).data('web-order-num');
		    var WebOrderDate = $(e.relatedTarget).data('web-order-date');
	
		    //populate the textbox with the id of the clicked order
		    $(e.currentTarget).find('input[name="txtIntRecID"]').val(InternalRecordIdentifier);
		    $(e.currentTarget).find('input[name="txtCID"]').val(CustID);
		    	    
		    var $modal = $(this);
	
    		$modal.find('#myWebOrdersLabel').html("Order Remarks for Order #" + WebOrderID + " placed on " + WebOrderDate);
    		
	    	$.ajax({
				type:"POST",
				url: "../../../inc/InSightFuncs_AjaxForInvoicingModals.asp",
				cache: false,
				data: "action=GetContentForWebOrderRemarksModal&InternalRecordIdentifier="+encodeURIComponent(InternalRecordIdentifier)+"&CustID="+encodeURIComponent(CustID),
				success: function(response)
				 {
	               	 $modal.find('#webOrderRemarksModalContent').html(response);
	             },
	             failure: function(response)
				 {
				   $modal.find('#webOrderRemarksModalContent').html("Failed");
	             }
			});
		    
		});
	    
	    
		// hide web order from appearing in report
		$('[name="chkDontIncludeOnReport"]').click(function () {	

			var checkboxID = this.id;
			
			// if the checkbox was checked, then update the SQL record
			// and hide web order from appearing in report (only if hide order setting is not set)
			if ($("#" + checkboxID).is(':checked')) {
			
					var checkboxIDLength = checkboxID.length;
					var InternalRecordIdentifier = checkboxID.substring(22, checkboxIDLength+1);
					var showOrdersThatAreHiddenReportSetting = $("#txtShowOrdersThatAreHidden").val();
					
					//alert(checkboxID + "---" + checkboxIDLength + "---" + InternalRecordIdentifier);
			    	
			    	$.ajax({
						type:"POST",
						url: "../../../inc/InSightFuncs_AjaxForInvoicingModals.asp",
						cache: false,
						data: "action=DoNotShowWebFulfillmentOrder&InternalRecordIdentifier="+encodeURIComponent(InternalRecordIdentifier),
						success: function(response)
						 {	
						 	if (showOrdersThatAreHiddenReportSetting == "0") {
			               	 	//This will temporarily hide the order row from view until the page refreshes
			               	 	$('#' + InternalRecordIdentifier).hide(); 
			               	 }              	 
			             },
			             failure: function(response)
						 {
			             }
					});	
			}
			
			// if the checkbox was UN-checked, then update the SQL record			
			else {
			
					var checkboxIDLength = checkboxID.length;
					var InternalRecordIdentifier = checkboxID.substring(22, checkboxIDLength+1);
					var showOrdersThatAreHiddenReportSetting = $("#txtShowOrdersThatAreHidden").val();
					
			    	$.ajax({
						type:"POST",
						url: "../../../inc/InSightFuncs_AjaxForInvoicingModals.asp",
						cache: false,
						data: "action=ShowWebFulfillmentOrder&InternalRecordIdentifier="+encodeURIComponent(InternalRecordIdentifier),
						success: function(response)
						 {	                  	 
			             },
			             failure: function(response)
						 {
			             }
					});	
			}
		});	
		
	    
	    
	});

</script>

	<style>
	.form-control[disabled], .form-control[readonly], fieldset[disabled] .form-control{
		background-color:#fff;
		border: 1px solid #eee;
	}
	
	.invoicerangedatepicker {
		position: absolute;
		bottom: 25px;
		right: 24px;
		top: auto;
		cursor: pointer;
	}
	
	.activefilter {
	    background: #f0ad4e !important;
	}
		

	.modal-footer {
	    /*padding: 0px !important;*/
	    text-align: right !important;
	    border-top: 0px !important;
	}
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


	.bs-example-modal-lg-customize .row{
		margin-bottom: 10px;
	 	width: 100%;
		overflow: hidden;
	}

	.bs-example-modal-lg-customize .left-column{
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

	.filter-search-width{
		max-width: 36%;
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
	
	markY {
	    background-color: yellow;
	    color: black;
	} 
	
</style>


<%
Response.Write("<div id=""PleaseWaitPanel"">")
Response.Write("<br><br>Processing Web Fulfillment and Invoice Cross Reference By Period Detail Summary Data<br><br>This may take up to a full minute, please wait...<br><br>")
Response.Write("<img src=""../../img/loading.gif"" />")
Response.Write("</div>")
Response.Flush()


Set cnnMasterWebFulfillment = Server.CreateObject("ADODB.Connection")
cnnMasterWebFulfillment.open (Session("ClientCnnString"))
Set rsMasterWebFulfillment = Server.CreateObject("ADODB.Recordset")
rsMasterWebFulfillment.CursorLocation = 3 

'**************************************************************************************
'Begin Build SQL STMT To Select From IN_WebFulfillment
'**************************************************************************************	

SQLMasterWebFulfillment = "SELECT InternalRecordIdentifier, RecordCreationDateTime, OCSAccessOrderID, "
SQLMasterWebFulfillment = SQLMasterWebFulfillment & " OCSAccessOrderDate, CustID, CustClassCode, CustTypeNum, MDSInvoiceID, "
SQLMasterWebFulfillment = SQLMasterWebFulfillment & " MDSInvoiceDate, OCSAccessMerchTotal, MDSInvoiceTotal, DontIncludeOnReport, Remarks, OCSAccessOrderComments "
SQLMasterWebFulfillment = SQLMasterWebFulfillment & " FROM IN_WebFulfillment "

'----------------------------------------------------------------------
'Check for customization by OCSAccessOrderDate (Periods or Dates)
'If no customization, use current period to date
'----------------------------------------------------------------------	
If OCSWebOrderOrMDSInvoice = "OCS" Then
	SQLMasterWebFulfillment = SQLMasterWebFulfillment & " WHERE (OCSAccessOrderDate BETWEEN '" & PeriodStartDate & "' AND '" & PeriodEndDate & "') "
ElseIf OCSWebOrderOrMDSInvoice = "MDS" Then
	SQLMasterWebFulfillment = SQLMasterWebFulfillment & " WHERE (MDSInvoiceDate BETWEEN '" & PeriodStartDate & "' AND '" & PeriodEndDate & "') "
Else
	SQLMasterWebFulfillment = SQLMasterWebFulfillment & " WHERE (OCSAccessOrderDate BETWEEN '" & PeriodStartDate & "' AND '" & PeriodEndDate & "') "
End If

'---------------------------------------------------------------------------
'Check for customization by Customer Class, Remarks, Invoiced, and Hidden
'---------------------------------------------------------------------------	
If WHERE_CLAUSE_REMARKS <> "" Then SQLMasterWebFulfillment = SQLMasterWebFulfillment & WHERE_CLAUSE_REMARKS
If WHERE_CLAUSE_INVOICED <> "" Then SQLMasterWebFulfillment = SQLMasterWebFulfillment & WHERE_CLAUSE_INVOICED
If WHERE_CLAUSE_HIDDEN <> "" Then SQLMasterWebFulfillment = SQLMasterWebFulfillment & WHERE_CLAUSE_HIDDEN

SQLMasterWebFulfillment = SQLMasterWebFulfillment & " AND (CustClassCode = '" & CustClassCode & "') "

If OCSWebOrderOrMDSInvoice = "OCS" OR OCSWebOrderOrMDSInvoice = "" Then
	SQLMasterWebFulfillment = SQLMasterWebFulfillment & " ORDER BY OCSAccessOrderDate DESC"
ElseIf OCSWebOrderOrMDSInvoice = "MDS" Then
	SQLMasterWebFulfillment = SQLMasterWebFulfillment & " ORDER BY MDSInvoiceDate DESC"
End If

'**************************************************************************************
'END Build SQL STMT To Select From IN_WebFulfillment
'**************************************************************************************	

'Response.write("<br><br>" & SQLMasterWebFulfillment)

Set rsMasterWebFulfillment = cnnMasterWebFulfillment.Execute(SQLMasterWebFulfillment)

%> 
<h3 class="page-header">

	<a href="<%= BaseURL %>accountsreceivable/reports/WebFulfillmentInvoiceXRefSummaryByPeriod.asp"><button type="button" class="btn btn-success"><i class="fa fa-arrow-left" aria-hidden="true"></i> Back To All Periods</button></a><br><br>

	<i class="fa fa-file-text" aria-hidden="true"></i> 
	<%= PageHeaderText %>

</h3>

	<h6 class="page-header">
	<table id="table-search" class='table table-striped table-condensed table-hover display'>
	<tr>
		<td width="20%">
			<% If UseSettings_Reports = True Then
				Response.Write("<span class='markY'>" & "Using Saved Customization Values</br>" & "</span>")
			End If %>
			
			<% If ShowOrdersThatAreHidden = 1 Then %>
				Include Hidden Orders is <strong>ON</strong><br>
			<% End If %>
			
			<% If ShowOrdersThatAreHidden = 0 Then %>
				Include Hidden Orders is <strong>OFF</strong><br>
			<% End If %>
			
			<% If ShowOrdersThatAreInvoiced = 1 Then %>
				Include Orders That Are Invoiced is <strong>ON</strong><br>
			<% End If %>
			
			<% If ShowOrdersThatAreNotInvoiced = 1 Then %>
				Include Orders That Are NOT Invoiced is <strong>ON</strong><br>
			<% End If %>
			
			<% If ShowOrdersThatAreInvoiced = 0 AND ShowOrdersThatAreNotInvoiced = 0 Then %>
				Include Both Invoiced and Non-Invoiced Orders is <strong>ON</strong><br>
			<% End If %>
			
			<% If ShowOrdersWithRemarks = 1 Then %>
				Include Orders With Remarks is <strong>ON</strong><br>
			<% End If %>
			
			<% If ShowOrdersWithoutRemarks = 1 Then %>
				Include Orders Without Remarks <strong>ON</strong><br>
			<% End If %>
			
			<% If ShowOrdersWithRemarks = 0 AND ShowOrdersWithoutRemarks = 0 Then %>
				Include Orders With and Without Remarks is <strong>ON</strong><br>
			<% End If %>
		
		</td>
		<td>
		
			<% If OCSWebOrderOrMDSInvoice = "" Then %>
			
				Filter Orders By Date <strong>OFF</strong><br>
				
			<% ElseIf OCSWebOrderOrMDSInvoice = "OCS" Then %>
		
				Filter Orders By OCS Web Order Date Within a Period Range is <strong>ON</strong><br>
				Showing OCS Web Orders from <strong>Period <%= PeriodStartNum %> of  <%= PeriodStartYear %></strong> to <strong>Period <%= PeriodEndNum %> of <%= PeriodEndYear %></strong><br>
				These periods span the dates: <strong><%= FormatDateTime(PeriodStartDate,2) %> to <%= FormatDateTime(PeriodEndDate,2) %></strong><br>
		
			<% ElseIf OCSWebOrderOrMDSInvoice = "MDS" Then %>
		
				Filter Orders By MDS Invoice Date Within a Period Range is <strong>ON</strong><br>
				Showing MDS Invoiced Orders from <strong>Period <%= PeriodStartNum %> of  <%= PeriodStartYear %></strong> to <strong><%= PeriodEndNum %> of <%= PeriodEndYear %></strong><br>
				These periods span the dates: <strong><%= FormatDateTime(PeriodStartDate,2) %> to <%= FormatDateTime(PeriodEndDate,2) %></strong><br>
		
			<% End If %>
		
		</td>
		
		<td>
			<% For z = 0 to UBound(CustomerTypeArray)
					If CustomerTypeArray(z) <> "" Then
						currentCustType = CustomerTypeArray(z)
						%>Customer Type <strong><%= GetCustTypeByCustTypeNum(currentCustType) %></strong> - (Cust Type <%= currentCustType %>)<br><%
					End If
			   Next
			%>
		</td>

	
	</tr>
	</table>
	</h6>
	

<!-- row !-->
<div class="row">

<!-- responsive tables !-->
<div class="table-responsive">
	
<div class="input-group"> <span class="input-group-addon">Narrow Results</span>

    <input id="filter" type="text" class="form-control filter-search-width" placeholder="Type here...">
</div>

	<input type="hidden" name="txtShowOrdersThatAreHidden" id="txtShowOrdersThatAreHidden" value="<%= ShowOrdersThatAreHidden %>">
    <table id="tableSuperSum" class="food_planner sortable table table-striped table-condensed table-hover">
      <thead>
        <tr>
          <th class="sorttable numeric">Web Order #</th>
          <th class="sorttable">Date Submitted</th>
          <th class="sorttable numeric">Invoice Number</th> 
          <th class="sorttable">Invoice Date</th>
          <th class="sorttable numeric">Order Amt</th>
          <th class="sorttable numeric">Invoice Amt</th> 
          <th class="sorttable numeric">Fulfillment Rate</th>
          <th class="sorttable">Remarks</th>
          <th class="sorttable">Web Comments</th>
          <th>Don't Include</th>
        </tr>
      </thead>
<%		
		Response.Write("<tbody class='searchable'>")
		
		Do While Not rsMasterWebFulfillment.EOF
					
			InternalRecordIdentifier = rsMasterWebFulfillment("InternalRecordIdentifier")
			RecordCreationDateTime = rsMasterWebFulfillment("RecordCreationDateTime") 
			OCSAccessOrderID = rsMasterWebFulfillment("OCSAccessOrderID")	
			OCSAccessOrderDate = rsMasterWebFulfillment("OCSAccessOrderDate")
			MDSInvoiceID = rsMasterWebFulfillment("MDSInvoiceID") 
			MDSInvoiceDate = rsMasterWebFulfillment("MDSInvoiceDate") 
			OCSAccessMerchTotal = rsMasterWebFulfillment("OCSAccessMerchTotal")
			MDSInvoiceTotal = rsMasterWebFulfillment("MDSInvoiceTotal")
			DontIncludeOnReport = rsMasterWebFulfillment("DontIncludeOnReport")
			Remarks = rsMasterWebFulfillment("Remarks")
			OCSAccessOrderComments = rsMasterWebFulfillment("OCSAccessOrderComments")
			CustClassCode = rsMasterWebFulfillment("CustClassCode")
			CustClassCode = GetCustomerClassNameByID(CustClassCode)
			CustTypeNum = rsMasterWebFulfillment("CustTypeNum")
			CustTypeName = GetCustTypeByCustTypeNum(CustTypeNum)
			
			If MDSInvoiceDate = "1/1/1900" Then 
				MDSInvoiceDate = "---"
				MDSInvoiceID = "---"
				FulfillmentRate = "---"
			Else
				If OCSAccessMerchTotal > 0 Then
					If MDSInvoiceTotal > OCSAccessMerchTotal Then
						FulfillmentRate = FormatPercent(1,2)
					Else
						FulfillmentRate = FormatPercent(MDSInvoiceTotal / OCSAccessMerchTotal,2)
					End If
				Else
					FulfillmentRate = FormatPercent(0,2)
				End If
			End If

			If MDSInvoiceTotal <> "" Then
				If rsMasterWebFulfillment("MDSInvoiceDate") = "1/1/1900" Then
					MDSInvoiceTotal = "---"
				Else
					MDSInvoiceTotal = FormatCurrency(MDSInvoiceTotal,2)
				End If
			End If

			
			If OCSAccessOrderDate <> "1/1/1900" AND OCSAccessOrderDate <> "" Then
				OCSAccessOrderDateSortKey = OCSAccessOrderDate
				OCSAccessOrderDate = FormatDateTime(OCSAccessOrderDate,1)
			Else
				OCSAccessOrderDateSortKey = "1/1/1900"
			End If
			
			If MDSInvoiceDate <> "1/1/1900" AND MDSInvoiceDate <> "" AND MDSInvoiceDate <> "---" Then
				MDSInvoiceDateSortKey = MDSInvoiceDate
				MDSInvoiceDate = FormatDateTime(MDSInvoiceDate,1)
			Else
				MDSInvoiceDateSortKey = "1/1/1900"
			End If
			
			If OCSAccessMerchTotal <> "" Then
				OCSAccessMerchTotal = FormatCurrency(OCSAccessMerchTotal,2)
			End If
			
			If MDSInvoiceTotal <> "" AND MDSInvoiceTotal <> "---" Then
				MDSInvoiceTotal = FormatCurrency(MDSInvoiceTotal,2)
			End If
			
			%>
				<tr id="<%= InternalRecordIdentifier %>">
					<td><%= OCSAccessOrderID %></td>
					<td sorttable_customkey="<%= OCSAccessOrderDateSortKey %>"><%= OCSAccessOrderDate %></td>
					<td><%= MDSInvoiceID %></td>
					<td sorttable_customkey="<%= MDSInvoiceDateSortKey %>"><%= MDSInvoiceDate %></td>
					<td><%= OCSAccessMerchTotal %></td>
					<td><%= MDSInvoiceTotal %></td>
					<td><%= FulfillmentRate %></td>
					<td><button class="btn btn-success btn-xs" data-title="Edit" data-toggle="modal" data-target="#myWebOrdersModal" href="#" data-intrecid="<%= InternalRecordIdentifier %>" data-custid="<%= CustID %>" data-web-order-num="<%= OCSAccessOrderID %>" data-web-order-date="<%= OCSAccessOrderDate %>"><span class="glyphicon glyphicon-pencil"></span></button>&nbsp;<span id="remarks<%= InternalRecordIdentifier %>"><%= Remarks %></span></td>
					<td width="15%"><%= OCSAccessOrderComments %></td>
					<td align="center"><input type="checkbox" name="chkDontIncludeOnReport" id="chkDontIncludeOnReport<%= InternalRecordIdentifier %>" <% If DontIncludeOnReport = 1 Then Response.write("checked") %>></td>
				</tr>
			<%
			
			rsMasterWebFulfillment.movenext
				
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
</div>
<!-- eof row !-->



<!-- **************************************************************************************************************************** -->
<!-- MODAL FOR ORDERS BEGIN HERE !-->
<!-- **************************************************************************************************************************** -->


<div class="modal fade" id="myWebOrdersModal" tabindex="-1" role="dialog" aria-labelledby="myWebOrdersLabel">

	<div class="modal-dialog" role="document">
						
		<div class="modal-content">	
	
			<script type="text/javascript">
			
				$(document).ready(function() {
				
					$("#btnEditRemarkSave").on("click", function(e) {
					
					    var InternalRecordIdentifier = $("#txtIntRecID").val();
					    var CustID = $("#txtCID").val();	
					    var WebOrderRemarks = $("#txtWebOrderRemarks").val();	
		
				    	$.ajax({
							type:"POST",
							url: "../../../inc/InSightFuncs_AjaxForInvoicingModals.asp",
							cache: false,
							data: "action=EditWebOrderRemarksFromModal&InternalRecordIdentifier="+encodeURIComponent(InternalRecordIdentifier)+"&CustID="+encodeURIComponent(CustID)+"&WebOrderRemarks="+encodeURIComponent(WebOrderRemarks),
							success: function(response)
							 {
				               	 $("#webOrderRemarksModalContent").html("<div class='alert alert-success'><strong>Success!</strong> Remarks Updated.</div>");
				               	 $("#remarks" + InternalRecordIdentifier).html(WebOrderRemarks);
				               	 $("#btnDeleteRemark").hide();
				               	 $("#btnEditRemarkSave").hide();
				               	 $("#btnEditRemarkClose").html('Close Window');
				             },
				             failure: function(response)
							 {
							   $("#webOrderRemarksModalContent").html("<div class='alert alert-danger'><strong>Error</strong> Remarks Failed to Save.</div>");
				             }
						});
					    
					});
			
				
					$("#btnDeleteRemark").on("click", function(e) {
					
					    var InternalRecordIdentifier = $("#txtIntRecID").val();
					    var CustID = $("#txtCID").val();	
		
				    	$.ajax({
							type:"POST",
							url: "../../../inc/InSightFuncs_AjaxForInvoicingModals.asp",
							cache: false,
							data: "action=DeleteWebOrderRemarksFromModal&InternalRecordIdentifier="+encodeURIComponent(InternalRecordIdentifier)+"&CustID="+encodeURIComponent(CustID),
							success: function(response)
							 {
				               	 $("#webOrderRemarksModalContent").html("<div class='alert alert-success'><strong>Success!</strong> Remarks Deleted.</div>");
				               	 $("#remarks" + InternalRecordIdentifier).html("");
				               	 $("#btnDeleteRemark").hide();
				               	 $("#btnEditRemarkSave").hide();
				               	 $("#btnEditRemarkClose").html('Close Window');
				             },
				             failure: function(response)
							 {
							   $("#webOrderRemarksModalContent").html("<div class='alert alert-danger'><strong>Error</strong> Remarks Failed to Delete.</div>");
				             }
						});
					    
					});
	    
				});
			
			</script>
    
			<!-- modal header !-->
			<div class="modal-header">
				<button type="button" class="close" data-dismiss="modal" aria-label="Close"><span aria-hidden="true">&times;</span></button>
				<h4 class="modal-title" id="myWebOrdersLabel"></h4>
			</div>
			<!-- eof modal header !-->
	  
			<!-- modal body !-->
			<div class="modal-body">
				<input type="hidden" name="txtIntRecID" id="txtIntRecID" value="">
				<input type="hidden" name="txtCID" id="txtCID" value="">
				<div id="webOrderRemarksModalContent">
					<!-- Content for the modal will be generated and written here -->
					<!-- Content generated by Sub GetContentForWebOrderRemarksModal() in InsightFuncs_AjaxForInvoicingModals.asp -->
				</div>
			</div>
					  
			<!-- modal footer !-->
		    <div class="modal-footer"> 
				<!-- delete !-->
				<div class="col-lg-5" style="padding-top:30px">
					<button type="button" class="btn btn-danger btn-sm pull-left" id="btnDeleteRemark">Delete These Remarks</button>
				</div>
					      	      
				<!-- close / save !-->
				<div class="col-lg-7" style="padding-top:30px">
					<button type="button" class="btn btn-default btn-sm" data-dismiss="modal" id="btnEditRemarkClose">Close</button>
					<button type="button" class="btn btn-primary btn-sm" id="btnEditRemarkSave">Save Changes</button>
				</div>
				<!-- eof close / save !-->
			</div>
			<!-- eof modal footer !-->
	

		</div>
		<!-- eof modal content !-->
</div>
<!-- eof modal dialog !-->
</div>
<!-- eof modal !-->

<!-- **************************************************************************************************************************** -->
<!-- MODAL FOR ORDERS END HERE !-->
<!-- **************************************************************************************************************************** -->



<!--#include file="../../../inc/footer-main.asp"-->