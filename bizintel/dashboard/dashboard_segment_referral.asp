<%
Server.ScriptTimeout = 900000 'Default value

Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))

FilterReferral = Request.QueryString("p")


%>
<!--#include file="../../inc/header.asp"-->
<!--#include file="../../inc/jquery_table_search.asp"-->
<!--#include file="../../inc/InSightFuncs_BizIntel.asp"--> 
<!--#include file="../../inc/InSightFuncs_Equipment.asp"--> 

<%
CreateAuditLogEntry "Report","Report","Minor",0, MUV_Read("DisplayName") & " ran the report: Leakage Overview Single Referral Code"

PeriodBeingEvaluated = GetLastClosedPeriodAndYear()
PeriodSeqBeingEvaluated = GetLastClosedPeriodSeqNum()

ShowPercentageColumns = False

WorkDaysIn3PeriodBasis =  NumberOfWorkDays(GetPeriodBeginDateBySeq(PeriodSeqBeingEvaluated -3), GetPeriodEndDateBySeq(PeriodSeqBeingEvaluated -1))+1
WorkDaysIn12PeriodBasis =  NumberOfWorkDays(GetPeriodBeginDateBySeq(PeriodSeqBeingEvaluated -12), GetPeriodEndDateBySeq(PeriodSeqBeingEvaluated -1)) + 1 
WorkDaysInLastClosedPeriod = NumberOfWorkDays(GetPeriodBeginDateBySeq(PeriodSeqBeingEvaluated), GetPeriodEndDateBySeq(PeriodSeqBeingEvaluated )) + 1 
WorkDaysInCurrentPeriod = NumberOfWorkDays(GetPeriodBeginDateBySeq(PeriodSeqBeingEvaluated +1), GetPeriodEndDateBySeq(PeriodSeqBeingEvaluated +1)) + 1 
WorkDaysSoFar =  NumberOfWorkDays(GetPeriodBeginDateBySeq(PeriodSeqBeingEvaluated +1),Date()) + 1

%>

  
<style>

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

	.vpc-variance-header{
		background: #D43F3A;
		color:#fff;
		text-align:center;
		font-weight:bold;
	}
	
	.vpc-3pavg-header{
		background: #F0AD4E;
		color:#fff;
		text-align:center;
		font-weight:bold;
	}
	
	.vpc-lcp-header{
		background: #337AB7;
		color:#fff;
		text-align:center;
		font-weight:bold;
	}

	.vpc-current-header{
		background: #5CB85C;
		color:#fff;
		text-align:center;
		font-weight:bold;
	}

	.gen-info-header{
		background: #3B579D;
		color:#fff;
		text-align:center;
		font-weight:bold;
	}

	.negative{
		color:red;	
	}

	.neutral{
		font-weight:bold;
		color:black;
	}

	.smaller-header{
		font-size: 0.8em;
		vertical-align: top !important;
		text-align: center;
	}	

	.smaller-detail-line{
		font-size: 0.8em;
	}	

	.footer-total{
		font-size: 0.95em;
		vertical-align: top !important;
	}	
	
	.footer-total-negative
	{
		font-size: 1.5em;
		color:red;	
	}

	.modal.modal-wide .modal-dialog {
	  width: 50%;
	}
	.modal-wide .modal-body {
	  overflow-y: auto;
	}
	
	.modal.modal-xwide .modal-dialog {
	  width: 70%;
	}
	.modal-xwide .modal-body {
	  overflow-y: auto;
	  max-height:600px;
	}
	
	.bs-example-modal-lg-customize .row{
		margin-bottom: 10px;
	 	width: 100%;
		overflow: hidden;
		font-size:11px;
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

</style>
<link rel="stylesheet" href="https://cdn.datatables.net/1.10.16/css/jquery.dataTables.min.css" />
<script type="text/javascript" src="https://cdn.datatables.net/1.10.16/js/jquery.dataTables.min.js"></script>
<script type="text/javascript" src="https://cdn.datatables.net/plug-ins/1.10.18/sorting/currency.js"></script>
<script type="text/javascript">

$.fn.dataTable.ext.type.order['currency-pre'] = function ( data ) {
   
    var expression = /((\(\$))|(\$\()/g;
    //Check if its in the proper format
    if(data.match(expression)){
        //It matched - strip out parentheses & any characters we dont want and append - at front     
        data = '-' + data.replace(/[\$\(\),]/g,'');
    }else{
      data = data.replace(/[\$\,]/g,'');
    }
    return parseInt( data, 10 );
};

$(document).ready(function() {



	$('#modalEquipmentVPC').on('show.bs.modal', function(j) {

	    //get data-id attribute of the clicked order
	    var CustID = $(j.relatedTarget).data('cust-id');
	    var LCPGP = $(j.relatedTarget).data('lcp-gp');
 
	    //populate the textbox with the id of the clicked order
	    $(j.currentTarget).find('input[name="txtCustIDToPass"]').val(CustID);
	    $(j.currentTarget).find('input[name="txtLastClosedPeriodGP"]').val(LCPGP);
	    	    
	    var $modal = $(this);
 

    	$.ajax({
			type:"POST",
			url: "../../inc/InSightFuncs_AjaxForBizIntelModals.asp",
			cache: false,
			data: "action=GetTitleForEquipmentVPCModal&CustID="+encodeURIComponent(CustID)+"&LCPGP="+encodeURIComponent(LCPGP),
			success: function(response)
			 {
               	 $modal.find('#modalEquipmentVPCTitle').html(response);            	 
             },
             failure: function(response)
			 {
			  	$modal.find('#modalEquipmentVPCTitle').html("Failed");
             }
		});
		

    	$.ajax({
			type:"POST",
			url: "../../inc/InSightFuncs_AjaxForBizIntelModals.asp",
			cache: false,
			data: "action=GetContentForEquipmentVPCModal&CustID="+encodeURIComponent(CustID),
		  	beforeSend: function() {
		     	$('#PleaseWaitPanelModal').show();
		     	$modal.find('#modalCategoryVPCContent').html('');
		  	},
		  	complete: function(){
		     	$('#PleaseWaitPanelModal').hide();
		  	},
			success: function(response)
			 {
               	 $("#PleaseWaitPanelModal").hide();
               	 $modal.find('#modalCategoryVPCContent').html(response);               	 
             },
             failure: function(response)
			 {
			  	$modal.find('#modalCategoryVPCContent').html("Failed");
             }
		});
	});
	
	
	

	$('#modalEquipmentVPC').on('show.bs.modal', function(j) {

	    //get data-id attribute of the clicked order
	    var CustID = $(j.relatedTarget).data('cust-id');
	    var LCPGP = $(j.relatedTarget).data('lcp-gp');
 
	    //populate the textbox with the id of the clicked order
	    $(j.currentTarget).find('input[name="txtCustIDToPass"]').val(CustID);
	    $(j.currentTarget).find('input[name="txtLastClosedPeriodGP"]').val(LCPGP);
	    	    
	    var $modal = $(this);
	    //$modal.find('#PleaseWaitPanelModal').show();  

    	$.ajax({
			type:"POST",
			url: "../../inc/InSightFuncs_AjaxForBizIntelModals.asp",
			cache: false,
			data: "action=GetTitleForEquipmentVPCModal&CustID="+encodeURIComponent(CustID)+"&LCPGP="+encodeURIComponent(LCPGP),
			success: function(response)
			 {
               	 $modal.find('#modalEquipmentVPCTitle').html(response);            	 
             },
             failure: function(response)
			 {
			  	$modal.find('#modalEquipmentVPCTitle').html("Failed");
             }
		});
		

    	$.ajax({
			type:"POST",
			url: "../../inc/InSightFuncs_AjaxForBizIntelModals.asp",
			cache: false,
			data: "action=GetContentForEquipmentVPCModal&CustID="+encodeURIComponent(CustID),
		  	beforeSend: function() {
		     	$('#PleaseWaitPanelModal').show();
		     	$modal.find('#modalCategoryVPCContent').html('');
		  	},
		  	complete: function(){
		     	$('#PleaseWaitPanelModal').hide();
		  	},
			success: function(response)
			 {
               	 $("#PleaseWaitPanelModal").hide();
               	 $modal.find('#modalCategoryVPCContent').html(response);               	 
             },
             failure: function(response)
			 {
			  	$modal.find('#modalCategoryVPCContent').html("Failed");
             }
		});
	});



    $("#PleaseWaitPanel").hide();
    
    $('#tableSuperSum').DataTable({
        scrollY: 500,
        scrollCollapse: true,
        paging: false,
        order: [ 2, 'asc' ],
  			columnDefs: [
		        { targets: 2, type: 'currency' }
		    ],
			 

			footerCallback: function ( row, data, start, end, display ) {
			            var api = this.api();
			            // Remove the formatting to get integer data for summation
			            var intVal = function ( i ) {

			                return typeof i === 'string' ?
			                    i.replace(/[\$,()]/g, '')*1 :
			                    typeof i === 'number' ?
			                        i : 0;
			            };
			  
			            // Total over all pages
			            var totalLCPv3PVar = api
			                .column( 2 )
			                .data()
			                .reduce( function (a, b) {
			                    return intVal(a) + intVal(b);
			                } );
			                
			            var totalLCPSales = api
			                .column( 8 )
			                .data()
			                .reduce( function (a, b) {
			                    return intVal(a) + intVal(b);
			                } );
			                
			            var total3PASalesCol = api
			                .column( 9 )
			                .data()
			                .reduce( function (a, b) {
			                    return intVal(a) + intVal(b);
			                } );

			            var totalCurrentSalesCol = api
			                .column( 11 )
			                .data()
			                .reduce( function (a, b) {
			                    return intVal(a) + intVal(b);
			                } );
			  
			            // Update footer			            
			            $("#totalLCPv3PVarCol").html('$'+ totalLCPv3PVar.toFixed(2).replace(/\d(?=(\d{3})+\.)/g, '$&,'))
			            $("#totalLCPSalesCol").html('$'+ totalLCPSales.toFixed(2).replace(/\d(?=(\d{3})+\.)/g, '$&,'))
			            $("#total3PASalesCol").html('$'+ total3PASalesCol.toFixed(2).replace(/\d(?=(\d{3})+\.)/g, '$&,'))
			            $("#totalCurrentSalesCol").html('$'+ totalCurrentSalesCol.toFixed(2).replace(/\d(?=(\d{3})+\.)/g, '$&,'))
			            
			            			            
			        }		        



   });
});
</script>


<%
Response.Write("<div id=""PleaseWaitPanel"" class=""container"">")
Response.Write("<br><br>Creating Leakage Overview Single Referral Code<br><br>Please wait...<br><br>")
Response.Write("<img src='" & baseURL & "/img/loading.gif'/>")
Response.Write("</div>")
Response.Flush()

%>





<h3 class="page-header">Leakage Overview for Referral Code <%=FilterReferral%> for Period <%=PeriodBeingEvaluated %>
&nbsp;&nbsp;
</h3>


<!-- row !-->
<div class="row">


<%
SQL = "SELECT Distinct CustCatPeriodSales_ReportData.CustNum,LCPTotSalesAllCats as LCPSales, Total3PPAvgAllCats, TotalCostAllCats, TotalTPLYAllCats "
SQL = SQL & ",Total3PPSalesAllCats AS ThreePPSales "
SQL = SQL & ", Total12PPSalesAllCats As TwelvePPSales "
SQL = SQL & " FROM CustCatPeriodSales_ReportData "
SQL = SQL & " WHERE ThisPeriodSequenceNumber = " & PeriodSeqBeingEvaluated 
SQL = SQL & " AND CustCatPeriodSales_ReportData.CustNum IN (SELECT CUSTNUM FROM AR_Customer "
SQL = SQL & " INNER JOIN Referal ON Referal.ReferalCode = AR_Customer.ReferalCode"
SQL = SQL & " WHERE Referal.Description2 = '" & FilterReferral & "') "

'Response.write(SQL)



Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.ConnectionTimeout = 120
cnn8.open (Session("ClientCnnString"))

Set rs = Server.CreateObject("ADODB.Recordset")

rs.CursorLocation = 3
Set rs = cnn8.Execute(SQL)
'Response.Write(now()&"<br>")
%>

 

<!-- responsive tables !-->

<!--	
<div class="input-group"> <span class="input-group-addon">Narrow Results</span>
    <input id="filter" type="text" class="form-control filter-search-width" placeholder="Type here...">
</div><br>
!-->
<div class="container-fluid">
    <div class="row">
           <table id="tableSuperSum" class="display  compact" style="width:100%;">
              <thead>
                  <tr>	
					<th rowspan="2"  class="sorttable numeric smaller-header"><br>Acct</th>
					<th rowspan="2"  class="sorttable numeric smaller-header"><br>Client</th>
					<% If ShowPercentageColumns = True Then %>
						<th class="td-align1 vpc-variance-header" colspan="6" style="border-right: 2px solid #555 !important;">Variances</th>
					<% Else %>
						<th class="td-align1 vpc-variance-header" colspan="4" style="border-right: 2px solid #555 !important;">Variances</th>
					<% End If %>
					<th class="td-align1 vpc-3pavg-header" colspan="5" style="border-right: 2px solid #555 !important;">Sales</th>
					<th class="td-align1 vpc-lcp-header" colspan="5" style="border-right: 2px solid #555 !important;">MCS</th>
					<th class="td-align1 vpc-current-header" colspan="3" style="border-right: 2px solid #555 !important;">EQUIP ROI</th>
					<th class="td-align1 gen-info-header" colspan="2" style="border-right: 2px solid #555 !important;">General</th>

				</tr>
                <tr>
                  
                  
					<th class="td-align sorttable_numeric smaller-header" style="border-left: 2px solid #555 !important; border-top: 2px solid #555 !important;" id="salesColumn">LCP vs<br>3P avg $</th> 
					<% If ShowPercentageColumns = True Then %>
						<th class="td-align sorttable_numeric smaller-header" style="border-top: 2px solid #555 !important;" id="salesColumn">LCP vs<br>3P avg %</th> 
					<%End If %>
					<th class="td-align sorttable_numeric smaller-header" style="border-top: 2px solid #555 !important;" id="salesColumn">Day<br>Impact</th>  
					<th class="td-align sorttable_numeric smaller-header" style="border-top: 2px solid #555 !important;" id="salesColumn"><br>ADS</th> 
					<th class="td-align sorttable_numeric smaller-header" style="border-top: 2px solid #555 !important;" id="salesColumn">LCP vs<br>12P avg $</th>
					<% If ShowPercentageColumns = True Then %>
						<th class="td-align sorttable_numeric smaller-header" style="border-top: 2px solid #555 !important;" id="salesColumn">LCP vs<br>12P avg %</th>
					<% End If %>
                  
                  
                  <th class="td-align sorttable_numeric smaller-header" style="border-left: 2px solid #555 !important; border-top: 2px solid #555 !important;" id="salesColumn"><br>LCP $</th>
                  <th class="td-align sorttable_numeric smaller-header" style="border-top: 2px solid #555 !important;" id="salesColumn"><br>3P avg $</th>
                  <th class="td-align sorttable_numeric smaller-header" style="border-top: 2px solid #555 !important;" id="salesColumn"><br>12P avg $</th>
                  <th class="td-align sorttable_numeric smaller-header" style="border-top: 2px solid #555 !important;" id="salesColumn"><br>Current $</th>
                  <th class="td-align sorttable_numeric smaller-header" style="border-top: 2px solid #555 !important;" id="salesColumn"><br>SPLY $</th> 


                  <th class="td-align sorttable_numeric smaller-header" style="border-left: 2px solid #555 !important; border-top: 2px solid #555 !important;" id="salesColumn"><br>MCS</th>
                  <th class="td-align sorttable_numeric smaller-header" style="border-top: 2px solid #555 !important;" id="salesColumn">LCP vs<br> MCS</th>
                  <th class="td-align sorttable_numeric smaller-header" style="border-top: 2px solid #555 !important;" id="salesColumn">3P avg vs<br> MCS</th>
                  <th class="td-align sorttable_numeric smaller-header" style="border-top: 2px solid #555 !important;" id="salesColumn">12P avg vs<br> MCS</th>
                  <th class="td-align sorttable_numeric smaller-header" style="border-top: 2px solid #555 !important;" id="salesColumn">Current vs<br> MCS</th>
                  
                  
                  <th class="td-align sorttable_numeric smaller-header" style="border-left: 2px solid #555 !important; border-top: 2px solid #555 !important;" id="salesColumn">LCP<br>ROI</th>
                  <th class="td-align sorttable_numeric smaller-header" style="border-top: 2px solid #555 !important;" id="salesColumn">3P avg<br>ROI</th>
                  <th class="td-align sorttable_numeric smaller-header" style="border-right: 2px solid #555 !important; border-top: 2px solid #555 !important;" id="salesColumn">Equipment<br>Value</th>
                  
                  <th class="td-align sorttable_numeric smaller-header" style="border-left: 2px solid #555 !important; border-top: 2px solid #555 !important;" id="salesColumn">Primary<br> Slsmn</th>
                  <th class="td-align sorttable_numeric smaller-header" style="border-top: 2px solid #555 !important;" id="salesColumn">Secondary<br> Slsmn</th>


                </tr>
              </thead>
              
			

<%		
		'Response.Write("<tbody class='searchable'>")
        Response.Write("<tbody>")

		GrandTotLCPvs3PAvgSales = 0
				
		Do While Not rs.EOF

			ShowThisRecord = True

				
			If ShowThisRecord <> False Then			
			
				PrimarySalesMan =  ""
				SecondarySalesMan =  ""
				CustomerType =  ""
				SelectedCustomerID = rs("CustNum")
				CustName = GetCustNameByCustNum(SelectedCustomerID)	
				
				'Extra Fields for Filtering
				SQL4 = "SELECT * FROM AR_Customer WHERE CustNum = '" & SelectedCustomerID & "'"
				
				Set rs4 = Server.CreateObject("ADODB.Recordset")
				rs4.CursorLocation = 3
				Set rs4= cnn8.Execute(SQL4 )

				If Not rs4.Eof Then

					PrimarySalesMan = rs4("Salesman")
					SecondarySalesMan = rs4("SecondarySalesman")
					ReferralCode = rs4("ReferalCode")
					CustomerType = rs4("CustType")
					
				Else
					' Customer not found un AR_Customer
					ShowThisRecord = False
				End If

			End If
			
			
			If ShowThisRecord <> False Then
			
				'TotalCustsReported = TotalCustsReported + 1

				GrandTotLCPvs3PAvgSales = GrandTotLCPvs3PAvgSales + LCPvs3PAvgSales
				
			

				LCPSales = rs("LCPSales")
				If Not IsNumeric(LCPSales) Then LCPSales = 0
				ThreePPSales = rs("ThreePPSales")
				TwelvePPSales = rs("TwelvePPSales")
				CurrentPSales = GetCurrent_PostedTotal_ByCust(SelectedCustomerID,PeriodSeqBeingEvaluated) + GetCurrent_UnPostedTotal_ByCust(SelectedCustomerID,PeriodSeqBeingEvaluated)
				LCPvs3PAvgSales = 0
				LCPvs3PAvgSales = LCPSales - (ThreePPSales/3)
				If Not IsNumeric(LCPvs3PAvgSales) Then LCPvs3PAvgSales = 0
				
'Response.Write(	"x<br>")				
'Response.Write(	LCPSales &"x<br>")
'Response.Write(	ThreePPSales&"x<br>")
'Response.Write(	LCPvs3PAvgSales &"x<br>")

				ImpactDays = (WorkDaysIn3PeriodBasis/3)- WorkDaysInLastClosedPeriod
				DayImpact = ImpactDays  * (LCPSales/WorkDaysInLastClosedPeriod)
				DayImpact = Round(DayImpact,2)
				ADS_LastClosed = (LCPSales/WorkDaysInLastClosedPeriod)
				ADS_3PA = ThreePPSales / (WorkDaysIn3PeriodBasis /3)
				ADS_Variance = ADS_LastClosed -  ADS_3PA 
				If Not IsNumeric(ADS_Variance) Then ADS_Variance = 0
				LCPvs12PAvgSales = LCPSales - (TwelvePPSales/12)
				If Not IsNumeric(LCPvs12PAvgSales) Then LCPvs12PAvgSales = 0
				If LCPvs12PAvgSales <> 0 Then LCPvs12PAvgPercent = ((LCPSales - LCPvs12PAvgSales) / LCPvs12PAvgSales)  * 100 Else LCPvs12PAvgPercent = 0
				SamePLYSales = TotalTPLYAllCats(PeriodSeqBeingEvaluated,SelectedCustomerID)
				If Not IsNumeric(SamePLYSales) Then SamePLYSales = 0
				ThreePPAvgSales = ThreePPSales / 3
				TwelvePPAvgSales = TwelvePPSales / 12
				If ThreePPAvgSales <> 0 Then LCPvs3PAvgPercent = ((LCPSales - ThreePPAvgSales ) / ThreePPAvgSales )  * 100  Else LCPvs3PAvgPercent = 0
				'ROI***********
				TotalEquipmentValue = GetTotalValueOfEquipmentForCustomer(SelectedCustomerID)
				'If CustHasEquipment(SelectedCustomerID) Then
				If TotalEquipmentValue > 0 Then	
					'LCPGP = LCPSales - TotalCostByPeriodSeq(PeriodSeqBeingEvaluated,SelectedCustomerID)
					LCPGP = LCPSales - rs("TotalCostAllCats")
					ThreePAvgGP = ThreePPAvgSales - ( TotalCostByPeriodSeqPrior3P(PeriodSeqBeingEvaluated,SelectedCustomerID) / 3 )
					If LCPGP <> 0 Then ROI = TotalEquipmentValue/LCPGP Else ROI = ""
					If ThreePAvgGP <> 0 Then ROI3P = TotalEquipmentValue/ThreePAvgGP Else ROI3P = ""
				End If

				If Not IsNumeric(ThreePPAvgSales) Then ThreePPAvgSales = 0
				If Not IsNumeric(TwelvePPAvgSales) Then TwelvePPAvgSales = 0
				
				If ShowThisRecord <> False Then
				
					TotalCustsReported = TotalCustsReported + 1
	
					Response.Write("<tr>")
				    Response.Write("<td class='smaller-detail-line'><a href='tools/CatAnalByPeriod/CatAnalByPeriod_SingleCustomer.asp?CID=" & SelectedCustomerID & "&ZDC=0&VB=3Periods&oon=new' target='_blank'>"& SelectedCustomerID  & "</a></td>")
				    Response.Write("<td class='smaller-detail-line'><a href='tools/CatAnalByPeriod/CatAnalByPeriod_SingleCustomer.asp?CID=" & SelectedCustomerID & "&ZDC=0&VB=3Periods&oon=new' target='_blank'>"& CustName & "</a></td>")

					If LCPvs3PAvgSales >= 0 Then
						Response.Write("<td align='" & ColumnAlign & "'  class='smaller-detail-line' data-line-amount='" & LCPvs3PAvgSales & "'>" & FormatCurrency(LCPvs3PAvgSales,0,-2,0) & "</td>")
					Else
						Response.Write("<td align='" & ColumnAlign & "'  class='smaller-detail-line negative' data-line-amount='" & LCPvs3PAvgSales & "'>" & FormatCurrency(LCPvs3PAvgSales,0,-2,0) & "</td>")
					End If
				    If ShowPercentageColumns = True Then
					    Response.Write("<td align='" & ColumnAlign & "' class='smaller-detail-line'>" & FormatNumber(LCPvs3PAvgPercent,0)  & "%</td>")
					End If
				    Response.Write("<td align='" & ColumnAlign & "'  class='smaller-detail-line'>" & FormatCurrency(DayImpact,0) & "</td>")
				    Response.Write("<td align='" & ColumnAlign & "'  class='smaller-detail-line'>" & FormatCurrency(ADS_Variance,0) & "</td>")
		   		    Response.Write("<td align='" & ColumnAlign & "'  class='smaller-detail-line'>" & FormatCurrency(LCPvs12PAvgSales,0) & "</td>")
				    If ShowPercentageColumns = True Then		   		    
					    Response.Write("<td align='" & ColumnAlign & "'  class='smaller-detail-line'>" & FormatNumber(LCPvs12PAvgPercent,0)  & "%</td>")
					End If
					Response.Write("<td align='" & ColumnAlign & "'  class='smaller-detail-line'>" & FormatCurrency(LCPSales,0,-2,0) & "</td>")
					Response.Write("<td align='" & ColumnAlign & "'  class='smaller-detail-line'>" & FormatCurrency(ThreePPAvgSales,0,-2,0) & "</td>")
					Response.Write("<td align='" & ColumnAlign & "'  class='smaller-detail-line'>" & FormatCurrency(TwelvePPAvgSales,0) & "</td>")
				   	Response.Write("<td align='" & ColumnAlign & "'  class='smaller-detail-line'>" & FormatCurrency(CurrentPSales,0,-2,0) & "</td>")
				   	Response.Write("<td align='" & ColumnAlign & "'  class='smaller-detail-line'>" & FormatCurrency(SamePLYSales,0) & "</td>")


					If Not IsNull(MonthlyContractedSalesDollars) Then
						Response.Write("<td align='" & ColumnAlign & "'  class='smaller-detail-line'>" & FormatCurrency(MonthlyContractedSalesDollars,0) & " </td>")
					Else
						Response.Write("<td align='" & ColumnAlign & "'  class='smaller-detail-line'>&nbsp;</td>")
					End If
					

					If Not IsNull(MonthlyContractedSalesDollars) Then
						If (LCPSales-MonthlyContractedSalesDollars) >= 0 Then
							Response.Write("<td align='" & ColumnAlign & "'  class='smaller-detail-line' data-line-amount='" & LCPSales-MonthlyContractedSalesDollars & "'>" & FormatCurrency(LCPSales-MonthlyContractedSalesDollars,0,-2,0) & "</td>")
						Else
							Response.Write("<td align='" & ColumnAlign & "'  class='smaller-detail-line negative' data-line-amount='" & LCPSales-MonthlyContractedSalesDollars & "'>" & FormatCurrency(LCPSales-MonthlyContractedSalesDollars,0,-2,0) & "</td>")
						End If
					Else
						Response.Write("<td align='" & ColumnAlign & "'  class='smaller-detail-line'>&nbsp;</td>")
					End If

					

					If Not IsNull(MonthlyContractedSalesDollars) Then
						If (ThreePPAvgSales-MonthlyContractedSalesDollars) >= 0 Then
							Response.Write("<td align='" & ColumnAlign & "'  class='smaller-detail-line' data-line-amount='" & ThreePPAvgSales-MonthlyContractedSalesDollars & "'>" & FormatCurrency(ThreePPAvgSales-MonthlyContractedSalesDollars,0,-2,0) & "</td>")
						Else
							Response.Write("<td align='" & ColumnAlign & "'  class='smaller-detail-line negative' data-line-amount='" & ThreePPAvgSales-MonthlyContractedSalesDollars & "'>" & FormatCurrency(ThreePPAvgSales-MonthlyContractedSalesDollars,0,-2,0) & "</td>")
						End If
					Else
						Response.Write("<td align='" & ColumnAlign & "'  class='smaller-detail-line'>&nbsp;</td>")
					End If
					
										

					If Not IsNull(MonthlyContractedSalesDollars) Then
						If (TwelvePPAvgSales-MonthlyContractedSalesDollars) >= 0 Then
							Response.Write("<td align='" & ColumnAlign & "'  class='smaller-detail-line' data-line-amount='" & TwelvePPAvgSales-MonthlyContractedSalesDollars & "'>" & FormatCurrency(TwelvePPAvgSales-MonthlyContractedSalesDollars,0,-2,0) & "</td>")
						Else
							Response.Write("<td align='" & ColumnAlign & "'  class='smaller-detail-line negative' data-line-amount='" & TwelvePPAvgSales-MonthlyContractedSalesDollars & "'>" & FormatCurrency(TwelvePPAvgSales-MonthlyContractedSalesDollars,0,-2,0) & "</td>")
						End If
					Else
						Response.Write("<td align='" & ColumnAlign & "'  class='smaller-detail-line'>&nbsp;</td>")
					End If
					



					If Not IsNull(MonthlyContractedSalesDollars) Then
						If (CurrentPSales-MonthlyContractedSalesDollars) >= 0 Then
							Response.Write("<td align='" & ColumnAlign & "'  class='smaller-detail-line' data-line-amount='" & CurrentPSales-MonthlyContractedSalesDollars & "'>" & FormatCurrency(CurrentPSales-MonthlyContractedSalesDollars,0,-2,0) & "</td>")
						Else
							Response.Write("<td align='" & ColumnAlign & "'  class='smaller-detail-line negative' data-line-amount='" & CurrentPSales-MonthlyContractedSalesDollars & "'>" & FormatCurrency(CurrentPSales-MonthlyContractedSalesDollars,0,-2,0) & "</td>")
						End If
					Else
						Response.Write("<td align='" & ColumnAlign & "'  class='smaller-detail-line'>&nbsp;</td>")
					End If




	
					If TotalEquipmentValue > 0 Then	
						If IsNumeric(ROI) and IsNumeric(ROI3P) Then
							If ROI >=10 and ROI3P >= 10 Then ' If both over 10 use red
								Response.Write("<td align='" & ColumnAlign & "'  class='negative smaller-detail-line'>" & FormatNumber(ROI,1)  & "</td>")
								Response.Write("<td align='" & ColumnAlign & "'  class='negative smaller-detail-line'>" & FormatNumber(ROI3P,1)  & "</td>")
							Else
								If ROI <> "" Then
									Response.Write("<td align='" & ColumnAlign & "'  class='smaller-detail-line'>" & FormatNumber(ROI,1)  & "</td>")
								Else
									Response.Write("<td align='" & ColumnAlign & "'  class='smaller-detail-line'>No Sales</td>")
								End If
								Response.Write("<td align='" & ColumnAlign & "'  class='smaller-detail-line'>" & FormatNumber(ROI3P,1)  & "</td>")
							End If
						Else
							If IsNumeric(ROI) Then
								Response.Write("<td align='" & ColumnAlign & "'  class='smaller-detail-line'>" & FormatNumber(ROI,1)  & "</td>")
							Else
								Response.Write("<td align='" & ColumnAlign & "'  class='smaller-detail-line'>No Sales</td>")
							End If
							If IsNumeric(ROI3P) Then
								Response.Write("<td align='" & ColumnAlign & "'  class='smaller-detail-line'>" & FormatNumber(ROI3P,1)  & "</td>")
							Else
								Response.Write("<td>&nbsp;</td>")
							End If
						End If
						' Write equipment value regardless of ROI
						Response.Write("<td align='" & ColumnAlign & "'  class='smaller-detail-line'>")
						Response.Write("<a data-toggle='modal' data-show='true' href='#' data-cust-id='" & SelectedCustomerID & "' data-lcp-gp='0' data-target='#modalEquipmentVPC' data-tooltip='true' data-title='View Customer Equipment'>" & FormatCurrency(TotalEquipmentValue,0) & "</a>")
						Response.Write("</td>")
					Else
						Response.Write("<td align='" & ColumnAlign & "'  class='smaller-detail-line'>&nbsp;</td>")
						Response.Write("<td align='" & ColumnAlign & "'  class='smaller-detail-line'>&nbsp;</td>")
						Response.Write("<td align='" & ColumnAlign & "'  class='smaller-detail-line'>&nbsp;</td>")
					End If
	
	
					' General info
					PrimarySalesPerson = GetSalesmanNameBySlsmnSequence(PrimarySalesMan)
				    SecondarySalesPerson = GetSalesmanNameBySlsmnSequence(SecondarySalesman)
				    If Instr(PrimarySalesPerson ," ") <> 0 Then
						Response.Write("<td class='smaller-detail-line'>" & Left(PrimarySalesPerson,Instr(PrimarySalesPerson ," ")+1) & "</td>")
					Else
						Response.Write("<td class='smaller-detail-line'>" & PrimarySalesPerson & "</td>")
					End If
					If Instr(SecondarySalesPerson," ") <> 0 Then
						Response.Write("<td class='smaller-detail-line'>" & Left(SecondarySalesPerson,Instr(SecondarySalesPerson," ")+1) & "</td>")
					Else
						Response.Write("<td class='smaller-detail-line'>" & SecondarySalesPerson & "</td>")
					End If
	                
				    Response.Write("</tr>")
			    
			    End If

			End If
			
			rs.movenext
				
		Loop
		
		Response.Write("</tbody>")
		
		%>
		<!-------------------- TOTALS FOOTER ---------------------------->
		<!--	
        <tfoot>
            <tr>
                <th style="text-align:left" class="footer-total">TOTALS:</th>
                <th style="text-align:left" class="footer-total" id="totalLCPv3PVarCol"></th>
                <th style="text-align:left" class="footer-total" id="totalLCPSalesCol"></th>
                <th style="text-align:left" class="footer-total" id="total3PASalesCol"></th>
                <th style="text-align:left" class="footer-total" id="totalCurrentSalesCol"></th>
            </tr>
        </tfoot>	
       -->
        
      <tfoot>
         <tr>	
			<th class="footer-total">&nbsp;</th>
    		<th class="footer-total">&nbsp;</th>
    		<% If ShowPercentageColumns = True Then %>
				<th class="vpc-variance-header" colspan="6" style="border-right: 2px solid #555 !important;">&nbsp;</th>
			<% Else %>
				<th class="vpc-variance-header" colspan="4" style="border-right: 2px solid #555 !important;">&nbsp;</th>			
			<% End IF%>

			<th class="vpc-3pavg-header" colspan="5" style="border-right: 2px solid #555 !important;">&nbsp;</th>
			<th class="vpc-lcp-header" colspan="5" style="border-right: 2px solid #555 !important;">&nbsp;</th>
			<th class="vpc-current-header" colspan="2" style="border-right: 2px solid #555 !important;">&nbsp;</th>
			<th class="gen-info-header" colspan="2" style="border-right: 2px solid #555 !important;">&nbsp;</th>
		</tr>
        <tr>
          <th colspan="2" style="text-align:right; border-top: 2px solid #FFF !important;" class="footer-total">Totals:</th>
          
          <% If GrandTotLCPvs3PAvgSales < 0 Then %>
          	<th class="footer-total negative" style="border-left: 2px solid #555 !important; border-top: 2px solid #555 !important; text-align:right;" id="totalLCPv3PVarCol"></th> 
          <% Else %>
          	<th class="footer-total" style="border-left: 2px solid #555 !important; border-top: 2px solid #555 !important; text-align:right;" id="totalLCPv3PVarCol"></th>
          <% End If %>
          
			<% If ShowPercentageColumns = True Then %>
				<th class="footer-total" style="border-top: 2px solid #555 !important;">&nbsp;</th> 
				<th class="footer-total" style="border-top: 2px solid #555 !important;">&nbsp;</th>  
			<% End If %>
  
          <th class="footer-total" style="border-top: 2px solid #555 !important;">&nbsp;</th> 
          <th class="footer-total" style="border-top: 2px solid #555 !important;">&nbsp;</th>
          <th class="footer-total" style="border-top: 2px solid #555 !important;">&nbsp;</th>
          
          
          <th class="footer-total" style="border-left: 2px solid #555 !important; border-top: 2px solid #555 !important; text-align:right;" id="totalLCPSalesCol"></th>
          <th class="footer-total" style="border-top: 2px solid #555 !important; text-align:right;" id="total3PASalesCol"></th>
          <th class="footer-total" style="border-top: 2px solid #555 !important;">&nbsp;</th>
          <th class="footer-total" style="border-top: 2px solid #555 !important; text-align:right;" id="totalCurrentSalesCol"></th>
          <th class="footer-total" style="border-top: 2px solid #555 !important;">&nbsp;</th> 


          <th class="footer-total" style="border-left: 2px solid #555 !important; border-top: 2px solid #555 !important;">&nbsp;</th>
          <th class="footer-total" style="border-top: 2px solid #555 !important;">&nbsp;</th>
          <th class="footer-total" style="border-top: 2px solid #555 !important;">&nbsp;</th>
          <th class="footer-total" style="border-top: 2px solid #555 !important;">&nbsp;</th>
          <th class="footer-total" style="border-top: 2px solid #555 !important;">&nbsp;</th>
          
          
          <th class="footer-total" style="border-top: 2px solid #555 !important;">&nbsp;</th>
          <th class="footer-total" style="border-top: 2px solid #555 !important;">&nbsp;</th>
          
          <th class="footer-total" style="border-top: 2px solid #555 !important;">&nbsp;</th>
          <th class="footer-total" style="border-top: 2px solid #555 !important;">&nbsp;</th>
        </tr>
      </tfoot>
		
			
		<!-------------------- END TOTALS FOOTER ---------------------------->
		<%
		
		Response.Write("</table>")		
		Response.Write("</div>")

'		Response.Write(now()&"<br>")
%>


            </table>
    </div>
         
          </div>
<!-- eof responsive tables !-->



<!-- eof row !-->

<!-- row !-->
<div class="row">
   <%
'Response.Write("<div class='col-lg-12'><h3>" & "Total Customers Listed:" & TotalCustsReported  & "</h3></div>")
%>

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
<!--#include file="../tools/CatAnalByPeriod/CatAnalByPeriod_Modals.asp"-->
<!--#include file="../../inc/footer-main.asp"-->