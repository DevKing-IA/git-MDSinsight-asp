<%
Response.buffer=true
Response.Expires = -1
Response.ExpiresAbsolute = Now() -1 
Response.AddHeader "pragma", "no-store"
Response.AddHeader "cache-control","no-store, no-cache, must-revalidate"
%>
<!--#include file="../../../inc/header-accounts-receivable.asp"-->

<%
'Init MUV_Remove
dummy = MUV_Remove("ConStmt-StartDate") 
dummy = MUV_Remove("ConStmt-EndDate") 
dummy = MUV_Remove("ConStmt-SelectedPeriod")
dummy = MUV_Remove("ConStmt-SkipZeroDollar")	
dummy = MUV_Remove("ConStmt-SkipLessThanZero")
dummy = MUV_Remove("ConStmt-IncludedType") 
dummy = MUV_Remove("ConStmt-CustomOrPredefined")
dummy = MUV_Remove("ConStmt-Account")
dummy = MUV_Remove("ConStmt-IncludeIndividuals")
dummy = MUV_Remove("ConStmt-Chain")
dummy = MUV_Remove("ConStmt-DueDateDays")
dummy = MUV_Remove("ConStmt-DueDateSingleDate")
dummy = MUV_Remove("ConStmt-DoNotShowDueDate")


' Drop & create temporary table
Set cnnTmpTable = Server.CreateObject("ADODB.Connection")
cnnTmpTable.open (Session("ClientCnnString"))
Set rsTmpTable = Server.CreateObject("ADODB.Recordset")
rsTmpTable.CursorLocation = 3 

on error resume next
SQLTmpTable = "DROP TABLE zReportConsolidatedInvoiceOmit_" & Trim(Session("userNo"))
Set rsTmpTable = cnnTmpTable.Execute(SQLTmpTable)
on error goto 0

on error resume next
SQLTmpTable = "DROP TABLE zReportConsolidatedInvoiceInclude_" & Trim(Session("userNo"))
Set rsTmpTable = cnnTmpTable.Execute(SQLTmpTable)
on error goto 0

SQLTmpTable = "CREATE TABLE zReportConsolidatedInvoiceOmit_" & Trim(Session("userNo")) & " (IvsHistSequence varchar(500))"
Set rsTmpTable = cnnTmpTable.Execute(SQLTmpTable)

SQLTmpTable = "CREATE TABLE zReportConsolidatedInvoiceInclude_" & Trim(Session("userNo")) & " (IvsHistSequence varchar(500))"
Set rsTmpTable = cnnTmpTable.Execute(SQLTmpTable)

set rsTmpTable = Nothing
cnnTmpTable.close
set cnnTmpTable = Nothing

		
%>

<%
ClientKeyForFileNames = MUV_READ("ClientKeyForFileNames")
%>


<script type="text/javascript">

	$(function () {
		var autocompleteJSONFileURLAccount = "../../../clientfiles/<%= ClientKeyForFileNames %>/autocomplete/customer_account_list_CSZ_<%= ClientKeyForFileNames %>.json";
		var autocompleteJSONFileURLChain = "../../../clientfiles/<%= ClientKeyForFileNames %>/autocomplete/customer_chain_list_<%= ClientKeyForFileNames %>.json";
		
		var optionsAccount = {
		  url: autocompleteJSONFileURLAccount,
		  placeholder: "Search for a customer by name, account, city, state, zip",
		  getValue: "name",
		  list: {	
	        onChooseEvent: function() {
	            var custID = $("#txtCustID").getSelectedItemData().code;
	            $("#optAccount").prop("checked","checked");
	            $("#txtCustIDToPass").val(custID);
        	},		  
		    match: {
		      enabled: true
			},
			maxNumberOfElements: 20		
		  },
		  theme: "round"
		};
		$("#txtCustID").easyAutocomplete(optionsAccount);
		
		
		var optionsChain = {
		  url: autocompleteJSONFileURLChain,
		  placeholder: "Search for a chain by chain number or name",
		  getValue: "name",
		  list: {	
	        onChooseEvent: function() {
	            var chainID = $("#txtChainID").getSelectedItemData().code;
	            $("#optChain").prop("checked","checked");
	            $("#txtChainIDToPass").val(chainID);
        	},		  
		    match: {
		      enabled: true
			},
			maxNumberOfElements: 20		
		  },
		  theme: "round"
		};
		$("#txtChainID").easyAutocomplete(optionsChain);
		

	})
</script>


<style>
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

<div id="PleaseWaitPanel">
	<br><br>Processing, please wait...<br><br>
	<img src="../../../img/loading.gif"/>
</div>

<script type="text/javascript">
$(document).ready(function() {
    $("#PleaseWaitPanel").hide();
});
</script>


<script type="text/javascript">
function OnSubmitForm()
{
  if(document.frmconsolidatedStatement.optAccountOrChain[0].checked == true)
  {
    document.frmconsolidatedStatement.action ="consolidated_stmt_frm_acct.asp";
  }
  else
  if(document.frmconsolidatedStatement.optAccountOrChain[1].checked == true)
  {
    document.frmconsolidatedStatement.action ="consolidated_stmt_frm_chn.asp";
  }
  return true;
}
</script>
   
<!-- css !-->
<style type="text/css">
 .beatpicker-clear{
	 display: none;
 } 

.row-line{
	margin-bottom: 20px;
}	 

.date-ranges h4{
	text-align: center;
	margin-top: 25px;
} 

	.account-chain h4{
	text-align: center;
	margin-top: 25px;
}

.col-box{
	border: 1px solid #ccc;
	padding: 15px;
} 

.due-date{
	margin-top:
	20px;
}

.due-date strong{
	display:block;
	width:100%;
	margin-bottom:15px;
}

.due-date select{
	max-width:40%;
	margin:0 auto;
}

select option{
	width: auto;
}

.table>tbody>tr>td, .table>tbody>tr>th, .table>tfoot>tr>td, .table>tfoot>tr>th, .table>thead>tr>td, .table>thead>tr>th{
	border: 0px;
}
 
 
 </style>
<!-- eof css !-->

<h1 class="page-header"><i class="fa fa-file-text"></i> Consolidated Invoice</h1>


<form method="POST" name="frmconsolidatedStatement" id="frmconsolidatedStatement" onsubmit="return OnSubmitForm();">		    
	
<!-- row !-->
<div class="row row-line">

    
    <!-- date ranges starts here !-->
    <div class="col-lg-6">
	    <div class="col-box">
	    <!-- date !-->
	    <div class="row date-ranges row-line">
		    
		    <div class="col-lg-4"><h4><input type="radio" id="optCustom" name="optCustomOrPredefined" value="Custom"> Custom Range</h4></div>
		    
			    <div class="col-lg-4">
					<div class="form-group">
						<input type="hidden" id="txtStartDate" name="txtStartDate">
						<input type="hidden" id="txtEndDate" name="txtEndDate">
						Select Dates
						<div class="btn btn-default" id="reportrange">
							<i class="fa fa-calendar"></i> &nbsp;
							<span></span>
							<b class="fa fa-angle-down"></b>
						</div>
					</div>
			    </div>
		    
		    </div>
	    
	    <!-- month !-->
	    <div class="row date-ranges row-line">
		    <div class="col-lg-4"><h4><input type="radio" id="optPredefined" name="optCustomOrPredefined" checked value="Predefined"> Predfined</h4></div>
			    <div class="col-lg-4">Select Period
				   <select class="form-control" id="selPeriod" name="selPeriod" onchange="setPredefined()">
				      	
				      	<%'Dont go past the last closed period
				      	 
				      	'Get all period info
			      	  	SQL = "SELECT * FROM " & Session("dbowner") & ".BillingPeriodHistory "
			      	  	SQL = SQL & "WHERE BillPerSequence < " & GetLastClosedPeriodSeqNum() + 1
			      	  	SQL = SQL & " order by [Year] desc, Period desc"
	
						Set cnn8 = Server.CreateObject("ADODB.Connection")
						cnn8.open (Session("ClientCnnString"))
						Set rs = Server.CreateObject("ADODB.Recordset")
						rs.CursorLocation = 3 
						Set rs = cnn8.Execute(SQL)
					
						If not rs.EOF Then
							Do
								Response.Write("<option value='" & FormatDateTime(rs("BeginDate")) & "~" & FormatDateTime(rs("EndDate")) & "'>" & FormatDateTime(rs("BeginDate")) & " - " & FormatDateTime(rs("EndDate")) & "</option>")
								rs.movenext
							Loop until rs.eof
						End If
						set rs = Nothing
						cnn8.close
						set cnn8 = Nothing
				      	%>
					</select>	    
				</div>
			</div>
	    </div>
		</div>

    
    
	    <!-- account / chain starts here !-->
	    <div class="col-lg-6">
	    <div class="col-box">
		    <!-- account !-->
		    <div class="row account-chain row-line">
		    
			    <div class="col-lg-2">
				     <h4><input type="radio" id="optAccount" name="optAccountOrChain" checked value="Account"> <%=GetTerm("Account")%></h4>
	 		    </div>
	    		    
			    <div class="col-lg-6">Please select a <%=GetTerm("customer")%> from the list below
	        		<!-- select company !-->
						<input id="txtCustID" name="txtCustID" onchange='$("#optAccount").prop("checked","checked");'>
						<input type="hidden" id="txtCustIDToPass" name="txtCustIDToPass">
						<i id="searchIcon" class="fa fa-search fa-2x"></i>
					<!-- eof select company !-->
				</div>
				
			</div>
     
		    <!-- chain !-->
		    <div class="row account-chain row-line">
		    
				<div class="col-lg-2">
					<h4><input type="radio" id="optChain" name="optAccountOrChain" value="Chain"> Chain</h4>
 			    </div>
		    
			    <div class="col-lg-6">Please select a chain from the list below
			    
	        		<!-- select company !-->
						<input id="txtChainID" name="txtChainID" onchange='$("#optChain").prop("checked","checked");'>
						<input type="hidden" id="txtChainIDToPass" name="txtChainIDToPass">
						<i id="searchIcon" class="fa fa-search fa-2x"></i>
					<!-- eof select company !-->

			    </div>
			</div>
		</div>
	    </div>
	</div>

	
<!-- individual invoices / include following types !-->
<div class="row row-line">
	<div class="col-lg-6">
			<div class="row">
				
		<!-- individual invoices starts here !-->
		<div class="col-lg-6">
			
			
            <!-- skip box !-->
			<div class="col-box">
			<div class="table-responsive">
  <table class="table">
  	<tbody>
  		
  		<!-- line !-->
  		<tr>
	  		<td>Generate copies of individual invoices</td>
	  		<td> <input type="checkbox" id="chkIncludeIndividuals" name="chkIncludeIndividuals" checked></td>
  		</tr>
  		<!-- eof line !-->
  		
  		<!-- line !-->
  		<tr>
	  		<td>Skip $0 invoices</td>
	  		<td> <input type="checkbox" id="chkZeroDollar" name="chkZeroDollar" checked></td>
  		</tr>
  		<!-- eof line !-->
  		
  		<!-- line !-->
  		<tr>
	  		<td>Skip invoices less then $0 (credits)</td>
	  		<td> <input type="checkbox" id="chkLessThanZero" name="chkLessThanZero"></td>
  		</tr>
  		<!-- eof line !-->

  	
  	</tbody>
  </table>
</div>

			</div>
            
            <!-- eof skip box !-->
            
            <!-- due date !-->
            <div class="col-box due-date">
            <input type="radio" name="radInvoiceDueDate" id="radInvoiceDueDate" value="DAYS" checked="checked">&nbsp;<b>Due Date Day Range- </b>
            Please specify the number of days from the options below. The system will add this many days to the end date specified by your date range criteria to generate a due date for the consolidated invoice.
            <br><br>
            <select class="form-control" name="selDueDate" id="selDueDate">
            	<%	For x = 1 to 365 
            			If x = 1 then
            				Response.Write("<option value='" & x & "'>" & x & " day</option>")
            			Elseif x = 30 Then
            				Response.Write("<option value='" & x & "' selected>" & x & " days</option>") ' default to 30 days            			            			
            			Else
            				Response.Write("<option value='" & x & "'>" & x & " days</option>")            			
            			End If
					Next%>
            </select>

			<hr>
			
            <input type="radio" name="radInvoiceDueDate" id="radInvoiceDueDate" value="SINGLEDATE">&nbsp;<b>Due Date Single - </b>
            or select a specific due date for the consolidate invoice.
            <br><br>
            <div align="center">
	            <div class="input-group input-append date" id="datePicker" style="width:200px;">
	                <input type="text" class="form-control" name="txtDueDate" id="txtDueDate">
	                <span class="input-group-addon add-on"><span class="glyphicon glyphicon-calendar"></span></span>
	            </div>
            </div>
            
            <hr>
            
            <b><input type="checkbox" name="chkDoNotShowDueDate" id="chkDoNotShowDueDate">&nbsp; Do Not Display Due Date - </b>
            check this option if you do not want the invoice due date to appear on the generated invoices.

            </div>
         
            
		</div>
		<!-- individual invoices ends here !-->
		
		<!-- include following types starts here !-->
		<div class="col-lg-6">
			<div class="col-box">
				
				<h4>Include the following types</h4>
				
			<div class="table-responsive">
  <table class="table">
  	<tbody>
  		
  		<!-- line !-->
  		<tr>
	  		<td>Backorder (B)</td>
	  		<td><input type="checkbox" id="chkBackOrder" name="chkBackOrder"></td>
  		</tr>
  		<!-- eof line !-->
  		
  		<!-- line !-->
  		<tr>
	  		<td>Credit Memo (C)</td>
	  		<td><input type="checkbox" id="chkCreditMemo" name="chkCreditMemo"></td>
  		</tr>
  		<!-- eof line !-->
  		
  		<!-- line !-->
  		<tr>
	  		<td>Simple A/R Debit (E)</td>
	  		<td><input type="checkbox" id="chkSimpleDebit" name="chkSimpleDebit"></td>
  		</tr>
  		<!-- eof line !-->
  		
  		<!-- line !-->
  		<tr>
	  		<td>Rental (G)</td>
	  		<td><input type="checkbox" id="chkRental" name="chkRental" checked></td>
  		</tr>
  		<!-- eof line !-->
  		
  		<!-- line !-->
  		<tr>
	  		<td>Route Invoicing (I)</td>
	  		<td><input type="checkbox" id="chkRouteInvoicing" name="chkRouteInvoicing"></td>
  		</tr>
  		<!-- eof line !-->
  		
  		<!-- line !-->
  		<tr>
	  		<td>Interest (O)</td>
	  		<td><input type="checkbox" id="chkInterest" name="chkInterest"></td>
  		</tr>
  		<!-- eof line !-->
  		
  		<!-- line !-->
  		<tr>
	  		<td>Telsel Invoicing (T)</td>
	  		<td><input type="checkbox" id="chkTelsel" name="chkTelsel" checked></td>
  		</tr>
  		<!-- eof line !-->
  	
  	</tbody>
  </table>
</div>

			</div>
		</div>
		<!-- include following types ends here !-->
		
			</div>
	</div>
</div>
<!-- eof individual invoices / include following types !-->

<!-- buttons row !-->
	<div class="row">
		<div class="col-lg-12">
			<p align="right">
				<br>
				<button type="button" class="btn btn-default">Cancel</button>
				<button type="submit" class="btn btn-primary">Run Report</button>
			</p>
		</div>
	</div>

</form>


<!--#include file="../../../inc/footer-main.asp"-->
<style type="text/css">
.datepicker.dropdown-menu {right: auto;}
</style>
<script src="<%= baseURL %>js/moment.min.js" type="text/javascript"></script>

<!-- Include Bootstrap DaterangePicker For Invoice Date Range Selection -->
<link href="<%= baseURL %>js/bootstrap-daterangepicker/daterangepicker.min.css" rel="stylesheet" type="text/css" />
<script src="<%= baseURL %>js/bootstrap-daterangepicker/daterangepicker.min.js" type="text/javascript"></script>

<!-- Include Bootstrap DatePicker For Due Date SINGLE Date Selection -->
<link rel="stylesheet" href="http://cdnjs.cloudflare.com/ajax/libs/bootstrap-datepicker/1.3.0/css/datepicker.min.css" />
<link rel="stylesheet" href="http://cdnjs.cloudflare.com/ajax/libs/bootstrap-datepicker/1.3.0/css/datepicker3.min.css" />
<script src="http://cdnjs.cloudflare.com/ajax/libs/bootstrap-datepicker/1.3.0/js/bootstrap-datepicker.min.js"></script>

<script type="text/javascript">
		function setPredefined(){
			$("#optPredefined").prop("checked","checked");
		}
        $('#reportrange').daterangepicker({
                opens: 'right',
                startDate: moment(),
                endDate: moment(),
                showWeekNumbers: true,
                timePicker: false,
                linkedCalendars: false,
                autoUpdateInput:false,
                autoApply:true,
                ranges: {
                    'Today': [moment(), moment()],
                    'Yesterday': [moment().subtract('days', 1), moment().subtract('days', 1)],
                    'Last 7 Days': [moment().subtract('days', 6), moment()],
                    'Last 30 Days': [moment().subtract('days', 29), moment()],
                    'This Month': [moment().startOf('month'), moment().endOf('month')],
                    'Last Month': [moment().subtract('month', 1).startOf('month'), moment().subtract('month', 1).endOf('month')]
                },
                buttonClasses: ['btn'],
                applyClass: 'green',
                cancelClass: 'default',
                format: 'MM/DD/YYYY',
                separator: ' to ',
                locale: {
                    applyLabel: 'Apply',
                    fromLabel: 'From',
                    toLabel: 'To',
                    customRangeLabel: 'Custom Range',
                    daysOfWeek: ['Su', 'Mo', 'Tu', 'We', 'Th', 'Fr', 'Sa'],
                    monthNames: ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'],
                    firstDay: 1
                }
            },
            function (start, end) {
                $('#reportrange span').html(start.format('MM/DD/YYYY') + ' - ' + end.format('MM/DD/YYYY'));
                $('#txtStartDate').val(start.format('MM/DD/YYYY'));
                $('#txtEndDate').val(end.format('MM/DD/YYYY'));
				$("#optCustom").prop("checked","checked");
            }
        );
        //Set the initial state of the picker label
        $('#reportrange span').html(moment().format('MM/DD/YYYY') + ' - ' + moment().format('MM/DD/YYYY'));
		$('#txtStartDate').val(moment().format('MM/DD/YYYY'));
		$('#txtEndDate').val(moment().format('MM/DD/YYYY'));
</script>

<script>
    
  $(document).ready(function () {
        $('#datePicker').datepicker({
            format: 'mm/dd/yyyy',
            clearBtn: true,
            todayHighlight: true,
        }).on('changeDate', function(ev){
			$('input[name="radInvoiceDueDate"][value="SINGLEDATE"]').removeAttr('checked');
			$('input[name="radInvoiceDueDate"][value="DAYS"]').removeAttr('checked');
			$('input[name="radInvoiceDueDate"][value="SINGLEDATE"]').attr('checked','checked');
			$('input[name="radInvoiceDueDate"][value="SINGLEDATE"]').prop("checked", true);
        }).on('clearDate', function(ev){
			$('input[name="radInvoiceDueDate"][value="SINGLEDATE"]').removeAttr('checked');
			$('input[name="radInvoiceDueDate"][value="DAYS"]').removeAttr('checked');
			$('input[name="radInvoiceDueDate"][value="DAYS"]').attr('checked','checked');
			$('input[name="radInvoiceDueDate"][value="DAYS"]').prop("checked", true);

    	});  
    	
    	
	    $(document).on('change','[name="selDueDate"]',function(){

			var selectedDueDateRange = $(this).find(":selected").val();
			
			if (selectedDueDateRange != '') {
				$('input[name="radInvoiceDueDate"][value="SINGLEDATE"]').removeAttr('checked');
				$('input[name="radInvoiceDueDate"][value="DAYS"]').removeAttr('checked');
				$('input[name="radInvoiceDueDate"][value="DAYS"]').attr('checked','checked');
				$('input[name="radInvoiceDueDate"][value="SINGLEDATE"]').removeAttr('checked');
				$('input[name="radInvoiceDueDate"][value="DAYS"]').prop("checked", true);
			}
			else {
				$('input[name="radInvoiceDueDate"][value="SINGLEDATE"]').removeAttr('checked');
				$('input[name="radInvoiceDueDate"][value="DAYS"]').removeAttr('checked');
				$('input[name="radInvoiceDueDate"][value="SINGLEDATE"]').removeAttr('checked');
				$('input[name="radInvoiceDueDate"][value="DAYS"]').attr('checked','checked');
				$('input[name="radInvoiceDueDate"][value="SINGLEDATE"]').prop("checked", true);
			}
			
		}); 
					
    	
    });  
</script>

