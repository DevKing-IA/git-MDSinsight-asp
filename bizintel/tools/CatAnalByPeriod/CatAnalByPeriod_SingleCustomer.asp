<!--#include file="../../../inc/header.asp"-->
<!--#include file="../../../inc/InSightFuncs.asp"-->
<!--#include file="../../../inc/InSightFuncs_BizIntel.asp"-->
<!--#include file="../../../inc/InsightFuncs_Equipment.asp"-->
<!--#include file="../../../inc/InsightFuncs_AR_AP.asp"-->
<!--#include file="../../../css/fa_animation_styles.css"-->
<%

Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))


'DebugMsg = True
Session("DebugSpeed") = False
Session("SkipGetTerm") = False
Session("NewPostLogic") = True
If Session("DebugSpeed") = True Then Response.Write("<br><br>1:" & Now())	

Session("CalcTax") = False

Server.ScriptTimeout = 900000 'Default value

CustForDetail = Request.QueryString("CID")

If CustForDetail = "" Then 
	CustForDetail = Request.Form("txtCustIDToPassMainSearch")
End If

'If CustForDetail ="" Then CustForDetail ="6700"

ShowZeros = Request.QueryString("ZDC")

If ShowZeros = "" Then 

	ShowZeros = Request.Form("chkShowZeroDollarCategories")
	
	If (ShowZeros <> "" AND ShowZeros = "on") Then 
		ShowZeros = 1 
	Else 
		ShowZeros = 0
	End If
End If

ShowGPPercent = 1


VarianceBasis = Request.QueryString("VB")

If VarianceBasis = "" Then 
	VarianceBasis = Request.Form("optWhatToShow")
	
	If VarianceBasis = "" Then
		VarianceBasis = "3Periods"
	End If
	
End If


oldornew = "new"

PeriodBeingEvaluated = GetLastClosedPeriodAndYear()
PeriodSeqBeingEvaluated = GetLastClosedPeriodSeqNum()

CustName = GetCustNameByCustNum(CustForDetail)
dummy=MUV_WRITE("VarianceBasis",VarianceBasis)


WorkDaysIn3PeriodBasis =  NumberOfWorkDays(GetPeriodBeginDateBySeq(PeriodSeqBeingEvaluated -3), GetPeriodEndDateBySeq(PeriodSeqBeingEvaluated -1))+1
WorkDaysIn12PeriodBasis =  NumberOfWorkDays(GetPeriodBeginDateBySeq(PeriodSeqBeingEvaluated -12), GetPeriodEndDateBySeq(PeriodSeqBeingEvaluated -1)) + 1 
WorkDaysInLastClosedPeriod = NumberOfWorkDays(GetPeriodBeginDateBySeq(PeriodSeqBeingEvaluated), GetPeriodEndDateBySeq(PeriodSeqBeingEvaluated )) + 1 
WorkDaysInCurrentPeriod = NumberOfWorkDays(GetPeriodBeginDateBySeq(PeriodSeqBeingEvaluated +1), GetPeriodEndDateBySeq(PeriodSeqBeingEvaluated +1)) + 1 
WorkDaysSoFar =  NumberOfWorkDays(GetPeriodBeginDateBySeq(PeriodSeqBeingEvaluated +1),Date()) + 1


If CustForDetail <> "" Then CreateAuditLogEntry "Report","Report","Minor",0, MUV_Read("DisplayName") & " ran the report: Category Analysis by Period For Acct# " & CustForDetail & " " & CustName

ClientKeyForFileNames = MUV_READ("ClientKeyForFileNames")

'Get all the information we need for the customer here so we only do it once
Cust_InstallDate = "" : Cust_MGP = "" : Cust_MGPSales = "" : Cust_LastBuy = "" 
Cust_ContStart = "" : Cust_ContEnd = "" : Cust_ContTerm = ""
Cust_SalesPerson1 = "" : Cust_SalesPerson2 = "" : Cust_Type = "" : Cust_ChainID = "" : Cust_ChainName = "" : Cust_ReferralCode = ""

SQL4 = "SELECT * FROM AR_Customer WHERE CustNum = '" & CustForDetail & "'"
Set rs4 = Server.CreateObject("ADODB.Recordset")
rs4.CursorLocation = 3
rs4.Open SQL4 , Session("ClientCnnString")

If Not rs4.Eof Then

	Cust_InstallDate = rs4("InstallDate")
	If Cust_InstallDate <> "" Then Cust_InstallDate = FormatDateTime(Cust_InstallDate,2)

	Cust_MGP = rs4("ProjGpPerMonth")
	If rs4("ProjSalesPerMonth") <> "" Then Cust_MGPSales = FormatCurrency(rs4("ProjSalesPerMonth"),0)
		
	
	Cust_LastBuy = rs4("LastBuyDate")
	If Cust_LastBuy <> "" Then Cust_LastBuy = FormatDateTime(Cust_LastBuy,2)

	Cust_SalesPerson1 = GetSalesmanNameBySlsmnSequence(rs4("Salesman"))
	Cust_SalesPerson2 = GetSalesmanNameBySlsmnSequence(rs4("SecondarySalesman"))
	Cust_Type = GetCustTypeByCode(rs4("CustType"))
	Cust_ChainID = rs4("ChainNum")
	If Cust_ChainID = 0 Then Cust_ChainID = ""
	If Cust_ChainID <> "" Then Cust_ChainName = GetChainDescByChainNum(Cust_ChainID)
	
	Cust_ContStart = rs4("ContractStartDate")
	If Cust_ContStart <> "" Then Cust_ContStart = FormatDateTime(Cust_ContStart,2)
	Cust_ContEnd = rs4("ContractEndDate")
	If Cust_ContEnd <> "" Then Cust_ContEnd = FormatDateTime(Cust_ContEnd,2)
	If rs4("ContractMonths") <> 0 Then Cust_ContTerm = rs4("ContractMonths") & " months"

	Cust_ReferralCode = rs4("ReferalCode")
	If Cust_ReferralCode <> "" Then Cust_ReferralCode = GetReferralNameByCode(Cust_ReferralCode)
	
	Cust_ContactName = rs4("contact")

End If
Set rs4 = Nothing

%>

<%
SQL4 = "SELECT * FROM ARAP WHERE CustNum = '" & CustForDetail & "'"
Set rs4 = Server.CreateObject("ADODB.Recordset")
rs4.CursorLocation = 3
rs4.Open SQL4 , Session("ClientCnnString")

If Not rs4.Eof Then

	Cust_ContactTel = rs4("contactTel")
	Cust_ContactEmail = rs4("generalEmailAddr")

End If
Set rs4 = Nothing

%>


<%
'Get all the posted, unposted for this customer & put in an array by category
' write to handle tax
If Session("NewPostLogic") = True Then

	Redim PostUnPostCatArray(22,2) 'Cat, Unposted element 0, Posted element 1
	Redim PostUnPostCatArrayCost(22,2)
	
	For x = 0 to 21
		PostUnPostCatArray(x,0) = 0
		PostUnPostCatArray(x,1) = 0
		PostUnPostCatArrayCost(x,0) = 0
		PostUnPostCatArrayCost(x,1) = 0
	Next 

	SQL4 = "SELECT * FROM BI_PostedUnpostedByCustCatPeriod WHERE CustID = '" & CustForDetail & "'"
	'Response.Write("SQL4:" & SQL4 & "<br>")
	Set rs4 = Server.CreateObject("ADODB.Recordset")
	rs4.CursorLocation = 3
	rs4.Open SQL4 , Session("ClientCnnString")

	If Not rs4.Eof Then
		Do While Not rs4.Eof
		
			If rs4("PostedOrUnposted") = "U" Then 
				PostUnPostCatArray(rs4("CategoryID"),0) = rs4("TotalSales")
				PostUnPostCatArrayCost(rs4("CategoryID"),0) = rs4("TotalCost")
			End If
			If rs4("PostedOrUnposted") = "P" Then 
				PostUnPostCatArray(rs4("CategoryID"),1) =  rs4("TotalSales")
				PostUnPostCatArrayCost(rs4("CategoryID"),1) =  rs4("TotalCost")
			End If
			
			rs4.MoveNext
		Loop
	
	End If

	If Session("DebugSpeed") = True Then 
		For x = 0 to Ubound(PostUnPostCatArray)-1
			For z = 0 to 1 
				Response.Write(x & "---" & PostUnPostCatArray(x,z) & ":")
			Next
			Response.Write("<br>")
		Next
	End If

End If

%>
<!---------------------------------------------------------------------------------------------------------->
<!----------THIS IS A CUSTOM STYLESHEET ADDED FOR THE AUTOCOMPLETE SEARCH FOR CATEGORY ANALYSIS ONLY-------->
<!-----------IT OVERRIDES THE STYLES THAT ARE STILL LOADED IN HEADER.ASP------------------------------------>

<!---------------------------------------------------------------------------------------------------------->
<link rel="stylesheet" href="<%= BaseURL %>js/easyautocomplete/EasyAutocomplete-1.4.0/easy-autocomplete-cat-analysis.css"> 

<!---------------------------------------------------------------------------------------------------------->
<!---------------------------------------------------------------------------------------------------------->
<!--
	js/easyautocomplete/EasyAutocomplete-1.4.0/easy-autocomplete.themes.css ALSO CONTAINS A CUSTOM STYLE
	SET CALLED "easy-autocomplete.eac-cat-analysis"
-->
<!---------------------------------------------------------------------------------------------------------->
<!---------------------------------------------------------------------------------------------------------->


<script type="text/javascript">


$(document).ready(function(){

	$('#modalEditCustomerNotes').on('show.bs.modal', function(e) {

	    //get data-id attribute of the clicked order
	    var CustID = $(e.relatedTarget).data('cust-id');
	    var CategoryID = $(e.relatedTarget).data('category-id');
	    
	    //populate the textbox with the id of the clicked order
	    $(e.currentTarget).find('input[name="txtCustIDToPassMainSearch"]').val(CustID);
	    $(e.currentTarget).find('input[name="txtCustIDToPassToGenerateNotes"]').val(CustID);
	    $(e.currentTarget).find('input[name="txtCategoryID"]').val(CategoryID);
	    	    
	    var $modal = $(this);

    	$.ajax({
			type:"POST",
			url: "../../../../inc/InSightFuncs_AjaxForARAP.asp",
			cache: false,
			data: "action=GetContentForCustomerNotesModal&CustID="+encodeURIComponent(CustID),
			success: function(response)
			 {
               	 $modal.find('#modalEditCustomerNotesContent').html(response);               	 
             },
             failure: function(response)
			 {
			  	$modal.find('#modalEditCustomerNotesContent').html("Failed");
	            //var height = $(window).height() - 600;
	            //$(this).find(".modal-body").css("max-height", height);
             }
		});
	});
	

	$('#modalCategoryVPC').on('show.bs.modal', function(j) {

	    //get data-id attribute of the clicked order
	    var CustID = $(j.relatedTarget).data('cust-id');
	    var CategoryID = $(j.relatedTarget).data('category-id');
	    var PeriodSeq = $(j.relatedTarget).data('period-seq');
	    var VarBasis = $(j.relatedTarget).data('variance-basis');
 
	    
	    
	    //populate the textbox with the id of the clicked order
	    $(j.currentTarget).find('input[name="txtCustIDToPassMainSearch"]').val(CustID);
	    $(j.currentTarget).find('input[name="txtCategoryID"]').val(CategoryID);
	    $(j.currentTarget).find('input[name="txtPeriodSeqToPass"]').val(PeriodSeq);
	    $(j.currentTarget).find('input[name="txtVarianceBasisToPass"]').val(VarBasis);
	    	    
	    var $modal = $(this);
	    
    	$.ajax({
			type:"POST",
			url: "../../../../inc/InSightFuncs_AjaxForBizIntelModals.asp",
			cache: false,
			data: "action=GetTitleForCategoryVPCModal&CategoryID="+encodeURIComponent(CategoryID)+"&PeriodSeq="+encodeURIComponent(PeriodSeq)+"&VarBasis="+encodeURIComponent(VarBasis)+"&CustID="+encodeURIComponent(CustID),
			success: function(response)
			 {
               	 $modal.find('#modalCategoryVPCTitle').html(response);            	 
             },
             failure: function(response)
			 {
			  	$modal.find('#modalCategoryVPCTitle').html("Failed");
             }
		});
		
	    

    	$.ajax({
			type:"POST",
			url: "../../../../inc/InSightFuncs_AjaxForBizIntelModals.asp",
			cache: false,
			data: "action=GetContentForCategoryVPCModal&CategoryID="+encodeURIComponent(CategoryID)+"&CustID="+encodeURIComponent(CustID)+"&PeriodSeq="+encodeURIComponent(PeriodSeq)+"&VarBasis="+encodeURIComponent(VarBasis),
			success: function(response)
			 {
               	 $modal.find('#modalCategoryVPCContent').html(response);  
               	 //sorttable.innerSortFunction.call(document.getElementById($modal.find('#salesColumn')));             	 
             },
             failure: function(response)
			 {
			  	$modal.find('#modalCategoryVPCContent').html("Failed");
	            //var height = $(window).height() - 600;
	            //$(this).find(".modal-body").css("max-height", height);
             }
		});
		
	
    
	});
	
	
	
	    

	$('#modalEquipmentVPC').on('show.bs.modal', function(j) {

	    //get data-id attribute of the clicked order
	    var CustID = $(j.relatedTarget).data('cust-id');
	    var LCPGP = $(j.relatedTarget).data('lcp-gp');
 
	    //populate the textbox with the id of the clicked order
	    $(j.currentTarget).find('input[name="txtCustIDToPassMainSearch"]').val(CustID);
	    $(j.currentTarget).find('input[name="txtLastClosedPeriodGP"]').val(LCPGP);
	    	    
	    var $modal = $(this);
	    //$modal.find('#PleaseWaitPanelModal').show();  

    	$.ajax({
			type:"POST",
			url: "../../../../inc/InSightFuncs_AjaxForBizIntelModals.asp",
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
		
	});
	

	$("#PleaseWaitPanel").hide();

	var $content = $(".content-line").hide();
	
	$(".toggle-line").on("click", function(e){
		$(this).toggleClass("expanded");
		$content.slideToggle();
	});
  
  
	var autocompleteJSONFileURLAccount = "../../../clientfiles/<%= ClientKeyForFileNames %>/autocomplete/customer_account_list_CSZ_<%= ClientKeyForFileNames %>.json";


	var optionsAccountMainPage = {
	  url: autocompleteJSONFileURLAccount,
	  placeholder: "Search for a customer by name, account, city, state, zip",
	  getValue: "name",
	  list: {	
        onChooseEvent: function() {
        
            var custID = $("#txtCustIDMainSearch").getSelectedItemData().code;
            
            //$("#AddMCSClientCustNameToPass").val($("#txtCustID").getSelectedItemData().name.split(" --- ")[1])
            
            $("#txtCustIDToPassMainSearch").val(custID);
            
            var zeroDollarCategories = $("#chkShowZeroDollarCategories").is(":checked");
            
            if (zeroDollarCategories == true) {
            	zeroDollarCats = 1
            }
            else{
            	zeroDollarCats = 0 
            }
            
            if ($('#optWhatToShow3Periods').attr("checked") == "checked") {
            	varianceBasis = "3Periods";
      		}

            if ($('#optWhatToShow12Periods').attr("checked") == "checked") {
            	varianceBasis = "12Periods";
      		}
      		
            if ($('#optold').attr("checked") == "checked") {
            	oldornew = "old";
      		}

            if ($('#optnew').attr("checked") == "checked") {
            	oldornew = "new";
      		}
			
			oldornew = "new";
			
            window.location.href = "CatAnalByPeriod_SingleCustomer.asp?CID=" + custID + "&ZDC=0&VB=3Periods" + "&oon=" + oldornew;
            
    	},		  
	    match: {
	      enabled: true
		},
		maxNumberOfElements: 20		
	  },
	  theme: "cat-analysis"
	};
	
	$("#txtCustIDMainSearch").easyAutocomplete(optionsAccountMainPage);
	
	
	$("#chkShowZeroDollarCategories").change(function() {
		$("#frmCatAnalByPeriodSingleCust").submit();
	});
	
	//$("#chkShowGPPercent").change(function() {
		//$("#frmCatAnalByPeriodSingleCust").submit();
	//});
	

	$("#optWhatToShow3Periods").change(function() {
		$("#frmCatAnalByPeriodSingleCust").submit();
	});
	
	$("#optWhatToShow12Periods").change(function() {
		$("#frmCatAnalByPeriodSingleCust").submit();
	});
	

	$("#optold").change(function() {
		$("#frmCatAnalByPeriodSingleCust").submit();
	});
	
	$("#optnew").change(function() {
		$("#frmCatAnalByPeriodSingleCust").submit();
	});
  
});

	function ajaxRowMode(type, id, mode) {
	
		$('#ajaxRow'+type+'-'+id).attr("class", "ajaxRow"+mode);
		if(id==0){
			$('#ajaxRow'+type+'-' + 0 + '').remove();
		}	
	
		 $(".ajaxRowEdit").find('input[disabled="true"]').each(function () {
		     $(this).removeAttr("disabled");
		 });
		
	}


</script>





<style>


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
	
	
	#PleaseWaitPanelModal{
		position: relative;
		z-index: 9999;
		background-color: #fff;
		opacity:1.0;
		text-align:center;
	}  
	
	 /* .customer-leakage-td{
		width:20%;
		max-width: 20%;
	}  */
	
	.three-cols-bg-blue{
		background: #ddf3ff;
	}
	
	.three-cols-bg-green{
		background: #e4fbed;
	}
	
	.three-cols-bg-red{
		background: #fee5e5;
	}
	
	
	.large-table{
		font-size: 12px;
	}
	
	.well{
		padding: 0px;
		margin: 20px 0px 0px 0px;
		background: transparent;
		border: 0px;
		border-radius:0px 0px 0px;
	}
	
	.account-info-table{
		margin-top: -10px;
	}
	
	
	.toggle-line::after{
		display: inline-block;
		content: "+";
		cursor: pointer;
		float:right;
		background: #337ab7;
		color: #fff;
		padding: 0px 5px 0px 5px;
	}
	
	.toggle-line.expanded::after{
		display: inline-block;
		content: "-";
		cursor: pointer;
	}
	
	
	.table-top{
	 	margin-top: 40px;
	}
	
	.table-top .table > tbody > tr > td, .table > tbody > tr > th, .table > tfoot > tr > td, .table > tfoot > tr > th, .table > thead > tr > td, .table > thead > tr > th{
		border: 1px solid #ddd !important;
	}
	
	
	.table-body{
	 	margin-top: -6px;
	}
	
	.table-body {
		border: 1px solid #ddd !important;
	}
	
	.table-body .table{
		margin-bottom: 0px !important;
	}
	
	 
	.inner-table{
		margin-top: 5px;
	}
	
	.inner-table .table{
		margin-bottom: 0px; !important;
	}
	
	.period-end{
			width:8%;
		}
	
	.wrapper{
		margin:0px !important;
	}
	
	.break{
		word-break: break-all;
	}	
	
	.sorttable,
	.sorttable_nosort
	{
		vertical-align: top !important;
		text-align:center;
	}

	.sorttable tr{
	 	text-align:center;
	}

	.sorttable td{
	 	text-align:center;
	}
	
	.large-table-col{
		padding:0px !important;
	}
	
	
	.col-gen-width{
		width: 80px !important;
	}
	
	
	
	.col-left{
		width: 80px !important;
		float: left;
		border-right: 1px solid #ddd;
	}
	
	table{
		table-layout: fixed;
	 }
	
	.table-border{
		border: 1px solid #ddd;
	}
	
	.td-border{
		border-left: 1px solid #ddd;
		text-align:center;
	}
	
	.td-noborder{
		border-top: 0px !important;
	}
	
	.valign{
		vertical-align: middle !important;
	}
	
	 .easy-autocomplete.eac-round input{
		width: 100% !important;
	} 
	
	.blueBg{
		background: #73c1f2 !important;
	}
	
	.greenBg{
		background: #65d96d !important;
	}
	
	.paleYellow{
		background: #fffbbb !important;
	}
	
	/*
	.white-border{
		border-bottom: 1px solid #fff !important;
		vertical-align: middle !important;
		text-align: center;
	}
	
	.white-border h4{
		line-height: 1 !important;
		font-weight: bold;
		padding:0px;
		margin: 5px 0px 0px 25px !important;
		position: absolute;
	}
	
	
	.white-border .current{
		margin-top: 15px !important;
	}
	
	.white-border-1{
		text-align: center;
	}
	
	.white-border-1 h4{
		line-height: 1 !important;
		font-weight: bold;
		padding:0px;
		margin: 0px 0px 0px 25px !important;
		position: absolute;
	}
	
	 */
	
	.td-border .variance{
		line-height: 1 !important;
		font-weight: bold;
		padding: 0px;
		margin:4px 0px 0px 25px;
		text-align: left;
	}
	
	.memos-equipment{
		vertical-align: middle !important;
	}
	
	.period-title{
		width:80%;
		float: left;
	}
	
	.category-title{
		width:80%;
		float: left;
		font-size: 1.27em;
		display:block;
		width: 110px;
		
	}
	
	.category-title a{
		color:#FFF;
		cursor:pointer;
		
	}
	
	.category-title a:hover{
		color:#23527c;	
		cursor:pointer;
	}
	
	.category-note {
		display:block;
		float:right;
		cursor: pointer;
	}
	
	.category-note a{
		color:#FFF;
		cursor:pointer;
		
	}
	
	.category-note a:hover{
		color:#23527c;
		cursor:pointer;	
	}
	
	
	.blue-bg{
		background-color:#80B8FF;
	}
	.modal.modal-wide .modal-dialog {
	  width: 75%;
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
	
	.modalResponsiveTable {
		margin-left: 25px;
		margin-right: 25px;
	}
	
	
	.ajaxRowView .visibleRowEdit, .ajaxRowEdit .visibleRowView { display: none; }
	
	
	.ajax-loading {
	    position: relative;
	}
	.ajax-loading::after {
	    background-image: url("/img/loading.gif");
	    background-position: center top;
	    background-repeat: no-repeat;
	    content: "";
	    display: block;
	    height: 100%;
	    min-height: 100px;
	    position: absolute;
	    top: 0;
	    width: 100%;
	}
	
	.positive{
		font-weight:bold;
		color:blue;
	}
	
	.neutral{
		font-weight:bold;
		color:black;
	}
	
	
	.negative{
		font-weight:bold;
		color:red;	
	}
	
	.zero{
		font-weight:bold;
		color:gray;	
	}
	
	.negative-large-with-bg{
		color:red;
		background: #FDD7D7 !important;
		font-weight:bold;
		font-size:18px;
	}
	
	.positive-large-with-bg{
		color:blue;
		background: #CBE7F9 !important;
		font-weight:bold;
		font-size:18px;
	}
	
	.zero-large-with-bg{
		color:gray;
		font-weight:bold;
		font-size:18px;
		background: #f9f9f9 !important;
	}
	
	.negative-large{
		color:red;
		font-weight:bold;
		font-size:18px;
	}
	
	.positive-large{
		color:blue;
		font-weight:bold;
		font-size:18px;
	}
	
	.zero-large{
		color:gray;
		font-weight:bold;
		font-size:18px;
	}
	
	.company-name-background{
		background: #ccc;
	    border: 1px solid #bbb;
	    border-radius: 7px;
	    font-size: 24px;
	    font-weight: bold;
	    text-align: center;
	    vertical-align: middle;
	    margin-right: 11px;
	    line-height: 2em;
	    text-transform: uppercase;
	    min-height: 150px;
	    padding-top: 5px;
	    letter-spacing: -0.02em;
	} 
	
	.contact-info-background{
		background: #F9F9F9;
	    border: 1px solid #bbb;
	    border-radius: 7px;
	    font-size: 12px;
	    text-align: center;
	    vertical-align: middle;
	    margin-right: 11px;
	    margin-bottom: 0px;
	    margin-top: 4px;
	    line-height: 2em;
	    min-height: 50px;
	    padding-top: 5px;
	} 
	
	
	.center-bold{
		text-align:center;
		font-weight:bold;
		margin-top: 0px !important;
	    margin-bottom: 0px !important;	
	}
	
	.table-valign-middle tr td {
	    vertical-align: middle !important;
	} 
	
	.avg-daily-sales-header{
		background: #5133AB;
		color:#fff;
		text-align:center;
	
	}
	
	table .collapse.in {
		display:table-row;
	}
	
	.no-padding {
		padding:2px !important;
	}
	
	.ROI-tile {
	/*  width: 50%; */
	  display: inline-block;
	  box-sizing: border-box;
	  background: #fff;
	  padding-top: 10px;
	  padding-left: 5px;
	  padding-right: 5px;
	  padding-bottom: 5px;
	  color:#FFF;
	  margin-bottom: 10px;
	}
	
	.ROI-tile .title {
	  margin-top: 0px;
	}
	
	.ROI-tile.red {
	  background: #AC193D;
	  color:#FFF;
	}
	
	.ROI-tile.red:hover {
	  background: #7f132d;
	  color:#FFF;
	}
	
	.ROI-tile.blue {
	  background: #2672EC;
	  color:#FFF;
	}
	
	.ROI-tile.blue:hover {
	  background: #125acd;
	  color:#FFF;
	}
	
	/*Business Card Css */
	.business-card {
		background: #F9F9F9;
		border: 1px solid #bbb;
		border-radius: 7px;
		background: #f8f8f8;
		padding: 10px;
		margin-bottom: 10px;
	    vertical-align: middle;
	    margin-right: 11px;
	    margin-bottom: 0px;
	    margin-top: 4px;
	    line-height: 2em;
	    min-height: 50px;
	    padding-top: 5px;
	    text-align:center;
		
	}
	.media-heading{
		color: #666666;
		font-size: 14px;
		margin-top: 8px;
		margin-bottom: 5px;
	  
	}
	
	.mail {
	  font-size: 12px;
	 }
	 
	.media-body{
	    border-left: 1px solid #999;
	    width:75%;
	    line-height: 1.5em;
	 }
	 
	.media-left{
		vertical-align: middle;
		font-weight:bold;
	 }
	

	.yes-unread-notes-button {
	  text-align: center;
	  color: white;
	  border: none;
	  font-size:1.1em;
	  background: #ffc12c;
	  cursor: pointer;
	  box-shadow: 0 0 0 0 rgba(#ffc12c, .5);
	  -webkit-animation: pulse 2.5s infinite;
	}
	.yes-unread-notes-button:hover {
	  -webkit-animation: none;
	}
	
	@-webkit-keyframes pulse {
	  0% {
	    @include transform(scale(.9));
	  }
	  70% {
	    @include transform(scale(1));
	    box-shadow: 0 0 0 25px rgba(#5a99d4, 0);
	  }
	    100% {
	    @include transform(scale(.9));
	    box-shadow: 0 0 0 0 rgba(#5a99d4, 0);
	  }
	}	
	
	.no-unread-notes-button {
	  text-align: center;
	  color: white;
	  border: none;
	  font-size:1.1em;
	  background: #ffc12c;
	  cursor: pointer;
	  box-shadow: 0 0 0 0 rgba(#ffc12c, .5);
	  -webkit-animation: none;
	}
	.no-unread-notes-button:hover {
	  -webkit-animation: none;
	}
	
</style>

<%
If CustForDetail <> "" Then 
	Response.Write("<div id=""PleaseWaitPanel"">")
	Response.Write("<br><br><strong>Analyzing " & CustName & ", please wait...</strong><br><br>")
	Response.Write("<img src='" & baseURL & "/img/loading.gif'/>")
	Response.Write("</div>")
	Response.Flush()
	
	If Session("DebugSpeed") = True Then Response.Write("<br><br>2:" & Now())	
	Call PrepareTmpTable (CustForDetail , PeriodSeqBeingEvaluated, PeriodBeingEvaluated)
	If Session("DebugSpeed") = True Then Response.Write("<br><br>3:" & Now())

	TotalSalesAllCatsCurrentPeriod = GetCurrent_PostedTotal_ByCust(CustForDetail,PeriodSeqBeingEvaluated) + GetCurrent_UnPostedTotal_ByCust(CustForDetail,PeriodSeqBeingEvaluated)

	TotalClostAllCatsCurrentPeriod = GetCurrent_PostedTotalCost_ByCust() + GetCurrent_UnPostedTotalCost_ByCust()

	SQL_G="SELECT SUM(TotalSales)/3 AS ThreePeriodAverage, SUM(TotalCost)/3 AS ThreePeriodAverageCost FROM zCatAnalByPeriod_SingleCustomer_" & Session("UserNo") & "_" & CustForDetail
	SQL_G = SQL_G & " WHERE (ThisPeriodSequenceNumber < " & PeriodSeqBeingEvaluated & ") AND (ThisPeriodSequenceNumber > " & PeriodSeqBeingEvaluated - 4 & ") "
	Set	rsGScreen= cnn8.Execute(SQL_G)
			
	If NOT rsGScreen.EOF Then 
		TotalSalesAllCats3PeriodAverage = rsGScreen("ThreePeriodAverage") 
		TotalCostAllCats3PeriodAverage = rsGScreen("ThreePeriodAverageCost") 
	Else
		TotalSalesAllCats3PeriodAverage = 0
		TotalCostAllCats3PeriodAverage = 0
	End If
	
	SQL_G="SELECT SUM(TotalSales)/12 AS TwelvePeriodAverage, SUM(TotalCost)/12 AS TwelvePeriodAverageCost FROM zCatAnalByPeriod_SingleCustomer_" & Session("UserNo") & "_" & CustForDetail
	SQL_G = SQL_G & " WHERE (ThisPeriodSequenceNumber < " & PeriodSeqBeingEvaluated & ") AND (ThisPeriodSequenceNumber > " & PeriodSeqBeingEvaluated - 13 & ") "
	Set	rsGScreen= cnn8.Execute(SQL_G)
			
	If NOT rsGScreen.EOF Then 
		TotalSalesAllCats12PeriodAverage = rsGScreen("TwelvePeriodAverage")
		TotalCostAllCats12PeriodAverage = rsGScreen("TwelvePeriodAverageCost")
	Else
		TotalSalesAllCats12PeriodAverage = 0
		TotalCostAllCats12PeriodAverage = 0
	End If
	
	SQL_G="SELECT Sum(PriorPeriod1Sales+PriorPeriod2Sales+PriorPeriod3Sales) As Tot3, "
	
	
	SQL_G = SQL_G & " Sum(PriorPeriod1Sales+PriorPeriod2Sales+PriorPeriod3Sales+PriorPeriod4Sales+PriorPeriod5Sales+PriorPeriod6Sales+ "
	SQL_G = SQL_G & " PriorPeriod7Sales+PriorPeriod8Sales+PriorPeriod9Sales+PriorPeriod10Sales+PriorPeriod11Sales+PriorPeriod12Sales) As Tot12 "
	SQL_G = SQL_G & "  FROM zCatAnalByPeriod_SingleCustomer_" & Session("UserNo") & "_" & CustForDetail
	SQL_G = SQL_G & " WHERE ThisPeriodSequenceNumber = " & PeriodSeqBeingEvaluated 
	Set	rsGScreen= cnn8.Execute(SQL_G)

	ActualTotalSales3PriorClosedPeriods  =  rsGScreen("Tot3")

	ActualTotalSales12PriorClosedPeriods =  rsGScreen("Tot12")
		
	
If Session("DebugSpeed") = True Then Response.Write("<br><br>4:" & Now())

End If
%>

<!-- row !-->
<form method="POST" name="frmCatAnalByPeriodSingleCust" id="frmCatAnalByPeriodSingleCust" action="CatAnalByPeriod_SingleCustomer.asp">		    
	
	
 	
	<div class='table-responsive table-top'>
	
		 <table class='table table-condensed'>
	
		 	<tbody>
	
		 		<tr>

		 			<td style="padding: 0px; border-top: 0px !important; border-bottom: 0px !important; border-left: 0px !important; border-right: 0px !important;" rowspan="2">			    
		        		<!-- select company !-->
							<input id="txtCustIDMainSearch" name="txtCustIDMainSearch">
							<input type="hidden" id="txtCustIDToPassMainSearch" name="txtCustIDToPassMainSearch" value="<%= CustForDetail %>" >
							<i id="searchIcon" class="fa fa-search fa-2x"></i>
						<!-- eof select company !-->
						
		 				<h3 class="company-name-background">
		 					<%=CustName %>
		 					<% If CustForDetail <> "" Then %>
			 					<br>Acct#&nbsp;<%=CustForDetail%>
			 					<br>Period&nbsp;<%= Left(PeriodBeingEvaluated ,Instr(PeriodBeingEvaluated ,"-")-2) %>&nbsp;
			 					FY'<%= Right(PeriodBeingEvaluated ,Len(PeriodBeingEvaluated) - Instr(PeriodBeingEvaluated ,"-")-3) %>
			 				<% Else %>
			 					Please select an account to analyze
			 				<% End If %>
		 				</h3>
				 				
						<div class="business-card">
		                	<div class="media">
		                		<% If CustForDetail <> "" Then %>
				                    <div class="media-left">
				                        Contact Info
				                    </div>
				                    <div class="media-body">
				                        <h2 class="media-heading"><%= Cust_ContactName %></h2>
				                        <div class="mail"><a href="<%= Cust_ContactEmail %>"><%= Cust_ContactEmail %></a></div>
				                        <div class="mail"><%= Cust_ContactTel %></div>
				                    </div>
				                 <% End If %>
			                </div>
		                </div>		 				
						
					</td>
		 			
		 			<td width="25%" rowspan="2">
		 				 <div class="table-striped table-condensed table-hover account-info-table inner-table">
			             <table class="table table-striped table-condensed table-hover" >
			             	<tr>
					            <td width="60%">
		  		 					<input type="radio" class="active_radio"  name="optWhatToShow" id="optWhatToShow3Periods" value="3Periods" <% If VarianceBasis = "3Periods" Then Response.Write("checked='checked'") %>>&nbsp;Base variance on 3 period average <br>
				 					<input type="radio"  class="active_radio" name="optWhatToShow" id="optWhatToShow12Periods" value="12Periods" <% If VarianceBasis = "12Periods" Then Response.Write("checked") %>>&nbsp;Base variance on 12 period average<br>
				 					<input type="checkbox" name="chkShowZeroDollarCategories" id="chkShowZeroDollarCategories" <% If ShowZeros = 1 Then Response.Write("checked") %>>&nbsp;Show $0 categories<br>
				 					<!--<input type="checkbox" name="chkShowGPPercent" id="chkShowGPPercent" <% If ShowGPPercent = 1 Then Response.Write("checked") %> checked="checked">&nbsp;Show GP%-->
				 				</td>
				 				<td width="40%" style="text-align:center;vertical-align:bottom;">
				 				
									<% 
									If CustForDetail <> "" then
									
										TotalEquipmentValue = GetTotalValueOfEquipmentForCustomer(CustForDetail)
										LCPGP = TotalSalesByPeriodSeq(PeriodSeqBeingEvaluated,CustForDetail) - TotalCostByPeriodSeq(PeriodSeqBeingEvaluated,CustForDetail)
										If LCPGP <> 0 Then ROI =  Round(TotalEquipmentValue/LCPGP,1) Else ROI = 0
										P3PGP = (TotalSalesAllCats3PeriodAverage- TotalCostAllCats3PeriodAverage)
										If P3PGP <> 0 Then ROI3PA =  TotalEquipmentValue/P3PGP  Else ROI3PA = 0
										
										If TotalSalesPeriodBeingEvaluated<> 0 then GP = ((TotalSalesPeriodBeingEvaluated - TotalCostByPeriodSeq(PeriodSeqBeingEvaluated,CustForDetail))/TotalSalesPeriodBeingEvaluated)*100
										
										'Response.Write("TotalEquipmentValue" & TotalEquipmentValue & "<br>")
										'Response.Write("S:" & TotalSalesByPeriodSeq(PeriodSeqBeingEvaluated,CustForDetail) & ":<br>")
										'Response.Write("C:" & TotalCostByPeriodSeq(PeriodSeqBeingEvaluated,CustForDetail)& ":<br>")
										'Response.Write("LCPGP:" & LCPGP  & ":<br>")
										
										If TotalEquipmentValue <> 0 Then ' Only if the customer has equipment
											If TotalSalesByPeriodSeq(PeriodSeqBeingEvaluated,CustForDetail)  > 0 Then
												If (cInt(ROI)) < 10 Then %>
													<div class="ROI-tile blue">
														<h5 class="title" style="margin-bottom: 3px;">ROI <%= Round(ROI,1) %></h5>
														<small>LCP</small>
													</div>	
												<% Else %>
													<div class="ROI-tile red">
														<h5 class="title" style="margin-bottom: 3px;">ROI <%= Round(ROI,1) %></h5>
														<small>LCP</small>
													</div>	
												<% End If 
											ElseIf TotalSalesByPeriodSeq(PeriodSeqBeingEvaluated,CustForDetail)  < 0 Then %>
												&nbsp;
											<%Else %>
													<div class="ROI-tile red">
														<h5 class="title" style="margin-bottom: 3px;">No Sales</h5>
														<small>In LCP</small>
													</div>	
											<% End If
											If (cInt(ROI3PA)) < 10 Then %>
												<div class="ROI-tile blue">
													<h5 class="title" style="margin-bottom: 3px;">ROI <%= Round(ROI3PA,1) %></h5>
													<small>3Pavg</small>
												</div>	
											<% Else %>
												<div class="ROI-tile red">
													<h5 class="title" style="margin-bottom: 3px;">ROI <%= Round(ROI3PA,1) %></h5>
													<small>3Pavg</small>
												</div>	
											<% End If 
										End If
				 					 End If %>
				 					 
				 					<button type="button" class="btn btn-success btn-md"><i class="fa fa-eye" aria-hidden="true"></i> Add To Watch List</button>
				 				</td>
							</tr>
							<%		
								If CustForDetail <> "" Then
									LastInvoiceFromWebDate = GetLastInvoiceFromWebDate(CustForDetail)
									
									LastInvoiceFromWebDate = GetLastWebOrderDateFromOCSAccess(CustForDetail)
									
									Response.Write("<tr>")
									Response.Write("<td colspan='4'>Web User:&nbsp;")
									If CustHasWebUserID(CustForDetail) <> True Then
										Response.Write("<strong>No</strong>")
									Else
										Response.Write("<strong>Yes</strong>")
									End If
									
									Response.Write("&nbsp;&nbsp;&nbsp;&nbsp;Last Web Order:&nbsp;")
									If LastInvoiceFromWebDate <> "" Then
										Response.Write("<strong>" & FormatDateTime(LastInvoiceFromWebDate,2)  & "</strong></td>")
									Else
										Response.Write("&nbsp;</td>")
									End If
									Response.Write("</tr>")
								End If
								
							%>
							<tr><td colspan="4">&nbsp;</td></tr>
						</table>
						</div>
	 					<!-- eof content -->
	 					
					<!-- Average Daily Sales !-->
		 			<!--<td>-->
		 			<% If CustForDetail <> "" Then %>
						 <div class="table-striped table-condensed table-hover account-info-table inner-table">
			             <table class="table table-striped table-condensed table-hover" >

			
							<tr>
							<td colspan="4" class="avg-daily-sales-header">Average Daily Sales</td>
							</tr>	
							
							<tr>
								<%
								ADS_Current = TotalSalesAllCatsCurrentPeriod 
								If ADS_Current = "" Then
									ADS_Current = 0
								Else
									ADS_Current = ADS_Current / WorkDaysSoFar
								End If
								%>
								
								<td>Day Impact</td>
								
								<% ImpactDays = (WorkDaysIn3PeriodBasis/3)- WorkDaysInLastClosedPeriod
								DayImpact = ImpactDays  * (TotalSalesByPeriodSeq(PeriodSeqBeingEvaluated,CustForDetail)/WorkDaysInLastClosedPeriod)
								DayImpact = Round(DayImpact,2)
									
								 If DayImpact > 0 Then %>
									<td align="right"><span class="neutral"><%= FormatCurrency(DayImpact,0) %></span></td>
								<% ElseIf DayImpact < 0 Then %>
									<td align="right"><span class="neutral"><%= FormatCurrency(DayImpact,0) %></span></td>
								<% Else %>
									<td align="right"><span class="neutral"><%= FormatCurrency(DayImpact,0) %></span></td>
								<% End If %>
							
								<%
								ADS_12PA = TotalSalesAllCats12PeriodAverage 
								If ADS_12PA = "" Then
									ADS_12PA = 0
								Else
									ADS_12PA = ADS_12PA / (WorkDaysIn12PeriodBasis /12)
								End If
								
								ADS_3PA = TotalSalesAllCats3PeriodAverage 
								If ADS_3PA = "" Then
									ADS_3PA = 0
								Else
									ADS_3PA = ADS_3PA / (WorkDaysIn3PeriodBasis /3)
								End If

								ADS_LastClosed = TotalSalesByPeriodSeq(PeriodSeqBeingEvaluated,CustForDetail)
								If ADS_LastClosed = "" Then
									ADS_LastClosed = 0
								Else
									ADS_LastClosed = ADS_LastClosed / WorkDaysInLastClosedPeriod 
								End If

								ADS_Current = TotalSalesAllCatsCurrentPeriod 
								If ADS_Current = "" Then
									ADS_Current = 0
								Else
									ADS_Current = ADS_Current / WorkDaysSoFar
								End If
								%>


								<td>Daily Diff</td>
								
								<% DailyDiff = ADS_LastClosed -  ADS_3PA 
								
								If DailyDiff  > 0 Then %>
									<td align="right"><span class="positive"><%= FormatCurrency(DailyDiff ,0) %></span></td>
								<% ElseIf DailyDiff < 0 Then %>
									<td align="right"><span class="negative"><%= FormatCurrency(DailyDiff ,0) %></span></td>
								<% Else %>
									<td align="right"><span class="neutral"><%= FormatCurrency(DailyDiff ,0) %></span></td>
								<% End If %>
							
							</tr>
		
							<tr>
								
								<td>LCP&nbsp;<small>(<%=WorkDaysInLastClosedPeriod%>&nbsp;days</small>)</td>


														
								<td align="right"><span class="neutral"><%= FormatCurrency(ADS_LastClosed,0) %></span></td>

								
								
								<td>3Pavg&nbsp;<small>(<%=Round((WorkDaysIn3PeriodBasis/3),1)%>)</small></td>
								
								<% If ADS_3PA > 0 Then %>
									<td align="right"><span class="neutral"><%= FormatCurrency(ADS_3PA,0) %></span></td>
								<% ElseIf ADS_3PA < 0 Then %>
									<td align="right"><span class="neutral"><%= FormatCurrency(ADS_3PA,0) %></span></td>
								<% Else %>
									<td align="right"><span class="neutral"><%= FormatCurrency(ADS_3PA,0) %></span></td>
								<% End If %>
								
								
							</tr>			
							<tr>
								
								<td>Current&nbsp;<small>(<%=WorkDaysSoFar%>)</small></td>
									
								<% If ADS_Current > 0 Then %>
									<td align="right"><span class="neutral"><%= FormatCurrency(ADS_Current,0) %></span></td>
								<% ElseIf ADS_Current < 0 Then %>
									<td align="right"><span class="neutral"><%= FormatCurrency(ADS_Current,0) %></span></td>
								<% Else %>
									<td align="right"><span class="neutral"><%= FormatCurrency(ADS_Current,0) %></span></td>
								<% End If %>
							
								<%
								ADS_12PA = TotalSalesAllCats12PeriodAverage 
								If ADS_12PA = "" Then
									ADS_12PA = 0
								Else
									ADS_12PA = ADS_12PA / (WorkDaysIn12PeriodBasis /12)
								End If
								%>
								<td>12Pavg&nbsp;<small>(<%=Round(WorkDaysIn12PeriodBasis/12,1)%>)</small></td>
								
								<% If ADS_12PA > 0 Then %>
									<td align="right"><span class="neutral"><%= FormatCurrency(ADS_12PA,0) %></span></td>
								<% ElseIf ADS_12PA< 0 Then %>
									<td align="right"><span class="neutral"><%= FormatCurrency(ADS_12PA,0) %></span></td>
								<% Else %>
									<td align="right"><span class="neutral"><%= FormatCurrency(ADS_12PA,0) %></span></td>
								<% End If %>
							
							</tr>
							
							

						</table>
						</div>
					<% End If%>
		 			<!--</td>-->
					<!-- eof Average Daily Sales !-->		 				 					

		 			</td>
		 			
		 			
		 			<!-- If you want to center the table vertically  align="center" class="valign" -->
		 			<td width="25%">
						<% If CustForDetail <> "" Then %>
							 <div class="table-striped table-condensed table-hover account-info-table inner-table">
				             <table class="table table-striped table-condensed table-hover" >
								<%		
									Response.Write("<tr>")
									Response.Write("<td width='30%'>Referral</td>")
									Response.Write("<td width='70%' colspan='3'><strong>" & Cust_ReferralCode & "</strong></td>")
									Response.Write("</tr>")
								
									Response.Write("<tr>")
									Response.Write("<td width='30%'>Customer Type</td>")
									Response.Write("<td width='70%' colspan='3'><strong>" & Cust_Type & "</strong></td>")
									Response.Write("</tr>")
									
									MGPTerm = "MES" 
									MGPVerbiage = "Not Set"

									' Determine what CCS is going to call it
									If Cust_MGPSales > 0 Then
										If cint(Cust_MGP) = 1 Then
											MGPTerm = "MES" 
											MGPVerbiage = "Internal"
										Else
											MGPTerm = "MCS" 
											MGPVerbiage = "Contracted Minimum"
										End If
									End If
									

									If Cust_MGPSales < 1 Then
									
										Response.Write("<tr>")
											Response.Write("<td width='40%'>" & MGPTerm & " vs LCP</td>")
											%><td width="10%">&nbsp;</td><%
											Response.Write("<td width='40%'>" & MGPTerm & " vs 3Pavg</td>")
											%><td width="10%">&nbsp;</td><%
										Response.Write("</tr>")
										
									Else
									
										Response.Write("<tr>")
										
											Response.Write("<td width='40%'>" & MGPTerm & " vs LCP</td>")
											
											VarianceLCP = Round(TotalSalesByPeriodSeq(PeriodSeqBeingEvaluated,CustForDetail)-Cust_MGPSales,0)
											
											If VarianceLCP > 0 Then %>
												<td width="10%"><span class="positive"><%= FormatCurrency(VarianceLCP,0) %></span></td>
											<% ElseIf VarianceLCP < 0 Then %>
												<td width="10%"><span class="negative"><%= FormatCurrency(VarianceLCP,0) %></span></td>
											<% Else %>
												<td width="10%"><span class="zero"><%= FormatCurrency(VarianceLCP,0) %></span></td>
											<% End If 
																					
											Response.Write("<td width='40%'>" & MGPTerm & " vs 3Pavg</td>")
											Variance3PAvg = Round(TotalSalesAllCats3PeriodAverage-Cust_MGPSales,0)
											
											If Variance3PAvg > 0 Then %>
												<td width="10%"><span class="positive"><%= FormatCurrency(Variance3PAvg,0) %></span></td>
											<% ElseIf Variance3PAvg < 0 Then %>
												<td width="10%"><span class="negative"><%= FormatCurrency(Variance3PAvg,0) %></span></td>
											<% Else %>
												<td width="10%"><span class="zero"><%= FormatCurrency(Variance3PAvg,0) %></span></td>
											<% End If 
										
										Response.Write("</tr>")
									End If									
									
									Response.Write("<tr>")
									Response.Write("<td width='30%'>" & MGPTerm & " Sales</td>")
									Response.Write("<td width='35%'><strong>" & Cust_MGPSales & "</strong></td>")

									Response.Write("<td width='35%' colspan='2'><strong>" & MGPVerbiage  & "</strong></td>")									

									Response.Write("</tr>")
									Response.Write("<tr>")
										Response.Write("<td width='30%'>Last Buy Date</td>")
										Response.Write("<td width='70%' colspan='3'><strong>" & Cust_LastBuy & "</strong></td>")
									Response.Write("</tr>")
								%>
							</table>
							</div>
			
							 <!-- content -->
							 <div class="collapse" id="collapseExample">
								<div class="well">
									Your content goes here.
								</div>
							</div>
							<!-- eof content -->
					 	<% End If %>
					</td>
	
					<td width="25%">
						 <div class="table-striped table-condensed table-hover account-info-table inner-table">
			             <table class="table table-striped table-condensed table-hover" >
							<%		
								Response.Write("<tr>")
								Response.Write("<td width='30%'>" & GetTerm("Sales Person 1") & "</td>")
								Response.Write("<td width='70%'><strong>" & Cust_SalesPerson1 & "</strong></td>")
								Response.Write("</tr>")
								
								Response.Write("<tr>")
								Response.Write("<td width='30%'>" & GetTerm("Sales Person 2") & "</td>")
								Response.Write("<td width='70%'><strong>" & Cust_SalesPerson2 & "</strong></td>")
								Response.Write("</tr>")
								
								Response.Write("<tr>")
								Response.Write("<td width='30%'>Install Date</td>")
								Response.Write("<td width='70%'><strong>" & Cust_InstallDate & "</strong></td>")
								Response.Write("</tr>")

								Response.Write("<td width='30%'>Chain ID / Name</td>")
								If Cust_ChainID <> "" Then 
									Response.Write("<td width='70%'><strong>" & Cust_ChainID & " / " & Cust_ChainName & "</strong></td>")
								Else
									Response.Write("<td width='70%'><strong>&nbsp;</strong></td>")
								End If
								Response.Write("</tr>")
							%>
						</table>
						</div>
	 					<!-- eof content -->
	
		 			</td>
		 		</tr>
	
		 		<tr>		

		 			
		 			
		 			<td>
						 <div class="table-striped table-condensed table-hover account-info-table inner-table">
			             <table class="table table-striped table-condensed table-hover" >
							<%		
								EverGreen_EndDate = ""
								EverGreen_Term = ""
								If Cust_ContEnd <> "" Then
									If cDate(DateAdd("d",90,Now())) > cDate(Cust_ContEnd)  Then ' Contract expired, show evergreen
										EverGreen_EndDate = Cust_ContEnd
										'Keep going unti we get a non-expired date
										Do While cDate(DateAdd("d",90,Now())) > cDate(EverGreen_EndDate)
											EverGreen_EndDate = DateAdd("m",12,EverGreen_EndDate)
										Loop
										EverGreen_Term = "12 months"
									End If
								End If
								
								Response.Write("<tr>")
								Response.Write("<td width='33%'>Contract Start</td>")
								Response.Write("<td width='67%' colspan='2'><strong>" & Cust_ContStart & "</strong></td>")
								Response.Write("</tr>")
								
								Response.Write("<tr>")
								Response.Write("<td width='33%'>Contract End</td>")
								Response.Write("<td width='33%'><strong>" & Cust_ContEnd & "</strong></td>")
								If EverGreen_EndDate <> "" Then
									Response.Write("<td width='34%' style='color:green'><strong>" & EverGreen_EndDate & "</strong></td>")
								Else
									Response.Write("<td width='34%'><strong>" & EverGreen_EndDate & "</strong></td>")								
								End If

								Response.Write("</tr>")
								
								Response.Write("<tr>")
								Response.Write("<td width='33%'>Contract Term</td>")
								Response.Write("<td width='33%'><strong>" & Cust_ContTerm & "</strong></td>")
								If EverGreen_Term <> "" Then 
									Response.Write("<td width='34%' style='color:green'><strong>" & EverGreen_Term & "</strong></td>")
								Else
									Response.Write("<td width='34%'><strong>" & EverGreen_Term & "</strong></td>")
								End If								
								Response.Write("</tr>")
							%>
						</table>
						</div>
	 					<!-- eof content -->
	 					
					</td>
					
		 			<td class="memos-equipment">

		 				<table width="100%">
		 					<% If CustForDetail <> "" Then %>
			 					<tr>
			 						<td valign="middle" align="center">										

										<div class="container-fluid">     
										  <div class="row">
										    <div class="col-sm-6">
													
										
											<%								
											'Allow for a note here as a way to put in a note for the customer in general
											'Use -1 as the category number
											
											%>
											
											<% If UserHasAnyUnviewedNotes(CustForDetail) Then %>
												<button type="button" class="btn btn-warning btn-block yes-unread-notes-button" data-toggle="modal" data-target="#modalEditCustomerNotes" data-category-id="-1" data-cust-id="<%= CustForDetail %>"><span class="fa fa-file-text-o faa-pulse animated" aria-hidden="true"></span> Client Notes</button>
											<% Else %>
												<button type="button" class="btn btn-warning btn-block no-unread-notes-button" data-toggle="modal" data-target="#modalEditCustomerNotes" data-category-id="-1" data-cust-id="<%= CustForDetail %>"><span class="fa fa-file-text-o" aria-hidden="true"></span> Client Notes</button>
											<% End If %>
											
										    </div>
										    <div class="col-sm-6">
										    <%
										    	TotalSalesPeriodBeingEvaluated  = TotalSalesByPeriodSeq(PeriodSeqBeingEvaluated,CustForDetail)
												LCPGP = TotalSalesPeriodBeingEvaluated - TotalCostByPeriodSeq(PeriodSeqBeingEvaluated,CustForDetail)

										    If TotalEquipmentValue <> 0 Then%>
												<button type="button" class="btn btn-success btn-block" data-toggle="modal" data-target="#modalEquipmentVPC" data-cust-id="<%= CustForDetail %>" data-lcp-gp="<%= LCPGP %>">Equipment</button>    
											<%Else%>
												<h6 class="title"><%=GetTerm("Customer")%> has no equipment</h6>
											<%End If%>
										    </div>
										  </div>
										    <div class="row" style="margin-top: 10px;">
										    <div class="col-sm-6">
										      <button type="button" class="btn btn-primary btn-block" data-toggle="modal" data-target="#modalServiceTickets">Service Tickets</button>
										    </div>
										    <div class="col-sm-6">
										      <button type="button" class="btn btn-danger btn-block" data-toggle="modal" data-target="#modalXPage">X Page</button>    
										    </div>
										  </div>
										</div>		 		
														
			 						</td>
			 					</tr>
			 					
		 						<% Else %>
		 						
		 							<tr><td>&nbsp;</td></tr>
		 							
		 						<% End If %>
		 					
		 				</table>
		 				 

		 			</td>
		 			
				</tr>
				
				
	
			</tbody>
		 </table>
		</div>
	</div> 
	
</form>		 

<%If CustForDetail <> "" Then 
If Session("DebugSpeed") = True Then Response.Write("<br><br>5:" & Now())

'***************************************************************************************
'BUILD THE MASTER REPORT ARRAY
'***************************************************************************************
Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rsBuildArray = Server.CreateObject("ADODB.Recordset")

SQLBuildArray = "SELECT TOP 1 * FROM zCatAnalByPeriod_SingleCustomer_" & Session("UserNo") & "_" & CustForDetail

Set rsBuildArray = cnn8.Execute(SQLBuildArray )

'Array needs 1 for each cat
Redim MasterReportArray(22,rsBuildArray.Fields.Count+1)
MaterArrayFieldCount = rsBuildArray.Fields.Count+1
Set rsBuildArray  = Nothing
'***************************************************************************************
'END BUILD THE MASTER REPORT ARRAY
'***************************************************************************************


'***************************************************************************************
'BUILD THE MASTER ORDER BY CLAUSE HERE
'****************************************************************************************
Set rsCatOrder = Server.CreateObject("ADODB.Recordset")
Redim MasterReportArrayOrder(22)

SQLCatOrder = "SELECT * FROM zCatAnalByPeriod_SingleCustomer_" & Session("UserNo")  & "_" & CustForDetail & " WHERE ThisPeriodSequenceNumber = " & PeriodSeqBeingEvaluated  

If MUV_READ("VarianceBasis") = "3Periods" Then
	SQLCatOrder = SQLCatOrder & " ORDER BY DiifThisPeriodVSLast3Dollars "
Else
	SQLCatOrder = SQLCatOrder & " ORDER BY DiifThisPeriodVSLast12Dollars "	
End If

Set rsCatOrder = cnn8.Execute(SQLCatOrder)

If Not rsCatOrder.EOF Then

	CatOrderByClauseCustom = " CASE Category "
	SortCount = 0

	Do While NOT rsCatOrder.EOF

		MasterReportArrayOrder(SortCount) = rsCatOrder("Category")
		
		CatOrderByClauseCustom = CatOrderByClauseCustom & " WHEN " & rsCatOrder("Category") & " THEN " & Trim(SortCount) & " "
		SortCount = SortCount + 1
		
		'Insert into the mater report array
		FieldCount = 0
		For Each item in rsCatOrder.Fields
			MasterReportArray(rsCatOrder("Category"),FieldCount) = rsCatOrder.Fields(FieldCount)
			FieldCount = FieldCount + 1
		Next

		rsCatOrder.MoveNext
	Loop
	
	CatOrderByClauseCustom = CatOrderByClauseCustom & " END "

End If

Set rsCatOrder = Nothing
If Session("DebugSpeed") = True Then Response.Write("<br><br>6:" & Now())

If Session("DebugSpeed") = True Then 
	For x = 0 to Ubound(MasterReportArray)-1
		For z = 0 to MaterArrayFieldCount 
			Response.Write(MasterReportArray(x,z) & ":")
		Next
		Response.Write("<br>")
	Next
End If

%>
<!-- row -->
<div class="col-lg-12">
	<div class='table-striped table-condensed table-responsive table-body'>
		<table class='table table-condensed large-table table-valign-middle'>


<%
'********************************
' Duplicate Category Heading Row
'********************************
				Response.Write("<tr>")
				Response.Write("<th scope='col' class='sorttable_nosort' width='80'>&nbsp;</th>")
				Response.Write("<th scope='col' class='break sorttable ' valign='top' width='60'>&nbsp;</th>")
				Response.Write("<th scope='col' class='break sorttable ' valign='top' width='60'>&nbsp;</th>")
				
				'********************************************************
				' GP% Column Show/Hide
				'********************************************************
				
				If ShowGPPercent = 1 Then
					Response.Write("<th scope='col' class='break sorttable'valign='top' width='30'")
					Response.Write("<span class='period-title gp-percent'>GP%</span></th>")
				End If
				
If Session("DebugSpeed") = True Then Response.Write("<br><br>6.5:" & Now())


				SQL_G="SELECT Category, CategoryNameGetTerm FROM zCatAnalByPeriod_SingleCustomer_" & Session("UserNo") & "_" & CustForDetail & " WHERE ThisPeriodSequenceNumber = " & PeriodSeqBeingEvaluated  
				SQL_G = SQL_G & " ORDER BY " & CatOrderByClauseCustom 

				Set	rsGScreen = cnn8.Execute(SQL_G)
				
				If not rsGScreen.eof Then
					
					'Category names get written below
					
					Do
						If ShowZeros = 1 Then
						
							Response.Write("<th scope='col' class='break sorttable blue-bg' valign='top' width='60'>")
							
							If Session("SkipGetTerm") <> True Then
								Response.Write("<span class='category-title'><a data-toggle='modal' data-target='#modalCategoryVPC' data-category-id='" & rsGScreen("Category") & "' data-cust-id='" & CustForDetail & "' data-period-seq='" & PeriodSeqBeingEvaluated & "' data-variance-basis='" & VarianceBasis & "'>" & GetTerm(GetCategoryByID(rsGScreen("Category"))) & "</a></span>")
							Else
								Response.Write("<span class='category-title'><a data-toggle='modal' data-target='#modalCategoryVPC' data-category-id='" & rsGScreen("Category") & "' data-cust-id='" & CustForDetail & "' data-period-seq='" & PeriodSeqBeingEvaluated & "' data-variance-basis='" & VarianceBasis & "'>" & rsGScreen("CategoryNameGetTerm") & "</a></span>")									
							End If
							
							Response.Write("</th>")
							
						Else
							' Cat total handles the current period, which must be included now
							If Session("NewPostLogic") <> True Then CatTotal = GetCurrent_PostedTotal_ByCustByCat(CustForDetail,PeriodSeqBeingEvaluated ,rsGScreen("Category")) +GetCurrent_UnPostedTotal_ByCustByCat(CustForDetail,PeriodSeqBeingEvaluated ,rsGScreen("Category"))						
							If Session("NewPostLogic") = True Then CatTotal = PostUnPostCatArray(rsGScreen("Category"),0) + PostUnPostCatArray(rsGScreen("Category"),1)
							
							If GetTotalSalesForCategByPeriodRange(rsGScreen("Category"),PeriodSeqBeingEvaluated - 11,PeriodSeqBeingEvaluated,CustForDetail) + CatTotal <> 0 Then
									
								Response.Write("<th scope='col' class='break sorttable blue-bg' valign='top' width='60'>")
								
								If Session("SkipGetTerm") <> True Then
									Response.Write("<span class='category-title'><a data-toggle='modal' data-target='#modalCategoryVPC' data-category-id='" & rsGScreen("Category") & "' data-cust-id='" & CustForDetail & "' data-period-seq='" & PeriodSeqBeingEvaluated & "' data-variance-basis='" & VarianceBasis & "'>" & GetTerm(GetCategoryByID(rsGScreen("Category"))) & "</a></span>")
								Else
									Response.Write("<span class='category-title'><a data-toggle='modal' data-target='#modalCategoryVPC' data-category-id='" & rsGScreen("Category") & "' data-cust-id='" & CustForDetail & "' data-period-seq='" & PeriodSeqBeingEvaluated & "' data-variance-basis='" & VarianceBasis & "'>" & rsGScreen("CategoryNameGetTerm") & "</a></span>")									
								End If
								
								Response.Write("</th>")
						
							End If
						End If
						
						rsGScreen.movenext
					loop until rsGScreen.eof
				End If
				Response.Write("</tr>")
'************************************
' End Duplicate Category Heading Row
'************************************
If Session("DebugSpeed") = True Then Response.Write("<br><br>7:" & Now())


%>
					<tr>
					
						<td width="80"><h4 class="center-bold">Last Closed</h4></td>
						<td width="60" class="td-border"><strong>Period&nbsp;<%=GetPeriodBySeq(PeriodSeqBeingEvaluated)%></strong></td>
						<td width="60" align="right" class="td-border"><% = FormatCurrency(TotalSalesByPeriodSeq(PeriodSeqBeingEvaluated,CustForDetail),0)%></td>
						
						<%
						If ShowGPPercent = 1 Then
							TotalSalesPeriodBeingEvaluated  = TotalSalesByPeriodSeq(PeriodSeqBeingEvaluated,CustForDetail)
							If TotalSalesPeriodBeingEvaluated = "" or TotalSalesPeriodBeingEvaluated = 0 Then
								Response.Write("<td align='right' class='td-border'>" & FormatCurrency(0,0) & "</td>")
							Else
								GP = ((TotalSalesPeriodBeingEvaluated - TotalCostByPeriodSeq(PeriodSeqBeingEvaluated,CustForDetail))/TotalSalesPeriodBeingEvaluated)*100
								Response.Write("<td align='right' class='td-border'>" & Round(GP,1) & "%</td>")
							End if
						End If

						' Current Last Closed
							
							SQL_G="SELECT TotalSales AS GroupTot, Category FROM zCatAnalByPeriod_SingleCustomer_" & Session("UserNo") & "_" & CustForDetail & " WHERE ThisPeriodSequenceNumber = " & PeriodSeqBeingEvaluated 
							SQL_G = SQL_G & " ORDER BY " & CatOrderByClauseCustom 
							
							'Response.Write(SQL_G & "<BR>")
							Set	rsGScreen= cnn8.Execute(SQL_G)
				
							IF NOT rsGScreen.EOF Then 
								Do While Not rsGScreen.EOF 
								
									If IsNull(rsGScreen("GroupTot")) Then GroupTot = 0 Else GroupTot = rsGScreen("GroupTot")
									
									If ShowZeros = 1 Then					
										Response.Write("<td width='60' align='right' class='td-border'>" & FormatCurrency(GroupTot,0) & "</td>")
									Else
										' Cat total handles the current period, which must be included now
										If Session("NewPostLogic") <> True Then CatTotal = GetCurrent_PostedTotal_ByCustByCat(CustForDetail,PeriodSeqBeingEvaluated ,rsGScreen("Category")) +GetCurrent_UnPostedTotal_ByCustByCat(CustForDetail,PeriodSeqBeingEvaluated ,rsGScreen("Category"))					
										If Session("NewPostLogic") = True Then CatTotal = PostUnPostCatArray(rsGScreen("Category"),0) + PostUnPostCatArray(rsGScreen("Category"),1)
										If GetTotalSalesForCategByPeriodRange(rsGScreen("Category"),PeriodSeqBeingEvaluated - 11,PeriodSeqBeingEvaluated,CustForDetail) + CatTotal <> 0  Then
											Response.Write("<td width='60' align='right' class='td-border'>" & FormatCurrency(GroupTot,0) & "</td>")
										End If
									End If
								
									rsGScreen.Movenext
								Loop
										
							End If

If Session("DebugSpeed") = True Then Response.Write("<br><br>7.1:" & Now())		
				%>
					</tr>
				
		
					<tr>
						<td width="80" rowspan="2"><h4 class="center-bold">Average</h4></td>
						<td width="60" class="td-border td-three-period"><strong>3 Period&nbsp;&nbsp;(<%=GetPeriodBySeq(PeriodSeqBeingEvaluated-1)%>-<%=GetPeriodBySeq(PeriodSeqBeingEvaluated-3)%>)</strong></td>
						<%
						If VarianceBasis = "3Periods" Then
							StrongTag = "<strong>"
							StrongTagEnd = "</strong>"
						Else
							StrongTag = ""
							StrongTagEnd = ""
						End If
						Response.Write("<td width='60' align='right' class='td-border td-three-period'>" & StrongTag & FormatCurrency(TotalSalesAllCats3PeriodAverage,0) & StrongTagEnd & "</td>")
						
						If ShowGPPercent = 1 Then
							If TotalSalesAllCats3PeriodAverage = 0 Then
								Response.Write("<td align='right' class='td-border'>" & FormatCurrency(0,0) & "</td>")
							Else
								GP = ((TotalSalesAllCats3PeriodAverage - TotalCostAllCats3PeriodAverage)/TotalSalesAllCats3PeriodAverage )*100
								Response.Write("<td align='right' class='td-border'>" & Round(GP,1) & "%</td>")
							End if
						End If
						
						
						SQL_G="SELECT AVG(TotalSales) AS ThreePeriodAverage, Category FROM zCatAnalByPeriod_SingleCustomer_" & Session("UserNo") & "_" & CustForDetail
						SQL_G = SQL_G & " WHERE (ThisPeriodSequenceNumber < " & PeriodSeqBeingEvaluated & ") AND (ThisPeriodSequenceNumber > " & PeriodSeqBeingEvaluated - 4 & ") "
						SQL_G = SQL_G & " GROUP BY Category "
						SQL_G = SQL_G & " ORDER BY " & CatOrderByClauseCustom 
						
						'Response.Write(SQL_G & "<BR>")
						Set	rsGScreen= cnn8.Execute(SQL_G)
			
						IF NOT rsGScreen.EOF Then 
							Do While Not rsGScreen.EOF 
							
								If IsNull(rsGScreen("ThreePeriodAverage")) Then ThreePeriodAverage = 0 Else ThreePeriodAverage = rsGScreen("ThreePeriodAverage")
								
								If ShowZeros = 1 Then					
									Response.Write("<td width='60' align='right' class='td-border td-three-period'>" & StrongTag & FormatCurrency(ThreePeriodAverage,0) & StrongTagEnd & "</td>")
								Else
									' Cat total handles the current period, which must be included now
									If Session("NewPostLogic") <> True Then CatTotal = GetCurrent_PostedTotal_ByCustByCat(CustForDetail,PeriodSeqBeingEvaluated ,rsGScreen("Category")) +GetCurrent_UnPostedTotal_ByCustByCat(CustForDetail,PeriodSeqBeingEvaluated ,rsGScreen("Category"))
									If Session("NewPostLogic") = True Then CatTotal = PostUnPostCatArray(rsGScreen("Category"),0) + PostUnPostCatArray(rsGScreen("Category"),1)
									If GetTotalSalesForCategByPeriodRange(rsGScreen("Category"),PeriodSeqBeingEvaluated - 11,PeriodSeqBeingEvaluated,CustForDetail) + CatTotal <> 0  Then
										Response.Write("<td width='60' align='right' class='td-border td-three-period'>" & StrongTag  & FormatCurrency(ThreePeriodAverage,0) & StrongTagEnd & "</td>")
									End If
								End If
							
								rsGScreen.Movenext
							Loop
									
						End If
If Session("DebugSpeed") = True Then Response.Write("<br><br>7.2:" & Now())	
					%>
					</tr>

					 <tr>
						<td width="60" class="td-border td-twelve-period"><strong>12 Period&nbsp;&nbsp;(<%=GetPeriodBySeq(PeriodSeqBeingEvaluated-1)%>-<%=GetPeriodBySeq(PeriodSeqBeingEvaluated-12)%>)</strong></td>
						<%
						If VarianceBasis <> "3Periods" Then
							StrongTag = "<strong>"
							StrongTagEnd = "</strong>"
						Else
							StrongTag = ""
							StrongTagEnd = ""
						End If

						Response.Write("<td width='60' align='right' class='td-border td-twelve-period'>" & StrongTag & FormatCurrency(TotalSalesAllCats12PeriodAverage ,0) & StrongTagEnd & "</td>")

						If ShowGPPercent = 1 Then
							If TotalSalesAllCats12PeriodAverage = 0 Then
								Response.Write("<td align='right' class='td-border'>" & FormatCurrency(0,0) & "</td>")
							Else
								GP = ((TotalSalesAllCats12PeriodAverage - TotalCostAllCats12PeriodAverage) /TotalSalesAllCats12PeriodAverage )*100
								Response.Write("<td align='right' class='td-border'>" & Round(GP,1) & "%</td>")
							End if
						End If
						
						' 12 prior periods average
						SQL_G="SELECT AVG(TotalSales) AS TwelvePeriodAverage, Category FROM zCatAnalByPeriod_SingleCustomer_" & Session("UserNo") & "_" & CustForDetail
						SQL_G = SQL_G & " WHERE (ThisPeriodSequenceNumber < " & PeriodSeqBeingEvaluated & ") AND (ThisPeriodSequenceNumber > " & PeriodSeqBeingEvaluated - 13 & ") "
						SQL_G = SQL_G & " GROUP BY Category "
						SQL_G = SQL_G & " ORDER BY " & CatOrderByClauseCustom 

						'Response.Write(SQL_G & "<BR>")
						Set	rsGScreen= cnn8.Execute(SQL_G)
			
						IF NOT rsGScreen.EOF Then 
							Do While Not rsGScreen.EOF 
							
								If IsNull(rsGScreen("TwelvePeriodAverage")) Then TwelvePeriodAverage = 0 Else TwelvePeriodAverage = rsGScreen("TwelvePeriodAverage")
								
								If ShowZeros = 1 Then					
									Response.Write("<td width='60' align='right' class='td-border td-twelve-period'>" & StrongTag & FormatCurrency(TwelvePeriodAverage,0) & StrongTagEnd & "</td>")
								Else 
									' Cat total handles the current period, which must be included now
									If Session("NewPostLogic") <> True Then CatTotal = GetCurrent_PostedTotal_ByCustByCat(CustForDetail,PeriodSeqBeingEvaluated ,rsGScreen("Category")) +GetCurrent_UnPostedTotal_ByCustByCat(CustForDetail,PeriodSeqBeingEvaluated ,rsGScreen("Category"))
									If Session("NewPostLogic") = True Then CatTotal = PostUnPostCatArray(rsGScreen("Category"),0) + PostUnPostCatArray(rsGScreen("Category"),1)
									If GetTotalSalesForCategByPeriodRange(rsGScreen("Category"),PeriodSeqBeingEvaluated - 11,PeriodSeqBeingEvaluated,CustForDetail) + CatTotal <> 0  Then
										Response.Write("<td width='60' align='right' class='td-border td-twelve-period'>" & StrongTag & FormatCurrency(TwelvePeriodAverage,0) & StrongTagEnd &  "</td>")
									End If
								End If
							
								rsGScreen.Movenext
							Loop
									
						End If
If Session("DebugSpeed") = True Then Response.Write("<br><br>7.3:" & Now())
If Session("DebugSpeed") = True Then Response.Write("<br><br>CURRENT" & Now())					
%>	
				
					</tr>

					<tr class="clickable" data-toggle="collapse" id="row1" data-target=".row1">
						<td width="80"><h4 class="center-bold">Current <i class="fa fa-plus-square-o" aria-hidden="true"></i></h4></td>
						
						<% CurrentPeriodName = "Period&nbsp;" & GetPeriodBySeq(PeriodSeqBeingEvaluated  + 1)%>
						
						<td width="60" class="td-border"><strong><%= CurrentPeriodName %></strong></td>
						<td width="60" align="right" class="td-border"><% = FormatCurrency(TotalSalesAllCatsCurrentPeriod,0)%></td>
						
						<%
						If ShowGPPercent = 1 Then
							If TotalSalesAllCatsCurrentPeriod = 0 Then
								Response.Write("<td align='right' class='td-border'>" & FormatCurrency(0,0) & "</td>")
							Else
								GP = ((TotalSalesAllCatsCurrentPeriod - TotalClostAllCatsCurrentPeriod ) / TotalSalesAllCatsCurrentPeriod)*100
								Response.Write("<td align='right' class='td-border'>" & Round(GP,1) & "%</td>")
							End if
						End If
						%>
						<!--<td width="60" align="right" class="td-border"><% = FormatCurrency(GetCurrent_PostedTotal_ByCust(CustForDetail,PeriodSeqBeingEvaluated)+GetCurrent_UnPostedTotal_ByCust(CustForDetail,PeriodSeqBeingEvaluated),0)%></td>-->
						<%
						' Current Pposted
						SQL_G="SELECT Category, Count(*) FROM zCatAnalByPeriod_SingleCustomer_" & Session("UserNo") & "_" & CustForDetail
						SQL_G = SQL_G & " WHERE (ThisPeriodSequenceNumber < " & PeriodSeqBeingEvaluated & ") AND (ThisPeriodSequenceNumber > " & PeriodSeqBeingEvaluated - 13 & ") "
						SQL_G = SQL_G & " GROUP BY Category "
						SQL_G = SQL_G & " ORDER BY " & CatOrderByClauseCustom 

						SQL_G="SELECT Category, Count(*) FROM zCatAnalByPeriod_SingleCustomer_" & Session("UserNo") & "_" & CustForDetail
						'SQL_G = SQL_G & " WHERE (ThisPeriodSequenceNumber < " & PeriodSeqBeingEvaluated & ") AND (ThisPeriodSequenceNumber > " & PeriodSeqBeingEvaluated - 13 & ") "
						SQL_G = SQL_G & " WHERE ThisPeriodSequenceNumber = " & PeriodSeqBeingEvaluated + 1 & " "
						SQL_G = SQL_G & " GROUP BY Category "
						SQL_G = SQL_G & " ORDER BY " & CatOrderByClauseCustom 

						'Response.Write(SQL_G & "<BR>")
						Set	rsGScreen= cnn8.Execute(SQL_G)
			
						IF NOT rsGScreen.EOF Then 
							Do While Not rsGScreen.EOF 
							
								If Session("NewPostLogic") <> True Then CatTotal = GetCurrent_PostedTotal_ByCustByCat(CustForDetail,PeriodSeqBeingEvaluated ,rsGScreen("Category")) +GetCurrent_UnPostedTotal_ByCustByCat(CustForDetail,PeriodSeqBeingEvaluated ,rsGScreen("Category"))
								
								If Session("NewPostLogic") = True Then CatTotal = PostUnPostCatArray(rsGScreen("Category"),0) + PostUnPostCatArray(rsGScreen("Category"),1)
								
								If ShowZeros = 1 Then					
									Response.Write("<td width='60' align='right' class='td-border'>" & FormatCurrency(CatTotal,0) & "</td>")
								Else
									' EvalTotal total handles the current period, which must be included now
									If Session("NewPostLogic") <> True Then EvalTotal = GetCurrent_PostedTotal_ByCustByCat(CustForDetail,PeriodSeqBeingEvaluated ,rsGScreen("Category")) +GetCurrent_UnPostedTotal_ByCustByCat(CustForDetail,PeriodSeqBeingEvaluated ,rsGScreen("Category"))										
									If Session("NewPostLogic") = True Then EvalTotal  = PostUnPostCatArray(rsGScreen("Category"),0) + PostUnPostCatArray(rsGScreen("Category"),1)
									If GetTotalSalesForCategByPeriodRange(rsGScreen("Category"),PeriodSeqBeingEvaluated - 11,PeriodSeqBeingEvaluated+1,CustForDetail)  + EvalTotal <> 0  Then
										Response.Write("<td width='60' align='right' class='td-border'>" & FormatCurrency(CatTotal ,0) & "</td>")
									End If
								End If
							
								rsGScreen.Movenext
							Loop
									
						End If
						%>
					</tr>

					<tr class="collapse row1">
						<td width="80">&nbsp;</td>
						<td width="60" class="td-border"><strong>Unposted</strong></td>
						<td width="60" align="right" class="td-border"><% = FormatCurrency(GetCurrent_UnPostedTotal_ByCust(CustForDetail,PeriodSeqBeingEvaluated),0)%></td>
						
						<%
						If ShowGPPercent = 1 Then
							If GetCurrent_UnPostedTotal_ByCust(CustForDetail,PeriodSeqBeingEvaluated) = 0 Then
								Response.Write("<td align='right' class='td-border'>" & FormatCurrency(0,0) & "</td>")
							Else
								GP = ((GetCurrent_UnPostedTotal_ByCust(CustForDetail,PeriodSeqBeingEvaluated) - GetCurrent_UnPostedTotalCost_ByCust()) /GetCurrent_UnPostedTotal_ByCust(CustForDetail,PeriodSeqBeingEvaluated))*100
								Response.Write("<td align='right' class='td-border'>" & Round(GP,1) & "%</td>")
							End if
						End If
						%>

						<%
						' Current Unposted
						SQL_G="SELECT Category, Count(*) FROM zCatAnalByPeriod_SingleCustomer_" & Session("UserNo") & "_" & CustForDetail
						'SQL_G = SQL_G & " WHERE (ThisPeriodSequenceNumber < " & PeriodSeqBeingEvaluated & ") AND (ThisPeriodSequenceNumber > " & PeriodSeqBeingEvaluated - 13 & ") "
						SQL_G = SQL_G & " WHERE ThisPeriodSequenceNumber = " & PeriodSeqBeingEvaluated + 1 & " "
						SQL_G = SQL_G & " GROUP BY Category "
						SQL_G = SQL_G & " ORDER BY " & CatOrderByClauseCustom 

						'Response.Write(SQL_G & "<BR>")
						Set	rsGScreen= cnn8.Execute(SQL_G)
			
						IF NOT rsGScreen.EOF Then 
							Do While Not rsGScreen.EOF 
							
								If Session("NewPostLogic") <> True Then CatTotal = GetCurrent_UnpostedTotal_ByCustByCat(CustForDetail,PeriodSeqBeingEvaluated ,rsGScreen("Category"))
	
								If Session("NewPostLogic") = True Then CatTotal = PostUnPostCatArray(rsGScreen("Category"),0)
								
								If ShowZeros = 1 Then					
									Response.Write("<td width='60' align='right' class='td-border'>" & FormatCurrency(CatTotal,0) & "</td>")
								Else
									' Eval total handles the current period, which must be included now
									If Session("NewPostLogic") <> True Then EvalTotal = GetCurrent_PostedTotal_ByCustByCat(CustForDetail,PeriodSeqBeingEvaluated ,rsGScreen("Category")) +GetCurrent_UnPostedTotal_ByCustByCat(CustForDetail,PeriodSeqBeingEvaluated ,rsGScreen("Category"))
									If Session("NewPostLogic") = True Then EvalTotal =PostUnPostCatArray(rsGScreen("Category"),0) + PostUnPostCatArray(rsGScreen("Category"),1)
									'Response.Write("Cat : " & rsGScreen("Category")& " - " & GetCategoryByID(rsGScreen("Category"))& "<br>")
									'Response.Write("Posted : " & GetCurrent_PostedTotal_ByCustByCat(CustForDetail,PeriodSeqBeingEvaluated ,rsGScreen("Category"))& "<br>")
									'Response.Write("Unposted : " & GetCurrent_UnPostedTotal_ByCustByCat(CustForDetail,PeriodSeqBeingEvaluated ,rsGScreen("Category"))& "<br>")
									'Response.Write("EvalTotal : " & EvalTotal  & "<br>")
									If GetTotalSalesForCategByPeriodRange(rsGScreen("Category"),PeriodSeqBeingEvaluated - 11,PeriodSeqBeingEvaluated,CustForDetail)  + EvalTotal <> 0  Then
									Response.Write("<td width='60' align='right' class='td-border'>" & FormatCurrency(CatTotal ,0) & "</td>")
									End If
								End If
							
								rsGScreen.Movenext
							Loop
							
						End If

						
						%>
					</tr>
		
					<tr class="collapse row1">
						<td width="80">&nbsp;</td>
						<td width="60" class="td-border"><strong>Posted</strong></td>
						<td width="60" align="right" class="td-border"><% = FormatCurrency(GetCurrent_PostedTotal_ByCust(CustForDetail,PeriodSeqBeingEvaluated),0)%></td>
						
						<%
						If ShowGPPercent = 1 Then
							If GetCurrent_PostedTotal_ByCust(CustForDetail,PeriodSeqBeingEvaluated) = 0 Then
								Response.Write("<td align='right' class='td-border'>" & FormatCurrency(0,0) & "</td>")
							Else
								GP = ((GetCurrent_PostedTotal_ByCust(CustForDetail,PeriodSeqBeingEvaluated) - GetCurrent_PostedTotalCost_ByCust()) /GetCurrent_PostedTotal_ByCust(CustForDetail,PeriodSeqBeingEvaluated))*100
								Response.Write("<td align='right' class='td-border'>" & Round(GP,1) & "%</td>")
							End if
						End If
						%>

						<%
						' Current Pposted
						SQL_G="SELECT Category, Count(*) FROM zCatAnalByPeriod_SingleCustomer_" & Session("UserNo") & "_" & CustForDetail
						SQL_G = SQL_G & " WHERE (ThisPeriodSequenceNumber < " & PeriodSeqBeingEvaluated & ") AND (ThisPeriodSequenceNumber > " & PeriodSeqBeingEvaluated - 13 & ") "
						SQL_G = SQL_G & " GROUP BY Category "
						SQL_G = SQL_G & " ORDER BY " & CatOrderByClauseCustom 

						SQL_G="SELECT Category, Count(*) FROM zCatAnalByPeriod_SingleCustomer_" & Session("UserNo") & "_" & CustForDetail
						'SQL_G = SQL_G & " WHERE (ThisPeriodSequenceNumber < " & PeriodSeqBeingEvaluated & ") AND (ThisPeriodSequenceNumber > " & PeriodSeqBeingEvaluated - 13 & ") "
						SQL_G = SQL_G & " WHERE ThisPeriodSequenceNumber = " & PeriodSeqBeingEvaluated + 1 & " "
						SQL_G = SQL_G & " GROUP BY Category "
						SQL_G = SQL_G & " ORDER BY " & CatOrderByClauseCustom 

						'Response.Write(SQL_G & "<BR>")
						Set	rsGScreen= cnn8.Execute(SQL_G)
			
						IF NOT rsGScreen.EOF Then 
							Do While Not rsGScreen.EOF 
							
								If Session("NewPostLogic") <> True Then CatTotal = GetCurrent_PostedTotal_ByCustByCat(CustForDetail,PeriodSeqBeingEvaluated ,rsGScreen("Category"))
								
								If Session("NewPostLogic") = True Then CatTotal = PostUnPostCatArray(rsGScreen("Category"),1)

								If ShowZeros = 1 Then					
									Response.Write("<td width='60' align='right' class='td-border'>" & FormatCurrency(CatTotal,0) & "</td>")
								Else
									' EvalTotal total handles the current period, which must be included now
									If Session("NewPostLogic") <> True Then EvalTotal = GetCurrent_PostedTotal_ByCustByCat(CustForDetail,PeriodSeqBeingEvaluated ,rsGScreen("Category")) +GetCurrent_UnPostedTotal_ByCustByCat(CustForDetail,PeriodSeqBeingEvaluated ,rsGScreen("Category"))
									If Session("NewPostLogic") = True Then EvalTotal = PostUnPostCatArray(rsGScreen("Category"),0) + PostUnPostCatArray(rsGScreen("Category"),1)
									
									If GetTotalSalesForCategByPeriodRange(rsGScreen("Category"),PeriodSeqBeingEvaluated - 11,PeriodSeqBeingEvaluated+1,CustForDetail)  + EvalTotal <> 0  Then
									 	Response.Write("<td width='60' align='right' class='td-border'>" & FormatCurrency(CatTotal ,0) & "</td>")
									End If
								End If
							
								rsGScreen.Movenext
							Loop
									
						End If
If Session("DebugSpeed") = True Then Response.Write("<br><br>EOF CURRENT" & Now())						%>
					</tr>

					
					<tr>

					<td width="80" class="no-padding"><h4 class="center-bold">Variance</h4></td>
						<td width="60" class="td-border no-padding"><strong>Sales</strong></td>
					<%
					' Variance Row
					If VarianceBasis = "3Periods" Then
						PriorThreePeriodAverageAll = GetPriorThreePeriodAverageAllCatsByPerSeq(PeriodSeqBeingEvaluated,CustForDetail)
						If TotalSalesByPeriodSeq(PeriodSeqBeingEvaluated,CustForDetail) - PriorThreePeriodAverageAll  <= 0 Then
							Response.Write("<td class='td-border no-padding negative-large-with-bg' width='60' align='right'>" & FormatCurrency(TotalSalesByPeriodSeq(PeriodSeqBeingEvaluated,CustForDetail) - PriorThreePeriodAverageAll ,0) & "</td>")
						Else
							Response.Write("<td class='td-border no-padding positive-large-with-bg' width='60' align='right'>" & FormatCurrency(TotalSalesByPeriodSeq(PeriodSeqBeingEvaluated,CustForDetail) - PriorThreePeriodAverageAll ,0) & "</td>")
						End If
					End If	
					If VarianceBasis = "12Periods" Then
						PriorTwelvePeriodAverageAll = GetPriorTwelvePeriodAverageAllCatsByPerSeq(PeriodSeqBeingEvaluated,CustForDetail)
						If TotalSalesByPeriodSeq(PeriodSeqBeingEvaluated,CustForDetail) - PriorTwelvePeriodAverageAll <= 0 Then
							Response.Write("<td class='td-border no-padding negative-large-with-bg' width='60' align='right'>" & FormatCurrency(TotalSalesByPeriodSeq(PeriodSeqBeingEvaluated,CustForDetail) - PriorTwelvePeriodAverageAll ,0) & "</td>")
						Else
							Response.Write("<td class='td-border no-padding positive-large-with-bg' width='60' align='right'>" & FormatCurrency(TotalSalesByPeriodSeq(PeriodSeqBeingEvaluated,CustForDetail) - PriorTwelvePeriodAverageAll ,0) & "</td>")
						End If
					End If	
					
					%>						
					<% If ShowGPPercent = 1 Then %>
							<td class="td-border gp-percent no-padding" align="right">&nbsp;</td>
					<% End If %>
					<%

					SQL_G="SELECT Category, CategoryNameGetTerm FROM zCatAnalByPeriod_SingleCustomer_" & Session("UserNo") & "_" & CustForDetail & " WHERE ThisPeriodSequenceNumber = " & PeriodSeqBeingEvaluated 
					SQL_G = SQL_G & " ORDER BY " & CatOrderByClauseCustom 
	
					Set	rsGScreen = cnn8.Execute(SQL_G)
					
					If not rsGScreen.eof Then
						
						ContribCount = 0
					
						Do

							If ShowZeros = 1 Then
								If ContribCount = 0 Then 
									'Indicator to show dollars difference
									TotalSlsThisPeriod = GetTotalSalesForCategByPeriodRange(rsGScreen("Category"),PeriodSeqBeingEvaluated,PeriodSeqBeingEvaluated,CustForDetail)
									If VarianceBasis = "3Periods" Then TotalSalesDiff = TotalSlsThisPeriod - GetPriorThreePeriodAverageByCatByPerSeq(rsGScreen("Category"),PeriodSeqBeingEvaluated,CustForDetail)
									If VarianceBasis = "12Periods" Then TotalSalesDiff =TotalSlsThisPeriod - GetPriorTwelvePeriodAverageByCatByPerSeq(rsGScreen("Category"),PeriodSeqBeingEvaluated,CustForDetail)
									If TotalSalesDiff > 0 Then 
										Response.Write("<td class='td-border no-padding positive-large' width='60' align='right'>" & FormatCurrency(TotalSalesDiff ,0) & "</td>")
									ElseIf TotalSalesDiff = 0 Then 
										Response.Write("<td class='td-border no-padding zero-large' width='60' align='right'>0</td>")
										'Response.Write("<td class='td-border no-padding zero-large' width='60' align='right'>" & FormatCurrency(TotalSalesDiff ,0) & "</td>")
										'Response.Write("<td class='td-border no-padding zero-large' width='60' align='right'>&nbsp;</td>")
									Else
										Response.Write("<td class='td-border no-padding negative-large' width='60' align='right'>" & FormatCurrency(TotalSalesDiff ,0) & "</td>")
									End If
								End If
							Else
								' Cat total handles the current period, which must be included now
								If Session("NewPostLogic") <> True Then CatTotal = GetCurrent_PostedTotal_ByCustByCat(CustForDetail,PeriodSeqBeingEvaluated ,rsGScreen("Category")) +GetCurrent_UnPostedTotal_ByCustByCat(CustForDetail,PeriodSeqBeingEvaluated ,rsGScreen("Category"))
								If Session("NewPostLogic") = True Then CatTotal = PostUnPostCatArray(rsGScreen("Category"),0) + PostUnPostCatArray(rsGScreen("Category"),1)
								If GetTotalSalesForCategByPeriodRange(rsGScreen("Category"),PeriodSeqBeingEvaluated - 11,PeriodSeqBeingEvaluated,CustForDetail) + CatTotal <> 0 Then
									If ContribCount = 0 Then 
										'Indicator to show dollars difference
										TotalSlsThisPeriod = GetTotalSalesForCategByPeriodRange(rsGScreen("Category"),PeriodSeqBeingEvaluated,PeriodSeqBeingEvaluated,CustForDetail)
										If VarianceBasis = "3Periods" Then TotalSalesDiff = TotalSlsThisPeriod - GetPriorThreePeriodAverageByCatByPerSeq(rsGScreen("Category"),PeriodSeqBeingEvaluated,CustForDetail)
										If VarianceBasis = "12Periods" Then TotalSalesDiff =TotalSlsThisPeriod - GetPriorTwelvePeriodAverageByCatByPerSeq(rsGScreen("Category"),PeriodSeqBeingEvaluated,CustForDetail)
										If TotalSalesDiff > 0 Then 
											Response.Write("<td class='td-border no-padding positive-large' width='60' align='right'>" & FormatCurrency(TotalSalesDiff ,0) & "</td>")
										ElseIF TotalSalesDiff = 0 Then 
											'Response.Write("<td class='td-border no-padding' width='60' align='right'>&nbsp;</td>")
											Response.Write("<td class='td-border no-padding zero-large' width='60' align='right'><strong>0</strong></font></td>")
										Else
											Response.Write("<td class='td-border no-padding negative-large' width='60' align='right'>" & FormatCurrency(TotalSalesDiff ,0) & "</td>")
										End If
									End If
								End If
	
							End If
						
							'ContribCount = ContribCount + 1
	
							rsGScreen.movenext
						loop until rsGScreen.eof
					End If
					
If Session("DebugSpeed") = True Then Response.Write("<br><br>7.7:" & Now())	
				'<td class='td-border' width="60">&nbsp;</td>
					
	
					%>
					</tr>

					 <tr>
						<td width="80" class="no-padding">&nbsp;</td>
						<td width="60" class="td-border no-padding"><strong>Cases</strong></td>
											<%
					' Variance Row
					VarianceResult = GetCaseVarianceByCust_ALLCats(PeriodSeqBeingEvaluated,VarianceBasis,CustForDetail)
					If VarianceResult < 0 Then
						Response.Write("<td class='td-border no-padding negative-large-with-bg' width='60' align='right'>(" & Round(VarianceResult ,0) & ")</td>")
					ElseIf VarianceResult = 0 Then	
						Response.Write("<td class='td-border no-padding negative-large-with-bg' width='60' align='right'>0</td>")
					Else
						Response.Write("<td class='td-border no-padding positive-large-with-bg' width='60' align='right'>" & Round(VarianceResult ,0) & "</td>")
					End If
					
					%>						
					<% If ShowGPPercent = 1 Then %>
							<td class="td-border gp-percent no-padding" align="right">&nbsp;</td>
					<% End If %>
					<%

					SQL_G="SELECT Category, CategoryNameGetTerm FROM zCatAnalByPeriod_SingleCustomer_" & Session("UserNo") & "_" & CustForDetail & " WHERE ThisPeriodSequenceNumber = " & PeriodSeqBeingEvaluated 
					SQL_G = SQL_G & " ORDER BY " & CatOrderByClauseCustom 
	
					Set	rsGScreen = cnn8.Execute(SQL_G)
					
					If not rsGScreen.eof Then
						
						ContribCount = 0
					
						Do

							VarianceResult = GetCaseVarianceByCustByCat(PeriodSeqBeingEvaluated,rsGScreen("Category"),VarianceBasis ,CustForDetail)
							
							If ShowZeros = 1 Then
								If VarianceResult > 0 Then 
									Response.Write("<td class='td-border no-padding positive-large' width='60' align='right'>" & Round(VarianceResult ,0) & "</td>")
								ElseIf VarianceResult = 0 Then 
									Response.Write("<td class='td-border no-padding zero-large' width='60' align='right'>0</td>")
								Else
									Response.Write("<td class='td-border no-padding negative-large' width='60' align='right'>" & Round(VarianceResult ,0) & "</td>")
								End If
							Else
								' Cat total handles the current period, which must be included now
								If Session("NewPostLogic") <> True Then CatTotal = GetCurrent_PostedTotal_ByCustByCat(CustForDetail,PeriodSeqBeingEvaluated ,rsGScreen("Category")) +GetCurrent_UnPostedTotal_ByCustByCat(CustForDetail,PeriodSeqBeingEvaluated ,rsGScreen("Category"))							
								If Session("NewPostLogic") = True Then CatTotal = PostUnPostCatArray(rsGScreen("Category"),0) + PostUnPostCatArray(rsGScreen("Category"),1)
								If GetTotalSalesForCategByPeriodRange(rsGScreen("Category"),PeriodSeqBeingEvaluated - 11,PeriodSeqBeingEvaluated,CustForDetail) + CatTotal <> 0 Then
									If VarianceResult > 0 Then 
										Response.Write("<td class='td-border no-padding positive-large' width='60' align='right'>" & Round(VarianceResult ,0) & "</td>")
									ElseIF VarianceResult = 0 Then 
										Response.Write("<td class='td-border no-padding zero-large' width='60' align='right'>0</td>")
									Else
										Response.Write("<td class='td-border no-padding negative-large' width='60' align='right'>" & Round(VarianceResult ,0) & "</td>")
									End If
								End If
							End If
						
	
							rsGScreen.movenext
						loop until rsGScreen.eof
					End If
If Session("DebugSpeed") = True Then Response.Write("<br><br>8:" & Now())

	
					%>
				</tr>


				

		</table>
	</div>
 </div>
 <!-- eof row -->

  

<!-- row -->
 <div class="col-lg-12">
	<div class='table-striped table-condensed table-hover table-responsive table-border'>
		<table class='table table-striped table-condensed table-hover large-table' >
		<!-- <table class='table table-striped table-condensed table-hover large-table sortable' > -->


		<%

		For x = PeriodSeqBeingEvaluated to (PeriodSeqBeingEvaluated - 12) Step -1
			
			If x = PeriodSeqBeingEvaluated  Then
				Response.Write("<tr>")
				Response.Write("<th scope='col' class='sorttable_nosort' width='80'>Period End</th>")
				Response.Write("<th scope='col' class='break sorttable ' valign='top' width='60'>Period #</th>")
				
				'Allow for a note here as a way to put in a note for the customer in general
				'Use -1 as the category number
				If CustHasCategoryAnalNotes(CustForDetail,-1) = True Then
					If NoteNewCatAnalForUser(CustForDetail,-1) = True Then
						'Pulsing icon
						Response.Write("<th scope='col' class='break sorttable ' valign='top' width='60'>")
						Response.Write("<span class='period-title'>Period Total</span></a>")									
						Response.Write("</th>")
					Else
						'Regular icon
						Response.Write("<th scope='col' class='break sorttable ' valign='top' width='60'>")
						Response.Write("<span class='period-title'>Period Total</span></a>")
						Response.Write("</th>")
					End If
				Else
					'Pencil icon
					Response.Write("<th scope='col' class='break sorttable ' valign='top' width='60'>")
					Response.Write("<span class='period-title'>Period Total</span></a>")										
					Response.Write("</th>")
				End If
				

				
				
				'********************************************************
				' GP% Column Show/Hide
				'********************************************************
		
				If ShowGPPercent = 1 Then
					Response.Write("<th scope='col' class='break sorttable'valign='top' width='30'><span class='period-title gp-percent'>GP%</span></th>")
				End If

				
				SQL_G="SELECT * FROM zCatAnalByPeriod_SingleCustomer_" & Session("UserNo") & "_" & CustForDetail & " WHERE ThisPeriodSequenceNumber = " & x 
				SQL_G = SQL_G & " ORDER BY " & CatOrderByClauseCustom 

				Set	rsGScreen = cnn8.Execute(SQL_G)
				
				If not rsGScreen.eof Then
					
					'Category names get written below
					
					Do
						If ShowZeros = 1 Then
											
							Response.Write("<th scope='col' class='break sorttable blue-bg' valign='top' width='60'>")
							
							If Session("SkipGetTerm") <> True Then
								Response.Write("<span class='category-title'><a data-toggle='modal' data-target='#modalCategoryVPC' data-category-id='" & rsGScreen("Category") & "' data-cust-id='" & CustForDetail & "' data-period-seq='" & PeriodSeqBeingEvaluated & "' data-variance-basis='" & VarianceBasis & "'>" & GetTerm(GetCategoryByID(rsGScreen("Category"))) & "</a></span>")
							Else
								Response.Write("<span class='category-title'><a data-toggle='modal' data-target='#modalCategoryVPC' data-category-id='" & rsGScreen("Category") & "' data-cust-id='" & CustForDetail & "' data-period-seq='" & PeriodSeqBeingEvaluated & "' data-variance-basis='" & VarianceBasis & "'>" & rsGScreen("CategoryNameGetTerm") & "</a></span>")									
							End If
							
							Response.Write("</th>")

						Else
							' Cat total handles the current period, which must be included now
							If Session("NewPostLogic") <> True Then CatTotal = GetCurrent_PostedTotal_ByCustByCat(CustForDetail,PeriodSeqBeingEvaluated ,rsGScreen("Category")) +GetCurrent_UnPostedTotal_ByCustByCat(CustForDetail,PeriodSeqBeingEvaluated ,rsGScreen("Category"))						
							If Session("NewPostLogic") = True Then CatTotal = PostUnPostCatArray(rsGScreen("Category"),0) + PostUnPostCatArray(rsGScreen("Category"),1)
							If GetTotalSalesForCategByPeriodRange(rsGScreen("Category"),PeriodSeqBeingEvaluated - 11,PeriodSeqBeingEvaluated,CustForDetail) + CatTotal <> 0 Then

								Response.Write("<th scope='col' class='break sorttable blue-bg' valign='top' width='60'>")
								
								If Session("SkipGetTerm") <> True Then
									Response.Write("<span class='category-title'><a data-toggle='modal' data-target='#modalCategoryVPC' data-category-id='" & rsGScreen("Category") & "' data-cust-id='" & CustForDetail & "' data-period-seq='" & PeriodSeqBeingEvaluated & "' data-variance-basis='" & VarianceBasis & "'>" & GetTerm(GetCategoryByID(rsGScreen("Category"))) & "</a></span>")
								Else
									Response.Write("<span class='category-title'><a data-toggle='modal' data-target='#modalCategoryVPC' data-category-id='" & rsGScreen("Category") & "' data-cust-id='" & CustForDetail & "' data-period-seq='" & PeriodSeqBeingEvaluated & "' data-variance-basis='" & VarianceBasis & "'>" & rsGScreen("CategoryNameGetTerm") & "</a></span>")									
								End If
								
								Response.Write("</th>")
							End If
						End If
						
						rsGScreen.movenext
					loop until rsGScreen.eof
				End If
				Response.Write("</tr>")

			End If

			Response.Write("<tr>")
			
			SDate = FormatDateTime(GetPeriodBeginDateBySeq(x),2)
			EDate = FormatDateTime(GetPeriodEndDateBySeq(x),2)
			
			SDate = Left(SDate ,Len(SDate )-4) & Right(SDate ,2)
			EDate = Left(EDate ,Len(EDate )-4) & Right(EDate ,2)
			
			Response.Write("<td align='center'>" & SDate  & " - " & EDate & "</td>")
			Response.Write("<td class='td-border' style='text-align:center;'>" & Left(GetPeriodAndYearBySeq(x),Instr(GetPeriodAndYearBySeq(x),"-")-2) & "</td>")
		
			'Now get the totals for the entire period
			TotalSalesPeriodBeingEvaluated = TotalSalesByPeriodSeq(x,CustForDetail)
			If TotalSalesPeriodBeingEvaluated = "" Then
				Response.Write("<td align='right' class='td-border'>" & FormatCurrency(0,0) & "</td>")
			Else
				Response.Write("<td align='right' class='td-border'>" & FormatCurrency(TotalSalesPeriodBeingEvaluated,0) & "</td>")
			End if
			
			If ShowGPPercent = 1 Then
				If TotalSalesPeriodBeingEvaluated = "" or TotalSalesPeriodBeingEvaluated = 0 Then
					Response.Write("<td align='right' class='td-border'>" & FormatCurrency(0,0) & "</td>")
				Else
					GP = ((TotalSalesPeriodBeingEvaluated - TotalCostByPeriodSeq(x,CustForDetail))/TotalSalesPeriodBeingEvaluated)*100
					Response.Write("<td align='right' class='td-border'>" & Round(GP,1) & "%</td>")
				End if
			End If
			
			
			
			
			%>						
			<%
If Session("DebugSpeed") = True Then Response.Write("<br><br>8.5:" & Now())		
	
			SQL_G="SELECT TotalSales AS GroupTot, Category FROM zCatAnalByPeriod_SingleCustomer_" & Session("UserNo") & "_" & CustForDetail & " WHERE ThisPeriodSequenceNumber = " & x 
			SQL_G = SQL_G & " ORDER BY " & CatOrderByClauseCustom 
			
			'Response.Write(SQL_G & "<BR>")
			Set	rsGScreen= cnn8.Execute(SQL_G)

			IF NOT rsGScreen.EOF Then 
				Do While Not rsGScreen.EOF 
				
					If IsNull(rsGScreen("GroupTot")) Then GroupTot = 0 Else GroupTot = rsGScreen("GroupTot")
					
					If ShowZeros = 1 Then					
						Response.Write("<td align='right' class='td-border'>" & FormatCurrency(GroupTot,0))
				
						Response.Write("</td>")
					Else
						' Cat total handles the current period, which must be included now
						If Session("NewPostLogic") <> True Then CatTotal = GetCurrent_PostedTotal_ByCustByCat(CustForDetail,PeriodSeqBeingEvaluated ,rsGScreen("Category")) +GetCurrent_UnPostedTotal_ByCustByCat(CustForDetail,PeriodSeqBeingEvaluated ,rsGScreen("Category"))					
						If Session("NewPostLogic") = True Then CatTotal = PostUnPostCatArray(rsGScreen("Category"),0) + PostUnPostCatArray(rsGScreen("Category"),1)
						If GetTotalSalesForCategByPeriodRange(rsGScreen("Category"),PeriodSeqBeingEvaluated - 11,PeriodSeqBeingEvaluated,CustForDetail) + CatTotal <> 0  Then
							Response.Write("<td align='right' class='td-border'>" & FormatCurrency(GroupTot,0))
							Response.Write("</td>")
						End If
					End If
				
					rsGScreen.Movenext
				Loop
						
			End If
				
			Response.Write("</tr>")

		Next
		 
		'Now do the grand totals at the bottom of the page, just for the previous 12 periods, don't have to break it out by cat
		SQL_G="SELECT Sum(TotalSales) AS GrandPeriodTot FROM zCatAnalByPeriod_SingleCustomer_" & Session("UserNo") & "_" & CustForDetail & " "
		SQL_G = SQL_G & " WHERE ThisPeriodSequenceNumber < " & PeriodSeqBeingEvaluated & " AND ThisPeriodSequenceNumber > " & PeriodSeqBeingEvaluated -13
					
		'Response.Write(SQL_G & "<BR>")
		Set	rsGScreen= cnn8.Execute(SQL_G)

		IF NOT rsGScreen.EOF Then 
			Response.Write("<tr>")
			'Response.Write("<td>" & "&nbsp;" & "</td>")
			Response.Write("<td colspan ='2' class='td-border' align='right'><h4>" & "Last 12 periods (")
			Response.Write(GetPeriodBySeq(PeriodSeqBeingEvaluated-1) & "-"  & GetPeriodBySeq(PeriodSeqBeingEvaluated -12) & ")</h4></td>")
			Response.Write("<td align='right' class='td-border'><h4>" & FormatCurrency(rsGScreen("GrandPeriodTot"),0) & "</h4></td>")
			Response.Write("</tr>")
		End If

If Session("DebugSpeed") = True Then Response.Write("<br><br>9:" & Now())	
%>

</table>
 
</div>
 <!-- eof row -->

<!-- row !-->
<div class="row">

<div class="col-lg-12"><hr></div>
</div>
<!-- eof row !-->
 
<%End If ' This is the endif for no custoer selected

Sub PrepareTmpTable (passedCustForDetail , passedPeriodSeqBeingEvaluated, passedPeriodBeingEvaluated)

	tblName = "zCatAnalByPeriod_SingleCustomer_" & Session("UserNo") & "_" & passedCustForDetail
	If TimeToRebuild(tblName) <> True Then Exit Sub
	
	Set cnnCatTmp = Server.CreateObject("ADODB.Connection")
	cnnCatTmp.open Session("ClientCnnString")
'Response.Write("<br><br>1-" & Now() & "<br>")
	
	On Error Resume Next
	Set rsCatTmp = Server.CreateObject("ADODB.Recordset")
	Set rsCatTmpForUpdates  = Server.CreateObject("ADODB.Recordset")
	Set rsLeakageLookup  = Server.CreateObject("ADODB.Recordset")
	rsCatTmp.CursorLocation = 3
	rsCatTmpForUpdates.CursorLocation = 3
	rsLeakageLookup.CursorLocation = 3
	Set rsCatTmp = cnnCatTmp.Execute("DROP TABLE zCatAnalByPeriod_SingleCustomer_" & Session("UserNo") & "_" & passedCustForDetail )

	On Error Goto 0
	SQLCatTmp = "CREATE TABLE zCatAnalByPeriod_SingleCustomer_" & Session("UserNo") & "_" & passedCustForDetail & " ( "
	SQLCatTmp = SQLCatTmp & "[RecordCreationDateTime] [datetime] NULL CONSTRAINT [DF_zCatAnalByPeriod_SingleCustomer_" & Session("UserNo") & "_" & passedCustForDetail & "]  DEFAULT (getdate()), "
	SQLCatTmp = SQLCatTmp & "[CustNum] [int] NULL, "
	SQLCatTmp = SQLCatTmp & "[Period] [int] NULL, "
	SQLCatTmp = SQLCatTmp & "[Category] [int] NULL, "
	SQLCatTmp = SQLCatTmp & "[PeriodYear] [int] NULL, "
	SQLCatTmp = SQLCatTmp & "[TotalSales] [money] NULL, "
	SQLCatTmp = SQLCatTmp & "[TotalCost] [money] NULL, "
	SQLCatTmp = SQLCatTmp & "[ThisPeriodSequenceNumber] [int] NULL, "
	SQLCatTmp = SQLCatTmp & "[ThisPeriodLastYearSales] [money] NULL, "
	SQLCatTmp = SQLCatTmp & "[3PriorPeriodsAeverage] [money] NULL, "
	SQLCatTmp = SQLCatTmp & "[3PriorPeriodsTotalSales] [money] NULL, "
	SQLCatTmp = SQLCatTmp & "[DiifThisPeriodVSLastYearDollars] [money] NULL, "
	SQLCatTmp = SQLCatTmp & "[DiifThisPeriodVSLastYearPercent] [float] NULL, "
	SQLCatTmp = SQLCatTmp & "[DiifThisPeriodVSLast3Dollars] [money] NULL, "
	SQLCatTmp = SQLCatTmp & "[DiifThisPeriodVSLast3Percent] [float] NULL, "
	SQLCatTmp = SQLCatTmp & "[DiifThisPeriodVSLast12Dollars] [money] NULL, "
	SQLCatTmp = SQLCatTmp & "[DiifThisPeriodVSLast12Percent] [float] NULL, "
	SQLCatTmp = SQLCatTmp & "[PriorPeriod1Sales] [money] NULL, "
	SQLCatTmp = SQLCatTmp & "[PriorPeriod2Sales] [money] NULL, "
	SQLCatTmp = SQLCatTmp & "[PriorPeriod3Sales] [money] NULL, "
	SQLCatTmp = SQLCatTmp & "[PriorPeriod4Sales] [money] NULL, "
	SQLCatTmp = SQLCatTmp & "[PriorPeriod5Sales] [money] NULL, "
	SQLCatTmp = SQLCatTmp & "[PriorPeriod6Sales] [money] NULL, "
	SQLCatTmp = SQLCatTmp & "[PriorPeriod7Sales] [money] NULL, "
	SQLCatTmp = SQLCatTmp & "[PriorPeriod8Sales] [money] NULL, "
	SQLCatTmp = SQLCatTmp & "[PriorPeriod9Sales] [money] NULL, "
	SQLCatTmp = SQLCatTmp & "[PriorPeriod10Sales] [money] NULL, "
	SQLCatTmp = SQLCatTmp & "[PriorPeriod11Sales] [money] NULL, "
	SQLCatTmp = SQLCatTmp & "[PriorPeriod12Sales] [money] NULL, "
	SQLCatTmp = SQLCatTmp & "[3ContributionDollars] [money] NULL, "
	SQLCatTmp = SQLCatTmp & "[3ContributionPercent] [decimal] (18,2) NULL, "
	SQLCatTmp = SQLCatTmp & "[12ContributionDollars] [money] NULL, "
	SQLCatTmp = SQLCatTmp & "[12ContributionPercent] [decimal] (18,2) NULL, "
	SQLCatTmp = SQLCatTmp & "[12PriorPeriodAverage] [money] NULL, "
	SQLCatTmp = SQLCatTmp & "[GroupName] [varchar](255) NULL, "
	SQLCatTmp = SQLCatTmp & "[CategoryName] [varchar](255) NULL, "
	SQLCatTmp = SQLCatTmp & "[CategoryNameGetTerm] [varchar](255) NULL "
	SQLCatTmp = SQLCatTmp & ") ON [PRIMARY]"
	
	Set rsCatTmp = cnnCatTmp.Execute(SQLCatTmp)

	'Now create records for each period for each category
	For x = passedPeriodSeqBeingEvaluated - 12 to passedPeriodSeqBeingEvaluated + 1 ' To include current period
	
		For z = 0 to 21 ' One for each cat
	
			SQLCatTmp = "INSERT INTO zCatAnalByPeriod_SingleCustomer_" & Session("UserNo")  & "_" & passedCustForDetail 
			SQLCatTmp = SQLCatTmp & " (CustNum ,"
			SQLCatTmp = SQLCatTmp & "Period , "
			SQLCatTmp = SQLCatTmp & "PeriodYear , "
			SQLCatTmp = SQLCatTmp & "Category , "
			SQLCatTmp = SQLCatTmp & "ThisPeriodSequenceNumber , "
			SQLCatTmp = SQLCatTmp & "TotalSales , "
			SQLCatTmp = SQLCatTmp & "TotalCost , "
			SQLCatTmp = SQLCatTmp & "ThisPeriodLastYearSales , "
			SQLCatTmp = SQLCatTmp & "[3PriorPeriodsAeverage] , "
			SQLCatTmp = SQLCatTmp & "[3PriorPeriodsTotalSales] ,"
			SQLCatTmp = SQLCatTmp & "DiifThisPeriodVSLastYearDollars , "
			SQLCatTmp = SQLCatTmp & "DiifThisPeriodVSLastYearPercent , "
			SQLCatTmp = SQLCatTmp & "DiifThisPeriodVSLast3Dollars , "
			SQLCatTmp = SQLCatTmp & "DiifThisPeriodVSLast3Percent , "
			SQLCatTmp = SQLCatTmp & "DiifThisPeriodVSLast12Dollars , "
			SQLCatTmp = SQLCatTmp & "DiifThisPeriodVSLast12Percent , "
			SQLCatTmp = SQLCatTmp & "PriorPeriod1Sales , "
			SQLCatTmp = SQLCatTmp & "PriorPeriod2Sales , "
			SQLCatTmp = SQLCatTmp & "PriorPeriod3Sales , "
			SQLCatTmp = SQLCatTmp & "PriorPeriod4Sales , "
			SQLCatTmp = SQLCatTmp & "PriorPeriod5Sales , "
			SQLCatTmp = SQLCatTmp & "PriorPeriod6Sales , "
			SQLCatTmp = SQLCatTmp & "PriorPeriod7Sales , "
			SQLCatTmp = SQLCatTmp & "PriorPeriod8Sales , "
			SQLCatTmp = SQLCatTmp & "PriorPeriod9Sales , "
			SQLCatTmp = SQLCatTmp & "PriorPeriod10Sales , "
			SQLCatTmp = SQLCatTmp & "PriorPeriod11Sales , "
			SQLCatTmp = SQLCatTmp & "PriorPeriod12Sales, "	
			SQLCatTmp = SQLCatTmp & "[3ContributionDollars], "
			SQLCatTmp = SQLCatTmp & "[3ContributionPercent], "
			SQLCatTmp = SQLCatTmp & "[12ContributionDollars],"
			SQLCatTmp = SQLCatTmp & "[12ContributionPercent],"
			SQLCatTmp = SQLCatTmp & "[12PriorPeriodAverage]"
			SQLCatTmp = SQLCatTmp & " ) "
			SQLCatTmp = SQLCatTmp & " VALUES ("
			SQLCatTmp = SQLCatTmp & passedCustForDetail & "," 
			SQLCatTmp = SQLCatTmp & Trim(Left(passedPeriodBeingEvaluated,Instr(passedPeriodBeingEvaluated,"-") - 1)) & ","
			SQLCatTmp = SQLCatTmp & Trim(Right(PeriodBeingEvaluated,Len(PeriodBeingEvaluated) - Instr(PeriodBeingEvaluated,"-"))) & ","
			SQLCatTmp = SQLCatTmp & z & ", "
			SQLCatTmp = SQLCatTmp & x & ", "
			SQLCatTmp = SQLCatTmp & "0, "
			SQLCatTmp = SQLCatTmp & "0, "
			SQLCatTmp = SQLCatTmp & "0, "	
			SQLCatTmp = SQLCatTmp & "0, "					
			SQLCatTmp = SQLCatTmp & "0, "			
			SQLCatTmp = SQLCatTmp & "0, "			
			SQLCatTmp = SQLCatTmp & "0, "						
			SQLCatTmp = SQLCatTmp & "0, "									
			SQLCatTmp = SQLCatTmp & "0, "									
			SQLCatTmp = SQLCatTmp & "0, "									
			SQLCatTmp = SQLCatTmp & "0, "												
			SQLCatTmp = SQLCatTmp & "0, "						
			SQLCatTmp = SQLCatTmp & "0, "									
			SQLCatTmp = SQLCatTmp & "0, "									
			SQLCatTmp = SQLCatTmp & "0, "									
			SQLCatTmp = SQLCatTmp & "0, "						
			SQLCatTmp = SQLCatTmp & "0, "			
			SQLCatTmp = SQLCatTmp & "0, "
			SQLCatTmp = SQLCatTmp & "0, "
			SQLCatTmp = SQLCatTmp & "0, "
			SQLCatTmp = SQLCatTmp & "0, "						
			SQLCatTmp = SQLCatTmp & "0, "
			SQLCatTmp = SQLCatTmp & "0, "										
			SQLCatTmp = SQLCatTmp & "0, "													
			SQLCatTmp = SQLCatTmp & "0, "													
			SQLCatTmp = SQLCatTmp & "0, "													
			SQLCatTmp = SQLCatTmp & "0, "													
			SQLCatTmp = SQLCatTmp & "0 " 
			SQLCatTmp = SQLCatTmp & ")"
			rsCatTmp.CursorLocation = 3
			Set rsCatTmp = cnnCatTmp.Execute(SQLCatTmp)

		Next 

	Next

	'OK, now update it with the values from the report data file
	SQLCatTmp = "UPDATE zCatAnalByPeriod_SingleCustomer_" & Session("UserNo")  & "_" & passedCustForDetail 
	SQLCatTmp = SQLCatTmp & " SET zCatAnalByPeriod_SingleCustomer_" & Session("UserNo") & "_" & passedCustForDetail & ".ThisPeriodLastYearSales = CustCatPeriodSales.ThisPeriodLastYearSales "
	SQLCatTmp = SQLCatTmp & " , zCatAnalByPeriod_SingleCustomer_" & Session("UserNo") & "_" & passedCustForDetail & ".[TotalSales] = CustCatPeriodSales.[TotalSales]"
	SQLCatTmp = SQLCatTmp & " , zCatAnalByPeriod_SingleCustomer_" & Session("UserNo") & "_" & passedCustForDetail & ".[TotalCost] = CustCatPeriodSales.[TotalCost]"
	SQLCatTmp = SQLCatTmp & " , zCatAnalByPeriod_SingleCustomer_" & Session("UserNo") & "_" & passedCustForDetail & ".[3PriorPeriodsAeverage] = CustCatPeriodSales.[3PriorPeriodsAeverage]"
	SQLCatTmp = SQLCatTmp & " , zCatAnalByPeriod_SingleCustomer_" & Session("UserNo") & "_" & passedCustForDetail & ".[3PriorPeriodsTotalSales] = CustCatPeriodSales.[3PriorPeriodsTotalSales]"
	SQLCatTmp = SQLCatTmp & " , zCatAnalByPeriod_SingleCustomer_" & Session("UserNo") & "_" & passedCustForDetail & ".[DiifThisPeriodVSLastYearDollars] = CustCatPeriodSales.[DiifThisPeriodVSLastYearDollars]"
	SQLCatTmp = SQLCatTmp & " , zCatAnalByPeriod_SingleCustomer_" & Session("UserNo") & "_" & passedCustForDetail & ".[DiifThisPeriodVSLastYearPercent] = CustCatPeriodSales.[DiifThisPeriodVSLastYearPercent]"
	SQLCatTmp = SQLCatTmp & " , zCatAnalByPeriod_SingleCustomer_" & Session("UserNo") & "_" & passedCustForDetail & ".[DiifThisPeriodVSLast3Dollars] = CustCatPeriodSales.[DiifThisPeriodVSLast3Dollars]"
	SQLCatTmp = SQLCatTmp & " , zCatAnalByPeriod_SingleCustomer_" & Session("UserNo") & "_" & passedCustForDetail & ".[DiifThisPeriodVSLast3Percent] = CustCatPeriodSales.[DiifThisPeriodVSLast3Percent]"
	SQLCatTmp = SQLCatTmp & " , zCatAnalByPeriod_SingleCustomer_" & Session("UserNo") & "_" & passedCustForDetail & ".[DiifThisPeriodVSLast12Dollars] = CustCatPeriodSales.[DiifThisPeriodVSLast12Dollars]"
	SQLCatTmp = SQLCatTmp & " , zCatAnalByPeriod_SingleCustomer_" & Session("UserNo") & "_" & passedCustForDetail & ".[DiifThisPeriodVSLast12Percent] = CustCatPeriodSales.[DiifThisPeriodVSLast12Percent]"
	SQLCatTmp = SQLCatTmp & " , zCatAnalByPeriod_SingleCustomer_" & Session("UserNo") & "_" & passedCustForDetail & ".[PriorPeriod1Sales] = CustCatPeriodSales.[PriorPeriod1Sales]"
	SQLCatTmp = SQLCatTmp & " , zCatAnalByPeriod_SingleCustomer_" & Session("UserNo") & "_" & passedCustForDetail & ".[PriorPeriod2Sales] = CustCatPeriodSales.[PriorPeriod2Sales]"
	SQLCatTmp = SQLCatTmp & " , zCatAnalByPeriod_SingleCustomer_" & Session("UserNo") & "_" & passedCustForDetail & ".[PriorPeriod3Sales] = CustCatPeriodSales.[PriorPeriod3Sales]"
	SQLCatTmp = SQLCatTmp & " , zCatAnalByPeriod_SingleCustomer_" & Session("UserNo") & "_" & passedCustForDetail & ".[PriorPeriod4Sales] = CustCatPeriodSales.[PriorPeriod4Sales]"
	SQLCatTmp = SQLCatTmp & " , zCatAnalByPeriod_SingleCustomer_" & Session("UserNo") & "_" & passedCustForDetail & ".[PriorPeriod5Sales] = CustCatPeriodSales.[PriorPeriod5Sales]"
	SQLCatTmp = SQLCatTmp & " , zCatAnalByPeriod_SingleCustomer_" & Session("UserNo") & "_" & passedCustForDetail & ".[PriorPeriod6Sales] = CustCatPeriodSales.[PriorPeriod6Sales]"
	SQLCatTmp = SQLCatTmp & " , zCatAnalByPeriod_SingleCustomer_" & Session("UserNo") & "_" & passedCustForDetail & ".[PriorPeriod7Sales] = CustCatPeriodSales.[PriorPeriod7Sales]"
	SQLCatTmp = SQLCatTmp & " , zCatAnalByPeriod_SingleCustomer_" & Session("UserNo") & "_" & passedCustForDetail & ".[PriorPeriod8Sales] = CustCatPeriodSales.[PriorPeriod8Sales]"
	SQLCatTmp = SQLCatTmp & " , zCatAnalByPeriod_SingleCustomer_" & Session("UserNo") & "_" & passedCustForDetail & ".[PriorPeriod9Sales] = CustCatPeriodSales.[PriorPeriod9Sales]"
	SQLCatTmp = SQLCatTmp & " , zCatAnalByPeriod_SingleCustomer_" & Session("UserNo") & "_" & passedCustForDetail & ".[PriorPeriod10Sales] = CustCatPeriodSales.[PriorPeriod10Sales]"
	SQLCatTmp = SQLCatTmp & " , zCatAnalByPeriod_SingleCustomer_" & Session("UserNo") & "_" & passedCustForDetail & ".[PriorPeriod11Sales] = CustCatPeriodSales.[PriorPeriod11Sales]"
	SQLCatTmp = SQLCatTmp & " , zCatAnalByPeriod_SingleCustomer_" & Session("UserNo") & "_" & passedCustForDetail & ".[PriorPeriod12Sales] = CustCatPeriodSales.[PriorPeriod12Sales]"
	
	SQLCatTmp = SQLCatTmp & " , zCatAnalByPeriod_SingleCustomer_" & Session("UserNo") & "_" & passedCustForDetail & ".[GroupName] = CustCatPeriodSales.[GroupName]"
	SQLCatTmp = SQLCatTmp & " , zCatAnalByPeriod_SingleCustomer_" & Session("UserNo") & "_" & passedCustForDetail & ".[CategoryName] = CustCatPeriodSales.[CategoryName]"
	SQLCatTmp = SQLCatTmp & " , zCatAnalByPeriod_SingleCustomer_" & Session("UserNo") & "_" & passedCustForDetail & ".[CategoryNameGetTerm] = Upper(CustCatPeriodSales.[CategoryNameGetTerm])"
				
	SQLCatTmp = SQLCatTmp & " FROM zCatAnalByPeriod_SingleCustomer_" & Session("UserNo") & "_" & passedCustForDetail & " INNER JOIN "
	SQLCatTmp = SQLCatTmp & " CustCatPeriodSales ON zCatAnalByPeriod_SingleCustomer_" & Session("UserNo") & "_" & passedCustForDetail & ".ThisPeriodSequenceNumber = CustCatPeriodSales.ThisPeriodSequenceNumber "
   	SQLCatTmp = SQLCatTmp & " AND zCatAnalByPeriod_SingleCustomer_" & Session("UserNo") & "_" & passedCustForDetail & ".CustNum = CustCatPeriodSales.CustNum "
    SQLCatTmp = SQLCatTmp & " AND zCatAnalByPeriod_SingleCustomer_" & Session("UserNo") & "_" & passedCustForDetail & ".Category = CustCatPeriodSales.Category "
    
	Set rsCatTmp = cnnCatTmp.Execute(SQLCatTmp)


	'Now we need to update the P3P and P12P info for any category that has $0 sales because the are not included 
	'in the overnight leakage data build BUT ONLY FOR THIS PERIOD (Whew!)
	
	If MUV_READ("VarianceBasis") = "3Periods" Then 
	
		SQLCatTmp = "SELECT * FROM zCatAnalByPeriod_SingleCustomer_" & Session("UserNo")  & "_" & passedCustForDetail 
		SQLCatTmp = SQLCatTmp & " WHERE DiifThisPeriodVSLast3Dollars = 0 AND"
		SQLCatTmp = SQLCatTmp & " ThisPeriodSequenceNumber = " & passedPeriodSeqBeingEvaluated
    
		Set rsCatTmp = cnnCatTmp.Execute(SQLCatTmp)

		If Not rsCatTmp.EOF Then
			Do While Not rsCatTmp.EOF
			
				SQLLeakageLookup = "SELECT SUM(TotalSales) AS ThreePeriodSales FROM CustCatPeriodSales "
				SQLLeakageLookup = SQLLeakageLookup & " WHERE ThisPeriodSequenceNumber < " & rsCatTmp("ThisPeriodSequenceNumber") & " "
				SQLLeakageLookup = SQLLeakageLookup & " AND ThisPeriodSequenceNumber > " & rsCatTmp("ThisPeriodSequenceNumber") - 4 & " "
				SQLLeakageLookup = SQLLeakageLookup & " AND Category = " & rsCatTmp("Category")
				SQLLeakageLookup = SQLLeakageLookup & " AND CustNum = " & rsCatTmp("CustNum")
	
				Set rsLeakageLookup = cnnCatTmp.Execute(SQLLeakageLookup)

				If Not rsLeakageLookup.EOF Then ThreePeriodSales = rsLeakageLookup("ThreePeriodSales") Else ThreePeriodSales = 0
				If Not IsNumeric(ThreePeriodSales) Then ThreePeriodSales = 0
				
				SQLCatTmpForUpdates ="UPDATE zCatAnalByPeriod_SingleCustomer_" & Session("UserNo") & "_" & passedCustForDetail & " SET "
				SQLCatTmpForUpdates = SQLCatTmpForUpdates & "DiifThisPeriodVSLast3Dollars = TotalSales - " & (ThreePeriodSales/3) 
				SQLCatTmpForUpdates = SQLCatTmpForUpdates & ", [3PriorPeriodsAeverage] = " & (ThreePeriodSales/3) & " WHERE "
				SQLCatTmpForUpdates = SQLCatTmpForUpdates & " ThisPeriodSequenceNumber = " & rsCatTmp("ThisPeriodSequenceNumber") & " "
				SQLCatTmpForUpdates = SQLCatTmpForUpdates & " AND Category = " & rsCatTmp("Category")

				Set rsCatTmpForUpdates = cnnCatTmp.Execute(SQLCatTmpForUpdates)

				rsCatTmp.MoveNext
				
			Loop
		End If
		
	Else

		SQLCatTmp = "SELECT * FROM zCatAnalByPeriod_SingleCustomer_" & Session("UserNo")  & "_" & passedCustForDetail 
		SQLCatTmp = SQLCatTmp & " WHERE DiifThisPeriodVSLast12Dollars = 0 AND"
		SQLCatTmp = SQLCatTmp & " ThisPeriodSequenceNumber = " & passedPeriodSeqBeingEvaluated
    
		Set rsCatTmp = cnnCatTmp.Execute(SQLCatTmp)

		If Not rsCatTmp.EOF Then
			Do While Not rsCatTmp.EOF
			
				SQLLeakageLookup = "SELECT SUM(TotalSales) AS TwelvePeriodSales FROM CustCatPeriodSales "
				SQLLeakageLookup = SQLLeakageLookup & " WHERE ThisPeriodSequenceNumber < " & rsCatTmp("ThisPeriodSequenceNumber") & " "
				SQLLeakageLookup = SQLLeakageLookup & " AND ThisPeriodSequenceNumber > " & rsCatTmp("ThisPeriodSequenceNumber") - 13 & " "
				SQLLeakageLookup = SQLLeakageLookup & " AND Category = " & rsCatTmp("Category")
				SQLLeakageLookup = SQLLeakageLookup & " AND CustNum = " & rsCatTmp("CustNum")
	
				Set rsLeakageLookup = cnnCatTmp.Execute(SQLLeakageLookup)

				If Not rsLeakageLookup.EOF Then TwelvePeriodSales = rsLeakageLookup("TwelvePeriodSales") Else TwelvePeriodSales = 0
				If Not IsNumeric(TwelvePeriodSales) Then TwelvePeriodSales  = 0
				
			
				SQLCatTmpForUpdates ="UPDATE zCatAnalByPeriod_SingleCustomer_" & Session("UserNo") & "_" & passedCustForDetail & " SET "
				SQLCatTmpForUpdates = SQLCatTmpForUpdates & "DiifThisPeriodVSLast12Dollars = TotalSales - " & (TwelvePeriodSales/12) 
				SQLCatTmpForUpdates = SQLCatTmpForUpdates & ", [12PriorPeriodAverage] = " & (TwelvePeriodSales/12) & " WHERE "
				SQLCatTmpForUpdates = SQLCatTmpForUpdates & " ThisPeriodSequenceNumber = " & rsCatTmp("ThisPeriodSequenceNumber") & " "
				SQLCatTmpForUpdates = SQLCatTmpForUpdates & " AND Category = " & rsCatTmp("Category")

				Set rsCatTmpForUpdates = cnnCatTmp.Execute(SQLCatTmpForUpdates)

				rsCatTmp.MoveNext
				
			Loop
		End If

			
	End If
			
	Set rsCatTmpForUpdates = cnnCatTmp.Execute(SQLCatTmpForUpdates)
'Response.Write("10-" & Now() & "<br>")
	Set rsCatTmp = Nothing
	Set rsCatTmpForUpdates = Nothing
	Set rsLeakageLookup = Nothing
End Sub

Function GetTotalSalesForCategByPeriodRange(passedCatID,passedStartPeriodSeq,passedEndPeriodSeq,passedCustForDetail)

	resultGetTotalSalesForCategByPeriodRange = ""

	Set cnnGetTotalSalesForCategByPeriodRange = Server.CreateObject("ADODB.Connection")
	cnnGetTotalSalesForCategByPeriodRange.open Session("ClientCnnString")
		
	SQLGetTotalSalesForCategByPeriodRange = "SELECT Sum(TotalSales) AS GroupTot FROM zCatAnalByPeriod_SingleCustomer_" & Session("UserNo") & "_" & passedCustForDetail & " WHERE "
	SQLGetTotalSalesForCategByPeriodRange = SQLGetTotalSalesForCategByPeriodRange & " Category = " & passedCatID & " AND "
	SQLGetTotalSalesForCategByPeriodRange = SQLGetTotalSalesForCategByPeriodRange & " ThisPeriodSequenceNumber >= " & passedStartPeriodSeq & " AND "
	SQLGetTotalSalesForCategByPeriodRange = SQLGetTotalSalesForCategByPeriodRange & " ThisPeriodSequenceNumber <= " & passedEndPeriodSeq 
 
	Set rsGetTotalSalesForCategByPeriodRange = Server.CreateObject("ADODB.Recordset")
	rsGetTotalSalesForCategByPeriodRange.CursorLocation = 3 
	Set rsGetTotalSalesForCategByPeriodRange = cnnGetTotalSalesForCategByPeriodRange.Execute(SQLGetTotalSalesForCategByPeriodRange)

	If not rsGetTotalSalesForCategByPeriodRange.EOF Then
		resultGetTotalSalesForCategByPeriodRange = rsGetTotalSalesForCategByPeriodRange("GroupTot")
	Else
		resultGetTotalSalesForCategByPeriodRange = 0
	End If

	rsGetTotalSalesForCategByPeriodRange.Close
	set rsGetTotalSalesForCategByPeriodRange= Nothing
	cnnGetTotalSalesForCategByPeriodRange.Close	
	set cnnGetTotalSalesForCategByPeriodRange= Nothing
	
	GetTotalSalesForCategByPeriodRange = resultGetTotalSalesForCategByPeriodRange


End Function

Function TotalSalesByPeriodSeq(passedPeriodSeq,passedCustForDetail)

	resultTotalSalesByPeriodSeq = ""

	Set cnnTotalSalesByPeriodSeq = Server.CreateObject("ADODB.Connection")
	cnnTotalSalesByPeriodSeq.open Session("ClientCnnString")
		
	SQLTotalSalesByPeriodSeq = "SELECT Sum(TotalSales) AS PeriodTotSales FROM zCatAnalByPeriod_SingleCustomer_" & Session("UserNo") & "_" & passedCustForDetail & " WHERE ThisPeriodSequenceNumber = " & passedPeriodSeq
 
	Set rsTotalSalesByPeriodSeq = Server.CreateObject("ADODB.Recordset")
	rsTotalSalesByPeriodSeq.CursorLocation = 3 
	Set rsTotalSalesByPeriodSeq = cnnTotalSalesByPeriodSeq.Execute(SQLTotalSalesByPeriodSeq)

	If not rsTotalSalesByPeriodSeq.EOF Then resultTotalSalesByPeriodSeq = rsTotalSalesByPeriodSeq("PeriodTotSales")

	rsTotalSalesByPeriodSeq.Close
	set rsTotalSalesByPeriodSeq= Nothing
	cnnTotalSalesByPeriodSeq.Close	
	set cnnTotalSalesByPeriodSeq= Nothing
	
	TotalSalesByPeriodSeq = resultTotalSalesByPeriodSeq

End Function

Function GetPriorThreePeriodAverageByCatByPerSeq(passedCatID,passedPeriodSeq,passedCustForDetail)

	resultGetPriorThreePeriodAverageByCatByPerSeq = ""

	Set cnnGetPriorThreePeriodAverageByCatByPerSeq = Server.CreateObject("ADODB.Connection")
	cnnGetPriorThreePeriodAverageByCatByPerSeq.open Session("ClientCnnString")
		
	SQLGetPriorThreePeriodAverageByCatByPerSeq ="SELECT SUM(TotalSales)/3 AS ThreePeriodAverage FROM zCatAnalByPeriod_SingleCustomer_" & Session("UserNo") & "_" & passedCustForDetail
	SQLGetPriorThreePeriodAverageByCatByPerSeq  = SQLGetPriorThreePeriodAverageByCatByPerSeq  & " WHERE ThisPeriodSequenceNumber < " & passedPeriodSeq & " "
	SQLGetPriorThreePeriodAverageByCatByPerSeq  = SQLGetPriorThreePeriodAverageByCatByPerSeq  & " AND ThisPeriodSequenceNumber > " & passedPeriodSeq - 4 & " "
	SQLGetPriorThreePeriodAverageByCatByPerSeq  = SQLGetPriorThreePeriodAverageByCatByPerSeq  & " AND Category = " & passedCatID
 
	Set rsGetPriorThreePeriodAverageByCatByPerSeq = Server.CreateObject("ADODB.Recordset")
	rsGetPriorThreePeriodAverageByCatByPerSeq.CursorLocation = 3 
	Set rsGetPriorThreePeriodAverageByCatByPerSeq = cnnGetPriorThreePeriodAverageByCatByPerSeq.Execute(SQLGetPriorThreePeriodAverageByCatByPerSeq)

	If not rsGetPriorThreePeriodAverageByCatByPerSeq.EOF Then resultGetPriorThreePeriodAverageByCatByPerSeq = rsGetPriorThreePeriodAverageByCatByPerSeq("ThreePeriodAverage")

	rsGetPriorThreePeriodAverageByCatByPerSeq.Close
	set rsGetPriorThreePeriodAverageByCatByPerSeq= Nothing
	cnnGetPriorThreePeriodAverageByCatByPerSeq.Close	
	set cnnGetPriorThreePeriodAverageByCatByPerSeq= Nothing
	
	GetPriorThreePeriodAverageByCatByPerSeq = resultGetPriorThreePeriodAverageByCatByPerSeq

End Function

Function GetPriorTwelvePeriodAverageByCatByPerSeq(passedCatID,passedPeriodSeq,PassedCustForDetail)

	resultGetPriorTwelvePeriodAverageByCatByPerSeq = ""

	Set cnnGetPriorTwelvePeriodAverageByCatByPerSeq = Server.CreateObject("ADODB.Connection")
	cnnGetPriorTwelvePeriodAverageByCatByPerSeq.open Session("ClientCnnString")
		
	SQLGetPriorTwelvePeriodAverageByCatByPerSeq ="SELECT SUM(TotalSales)/12 AS TwelvePeriodAverage FROM zCatAnalByPeriod_SingleCustomer_" & Session("UserNo") & "_" & PassedCustForDetail
	SQLGetPriorTwelvePeriodAverageByCatByPerSeq  = SQLGetPriorTwelvePeriodAverageByCatByPerSeq  & " WHERE ThisPeriodSequenceNumber < " & passedPeriodSeq & " "
	SQLGetPriorTwelvePeriodAverageByCatByPerSeq  = SQLGetPriorTwelvePeriodAverageByCatByPerSeq  & " AND ThisPeriodSequenceNumber > " & passedPeriodSeq - 13 & " "
	SQLGetPriorTwelvePeriodAverageByCatByPerSeq  = SQLGetPriorTwelvePeriodAverageByCatByPerSeq  & " AND Category = " & passedCatID
 
	Set rsGetPriorTwelvePeriodAverageByCatByPerSeq = Server.CreateObject("ADODB.Recordset")
	rsGetPriorTwelvePeriodAverageByCatByPerSeq.CursorLocation = 3 
	Set rsGetPriorTwelvePeriodAverageByCatByPerSeq = cnnGetPriorTwelvePeriodAverageByCatByPerSeq.Execute(SQLGetPriorTwelvePeriodAverageByCatByPerSeq)

	If not rsGetPriorTwelvePeriodAverageByCatByPerSeq.EOF Then resultGetPriorTwelvePeriodAverageByCatByPerSeq = rsGetPriorTwelvePeriodAverageByCatByPerSeq("TwelvePeriodAverage")

	rsGetPriorTwelvePeriodAverageByCatByPerSeq.Close
	set rsGetPriorTwelvePeriodAverageByCatByPerSeq= Nothing
	cnnGetPriorTwelvePeriodAverageByCatByPerSeq.Close	
	set cnnGetPriorTwelvePeriodAverageByCatByPerSeq= Nothing
	
	GetPriorTwelvePeriodAverageByCatByPerSeq = resultGetPriorTwelvePeriodAverageByCatByPerSeq

End Function

Function GetPriorThreePeriodAverageAllCatsByPerSeq(passedPeriodSeq,PassedCustForDetail)

	resultGetPriorThreePeriodAverageAllCatsByPerSeq = ""

	Set cnnGetPriorThreePeriodAverageAllCatsByPerSeq = Server.CreateObject("ADODB.Connection")
	cnnGetPriorThreePeriodAverageAllCatsByPerSeq.open Session("ClientCnnString")
		
	SQLGetPriorThreePeriodAverageAllCatsByPerSeq ="SELECT SUM(TotalSales)/3 AS ThreePeriodAverage FROM zCatAnalByPeriod_SingleCustomer_" & Session("UserNo") & "_" & PassedCustForDetail
	SQLGetPriorThreePeriodAverageAllCatsByPerSeq  = SQLGetPriorThreePeriodAverageAllCatsByPerSeq  & " WHERE ThisPeriodSequenceNumber < " & passedPeriodSeq & " "
	SQLGetPriorThreePeriodAverageAllCatsByPerSeq  = SQLGetPriorThreePeriodAverageAllCatsByPerSeq  & " AND ThisPeriodSequenceNumber > " & passedPeriodSeq - 4 & " "
 
	Set rsGetPriorThreePeriodAverageAllCatsByPerSeq = Server.CreateObject("ADODB.Recordset")
	rsGetPriorThreePeriodAverageAllCatsByPerSeq.CursorLocation = 3 
	Set rsGetPriorThreePeriodAverageAllCatsByPerSeq = cnnGetPriorThreePeriodAverageAllCatsByPerSeq.Execute(SQLGetPriorThreePeriodAverageAllCatsByPerSeq)

	If not rsGetPriorThreePeriodAverageAllCatsByPerSeq.EOF Then resultGetPriorThreePeriodAverageAllCatsByPerSeq = rsGetPriorThreePeriodAverageAllCatsByPerSeq("ThreePeriodAverage")

	rsGetPriorThreePeriodAverageAllCatsByPerSeq.Close
	set rsGetPriorThreePeriodAverageAllCatsByPerSeq= Nothing
	cnnGetPriorThreePeriodAverageAllCatsByPerSeq.Close	
	set cnnGetPriorThreePeriodAverageAllCatsByPerSeq= Nothing
	
	GetPriorThreePeriodAverageAllCatsByPerSeq = resultGetPriorThreePeriodAverageAllCatsByPerSeq

End Function

Function GetPriorTwelvePeriodAverageAllCatsByPerSeq(passedPeriodSeq,passedCustForDetail)

	resultGetPriorTwelvePeriodAverageAllCatsByPerSeq = ""

	Set cnnGetPriorTwelvePeriodAverageAllCatsByPerSeq = Server.CreateObject("ADODB.Connection")
	cnnGetPriorTwelvePeriodAverageAllCatsByPerSeq.open Session("ClientCnnString")
		
	SQLGetPriorTwelvePeriodAverageAllCatsByPerSeq ="SELECT SUM(TotalSales)/12 AS ThreePeriodAverage FROM zCatAnalByPeriod_SingleCustomer_" & Session("UserNo") & "_" & passedCustForDetail
	SQLGetPriorTwelvePeriodAverageAllCatsByPerSeq  = SQLGetPriorTwelvePeriodAverageAllCatsByPerSeq  & " WHERE ThisPeriodSequenceNumber < " & passedPeriodSeq & " "
	SQLGetPriorTwelvePeriodAverageAllCatsByPerSeq  = SQLGetPriorTwelvePeriodAverageAllCatsByPerSeq  & " AND ThisPeriodSequenceNumber > " & passedPeriodSeq - 13 & " "
 
	Set rsGetPriorTwelvePeriodAverageAllCatsByPerSeq = Server.CreateObject("ADODB.Recordset")
	rsGetPriorTwelvePeriodAverageAllCatsByPerSeq.CursorLocation = 3 
	Set rsGetPriorTwelvePeriodAverageAllCatsByPerSeq = cnnGetPriorTwelvePeriodAverageAllCatsByPerSeq.Execute(SQLGetPriorTwelvePeriodAverageAllCatsByPerSeq)

	If not rsGetPriorTwelvePeriodAverageAllCatsByPerSeq.EOF Then resultGetPriorTwelvePeriodAverageAllCatsByPerSeq = rsGetPriorTwelvePeriodAverageAllCatsByPerSeq("ThreePeriodAverage")

	rsGetPriorTwelvePeriodAverageAllCatsByPerSeq.Close
	set rsGetPriorTwelvePeriodAverageAllCatsByPerSeq= Nothing
	cnnGetPriorTwelvePeriodAverageAllCatsByPerSeq.Close	
	set cnnGetPriorTwelvePeriodAverageAllCatsByPerSeq= Nothing
	
	GetPriorTwelvePeriodAverageAllCatsByPerSeq = resultGetPriorTwelvePeriodAverageAllCatsByPerSeq

End Function


Function GetCurrent_UnpostedTotal_ByCust(passedCustID,passedPeriodBeingEvaluated)

If Session("NewPostLogic") = False Then

	StartDateToFind = GetPeriodBeginDateBySeq(passedPeriodBeingEvaluated+1)
	EndDateToFind = GetPeriodEndDateBySeq(passedPeriodBeingEvaluated+1)
	
	resultGetCurrent_UnpostedTotal_ByCust = 0

	Set cnnGetCurrent_UnpostedTotal_ByCust = Server.CreateObject("ADODB.Connection")
	cnnGetCurrent_UnpostedTotal_ByCust.open Session("ClientCnnString")
		
	SQLGetCurrent_UnpostedTotal_ByCust = "SELECT SUM(InvoiceTotal-SalesTaxCharge-Deposit) AS TotalForCurrent FROM Telsel WHERE (InvoiceTFlag = 'O' OR InvoiceTFlag = 'T') AND CustNum = " & passedCustID & " AND ("
	SQLGetCurrent_UnpostedTotal_ByCust = SQLGetCurrent_UnpostedTotal_ByCust & "InvoiceDate >= '" & StartDateToFind & "' AND "
	SQLGetCurrent_UnpostedTotal_ByCust = SQLGetCurrent_UnpostedTotal_ByCust & "InvoiceDate <= '" & EndDateToFind & "')"
'Response.Write(SQLGetCurrent_UnpostedTotal_ByCust)
	Set rsGetCurrent_UnpostedTotal_ByCust = Server.CreateObject("ADODB.Recordset")
	rsGetCurrent_UnpostedTotal_ByCust.CursorLocation = 3 
	Set rsGetCurrent_UnpostedTotal_ByCust = cnnGetCurrent_UnpostedTotal_ByCust.Execute(SQLGetCurrent_UnpostedTotal_ByCust)

	If not rsGetCurrent_UnpostedTotal_ByCust.EOF Then resultGetCurrent_UnpostedTotal_ByCust = rsGetCurrent_UnpostedTotal_ByCust("TotalForCurrent")

	If Not IsNumeric(resultGetCurrent_UnpostedTotal_ByCust) Then resultGetCurrent_UnpostedTotal_ByCust = 0 ' In case there are no results
	
	rsGetCurrent_UnpostedTotal_ByCust.Close
	set rsGetCurrent_UnpostedTotal_ByCust= Nothing
	cnnGetCurrent_UnpostedTotal_ByCust.Close	
	set cnnGetCurrent_UnpostedTotal_ByCust= Nothing
	
Else

	For z=0 to 21
		resultGetCurrent_UnpostedTotal_ByCust = resultGetCurrent_UnpostedTotal_ByCust + PostUnPostCatArray(z,0)
	next  

End If	
	
	GetCurrent_UnpostedTotal_ByCust = resultGetCurrent_UnpostedTotal_ByCust 

End Function



Function GetCurrent_PostedTotal_ByCust(passedCustID,passedPeriodBeingEvaluated)

If Session("NewPostLogic") = False Then

	StartDateToFind = GetPeriodBeginDateBySeq(passedPeriodBeingEvaluated+1)
	EndDateToFind = GetPeriodEndDateBySeq(passedPeriodBeingEvaluated+1)
	
	resultGetCurrent_PostedTotal_ByCust = 0

	Set cnnGetCurrent_PostedTotal_ByCust = Server.CreateObject("ADODB.Connection")
	cnnGetCurrent_PostedTotal_ByCust.open Session("ClientCnnString")
		
	SQLGetCurrent_PostedTotal_ByCust = "SELECT SUM(IvsTotalAmt-IvsSalesTax-IvsDepositChg) AS TotalForCurrent FROM InvoiceHistory WHERE CustNum = " & passedCustID & " AND ("
	SQLGetCurrent_PostedTotal_ByCust = SQLGetCurrent_PostedTotal_ByCust & "IvsDate >= '" & StartDateToFind & "' AND "
	SQLGetCurrent_PostedTotal_ByCust = SQLGetCurrent_PostedTotal_ByCust & "IvsDate <= '" & EndDateToFind & "')"

	Set rsGetCurrent_PostedTotal_ByCust = Server.CreateObject("ADODB.Recordset")
	rsGetCurrent_PostedTotal_ByCust.CursorLocation = 3 
	Set rsGetCurrent_PostedTotal_ByCust = cnnGetCurrent_PostedTotal_ByCust.Execute(SQLGetCurrent_PostedTotal_ByCust)

	If not rsGetCurrent_PostedTotal_ByCust.EOF Then resultGetCurrent_PostedTotal_ByCust = rsGetCurrent_PostedTotal_ByCust("TotalForCurrent")

	If Not IsNumeric(resultGetCurrent_PostedTotal_ByCust) Then resultGetCurrent_PostedTotal_ByCust = 0 ' In case there are no results
	
	rsGetCurrent_PostedTotal_ByCust.Close
	set rsGetCurrent_PostedTotal_ByCust= Nothing
	cnnGetCurrent_PostedTotal_ByCust.Close	
	set cnnGetCurrent_PostedTotal_ByCust= Nothing

Else

	For z=0 to 21
		resultGetCurrent_PostedTotal_ByCust = resultGetCurrent_PostedTotal_ByCust + PostUnPostCatArray(z,1)
	next  

End If	
	
	GetCurrent_PostedTotal_ByCust = resultGetCurrent_PostedTotal_ByCust

End Function


Function GetCaseVarianceByCust_ALLCats(passedPeriodSeq,passedVarianceBasis,passedCustID)

	resultGetCaseVarianceByCust_ALLCats = 0
	
	FirstPeriod_TotalCases = 0
	
	SQLGetCaseVarianceByCust_ALLCats = "SELECT SUM(NumberOfCases) AS TotCases FROM " & MUV_Read("SQL_Owner") & ".InvoiceHistoryDetail"
	If passedVarianceBasis = "3Periods" Then
		SQLGetCaseVarianceByCust_ALLCats = SQLGetCaseVarianceByCust_ALLCats & " WHERE ivsDate BETWEEN '" & GetPeriodBeginDateBySeq(passedPeriodSeq-3) & "' AND '" & GetPeriodEndDateBySeq(passedPeriodSeq-1) & "' "
	Else
		SQLGetCaseVarianceByCust_ALLCats = SQLGetCaseVarianceByCust_ALLCats & " WHERE ivsDate BETWEEN '" & GetPeriodBeginDateBySeq(passedPeriodSeq-12) & "' AND '" & GetPeriodEndDateBySeq(passedPeriodSeq-1) & "' "
	End If
	SQLGetCaseVarianceByCust_ALLCats = SQLGetCaseVarianceByCust_ALLCats & " AND CustNum = " & passedCustID & " "

	Set cnnGetCaseVarianceByCust_ALLCats = Server.CreateObject("ADODB.Connection")
	cnnGetCaseVarianceByCust_ALLCats.open Session("ClientCnnString")

	Set rsGetCaseVarianceByCust_ALLCats = Server.CreateObject("ADODB.Recordset")
	rsGetCaseVarianceByCust_ALLCats.CursorLocation = 3
	'rsGetCaseVarianceByCust_ALLCats.Open SQLGetCaseVarianceByCust_ALLCats, Session("ClientCnnString")

	Set rsGetCaseVarianceByCust_ALLCats = cnnGetCaseVarianceByCust_ALLCats.Execute(SQLGetCaseVarianceByCust_ALLCats)
	
	If not rsGetCaseVarianceByCust_ALLCats.eof Then

		If VarianceBasis = "3Periods" Then		
			FirstPeriod_TotalCases = rsGetCaseVarianceByCust_ALLCats("TotCases")/3
		Else
			FirstPeriod_TotalCases = rsGetCaseVarianceByCust_ALLCats("TotCases")/12
		End If

	End If
	rsGetCaseVarianceByCust_ALLCats.Close
	
	SecondPeriod_TotalCases = 0

	SQLGetCaseVarianceByCust_ALLCats = "SELECT SUM(NumberOfCases) AS TotCases FROM " & MUV_Read("SQL_Owner") & ".InvoiceHistoryDetail"
	SQLGetCaseVarianceByCust_ALLCats = SQLGetCaseVarianceByCust_ALLCats & " WHERE ivsDate BETWEEN '" & GetPeriodBeginDateBySeq(passedPeriodSeq) & "' AND '" & GetPeriodEndDateBySeq(passedPeriodSeq) & "' "
	SQLGetCaseVarianceByCust_ALLCats = SQLGetCaseVarianceByCust_ALLCats & " AND CustNum = " & passedCustID & " "

	Set rsGetCaseVarianceByCust_ALLCats2 = Server.CreateObject("ADODB.Recordset")
	rsGetCaseVarianceByCust_ALLCats2.CursorLocation = 3
	'rsGetCaseVarianceByCust_ALLCats2.Open SQLGetCaseVarianceByCust_ALLCats, Session("ClientCnnString")

	Set rsGetCaseVarianceByCust_ALLCats2 = cnnGetCaseVarianceByCust_ALLCats.Execute(SQLGetCaseVarianceByCust_ALLCats)
	
	If not rsGetCaseVarianceByCust_ALLCats2.eof Then SecondPeriod_TotalCases = rsGetCaseVarianceByCust_ALLCats2("TotCases")

	rsGetCaseVarianceByCust_ALLCats2.Close

	resultGetCaseVarianceByCust_ALLCats = (SecondPeriod_TotalCases - FirstPeriod_TotalCases)

	If Not IsNumeric(resultGetCaseVarianceByCust_ALLCats) Then resultGetCaseVarianceByCust_ALLCats = 0

	GetCaseVarianceByCust_ALLCats = resultGetCaseVarianceByCust_ALLCats

End Function


Function GetCaseVarianceByCustByCat(passedPeriodSeq,passedCategory,passedVarianceBasis,passedCustID)

	resultGetCaseVarianceByCustByCat = ""
	
	SQLGetCaseVarianceByCustByCat = "SELECT SUM(NumberOfCases) AS TotCases FROM " & MUV_Read("SQL_Owner") & ".InvoiceHistoryDetail"
	If passedVarianceBasis = "3Periods" Then
		SQLGetCaseVarianceByCustByCat = SQLGetCaseVarianceByCustByCat & " WHERE ivsDate BETWEEN '" & GetPeriodBeginDateBySeq(passedPeriodSeq-3) & "' AND '" & GetPeriodEndDateBySeq(passedPeriodSeq-1) & "' "
	Else
		SQLGetCaseVarianceByCustByCat = SQLGetCaseVarianceByCustByCat & " WHERE ivsDate BETWEEN '" & GetPeriodBeginDateBySeq(passedPeriodSeq-12) & "' AND '" & GetPeriodEndDateBySeq(passedPeriodSeq-1) & "' "
	End If
	SQLGetCaseVarianceByCustByCat = SQLGetCaseVarianceByCustByCat & " AND CustNum = " & passedCustID & " "
	SQLGetCaseVarianceByCustByCat = SQLGetCaseVarianceByCustByCat & " AND prodCategory = " & passedCategory

	Set rsGetCaseVarianceByCustByCat = Server.CreateObject("ADODB.Recordset")
	rsGetCaseVarianceByCustByCat.CursorLocation = 3
	rsGetCaseVarianceByCustByCat.Open SQLGetCaseVarianceByCustByCat, Session("ClientCnnString")
	
	If Not rsGetCaseVarianceByCustByCat.eof Then
		If Not IsNull(rsGetCaseVarianceByCustByCat("TotCases")) Then
			If VarianceBasis = "3Periods" Then		
				FirstPeriod_TotalCases = rsGetCaseVarianceByCustByCat("TotCases")/3
			Else
				FirstPeriod_TotalCases = rsGetCaseVarianceByCustByCat("TotCases")/12
			End If
		Else
			FirstPeriod_TotalCases = 0
		End If
	Else
		FirstPeriod_TotalCases = 0
	End If

	rsGetCaseVarianceByCustByCat.Close

	SQLGetCaseVarianceByCustByCat = "SELECT SUM(NumberOfCases) AS TotCases FROM " & MUV_Read("SQL_Owner") & ".InvoiceHistoryDetail"
	SQLGetCaseVarianceByCustByCat = SQLGetCaseVarianceByCustByCat & " WHERE ivsDate BETWEEN '" & GetPeriodBeginDateBySeq(passedPeriodSeq) & "' AND '" & GetPeriodEndDateBySeq(passedPeriodSeq) & "' "
	SQLGetCaseVarianceByCustByCat = SQLGetCaseVarianceByCustByCat & " AND CustNum = " & passedCustID & " "
	SQLGetCaseVarianceByCustByCat = SQLGetCaseVarianceByCustByCat & " AND prodCategory = " & passedCategory 
	
	Set rsGetCaseVarianceByCustByCat2 = Server.CreateObject("ADODB.Recordset")
	rsGetCaseVarianceByCustByCat2.CursorLocation = 3
	rsGetCaseVarianceByCustByCat2.Open SQLGetCaseVarianceByCustByCat, Session("ClientCnnString")

	If Not rsGetCaseVarianceByCustByCat2.eof Then
		If Not IsNull(rsGetCaseVarianceByCustByCat2("TotCases")) Then
			SecondPeriod_TotalCases = rsGetCaseVarianceByCustByCat2("TotCases")
		Else
			SecondPeriod_TotalCases = 0
		End If
	Else
		SecondPeriod_TotalCases = 0
	End If

	rsGetCaseVarianceByCustByCat2.Close

	resultGetCaseVarianceByCustByCat = SecondPeriod_TotalCases - FirstPeriod_TotalCases
	
	GetCaseVarianceByCustByCat = cint(resultGetCaseVarianceByCustByCat)

End Function


Function TotalCostByPeriodSeq(passedPeriodSeq,passedCustForDetail)

	resultTotalCostByPeriodSeq = ""

	Set cnnTotalCostByPeriodSeq = Server.CreateObject("ADODB.Connection")
	cnnTotalCostByPeriodSeq.open Session("ClientCnnString")
		
	SQLTotalCostByPeriodSeq = "SELECT Sum(TotalCost) AS PeriodTotCost FROM zCatAnalByPeriod_SingleCustomer_" & Session("UserNo") & "_" & passedCustForDetail & " WHERE ThisPeriodSequenceNumber = " & passedPeriodSeq
 
	Set rsTotalCostByPeriodSeq = Server.CreateObject("ADODB.Recordset")
	rsTotalCostByPeriodSeq.CursorLocation = 3 
	Set rsTotalCostByPeriodSeq = cnnTotalCostByPeriodSeq.Execute(SQLTotalCostByPeriodSeq)

	If not rsTotalCostByPeriodSeq.EOF Then resultTotalCostByPeriodSeq = rsTotalCostByPeriodSeq("PeriodTotCost")

	rsTotalCostByPeriodSeq.Close
	set rsTotalCostByPeriodSeq= Nothing
	cnnTotalCostByPeriodSeq.Close	
	set cnnTotalCostByPeriodSeq= Nothing
	
	TotalCostByPeriodSeq = resultTotalCostByPeriodSeq

End Function



Function GetCurrent_PostedTotalCost_ByCust()

	For z=0 to 21
		resultGetCurrent_PostedTotalCost_ByCust = resultGetCurrent_PostedTotalCost_ByCust + PostUnPostCatArrayCost(z,1)
	next  
	
	GetCurrent_PostedTotalCost_ByCust = resultGetCurrent_PostedTotalCost_ByCust

End Function


Function GetCurrent_UnpostedTotalCost_ByCust()

	For z=0 to 21
		resultGetCurrent_UnpostedTotalCost_ByCust = resultGetCurrent_UnpostedTotalCost_ByCust + PostUnPostCatArrayCost(z,0)
	next  
	
	GetCurrent_UnpostedTotalCost_ByCust = resultGetCurrent_UnpostedTotalCost_ByCust 

End Function

Function TimeToRebuild(passedTableName)

	resultTimeToRebuild = True
	
	Set cnnTimeToRebuild = Server.CreateObject("ADODB.Connection")
	cnnTimeToRebuild.open Session("ClientCnnString")
	Set rsTimeToRebuild = Server.CreateObject("ADODB.Recordset")
	rsTimeToRebuild.CursorLocation = 3
	
	'First see if the table exists at all
	Err.Clear
	on error resume next
	
	SQLTimeToRebuild = "SELECT TOP 1 RecordCreationDateTime AS OldestRecord FROM " & passedTableName & " ORDER BY RecordCreationDateTime"
	Set rsTimeToRebuild = cnnTimeToRebuild.Execute(SQLTimeToRebuild)

	
	If Err.Description = "" Then ' table is there
		If Not rsTimeToRebuild.EOF Then
			'Check oldest record date & time 
			If DateDiff("h",rsTimeToRebuild("OldestRecord"),Now()) < 9 Then resultTimeToRebuild = False ' Table is there and less than 9 hours old
			If Hour(Now()) < 9 Then ' If it is earlier than 9AM, then we need to make sure the oldest record is more recent
				If Hour(rsTimeToRebuild("OldestRecord")) - Hour(Now()) > 4 Then resultTimeToRebuild = True ' It's 9 and the record is older than 5am
			End If
		End If
	End If
	
	set rsTimeToRebuild = Nothing
	cnnTimeToRebuild.Close	
	set cnnTimeToRebuild = Nothing
	
	TimeToRebuild = resultTimeToRebuild 

End Function

%>


<!-- add background on click jQuery -->
<script>
	if($("input:radio[class='active_radio'][value='3Periods']").is(":checked")) {
      $('.td-three-period').addClass('paleYellow'); 
}

	else if($("input:radio[class='active_radio'][value='12Periods']").is(":checked")) {
      $('.td-twelve-period').addClass('paleYellow'); 
}
</script>
<!-- eof add background on click jQuery -->




<!-- ************************************************************************** -->
<!-- MODALS FOR EDITING CATEGORY NOTES, MEMOS AND EQUIPMENT                     -->
<!-- ************************************************************************** -->

<!--#include file="CatAnalByPeriod_Modals.asp"-->

<!-- ************************************************************************** -->
<!-- ************************************************************************** -->


<!--#include file="../../../inc/footer-main.asp"-->