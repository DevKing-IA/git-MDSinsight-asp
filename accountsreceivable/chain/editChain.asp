<!--#include file="../../inc/header.asp"-->

<% InternalRecordIdentifier = Request.QueryString("i") 
If InternalRecordIdentifier = "" Then Response.Redirect("main.asp")
%>

<link rel="stylesheet" type="text/css" href="<%= BaseURL %>css/tabs.css" />


<SCRIPT LANGUAGE="JavaScript">
<!--
    function validateEditChainForm()
    {

        if (document.frmEditChain.txtChainDescription.value == "") {
            swal("Chain Description can not be blank.");
            return false;
        }

        return true;

    }
// -->
</SCRIPT>          

    <script>
      function countChar(val,i) {
        var len = val.value.length;
        if (len > 1000) {
          val.value = val.value.substring(0, 1000);
        } else {
		var remain = 1000 - len;
          $('#charNum'+i).text("("+ remain +" chars remaining}");
        }
      };
    </script>
	
<!-- password strength meter !-->

<style type="text/css">

.pass-strength h5{
	margin-top: 0px;
	color: #000;
}
.popover.primary {
    border-color:#337ab7;
}
.popover.primary>.arrow {
    border-top-color:#337ab7;
}
.popover.primary>.popover-title {
    color:#fff;
    background-color:#337ab7;
    border-color:#337ab7;
}
.popover.success {
    border-color:#d6e9c6;
}
.popover.success>.arrow {
    border-top-color:#d6e9c6;
}
.popover.success>.popover-title {
    color:#3c763d;
    background-color:#dff0d8;
    border-color:#d6e9c6;
}
.popover.info {
    border-color:#bce8f1;
}
.popover.info>.arrow {
    border-top-color:#bce8f1;
}
.popover.info>.popover-title {
    color:#31708f;
    background-color:#d9edf7;
    border-color:#bce8f1;
}
.popover.warning {
    border-color:#faebcc;
}
.popover.warning>.arrow {
    border-top-color:#faebcc;
}
.popover.warning>.popover-title {
    color:#8a6d3b;
    background-color:#fcf8e3;
    border-color:#faebcc;
}
.popover.danger {
    border-color:#ebccd1;
}
.popover.danger>.arrow {
    border-top-color:#ebccd1;
}
.popover.danger>.popover-title {
    color:#a94442;
    background-color:#f2dede;
    border-color:#ebccd1;
}

.select-line{
	margin-bottom: 15px;
}

.enable-disable{
	margin-top:20px;
}

.row-line{
	margin-bottom: 25px;
}

.table th, tr, td{
	font-weight: normal;
}

.table>thead>tr>th{
	border: 0px;
}
.table thead>tr>th,.table tbody>tr>th,.table tfoot>tr>th,.table thead>tr>td,.table tbody>tr>td,.table tfoot>tr>td{
border:0px;
}

.when-col{
	width: 10%;
}

.reference-col{
	width: 45%;
}

.has-more-col{
	width: 12%;
}

.form-control{
	min-width: 100px;
}

.textarea-box{
	min-width: 260px;
}

.custom-container{
	max-width:600px;
	margin:0 auto;
}

.control-label{
	font-size:12px;
	font-weight:normal;
	padding-top:10px;
}
.control-label-last{
	padding-top:0px;
}

.required{
	border-left:3px solid red;
}
.p-h-n{
	padding-right: 0!important;
    padding-left: 0!important;
}
.custTable .col-lg-12, .custTable .col-sm-8, .custTable .col-sm-9{
	padding-right: 0!important;
    padding-left: 0!important;
}

.custTable{
	padding-right: 5!important;
    padding-left: 5!important;
}


.custTable .col-sm-4{
	padding-right: 5!important;
    padding-left: 0!important;
}

.custTable .form-control-sm{
	width: 40%!important;
}
	</style>
<!-- eof password strength meter !-->



<%
SQL = "SELECT * FROM AR_Chain where InternalRecordIdentifier = " & InternalRecordIdentifier 

Set cnncust = Server.CreateObject("ADODB.Connection")
cnncust.open (Session("ClientCnnString"))
Set rscust = Server.CreateObject("ADODB.Recordset")
rscust.CursorLocation = 3 
Set rscust = cnncust.Execute(SQL)
	
If not rscust.EOF Then	
	Description = rscust("Description")
	updateDiscount = rscust("updateDiscount")
	SellOnlyQuoted = rscust("SellOnlyQuoted")
	chainPrice = rscust("chainPrice")
	poFlag = rscust("poFlag")
	purchaseOrder = rscust("purchaseOrder")
	programType = rscust("programType")
	primarySalesman = rscust("primarySalesman")
	webRequiredFields = rscust("webRequiredFields")
	defQuoteValidDate = rscust("defQuoteValidDate")
End If
set rscust = Nothing
cnncust.close
set cnncust = Nothing

%>


<h1 class="page-header"> Edit Chain</h1>
<form method="POST" action="editChain_submit.asp" name="frmEditChain" id="frmEditChain" onsubmit="return validateEditChainForm();">

<div class="col-md-12 custTable clearfix p-h-n">
	<div class="row">		
		<div class="col-md-4">
			<input type="hidden" id="txtInternalRecordIdentifier" name="txtInternalRecordIdentifier" value="<%=InternalRecordIdentifier%>">

			<div class="form-group col-lg-12">
				<label for="txtChainDescription" class="col-sm-4 control-label"><strong>Chain Description</strong></label>	
    			<div class="col-sm-8">
    				<!--<input type="text" class="form-control required" id="txtChainDescription" name="txtChainDescription" value="<%=Description%>">-->
					<textarea class="form-control required" id="txtChainDescription" name="txtChainDescription" rows="4" onkeyup="countChar(this, 1)"><%=Description%></textarea>
					<%
					If len(Description)>0 Then
						commentCount = len(Description)
					Else
						commentCount = 0
					End If	
					%>					
					<div id="charNum1" align="right"><strong><small>(<%=1000-(commentCount) %> Chars Remaining)</small></strong></div>
    			</div>
			</div>
		</div>

		<div class="col-md-4">	
			<div class="form-group col-lg-12">
				<label for="txtUpdateDiscount" class="col-sm-4 control-label"><strong>Update Discount</strong></label>	
    			<div class="col-sm-8">
    				<input type="text" class="form-control" id="txtUpdateDiscount" name="txtUpdateDiscount" value="<%=updateDiscount%>">
    			</div>
			</div>
			
			<div class="form-group col-lg-12">
				<label for="txtSellOnlyQuoted" class="col-sm-4 control-label"><strong>Sell Only Quoted Items</strong></label>
    			<div class="col-sm-8">
    				<input type="text" class="form-control" id="txtSellOnlyQuoted" name="txtSellOnlyQuoted" value="<%=SellOnlyQuoted%>">
    			</div>
			</div>

			<div class="form-group col-lg-12">
				<label for="txtChainPrice" class="col-sm-4 control-label"><strong>Chain Price</strong></label>
    			<div class="col-sm-8">
    				<input type="text" class="form-control" id="txtChainPrice" name="txtChainPrice" value="<%=chainPrice%>">
    			</div>
			</div>

			<div class="form-group col-lg-12">
				<label for="txtPoFlag" class="col-sm-4 control-label"><strong>Purchase Order Flag</strong></label>
    			<div class="col-sm-8">
    				<input type="text" class="form-control" id="txtPoFlag" name="txtPoFlag" value="<%=poFlag%>">
    			</div>
			</div>

			<div class="form-group col-lg-12">
				<label for="txtPurchaseOrder" class="col-sm-4 control-label"><strong>Purchase Order</strong></label>
    			<div class="col-sm-8">
    				<input type="text" class="form-control" id="txtPurchaseOrder" name="txtPurchaseOrder" value="<%=purchaseOrder%>">
    			</div>
			</div>

			<div class="form-group col-lg-12">
				<label for="txtProgramType" class="col-sm-4 control-label"><strong>Program Type</strong></label>
    			<div class="col-sm-8">
    				<input type="text" class="form-control" id="txtProgramType" name="txtProgramType" value="<%=programType%>">
    			</div>
			</div>

			<div class="form-group col-lg-12">
				<label for="txtPrimarySalesman" class="col-sm-4 control-label"><strong>Primary Salesman</strong></label>
    			<div class="col-sm-8">
    				<input type="text" class="form-control" id="txtPrimarySalesman" name="txtPrimarySalesman" value="<%=primarySalesman%>">
    			</div>
			</div>

			<div class="form-group col-lg-12">
				<label for="txtWebRequiredFields" class="col-sm-4 control-label"><strong>Web Required Fields</strong></label>
    			<div class="col-sm-8">
    				<input type="text" class="form-control" id="txtWebRequiredFields" name="txtWebRequiredFields" value="<%=webRequiredFields%>">
    			</div>
			</div>

			<div class="form-group col-lg-12">
				<label for="txtDefQuoteValidDate" class="col-sm-4 control-label"><strong>Default Quote Valid Date</strong></label>
    			<div class="col-sm-8">
    				<input type="text" class="form-control" id="txtDefQuoteValidDate" name="txtDefQuoteValidDate" value="<%=defQuoteValidDate%>">
    			</div>
			</div>

	    <!-- cancel / submit !-->
			<div class="col-lg-12 alertbutton">
				<div class="col-lg-12">
					<a href="<%= BaseURL %>accountsreceivable/chain/main.asp">
	    				<button type="button" class="btn btn-default">&lsaquo; Cancel &amp; Go Back To Chain List</button>
					</a>
					<button type="submit" class="btn btn-primary"><i class="far fa-save"></i> Save</button>
				</div>
		    </div>
			
		</div>	
	</div>
</form>

<!--#include file="../../inc/footer-main.asp"-->
