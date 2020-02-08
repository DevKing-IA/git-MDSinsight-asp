<!--#include file="../../inc/header.asp"-->


<link rel="stylesheet" type="text/css" href="<%= BaseURL %>css/tabs.css" />

<SCRIPT LANGUAGE="JavaScript">
<!--
    function validatecustomerForm()
    {

        if (document.frmAddCustomer.txtCustDescription.value == "") {
            swal("Customer Type Description can not be blank.");
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

<h1 class="page-header"> Add New Customer Type</h1>
<form method="POST" action="addcustomerType_submit.asp" name="frmAddCustomer" id="frmAddCustomer" onsubmit="return validatecustomerForm();">

<div class="col-md-12 custTable clearfix p-h-n">
	<div class="row">
		<div class="col-md-4">
			<div class="form-group col-lg-12">
				<label for="txtCustDescription" class="col-sm-4 control-label"><strong>Customer Type</strong></label>	
    			<div class="col-sm-8">
    				<input type="text" class="form-control required" id="txtCustDescription" name="txtCustDescription" >
    			</div>
			</div>
		</div>	
		
		<div class="col-md-4">		
			<div class="form-group col-lg-12">
				<label for="txtIvsComment1" class="col-sm-4 control-label"><strong>Invoice Comment 1</strong></label>
    			<div class="col-sm-8">
    				<!--<input type="text" class="form-control" id="txtIvsComment1" name="txtIvsComment1" >-->
					<textarea class="form-control" id="txtIvsComment1" name="txtIvsComment1" rows="4" onkeyup="countChar(this, 1)"></textarea>
					<div id="charNum1" align="right"><strong><small>(1000 Chars Remaining)</small></strong></div>
    			</div>
			</div>

			<div class="form-group col-lg-12">
				<label for="txtIvsComment2" class="col-sm-4 control-label"><strong>Invoice Comment 2</strong></label>
    			<div class="col-sm-8">
    				<!--<input type="text" class="form-control" id="txtIvsComment2" name="txtIvsComment2" >-->
					<textarea class="form-control" id="txtIvsComment2" name="txtIvsComment2" rows="4" onkeyup="countChar(this, 2)"></textarea>
					<div id="charNum2" align="right"><strong><small>(1000 Chars Remaining)</small></strong></div>
    			</div>
			</div>

			<div class="form-group col-lg-12">
				<label for="txtIvsComment3" class="col-sm-4 control-label"><strong>Invoice Comment 3</strong></label>
    			<div class="col-sm-8">
    				<!--<input type="text" class="form-control" id="txtIvsComment3" name="txtIvsComment3" >-->
					<textarea class="form-control" id="txtIvsComment3" name="txtIvsComment3" rows="4" onkeyup="countChar(this, 3)"></textarea>
					<div id="charNum3" align="right"><strong><small>(1000 Chars Remaining)</small></strong></div>
    			</div>
			</div>

			<div class="form-group col-lg-12">
				<label for="txtIvsComment4" class="col-sm-4 control-label"><strong>Invoice Comment 4</strong></label>
    			<div class="col-sm-8">
    				<!--<input type="text" class="form-control" id="txtIvsComment4" name="txtIvsComment4" >-->
					<textarea class="form-control" id="txtIvsComment4" name="txtIvsComment4" rows="4" onkeyup="countChar(this, 4)"></textarea>
					<div id="charNum4" align="right"><strong><small>(1000 Chars Remaining)</small></strong></div>
    			</div>
			</div>

			<div class="form-group col-lg-12">
				<label for="txtIvsComment5" class="col-sm-4 control-label"><strong>Invoice Comment 5</strong></label>
    			<div class="col-sm-8">
    				<!--<input type="text" class="form-control" id="txtIvsComment5" name="txtIvsComment5" >-->
					<textarea class="form-control" id="txtIvsComment5" name="txtIvsComment5" rows="4" onkeyup="countChar(this, 5)"></textarea>
					<div id="charNum5" align="right"><strong><small>(1000 Chars Remaining)</small></strong></div>
    			</div>
			</div>
		</div>

		<div class="col-md-4">		
			<div class="form-group col-lg-12">
				<label for="txtHoldDays" class="col-sm-4 control-label"><strong>Hold Days</strong></label>	
    			<div class="col-sm-8">
    				<input type="text" class="form-control form-control-sm" id="txtHoldDays" name="txtHoldDays" >
    			</div>
			</div>

			<div class="form-group col-lg-12">
				<label for="txtHoldAmt" class="col-sm-4 control-label"><strong>Hold Amount</strong></label>	
    			<div class="col-sm-8">
    				<input type="text" class="form-control form-control-sm" id="txtHoldAmt" name="txtHoldAmt" >
    			</div>
			</div>

			<div class="form-group col-lg-12">
				<label for="txtWholesaleFlag" class="col-sm-4 control-label"><strong>Wholesale Flag</strong></label>	
    			<div class="col-sm-8">
    				<input type="text" class="form-control  form-control-sm" id="txtWholesaleFlag" name="txtWholesaleFlag" >
    			</div>
			</div>

			<div class="form-group col-lg-12">
				<label for="txtMemoMessagingFlag" class="col-sm-4 control-label"><strong>Memo Messaging Flag</strong></label>	
    			<div class="col-sm-8">
    				<input type="text" class="form-control  form-control-sm" id="txtMemoMessagingFlag" name="txtMemoMessagingFlag" >
    			</div>
			</div>			
			
	    <!-- cancel / submit !-->
			<div class="col-lg-12 alertbutton">
				<div class="col-lg-12 text-right">
					<a href="<%= BaseURL %>accountsreceivable/customertype/main.asp">
	    				<button type="button" class="btn btn-default">&lsaquo; Cancel &amp; Go Back To Customer Type List</button>
					</a>
					<button type="submit" class="btn btn-primary"><i class="far fa-save"></i> Save</button>
				</div>
		    </div>
		</div>	

	</div>		
</div>
</form>

<!--#include file="../../inc/footer-main.asp"-->
