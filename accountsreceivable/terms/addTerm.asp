<!--#include file="../../inc/header.asp"-->


<link rel="stylesheet" type="text/css" href="<%= BaseURL %>css/tabs.css" />

<SCRIPT LANGUAGE="JavaScript">
<!--
    function validatecustomerForm()
    {

        if (document.frmAddTerm.txtTermDescription.value == "") {
            swal("Term Description can not be blank.");
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

<h1 class="page-header"> Add New Term</h1>
<div class="custom-container">

<form method="POST" action="addTerm_submit.asp" name="frmAddTerm" id="frmAddTerm" onsubmit="return validatecustomerForm();">

	<div class="row row-line">
		
			<div class="form-group col-lg-12">
				<label for="txtTermDescription" class="col-sm-4 control-label"><strong>Term Description</strong></label>	
    			<div class="col-sm-8">
    				<input type="text" class="form-control required" id="txtTermDescription" name="txtTermDescription" >
    			</div>
			</div>
		
		
			<div class="form-group col-lg-12">
				<label for="txtfirstTermsPercent" class="col-sm-4 control-label"><strong>First Terms Percent</strong></label>
    			<div class="col-sm-8">
    				<input type="text" class="form-control" id="txtfirstTermsPercent" name="txtfirstTermsPercent" >
    			</div>
			</div>

			<div class="form-group col-lg-12">
				<label for="txtfirstTermsPeriod" class="col-sm-4 control-label"><strong>First Terms Period</strong></label>
    			<div class="col-sm-8">
    				<input type="text" class="form-control" id="txtfirstTermsPeriod" name="txtfirstTermsPeriod" >
    			</div>
			</div>

			<div class="form-group col-lg-12">
				<label for="txtsecondTermsPeriod" class="col-sm-4 control-label"><strong>Second Terms Period</strong></label>
    			<div class="col-sm-8">
    				<input type="text" class="form-control" id="txtsecondTermsPeriod" name="txtsecondTermsPeriod" >
    			</div>
			</div>

			<div class="form-group col-lg-12">
				<label for="txtTermsType" class="col-sm-4 control-label"><strong>Terms Type</strong></label>
    			<div class="col-sm-8">
    				<input type="text" class="form-control" id="txtTermsType" name="txtTermsType" >
    			</div>
			</div>

			<div class="form-group col-lg-12">
				<label for="txtCreditCardBill" class="col-sm-4 control-label"><strong>Credit Card Bill</strong></label>
    			<div class="col-sm-8">
    				<input type="text" class="form-control" id="txtCreditCardBill" name="txtCreditCardBill" >
    			</div>
			</div>
		</div>

			
		<div class="row row-line">	
	    <!-- cancel / submit !-->
			<div class="col-lg-12 alertbutton">
				<div class="col-lg-12">
					<a href="<%= BaseURL %>accountsreceivable/terms/main.asp">
	    				<button type="button" class="btn btn-default">&lsaquo; Cancel &amp; Go Back To Term List</button>
					</a>
					<button type="submit" class="btn btn-primary"><i class="far fa-save"></i> Save</button>
				</div>
		    </div>
		</div>	

</form>
</div>


<!--#include file="../../inc/footer-main.asp"-->
