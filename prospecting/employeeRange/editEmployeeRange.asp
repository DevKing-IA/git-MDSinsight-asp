<!--#include file="../../inc/header.asp"-->

<% InternalRecordIdentifier = Request.QueryString("i") 
If InternalRecordIdentifier = "" Then Response.Redirect("main.asp")
%>

<link rel="stylesheet" type="text/css" href="<%= BaseURL %>css/tabs.css" />


<SCRIPT LANGUAGE="JavaScript">
<!--

	function isInt(value) {
	  return !isNaN(value) && (function(x) { return (x | 0) === x; })(parseFloat(value))
	}

    function validateEditEmployeeRangeForm()
    {

        if (document.frmeditEmployeeRange.txtEmployeeRange.value == "") {
            swal("Employee Range can not be blank.");
            return false;
        }

        if (document.frmeditEmployeeRange.txtProjectedGPSpend.value == "") {
            swal("Projected GP Spend cannot be blank.");
            return false;
        }
        
        if (!isInt(document.frmeditEmployeeRange.txtProjectedGPSpend.value)) {
             swal("Projected GP Spend must be a whole number.");
            return false;
        }


        return true;

    }
// -->
</SCRIPT>          

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
	</style>
<!-- eof password strength meter !-->



<%
SQL = "SELECT * FROM PR_EmployeeRangeTable where InternalRecordIdentifier = " & InternalRecordIdentifier 

Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3 
Set rs = cnn8.Execute(SQL)
	
If not rs.EOF Then
	EmployeeRange = rs("Range")
	ProjectedGPSpend = rs("ProjectedGPSpend")
End If
set rs = Nothing
cnn8.close
set cnn8 = Nothing

%>


<h1 class="page-header"> Edit Employee Range</h1>

<div class="custom-container">

	<form method="POST" action="editEmployeeRange_submit.asp" name="frmeditEmployeeRange" id="frmeditEmployeeRange" onsubmit="return validateEditEmployeeRangeForm();">

		<div class="row row-line">

			<input type="hidden" id="txtInternalRecordIdentifier" name="txtInternalRecordIdentifier" value="<%=InternalRecordIdentifier%>">
    				
			<div class="form-group col-lg-12">
				<label for="txtEmployeeRange" class="col-sm-3 control-label">Employee Range</label>	
    			<div class="col-sm-6">
    				<input type="text" class="form-control required" id="txtEmployeeRange" name="txtEmployeeRange" value="<%=EmployeeRange%>">
    			</div>
			</div>
    				
			<div class="form-group col-lg-12">
				<label for="txtProjectedGPSpend" class="col-sm-3 control-label">Projected GP Spend</label>	
    			<div class="col-sm-6">
    				<input type="text" class="form-control required" id="txtProjectedGPSpend" name="txtProjectedGPSpend" value="<%=ProjectedGPSpend%>">
    			</div>
			</div>

		</div>
		
	    <!-- cancel / submit !-->
		<div class="row row-line">
			<div class="col-lg-12 alertbutton">
				<div class="col-lg-12">
					<a href="<%= BaseURL %>filemaint/prospecting/EmployeeRange/main.asp">
	    				<button type="button" class="btn btn-default">&lsaquo; Cancel &amp; Go Back To Employee Range List</button>
					</a>
					<button type="submit" class="btn btn-primary"><i class="far fa-save"></i> Save</button>
				</div>
		    </div>
		</div>
		
	</form>
</div>

<!--#include file="../../inc/footer-main.asp"-->
