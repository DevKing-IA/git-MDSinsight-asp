<!--#include file="../../inc/header.asp"-->

<% InternalRecordIdentifier = Request.QueryString("i") 
If InternalRecordIdentifier = "" Then Response.Redirect("main.asp")
%>

<link rel="stylesheet" type="text/css" href="<%= BaseURL %>css/tabs.css" />



<SCRIPT LANGUAGE="JavaScript">
<!--
    function validateCompetitorForm()
    {

        if (document.frmAddCompetitor.txtCompetitorName.value == "") {
            swal("Competitor name cannot be blank.");
            return false;
        }

        if (document.frmAddCompetitor.txtCompetitorAddressInfo.value == "") {
            swal("Please enter address information for competitor.");
            return false;
        }
        
        var oneCheckboxSelected = false;
         
		var c = document.getElementsByTagName('input'); 
	    for (var i = 0; i < c.length; i++) { 
	    	if (c[i].type == 'checkbox' && c[i].checked == true) { 
			    // At least one checkbox IS checked     
			    oneCheckboxSelected = true;
	   		}     
		}
		
        if (oneCheckboxSelected == false) {
            swal("Please select at least one industry for this competitor.");
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


.checkboxrow {
    border: 0px;
    overflow: hidden;
    padding: 5px;
}

.checkboxcol {
 	border: 1px solid #ccc;
    border-radius: 4px;
    /*float: left;*/
    padding: 5px;
    margin-top:10px;
    margin-right: 5px;
}

	</style>
<!-- eof password strength meter !-->



<%
SQL = "SELECT * FROM PR_Competitors where InternalRecordIdentifier = " & InternalRecordIdentifier 

Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3 
Set rs = cnn8.Execute(SQL)
	
If not rs.EOF Then
	CompetitorName = rs("CompetitorName")
	CompetitorAddressInfo = rs("AddressInformation")
	BottledWater = rs("BottledWater")
	FilteredWater = rs("FilteredWater")
	OCS = rs("OCS")
	OCS_Supply = rs("OCS_Supply")
	OfficeSupplies = rs("OfficeSupplies")
	Vending = rs("Vending")
	Micromarket = rs("Micromarket")
	Pantry = rs("Pantry")	
End If
set rs = Nothing
cnn8.close
set cnn8 = Nothing

%>


<h1 class="page-header"> Edit Competitor</h1>

<div class="custom-container">

	<form method="POST" action="editCompetitor_submit.asp" name="frmeditCompetitor" id="frmeditCompetitor" onsubmit="return validateEditCompetitorForm();">
		
		<div class="row row-line">
		
			<input type="hidden" id="txtInternalRecordIdentifier" name="txtInternalRecordIdentifier" value="<%=InternalRecordIdentifier%>">

			<div class="form-group col-lg-12">
				<label for="txtCompetitorName" class="col-sm-3 control-label">Competitor Name</label>	
    			<div class="col-sm-9">
    				<input type="text" class="form-control required" id="txtCompetitorName" name="txtCompetitorName" value="<%= CompetitorName %>">
    			</div>
			</div>
			<div class="form-group col-lg-12">
				<label for="txtCompetitorAddressInfo" class="col-sm-3 control-label">Address Information</label>	
    			<div class="col-sm-9">
    				<textarea class="form-control textarea required" rows="4" id="txtCompetitorAddressInfo" name="txtCompetitorAddressInfo"><%= CompetitorAddressInfo %></textarea>
    				<strong>Enter full mailing address or city/state/zip.</strong>
    			</div>
 			</div>
 			
 			<div class="form-group col-lg-12">
 				
 				 	 <label for="checkboxes" class="col-sm-3 control-label">Select Industries (choose all that apply)</label>
 				 	 
 				 	 <div class="col-sm-9">
						 <div class="checkboxrow">
						    <div class="checkboxcol"><input type="checkbox" name="chkOfficeCoffee" value="Office Coffee" <% If OCS = True Then Response.write("checked") %>><label>&nbsp;Office Coffee</label></div>
						    <div class="checkboxcol"><input type="checkbox" name="chkOfficeCoffeeSupply" value="Office Coffee Supply" <% If OCS_Supply = True Then Response.write("checked") %>><label>&nbsp;Office Coffee Supply</label></div>
						    <div class="checkboxcol"><input type="checkbox" name="chkBottledWater" value="Bottled Water" <% If BottledWater = True Then Response.write("checked") %>><label>&nbsp;Bottled Water</label></div>
						    <div class="checkboxcol"><input type="checkbox" name="chkFilteredWater" value="Filtered Water" <% If FilteredWater = True Then Response.write("checked") %>><label>&nbsp;Filtered Water</label></div>
						    <div class="checkboxcol"><input type="checkbox" name="chkVending" value="Vending" <% If Vending = True Then Response.write("checked") %>><label>&nbsp;Vending</label></div>
						    <div class="checkboxcol"><input type="checkbox" name="chkMicromarket" value="Micromarket" <% If Micromarket = True Then Response.write("checked") %>><label>&nbsp;Micromarket</label></div>
						    <div class="checkboxcol"><input type="checkbox" name="chkPantry" value="Pantry" <% If Pantry = True Then Response.write("checked") %>><label>&nbsp;Pantry</label></div>
						    <div class="checkboxcol"><input type="checkbox" name="chkOfficeSupplies" value="Office Supplies" <% If OfficeSupplies = True Then Response.write("checked") %>><label>&nbsp;Office Supplies</label></div>
						</div>			
					</div>
 			</div>	
			
		</div>
		
		
	    <!-- cancel / submit !-->
		<div class="row row-line">
			<div class="col-lg-12 alertbutton">
				<div class="col-lg-12">
					<a href="<%= BaseURL %>filemaint/prospecting/competitors/main.asp">
	    				<button type="button" class="btn btn-default">&lsaquo; Cancel &amp; Go Back To Competitor List</button>
					</a>
					<button type="submit" class="btn btn-primary"><i class="far fa-save"></i> Save</button>
				</div>
		    </div>
		</div>
		
	</form>
</div>

<!--#include file="../../inc/footer-main.asp"-->
