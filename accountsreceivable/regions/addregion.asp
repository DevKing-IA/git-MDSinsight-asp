<!--#include file="../../inc/header.asp"-->

<link rel="stylesheet" href="<%= BaseURL %>js/bootstrap-multiselect/bootstrap-multiselect.css" type="text/css">
<script type="text/javascript" src="<%= BaseURL %>js/bootstrap-multiselect/bootstrap-multiselect.js"></script>


<link rel="stylesheet" type="text/css" href="<%= BaseURL %>css/tabs.css" />

<SCRIPT LANGUAGE="JavaScript">
<!--
    function validateregionForm()
    {

        if (document.frmAddregion.txtRegion.value == "") {
            swal("Region can not be blank.");
            return false;
        }
        /*else if (document.frmAddregion.txtCities.value == "") {
            swal("Cities can not be blank.");
            return false;
        }
        else if (document.frmAddregion.txtZipPostalCodes.value == "") {
            swal("Zip or Postal Codes can not be blank.");
            return false;
        }
        else if (document.frmAddregion.txtStatesProvinces.value == "") {
            swal("States or Provinces can not be blank.");
            return false;
        }*/		
		
        return true;

    }
// -->
</SCRIPT>   


	<script>
	
		$(document).ready(function() {	
			
			$('#lstExistingRegionList').multiselect({
			   buttonTitle: function(options, select) {
				    var selected = '';
				    options.each(function () {
				      selected += $(this).text() + ', ';
				    });
				    return selected.substr(0, selected.length - 2);
				  },
				buttonClass: 'btn btn-primary',
				buttonWidth: '425px',
				maxHeight: 400,
				dropRight:true,
				enableFiltering:true,
				filterPlaceholder:'Search',
				enableCaseInsensitiveFiltering:true,
				// possible options: 'text', 'value', 'both'
				filterBehavior:'text',
				includeFilterClearBtn:true,
				nonSelectedText:'No Regions Selected',
				numberDisplayed: 20,
			    onChange: function() {
			        var selected = this.$select.val();
			        $("#lstSelectedRegionList").val(selected);
			        console.log(selected);
			        // ...
			    }
	    			
		    });	
		    
			//*************************************************************************************************
			//Load the bootstrap multiselect box with the current threshhold report users preselected
			//*************************************************************************************************
			var data= $("#lstSelectedRegionList").val();
			//Make an array
			
			if (data) {
				var dataarray=data.split(",");
				// Set the value
				$("#lstExistingRegionList").val(dataarray);
				// Then refresh
				$("#lstExistingRegionList").multiselect("refresh");
			}
			//*************************************************************************************************
					
	        
		});
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
	</style>
<!-- eof password strength meter !-->

<h1 class="page-header"> Add New Region</h1>

<div class="custom-container">

	<form method="POST" action="addregion_submit.asp" name="frmAddregion" id="frmAddregion" onsubmit="return validateregionForm();">

		<div class="row row-line">

			<div class="form-group col-lg-12">
				<label for="txtRegion" class="col-sm-3 control-label">Region</label>	
    			<div class="col-sm-6">
    				<input type="text" class="form-control required" id="txtRegion" name="txtRegion" >
    			</div>
			</div>
			
			<div class="form-group col-lg-12">
				<label for="txtCities" class="col-sm-3 control-label">Cities</label>	
    			<div class="col-sm-6">    				
					<textarea class="form-control" id="txtCities" name="txtCities" rows="4"></textarea>
    			</div>
			</div>

			<div class="form-group col-lg-12">
				<label for="txtCities" class="col-sm-3 control-label">State or Province for Cities</label>	
    			<div class="col-sm-6">    				
					<select class="form-control" id="selStateOrProvince" name="selStateOrProvince">
						<!--#include file="..\customermininfo\statelist.asp"-->
					</select>
    			</div>
			</div>

			<div class="form-group col-lg-12">
				<label for="txtZipPostalCodes" class="col-sm-3 control-label">Zip Or Postal Codes</label>	
    			<div class="col-sm-6">    				
					<textarea class="form-control" id="txtZipPostalCodes" name="txtZipPostalCodes" rows="4"></textarea>
    			</div>
			</div>
			
			<div class="form-group col-lg-12">
				<label for="txtStatesProvinces" class="col-sm-3 control-label">States Or Provinces</label>	
    			<div class="col-sm-6">
    				<input type="text" class="form-control" id="txtStatesProvinces" name="txtStatesProvinces" >
    			</div>
			</div>

			<div class="form-group col-lg-12">
				<label for="chkUseRegionForServiceTickets" class="col-sm-3 control-label"></label>	
				<div class="col-sm-9">
					<input type="checkbox" id="chkUseRegionForServiceTickets" name="chkUseRegionForServiceTickets">&nbsp;&nbsp;Use Region For Service Tickets
				</div>
			</div>

			<div class="form-group col-lg-12">
				<label for="chkAutomaticFilter" class="col-sm-3 control-label"></label>	
    			<div class="col-sm-9">
    				<input type="checkbox" id="chkAutomaticFilter" name="chkAutomaticFilter">&nbsp;&nbsp;Include in Automatic Filter Ticket Generation
    			</div>
			</div>

			<div class="form-group col-lg-12">
				<label for="chkSuggestedFilter" class="col-sm-3 control-label"></label>	
    			<div class="col-sm-9">
    				<input type="checkbox" id="chkSuggestedFilter" name="chkSuggestedFilter">&nbsp;&nbsp;Include in Suggested Filter Ticket Generation
    			</div>
			</div>
			
			<div class="form-group col-lg-12">
				<label for="selAutoTicket" class="col-sm-3 control-label">Auto Filter Tickets (Max # of Tickets Per Day)</label>	
    			<div class="col-sm-6">
    				<select class="form-control" name="selAutoTicket" id="selAutoTicket">
						<!--<option value="">Max # of Tickets Per Day</option>-->
    					<option value="0">0</option>
						<option value="5">5</option>
    					<option value="10">10</option>
						<option value="15">15</option>
    					<option value="20">20</option>
						<option value="25">25</option>
    					<option value="30">30</option>
						<option value="35">35</option>
    					<option value="40">40</option>
						<option value="45">45</option>
    					<option value="50">50</option>
						<option value="55">55</option>
    					<option value="60">60</option>
						<option value="65">65</option>
    					<option value="70">70</option>
						<option value="75">75</option>
    					<option value="80">80</option>
						<option value="85">85</option>
    					<option value="90">90</option>
						<option value="95">95</option>
    					<option value="100">100</option>
						<option value="105">105</option>
    					<option value="110">110</option>
						<option value="115">115</option>
    					<option value="120">120</option>
						<option value="125">125</option>
    					<option value="130">130</option>
						<option value="135">135</option>
    					<option value="140">140</option>
						<option value="145">145</option>
    					<option value="150">150</option>
						<option value="155">155</option>
    					<option value="160">160</option>
						<option value="165">165</option>
    					<option value="170">170</option>
						<option value="175">175</option>
    					<option value="180">180</option>
						<option value="185">185</option>
    					<option value="190">190</option>
						<option value="195">195</option>
    					<option value="200">200</option>
    					<option value="205">205</option>
    					<option value="210">210</option>
    					<option value="215">215</option>
    					<option value="220">220</option>
    					<option value="225">225</option>
    					<option value="230">230</option>
    					<option value="235">235</option>
    					<option value="240">240</option>
    					<option value="245">245</option>
    					<option value="250">250</option>						
    				</select>
					<span><strong>Note: ZERO DENOTES NO LIMIT</strong></span>
    			</div>
			</div>			

			<div class="form-group col-lg-12">
				<label for="selSuggestedTicket" class="col-sm-3 control-label">Suggested Filter Tickets (Max # of Tickets Per Day)</label>	
    			<div class="col-sm-6">
    				<select class="form-control" name="selSuggestedTicket" id="selSuggestedTicket">
						<!--<option value="">Max # of Tickets Per Day</option>-->
    					<option value="0">0</option>
						<option value="5">5</option>
    					<option value="10">10</option>
						<option value="15">15</option>
    					<option value="20">20</option>
						<option value="25">25</option>
    					<option value="30">30</option>
						<option value="35">35</option>
    					<option value="40">40</option>
						<option value="45">45</option>
    					<option value="50">50</option>
						<option value="55">55</option>
    					<option value="60">60</option>
						<option value="65">65</option>
    					<option value="70">70</option>
						<option value="75">75</option>
    					<option value="80">80</option>
						<option value="85">85</option>
    					<option value="90">90</option>
						<option value="95">95</option>
    					<option value="100">100</option>
						<option value="105">105</option>
    					<option value="110">110</option>
						<option value="115">115</option>
    					<option value="120">120</option>
						<option value="125">125</option>
    					<option value="130">130</option>
						<option value="135">135</option>
    					<option value="140">140</option>
						<option value="145">145</option>
    					<option value="150">150</option>
						<option value="155">155</option>
    					<option value="160">160</option>
						<option value="165">165</option>
    					<option value="170">170</option>
						<option value="175">175</option>
    					<option value="180">180</option>
						<option value="185">185</option>
    					<option value="190">190</option>
						<option value="195">195</option>
    					<option value="200">200</option>
    					<option value="205">205</option>
    					<option value="210">210</option>
    					<option value="215">215</option>
    					<option value="220">220</option>
    					<option value="225">225</option>
    					<option value="230">230</option>
    					<option value="235">235</option>
    					<option value="240">240</option>
    					<option value="245">245</option>
    					<option value="250">250</option>						
    				</select>
					<span><strong>Note: ZERO DENOTES NO LIMIT</strong></span>
    			</div>
			</div>			
		</div>

		<div class="row row-line">

			<div class="form-group col-lg-12">
				<% RegionList = "" %>	
    			<div class="col-sm-12">
					<p>Include All Addresses <i class="fas fa-map-marked-alt"></i> Not Covered By The Following Regions:</p>
					<input type="hidden" name="lstSelectedRegionList" id="lstSelectedRegionList" value="<%= RegionList %>">
					<select id="lstExistingRegionList" multiple="multiple" name="lstExistingRegionList">
						<%	
							
						Set cnnRegionList = Server.CreateObject("ADODB.Connection")
						cnnRegionList.open Session("ClientCnnString")
		
						SQLRegionList = "SELECT * FROM AR_Regions ORDER BY InternalRecordIdentifier"
						
						Set rsRegionList = Server.CreateObject("ADODB.Recordset")
						rsRegionList.CursorLocation = 3 
						Set rsRegionList = cnnRegionList.Execute(SQLRegionList)
						
						If Not rsRegionList.EOF Then
							Do While Not rsRegionList.EOF
							
								RegionName = rsRegionList("Region")
								Response.Write("<option value='" & rsRegionList("InternalRecordIdentifier") & "'>" & RegionName & "</option>")
						
								rsRegionList.MoveNext
							Loop
						End If
			
						Set rsRegionList = Nothing
						cnnRegionList.Close
						Set cnnRegionList = Nothing
							
						%>
					</select>				
    			</div>
			</div>
			
		</div>

		
	    <!-- cancel / submit !-->
		<div class="row row-line">
			<div class="col-lg-12 alertbutton">
				<div class="col-lg-12">
					<a href="<%= BaseURL %>accountsreceivable/regions/main.asp">
	    				<button type="button" class="btn btn-default">&lsaquo; Cancel &amp; Go Back To Region List</button>
					</a>
					<button type="submit" class="btn btn-primary"><i class="far fa-save"></i> Save</button>
				</div>
		    </div>
		</div>
		
	</form>
</div>

<!--#include file="../../inc/footer-main.asp"-->
