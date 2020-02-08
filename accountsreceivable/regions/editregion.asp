<!--#include file="../../inc/header.asp"-->

<% InternalRecordIdentifier = Request.QueryString("i") 
If InternalRecordIdentifier = "" Then Response.Redirect("main.asp")
%>

<link rel="stylesheet" href="<%= BaseURL %>js/bootstrap-multiselect/bootstrap-multiselect.css" type="text/css">
<script type="text/javascript" src="<%= BaseURL %>js/bootstrap-multiselect/bootstrap-multiselect.js"></script>
<link rel="stylesheet" type="text/css" href="<%= BaseURL %>css/tabs.css" />


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


<SCRIPT LANGUAGE="JavaScript">
<!--
    function validateEditregionForm()
    {

        if (document.frmeditregion.txtRegion.value == "") {
            swal("Region can not be blank.");
            return false;
        }
        /*else if (document.frmeditregion.txtCities.value == "") {
            swal("Cities can not be blank.");
            return false;
        }
        else if (document.frmeditregion.txtZipPostalCodes.value == "") {
            swal("Zip or Postal Codes can not be blank.");
            return false;
        }
        else if (document.frmeditregion.txtStatesProvinces.value == "") {
            swal("States or Provinces can not be blank.");
            return false;
        }*/

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
	max-width:850px;
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
SQL = "SELECT * FROM AR_Regions where InternalRecordIdentifier = " & InternalRecordIdentifier 

Set cnnregions = Server.CreateObject("ADODB.Connection")
cnnregions.open (Session("ClientCnnString"))
Set rsregions = Server.CreateObject("ADODB.Recordset")
rsregions.CursorLocation = 3 
Set rsregions = cnnregions.Execute(SQL)
	
If not rsregions.EOF Then

	State =	rsregions("StateForCities")
	
	Region = rsregions("Region")
	Cities1 = rsregions("Cities1")
	Cities2 = rsregions("Cities2")
	Cities3 = rsregions("Cities3")
	'Cities = Cities1 & "," & Cities2 & "," & Cities3
	
	Cities = Cities1
	
	If Cities2 <> "" Then
		Cities = Cities & "," & Cities2
	End If
	
	If Cities3 <> "" Then
		Cities = Cities & "," & Cities3
	End If
							
	StatesOrProvinces = rsregions("StatesOrProvinces")
	ZipOrPostalCodes1 = rsregions("ZipOrPostalCodes1")
	ZipOrPostalCodes2 = rsregions("ZipOrPostalCodes2")
	
	IncludeInAutoFilterTickets = rsregions("IncludeInAutoFilterTickets")
	IncludeInSuggestedFilterTickets = rsregions("IncludeInSuggestedFilterTickets")
	AutoTickets = rsregions("AutoFilterChangeMaxNumTicketsPerDay")
	SuggestedTickets = rsregions("SuggestedFilterChangeMaxNumTicketsPerDay")

	ZipOrPostalCodes = ZipOrPostalCodes1 & "," & ZipOrPostalCodes2
	
	ZipOrPostalCodes = ZipOrPostalCodes1
	
	If ZipOrPostalCodes2 <> "" Then
		ZipOrPostalCodes = ZipOrPostalCodes & "," & ZipOrPostalCodes2
	End If	
	
	CatchAllRegionIntRecIDs = rsregions("CatchAllRegionIntRecIDs")
	
	If CatchAllRegionIntRecIDs <> "" Then
		CatchAllRegionIntRecIDs = Replace(CatchAllRegionIntRecIDs, " ", "")
	End If		
	
	UseRegionForServiceTickets = rsregions("UseRegionForServiceTickets")
	
End If
set rsregions = Nothing
cnnregions.close
set cnnregions = Nothing

%>

<%
	If IncludeInAutoFilterTickets = 0 Then
		AutoFilterTickets = ""
	Else
		AutoFilterTickets = "checked"
	End If

	If IncludeInSuggestedFilterTickets = 0 Then
		SuggestedFilterTickets = ""
	Else
		SuggestedFilterTickets = "checked"
	End If	
	
	If UseRegionForServiceTickets = 0 Then
		UseForServiceTickets = ""
	Else
		UseForServiceTickets = "checked"
	End If	

%>

<h1 class="page-header"> Edit Region</h1>

<div class="custom-container">

	<form method="POST" action="editregion_submit.asp" name="frmeditregion" id="frmeditregion" onsubmit="return validateEditregionForm();">

		<div class="row row-line">
		
			<input type="hidden" id="txtInternalRecordIdentifier" name="txtInternalRecordIdentifier" value="<%=InternalRecordIdentifier%>">
			
			<div class="form-group col-lg-12">
				<label for="txtRegion" class="col-sm-3 control-label">Region</label>	
    			<div class="col-sm-6">
    				<input type="text" class="form-control required" id="txtRegion" name="txtRegion" value="<%=Region%>">
    			</div>
			</div>
		</div>	

		<div class="row row-line">		
			<div class="form-group col-lg-12">
				<label for="txtCities" class="col-sm-3 control-label">Cities</label>	
    			<div class="col-sm-6">
    				<textarea class="form-control" id="txtCities" name="txtCities" rows="4"><%=Cities%></textarea>
    			</div>
			</div>
		</div>	
		
		<div class="row row-line">		
			<div class="form-group col-lg-12">
				<label for="txtCities" class="col-sm-3 control-label">State or Province for Cities</label>	
    			<div class="col-sm-6">
					<select class="form-control" id="selStateOrProvince" name="selStateOrProvince">
						<!--#include file="..\customermininfo\statelist.asp"-->
					</select>
    			</div>
			</div>
		</div>	
		
		<div class="row row-line">		
			<div class="form-group col-lg-12">
				<label for="txtZipPostalCodes" class="col-sm-3 control-label">Zip Or Postal Codes</label>	
    			<div class="col-sm-6">    				
					<textarea class="form-control" id="txtZipPostalCodes" name="txtZipPostalCodes" rows="4"><%=ZipOrPostalCodes%></textarea>
    			</div>
			</div>
		</div>

		<div class="row row-line">
			<div class="form-group col-lg-12">
				<label for="txtStatesProvinces" class="col-sm-3 control-label">States Or Provinces</label>	
				<div class="col-sm-6">
					<input type="text" class="form-control" id="txtStatesProvinces" name="txtStatesProvinces" value="<%=StatesOrProvinces%>">
				</div>
			</div>	
		</div>	

		<div class="row row-line">
			<div class="form-group col-lg-12">
				<label for="chkSuggestedFilter" class="col-sm-3 control-label"></label>	
    			<div class="col-sm-9">
    				<input type="checkbox" id="chkUseRegionForServiceTickets" name="chkUseRegionForServiceTickets" <%=UseForServiceTickets %>>&nbsp;&nbsp;Use Region For Service Tickets
    			</div>
			</div>
		</div>
		
		<div class="row row-line">
			<div class="form-group col-lg-12">
				<label for="chkAutomaticFilter" class="col-sm-3 control-label"></label>	
    			<div class="col-sm-9">
    				<input type="checkbox" id="chkAutomaticFilter" name="chkAutomaticFilter" <%=AutoFilterTickets%>>&nbsp;&nbsp;Include in Automatic Filter Ticket Generation
    			</div>
			</div>
		</div>

		<div class="row row-line">
			<div class="form-group col-lg-12">
				<label for="chkSuggestedFilter" class="col-sm-3 control-label"></label>	
    			<div class="col-sm-9">
    				<input type="checkbox" id="chkSuggestedFilter" name="chkSuggestedFilter" <%=SuggestedFilterTickets%>>&nbsp;&nbsp;Include in Suggested Filter Ticket Generation
    			</div>
			</div>
		</div>
			
		<div class="row row-line">
			<div class="form-group col-lg-12">
				<label for="selAutoTicket" class="col-sm-3 control-label">Auto Filter Tickets (Max # of Tickets Per Day)</label>	
    			<div class="col-sm-6">
    				<select class="form-control" name="selAutoTicket" id="selAutoTicket">
						<!--<option value="">Max # of Tickets Per Day</option>-->
    					<option value="0" <% If AutoTickets = "0" Then Response.Write(" selected")%>>0</option>
						<option value="5" <% If AutoTickets = "5" Then Response.Write(" selected")%>>5</option>
    					<option value="10" <% If AutoTickets = "10" Then Response.Write(" selected")%>>10</option>
						<option value="15" <% If AutoTickets = "15" Then Response.Write(" selected")%>>15</option>
    					<option value="20" <% If AutoTickets = "20" Then Response.Write(" selected")%>>20</option>
						<option value="25" <% If AutoTickets = "25" Then Response.Write(" selected")%>>25</option>
    					<option value="30" <% If AutoTickets = "30" Then Response.Write(" selected")%>>30</option>
						<option value="35" <% If AutoTickets = "35" Then Response.Write(" selected")%>>35</option>
    					<option value="40" <% If AutoTickets = "40" Then Response.Write(" selected")%>>40</option>
						<option value="45" <% If AutoTickets = "45" Then Response.Write(" selected")%>>45</option>
    					<option value="50" <% If AutoTickets = "50" Then Response.Write(" selected")%>>50</option>
						<option value="55" <% If AutoTickets = "55" Then Response.Write(" selected")%>>55</option>
    					<option value="60" <% If AutoTickets = "60" Then Response.Write(" selected")%>>60</option>
						<option value="65" <% If AutoTickets = "65" Then Response.Write(" selected")%>>65</option>
    					<option value="70" <% If AutoTickets = "70" Then Response.Write(" selected")%>>70</option>
						<option value="75" <% If AutoTickets = "75" Then Response.Write(" selected")%>>75</option>
    					<option value="80" <% If AutoTickets = "80" Then Response.Write(" selected")%>>80</option>
						<option value="85" <% If AutoTickets = "85" Then Response.Write(" selected")%>>85</option>
    					<option value="90" <% If AutoTickets = "90" Then Response.Write(" selected")%>>90</option>
						<option value="95" <% If AutoTickets = "95" Then Response.Write(" selected")%>>95</option>
    					<option value="100" <% If AutoTickets = "100" Then Response.Write(" selected")%>>100</option>
						<option value="105" <% If AutoTickets = "105" Then Response.Write(" selected")%>>105</option>
    					<option value="110" <% If AutoTickets = "110" Then Response.Write(" selected")%>>110</option>
						<option value="115" <% If AutoTickets = "115" Then Response.Write(" selected")%>>115</option>
    					<option value="120" <% If AutoTickets = "120" Then Response.Write(" selected")%>>120</option>
						<option value="125" <% If AutoTickets = "125" Then Response.Write(" selected")%>>125</option>
    					<option value="130" <% If AutoTickets = "130" Then Response.Write(" selected")%>>130</option>
						<option value="135" <% If AutoTickets = "135" Then Response.Write(" selected")%>>135</option>
    					<option value="140" <% If AutoTickets = "140" Then Response.Write(" selected")%>>140</option>
						<option value="145" <% If AutoTickets = "145" Then Response.Write(" selected")%>>145</option>
    					<option value="150" <% If AutoTickets = "150" Then Response.Write(" selected")%>>150</option>
						<option value="155" <% If AutoTickets = "155" Then Response.Write(" selected")%>>155</option>
    					<option value="160" <% If AutoTickets = "160" Then Response.Write(" selected")%>>160</option>
						<option value="165" <% If AutoTickets = "165" Then Response.Write(" selected")%>>165</option>
    					<option value="170" <% If AutoTickets = "170" Then Response.Write(" selected")%>>170</option>
						<option value="175" <% If AutoTickets = "175" Then Response.Write(" selected")%>>175</option>
    					<option value="180" <% If AutoTickets = "180" Then Response.Write(" selected")%>>180</option>
						<option value="185" <% If AutoTickets = "185" Then Response.Write(" selected")%>>185</option>
    					<option value="190" <% If AutoTickets = "190" Then Response.Write(" selected")%>>190</option>
						<option value="195" <% If AutoTickets = "195" Then Response.Write(" selected")%>>195</option>
    					<option value="200" <% If AutoTickets = "200" Then Response.Write(" selected")%>>200</option>
    					<option value="205" <% If AutoTickets = "205" Then Response.Write(" selected")%>>205</option>
    					<option value="210" <% If AutoTickets = "210" Then Response.Write(" selected")%>>210</option>
    					<option value="215" <% If AutoTickets = "215" Then Response.Write(" selected")%>>215</option>
    					<option value="220" <% If AutoTickets = "220" Then Response.Write(" selected")%>>220</option>
    					<option value="225" <% If AutoTickets = "225" Then Response.Write(" selected")%>>225</option>
    					<option value="230" <% If AutoTickets = "230" Then Response.Write(" selected")%>>230</option>
    					<option value="235" <% If AutoTickets = "235" Then Response.Write(" selected")%>>235</option>
    					<option value="240" <% If AutoTickets = "240" Then Response.Write(" selected")%>>240</option>
    					<option value="245" <% If AutoTickets = "245" Then Response.Write(" selected")%>>245</option>
    					<option value="250" <% If AutoTickets = "250" Then Response.Write(" selected")%>>250</option>
    				</select>
					<span><strong>Note: ZERO DENOTES NO LIMIT</strong></span>
    			</div>
			</div>	
		</div>		

		<div class="row row-line">
			<div class="form-group col-lg-12">
				<label for="selSuggestedTicket" class="col-sm-3 control-label">Suggested Filter Tickets (Max # of Tickets Per Day)</label>	
    			<div class="col-sm-6">
    				<select class="form-control" name="selSuggestedTicket" id="selSuggestedTicket">
						<!--<option value="">Max # of Tickets Per Day</option>-->
    					<option value="0" <% If SuggestedTickets = "0" Then Response.Write(" selected")%>>0</option>
						<option value="5" <% If SuggestedTickets = "5" Then Response.Write(" selected")%>>5</option>
    					<option value="10" <% If SuggestedTickets = "10" Then Response.Write(" selected")%>>10</option>
						<option value="15" <% If SuggestedTickets = "15" Then Response.Write(" selected")%>>15</option>
    					<option value="20" <% If SuggestedTickets = "20" Then Response.Write(" selected")%>>20</option>
						<option value="25" <% If SuggestedTickets = "25" Then Response.Write(" selected")%>>25</option>
    					<option value="30" <% If SuggestedTickets = "30" Then Response.Write(" selected")%>>30</option>
						<option value="35" <% If SuggestedTickets = "35" Then Response.Write(" selected")%>>35</option>
    					<option value="40" <% If SuggestedTickets = "40" Then Response.Write(" selected")%>>40</option>
						<option value="45" <% If SuggestedTickets = "45" Then Response.Write(" selected")%>>45</option>
    					<option value="50" <% If SuggestedTickets = "50" Then Response.Write(" selected")%>>50</option>
						<option value="55" <% If SuggestedTickets = "55" Then Response.Write(" selected")%>>55</option>
    					<option value="60" <% If SuggestedTickets = "60" Then Response.Write(" selected")%>>60</option>
						<option value="65" <% If SuggestedTickets = "65" Then Response.Write(" selected")%>>65</option>
    					<option value="70" <% If SuggestedTickets = "70" Then Response.Write(" selected")%>>70</option>
						<option value="75" <% If SuggestedTickets = "75" Then Response.Write(" selected")%>>75</option>
    					<option value="80" <% If SuggestedTickets = "80" Then Response.Write(" selected")%>>80</option>
						<option value="85" <% If SuggestedTickets = "85" Then Response.Write(" selected")%>>85</option>
    					<option value="90" <% If SuggestedTickets = "90" Then Response.Write(" selected")%>>90</option>
						<option value="95" <% If SuggestedTickets = "95" Then Response.Write(" selected")%>>95</option>
    					<option value="100" <% If SuggestedTickets = "100" Then Response.Write(" selected")%>>100</option>
						<option value="105" <% If SuggestedTickets = "105" Then Response.Write(" selected")%>>105</option>
    					<option value="110" <% If SuggestedTickets = "110" Then Response.Write(" selected")%>>110</option>
						<option value="115" <% If SuggestedTickets = "115" Then Response.Write(" selected")%>>115</option>
    					<option value="120" <% If SuggestedTickets = "120" Then Response.Write(" selected")%>>120</option>
						<option value="125" <% If SuggestedTickets = "125" Then Response.Write(" selected")%>>125</option>
    					<option value="130" <% If SuggestedTickets = "130" Then Response.Write(" selected")%>>130</option>
						<option value="135" <% If SuggestedTickets = "135" Then Response.Write(" selected")%>>135</option>
    					<option value="140" <% If SuggestedTickets = "140" Then Response.Write(" selected")%>>140</option>
						<option value="145" <% If SuggestedTickets = "145" Then Response.Write(" selected")%>>145</option>
    					<option value="150" <% If SuggestedTickets = "150" Then Response.Write(" selected")%>>150</option>
						<option value="155" <% If SuggestedTickets = "155" Then Response.Write(" selected")%>>155</option>
    					<option value="160" <% If SuggestedTickets = "160" Then Response.Write(" selected")%>>160</option>
						<option value="165" <% If SuggestedTickets = "165" Then Response.Write(" selected")%>>165</option>
    					<option value="170" <% If SuggestedTickets = "170" Then Response.Write(" selected")%>>170</option>
						<option value="175" <% If SuggestedTickets = "175" Then Response.Write(" selected")%>>175</option>
    					<option value="180" <% If SuggestedTickets = "180" Then Response.Write(" selected")%>>180</option>
						<option value="185" <% If SuggestedTickets = "185" Then Response.Write(" selected")%>>185</option>
    					<option value="190" <% If SuggestedTickets = "190" Then Response.Write(" selected")%>>190</option>
						<option value="195" <% If SuggestedTickets = "195" Then Response.Write(" selected")%>>195</option>
    					<option value="200" <% If SuggestedTickets = "200" Then Response.Write(" selected")%>>200</option>
    					<option value="205" <% If SuggestedTickets = "205" Then Response.Write(" selected")%>>205</option>
    					<option value="210" <% If SuggestedTickets = "210" Then Response.Write(" selected")%>>210</option>
    					<option value="215" <% If SuggestedTickets = "215" Then Response.Write(" selected")%>>215</option>
    					<option value="220" <% If SuggestedTickets = "220" Then Response.Write(" selected")%>>220</option>
    					<option value="225" <% If SuggestedTickets = "225" Then Response.Write(" selected")%>>225</option>
    					<option value="230" <% If SuggestedTickets = "230" Then Response.Write(" selected")%>>230</option>
    					<option value="235" <% If SuggestedTickets = "235" Then Response.Write(" selected")%>>235</option>
    					<option value="240" <% If SuggestedTickets = "240" Then Response.Write(" selected")%>>240</option>
    					<option value="245" <% If SuggestedTickets = "245" Then Response.Write(" selected")%>>245</option>
    					<option value="250" <% If SuggestedTickets = "250" Then Response.Write(" selected")%>>250</option>
    				</select>
					<span><strong>Note: ZERO DENOTES NO LIMIT</strong></span>
    			</div>
			</div>
		</div>
			
		<div class="row row-line">
			<div class="form-group col-lg-12">
    			<label for="selRegions" class="col-sm-3 control-label">Include All Addresses <i class="fas fa-map-marked-alt"></i> Not Covered By The Following Regions:</label>
				<div class="col-sm-6">
					<input type="hidden" name="lstSelectedRegionList" id="lstSelectedRegionList" value="<%= CatchAllRegionIntRecIDs %>">
					<select id="lstExistingRegionList" multiple="multiple" name="lstExistingRegionList">
						<%	
							
						Set cnnRegionList = Server.CreateObject("ADODB.Connection")
						cnnRegionList.open Session("ClientCnnString")
		
						SQLRegionList = "SELECT * FROM AR_Regions WHERE InternalRecordIdentifier <> " & InternalRecordIdentifier & " ORDER BY InternalRecordIdentifier"
						
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
