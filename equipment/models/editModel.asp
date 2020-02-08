<!--#include file="../../inc/header.asp"-->
<!--#include file="../../inc/InSightFuncs_Equipment.asp"-->

<!--------------------------------------------------------------------------->
<!-- THESE FILES ARE REQUIRED FOR THE UPLOADING OF DOCUMENTS, IMAGES, ETC. -->
<!--------------------------------------------------------------------------->


<link rel="stylesheet" href="https://ajax.googleapis.com/ajax/libs/jqueryui/1.12.1/themes/smoothness/jquery-ui.css">
<script src="https://ajax.googleapis.com/ajax/libs/jqueryui/1.12.1/jquery-ui.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/jquery.iframe-transport/1.0.1/jquery.iframe-transport.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/blueimp-file-upload/9.21.0/js/jquery.fileupload.js"></script>

<!-- The jQuery UI widget factory, can be omitted if jQuery UI is already included -->

<!--
<script src="<%= BaseURL %>js/fileupload/vendor/jquery.ui.widget.js"></script>
<!-- The Iframe Transport is required for browsers without support for XHR file uploads -->

<!--
<script src="<%= BaseURL %>js/fileupload/jquery.iframe-transport.js"></script>
<!-- The basic File Upload plugin -->

<!--
<script src="<%= BaseURL %>js/fileupload/jquery.fileupload.js"></script>
<!-- The File Upload processing plugin -->

<!--
<script src="<%= BaseURL %>js/fileupload/jquery.fileupload-process.js"></script>
<!-- The File Upload image preview & resize plugin -->

<!--
<script src="<%= BaseURL %>js/fileupload/jquery.fileupload-image.js"></script>
<!-- The File Upload validation plugin -->

<!--
<script src="<%= BaseURL %>js/fileupload/jquery.fileupload-validate.js"></script>
<!-- The File Upload user interface plugin -->

<!--
<script src="<%= BaseURL %>js/fileupload/jquery.fileupload-ui.js"></script>

-->

<!--------------------------------------------------------------------------->
<!--------------------------------------------------------------------------->


<% 

InternalRecordIdentifier = Request.QueryString("i") 
ModelIntRecID = Request.QueryString("i")
If InternalRecordIdentifier = "" Then Response.Redirect("main.asp")


SQL = "SELECT * FROM EQ_Models where InternalRecordIdentifier = " & InternalRecordIdentifier 

Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3 
Set rs = cnn8.Execute(SQL)
	
If not rs.EOF Then
	Model = rs("Model")
	BrandIntRecID = rs("BrandIntRecID")
	GroupIntRecID = rs("GroupIntRecID")
	ClassIntRecID = rs("ClassIntRecID")
	ManufacIntRecID = rs("ManufacIntRecID")
	DefaultRentalPrice = rs("DefaultRentalPrice")
	DefaultCostPrice = rs("DefaultCostPrice")
	ReplacementCost = rs("ReplacementCost")
	BackendSystemCode = rs("BackendSystemCode")	
	InsightAssetTagPrefix = rs("InsightAssetTagPrefix")
End If
set rs = Nothing
cnn8.close
set cnn8 = Nothing

%>


<link rel="stylesheet" type="text/css" href="<%= BaseURL %>css/tabs.css" />

<SCRIPT LANGUAGE="JavaScript">
<!--
	
	jQuery(document).ready(function($) {
	 
        $('#myCarousel').carousel({
                interval: 50000
        });
	 
        //Handles the carousel thumbnails
        $('[id^=carousel-selector-]').click(function () {
	        var id_selector = $(this).attr("id");
	        try {
	            var id = /-(\d+)$/.exec(id_selector)[1];
	            console.log(id_selector, id);
	            jQuery('#myCarousel').carousel(parseInt(id));
	        } catch (e) {
	            console.log('Regex failed!', e);
	        }
	    });
	    
        // When the carousel slides, auto update the text
        $('#myCarousel').on('slid.bs.carousel', function (e) {
                 var id = $('.item.active').data('slide-number');
                $('#carousel-text').html($('#slide-content-'+id).html());
        });
        
        
		function sliceSize(dataNum, dataTotal) {
		  return (dataNum / dataTotal) * 360;
		}
		
		function addSlice(sliceSize, pieElement, offset, sliceID, color) {
		  $(pieElement).append("<div class='slice "+sliceID+"'><span></span></div>");
		  var offset = offset - 1;
		  var sizeRotation = -179 + sliceSize;
		  $("."+sliceID).css({
		    "transform": "rotate("+offset+"deg) translate3d(0,0,0)"
		  });
		  $("."+sliceID+" span").css({
		    "transform"       : "rotate("+sizeRotation+"deg) translate3d(0,0,0)",
		    "background-color": color
		  });
		}
		
		function iterateSlices(sliceSize, pieElement, offset, dataCount, sliceCount, color) {
		  var sliceID = "s"+dataCount+"-"+sliceCount;
		  var maxSize = 179;
		  if(sliceSize<=maxSize) {
		    addSlice(sliceSize, pieElement, offset, sliceID, color);
		  } else {
		    addSlice(maxSize, pieElement, offset, sliceID, color);
		    iterateSlices(sliceSize-maxSize, pieElement, offset+maxSize, dataCount, sliceCount+1, color);
		  }
		}
		
		function createPie(dataElement, pieElement) {
		  var listData = [];
		  $(dataElement+" span").each(function() {
		    listData.push(Number($(this).html()));
		  });
		  var listTotal = 0;
		  for(var i=0; i<listData.length; i++) {
		    listTotal += listData[i];
		  }
		  var offset = 0;
		  var color = [
		    "cornflowerblue", 
		    "olivedrab", 
		    "orange", 
		    "tomato", 
		    "crimson", 
		    "purple", 
		    "turquoise", 
		    "forestgreen", 
		    "navy", 
		    "gray"
		  ];
		  for(var i=0; i<listData.length; i++) {
		    var size = sliceSize(listData[i], listTotal);
		    iterateSlices(size, pieElement, offset, i, 0, color[i]);
		    $(dataElement+" li:nth-child("+(i+1)+")").css("border-color", color[i]);
		    offset += size;
		  }
		}
		
		createPie(".pieID.legend", ".pieID.pie");
        
	});


    function validateEditModelForm()
    {

        if (document.frmEditModel.txtModel.value == "") {
            swal("Model can not be blank.");
            return false;
        }

		var ddlBrand = document.getElementById("selBrandIntRecID");
		var selectedValueBrand = ddlBrand.options[ddlBrand.selectedIndex].value;
		
		if (selectedValueBrand == "")
		{
			swal("Brand must be selected for this model.");
			return false;
		}

		var ddlGroup = document.getElementById("selGroupIntRecID");
		var selectedValueGroup = ddlGroup.options[ddlGroup.selectedIndex].value;
		
		if (selectedValueGroup == "")
		{
			swal("Group must be selected for this model.");
			return false;
		}

		var ddlClass = document.getElementById("selClassIntRecID");
		var selectedValueClass= ddlClass.options[ddlClass.selectedIndex].value;
		
		if (selectedValueClass == "")
		{
			swal("Class must be selected for this model.");
			return false;
		}
		
        if (document.frmEditModel.txtDefaultRentalPrice.value != "") {
        
        	if (isNaN(document.frmEditModel.txtDefaultRentalPrice.value)) {
            	swal("Please enter numbers only for the default rental price.");
            	return false;
           	}
        }
        if (document.frmEditModel.txtDefaultCost.value != "") {
        
        	if (isNaN(document.frmEditModel.txtDefaultCost.value)) {
            	swal("Please enter numbers only for the default cost.");
            	return false;
           	}
        }
        if (document.frmEditModel.txtReplacementCost.value != "") {
        
        	if (isNaN(document.frmEditModel.txtReplacementCost.value)) {
            	swal("Please enter numbers only for the replacement cost.");
            	return false;
           	}
        }
        if (document.frmEditModel.txtInsightAssetTagPrefix.value == "") {
            swal("The Insight asset tag prefix cannot be blank.");
            return false;
        }
        return true;
        

    }
    
    
		function ajaxRowMode(type, id, mode) {
		
			$('#ajaxRow'+type+'-'+id).attr("class", "ajaxRow"+mode);
			if(id==0){
				$('#ajaxRow'+type+'-' + 0 + '').remove();
			}	
		
			 $(".ajaxRowEdit").find('input[disabled="true"]').each(function () {
			     $(this).removeAttr("disabled");
		});	 
		 
		
	   var phonesTab = [{ "mask": "(###) ###-####"}];
	    $('#txtPhoneTab' + id).inputmask({ 
	        mask: phonesTab, 
	        greedy: false, 
	        definitions: { '#': { validator: "[0-9]", cardinality: 1}} });
	    $('#txtCellTab' + id).inputmask({ 
	        mask: phonesTab, 
	        greedy: false, 
	        definitions: { '#': { validator: "[0-9]", cardinality: 1}} });
	    $('#txtFaxTab' + id).inputmask({ 
	        mask: phonesTab, 
	        greedy: false, 
	        definitions: { '#': { validator: "[0-9]", cardinality: 1}} });
	
			
	}
    
// -->
</SCRIPT>   


<!-- password strength meter !-->

<style type="text/css">
	
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
	.ajaxRowView .visibleRowEdit, .ajaxRowEdit .visibleRowView { display: none; }
	
	
	.styled{
	 	cursor:pointer;
	}

	.plus-button{
		cursor:pointer;
	}
	
	.beatpicker-clear{
		display: none;
	}
	.nav-tabs{
		font-size: 12px;
	}
	
	.the-tabs .nav>li>a{
		padding: 5px 10px;
		font-weight: bold;
	}


	.tab-content{
		margin-top:20px;
		font-size:12px;
	}
	
	   
	.tab-content .split-arrows{
		 text-align:left;
		 margin-top:10px;
		 margin-bottom: 10px;
	 }
	 
	.tab-content .split-arrows a{
		display:inline-block;
		background:#f5f5f5;
		padding:5px;
	}
	
	.tab-content .split-arrows a:hover{
		background:#ccc;
		text-decoration:none;
	}

	.inside {
		position:absolute;
		text-indent:8px;
		margin-top:7px;
		color:green;
		font-size:20px;
	}
	
	.inp {
		text-indent:15px;
	}
	.select-line{
		margin-bottom: 15px;
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

	
	.form-control{
		min-width: 100px;
	}
	
	.textarea-box{
		min-width: 260px;
	}
	
	.container {
	    width: 100%;
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
	
	.bottom-tabs-section {
	    border: 1px solid #ccc;
	    padding: 10px;
	    margin-top: 20px;
	    float: left;
	    width: 100%;
	}
		
	.btn-custom{
		width:100%;
		text-align:left;
		color: #333;
	    background-color: #f5f5f5;
		border:1px solid #ddd;
		outline:none;
		border-top-left-radius:5px;
		border-top-right-radius:5px;
		font-size:16px;
		padding:10px;
	}
	
	.btn-custom:hover{
		background:#ccc;
	}
	
	
	.bottom-table table thead th{
		padding:6px;
		font-weight:bold;
		border:1px solid #ddd;
		vertical-align:top;
	}
	
	.bottom-table table>tbody>tr>td{
		padding:6px;
		font-weight:normal;
		border:1px solid #ddd;
		vertical-align: middle;
	}
	
	.narrow-results{
		margin-bottom:15px;
	}
	
	#filter-documents{
		width:40%;
		padding:10px;
		height:34px;
	}
	 
	#filter-images{
		width:40%;
		padding:10px;
		height:34px;
	}
	 
	#filter-links{
		width:40%;
		padding:10px;
		height:34px;
	}
	 
	#filter-companions{
		width:40%;
		padding:10px;
		height:34px;
	}
		
	.nav-tabs>li>a {
	color: #fff;
	font-size:16px;
	}
	
	
	.EquipmentTabImageColor{
		background:#cc4125 !important;
	}
	.EquipmentTabLinkColor{
		background:#ff9900 !important;
	}
	.EquipmentTabCompanionColor{
		background:#2ecc71 !important;
	}
	.EquipmentTabDocumentColor{
		background:#3d85c6 !important;
	}
		
	.fileicon {
		width:40%;
	}
	
	.nav-tabs > li.active > a,
	.nav-tabs > li.active > a:hover,
	.nav-tabs > li.active > a:focus{
	
		color: #fff;
		font-weight:normal;
		font-size:24px;
	    /*background-color: #111 !important;*/
	    border-color: #2e6da4 !important;
	    margin-bottom:20px;
	    margin-top:0px;
	     
	} 
			
	.hide-bullets {
	    list-style:none;
	    margin-left: -40px;
	    margin-top:20px;
	}
	
	.thumbnail {
	    padding: 0;
	}
	
	.carousel-inner>.item>img, .carousel-inner>.item>a>img {
	    width: 100%;
	}
	
	.carousel {
	    position: relative;
	    width: 400px;
	}	
	
	#slider-thumbs {
	    height: 400px;
	    overflow-y: scroll;
	    white-space: nowrap;
	}
			
	.pieID {
	  display: inline-block;
	  vertical-align: top;
	}
	
	.pie {
	  height: 200px;
	  width: 200px;
	  position: relative;
	  margin: 0 30px 30px 0;
	}

	.pie::before {
	  content: "";
	  display: block;
	  position: absolute;
	  z-index: 1;
	  width: 100px;
	  height: 100px;
	  background: #EEE;
	  border-radius: 50%;
	  top: 50px;
	  left: 50px;
	}
	
	.pie::after {
	  content: "";
	  display: block;
	  width: 120px;
	  height: 2px;
	  background: rgba(0, 0, 0, 0.1);
	  border-radius: 50%;
	  box-shadow: 0 0 3px 4px rgba(0, 0, 0, 0.1);
	  margin: 220px auto;
	}
	
	.slice {
	  position: absolute;
	  width: 200px;
	  height: 200px;
	  clip: rect(0px, 200px, 200px, 100px);
	  animation: bake-pie 1s;
	}

	.slice span {
	  display: block;
	  position: absolute;
	  top: 0;
	  left: 0;
	  background-color: black;
	  width: 200px;
	  height: 200px;
	  border-radius: 50%;
	  clip: rect(0px, 200px, 200px, 100px);
	}
	
	.legend {
	  list-style-type: none;
	  padding: 0;
	  margin: 0;
	  background: #FFF;
	  padding: 15px;
	  font-size: 13px;
	  box-shadow: 1px 1px 0 #DDD, 2px 2px 0 #BBB;
	  width: 200px;
	}
	
	.legend li {
	  width: 170px;
	  height: 1.25em;
	  margin-bottom: 0.7em;
	  padding-left: 0.5em;
	  border-left: 1.25em solid black;
	}
	
	.legend em {
	  font-style: normal;
	}
	
	.legend span {
	  float: right;
	}

	
</style>


<h1 class="page-header"> Edit <%= GetTerm("Equipment") %> Model - <%= GetBrandNameByModelIntRecID(InternalRecordIdentifier) %>&nbsp;<%= GetModelNameByIntRecID(InternalRecordIdentifier) %></h1>

<div class="container">
    <div class="row">
        <div class="col-md-6">

			<form method="POST" action="editModel_submit.asp" name="frmEditModel" id="frmEditModel" onsubmit="return validateEditModelForm();">
		
			<div class="col-lg-6">
			
				<input type="hidden" id="txtInternalRecordIdentifier" name="txtInternalRecordIdentifier" value="<%=InternalRecordIdentifier%>">
			
				<div class="row row-line">
					<div class="form-group col-lg-12">
						<label for="txtModel" class="col-sm-3 control-label">Model</label>	
		    			<div class="col-sm-8">
		    				<input type="text" class="form-control required" id="txtModel" name="txtModel" value="<%= Model %>">
		    			</div>
					</div>
				</div>
				
		
				<div class="row row-line">
					<div class="form-group col-lg-12">
						<label for="txtBrand" class="col-sm-3 control-label">Brand</label>	
		    			<div class="col-sm-8">
						  	<select class="form-control" name="selBrandIntRecID" id="selBrandIntRecID">
						  			<option value="">Select Brand For Model</option>
							      	<% 'Get all Brands 
							      	  	SQL9 = "SELECT * FROM EQ_Brands ORDER BY Brand ASC"
			
										Set cnn9 = Server.CreateObject("ADODB.Connection")
										cnn9.open (Session("ClientCnnString"))
										Set rs9 = Server.CreateObject("ADODB.Recordset")
										rs9.CursorLocation = 3 
										Set rs9 = cnn9.Execute(SQL9)
										If not rs9.EOF Then
											Do
												If cInt(BrandIntRecID) = cInt(rs9("InternalRecordIdentifier")) Then 
													Response.Write("<option value='" & rs9("InternalRecordIdentifier") & "' selected='selected'>" & rs9("Brand") & "</option>")
												Else
													Response.Write("<option value='" & rs9("InternalRecordIdentifier") & "'>" & rs9("Brand") & "</option>")
												End If
												rs9.movenext
											Loop until rs9.eof
										End If
										set rs9 = Nothing
										cnn9.close
										set cnn9 = Nothing
									%>
							</select>
		    			</div>
					</div>
				</div>
				
				
				
				<div class="row row-line">
					<div class="form-group col-lg-12">
						<label for="txtBrand" class="col-sm-3 control-label">Groups</label>	
		    			<div class="col-sm-8">
						  	<select class="form-control" name="selGroupIntRecID" id="selGroupIntRecID">
						  			<option value="">Select Group For Model</option>
							      	<% 'Get all Groups 
							      	  	SQL9 = "SELECT * FROM EQ_Groups ORDER BY GroupName ASC"
			
										Set cnn9 = Server.CreateObject("ADODB.Connection")
										cnn9.open (Session("ClientCnnString"))
										Set rs9 = Server.CreateObject("ADODB.Recordset")
										rs9.CursorLocation = 3 
										Set rs9 = cnn9.Execute(SQL9)
										If not rs9.EOF Then
											Do
												If cInt(GroupIntRecID) = cInt(rs9("InternalRecordIdentifier")) Then 
													Response.Write("<option value='" & rs9("InternalRecordIdentifier") & "' selected='selected'>" & rs9("GroupName") & "</option>")
												Else
													Response.Write("<option value='" & rs9("InternalRecordIdentifier") & "'>" & rs9("GroupName") & "</option>")
												End If
												rs9.movenext
											Loop until rs9.eof
										End If
										set rs9 = Nothing
										cnn9.close
										set cnn9 = Nothing
									%>
							</select>		
		    			</div>
					</div>
				</div>
				
				
				
				<div class="row row-line">
					<div class="form-group col-lg-12">
						<label for="txtBrand" class="col-sm-3 control-label">Class</label>	
		    			<div class="col-sm-8">
						  	<select class="form-control" name="selClassIntRecID" id="selClassIntRecID">
						  			<option value="">Select Class For Model</option>
							      	<% 'Get all Classes 
							      	  	SQL9 = "SELECT * FROM EQ_Classes ORDER BY Class ASC"
			
										Set cnn9 = Server.CreateObject("ADODB.Connection")
										cnn9.open (Session("ClientCnnString"))
										Set rs9 = Server.CreateObject("ADODB.Recordset")
										rs9.CursorLocation = 3 
										Set rs9 = cnn9.Execute(SQL9)
										If not rs9.EOF Then
											Do
												If cInt(ClassIntRecID) = cInt(rs9("InternalRecordIdentifier")) Then 
													Response.Write("<option value='" & rs9("InternalRecordIdentifier") & "' selected='selected'>" & rs9("Class") & "</option>")
												Else
													Response.Write("<option value='" & rs9("InternalRecordIdentifier") & "'>" & rs9("Class") & "</option>")
												End If
												rs9.movenext
											Loop until rs9.eof
										End If
										set rs9 = Nothing
										cnn9.close
										set cnn9 = Nothing
									%>
							</select>
		    			</div>
					</div>
				</div>
				
			</div>	
			<div class="col-lg-6">	
		
				<div class="row row-line">
					<div class="form-group col-lg-12">
						<label for="txtModel" class="col-sm-3 control-label">Default Rental Price</label>	
		    			<div class="col-sm-8">
		    				<i class="inside fa fa-usd"></i>
		    				<input type="text" class="form-control required inp" id="txtDefaultRentalPrice" name="txtDefaultRentalPrice" value="<%= DefaultRentalPrice %>">		    			</div>
					</div>
				</div>
				
				<div class="row row-line">
					<div class="form-group col-lg-12">
						<label for="txtModel" class="col-sm-3 control-label">Default Cost</label>	
		    			<div class="col-sm-8">
		    				 <i class="inside fa fa-usd"></i>
		    				<input type="text" class="form-control required inp" id="txtDefaultCost" name="txtDefaultCost" value="<%= DefaultCostPrice %>">
		    			</div>
					</div>
				</div>
				
				
				<div class="row row-line">
					<div class="form-group col-lg-12">
						<label for="txtModel" class="col-sm-3 control-label">Replacement Cost</label>	
		    			<div class="col-sm-8">
		    				 <i class="inside fa fa-usd"></i>
		    				<input type="text" class="form-control required inp" id="txtReplacementCost" name="txtReplacementCost" value="<%= ReplacementCost %>">
		    			</div>
					</div>
				</div>
				
				<div class="row row-line">
					<div class="form-group col-lg-12">
						<label for="txtModel" class="col-sm-3 control-label">Backend System Code</label>	
		    			<div class="col-sm-8">
		    				<input type="text" class="form-control inp" id="txtBackendSystemCode" name="txtBackendSystemCode" value="<%= BackendSystemCode %>">
		    			</div>
					</div>
				</div>		
				
				<div class="row row-line">
					<div class="form-group col-lg-12">
						<label for="txtModel" class="col-sm-3 control-label">Insight Asset Tag Prefix</label>	
		    			<div class="col-sm-6">
		    				<input type="text" class="form-control required inp" id="txtInsightAssetTagPrefix" name="txtInsightAssetTagPrefix" value="<%= InsightAssetTagPrefix %>">
		    			</div>
					</div>
				</div>					
				
				
			    <!-- cancel / submit !-->
				<div class="row row-line">
					&nbsp;
				</div>
				
				
			    <!-- cancel / submit !-->
				<div class="row row-line">
					<div class="col-lg-12 alertbutton">
						<div class="col-lg-12">
							<a href="<%= BaseURL %>equipment/models/main.asp">
			    				<button type="button" class="btn btn-default">&lsaquo; Cancel &amp; Go Back To Models List</button>
							</a>
							<button type="submit" class="btn btn-primary"><i class="far fa-save"></i> Save</button>
						</div>
				    </div>
				</div>
				
			</div>
				
			</form>
    </div><!-- eof col-md-6 -->
    
        
	
    <div class="col-md-2">
    
    	<% 
    	NumPiecesOfEquipByModel = NumberEquipmentRecsDefinedForModel(InternalRecordIdentifier)
    	
    	If cInt(NumPiecesOfEquipByModel) > 0 Then
    	%>
    
			<div class="pieID pie"></div>
			<ul class="pieID legend">
			
			<%
			 
			SQLStatusCodes = "SELECT * FROM EQ_StatusCodes"
			
			Set cnnStatusCodes = Server.CreateObject("ADODB.Connection")
			cnnStatusCodes.open (Session("ClientCnnString"))
			Set rsStatusCodes = Server.CreateObject("ADODB.Recordset")
			rsStatusCodes.CursorLocation = 3 
			Set rsStatusCodes = cnnStatusCodes.Execute(SQLStatusCodes)
			
			If NOT rsStatusCodes.EOF Then
				Do While NOT rsStatusCodes.EOF
				
					StatusIntRecID = rsStatusCodes("InternalRecordIdentifier")
					StatusBackendSystemCode = rsStatusCodes("statusBackendSystemCode")
					StatusDesc = rsStatusCodes("statusDesc")
					StatusAvailableForPlacement = rsStatusCodes("statusAvailableForPlacement")
					
	
					max=100
					min=1
					Randomize			
					%>
						<li>
							<em><%= StatusDesc %>-<%= StatusBackendSystemCode %></em>
							<span><%= NumberModelsWithStatusCode(StatusIntRecID,InternalRecordIdentifier) %></span>
							<!--<span><%= Int((max-min+1)*Rnd+min) %></span>-->
						</li>
					<% 
					rsStatusCodes.MoveNext
				Loop
			End If
			
			set rsStatusCodes = Nothing 
			%>
			
			</ul>
		<% Else %>
		
			No equipment matching this model found.<br><br>
			You may not own any pieces of equipment of this model type.
		
		<% End If%>
    </div>
    
    
    
    
    
    
    
    
    
    
    
    
    <div class="col-md-4">
     
     
	     <div class="container">
		    <div id="main_area">
		        <!-- Slider -->
		        <div class="row">
		            <div class="col-sm-3" id="slider-thumbs">
		                <!-- Bottom switcher of slider -->
		                <ul class="hide-bullets">
		                    <li class="col-sm-12">
		                        <a class="thumbnail" id="carousel-selector-0">
		                            <img src="http://placehold.it/400x400&text=zero">
		                        </a>
		                    </li>
		
		                    <li class="col-sm-12">
		                        <a class="thumbnail" id="carousel-selector-1"><img src="http://placehold.it/400x400&text=1"></a>
		                    </li>
		
		                    <li class="col-sm-12">
		                        <a class="thumbnail" id="carousel-selector-2"><img src="http://placehold.it/400x400&text=2"></a>
		                    </li>
		
		                    <li class="col-sm-12">
		                        <a class="thumbnail" id="carousel-selector-3"><img src="http://placehold.it/400x400&text=3"></a>
		                    </li>
		
		                    <li class="col-sm-12">
		                        <a class="thumbnail" id="carousel-selector-4"><img src="http://placehold.it/400x400&text=4"></a>
		                    </li>
		
		                    <li class="col-sm-12">
		                        <a class="thumbnail" id="carousel-selector-5"><img src="http://placehold.it/400x400&text=5"></a>
		                    </li>
		                    <li class="col-sm-12">
		                        <a class="thumbnail" id="carousel-selector-6"><img src="http://placehold.it/400x400&text=6"></a>
		                    </li>
		
		                    <li class="col-sm-12">
		                        <a class="thumbnail" id="carousel-selector-7"><img src="http://placehold.it/400x400&text=7"></a>
		                    </li>
		
		                    <li class="col-sm-12">
		                        <a class="thumbnail" id="carousel-selector-8"><img src="http://placehold.it/400x400&text=8"></a>
		                    </li>
		
		                    <li class="col-sm-12">
		                        <a class="thumbnail" id="carousel-selector-9"><img src="http://placehold.it/400x400&text=9"></a>
		                    </li>
		                    <li class="col-sm-12">
		                        <a class="thumbnail" id="carousel-selector-10"><img src="http://placehold.it/400x400&text=10"></a>
		                    </li>
		
		                    <li class="col-sm-12">
		                        <a class="thumbnail" id="carousel-selector-11"><img src="http://placehold.it/400x400&text=11"></a>
		                    </li>
		
		                    <li class="col-sm-12">
		                        <a class="thumbnail" id="carousel-selector-12"><img src="http://placehold.it/400x400&text=12"></a>
		                    </li>
		
		                    <li class="col-sm-12">
		                        <a class="thumbnail" id="carousel-selector-13"><img src="http://placehold.it/400x400&text=13"></a>
		                    </li>
		                    <li class="col-sm-12">
		                        <a class="thumbnail" id="carousel-selector-14"><img src="http://placehold.it/400x400&text=14"></a>
		                    </li>
		
		                    <li class="col-sm-12">
		                        <a class="thumbnail" id="carousel-selector-15"><img src="http://placehold.it/400x400&text=15"></a>
		                    </li>
		                </ul>
		            </div>
		            <div class="col-sm-8">
		                <div class="col-xs-12" id="slider">
		                    <!-- Top part of the slider -->
		                    <div class="row">
		                        <div class="col-sm-12" id="carousel-bounding-box">
		                            <div class="carousel slide" id="myCarousel">
		                                <!-- Carousel items -->
		                                <div class="carousel-inner">
		                                    <div class="active item" data-slide-number="0">
		                                        <a href="http://placehold.it/470x480&text=zero" target="_blank"><img src="http://placehold.it/470x480&text=zero"></a></div>
		
		                                    <div class="item" data-slide-number="1">
		                                        <a href="http://placehold.it/470x480&text=1" target="_blank"><img src="http://placehold.it/470x480&text=1"></a></div>
		
		                                    <div class="item" data-slide-number="2">
		                                        <a href="http://placehold.it/470x480&text=2" target="_blank"><img src="http://placehold.it/470x480&text=2"></a></div>
		
		                                    <div class="item" data-slide-number="3">
		                                        <a href="http://placehold.it/470x480&text=3" target="_blank"><img src="http://placehold.it/470x480&text=3"></a></div>
		
		                                    <div class="item" data-slide-number="4">
		                                        <a href="http://placehold.it/470x480&text=4" target="_blank"><img src="http://placehold.it/470x480&text=4"></a></div>
		
		                                    <div class="item" data-slide-number="5">
		                                        <a href="http://placehold.it/470x480&text=5" target="_blank"><img src="http://placehold.it/470x480&text=5"></a></div>
		                                    
		                                    <div class="item" data-slide-number="6">
		                                        <a href="http://placehold.it/470x480&text=6" target="_blank"><img src="http://placehold.it/470x480&text=6"></a></div>
		                                    
		                                    <div class="item" data-slide-number="7">
		                                        <a href="http://placehold.it/470x480&text=7" target="_blank"><img src="http://placehold.it/470x480&text=7"></a></div>
		                                    
		                                    <div class="item" data-slide-number="8">
		                                        <a href="http://placehold.it/470x480&text=8" target="_blank"><img src="http://placehold.it/470x480&text=8"></a></div>
		                                    
		                                    <div class="item" data-slide-number="9">
		                                        <a href="http://placehold.it/470x480&text=9" target="_blank"><img src="http://placehold.it/470x480&text=9"></a></div>
		                                    
		                                    <div class="item" data-slide-number="10">
		                                        <a href="http://placehold.it/470x480&text=10" target="_blank"><img src="http://placehold.it/470x480&text=10"></a></div>
		                                    
		                                    <div class="item" data-slide-number="11">
		                                        <a href="http://placehold.it/470x480&text=11" target="_blank"><img src="http://placehold.it/470x480&text=11"></a></div>
		                                    
		                                    <div class="item" data-slide-number="12">
		                                        <a href="http://placehold.it/470x480&text=12" target="_blank"><img src="http://placehold.it/470x480&text=12"></a></div>
		
		                                    <div class="item" data-slide-number="13">
		                                        <a href="http://placehold.it/470x480&text=13" target="_blank"><img src="http://placehold.it/470x480&text=13"></a></div>
		
		                                    <div class="item" data-slide-number="14">
		                                        <a href="http://placehold.it/470x480&text=14" target="_blank"><img src="http://placehold.it/470x480&text=14"></a></div>
		
		                                    <div class="item" data-slide-number="15">
		                                        <a href="http://placehold.it/470x480&text=15" target="_blank"><img src="http://placehold.it/470x480&text=15"></a></div>
		                                </div>
		                                <!-- Carousel nav -->
		                                <a class="left carousel-control" href="#myCarousel" role="button" data-slide="prev">
		                                    <span class="glyphicon glyphicon-chevron-left"></span>
		                                </a>
		                                <a class="right carousel-control" href="#myCarousel" role="button" data-slide="next">
		                                    <span class="glyphicon glyphicon-chevron-right"></span>
		                                </a>
		                            </div>
		                        </div>
		                    </div>
		                </div>
		            </div>
		            <!--/Slider-->
		        </div>
		
		    </div>
		</div>
     
     </div>
     
</div><!-- eof row -->	
<div class="row">
		 
	<!-- tabs start here !-->
	<div class="bottom-table">
		<div class="row">
			<div class="col-lg-12">
				<div class="bottom-tabs-section">
	
					<!-- tab navigation !-->
					<ul class="nav nav-tabs" role="tablist">
						<li role='presentation' class="active"><a href='#documents' class='EquipmentTabDocumentColor' aria-controls='documents' role='tab' data-toggle='tab'>Documents <div style="display:inline" id="docsNum">(<%= NumberOfDocumentsByModelIntRecID(ModelIntRecID) %>)</div></a></li>
						<li role='presentation'><a href='#images' class='EquipmentTabImageColor' aria-controls='images' role='tab' data-toggle='tab'>Images <div style="display:inline" id="imagesNum">(<%= NumberOfImagesByModelIntRecID(ModelIntRecID) %>)</div></a></li>
						<li role='presentation'><a href='#links' class='EquipmentTabLinkColor' aria-controls='links' role='tab' data-toggle='tab'>Links <div style="display:inline" id="linksNum">(<%= NumberOfLinksByModelIntRecID(ModelIntRecID) %>)</div></a></li>
						<li role='presentation'><a href='#companions' class='EquipmentTabCompanionColor' aria-controls='companions' role='tab' data-toggle='tab' id="companionsNum">Companions</a></li>
					</ul>
					<!-- eof tab navigation -->
				
					<div class="tab-content">
						<!--#include file="editModel_documents_tab.asp"-->
						<!--#include file="editModel_images_tab.asp"-->
						<!--#include file="editModel_links_tab.asp"-->
						<!--#include file="editModel_compequipment_tab.asp"-->
					</div>
						
				</div><!-- eof bottom-tabs-section-->
			</div><!-- eof col-lg-12 -->
		</div><!-- eof row -->
	</div><!-- eof bottom-table -->
</div><!-- eof row -->


</div><!-- eof content container -->


<!-- tabs js  !-->
<script type="text/javascript">
		 $('a[data-toggle="tab"]').on('shown.bs.tab', function (e) {
		  e.target // newly activated tab
		  e.relatedTarget // previous active tab
	})
</script>

<script>
		$(document).ready(function(){
		  $("#demo").on("hide.bs.collapse", function(){
		    $(".btn-custom").html('<span class="glyphicon glyphicon-collapse-down"></span> Click to Expand');
		  });
		  $("#demo").on("show.bs.collapse", function(){
		    $(".btn-custom").html('<span class="glyphicon glyphicon-collapse-up"></span> Click to Collapse');
		  });
		});
</script>
 <!-- eof tabs js !-->

 <!-- custom table search !-->

<script>

$(document).ready(function () {

    (function ($) {
        
        $('#filter-documents').keyup(function () {

            var rex = new RegExp($(this).val(), 'i');
            $('.searchable-documents tr').hide();
            $('.searchable-documents tr').filter(function () {
                return rex.test($(this).text());
            }).show();
        })
        
        $('#filter-images').keyup(function () {

            var rex = new RegExp($(this).val(), 'i');
            $('.searchable-images tr').hide();
            $('.searchable-images tr').filter(function () {
                return rex.test($(this).text());
            }).show();
        })

        
        $('#filter-links').keyup(function () {

            var rex = new RegExp($(this).val(), 'i');
            $('.searchable-links tr').hide();
            $('.searchable-links tr').filter(function () {
                return rex.test($(this).text());
            }).show();
        })

    }(jQuery));

});
</script>
<!-- eof custom table search !-->


<!-- checkboxes JS !-->
<script type="text/javascript">
    function changeState(el) {
        if (el.readOnly) el.checked=el.readOnly=false;
        else if (!el.checked) el.readOnly=el.indeterminate=true;
    }
</script>
<!-- eof checkboxes JS !-->

 

<!--#include file="../../inc/footer-main.asp"-->
