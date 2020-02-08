<!--#include file="../../../inc/InsightFuncs.asp"-->
<!--#include file="../../../inc/InsightFuncs_InventoryControl.asp"-->
<!--#include file="../../../inc/mail.asp"-->
<% 
ClientKeyForFileNames = MUV_READ("ClientKeyForFileNames")

sURL = Request.ServerVariables("SERVER_NAME")
'baseURL should always have a trailing /slash, just in case, handle either way
If right(sURL ,1)="/" Then maildomain = Left(right(sURL ,len(sURL )-7),len(right(sURL ,len(sURL )-7))-1) Else maildomain = right(sURL ,len(sURL )-7)

%>

<style type="text/css">
	
	.tt-menu,
	.gist {
	  text-align: left;
	  width: 100%;
	}
	
	.typeahead,
	.tt-query,
	.tt-hint {
	 width: 100% !important;
	  height: 50px;
	  padding: 8px 12px;
	  font-size: 16px;
	  line-height: 30px;
	  -webkit-border-radius: 8px;
	     -moz-border-radius: 8px;
	          border-radius: 8px;
	  outline: none;
	  
		border: 1px solid #ccc;
	    border-radius: 4px;
	    -webkit-box-shadow: inset 0 1px 1px rgba(0,0,0,.075);
	    box-shadow: inset 0 1px 1px rgba(0,0,0,.075);
	    -webkit-transition: border-color ease-in-out .15s,-webkit-box-shadow ease-in-out .15s;
	    -o-transition: border-color ease-in-out .15s,box-shadow ease-in-out .15s;
	    transition: border-color ease-in-out .15s,box-shadow ease-in-out .15s;	  
	}
	
	.typeahead {
	  background-color: #fff;
	}
	
	.typeahead:focus {
	  border: 2px solid #0097cf;
	}
	
	.tt-query {
	  -webkit-box-shadow: inset 0 1px 1px rgba(0, 0, 0, 0.075);
	     -moz-box-shadow: inset 0 1px 1px rgba(0, 0, 0, 0.075);
	          box-shadow: inset 0 1px 1px rgba(0, 0, 0, 0.075);
	}
	
	.tt-hint {
	  color: #999
	}
	
	.tt-menu {
	  width: 100%;
	  margin: 12px 0;
	  padding: 8px 0;
	  background-color: #fff;
	  border: 1px solid #ccc;
	  border: 1px solid rgba(0, 0, 0, 0.2);
	  -webkit-border-radius: 8px;
	     -moz-border-radius: 8px;
	          border-radius: 8px;
	  -webkit-box-shadow: 0 5px 10px rgba(0,0,0,.2);
	     -moz-box-shadow: 0 5px 10px rgba(0,0,0,.2);
	          box-shadow: 0 5px 10px rgba(0,0,0,.2);
	}
	
	.tt-suggestion {
	  padding: 3px 20px;
	  font-size: 16px;
	  line-height: 18px;
	}
	
	.tt-suggestion:hover {
	  cursor: pointer;
	  color: #fff;
	  background-color: #0097cf;
	}
	
	.tt-suggestion.tt-cursor {
	  color: #fff;
	  background-color: #0097cf;
	
	}
	
	.tt-suggestion p {
	  margin: 0;
	}
		
	/* scrollable dropdown specific styles */
	/* ----------------------- */
	
	#scrollable-dropdown-menu .empty-message {
	  padding: 5px 10px;
	 text-align: center;
	}
		
	
	#scrollable-dropdown-menu .tt-menu {
	   max-height: 150px;
	   overflow-y: auto;
	 }
 
	/** Added tp make typeahead 100% screen width */
	.twitter-typeahead{
	     width: 98%;
	}
	.tt-dropdown-menu{
	    width: 102%;
	}
	input.typeahead.tt-query{ /* This is optional */
	    width: 300px !important;
	}	
	
</style>

<script type="text/javascript">

	$(document).ready(function() { 
	
		var productList = new Bloodhound({
		  datumTokenizer: Bloodhound.tokenizers.obj.whitespace(['value','description']),
		  queryTokenizer: Bloodhound.tokenizers.whitespace,
		  prefetch: "../../../clientfiles/<%= ClientKeyForFileNames %>/autocomplete/product_list_mobile_<%= ClientKeyForFileNames %>.json",
		});
		
		productList.initialize();
		productList.clearPrefetchCache();		
	
		$('#scrollable-dropdown-menu .typeahead').typeahead(null, {
		  name: 'product-list',
		  limit: 10,
		  display: 'display',
		  source: productList,
		  hint: false,
		  highlight: true,
		  minLength: 1,	  
		  templates: {
		    empty: [
		      '<div class="empty-message">',
		        'unable to find any products that match the current query',
		      '</div>'
		    ].join('\n'),
		    suggestion: function(data) {
	    		return '<p><strong>' + data.value + '</strong> – ' + data.description + '</p>';
			}
		  }
		  
		}).on('typeahead:selected', function (obj, datum) {
		
		    //console.log(obj);
		    //console.log(datum);
		    //console.log(datum.value);
		    
		    var prodSKU = datum.value;
	    
             $("#txtProdSKUSelected").val(prodSKU);
             
			 if (prodSKU!=""){
			 	$.ajax({
					type:"POST",
					url: "../../../inc/InsightFuncs_AjaxForInventoryControl.asp",
					data: "action=ReturnUMInfoForProduct&prodSKU="+encodeURIComponent(prodSKU),
						success: function(msg){
							$("#productUMInfo").html(msg);
							$("#yesSavedProdSKU").show();
							$("#btnHoldLastSKUEntered").prop('value', 'SAME AS LAST SKU: ' + prodSKU);
							$("#noSavedProdSKU").hide();
						}
				}) 
			  }
		    
		});
		
		
		$("#btnHoldLastSKUEntered").click(function() {
		
			var prodSKU = $("#btnHoldLastSKUEntered").text().split(':').pop().trim();
			$("#txtProdSKUSelected").val(prodSKU);
			$('.typeahead').typeahead('val', prodSKU);
			$(".typeahead").trigger("click");

			 if (prodSKU!=""){
			 	$.ajax({
					type:"POST",
					url: "../../../inc/InsightFuncs_AjaxForInventoryControl.asp",
					data: "action=ReturnUMInfoForProduct&prodSKU="+encodeURIComponent(prodSKU),
						success: function(msg){
							$("#productUMInfo").html(msg);
							$("#yesSavedProdSKU").show();
							$("#btnHoldLastSKUEntered").prop('value', 'SAME AS LAST SKU: ' + prodSKU);
							$("#noSavedProdSKU").hide();
						}
				}) 
			  }
		
		});
		  		
		
		$("#btnAssignUPCToProduct").click(function() {

             var prodSKU = $("#txtProdSKUSelected").val();
             var prodUM = $("#selProdUM option:selected").val();
             var prodUPCCode = $("#txtEnteredUPCCode").val();
             
			 if (prodSKU != ""){
			 
			 	if (prodUM != "") {
			 
				 	$.ajax({
						type:"POST",
						url: "../../../inc/InsightFuncs_AjaxForInventoryControl.asp",
						data: "action=AssignUPCCodeToProductAndUM&prodSKU="+encodeURIComponent(prodSKU)+"&prodUM="+encodeURIComponent(prodUM)+"&prodUPC="+encodeURIComponent(prodUPCCode),
							success: function(msg){
								window.location = "upclookup.asp";
								//$("#postInfo").html(msg);
							}
					}) 
				}
				else {
					swal("Please select a unit of measure");
				}
			  }
			  else {
			  	swal("Please select a product");
			  }
		});		
		
		
		
		$("#btnChangeProduct").click(function() {

             var prodUPCCode = $("#txtUPCCodeToPass").val();
             
			 if (prodUPCCode != ""){
			 	$.ajax({
					type:"POST",
					url: "../../../inc/InsightFuncs_AjaxForInventoryControl.asp",
					data: "action=RemoveUPCCodeFromICProduct&prodUPC="+encodeURIComponent(prodUPCCode),
						success: function(msg){
					        $(".scan-result").html("");
					        $(".scan-result").load("upcload.asp", { code: prodUPCCode }, function (response, status, xhr) {
					            if (status == "error") {
					                var msg = "Sorry but there was an error: ";
					                $("#error").html(msg + xhr.status + " " + xhr.statusText);
					             } 
					            $('#txtUPCCode').val(prodUPCCode);
					            setTimeout(function () { $('#txtUPCCode').focus(); event.preventDefault(); }, 50);
					        });
							//$("#postInfo").html(msg);
						}
				}) 
			  }
		});		
		
		
		$("#btnClearTypeahead").click(function() {
             $('.typeahead').typeahead('val', '');
             $('.typeahead').focus();
   		});		
		
		

	})
</script>

<%

If Request.Form("code") <> "" Then

	UPCCode = Request.Form("code")
	prodSKU =  GetProdSKUByUPC(UPCCode)

	prodDesc = ""
	prodUM = ""
	prodBin = ""

	' Maybe they typed a sku instead of scanning a UPC
	If prodSKU = "" Then

		Set cnnprodLookup = Server.CreateObject("ADODB.Connection")
		cnnprodLookup.open (Session("ClientCnnString"))
		Set rsprodLookup = Server.CreateObject("ADODB.Recordset")
		rsprodLookup.CursorLocation = 3 
			
		SQL_prodLookup = "SELECT * FROM IC_Product WHERE prodSKU = '" & UPCCode & "'"
		Set rsprodLookup = cnnprodLookup.Execute(SQL_prodLookup)
	
		If Not rsprodLookup.EOF Then
			prodSKU = rsprodLookup("prodSKU")
			UPCCode = ""
		End If
		
		Set rsprodLookup = Nothing
		cnnprodLookup.Close
		Set cnnprodLookup = Nothing

	End If

	If prodSKU <> "" Then  
	
		Set cnnprodLookup = Server.CreateObject("ADODB.Connection")
		cnnprodLookup.open (Session("ClientCnnString"))
		Set rsprodLookup = Server.CreateObject("ADODB.Recordset")
		rsprodLookup.CursorLocation = 3 
			
		SQL_prodLookup = "SELECT * FROM IC_Product WHERE prodUnitUPC = '" & UPCCode & "' OR prodCaseUPC = '" & UPCCode & "'"
		SQL_prodLookup = SQL_prodLookup & " OR prodSKU = '" & prodSKU & "'"
		
		Set rsprodLookup = cnnprodLookup.Execute(SQL_prodLookup)
		
		If Not rsprodLookup.EOF Then
			'First get the on hand qty
			QtyOnHand_Units = rsprodLookup("QtyOnHand_Units") ' Initially set it to ours in case the post fails
			QtyOnHand_LastUpdated = rsprodLookup("QtyOnHand_LastUpdated")
			
			'Post to the backend to try to get an up-to-the-minute on hand
			'Construct xml fields based on record
			xmlData = "<DATASTREAM>"
			xmlData = xmlData & "<IDENTITY>Pm8316wyc011</IDENTITY>"
					
			xmlData = xmlData & "<MODE>" & GetPOSTParams("BackendInventoryPostsMode") & "</MODE>"
					
			xmlData = xmlData & "  <RECORD_TYPE>INVENTORY</RECORD_TYPE>"
			xmlData = xmlData & "  <RECORD_SUBTYPE>QUERY_ONHAND</RECORD_SUBTYPE>"
					
			xmlData = xmlData & "<SERNO>" & MUV_READ("SERNO") & "</SERNO>"
					
						
			xmlData = xmlData & " <QUERY_ONHAND>"

			xmlData = xmlData & "        <PROD_ID>" & prodSKU & "</PROD_ID>"
			xmlData = xmlData & "        <RETURN_VALUE_UM>U</RETURN_VALUE_UM>"
			
			xmlData = xmlData & " </QUERY_ONHAND>"
				 
			xmlData = xmlData & "</DATASTREAM>"
					
					
			xmlDataForDisp = Replace(xmlData,"<","[")
			xmlDataForDisp = Replace(xmlDataForDisp ,">","]")
			xmlDataForDisp = Replace(xmlDataForDisp ,"][","]<br>[")
			xmlDataForDisp = Replace(xmlDataForDisp ,"[","</b>[")
			xmlDataForDisp = Replace(xmlDataForDisp ,"]","]<b>")
		
			Set httpRequest = Server.CreateObject("MSXML2.ServerXMLHTTP")
		
			'Response.Write(GetPOSTParams("BackendInventoryPostsURL"))
			
			httpRequest.Open "POST", GetPOSTParams("BackendInventoryPostsURL"), False
			httpRequest.SetRequestHeader "Content-Type", "text/xml"
			
			xmlData = Replace(xmlData,"&","&amp;")
			xmlData = Replace(xmlData,chr(34),"")			
			httpRequest.Send xmlData
		
			data = xmlData
		
			If (Err.Number <> 0 ) Then
				emailbody="httpRequest.status returned " & httpRequest.status & " when posting <RECORD_TYPE>INVENTORY and <RECORD_SUBTYPE>QUERY ONHAND"& "<br><br>"
				emailBody = "httpRequest.responseText:" & httpRequest.responseText & "<br><br>"
				emailBody = emailBody & "PAGE: " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "<br>"
				emailBody = emailBody & "Posted to " & GetPOSTParams("BackendInventoryPostsURL") & "<br><br>"
				emailBody = emailBody & "POSTED DATA:<br>" & xmlDataForDisp & "<br><br>"
				emailBody = emailBody & "SERNO: " & MUV_READ("SERNO") & "<br>"
				SendMail "mailsender@" & maildomain ,"rich@ocsaccess.com",SERNO & " Post Error Inventory Query On Hand",emailBody, "Inventory API", "Inventory API"
				Description = emailBody 
				Write_API_AuditLog_Entry Request.ServerVariables("SERVER_NAME"),Description,GetPOSTParams("BackendInventoryPostsMode"),MUV_READ("SERNO"),MUV_READ("SERNO"),"Inventory API"
			End If
		
			If httpRequest.status = 200 THEN 
			
				If IsNumeric(httpRequest.responseText) Then ' Success
			
					emailbody="httpRequest.status returned " & httpRequest.status & " when posting <RECORD_TYPE>INVENTORY and <RECORD_SUBTYPE>QUERY ONHAND<"& "<br><br>"
					emailBody = emailBody & "httpRequest.responseText:" & httpRequest.responseText & "<br><br>"
					emailBody = emailBody & "PAGE: " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "<br>"
					emailBody = emailBody & "Posted to " & GetPOSTParams("BackendInventoryPostsURL") & "<br><br>"
					emailBody = emailBody & "POSTED DATA:<br>" & xmlDataForDisp & "<br><br>"
					emailBody = emailBody & "SERNO: " & MUV_READ("SERNO") & "<br>"
					
					SendMail "mailsender@" & maildomain ,"insight@ocsaccess.com", MUV_READ("SERNO") & " Good Post Inventory Query On Hand",emailBody, "Inventory API", "Inventory API"
					
					Call Write_API_AuditLog_Entry(Identity ,emailBody,GetPOSTParams("BackendInventoryPostsMode"),"CountOnHand_submit.asp")
					
					QtyOnHand_Units = httpRequest.responseText
					QtyOnHand_UnitsStatus = "LIVE"
					QtyOnHand_UnitsStatus = "1 MIN AGO"
					
					
				Else
					'FAILURE
					emailbody="httpRequest.status returned " & httpRequest.status & " when posting <RECORD_TYPE>INVENTORY and <RECORD_SUBTYPE>QUERY ONHAND<"& "<br><br>"
					emailBody = emailBody & "httpRequest.responseText:" & httpRequest.responseText & "<br><br>"
					emailBody = emailBody & "PAGE: " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "<br>"
					emailBody = emailBody & "Posted to " & GetPOSTParams("BackendInventoryPostsURL") & "<br><br>"
					emailBody = emailBody & "POSTED DATA:<br>" & xmlDataForDisp & "<br><br>"
					emailBody = emailBody & "SERNO: " & MUV_READ("SERNO") & "<br>"
					
					SendMail "mailsender@" & maildomain ,"rich@ocsaccess.com",MUV_READ("SERNO") & " Post Error Inventory Query On Hand",emailBody, "Inventory API", "Inventory API"
				
					QtyOnHand_UnitsStatus = "CACHED"
					
					QtyOnHand_LastUpdatedTimeDiff = DateDiff("n",Now(),QtyOnHand_LastUpdated)
					QtyOnHand_LastUpdatedTimeDiff = Abs(QtyOnHand_LastUpdatedTimeDiff) 
					
					If cLng(QtyOnHand_LastUpdatedTimeDiff) >= cInt(1440) Then
					
						QtyOnHand_LastUpdatedDays = QtyOnHand_LastUpdatedTimeDiff \ 1440
						QtyOnHand_LastUpdatedHours = (QtyOnHand_LastUpdatedTimeDiff - (QtyOnHand_LastUpdatedDays * 1440)) \ 60
						QtyOnHand_LastUpdatedMinutes = (QtyOnHand_LastUpdatedTimeDiff - (QtyOnHand_LastUpdatedDays * 1440)) mod 60 
						QtyOnHand_UnitsStatus = QtyOnHand_LastUpdatedDays & " DAYS " & QtyOnHand_LastUpdatedHours & " HRS " & QtyOnHand_LastUpdatedMinutes & " MIN AGO "
						
					ElseIf cLng(QtyOnHand_LastUpdatedTimeDiff) >= cInt(60) AND cLng(QtyOnHand_LastUpdatedTimeDiff) < cInt(1440) Then
					
						QtyOnHand_LastUpdatedHours = QtyOnHand_LastUpdatedTimeDiff \ 60
						QtyOnHand_LastUpdatedMinutes = QtyOnHand_LastUpdatedTimeDiff - (QtyOnHand_LastUpdatedHours * 60)
						QtyOnHand_UnitsStatus = QtyOnHand_LastUpdatedHours & " HRS " & QtyOnHand_LastUpdatedMinutes & " MIN AGO "
						
					ElseIf cLng(QtyOnHand_LastUpdatedTimeDiff) < cInt(60) Then
					
						QtyOnHand_LastUpdatedMinutes = QtyOnHand_LastUpdatedTimeDiff
						QtyOnHand_UnitsStatus = QtyOnHand_LastUpdatedMinutes & " MIN AGO "
						
					End If
				
					Call Write_API_AuditLog_Entry(Identity ,emailBody ,GetPOSTParams("BackendInventoryPostsMode"),"CountOnHand_submit.asp")
					
				End If
				
			Else
			
					'FAILURE
					emailbody="NON 200 Response - httpRequest.status returned " & httpRequest.status & " when posting <RECORD_TYPE>INVENTORY and <RECORD_SUBTYPE>QUERY ONHAND<"& "<br><br>"
					emailBody = emailBody & "httpRequest.responseText:" & httpRequest.responseText & "<br><br>"
					emailBody = emailBody & "PAGE: " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "<br>"
					emailBody = emailBody & "Posted to " & GetPOSTParams("BackendInventoryPostsURL") & "<br><br>"
					emailBody = emailBody & "POSTED DATA:<br>" & xmlDataForDisp & "<br><br>"
					emailBody = emailBody & "SERNO: " & MUV_READ("SERNO") & "<br>"
					
					SendMail "mailsender@" & maildomain ,"rich@ocsaccess.com",MUV_READ("SERNO") & " Post Error Inventory Query On Hand",emailBody, "Inventory API", "Inventory API"
				
				
					QtyOnHand_UnitsStatus = "CACHED"
					
					QtyOnHand_LastUpdatedTimeDiff = DateDiff("n",Now(),QtyOnHand_LastUpdated)
					QtyOnHand_LastUpdatedTimeDiff = Abs(QtyOnHand_LastUpdatedTimeDiff) 
					
					If cLng(QtyOnHand_LastUpdatedTimeDiff) >= cInt(1440) Then
					
						QtyOnHand_LastUpdatedDays = QtyOnHand_LastUpdatedTimeDiff \ 1440
						QtyOnHand_LastUpdatedHours = (QtyOnHand_LastUpdatedTimeDiff - (QtyOnHand_LastUpdatedDays * 1440)) \ 60
						QtyOnHand_LastUpdatedMinutes = (QtyOnHand_LastUpdatedTimeDiff - (QtyOnHand_LastUpdatedDays * 1440)) mod 60 
						QtyOnHand_UnitsStatus = QtyOnHand_LastUpdatedDays & " DAYS " & QtyOnHand_LastUpdatedHours & " HRS " & QtyOnHand_LastUpdatedMinutes & " MIN AGO "
						
					ElseIf cLng(QtyOnHand_LastUpdatedTimeDiff) >= cInt(60) AND cLng(QtyOnHand_LastUpdatedTimeDiff) < cInt(1440) Then
					
						QtyOnHand_LastUpdatedHours = QtyOnHand_LastUpdatedTimeDiff \ 60
						QtyOnHand_LastUpdatedMinutes = QtyOnHand_LastUpdatedTimeDiff - (QtyOnHand_LastUpdatedHours * 60)
						QtyOnHand_UnitsStatus = QtyOnHand_LastUpdatedHours & " HRS " & QtyOnHand_LastUpdatedMinutes & " MIN AGO "
						
					ElseIf cLng(QtyOnHand_LastUpdatedTimeDiff) < cInt(60) Then
					
						QtyOnHand_LastUpdatedMinutes = QtyOnHand_LastUpdatedTimeDiff
						QtyOnHand_UnitsStatus = QtyOnHand_LastUpdatedMinutes & " MIN AGO "
						
					End If
				
					Call Write_API_AuditLog_Entry(Identity ,emailBody,GetPOSTParams("BackendInventoryPostsMode"),"CountOnHand_submit.asp")
		
			End If

			
			prodCaseConversionFactor = rsprodLookup("prodCaseConversionFactor")			
			prodCasePricing = rsprodLookup("prodCasePricing")
			
			'Figure out which description & u/m to get
			'Unit
			If UPCCode = rsprodLookup("prodUnitUPC") Then 
				prodDesc = rsprodLookup("prodDescription")
				prodUM = "Unit"
				prodUnitBin = rsprodLookup("prodUnitBin")
				prodCaseBin = rsprodLookup("prodCaseBin")
			End If
			'Case
			If UPCCode = rsprodLookup("prodCaseUPC") Then 
				prodDesc = rsprodLookup("prodDescription")
				prodUM = "Case"
				prodUnitBin = rsprodLookup("prodUnitBin")
				prodCaseBin = rsprodLookup("prodCaseBin")		
			End If
			' N products
			If prodCasePricing = "N" Then 
				prodDesc = rsprodLookup("prodDescription")
				prodUM = "Unit"
				prodBin = rsprodLookup("prodUnitBin")			
			End If
			
			If UPCCode = "" Then ' This means they typed a sku
				If prodCasePricing = "N" Then 
					prodDesc = rsprodLookup("prodDescription")
					prodUM = "Unit"
					prodBin = rsprodLookup("prodUnitBin")	
					prodUnitBin = rsprodLookup("prodUnitBin")
					prodCaseBin = rsprodLookup("prodCaseBin")							
				Else
					prodDesc = rsprodLookup("prodDescription")		' we have no way to know if it's a unit or bin so go to unit		
				End If
			End If
		
		End If
		
		
		Set rsprodLookup = Nothing
		cnnprodLookup.Close
		Set cnnprodLookup = Nothing
	
		prodImage = GetProdImage(prodSKU)%>
	
		<div class="container-fluid inventory-upc-container">	
	
		    <div class="row">
		    
				<div class="col-lg-7 col-md-7 col-sm-7 col-xs-7">
					<div class="container-fluid" style="padding-left:0; padding-right:0">
						<div class="row">
							<div class="col-xs-12" style="padding-bottom:10px;"><strong>UPC Code:</strong>&nbsp;<strong class="red"><%= UPCCode %></strong></div>
							<div class="col-xs-12" style="padding-bottom:10px;"><strong>Product ID:</strong>&nbsp;<strong class="red"><%= prodSKU %></strong></div>
							
							<input type="hidden" id="txtUPCCodeToPass" name="txtUPCCodeToPass" value="<%= UPCCode %>">
							
							<% If UPCCode <> "" Then %>
								<div class="col-xs-12" style="padding-bottom:10px;"><button class="btn btn-warning btn-go btn-sm" id="btnChangeProduct">Move UPC to New Product</button></div>
							<% End If %>
							
							<% If prodCasePricing = "N" Then %>
								<div class="col-xs-12" style="padding-bottom:10px;">
									<strong>U/M:</strong>&nbsp;<strong class="red">N</strong>
								</div>
							<% End If %>
							
							<% If prodCasePricing = "N" Then %>
								<div class="col-xs-12" style="padding-bottom:5px;"><strong>Bin:</strong>&nbsp;<strong class="red"><%= prodBin %></strong></div>
							<% Else %>
								<div class="col-xs-12" style="padding-bottom:5px;"><strong>Unit Bin:</strong>&nbsp;<strong class="red"><%= prodUnitBin %></strong></div>
								<div class="col-xs-12" style="padding-bottom:5px;"><strong>Case Bin:</strong>&nbsp;<strong class="red"><%= prodCaseBin %></strong></div>
							<% End If %>
						</div>
					</div>
				</div>
				
				<div class="col-lg-5 col-md-5 col-sm-5 col-xs-5">
					<% If prodImage <> "" Then %>
						<img src="<%=GetProdImage(prodSKU )%>" class="general-image mobile-image img-thumbnail" style="width:100%;">
					<% End IF %>
				</div>
				
			</div>
		        
			<div class="row row-line">
				<div class="col-xs-12" style="padding-bottom:10px;"><strong>Description:</strong>&nbsp;<strong class="red"><%= prodDesc %></strong></div>
			</div>
			
			<div class="row">	
			
				<div class="col-xs-12" style="border-top:2px solid #000000;"></div>
				
				<% If prodCasePricing = "N" Then %>
				
						<div class="row" style="padding-bottom:5px;">
							<div class="col-xs-6 text-center"><strong>On Hand</strong></div>
							<div class="col-xs-6 text-center"><strong>Tot Units</strong></div>
						</div>
						
						<div class="row row-info">
							<div class="col-xs-6 text-center"><%= QtyOnHand_UnitsStatus %></div>
							<div class="col-xs-6 text-center"><strong class="red"><%= QtyOnHand_Units %></strong></div>				
						</div>
						
		 		<% Else %>
		 		
		 				<div class="row" style="padding-bottom:5px;">
							<div class="col-xs-3"><strong>On Hand</strong></div>
							<div class="col-xs-3 text-center"><strong>Tot Units</strong></div>	
							<div class="col-xs-3 text-center"><strong>Cases</strong></div>
							<div class="col-xs-3 text-center"><strong>Units</strong></div>
						</div>
						
						<div class="row row-info">
						
							<div class="col-xs-3 text-center"><%= QtyOnHand_UnitsStatus %></div>

							<div class="col-xs-3 text-center">
								<strong class="red"><%= QtyOnHand_Units %></strong>
							</div>	
							
							<% If prodCaseConversionFactor <> "" Then %>
								<div class="col-xs-3 text-center"><strong class="red"><%= Int(QtyOnHand_Units / cInt(prodCaseConversionFactor))  %></strong></div>
							<% Else %>
								<div class="col-xs-3 text-center">---</div>
							<% End If %>
							
							<% If QtyOnHand_Units Mod cint(prodCaseConversionFactor) = 0 Then %>	
								<div class="col-xs-3 text-center"><strong class="red">---</strong></div>
							<% Else %>
								<div class="col-xs-3 text-center"><strong class="red"><%=  QtyOnHand_Units Mod cInt(prodCaseConversionFactor) %></strong></div>	
							<% End If %>	
							
								
						</div>
		 		<% End If %>
				
				</div>
				
				<form method="POST" action="CountOnHand_submit.asp" name="frmCountOnHand" onsubmit="return validateCountOnHandform();">

					<input type="hidden" name="txtprodSKU" id="txtprodSKU" value="<%= prodSKU %>">
					<input type="hidden" name="txtproductUPC" id="txtproductUPC" value="<%= UPCCode %>">    
			
					
					<% If prodCasePricing <> "N" Then %>
					
						<div class="row row-line">
							<div class="col-xs-4"><strong>New Count:</strong></div>
							<div class="col-xs-3"><strong>Cases</strong></div>
							<div class="col-xs-3"><strong>Units</strong></div>
							<div class="col-xs-4">&nbsp;</div>
							<div class="col-xs-3"><input type="search" class="form-control" name="txtCasesCounted" id="txtCasesCounted" AUTOCOMPLETE="off"></div>
							<div class="col-xs-3"><input type="search" class="form-control" name="txtUnitsCounted" id="txtUnitsCounted" AUTOCOMPLETE="off"></div>
						</div>
	
					<% Else %>
					
						<div class="row row-line">
							<div class="col-xs-4"><strong>New Count:</strong></div>
							<div class="col-xs-3"><strong>Units</strong></div>
							<div class="col-xs-3">&nbsp;</div>
							<div class="col-xs-4">&nbsp;</div>
							<div class="col-xs-3"><input type="search" class="form-control" name="txtUnitsCounted" id="txtUnitsCounted" AUTOCOMPLETE="off"></div>
							<div class="col-xs-3">&nbsp;</div>
						</div>
						
					<% End If %>
						
					<% If prodCasePricing = "N" Then %>
					
						<% If prodUM <> "" Then 'If they did it by SKU, we can't accept a bin - we dont know if it is unit or case %>
							<div class="row row-line">
								<div class="col-xs-4"><strong>New Bin:</strong></div>
								<div class="col-xs-5"><input type="search" class="form-control" name="txtBinLocation" id="txtBinLocation" AUTOCOMPLETE="off"></div>
							</div>
						<% End If %>
					
					<% Else %>
					
						<% If prodUM <> "" Then 'If they did it by SKU, we can't accept a bin - we dont know if it is unit or case %>
						
							<div class="row row-line">
								<div class="col-xs-4"><strong>New Bin:</strong></div>
								<div class="col-xs-4"><strong>Case Bin</strong></div>
								<div class="col-xs-4"><strong>Unit Bin</strong></div>
								<div class="col-xs-4">&nbsp;</div>
								<div class="col-xs-4"><input type="text" class="form-control" name="txtCaseBinLocation" id="txtCaseBinLocation" AUTOCOMPLETE="off"></div>
								<div class="col-xs-4"><input type="text" class="form-control" name="txtUnitBinLocation" id="txtUnitBinLocation" AUTOCOMPLETE="off"></div>
							</div>	
							
						<% End If %>
					
					<% End If %>
						
					<div class="row row-line">
						<div class="col-xs-12"><button class="btn btn-primary btn-go btn-md">SUBMIT INVENTORY CHANGES</button></div>
					</div>
				</form>
				 
		        
		    </div>
	
	
	<% Else %>
	
		<div class="container-fluid">
		
		    <div class="row row-line">
		        <div class="col-xs-4"><strong>UPC Code:</strong></div>
				<div class="col-xs-8"><strong class="red"><%= UPCCode %></strong></div>
		    </div>
		    
		    <div class="row row-line">
		        <div class="col-xs-12"><strong class="red">[No Product Matched Entered UPC Code]</strong></div>
		    </div>		    
	
		    <div class="row row-line">
		        <div class="col-xs-12"><strong>Select Product To Assign UPC Code To:</strong></div>
		    </div>
		    		    
		    <div class="row row-line">
		        <div class="col-lg-12 col-md-12 col-sm-12 col-xs-12">
					<div id="scrollable-dropdown-menu">
					  <input class="typeahead" type="text" placeholder="Search By SKU or Description">
					</div>		     
		        </div>
		     </div>
		     
		    <% savedProdSKUFromLookup = MUV_Read("savedProdSKUFromLookup") %>
		    
		    <% If savedProdSKUFromLookup <> "" Then %>
				<div class="row row-line" id="yesSavedProdSKU">
					<div class="col-xs-3 pull-right" style="padding-left:0px"><button class="btn btn-info btn-go btn-md" id="btnClearTypeahead">CLEAR</button></div>
					<div class="col-xs-9 pull-right"><button class="btn btn-primary btn-go btn-md" id="btnHoldLastSKUEntered">SAME AS LAST SKU: <%= savedProdSKUFromLookup %></button></div>
				</div>
			<% Else %>
				<div class="row row-line" id="noSavedProdSKU">
					<div class="col-xs-3 pull-right" style="padding:left:0px"><button class="btn btn-info btn-go btn-md" id="btnClearTypeahead">CLEAR</button></div>
					<div class="col-xs-9 pull-right">&nbsp;</div>
				</div>
			<% End If %>
		     
		     <div class="row row-line" id="productUMInfo"></div>
		     
		     <input type="hidden" id="txtEnteredUPCCode" name="txtEnteredUPCCode" value="<%= UPCCode %>">
		     <input type="hidden" id="txtProdSKUSelected" name="txtProdSKUSelected">
		     
		     
		     <!--<div class="row row-line" id="postInfo"></div>-->
		     
		     <div class="row row-line">
		     	<div class="col-xs-12 pull-right"><button class="btn btn-success btn-go btn-md" id="btnAssignUPCToProduct">ASSIGN UPC <%= UPCCode %></button></div>
		     </div>
			     
		</div>
		
	<%End If

End If%>