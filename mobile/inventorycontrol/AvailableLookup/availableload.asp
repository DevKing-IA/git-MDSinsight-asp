<!--#include file="../../../inc/InsightFuncs.asp"-->
<!--#include file="../../../inc/InsightFuncs_InventoryControl.asp"-->
<!--#include file="../../../inc/mail.asp"-->
<% 
ClientKeyForFileNames = MUV_READ("ClientKeyForFileNames")
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
					data: "action=ReturnAvailabilityInfoForProduct&prodSKU="+encodeURIComponent(prodSKU),
						success: function(msg){
							$("#productAvailabilityInfo").html(msg);
						}
				}) 
			  }
		    
		});
		
				
		$("#btnClearTypeahead").click(function() {
             $('.typeahead').typeahead('val', '');
             $('.typeahead').focus();
             $("#productAvailabilityInfo").html("");
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
			QtyAvailableUnits_Units = 0 ' Initially set it to ours in case the post fails
			QtyAvailableUnits_UnitsStatus = "NO BACKEND"
			
			'Post to the backend to try to get an up-to-the-minute on hand
			'Construct xml fields based on record
			xmlData = "<DATASTREAM>"
			xmlData = xmlData & "<IDENTITY>Pm8316wyc011</IDENTITY>"
					
			xmlData = xmlData & "<MODE>" & GetPOSTParams("BackendInventoryPostsMode") & "</MODE>"
					
			xmlData = xmlData & "  <RECORD_TYPE>INVENTORY</RECORD_TYPE>"
			xmlData = xmlData & "  <RECORD_SUBTYPE>QUERY_AVAILABLE</RECORD_SUBTYPE>"
					
			xmlData = xmlData & "<SERNO>" & MUV_READ("SERNO") & "</SERNO>"
					
						
			xmlData = xmlData & " <QUERY_AVAILABLE>"

			xmlData = xmlData & "        <PROD_ID>" & prodSKU & "</PROD_ID>"
			xmlData = xmlData & "        <RETURN_VALUE_UM>U</RETURN_VALUE_UM>"
			
			xmlData = xmlData & " </QUERY_AVAILABLE>"
				 
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
				emailbody="httpRequest.status returned " & httpRequest.status & " when posting <RECORD_TYPE>INVENTORY and <RECORD_SUBTYPE>QUERY AVAILABLE"& "<br><br>"
				emailBody = "httpRequest.responseText:" & httpRequest.responseText & "<br><br>"
				emailBody = emailBody & "PAGE: " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "<br>"
				emailBody = emailBody & "Posted to " & GetPOSTParams("BackendInventoryPostsURL") & "<br><br>"
				emailBody = emailBody & "POSTED DATA:<br>" & xmlDataForDisp & "<br><br>"
				emailBody = emailBody & "SERNO: " & MUV_READ("SERNO") & "<br>"
				SendMail "mailsender@" & maildomain ,"support@mdsinsight.com",SERNO & " Post Error Inventory Query Available",emailBody, "Inventory API", "Inventory API"
				Description = emailBody 
				Write_API_AuditLog_Entry Request.ServerVariables("SERVER_NAME"),Description,GetPOSTParams("BackendInventoryPostsMode"),MUV_READ("SERNO"),MUV_READ("SERNO"),"Inventory API"
			End If
		
			If httpRequest.status = 200 THEN 
			
				If IsNumeric(httpRequest.responseText) Then ' Success
			
					emailbody="httpRequest.status returned " & httpRequest.status & " when posting <RECORD_TYPE>INVENTORY and <RECORD_SUBTYPE>QUERY AVAILABLE<"& "<br><br>"
					emailBody = emailBody & "httpRequest.responseText:" & httpRequest.responseText & "<br><br>"
					emailBody = emailBody & "PAGE: " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "<br>"
					emailBody = emailBody & "Posted to " & GetPOSTParams("BackendInventoryPostsURL") & "<br><br>"
					emailBody = emailBody & "POSTED DATA:<br>" & xmlDataForDisp & "<br><br>"
					emailBody = emailBody & "SERNO: " & MUV_READ("SERNO") & "<br>"
					SendMail "mailsender@" & maildomain ,"insight@ocsaccess.com", MUV_READ("SERNO") & " Good Post Inventory Query Available",emailBody, "Inventory API", "Inventory API"
					
					Call Write_API_AuditLog_Entry(Identity ,emailBody,GetPOSTParams("BackendInventoryPostsMode"),"availableload.asp")
					
					QtyAvailableUnits_Units = httpRequest.responseText
					QtyAvailableUnits_UnitsStatus = "LIVE"
					
					
				Else
					'FAILURE
					emailbody="httpRequest.status returned " & httpRequest.status & " when posting <RECORD_TYPE>INVENTORY and <RECORD_SUBTYPE>QUERY AVAILABLE<"& "<br><br>"
					emailBody = emailBody & "httpRequest.responseText:" & httpRequest.responseText & "<br><br>"
					emailBody = emailBody & "PAGE: " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "<br>"
					emailBody = emailBody & "Posted to " & GetPOSTParams("BackendInventoryPostsURL") & "<br><br>"
					emailBody = emailBody & "POSTED DATA:<br>" & xmlDataForDisp & "<br><br>"
					emailBody = emailBody & "SERNO: " & MUV_READ("SERNO") & "<br>"
					SendMail "mailsender@" & maildomain ,"support@mdsinsight.com",MUV_READ("SERNO") & " Post Error Inventory Query Available",emailBody, "Inventory API", "Inventory API"
				
					Call Write_API_AuditLog_Entry(Identity ,emailBody ,GetPOSTParams("BackendInventoryPostsMode"),"availableload.asp")
					
				End If
				
			Else
			
					'FAILURE
					emailbody="NON 200 Response - httpRequest.status returned " & httpRequest.status & " when posting <RECORD_TYPE>INVENTORY and <RECORD_SUBTYPE>QUERY AVAILABLE<"& "<br><br>"
					emailBody = emailBody & "httpRequest.responseText:" & httpRequest.responseText & "<br><br>"
					emailBody = emailBody & "PAGE: " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "<br>"
					emailBody = emailBody & "Posted to " & GetPOSTParams("BackendInventoryPostsURL") & "<br><br>"
					emailBody = emailBody & "POSTED DATA:<br>" & xmlDataForDisp & "<br><br>"
					emailBody = emailBody & "SERNO: " & MUV_READ("SERNO") & "<br>"
					SendMail "mailsender@" & maildomain ,"support@mdsinsight.com",MUV_READ("SERNO") & " Post Error Inventory Query Available",emailBody, "Inventory API", "Inventory API"
				
					Call Write_API_AuditLog_Entry(Identity ,emailBody,GetPOSTParams("BackendInventoryPostsMode"),"availableload.asp")
		
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
					prodUnitBin = rsprodLookup("prodUnitBin")
					prodCaseBin = rsprodLookup("prodCaseBin")	
				End If
			End If
		
		End If
		
		
		Set rsprodLookup = Nothing
		cnnprodLookup.Close
		Set cnnprodLookup = Nothing
	
		prodImage = GetProdImage(prodSKU)%>
	
		<div class="container-fluid inventory-upc-container">	

			<div class="row">
				<div class="col-xs-12 text-center"><strong style="font-size:30px;">AVAILABLE UNITS:</strong></div>
			</div>
			
			<div class="row row-line">
				<div class="col-xs-12 text-center" style="padding-bottom:10px;">
					<% If QtyAvailableUnits_UnitsStatus = "NO BACKEND" Then %>
						<strong class="red" style="font-size:40px;">No Backend Connection</strong>
					<% Else %>
						<strong class="green" style="font-size:80px;"><%= QtyAvailable_Units %></strong>
					<% End If %>
				</div>
			</div>
			
	
		    <div class="row">
		    
				<div class="col-lg-7 col-md-7 col-sm-7 col-xs-7">
					<div class="container-fluid" style="padding-left:0; padding-right:0">
						<div class="row">
							<div class="col-xs-12" style="padding-bottom:10px;"><strong>UPC Code:</strong>&nbsp;<strong class="red"><%= UPCCode %></strong></div>
							<div class="col-xs-12" style="padding-bottom:10px;"><strong>Product ID:</strong>&nbsp;<strong class="red"><%= prodSKU %></strong></div>
							
							<input type="hidden" id="txtUPCCodeToPass" name="txtUPCCodeToPass" value="<%= UPCCode %>">
							
							
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
							<div class="col-xs-6 text-center"><%= QtyAvailableUnits_UnitsStatus %></div>
							<div class="col-xs-6 text-center"><strong class="red"><%= QtyAvailableUnits_Units %></strong></div>				
						</div>
						
		 		<% Else %>
		 		
		 				<div class="row" style="padding-bottom:5px;">
							<div class="col-xs-3"><strong>On Hand</strong></div>
							<div class="col-xs-3 text-center"><strong>Tot Units</strong></div>	
							<div class="col-xs-3 text-center"><strong>Cases</strong></div>
							<div class="col-xs-3 text-center"><strong>Units</strong></div>
						</div>
						
						<div class="row row-info">
						
							<div class="col-xs-3 text-center"><%= QtyAvailableUnits_UnitsStatus %></div>

							<div class="col-xs-3 text-center">
								<strong class="red"><%= QtyAvailableUnits_Units %></strong>
							</div>	
							
							<% If prodCaseConversionFactor <> "" Then %>
								<div class="col-xs-3 text-center"><strong class="red"><%= Int(QtyAvailableUnits_Units / cInt(prodCaseConversionFactor))  %></strong></div>
							<% Else %>
								<div class="col-xs-3 text-center">---</div>
							<% End If %>
							
							<% If QtyAvailableUnits_Units Mod cint(prodCaseConversionFactor) = 0 Then %>	
								<div class="col-xs-3 text-center"><strong class="red">---</strong></div>
							<% Else %>
								<div class="col-xs-3 text-center"><strong class="red"><%=  QtyAvailableUnits_Units Mod cInt(prodCaseConversionFactor) %></strong></div>	
							<% End If %>	
							
								
						</div>
		 		<% End If %>
				
				</div>				 
		        
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
		        <div class="col-xs-12"><strong>Select Product To Search For Availability:</strong></div>
		    </div>
		    		    
		    <div class="row row-line">
		        <div class="col-lg-12 col-md-12 col-sm-12 col-xs-12">
					<div id="scrollable-dropdown-menu">
					  <input class="typeahead" type="text" placeholder="Search By SKU or Description">
					</div>		     
		        </div>
		     </div>
		     
			<div class="row row-line" id="noSavedProdSKU">
				<div class="col-xs-3 pull-right" style="padding:left:0px"><button class="btn btn-info btn-go btn-md" id="btnClearTypeahead">CLEAR</button></div>
				<div class="col-xs-9 pull-right">&nbsp;</div>
			</div>
		     
		     <div class="row row-line" id="productAvailabilityInfo"></div>
		     
		     <input type="hidden" id="txtEnteredUPCCode" name="txtEnteredUPCCode" value="<%= UPCCode %>">
		     <input type="hidden" id="txtProdSKUSelected" name="txtProdSKUSelected">
		     
		     
		     <!--<div class="row row-line" id="postInfo"></div>-->
			     
		</div>
		
	<%End If

End If%>