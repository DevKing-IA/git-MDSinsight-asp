<!--#include file="../inc/header-field-service-mobile.asp"-->
<!--#include file="../inc/InsightFuncs_service.asp"-->
<%
SelectedMemoNumber = Request.Form("txtTicketNumber")
dummy = MUV_WRITE("PHP_PAGE",BaseURL & "/fieldserviceADV/upload.php")
If FS_SignatureOptional() Then
	dummy = MUV_WRITE("SignatureOptional",1)
Else
	dummy = MUV_WRITE("SignatureOptional",0) 
End If
%>

<!-- signature pad !-->
<script src="<%= BaseURL %>js/signature/signature_pad.js"></script>
<!-- eof signature pad !-->


<SCRIPT LANGUAGE="JavaScript">
	<!--
     function validateCloseServiceTicketform()
    {
        if (document.frmCloseServiceMemo.selAssetID.value != "") {
      	  if (document.frmCloseServiceMemo.txtAssetTagNumber.value != "") {
            swal("You selected an asset from the list AND typed in an asset tag number. Only one or the other is permitted. Please clear one entry before submitting.");
            return false;
      	  }
        }
        
        if (document.frmCloseServiceMemo.selAssetID.value != "") {
      	  if (document.frmCloseServiceMemo.txtAssetLocation.value == "") {
            swal("You selected an asset from the list but did not specify the location. Please enter the location before submitting.");
            return false;
      	  }
        }
        
        if (document.frmCloseServiceMemo.txtAssetTagNumber.value != "") {
      	  if (document.frmCloseServiceMemo.txtAssetLocation.value == "") {
            swal("You entered an asset tag but did not specify the location. Please enter the location before submitting.");
            return false;
      	  }
        }

                
        return true;
    }
    
    // -->
 </SCRIPT>  
 


<style type="text/css">

	.fieldservice-heading{
		line-height: 1;
		padding-top: 5px;
		padding-bottom: 5px;
	}
	
	.btn-home{
		color: #fff;
		margin-top: -2px;
		margin-left: 5px;
		float: left;
		background: transparent;
		border: 0px;
		cursor: pointer;
 	}

	.input-lg::-webkit-input-placeholder, textarea::-webkit-input-placeholder {
	  color: #666;
	}
	.input-lg:-moz-placeholder, textarea:-moz-placeholder {
	  color: #666;
	}
	
	.checkboxes label{
		font-weight: normal;
		margin-right: 20px;
	}
	.close-service-client-output{
		text-align: left;
	}

	.btn-link {
	    font-weight: 500;
	    font-size: 1.2em;
	    color: #343173;
	    background-color: transparent;
	    white-space: normal;
	    padding: 0px;
	}
	
	.btn-link:hover {
	    color: #007bff;
	    text-decoration:none;
	    background-color: transparent;
	    border-color: transparent;
	}	
	.accordion-box{
		margin-bottom:15px;
	}
	
	.close-service-h4{
		text-align:center;
	}
	
	.list-group{
		margin:5px;
	}
	
	label.btn span {
	  font-size: 1.5em ;
	}
	

	input[type=checkbox], input[type=radio] {
	    box-sizing: border-box;
	    padding: 0;
	    display: none;
	}    
		
	label input[type="checkbox"] ~ i.far.fa-square{
	    color: #007bff;    
	    display: inline;
	}
	label input[type="checkbox"] ~ i.far.fa-check-square{
	    display: none;
	}
	label input[type="checkbox"]:checked ~ i.far.fa-square{
	    display: none;
	}
	label input[type="checkbox"]:checked ~ i.far.fa-check-square{
	    color: #28a745;    
	    display: inline;
	}
	label:hover input[type="checkbox"] ~ i.far {
		color: #28a745;
	}
	
	div[data-toggle=buttons] label.active{
	    color: #007bff;
	}
	
	div[data-toggle=buttons] label {
		display: inline-block;
		padding: 6px 12px;
		margin-bottom: 0;
		font-size: 14px;
		font-weight: normal;
		line-height: 2em;
		text-align: left;
		white-space: nowrap;
		vertical-align: top;
		cursor: pointer;
		background-color: none;
		border: 0px solid #007bff;
		border-radius: 3px;
		color: #007bff;
		-webkit-user-select: none;
		-moz-user-select: none;
		-ms-user-select: none;
		-o-user-select: none;
		user-select: none;
	}

	.complete{
		display: inline-block;
		margin-bottom: 0;
		font-size: 14px;
		font-weight: normal;
		line-height: 1.4em;
		text-align: left;
		white-space: normal;
		vertical-align: top;
		cursor: pointer;
		background-color: none;
		color: #28a745;
		-webkit-user-select: none;
		-moz-user-select: none;
		-ms-user-select: none;
		-o-user-select: none;
		user-select: none;
	}
	
	div[data-toggle=buttons] label:hover {
		color: #28a745;
	}
	
	div[data-toggle=buttons] label:active, div[data-toggle=buttons] label.active {
		-webkit-box-shadow: none;
		box-shadow: none;
	}
	
</style>

<h1 class="fieldservice-heading">
	<form method="post" action="onSite.asp" name="frmOnSiteH" id="frmOnSiteH">
			<input type='hidden' id='txtTicketNumber' name='txtTicketNumber' value='<%=SelectedMemoNumber%>'>		 
			<button type="button" onclick="document.forms['frmOnSiteH'].submit();" class="btn-home"><i class="fa fa-arrow-left"></i></button>
	</form>
	
	Close Ticket #<%=SelectedMemoNumber%>
</h1>
 

<form action="CloseService_submit.asp" method="POST" ENCTYPE="multipart/form-data" name="frmCloseServiceMemo" id="frmCloseServiceMemo" onsubmit="return validateCloseServiceTicketform();"> 

<% NumberOfMinutesInServiceDayVar = GetNumberOfMinutesInServiceDay() %>

<div class="list-group">

	<%

	MemoNumber = SelectedMemoNumber
	custNum = GetServiceTicketCust(SelectedMemoNumber)
	
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rsCust = Server.CreateObject("ADODB.Recordset")
	rsCust.CursorLocation = 3 

	SQL = "SELECT Name,Addr1,Addr2,City,CityStateZip,Phone,Contact FROM AR_Customer WHERE Custnum = '" & custNum & "'"
	Set rsCust = cnn8.Execute(SQL)
	If NOT rsCust.EOF Then
		custName = rsCust("Name")
		custAddr1 = rsCust("Addr1")
		custAddr2 = rsCust("Addr2")
		custCity = rsCust("City")
		custCityStateZip = rsCust("CityStateZip")
		custPhone = rsCust("Phone")
		custContact = rsCust("Contact")
	End If 
	%>
				
	<span class="list-group-item list-group-item-action flex-column align-items-start">
		<div class="d-flex w-100 justify-content-between">
			<h6 class="mb-1 font-weight-bold" style="font-size:1.1em;"><%= custName %></h6>
			
			<%
				elapsedMinutes = ServiceCallElapsedMinutesOpenTicket(MemoNumber)
				minutesInServiceDay = NumberOfMinutesInServiceDayVar
				
				If elapsedMinutes < 1 Then elapsedMinutes = 1 ' If it has been less than 1 minute, just show 1 anyway
				elapsedMinutesForSorting = elapsedMinutes
				elapsedString = ""
				elapsedDays = 	elapsedMinutes \ minutesInServiceDay
				If int(elapsedDays) > 0 Then
					elapsedMinutes = elapsedMinutes - (int(elapsedDays) * minutesInServiceDay)
					elapsedString = elapsedDays & "d "
				End If
				elapsedHours = elapsedMinutes \ 60
				If int(elapsedHours) > 0 Then 
					elapsedMinutes = elapsedMinutes - (int(elapsedHours) * 60)
					elapsedString = elapsedString  & elapsedHours & "h "
				End If
				If int(elapsedMinutes) > 0 Then
					elapsedString = elapsedString  & elapsedMinutes & "m"
				End If

			%>							
			<small><%= elapsedString %></small>

		</div>
									
		<small><%= custAddr1 %>&nbsp;<%= custAddr2 %>&nbsp;<%= custCity %></small>

		<h6 class="mb-1 mt-1">Ticket #<%= MemoNumber %>

				<% If TicketIsUrgent(MemoNumber) Then %>
					<span class="badge badge-danger badge-pill"><i class="fas fa-exclamation"></i></span>
				<% End If %>
				
				<% If filterChangeModuleOn() = True Then %>
					<% If TicketIsFilterChange(MemoNumber) Then %>
						<span class="badge badge-warning badge-pill">F</span>
					<% Else %>
						<span class="badge badge-info badge-pill badge-pill-icon-letter"><i class="fas fa-cog"></i></span>
					<% End If %>
				<% Else %>
					<span class="badge badge-info badge-pill badge-pill-icon-letter"><i class="fas fa-cog"></i></span>
				<% End If %>
		
		</h6>
		
		<!--<p class="mb-1"><%= GetTerm("Account") %>&nbsp;<%= custNum %></p>-->
		<small class="mb-2 d-block"><%= custContact %>&nbsp;<%= custPhone %></small>
		
		
	</span>

</div>


	<%
	
	SQL = "SELECT * FROM FS_ServiceMemosFilterInfo WHERE (ServiceTicketID = '" & MemoNumber & "')"
	Set rsCust = cnn8.Execute(SQL)
	
	If NOT rsCust.EOF Then
		%><div class="list-group"><%
		
		DO WHILE NOT rsCust.EOF
		
			 InternalRecordIdentifier = rsCust("InternalRecordIdentifier") 
			 ServiceTicketID = rsCust("ServiceTicketID")
			 CustFilterIntRecID = rsCust("CustFilterIntRecID") 
			 ICFilterIntRecID = rsCust("ICFilterIntRecID")
			 
			Set cnnFilterInfo = Server.CreateObject("ADODB.Connection")
			cnnFilterInfo.open (Session("ClientCnnString"))
			Set rsFilterInfo = Server.CreateObject("ADODB.Recordset")
			rsFilterInfo.CursorLocation = 3 
			
	
			SQLFilterInfo = "SELECT * FROM FS_CustomerFilters WHERE (InternalRecordIdentifier = " & CustFilterIntRecID & ")"
			Set rsFilterInfo = cnnFilterInfo.Execute(SQLFilterInfo)
			
			If NOT rsFilterInfo.EOF Then
				FilterLocation = rsFilterInfo("Notes")
			End If
			
			 
			%>
			<span class="list-group-item list-group-item-action flex-column align-items-start">

			
				<h6 class="mb-1 mt-1"><%= GetFilterIDByIntRecID(ICFilterIntRecID) & " - " & GetFilterDescByIntRecID(ICFilterIntRecID) %></h6>
				<small class="mb-2 d-block"><strong>Location</strong>: <%= FilterLocation %></small>

			    <div class="col-xs-12">
				    <div class="btn-group btn-group-vertical" data-toggle="buttons">
						<label class="btn" for="chkComplete<%= InternalRecordIdentifier %>">
						<%
						
						Set cnnFilterUpdate = Server.CreateObject("ADODB.Connection")
						cnnFilterUpdate.open (Session("ClientCnnString"))
						Set rsFilterUpdate = Server.CreateObject("ADODB.Recordset")
						rsFilterUpdate.CursorLocation = 3 
						
						'******************************************************************************
						'CHECK TO SEE IF THE FILTER IS ALREADY COMPLETED
									
						SQLFilterUpdate = "SELECT * FROM FS_ServiceMemosFilterInfo WHERE (InternalRecordIdentifier = " & InternalRecordIdentifier & ")"
						Set rsFilterUpdate = cnnFilterUpdate.Execute(SQLFilterUpdate)
						
						Completed = 0 
						
						If NOT rsFilterUpdate.EOF Then
							Completed = rsFilterUpdate("Completed")
							CompletedByUserNo = rsFilterUpdate("CompletedByUserNo")
							CompletedDate = rsFilterUpdate("CompletedDate")
						End If
						'******************************************************************************
						
						%>
						
						<% If Completed = 1 Then %>
	          				<span class="complete">Completed by <%= GetUserDisplayNameByUserNo(CompletedByUserNo) %> on <%= CompletedDate %></span>
	          			<% Else %>
	          				<input type="checkbox" name="chkComplete<%= InternalRecordIdentifier %>" id="chkComplete<%= InternalRecordIdentifier %>"><i class="far fa-square fa-2x"></i><i class="far fa-check-square fa-2x"></i> <span> Complete</span>
	          			<% End If %>
	        			</label>
	        		</div>
			    </div>
			    

			</span>
			<%
		
		rsCust.MoveNext
		Loop
		
		%></div><%
	End If 
	%>
				




<div class="container-fluid">

		<input type='hidden' id='txtTicketNumber' name='txtTicketNumber' value='<%=SelectedMemoNumber%>'>		 


		<!-- order notes !-->
		<h4 class="close-service-h4"><i class="fa fa-comments"></i> Service Notes </h4>
		<div class="col-lg-12 close-service-box">
			<textarea class="form-control  input-lg" rows="5" name="ServiceNotes" spellcheck="True" id="ServiceNotes"></textarea>
		</div>
		<!-- eof order notes !-->
		
		
		<% If SelectedMemoNumber <> "" Then %>
			
			<div id="accordion">
				  <div class="card">
				    <div class="card-header" id="headingOne">
				      <h5 class="mb-0">
				        <button class="btn btn-link" type="button" data-toggle="collapse" data-target="#collapseOne" aria-expanded="true" aria-controls="collapseOne">
				          <i class="fas fa-map-marker-edit"></i> Update Asset Location
				        </button>
				      </h5>
				    </div>
				
				    <div id="collapseOne" class="collapse" aria-labelledby="headingOne" data-parent="#accordion">
				      <div class="card-body">

					        <div class="col-lg-12">

						        <select name="selAssetID" id="selAssetID" class="form-control input-lg">
									<option value="">Tap to select from assets assigned to this account</option>
									<option value="noneselected">-- NONE or NOT FOUND, USE THE NUMBER FROM THE BOX BELOW --</option>
									<%	
									
									Set cnn8 = Server.CreateObject("ADODB.Connection")
									cnn8.open (Session("ClientCnnString"))
									Set rs = Server.CreateObject("ADODB.Recordset")
									rs.CursorLocation = 3 
										
								'	SQL = "SELECT assetNumber,description,serno FROM " & MUV_Read("SQL_Owner") & ".Assets WHERE CustAcctNum = " & GetServiceTicketCust(SelectedMemoNumber) &" ORDER BY assetTypeNo, assetNumber"
									
								'	set rs = cnn8.Execute (SQL)
								'	If not rs.EOF Then
							
								'		Do While Not rs.EOF
								'			tempAssetNum = rs("assetNumber")
								'			tempAssetDescription = rs("description")
								'			tempSerialNumber = rs("serno")
								'			
								'			If tempAssetNum = CurrentService_AssetNum Then
								'				strSelect =  "<option selected value='"& tempAssetNum &"'>"& tempAssetNum  & " -- SN:" &  tempSerialNumber & "-- " &  tempAssetDescription & "</option>"
								'			Else
								'				strSelect =  "<option value='"& tempAssetNum &"'>"& tempAssetNum  & " -- SN:" & tempSerialNumber & "-- " & tempAssetDescription & "</option>"
								'			End If
							'
							'				Response.Write(strSelect)
							'				rs.MoveNext
							'			Loop
										
										
							'		End If
									Set rs = Nothing
									Set Cnn8 = Nothing
								%>
							</select>
						</div> 

						<!-- asset tag number !-->
						<div class="col-lg-12 selectedhidden" id="noneselected" style="display:none;">
							<h4 class="close-service-h4">If not found enter the asset tag below</h4>
							<input type="text" class="form-control input-lg" name="txtAssetTagNumber" id="txtAssetTagNumber">
						</div>
						<!-- eof asset tag number !-->

						<!-- asset location !-->
						<div class="col-lg-12">
							<h4 class="close-service-h4"><i class="fa fa-map-marker"></i> Asset Location </h4>
							<input type="text" class="form-control input-lg" name="txtAssetLocation" id="txtAssetLocation">
						</div>
						<!-- eof asset location !-->
						
				      </div>
				    </div>
				  </div>
			  
			  
			  
				  <div class="card">
				    <div class="card-header" id="headingTwo">
				      <h5 class="mb-0">
				        <button class="btn btn-link collapsed" type="button" data-toggle="collapse" data-target="#collapseTwo" aria-expanded="false" aria-controls="collapseTwo">
				          <i class="fas fa-camera-alt"></i> Attach Pictures
				        </button>
				      </h5>
				    </div>
				    <div id="collapseTwo" class="collapse" aria-labelledby="headingTwo" data-parent="#accordion">
				      <div class="card-body" id="uploadPictures">
							Picture 1:<input type="file" capture="camera" accept="image/*" id="cameraInput1" name="cameraInput1"><br><br>
							Picture 2:<input type="file" capture="camera" accept="image/*" id="cameraInput2" name="cameraInput2"><br><br>
							Picture 3:<input type="file" capture="camera" accept="image/*" id="cameraInput3" name="cameraInput3"><br><br>
							Picture 4:<input type="file" capture="camera" accept="image/*" id="cameraInput4" name="cameraInput4"><br><br>
							Picture 5:<input type="file" capture="camera" accept="image/*" id="cameraInput5" name="cameraInput5"><br><br>
							Picture 6:<input type="file" capture="camera" accept="image/*" id="cameraInput6" name="cameraInput6"><br>
				      </div>
				    </div>
				  </div>
			  
			</div>			

		
			<!-- signature id !-->
			<div  id="signature-pad">
				<h4 class="close-service-h4"><i class="fa fa-hand-o-down"></i> Please sign in the box below</h4>
				<div class="col-lg-12 close-service-box">
					<div class="panel panel-default">
				        <div class="panel-body">
							<div>
								<canvas class="signature-canvas" ></canvas>
								<canvas id="buffer" style="display:none;"></canvas>
							</div>
							<div>
								<% If FS_SignatureOptional() <> True Then %>
									<div class="alert alert-info">*NOTE* Signature area can not be left blank</div>
								<%End If%>
									<input type="text" class="form-control input-lg" placeholder="Print Your Name Here" name="txtPrintedName" id="txtPrintedName">
									<br>
									<button data-action="clear" class="btn btn-info" name="Submit" value="Clear" id="clear">Clear Signature Area</button>
								</div>
							</div>
						</div>
					</div>
					<!-- eof signature pad !-->

					<button class="btn btn-info btn-block btn-lg close-buttons" name="Submit" value="Submit Service Memo" type="submit" data-action="save" id="btn-download"><i class="fa fa-upload"></i> Submit Service Memo</button>
			</div>
			<!-- eof signature id !-->
			<!-- eof cancel / submit buttons !-->
			

			<% End If %>  
			
	</form>	
	

	<form method="post" action="onSite.asp" name="frmOnSite" id="frmOnSite">
		<input type='hidden' id='txtTicketNumber' name='txtTicketNumber' value='<%=SelectedMemoNumber%>'>			
		<button type="button" onclick="document.forms['frmOnSite'].submit();" class="btn btn-warning btn-block btn-lg close-buttons"><i class="fa fa-times-circle-o"></i> Cancel</button>
	</form>
		

</div>
  
  
    
<script>

    $(document).ready(function () {
    


        // Handler for .ready() called.

        var wrapper = document.getElementById("signature-pad"),
        clearButton = wrapper.querySelector("[data-action=clear]"),
        saveButton = wrapper.querySelector("[data-action=save]"),
        canvas = wrapper.querySelector("canvas"),
        signaturePad;
        

        // Adjust canvas coordinate space taking into account pixel ratio,
        // to make it look crisp on mobile devices.
        // This also causes canvas to be cleared.
        function resizeCanvas() {
            var ratio = window.devicePixelRatio || 50;
            canvas.width = canvas.offsetWidth * ratio;
            canvas.height = canvas.offsetHeight * ratio;
            canvas.getContext("2d").scale(ratio, ratio);
        }


        window.onresize = resizeCanvas;
        resizeCanvas();
        
        signaturePad = new SignaturePad(canvas);
            
        var canvas = document.getElementById('canvas');
		var buffer = document.getElementById('buffer');

		
		window.onresize = function(event) {
	    var w = $(window).width(); //Using jQuery for easy multi browser support.
	    var h = $(window).height();
	    buffer.width = w;
	    buffer.height = h;
	    buffer.getContext('2d').drawImage(canvas, 0, 0);
	    canvas.width = w;
	    canvas.height = h;
	    canvas.getContext('2d').drawImage(buffer, 0, 0);
		}
            
            
		clearButton.addEventListener("click", function (event) {
			signaturePad.clear();
		});

		saveButton.addEventListener("click", function (event) {
		
		    if (signaturePad.isEmpty()) {

				
				if(<%=MUV_READ("SignatureOptional")%> == 0){
					swal("Please provide a signature.");
				  	event.preventDefault();
				}    

					
		    } else {
	                
				var ticketid = "<%= SelectedMemoNumber %>";
                    
				var dataURL = signaturePad.toDataURL("image/png");
	                    
				$.ajax({
					
					url:'<%= MUV_READ("PHP_PAGE") %>', 	
				    type:'POST', 
				    async: false,
				    data: { 
				           imgBase64: dataURL,
				           ticketid: ticketid,
				           seno: '<%= MUV_READ("SERNO") %>'
				           //seno: '<%=Session("SerNoToPass")%>'
				         }      
	
			});

          }
      });
  });

	var uri = 'api/signatures';

	function SaveImage(dataURL) {
		dataURL = dataURL.replace('data:image/png;base64,', '');
        var data = JSON.stringify(
                       {
                       value: dataURL
               });
                               
		var image = document.getElementById("canvas").toDataURL("image/png");
		image = image.substr(23, image.length);
                    }

        function onWebServiceFailed(result, status, error) {
            var errormsg = eval("(" + result.responseText + ")");
            alert(errormsg.Message);
        }
        
        //prevent page from refresh on clicking Clear
                
     $("#clear").click(function(e) {
  e.preventDefault();
});
    
    // eof prevent page from refresh on clicking Clear
         
</script>

 <!-- show content if NONE or NOT FOUND is selected !-->
<script>
	 $(function() {
        $('#selAssetID').change(function(){
            $('.selectedhidden').hide();
            $('#' + $(this).val()).show();
        });
    });
</script>
<!-- eof show content !-->
 
<!--#include file="../inc/footer-field-service-noTimeout.asp"-->