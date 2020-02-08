<!--#include file="inc/header-tech-and-driver.asp"-->

<%SelectedMemoNumber = Request.Form("txtTicketNumber")%>

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
.ticket-details{
	margin-bottom: 15px;
} 

.accordion-box{
margin-bottom:15px;
}
</style>

<h1 class="fieldservice-heading" ><form method="post" action="onSite.asp" name="frmOnSiteH" id="frmOnSiteH">
		<input type='hidden' id='txtTicketNumber' name='txtTicketNumber' value='<%=SelectedMemoNumber%>'>		 
		<button type="button" onclick="document.forms['frmOnSiteH'].submit();" class="btn-home"><i class="fa fa-arrow-left"></i>  </button>
</form>
Close Ticket # <%=SelectedMemoNumber%></h1>
 

<div class="container-fluid fieldservice-container">
	<div class="row">
 
		<form action="CloseService_submit.asp" method="POST" ENCTYPE="multipart/form-data" name="frmCloseServiceMemo" id="frmCloseServiceMemo" onsubmit="return validateCloseServiceTicketform();"> 

		<input type='hidden' id='txtTicketNumber' name='txtTicketNumber' value='<%=SelectedMemoNumber%>'>		 


			<!-- order notes !-->
			<div class="row row-rest">
				<h4 class="close-service-h4"><i class="fa fa-comments"></i> Service Notes </h4>
				<div class="col-lg-12 close-service-box">
					<textarea class="form-control  input-lg" rows="5" name="ServiceNotes"  spellcheck="True" id="ServiceNotes" ></textarea>
				</div>
			</div>
			<!-- eof order notes !-->

			<!-- assets accordion !-->
			<div class="container-fluid accordion-box"  id="theaccordion"  >
				<div class="row row-rest">
					<% If SelectedMemoNumber <> "" Then %>		
						<div class="panel-group panel-group-asset" id="accordion" role="tablist" aria-multiselectable="true">
 
	  						<div class="panel panel-default">
		   
								<!-- title !-->
    							<div class="panel-heading" role="tab" id="headingOne">
       								<h4 class="panel-title close-service-h4-panel">
	       								<a role="button" data-toggle="collapse"   data-parent="#accordion" href="#collapseOne"  aria-expanded="false" aria-controls="collapseOne">
	       									<i class="fa fa-plus-square-o"></i> Update Asset Location
	       								</a>
      								</h4>
								</div>

								<div id="collapseOne" class="panel-collapse collapse" role="tabpanel" aria-labelledby="headingOne">
									<div class="panel-body">
								        <div class="col-lg-12">

									        <select name="selAssetID" id="selAssetID" class="form-control input-lg">
												<option value="">Tap to select from assets assigned to this account</option>
												<option value="noneselected">-- NONE or NOT FOUND, USE THE NUMBER FROM THE BOX BELOW --</option>
												<%	
												
												Set cnn8 = Server.CreateObject("ADODB.Connection")
												cnn8.open (Session("ClientCnnString"))
												Set rs = Server.CreateObject("ADODB.Recordset")
												rs.CursorLocation = 3 
													
												SQL = "SELECT assetNumber,description,serno FROM " & MUV_Read("SQL_Owner") & ".Assets WHERE CustAcctNum = " & GetServiceTicketCust(SelectedMemoNumber) &" ORDER BY assetTypeNo, assetNumber"
												
												set rs = cnn8.Execute (SQL)
												If not rs.EOF Then
										
													Do While Not rs.EOF
														tempAssetNum = rs("assetNumber")
														tempAssetDescription = rs("description")
														tempSerialNumber = rs("serno")
														
														If tempAssetNum = CurrentService_AssetNum Then
															strSelect =  "<option selected value='"& tempAssetNum &"'>"& tempAssetNum  & " -- SN:" &  tempSerialNumber & "-- " &  tempAssetDescription & "</option>"
														Else
															strSelect =  "<option value='"& tempAssetNum &"'>"& tempAssetNum  & " -- SN:" & tempSerialNumber & "-- " & tempAssetDescription & "</option>"
														End If
										
														Response.Write(strSelect)
														rs.MoveNext
													Loop
													
													
												End If
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
							<!-- eof content !-->
						</div>
					</div>
				</div>
			</div>
			<!-- eof assets accordion !-->

			<!-- cancel / submit buttons !-->
			<div class="container-fluid">
				<div class="row">

					<!-- left col with pictures !-->
 						<!-- pictures accordion !-->
					    <div class="panel-group panel-group-asset pictures" id="accordion" role="tablist" aria-multiselectable="true">
							<div class="panel panel-default">
								<!-- title !-->
								<div class="panel-heading " role="tab" >
									<h4 class="panel-title close-service-h4-panel">
										<a role="button" data-toggle="collapse"   data-parent="#accordion" href="#uploadPictures"  aria-expanded="false" aria-controls="uploadPictures">
											<i class="fa fa-plus-square-o"></i> Attach Pictures
										</a>
									</h4>
								</div>
								<!-- eof title !-->

								<!-- content !-->
								<div id="uploadPictures" class="panel-collapse collapse " role="tabpanel" aria-labelledby="pictures">
									<div class="panel-body">
										Picture 1:<input type="file" capture="camera" accept="image/*" id="cameraInput1" name="cameraInput1"><br>
										Picture 2:<input type="file" capture="camera" accept="image/*" id="cameraInput2" name="cameraInput2"><br>
										Picture 3:<input type="file" capture="camera" accept="image/*" id="cameraInput3" name="cameraInput3"><br>
										Picture 4:<input type="file" capture="camera" accept="image/*" id="cameraInput4" name="cameraInput4"><br>
										Picture 5:<input type="file" capture="camera" accept="image/*" id="cameraInput5" name="cameraInput5"><br>
										Picture 6:<input type="file" capture="camera" accept="image/*" id="cameraInput6" name="cameraInput6"><br>
									</div>
								</div>
								<!-- eof content !-->
							</div>
							</div>
 						<!-- eof pictures accordion !-->
					</div>
				</div>

		
				<!-- signature pad !-->
				
				<div class="row">
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
										<div class="alert alert-info">*NOTE* Signature area can not be left blank</div>
 											<input type="text" class="form-control input-lg" placeholder="Print Your Name Here" name="txtPrintedName" id="txtPrintedName">
											<br>
											<button data-action="clear" class="btn btn-info" name="Submit" value="Clear" id="clear" >Clear Signature Area</button>
										</div>
									</div>
								</div>
							</div>
							<!-- eof signature pad !-->

  								<button   class="btn btn-info btn-block  btn-lg close-buttons" name="Submit" value="Submit Service Memo" type="submit"   data-action="save" id="btn-download" onclick="uploadEx()" ><i class="fa fa-upload"></i> Submit Service Memo</button>
 						</div>
					</div>
					<!-- eof signature id !-->
					<!-- eof cancel / submit buttons !-->
				<% End If%>
			</div>
		</form>	
		
		<div class="row">
		<form method="post" action="onSite.asp" name="frmOnSite" id="frmOnSite">
				<input type='hidden' id='txtTicketNumber' name='txtTicketNumber' value='<%=SelectedMemoNumber%>'>
<div class="row">				
				<button type="button" onclick="document.forms['frmOnSite'].submit();" class="btn btn-warning btn-block  btn-lg close-buttons"><i class="fa fa-times-circle-o"></i> Cancel</button>
				</div>
		</form>
		</div>
		
	</div>
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
	                
                    swal("Please provide signature first.");
                  event.preventDefault();
                  
                } else {
	                
	                var ticketid = "<%= SelectedMemoNumber %>";
	             
	                
	               // var myMessage = "<%= SelectedMemoNumber %>";  
	                
	                    // save signature to server 
	                    
	                    var dataURL = signaturePad.toDataURL("image/png");
	                    
	                    
$.ajax({
	
		url:'http://dev.mdsinsight.com//mobile/TechAndDriver/upload.php', 
			
    type:'POST', 
    async: false,
    data: { 
           imgBase64: dataURL,
           ticketid: ticketid,
           seno: '<%=Session("ClientID")%>'
           
         }      
	
});

  
 // eof save signature to server


 
                     //window.open(signaturePad.toDataURL());
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
 
<!--#include file="inc/footer-tech-and-driver.asp"-->