
<!--#include file="../../../inc/header-deliveryboard-drivers-mobile.asp"-->
<!--#include file="../../../inc/InSightFuncs_Routing.asp"-->
<script src="<%= BaseURL %>js/signature/signature_pad.js"></script>

<%
IvsNum = Request.QueryString("i")
CustNum = Request.QueryString("c")
Status = Request.QueryString("s")

dummy = MUV_WRITE("PHP_PAGE_DEL",BaseURL & "/mobile/drivers/deliveryboard/upload.php")


%>



<style type="text/css">
.input-lg::-webkit-input-placeholder, textarea::-webkit-input-placeholder {
  color: #666;
}
.input-lg:-moz-placeholder, textarea:-moz-placeholder {
  color: #666;
}

.alert{
	padding: 0px;
	margin: 0px;
}

 
.ac-container{
	width: 100%;
 	text-align: left;
  }
.ac-container label{
	margin-top: 20px;
	float: left;
	width: 100%;
	font-family:  Arial, sans-serif;
	padding: 5px 20px;
	position: relative;
	z-index: 20;
	display: block;
 	cursor: pointer;
	color: #777;
	text-shadow: 1px 1px 1px rgba(255,255,255,0.8);
	line-height: 33px;
	font-size: 14px;
	background: #ffffff;
	background: -moz-linear-gradient(top, #ffffff 1%, #eaeaea 100%);
	background: -webkit-gradient(linear, left top, left bottom, color-stop(1%,#ffffff), color-stop(100%,#eaeaea));
	background: -webkit-linear-gradient(top, #ffffff 1%,#eaeaea 100%);
	background: -o-linear-gradient(top, #ffffff 1%,#eaeaea 100%);
	background: -ms-linear-gradient(top, #ffffff 1%,#eaeaea 100%);
	background: linear-gradient(top, #ffffff 1%,#eaeaea 100%);
	filter: progid:DXImageTransform.Microsoft.gradient( startColorstr='#ffffff', endColorstr='#eaeaea',GradientType=0 );
	box-shadow: 
		0px 0px 0px 1px rgba(155,155,155,0.3), 
		1px 0px 0px 0px rgba(255,255,255,0.9) inset, 
		0px 2px 2px rgba(0,0,0,0.1);
}
.ac-container label:hover{
	background: #fff;
}
.ac-container input:checked + label,
.ac-container input:checked + label:hover{
	background: #c6e1ec;
	color: #3d7489;
	text-shadow: 0px 1px 1px rgba(255,255,255, 0.6);
	box-shadow: 
		0px 0px 0px 1px rgba(155,155,155,0.3), 
		0px 2px 2px rgba(0,0,0,0.1);
}
.ac-container label:hover:after,
.ac-container input:checked + label:hover:after{
	content: '';
	position: absolute;
	width: 24px;
	height: 24px;
	right: 13px;
	top: 7px;
	background: transparent url(../../../img/accordion/arrow_down.png) no-repeat center center;	
}
.ac-container input:checked + label:hover:after{
	background-image: url(../../../img/accordion/arrow_up.png);
}
.ac-container input{
	display: none;
}
.ac-container article{
	background: rgba(255, 255, 255, 0.5);
	margin-top: -3px;
	overflow: hidden;
 	position: relative;
	z-index: 10;
	-webkit-transition: height 0.3s ease-in-out, box-shadow 0.6s linear;
	-moz-transition: height 0.3s ease-in-out, box-shadow 0.6s linear;
	-o-transition: height 0.3s ease-in-out, box-shadow 0.6s linear;
	-ms-transition: height 0.3s ease-in-out, box-shadow 0.6s linear;
	transition: height 0.3s ease-in-out, box-shadow 0.6s linear;
 }
.ac-container article p{
	font-style: italic;
	color: #777;
	line-height: 23px;
	font-size: 14px;
	padding: 20px;
	text-shadow: 1px 1px 1px rgba(255,255,255,0.8);
}
.ac-container input:checked ~ article{
	-webkit-transition: height 0.5s ease-in-out, box-shadow 0.1s linear;
	-moz-transition: height 0.5s ease-in-out, box-shadow 0.1s linear;
	-o-transition: height 0.5s ease-in-out, box-shadow 0.1s linear;
	-ms-transition: height 0.5s ease-in-out, box-shadow 0.1s linear;
	transition: height 0.5s ease-in-out, box-shadow 0.1s linear;
	box-shadow: 0px 0px 0px 1px rgba(155,155,155,0.3);
}
.ac-container input:checked ~ article.ac-small{
padding:10px 20px 30px 20px;
width: 100%;
float: left;
background: #fff;
display: block;
 }

article.ac-small{
	display: none;
}

h3{
	margin-top: 0px;
}
 
 

.row-line{
	margin-bottom:20px;
}

.alert-warning{
	padding-top:15px;
	padding-bottom:10px;
	margin-bottom:10px;
	width:100%;
	float:left;
}


.fieldservice-heading{
	font-size:14px;
	font-weight:bold;
}

.fieldservice-heading-h3{
    background-color: #ddd;
    color: #222;
    text-align: center;
    margin-top: 0px;
    padding-top: 15px;
    padding-bottom: 15px;
    font-size: 20px;
 }
 
 .back-arrow{
	color: #fff;
	text-decoration: none;
	margin-left:5px;
} 

.back-arrow:hover{
	color:#ccc;
}
</style>
 

<h1 class="fieldservice-heading"><a class="back-arrow pull-left" href="main.asp" role="button"><i class="fa fa-arrow-left" aria-hidden="true"></i></a> Driver Comments</h1>

<div class="container-fluid">

	<strong>
		<%=GetTerm("Customer")%>:&nbsp;<%=GetCustNameByCustNum(CustNum)%><br>
		<% If IvsNum <> "" Then %>
		Invoice:
		<%Else%>
		Invoice(s):
		<%End If
		
		If IvsNum <> "" Then 
			SQL = "SELECT IvsNum FROM RT_DeliveryBoard WHERE IvsNum = '" & IvsNum & "'"
		Else
			SQL = "SELECT IvsNum FROM RT_DeliveryBoard WHERE CustNum = '" & CustNum & "'"
		End If
		
		Set cnn8 = Server.CreateObject("ADODB.Connection")
		cnn8.open (Session("ClientCnnString"))
		Set rsCust = Server.CreateObject("ADODB.Recordset")
		rsCust.CursorLocation = 3 
		Set rsCust = cnn8.Execute(SQL)
		
		x = x +1
		
		If NOT rsCust.EOF Then
		
			Do While NOT rsCust.EOF
				
				If x > 1 Then Response.Write(", ")
				
				If x > 4 Then
					Response.Write("<BR>")
					x = 1
				End IF
				
				Response.Write(rsCust("IvsNum"))
			
				rsCust.movenext
			Loop
			
		End IF 
		
		Set rsCust = Nothing
		cnn8.Close
		Set cnn8 = Nothing
		%>
	</strong>
	
	<h2 class="fieldservice-heading-h3">Please enter your comments regarding this stop. If there are no comments, simply leave the box empty.</h2>

	<div class="alert-warning">

		<% If Request.QueryString("i") <> "" Then
			' By Invoice Number
			If Status="d" Then
				Response.Write("<form method='post' action='tap_delivered_invoice.asp' name='frmCustoemr' id='frmCustoemr'>")
			Else
				Response.Write("<form method='post' action='tap_no_delivery_invoice.asp' name='frmCustoemr' id='frmCustoemr'>")
			End If
			Response.Write("<input type='hidden' id='txtCustNum' name='txtCustNum' value='" & CustNum & "'>")
			Response.Write("<input type='hidden' id='txtIvsNum' name='txtIvsNum' value='" & IvsNum & "'>")
		Else		
			'By Customer Number
			If Status="d" Then
				Response.Write("<form method='post' action='tap_delivered_customer.asp' name='frmCustoemr' id='frmCustoemr'>")
			Else
				Response.Write("<form method='post' action='tap_no_delivery_customer.asp' name='frmCustoemr' id='frmCustoemr'>")
			End If
			Response.Write("<input type='hidden' id='txtCustNum' name='txtCustNum' value='" & CustNum & "'>")
		End If %>
		
			<!-- comments -->
			<div class="col-lg-12"> 
				<textarea class="form-control row-line" rows="5" id="txtdriverComments" name="txtdriverComments" ></textarea>
			</div>
			 <!-- eof comments -->

			<!-- clear / continue buttons -->
			<div class="col-lg-6 col-md-6 col-sm-6 col-xs-6"><p><button type="button" class="btn btn-warning" id="btnClear" name="btnClear">Clear Box &amp; Start Over</button></p></div>
			<!-- eof clear / continue buttons -->

	</div>

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
								<input type="text" class="form-control input-lg" placeholder="Print Your Name Here" name="txtPrintedName" id="txtPrintedName">
								<br>
								<button data-action="clear" class="btn btn-info" name="Submit" value="Clear" id="clear">Clear Signature Area</button>
						</div>
					</div>
				</div>
			</div>
				<!-- eof signature pad !-->
		<div class="col-lg-6 col-md-6 col-sm-6 col-xs-6"><p align="center"><button name="Submit" type="submit" data-action="save" id="btn-download" class="btn btn-info">Complete Delivery</button></p></div>				
		</div>
	</div>
		<!-- eof signature id !-->
		<!-- eof cancel / submit buttons !-->

</div>





</form>
			
</div>
 	 
<script type="text/javascript">

    $(document).ready(function () {
    

   $('#btnClear').click( function () {
         $('#txtdriverComments').val(""); 
   });


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
				
				if(0 == 0){
					swal("Please provide a signature.");
				  	event.preventDefault();
				}    

					
		    } else {
	                
	                
				var ticketid = "<%= CustNum  %>";
                    
				var dataURL = signaturePad.toDataURL("image/png");

//alert('<%= MUV_READ("PHP_PAGE_DEL") %>');			
//alert('<%= MUV_READ("SERNO") %>');	
alert('<%= CustNum  %>');	
	                    
				$.ajax({
					
					url:'<%= MUV_READ("PHP_PAGE_DEL") %>', 	
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
	
		alert('vd');	
	
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

 
<!--#include file="../../../inc/footer-field-service-noTimeout.asp"-->