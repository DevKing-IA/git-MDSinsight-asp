<div class="waitdiv d-none" style="position: fixed;z-index: 999999999; top: 0px; left: 0px; width: 100%; height:80%; background-color:transparent; text-align: center; padding-top: 20%; filter: alpha(opacity=0); opacity:0; "></div>
	<div id="waitdiv" class="waitdiv d-none small" style="padding-bottom: 90px;text-align: center; vertical-align:middle;padding-top:50px;background-color:#ebebeb;width:300px;height:100px;margin: 0 auto; top:40%; left:40%;position:absolute;-webkit-box-shadow: 0 5px 10px rgba(0, 0, 0, 0.2); -moz-box-shadow: 0 5px 10px rgba(0, 0, 0, 0.2); box-shadow: 0 5px 10px rgba(0, 0, 0, 0.2); z-index:999999999;">
		<img src="/img/loading_gray.gif" alt="" /><br /><span id="waitmsg">Loading Filter Change Information.</span> <br />Please wait ...
</div>


<div class="modal fade" id="confirmArchiveCustomer" tabindex="-1" role="dialog">

	<div class="modal-dialog" role="document">
						
		<div class="modal-content">	
		
			<div class="modal-header">
				<button type="button" class="close" data-dismiss="modal" aria-label="Close"><span aria-hidden="true">&times;</span></button>
				<h4 class="modal-title"></h4>
			</div>
			
			<div class="modal-body">
			</div>
			
			<div class="modal-footer">
				<button type="button" class="btn btn-default" data-dismiss="modal">Cancel</button>
				<button type="button" class="btn btn-primary" OnClick="">Yes, make inactive.</button>
			</div>
			
		</div>
		<!-- eof modal content !-->
	</div>
	<!-- eof modal dialog !-->
</div>


<script language="Javascript">

	var values=[];
	var currentCustomerID="";
		///*************************** datepicker area
			
	function toExclude(obj) {
		
		currentCustomerID=$(obj).attr("data-id");
		$("#confirmArchiveCustomer .modal-title").html("Please Confirm Marking As Inactive");
		$("#confirmArchiveCustomer .modal-body").html("Are you sure to mark this customer as inactive?");
		$("#confirmArchiveCustomer .modal-footer .btn.btn-primary").attr("onclick","javascript:doArchiveCustomer();");
		$("#confirmArchiveCustomer").modal("show");
	}
	
	function doArchiveCustomer() {
		if (currentCustomerID.length>0) {
		
			$.ajax({
				type:"GET",
				url: "archiveCustomer.asp",
				cache: false,
				data: "customerID="+currentCustomerID,
				success: function(response) {
					location.reload();
				},
				error:function(response) {
				
				},
				complete:function(response) {
				
				}
			});
		
		}
	}


</script>			
			
<style>
	.keepcenter {
		text-align: center;
	}
</style>

<!-- eof modal !-->	
