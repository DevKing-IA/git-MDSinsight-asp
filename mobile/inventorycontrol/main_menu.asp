<!--#include file="../../inc/header-mobile.asp"-->

<style type="text/css">

body{
	margin:0;
	padding: 0;
}

	.general-heading{
		line-height: 1;
		padding-top: 5px;
		padding-bottom: 5px;
		margin-bottom: 25px;
		background-color: #343173;
		color: #fff;
		text-align: center;
		margin-top: 0px;
		font-size: 30px;
	}

	.general-button{
		font-size: 24px;
		border-bottom: 4px solid #3c8f3c;
		margin-bottom: 15px;
		border-radius: 0px !important;
	}
	

	 

	.general-image{
		max-width: 100%;
		height: auto;
	}

	.magnifier{
		max-height: 30px;
	}

	@media (max-width: 768px) {

		.mobile-col{
			padding-left: 2px;
			padding-right: 2px;
 		}

 		.mobile-col .label{
 			width: 100%;
 			display: block;
 			font-size: 16px;
 			font-weight: bold;
 			margin-top: 5px;
 			padding: 0px !important;
 			white-space: normal !important;
 		}
 		 


}

</style>       

<h1 class="general-heading" >Inventory</h1>

 
<!-- driver menu starts here !-->
<div class="container-fluid general-container">
	
 	
		<!-- three buttons -->
		<div class="row">

			<!-- UPC Lookup -->
			<div class="col-lg-4 col-md-4 col-sm-4 col-xs-4 mobile-col">
				<a href="UPCLookup/upclookup.asp"> 
			<button type="button" class="btn btn-success btn-block col-lg-12 general-button"> <img src="../../img/magnifier.png" class="general-image magnifier"> <span class="label">Inventory<br>Lookup<br>(Scan)</span></button>
		</a>
			</div>
			<!-- eof UPC Lookup -->

			<!-- 2nd Button -->
			<div class="col-lg-4 col-md-4 col-sm-4 col-xs-4 mobile-col">
				<a href="SKULookup/skulookup.asp"> 
			<button type="button" class="btn btn-success btn-block col-lg-12 general-button"><img src="../../img/magnifier.png" class="general-image magnifier"> <span class="label">Inventory<br>Lookup<br>(No Scan)</span></button>
		</a>
			</div>
			<!-- eof 2nd Button -->

			<!-- 3rd Button -->
			<div class="col-lg-4 col-md-4 col-sm-4 col-xs-4 mobile-col">
				<a href="AvailableLookup/availablelookup.asp"> 
			<button type="button" class="btn btn-success btn-block col-lg-12 general-button"><img src="../../img/magnifier.png" class="general-image magnifier"> <span class="label">Check SKU<br>Availability<br>(Scan)</span></button>
		</a>
			</div>
			<!-- eof 3rd Button -->


		</div>
		<!-- eof three buttons -->

		<!-- Logout -->
		<div class="row">
			<div class="col-lg-12 mobile-col">
				<a href="../../../logout.asp"> 
			<button type="button" class="btn btn-success   btn-block  col-lg-12 general-button">Logout</button>
		</a>
			</div>
		</div>
			<!-- eof Logout -->
 
	
</div>


<!--#include file="../../inc/footer-mobile.asp"-->