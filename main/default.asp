<!--#include file="../inc/header.asp"-->


<!--#include file="service_call_activity_chart_data.asp"-->
<!--#include file="user_activity_chart_data.asp"-->
<!--#include file="daily_api_activity_summary_chart_data.asp"-->

<%
dummy = MUV_Remove("PSoftAjax")
dummy = MUV_Remove("PSoftStartDate")
dummy = MUV_Remove("PSoftEndDate")		
dummy = MUV_Remove("PSoftSelectedPeriod")		
dummy = MUV_Remove("PSoftSkipZeroDollar")		
dummy = MUV_Remove("PSoftSkipLessThanZero")		
dummy = MUV_Remove("PSoftIncludedType")
dummy = MUV_Remove("PSoftCustomOrPredefined")
dummy = MUV_Remove("PSoftAccount")
dummy = MUV_Remove("PSoftSkipLessThanZeroLines")
dummy = MUV_Remove("PSoftDueDateDays")
dummy = MUV_Remove("PSoftDueDateSingleDate")
dummy = MUV_Remove("PSoftDoNotShowDueDate")
dummy = MUV_Remove("PSofttypeOfAccounts")
dummy = MUV_Remove("PSoftChain")
dummy = MUV_Remove("PSofttxtDueDate")
dummy = MUV_Remove("PSoftselDueDate")
%>

<% ' See if there is a login landing page
	LandingPage = MUV_ReadAndRemove("LoginLandingPageURL") 
	If Len(LandingPage) > 1 Then Response.Redirect(LandingPage)		
%>
<style type="text/css">
	
	#map{
		width: 100%;
		height: 350px;
	}
	
 	.first-line{
		margin-top:30px;
	}
	.row-sections > div {
     padding: 20px;
    border: 1px solid #eaeaea;
	margin:0px -1px -1px 0px;
 }
 
 .row-sections > div h2{
	 line-height: 1;
	 margin-top: 0px;
	 margin-bottom: 10px;
 }
 
 .row-eq-height {
  display: -webkit-box;
  display: -webkit-flex;
  display: -ms-flexbox;
  display:         flex;
}

@media (max-width: 480px) {
	
	.row-eq-height{
		display: block;
	} 
}
	
	</style>


<!-- ROW 1 !-->
<div class="row row-sections first-line row-eq-height">
	<!-- Section 1 !-->
	<div class="col-lg-3">
	</div>
	<!-- eof Section 1 !-->
	
	<!-- Section 2 & 3 !-->
	<div class="col-lg-9">
	</div>
	<!-- eof Section 2 & 3 !-->
 </div>
<!-- eof ROW 1 !-->




<!-- sections line !-->
<div class="row row-sections row-eq-height">
	
	<!-- column !-->
	<div class="col-lg-3">
		<h2></h2>
	</div>
	<!-- eof column !-->
	
	<!-- column !-->
	<div class="col-lg-6">
		<h2></h2>
	</div>
	<!-- eof column !-->
	
	<!-- column !-->
	<div class="col-lg-3">
		<!--<h2>Section 8</h2>
				<div class="pre-scrollable">
			<p>Taller box with a lot of content. This is a test. Lorem ipsum dolor sit amet, consectetur adipisicing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum.</p>
			
			<p>Taller box with a lot of content. This is a test. Lorem ipsum dolor sit amet, consectetur adipisicing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum.</p>
			
			<p>Taller box with a lot of content. This is a test. Lorem ipsum dolor sit amet, consectetur adipisicing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum.</p>
			
			<p>Taller box with a lot of content. This is a test. Lorem ipsum dolor sit amet, consectetur adipisicing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum.</p>
		!--></div>

	</div>
	<!-- eof column !-->
	
</div>
<!-- eof sections line !-->

<!-- sections line !-->
<div class="row row-sections row-eq-height">
	
	<!-- column !-->
	<div class="col-lg-3">
		<h2></h2>
		
	</div>
	<!-- eof column !-->
	
	<!-- column !-->
	<div class="col-lg-3">
		<h2></h2>
		
	</div>
	<!-- eof column !-->
	
	<!-- column !-->
	<div class="col-lg-3">
		<h2></h2>

	</div>
	<!-- eof column !-->
	
	<!-- column !-->
	<div class="col-lg-3">
	<!-- 	<h2>Section 12</h2>
		<p>Taller box with a lot of content. This is a test. Lorem ipsum dolor sit amet, consectetur adipisicing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum.</p>
	!-->
</div>
	<!-- eof column !-->
	
</div>
<!-- eof sections line !-->

<!-- sections line !-->
<div class="row row-sections row-eq-height">
	
	<!-- column !-->
	<div class="col-lg-3">
		<h2></h2>

	</div>
	<!-- eof column !-->
	
	<!-- column !-->
	<div class="col-lg-3">
		<h2></h2>
		
	</div>
	<!-- eof column !-->
	
	<!-- column !-->
	<div class="col-lg-3">
		<h2></h2>
	
	</div>
	<!-- eof column !-->
	
	<!-- column !-->
	<div class="col-lg-3">
		<!-- <h2>Section 16</h2>
				<div id="map"></div>
    <p><b>Address</b>: <span id="address"></span></p>
    <p id="error"></p>		


	!--></div>
	<!-- eof column !-->
	
</div>
<!-- eof sections line !-->

         
    
    </div>

<!--#include file="../inc/footer-main.asp"-->