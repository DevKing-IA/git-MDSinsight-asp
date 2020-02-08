<!--#include file="inc/header-tech-and-driver.asp"-->

<%
Session("MemoNumber") = "" 'Init to empty 
Session("ServiceCustID") = ""
Session("ServiceCustName") = ""  'Init to empty
%>

<div class="container-fluid">
	<div class="row">
		<h1 class="fieldservice-heading" >Thank You.<br><br>The service ticket has been created.</h1>
		<div class="container-fluid">
			<div class="row">
				<div class="col-lg-12">
					<a href="<%= baseURL %>/mobile/TechAndDriver/main_menu.asp"> 
						<button type="button" class="btn btn-danger btn-lg"><i class="fa fa-sign-out"></i> Main Menu</button>
					</a>
					<a href="<%= baseURL %>logout.asp"> 
						<button type="button" class="btn btn-danger btn-lg"><i class="fa fa-sign-out"></i> Log Out</button>
					</a>
				</div>
			</div>
		</div>
	</div>
</div>

<!--#include file="inc/footer-tech-and-driver.asp"-->
