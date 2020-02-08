<!--#include file="../inc/header.asp"-->

<style>
	.container {
	    width: 100%;
	}
	
	.menu-option{
	  display: inline-block;
	  font-size: 14px;
	  text-align: left;
	  background-color: #fff;
	  height: 40px;
	  -webkit-box-shadow: 0 1px 1px 0 rgba(0,0,0,.2);
	  box-shadow: 0 1px 1px 0 rgba(0,0,0,.1);
	  margin-bottom: 10px;
	  width:360px;
	}
	
	.menu-option:hover{
	    cursor: pointer;
	    -webkit-box-shadow: 0 1px 1px 0 rgba(0,0,0,.4);
	  	box-shadow: 0 1px 1px 0 rgba(0,0,0,.3);
	}
	
	.menu-option > .menu-split{
	  background: #337ab7;
	  width: 33px;
	  float: left;
	  color: #fff!important;
	  height: 100%;
	  text-align: center;
	}
	
	.menu-option > .menu-split > .fa{
	  position:relative;
	  top: calc(50% - 9px)!important; /* 50% - 3/4 of icon height */
	}
	.menu-option > .menu-split.menu-success{
	  background: #5cb85c!important;
	}
	
	.menu-option > .menu-split.menu-danger{
	  background: #d9534f!important;
	}
	
	.menu-option > .menu-split.menu-info{
	  background: #5bc0de!important;
	}

	.menu-option > .menu-text{
	  line-height: 19px;
	  padding-top: 11px;
	  padding-left: 45px;
	  padding-right: 20px;
	}
</style>

<h1 class="page-header"><i class="fa fa-dollar"></i> <%= GetTerm("Accounts Receivable") %> Menu</h1>


<div class="container">

	<div class="col-lg-4">
	
		<% If MUV_Read("serviceModuleOn") = "Enabled" Then %>
			<div class="row">
				<div class="col-md-12">
					<div class="menu-option">
						<div class="menu-split menu-info"><i class="fa fa-arrow-right" aria-hidden="true"></i></div>
						<div class="menu-text"><a href="<%= BaseURL %>accountsreceivable/TicketsOnHold.asp">Service Tickets On Hold</a></div>
					</div>
				</div>
			</div>
		<% End If %>
			
		<div class="row">
			<div class="col-md-12">
				<div class="menu-option">
					<div class="menu-split menu-info"><i class="fa fa-arrow-right" aria-hidden="true"></i></div>
					<div class="menu-text"><a href="<%= BaseURL %>accountsreceivable/reports.asp">Invoices / Statements</a></div>
				</div>
			</div>
		</div>
		
		<div class="row">
			<div class="col-md-12">
				<div class="menu-option">
					<div class="menu-split menu-info"><i class="fa fa-arrow-right" aria-hidden="true"></i></div>
					<div class="menu-text"><a href="<%= BaseURL %>accountsreceivable/sage/main.asp">Export Invoices To SAGE Accounting&nbsp;<img src="<%= BaseURL %>img/partnericons/sage-logo.png"></a></div>
				</div>
			</div>
		</div>
	
	</div>
	
	<div class="col-lg-3">
		<% If userIsAdmin(Session("userNo")) = True Then %>

			<% If MUV_Read("ShowAddEditCustomer") = "1" Then %>
				<div class="row">
					<div class="col-md-12">
						<div class="menu-option">
							<div class="menu-split menu-info"><i class="fa fa-arrow-right" aria-hidden="true"></i></div>
							<div class="menu-text"><a href="<%= BaseURL %>accountsreceivable/customermininfo/main.asp">Add/Edit Customer</a></div>
						</div>
					</div>
				</div>
				
		
				<div class="row">
					<div class="col-md-12">
						<div class="menu-option">
							<div class="menu-split menu-info"><i class="fa fa-arrow-right" aria-hidden="true"></i></div>
							<div class="menu-text"><a href="<%= BaseURL %>accountsreceivable/customermininfo/contactTitles/main.asp">Add/Edit Contact Titles</a></div>
						</div>
					</div>
				</div>
				
			<% End If %>
	
			<div class="row">
				<div class="col-md-12">
					<div class="menu-option">
						<div class="menu-split menu-info"><i class="fa fa-arrow-right" aria-hidden="true"></i></div>
						<div class="menu-text"><a href="<%= BaseURL %>accountsreceivable/customerclass/main.asp">Add/Edit Customer Classes</a></div>
					</div>
				</div>
			</div>
	
			<div class="row">
				<div class="col-md-12">
					<div class="menu-option">
						<div class="menu-split menu-info"><i class="fa fa-arrow-right" aria-hidden="true"></i></div>
						<div class="menu-text"><a href="<%= BaseURL %>accountsreceivable/customermapping/main.asp">Add/Edit Customer Mapping</a></div>
					</div>
				</div>
			</div>

			<div class="row">
				<div class="col-md-12">
					<div class="menu-option">
						<div class="menu-split menu-info"><i class="fa fa-arrow-right" aria-hidden="true"></i></div>
						<div class="menu-text"><a href="<%= BaseURL %>accountsreceivable/regions/main.asp">Add/Edit Regions</a></div>
					</div>
				</div>
			</div>			

			<div class="row">
				<div class="col-md-12">
					<div class="menu-option">
						<div class="menu-split menu-info"><i class="fa fa-arrow-right" aria-hidden="true"></i></div>
						<div class="menu-text"><a href="<%= BaseURL %>accountsreceivable/customertype/main.asp">Add/Edit Customer Types</a></div>
					</div>
				</div>
			</div>

			<div class="row">
				<div class="col-md-12">
					<div class="menu-option">
						<div class="menu-split menu-info"><i class="fa fa-arrow-right" aria-hidden="true"></i></div>
						<div class="menu-text"><a href="<%= BaseURL %>accountsreceivable/customerreferral/main.asp">Add/Edit Referrals</a></div>
					</div>
				</div>
			</div>

			<div class="row">
				<div class="col-md-12">
					<div class="menu-option">
						<div class="menu-split menu-info"><i class="fa fa-arrow-right" aria-hidden="true"></i></div>
						<div class="menu-text"><a href="<%= BaseURL %>accountsreceivable/chain/main.asp">Add/Edit Chain</a></div>
					</div>
				</div>
			</div>

			<div class="row">
				<div class="col-md-12">
					<div class="menu-option">
						<div class="menu-split menu-info"><i class="fa fa-arrow-right" aria-hidden="true"></i></div>
						<div class="menu-text"><a href="<%= BaseURL %>accountsreceivable/terms/main.asp">Add/Edit Terms</a></div>
					</div>
				</div>
			</div>			

			<div class="row">
				<div class="col-md-12">
					<div class="menu-option">
						<div class="menu-split menu-info"><i class="fa fa-arrow-right" aria-hidden="true"></i></div>
						<div class="menu-text"><a href="<%= BaseURL %>accountsreceivable/paymentmethods/main.asp">Add/Edit Payment Method</a></div>
					</div>
				</div>
			</div>
			
		<% End If %>
	</div>
		
</div>



<!--#include file="../inc/footer-main.asp"-->