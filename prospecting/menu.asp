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
	  width:350px;
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

<h1 class="page-header"><i class="fa fa-asterisk"></i> <%= GetTerm("Prospecting") %> Menu</h1>


<div class="container">
	
		
	<% If userIsAdmin(Session("userNo")) or (userIsTelemarketing(Session("userNo")) or userIsOutsideSales(Session("userNo")) or userIsInsideSales(Session("userNo")) or userIsOutsideSalesManager(Session("userNo")) or userIsInsideSalesManager(Session("userNo")) AND (GetCRMPermissionLevel(Session("userNo")) <> "NONE")) Then %>
		
		<div class="col-lg-3">
		
			<div class="row">
				<div class="col-md-12">
					<div class="menu-option">
						<div class="menu-split menu-info"><i class="fa fa-arrow-right" aria-hidden="true"></i></div>
						<div class="menu-text"><a href="<%= BaseURL %>prospecting/main.asp"><%= GetTerm("Prospects") %></a></div>
					</div>
				</div>
			</div>
			
			<div class="row">
				<div class="col-md-12">
					<div class="menu-option">
						<div class="menu-split menu-info"><i class="fa fa-arrow-right" aria-hidden="true"></i></div>
						<div class="menu-text"><a href="<%= BaseURL %>prospecting/mainRecyclePool.asp"><%= GetTerm("Recycle Pool") %></a></div>
					</div>
				</div>
			</div>
			
			<div class="row">
				<div class="col-md-12">
					<div class="menu-option">
						<div class="menu-split menu-info"><i class="fa fa-arrow-right" aria-hidden="true"></i></div>
						<div class="menu-text"><a href="<%= BaseURL %>prospecting/mainWonPool.asp"><%= GetTerm("New Customer Pool") %></a></div>
					</div>
				</div>
			</div>
			
			<% If (userIsInsideSalesManager(Session("userNo")) = True OR userIsOutsideSalesManager(Session("userNo")) = True OR userIsAdmin(Session("userNo"))) Then %>
			
				<div class="row">
					<div class="col-md-12">
						<div class="menu-option">
							<div class="menu-split menu-info"><i class="fa fa-arrow-right" aria-hidden="true"></i></div>
							<div class="menu-text"><a href="<%= BaseURL %>admin/prospecting/dashboard_salesrep_display_setup.asp">Select Sales Reps for Report Display</a></div>
						</div>
					</div>
				</div>
				
			<% End If %>
			
		</div>
		
	<% End If %>
	
	<% If GetCRMAddEditMenuPermissionLevel(Session("UserNo")) = vbTrue OR userIsAdmin(Session("userNo")) Then %>
	
		<div class="col-lg-3">
	
	
			<div class="row">
				<div class="col-md-12">
					<div class="menu-option">
						<div class="menu-split menu-info"><i class="fa fa-arrow-right" aria-hidden="true"></i></div>
						<div class="menu-text"><a href="<%= BaseURL %>prospecting/stages/main.asp">Add/Edit Stages</a></div>
					</div>
				</div>
			</div>
		
			<div class="row">
				<div class="col-md-12">
					<div class="menu-option">
						<div class="menu-split menu-info"><i class="fa fa-arrow-right" aria-hidden="true"></i></div>
						<div class="menu-text"><a href="<%= BaseURL %>prospecting/leadSource/main.asp">Add/Edit Lead Sources</a></div>
					</div>
				</div>
			</div>
		
			<div class="row">
				<div class="col-md-12">
					<div class="menu-option">
						<div class="menu-split menu-info"><i class="fa fa-arrow-right" aria-hidden="true"></i></div>
						<div class="menu-text"><a href="<%= BaseURL %>prospecting/industries/main.asp">Add/Edit Industries</a></div>
					</div>
				</div>
			</div>

			<div class="row">
				<div class="col-md-12">
					<div class="menu-option">
						<div class="menu-split menu-info"><i class="fa fa-arrow-right" aria-hidden="true"></i></div>
						<div class="menu-text"><a href="<%= BaseURL %>prospecting/reasons/main.asp">Add/Edit Reasons</a></div>
					</div>
				</div>
			</div>
		
			<div class="row">
				<div class="col-md-12">
					<div class="menu-option">
						<div class="menu-split menu-info"><i class="fa fa-arrow-right" aria-hidden="true"></i></div>
						<div class="menu-text"><a href="<%= BaseURL %>prospecting/predefinedNotes/main.asp">Add/Edit Predefined Notes</a></div>
					</div>
				</div>
			</div>
		
			<div class="row">
				<div class="col-md-12">
					<div class="menu-option">
						<div class="menu-split menu-info"><i class="fa fa-arrow-right" aria-hidden="true"></i></div>
						<div class="menu-text"><a href="<%= BaseURL %>prospecting/contactTitles/main.asp">Add/Edit Contact Titles</a></div>
					</div>
				</div>
			</div>
	
	
			<div class="row">
				<div class="col-md-12">
					<div class="menu-option">
						<div class="menu-split menu-info"><i class="fa fa-arrow-right" aria-hidden="true"></i></div>
						<div class="menu-text"><a href="<%= BaseURL %>prospecting/employeeRange/main.asp">Add/Edit Employee Range</a></div>
					</div>
				</div>
			</div>
		
			<div class="row">
				<div class="col-md-12">
					<div class="menu-option">
						<div class="menu-split menu-info"><i class="fa fa-arrow-right" aria-hidden="true"></i></div>
						<div class="menu-text"><a href="<%= BaseURL %>prospecting/activities/main.asp">Add/Edit Activities</a></div>
					</div>
				</div>
			</div>
		
			<div class="row">
				<div class="col-md-12">
					<div class="menu-option">
						<div class="menu-split menu-info"><i class="fa fa-arrow-right" aria-hidden="true"></i></div>
						<div class="menu-text"><a href="<%= BaseURL %>prospecting/competitors/main.asp">Add/Edit Competitors</a></div>
					</div>
				</div>
			</div>
		
		</div>
	
	<% End If %>
		
</div>



<!--#include file="../inc/footer-main.asp"-->