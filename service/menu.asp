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

<h1 class="page-header"><i class="fa fa-wrench"></i> <%= GetTerm("Service") %> Menu</h1>


<div class="container">
	
	<div class="col-lg-3">
		
		<% If userIsServiceManager(Session("userNo")) or userIsAdmin(Session("userNo")) Then %> 
	
			<div class="row">
				<div class="col-md-12">
					<div class="menu-option">
						<div class="menu-split menu-info"><i class="fa fa-arrow-right" aria-hidden="true"></i></div>
						<div class="menu-text"><a href="<%= BaseURL %>service/main.asp">Service Tickets</a></div>
					</div>
				</div>
			</div>
	
	

			<% If filterChangeModuleOn() Then %>
				<div class="row">
					<div class="col-md-12">
						<div class="menu-option">
							<div class="menu-split menu-info"><i class="fa fa-arrow-right" aria-hidden="true"></i></div>
							<div class="menu-text"><a href="<%= BaseURL %>service/filters/custfilters/main.asp">Manage Filter Changes</a></div>
						</div>
					</div>
				</div>
			<% End If %>

	
	
	
			<% If advancedDispatchIsOn()Then %>
			
				<% If userCanAccessServiceDispatchCenter(Session("UserNo")) = true Then %>
					<div class="row">
						<div class="col-md-12">
							<div class="menu-option">
								<div class="menu-split menu-info"><i class="fa fa-arrow-right" aria-hidden="true"></i></div>
								<div class="menu-text"><a href="<%= BaseURL %>service/dispatchcenter/main.asp">Dispatch Center</a></div>
							</div>
						</div>
					</div>
				<% End If %>
	
				<div class="row">
					<div class="col-md-12">
						<div class="menu-option">
							<div class="menu-split menu-info"><i class="fa fa-arrow-right" aria-hidden="true"></i></div>
							<div class="menu-text"><a href="<%= BaseURL %>service/serviceBoard.asp">Service Board</a></div>
						</div>
					</div>
				</div>
				
			<% End If %>
		
		<% End If %>
	</div>


			
	<div class="col-lg-3">

		<div class="row">
			<div class="col-md-12">
				<div class="menu-option">
					<div class="menu-split menu-info"><i class="fa fa-arrow-right" aria-hidden="true"></i></div>
					<div class="menu-text"><a href="<%= BaseURL %>service/symptomCodes/main.asp">Add/Edit Symptom Codes</a></div>
				</div>
			</div>
		</div>

		<div class="row">
			<div class="col-md-12">
				<div class="menu-option">
					<div class="menu-split menu-info"><i class="fa fa-arrow-right" aria-hidden="true"></i></div>
					<div class="menu-text"><a href="<%= BaseURL %>service/problemCodes/main.asp">Add/Edit Problem Codes</a></div>
				</div>
			</div>
		</div>

		<div class="row">
			<div class="col-md-12">
				<div class="menu-option">
					<div class="menu-split menu-info"><i class="fa fa-arrow-right" aria-hidden="true"></i></div>
					<div class="menu-text"><a href="<%= BaseURL %>service/resolutionCodes/main.asp">Add/Edit Resolution Codes</a></div>
				</div>
			</div>
		</div>
	
		<% If filterChangeModuleOn() Then %>
			<div class="row">
				<div class="col-md-12">
					<div class="menu-option">
						<div class="menu-split menu-info"><i class="fa fa-arrow-right" aria-hidden="true"></i></div>
						<div class="menu-text"><a href="<%= BaseURL %>service/filters/addeditfilters/main.asp">Add/Edit Filters</a></div>
					</div>
				</div>
			</div>
		<% End If %>
	
		<div class="row">
			<div class="col-md-12">
				<div class="menu-option">
					<div class="menu-split menu-info"><i class="fa fa-arrow-right" aria-hidden="true"></i></div>
					<div class="menu-text"><a href="<%= BaseURL %>service/parts/main.asp">Add/Edit Parts</a></div>
				</div>
			</div>
		</div>
		
	</div>
	
</div>



<!--#include file="../inc/footer-main.asp"-->