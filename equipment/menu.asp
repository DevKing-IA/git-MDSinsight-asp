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

<h1 class="page-header"><i class="fa fa-fw fa-coffee"></i> <%= GetTerm("Equipment") %> Menu</h1>


<div class="container">

	<div class="col-lg-3">

		<div class="row">
			<div class="col-md-12">
				<div class="menu-option">
					<div class="menu-split menu-info"><i class="fa fa-arrow-right" aria-hidden="true"></i></div>
					<div class="menu-text"><a href="<%= BaseURL %>equipment/equipment/addEquipment.asp">Add/Edit Equipment</a></div>
				</div>
			</div>
		</div>
			
		<div class="row">
			<div class="col-md-12">
				<div class="menu-option">
					<div class="menu-split menu-info"><i class="fa fa-arrow-right" aria-hidden="true"></i></div>
					<div class="menu-text"><a href="<%= BaseURL %>equipment/equipment/findEquipment.asp">Find/Edit Equipment Piece</a></div>
				</div>
			</div>
		</div>
	</div>
	
	<div class="col-lg-3">
		
		<div class="row">
			<div class="col-md-12">
				<div class="menu-option">
					<div class="menu-split menu-info"><i class="fa fa-arrow-right" aria-hidden="true"></i></div>
					<div class="menu-text"><a href="<%= BaseURL %>equipment/classes/main.asp">Add/Edit Classes</a></div>
				</div>
			</div>
		</div>
		
		<div class="row">
			<div class="col-md-12">
				<div class="menu-option">
					<div class="menu-split menu-info"><i class="fa fa-arrow-right" aria-hidden="true"></i></div>
					<div class="menu-text"><a href="<%= BaseURL %>equipment/groups/main.asp">Add/Edit Groups</a></div>
				</div>
			</div>
		</div>
		
		<div class="row">
			<div class="col-md-12">
				<div class="menu-option">
					<div class="menu-split menu-info"><i class="fa fa-arrow-right" aria-hidden="true"></i></div>
					<div class="menu-text"><a href="<%= BaseURL %>equipment/manufacturers/main.asp">Add/Edit Manufacturers</a></div>
				</div>
			</div>
		</div>
		
		<div class="row">
			<div class="col-md-12">
				<div class="menu-option">
					<div class="menu-split menu-info"><i class="fa fa-arrow-right" aria-hidden="true"></i></div>
					<div class="menu-text"><a href="<%= BaseURL %>equipment/brands/main.asp">Add/Edit Brands</a></div>
				</div>
			</div>
		</div>
		
		<div class="row">
			<div class="col-md-12">
				<div class="menu-option">
					<div class="menu-split menu-info"><i class="fa fa-arrow-right" aria-hidden="true"></i></div>
					<div class="menu-text"><a href="<%= BaseURL %>equipment/models/main.asp">Add/Edit Models</a></div>
				</div>
			</div>
		</div>
		
		<div class="row">
			<div class="col-md-12">
				<div class="menu-option">
					<div class="menu-split menu-info"><i class="fa fa-arrow-right" aria-hidden="true"></i></div>
					<div class="menu-text"><a href="<%= BaseURL %>equipment/condition/main.asp">Add/Edit Conditions</a></div>
				</div>
			</div>
		</div>
		
		<div class="row">
			<div class="col-md-12">
				<div class="menu-option">
					<div class="menu-split menu-info"><i class="fa fa-arrow-right" aria-hidden="true"></i></div>
					<div class="menu-text"><a href="<%= BaseURL %>equipment/statuscodes/main.asp">Add/Edit Status Codes</a></div>
				</div>
			</div>
		</div>
		
		<div class="row">
			<div class="col-md-12">
				<div class="menu-option">
					<div class="menu-split menu-info"><i class="fa fa-arrow-right" aria-hidden="true"></i></div>
					<div class="menu-text"><a href="<%= BaseURL %>equipment/movementcodes/main.asp">Add/Edit Movement Codes</a></div>
				</div>
			</div>
		</div>
		
		<div class="row">
			<div class="col-md-12">
				<div class="menu-option">
					<div class="menu-split menu-info"><i class="fa fa-arrow-right" aria-hidden="true"></i></div>
					<div class="menu-text"><a href="<%= BaseURL %>equipment/acquisitioncodes/main.asp">Add/Edit Acquisition Codes</a></div>
				</div>
			</div>
		</div>
		
	</div>
	
</div>

<!--#include file="../inc/footer-main.asp"-->