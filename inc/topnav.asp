	
<% If userIsAdmin(Session("userNo")) = True Then %>	
	<div class="collapse navbar-collapse js-navbar-collapse pull-right">
	    
		<ul class="nav navbar-nav">
			<li class="dropdown dropdown-large">
				<a href="#" class="dropdown-toggle" data-toggle="dropdown">	<i class="fa fa-cog fa-lg"></i><b> Admin</b></a>
				
				<ul class="dropdown-menu dropdown-menu-large row">
					<li class="col-sm-3">
						<ul>
							<li class="dropdown-header">Settings</li>
							<li><a href="<%= BaseURL %>admin/company/main.asp">Company Settings</a></li>
							<li><a href="<%= BaseURL %>admin/global/main.asp">Global Settings</a></li>
						</ul>
					</li>
					
					<li class="col-sm-3">
						<ul>
							<li class="dropdown-header">Add / Edit</li>
							<li><a href="<%= BaseURL %>accountsreceivable/regions/main.asp">Regions</a></li>
							<li class="divider"></li>
						</ul>
					</li>

					
					<li class="col-sm-3">
						<ul>
							<li class="dropdown-header">Data Exchange</li>
							<li><a href="<%= BaseURL %>admin/dataexchange/main.asp">Upload Files</a></li>
							<li class="divider"></li>
						</ul>
					</li>


					<li class="col-sm-3">
						<ul>
							<li class="dropdown-header">Users</li>
							<li><a href="<%= BaseURL %>admin/users/main.asp">Manage Users</a></li>
							<li><a href="<%= BaseURL %>admin/teams/main.asp">Manage Teams</a></li>
							<li class="divider"></li>
						</ul>
					</li>
					
					<% If MUV_Read("FilterTrax") <> "1" Then %>
						<li class="col-sm-3">
							<ul>
								<li class="dropdown-header">Dashboard Setup</li>
								<li><a href="<%= BaseURL %>main/user_activity_chart_setup.asp">User Activity Chart</a></li>
								<li class="divider"></li>
							</ul>
						</li>
					<% End If %>
					
					<li class="col-sm-3">
						<ul>
							<li class="dropdown-header">Email Settings</li>
							<li><a href="<%= BaseURL %>admin/emailcustomize/main.asp">Customize System Emails</a></li>
							<li><a href="<%= BaseURL %>admin/emailsettings/service_tickets.asp">Service Tickets</a></li>
							<% If Session("MAILOFF") = 1 Then %>
								<li><a href="<%= BaseURL %>admin/toggle_email.asp"><strong><font color="green">START Email</font></strong></a></li>
							<% Else %>
								<li><a href="<%= BaseURL %>admin/toggle_email.asp"><strong><font color="red">STOP All Email</font></strong></a></li>
							<% End If%>
						</ul>
						
					</li>

					<li class="col-sm-3">
						<ul>
							<li class="dropdown-header">Sent Email</li>
							<li><a href="<%= BaseURL %>admin/emailsettings/allSentEmails.asp">All Sent Email</a></li>
							<li><a href="<%= BaseURL %>admin/emailsettings/allFailedEmails.asp">All Failed Email</a></li>
							<li><a href="<%= BaseURL %>admin/emailsettings/allArchivedEmails.asp">All Archived Email</a></li>
						</ul>
					</li>
										
				</ul>
				
			</li>
		</ul>
		
	</div><!-- /.nav-collapse -->
	
<% End If %>
