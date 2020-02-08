<!--#include file="../../inc/header.asp"-->

<%

currentEmailCategory1ViewedIDTab = Request.Querystring("cat1ID")
currentEmailCategory1ViewedIDTab = Replace(currentEmailCategory1ViewedIDTab,"_"," ")
currentEmailCategory2ViewedIDTab = Request.QueryString("tab")

%>


<script language="javascript">

	$(document).ready(function() {
		
		//we need to reload the page when the full email modal is closed because the hash tag
		//it leaves in the URL makes the page throw jQuery errors
	
	    $("[id^='myEmailModal']").on('hidden.bs.modal', function () {
	    
	    	$(this).removeData('bs.modal');
	    
			var counter = 1;
			
			//first, get all the 'active' tabs so we can return back the page
			//and open them after archiving the email(s)
			$("li.active").each(function() {
			    if (counter == 1) {
			    	cat1 = $(this).attr("id");
			    }
			    counter++;
			});
			
			cat2 = $("#emailTabLevel2.active").children(':first').attr("href");
			cat2 = cat2.substr(cat2.length - (cat2.length-1));

	        window.location.href = "allArchivedEmails.asp?cat1ID=" + cat1 + "#" + cat2;
	    });
	    
	

	    $('#unarchiveBtn').click(function() {
		
				var counter = 1;
				
				//first, get all the 'active' tabs so we can return back the page
				//and open them after archiving the email(s)
				$("li.active").each(function() {
				    if (counter == 1) {
				    	cat1 = $(this).attr("id");
				    }
				    counter++;
				});
				
				cat2 = $("#emailTabLevel2.active").children(':first').attr("href");
				cat2 = cat2.substr(cat2.length - (cat2.length-1))
				
				//second, get all the checked checkboxes and store their value in
				//an array to pass to the archive email page
				var chkBoxEmailIDArray = [];
						
				$('input:checkbox[name=chkEmail]:checked').each(function() 
				{
				   chkBoxEmailIDArray.push($(this).attr("id"));
				});	
				
				
				//post all obtained values to processing ASP page
				$.ajax({		
					type:"POST",
					data: "i="+chkBoxEmailIDArray+"&cat1="+cat1+"&cat2="+cat2,
					url: "unarchiveEmailFromTabView.asp",
					success: function (data) {
						swal('Email(s) Successfully Archived');
						window.location.href = "allArchivedEmails.asp?cat1ID=" + cat1 + "#" + cat2;
					}
				})	    
	    
	    });
	    
	    
	    
	    $('#forwardBtn').click(function() {

				var counter = 1;
				
				//first, get all the 'active' tabs so we can return back the page
				//and open them after archiving the email(s)
				$("li.active").each(function() {
				    if (counter == 1) {
				    	cat1 = $(this).attr("id");
				    }
				    counter++;
				});
				
				clientid = $("#txtClientID").val();
				cat2 = $("#emailTabLevel2.active").children(':first').attr("href");
				cat2 = cat2.substr(cat2.length - (cat2.length-1))
					
		
				//second, get all the checked checkboxes and store their value in
				//an array to pass to the archive email page
				var chkBoxEmailIDArray = [];
						
				$('input:checkbox[name=chkEmail]:checked').each(function() 
				{
				   chkBoxEmailIDArray.push($(this).attr("id"));
				});	
				
				window.location.href = "forwardEmailFromTabView.asp?i="+chkBoxEmailIDArray+"&cat1=" + cat1 + "&cat2=" + cat2 + "&cid=" + clientid;
	    
	    });	  
	    
  
	    
	    $('#resendBtn').click(function() {
				var counter = 1;
				
				//first, get all the 'active' tabs so we can return back the page
				//and open them after archiving the email(s)
				$("li.active").each(function() {
				    if (counter == 1) {
				    	cat1 = $(this).attr("id");
				    }
				    counter++;
				});
				
				cat2 = $("#emailTabLevel2.active").children(':first').attr("href");	
				cat2 = cat2.substr(cat2.length - (cat2.length-1));
				clientid = $("#txtClientID").val();				
		
				//second, get all the checked checkboxes and store their value in
				//an array to pass to the archive email page
				var chkBoxEmailIDArray = [];
						
				$('input:checkbox[name=chkEmail]:checked').each(function() 
				{
				   chkBoxEmailIDArray.push($(this).attr("id"));
				});	
				
				//post all obtained values to processing ASP page
				$.ajax({		
					type:"POST",
					data: "i="+chkBoxEmailIDArray+"&cat1="+cat1+"&cat2="+cat2,
					url: "resendEmailFromTabView.asp",
					success: function (data) {
						swal('Email(s) Successfully Re-Sent');
						window.location.href = "allArchivedEmails.asp?cat1ID=" + cat1 + "#" + cat2 + "&cid=" + clientid;
					}
				})	
	    
	    });	
	    


		//Functionality for check/uncheck all checkboxes that have the 'active' class name
		$("#checkAll").change(function() {	
		
		  	currentTab2 = window.location.hash;
		 	currentTab2 = currentTab2.substring(1);
		 	
		 	
		 	if (!currentTab2.trim()) {
		   		//The page either loaded for the first time and no level 2 tabs were clicked
		   		//or the page returned from a calling page with level 2 as a querystring parameter
		   		cat2 = $("#emailTabLevel2.active").children(':first').attr("href");
		   		cat2 = cat2.substring(1);
		   		$("input:checkbox." +cat2).prop('checked', $(this).prop("checked"));
		   	}
		   	else {
		   		//The page did not reload, so we take the hash tag value in the browsers URL
		   		$("input:checkbox." +currentTab2).prop('checked', $(this).prop("checked"));
			}
		});
		
		
		
		//refresh button functionality - reloads tabs
		 
		$('#refresh').click(function() { 
			window.location.reload();    
		});

	
		//function that gets the value of the tab when it is clicked and then
		//updates the value of a hidden form field so when the page posts, it returns
		//back to the tab that was previously opened
		
		//THIS IS SPECIAL TAB CODE FOR EMAILS ONLY - WILL NOT WORK ON OTHER INSIGHT TABS
			    
		$('a[data-toggle="tab"]').on('shown.bs.tab', function (e) {
			  var targetTab = $(e.target).attr("href") // activated tab
			  var LastTab = $(e.relatedTarget).attr("href"); // get last tab
			  //$('input[name="txtTab"]').val(targetTab);
			  
		});	    

			});
		
	
</script>

<style type="text/css">

	.nav-tabs .glyphicon:not(.no-margin) { margin-right:10px; }
	.tab-pane .list-group-item:first-child {border-top-right-radius: 0px;border-top-left-radius: 0px;}
	.tab-pane .list-group-item:last-child {border-bottom-right-radius: 0px;border-bottom-left-radius: 0px;}
	.tab-pane .list-group .checkbox { display: inline-block;margin: 0px; }
	.tab-pane .list-group input[type=checkbox]{ margin-top: 8px; }
	.tab-pane .list-group .glyphicon { margin-right:5px; }
	.tab-pane .list-group .glyphicon:hover { color:#FFBC00; }
	
	a.list-group-item.read { color: #222;background-color: #F3F3F3; }
	
	hr { margin-top: 5px;margin-bottom: 10px; }
	
	.nav-pills>li>a {padding: 5px 10px;}
	
	.nav-pills>li.active>a{
    color: #fff;
    background-color: #428bca;
	}

	.nav-pills>li>a:hover, .nav>li>a:focus {
	    text-decoration: none;
	    background-color: #eee;
	}	
	.inbox-head {
	    /*background: none repeat scroll 0 0 #41cac0;*/
	    border-radius: 0 4px 0 0;
	    color: #555;
	    min-height: 50px;
	    padding: 20px;
	    padding-top:0px;
	    padding-bottom:0px;
	}
	.inbox-head h3 {
	    display: inline-block;
	    font-weight: 300;
	    margin: 0;
	    padding-top: 6px;
	}
	.sr-input {
		border: 1px #ccc solid;
	    border-radius: 4px 0 0 4px;
	    box-shadow: none;
	    color: #8a8a8a;
	    float: left;
	    height: 40px;
	    padding: 0 10px;
	    width: 400px;
	}
	.sr-btn {
	    background: none repeat scroll 0 0 #337ab7;
	    border: medium none;
	    border-radius: 0 4px 4px 0;
	    color: #fff;
	    height: 40px;
	    padding: 0 20px;
	}
	.sr-btn:hover {
	    color: #fff;
	    background-color: #286090;
	    border-color: #204d74;	
	}
	
	.container-full {
	  margin: 0 auto;
	  width: 100%;
	}	
	
	.modal-link{
	cursor:pointer;
	}
	
	.modal-content{
		max-height:650px;
		overflow-y:auto;
		width:750px;
	}
	
	 .modal-content .row{
		 padding-bottom:20px;
	 }
	
	 .modal-content p{
		 margin-bottom:20px;
		 white-space:normal;
	 }
	 
	 .dropdown-menu{
		min-width:250px;
	}
	.inbox-head .sr-btn .fa{
		color: #fff;
	}
	
</style>


<h1 class="page-header"><i class="fa fa-envelope-o"></i> All Archived Emails</h1>


<div class="container-full">
    <div class="row">
        <div class="inbox-head">
          <h3>Sent Items That Were Archived</h3>
      </div>

        <div class="col-sm-3 col-md-2">
            <div class="btn-group">
                <button type="button" class="btn btn-primary dropdown-toggle" data-toggle="dropdown">
                    All Archived Mail <span class="caret"></span>
                </button>
                <ul class="dropdown-menu" role="menu">
                	<li><a href="allSentEmails.asp">Sent Mail</a></li>
                    <li><a href="allArchivedEmails.asp">Archived Mail</a></li>
                    <li><a href="allFailedEmails.asp">Failed Mail</a></li>
                </ul>
            </div>
        </div>
        <div class="col-sm-9 col-md-10">
             
            <!-- Split button -->
            <div class="btn-group">
                <button type="button" class="btn btn-default">
                    <div class="checkbox" style="margin: 0;">
                        <label style="display:block;">
                            <input type="checkbox" name="checkAll" id="checkAll">
                        </label>
                    </div>
                </button>
            </div>
            
            <button type="button" class="btn btn-default" data-toggle="tooltip" title="refresh" id="refresh">
                   <span class="fa fa-refresh"></span>&nbsp;Refresh View   
            </button>
            <button type="button" class="btn btn-default" data-toggle="tooltip" title="archive" id="unarchiveBtn">
                   <span class="fa fa-archive"></span>&nbsp;Un-Archive Selected 
            </button>
            <button type="button" class="btn btn-default" data-toggle="tooltip" title="forward" id="forwardBtn">
                   <span class="fa fa-mail-forward"></span>&nbsp;Forward Selected  
            </button>
            <button type="button" class="btn btn-default" data-toggle="tooltip" title="resend" id="resendBtn">
                   <span class="fa fa-retweet"></span>&nbsp;Resend To Original Recipients
            </button>

             <!--                                               
            <div class="pull-right">
		          <form action="#" class="pull-right position">
		              <div class="input-append">
		                  <input type="text" id="search" class="sr-input" placeholder="Search Sent Items...">
		                  <button class="btn sr-btn" type="button"><i class="fa fa-search"></i></button>
		              </div>
		          </form>
            </div>-->
            
        </div>
    </div>
    <hr />
    <div class="row">
        <div class="col-sm-3 col-md-2">
            <ul class="nav nav-pills nav-stacked">
			<%       
				Set cnnSentEmail = Server.CreateObject("ADODB.Connection")
				cnnSentEmail.open (Session("ClientCnnString"))
				Set rsSentEmail = Server.CreateObject("ADODB.Recordset")
				rsSentEmail.CursorLocation = 3 
				
				SQL_SentEmail = "SELECT COUNT(EmailCategory1) as catCount,EmailCategory1 FROM SC_EmailLog  WHERE Archived = 1 GROUP BY EmailCategory1 ORDER BY EmailCategory1"
				
				Set rsSentEmail = cnnSentEmail.Execute(SQL_SentEmail)
				
				IF Not rsSentEmail.EOF Then
					
					categoryCount = 0
					
					Do While NOT rsSentEmail.EOF
					
													
						If DefView = "" Then
							DefView = rsSentEmail("EmailCategory1")  
						End If
						
						'**********************************************************************
						'TO CREATE TAB NAME, REMOVE ALL SPACES FROM SQL FIELD NAME
						'**********************************************************************
						emailCategory1TabName = Trim(rsSentEmail("EmailCategory1"))
						emailCategory1TabName = Replace(emailCategory1TabName," ","_")
						
						%>
			                <li id="<%= emailCategory1TabName %>" <% If (currentEmailCategory1ViewedIDTab = rsSentEmail("EmailCategory1")) OR (currentEmailCategory1ViewedIDTab= "" AND categoryCount = 0) Then Response.write("class='active'") %>><a href="allArchivedEmails.asp?cat1ID=<%= rsSentEmail("EmailCategory1") %>"><span class="badge pull-right"><%= rsSentEmail("catCount") %></span> <%= rsSentEmail("EmailCategory1") %> - Archived</a></li>
		           		<%
		           		categoryCount = categoryCount + 1
		           
		           		rsSentEmail.MoveNext
		           		Loop
           		End If
           		
		
				set rsSentEmail = Nothing
				cnnSentEmail.close
				set cnnSentEmail = Nothing

           
           %>
            
                
            </ul>
            <hr />
            <!--<a href="allFailedEmails.asp" class="btn btn-danger btn-sm btn-block" role="button">Emails That Failed To Send</a>-->
            
            
        </div>
        
        
        <div class="col-sm-9 col-md-10">
        
        	<form method="post" action="allArchivedEmails.asp" name="frmAllSentEmails" id="frmAllSentEmails">

			<input type="hidden" name="txtTab" id="txtTab" value="">
			<input type="hidden" name="txtClientID" id="txtClientID" value="<%= MUV_Read("ClientID") %>">

			
            <!-- Nav tabs -->
            <ul class="nav nav-tabs" role="tablist" id="emailtabs">
            <%
            
            	Set cnnSentEmail = Server.CreateObject("ADODB.Connection")
				cnnSentEmail.open (Session("ClientCnnString"))
				Set rsSentEmail = Server.CreateObject("ADODB.Recordset")
				rsSentEmail.CursorLocation = 3 
				
				If  currentEmailCategory1ViewedIDTab = "" Then 
					SQL_SentEmail = "SELECT EmailCategory2 FROM SC_EmailLog WHERE EmailCategory1 = '" &  DefView & "' AND Archived = 1 GROUP BY EmailCategory2 ORDER BY EmailCategory2"
				Else
					SQL_SentEmail = "SELECT EmailCategory2 FROM SC_EmailLog WHERE EmailCategory1 = '" & currentEmailCategory1ViewedIDTab & "' AND Archived = 1 GROUP BY EmailCategory2 ORDER BY EmailCategory2"
				End IF
				
				'Response.write(SQL_SentEmail)

				Set rsSentEmail = cnnSentEmail.Execute(SQL_SentEmail)
				
				IF Not rsSentEmail.EOF Then

					categoryTabCount = 0
					
					Do While NOT rsSentEmail.EOF
					
						If currentEmailCategory2ViewedIDTab <> "" Then
							querystringTab2ID = Replace(Trim(currentEmailCategory2ViewedIDTab)," ","")
						Else
							querystringTab2ID = ""
						End If
						
						'**********************************************************************
						'TO CREATE TAB NAME, REMOVE ALL SPACES FROM SQL FIELD NAME
						'**********************************************************************
						emailCategory2TabName = Trim(rsSentEmail("EmailCategory2"))
						emailCategory2TabName = Replace(emailCategory2TabName," ","")
						
						'Response.write("emailCategory2TabName : " & emailCategory2TabName & "<br>")
						'Response.write("querystringTab2ID : " & querystringTab2ID & "<br>")
						'Response.write("currentEmailCategory2ViewedIDTab : " & currentEmailCategory2ViewedIDTab & "<br>")


						%>
			                
			                <li id="emailTabLevel2" role="presentation" 
			                <% If ((querystringTab2ID = emailCategory2TabName) OR (currentEmailCategory2ViewedIDTab = "" AND categoryTabCount =0)) Then Response.write("class='active'")%>>
			                <a href="#<%= emailCategory2TabName %>" role="tab" data-toggle="tab"><span class="glyphicon glyphicon-inbox"></span><%= rsSentEmail("EmailCategory2") %></a>
			                </li>
			                
		           		<%
		           		categoryTabCount = categoryTabCount + 1
		           
		           		rsSentEmail.MoveNext
		           		Loop
           		End If
           		
		
				set rsSentEmail = Nothing
				cnnSentEmail.close
				set cnnSentEmail = Nothing
				
			%>
            </ul>
            <!-- Tab panes -->
            
            
            <div class="tab-content" id="emailtabs-content">
            
            <%
            
            	Set cnnSentEmail = Server.CreateObject("ADODB.Connection")
				cnnSentEmail.open (Session("ClientCnnString"))
				Set rsSentEmail = Server.CreateObject("ADODB.Recordset")
				rsSentEmail.CursorLocation = 3 
				
				If  currentEmailCategory1ViewedIDTab = "" Then 
					SQL_SentEmail = "SELECT * FROM SC_EmailLog WHERE EmailCategory1 = '" & DefView & "' AND Archived = 1 ORDER BY EmailCategory2, EmailDate DESC, EmailTime DESC"
				Else
					SQL_SentEmail = "SELECT * FROM SC_EmailLog WHERE EmailCategory1 = '" & currentEmailCategory1ViewedIDTab & "' AND Archived = 1  ORDER BY EmailCategory2, EmailDate DESC, EmailTime DESC"		
				End IF
				
				'Response.write(SQL_SentEmail)
				
				Set rsSentEmail = cnnSentEmail.Execute(SQL_SentEmail)
				
				IF Not rsSentEmail.EOF Then
				
					oldTab2ID = ""
					currentTab2ID = ""
					categoryCount2 = 0
					
					Do While NOT rsSentEmail.EOF
										
						currentTab2ID = rsSentEmail("EmailCategory2")
						
						If currentEmailCategory2ViewedIDTab <> "" Then
							querystringTab2ID = Replace(Trim(currentEmailCategory2ViewedIDTab)," ","")
						Else
							querystringTab2ID = ""
						End If

						If currentEmailCategory1ViewedIDTab <> "" Then
							querystringTab1ID = Replace(Trim(currentEmailCategory1ViewedIDTab)," ","_")
						Else
							querystringTab1ID = ""
						End If
						
						'**********************************************************************
						'TO CREATE TAB PANEL ID, REMOVE ALL SPACES FROM SQL FIELD NAME
						'**********************************************************************
						emailCategory2LowerTabName = Trim(rsSentEmail("EmailCategory2"))
						emailCategory2LowerTabName = Replace(emailCategory2LowerTabName," ","")
		
							If (currentTab2ID <> oldTab2ID) Then
								%>
								<div role="tabpanel" class="tab-pane fade in <% If (querystringTab2ID = emailCategory2LowerTabName) OR (categoryCount2 = 0) Then Response.write("active") %>" id="<%= emailCategory2LowerTabName %>">
			            		<div class="list-group">
							<% End If %>
							
							<%
												
							 InternalRecordNumber = rsSentEmail("InternalRecordNumber")
							 RecordCreationDateTime = rsSentEmail("RecordCreationDateTime")
							 EmailDate = FormatDateTime(rsSentEmail("EmailDate"),2)
							 EmailTime = FormatDateTime(rsSentEmail("EmailTime"),3)
							 EmailSendTo = rsSentEmail("EmailTo")
							 EmailSendFrom = rsSentEmail("EmailFrom")
							 EmailSendFromName = rsSentEmail("EmailFromName")
							 Subject = rsSentEmail("Subject")
							 Body = stripHTML(rsSentEmail("Body"))
							 CCs = rsSentEmail("CCs")
							 BCCs = rsSentEmail("BCCs")
							 Attachment = rsSentEmail("Attachment")
							 ASPMailStatus = rsSentEmail("ASPMailStatus")
								
							%>

							<div class="list-group-item">
									<div class="checkbox">
										<label>
											<input type="checkbox" name="chkEmail" class="<%= emailCategory2LowerTabName %>" id="<%= InternalRecordNumber %>">
										</label>
									</div>
									<!--<span class="glyphicon glyphicon-star-empty"></span>-->
									<span class="name" style="min-width: 120px; display: inline-block;">TO: <%= EmailSendTo %></span> 
									<span class=""><strong>
									
									<% If currentEmailCategory1ViewedIDTab = "" Then %>
										<a data-toggle="modal" data-show="true" data-target="#myEmailModal<%= InternalRecordNumber %>" href="displayFullEmailModal.asp?i=<%= InternalRecordNumber %>&cat1=&cat2="><%= Subject %></a>
									<% Else %>	
										<a data-toggle="modal" data-show="true" data-target="#myEmailModal<%= InternalRecordNumber %>" href="displayFullEmailModal.asp?i=<%= InternalRecordNumber %>&cat1=<%= querystringTab1ID %>&cat2=<%= querystringTab2ID %>"><%= Subject %></a>
									<% End If %>
									
									</strong></span>
									<span class="text-muted" style="font-size: 11px;">- <%= Left(Body,50) %></span> 
									
									<% If Abs(dateDiff("d",EmailDate,Now())) = 1 Then %>
										<span class="badge">Today, <%= EmailTime %></span> 
									<% Else %>
										<span class="badge"><%= EmailDate %>, <%= EmailTime %></span>
									<% End If %>
									
									<% If Attachment <> "" Then %>
										<span class="pull-right"><span class="glyphicon glyphicon-paperclip"></span></span>
									<% Else %>
										<span class="pull-right"></span>
									<% End If %>
									

							</div>
							
							<!-- modal  starts here !-->
							 <!-- Modal -->
							<div class="modal fade" id="myEmailModal<%= InternalRecordNumber %>" tabindex="-1" role="dialog" aria-labelledby="myEmailModalLabel<%= InternalRecordNumber %>" aria-hidden="true">
							    <div class="modal-dialog">
							        <div class="modal-content">
							            <div class="modal-body"></div>
							        </div>
							        <!-- /.modal-content -->
							    </div>
							    <!-- /.modal-dialog -->
							</div>
							<!-- /.modal -->
							<!-- modal  ends here !-->
														

			            <% 
			           	oldTab2ID = rsSentEmail("EmailCategory2")
			           	
			           	
		           		rsSentEmail.MoveNext
		           		
		           		If NOT rsSentEmail.EOF Then
		           			currentTab2ID = rsSentEmail("EmailCategory2") 
		           		End If
		           		
		           		If (currentTab2ID <> oldTab2ID) AND oldTab2ID <> "" Then 
		           			categoryCount2 = categoryCount2 + 1%>                         
			                </div> <!-- end list group -->
			              </div> <!-- end tab pane -->
			             <% End If

		           		
           		Loop
           		           		
       		End If
       		%>
	                </div> <!-- end list group -->
	          </div> <!-- end tab pane -->
 		
			<%
			set rsSentEmail = Nothing
			cnnSentEmail.close
			set cnnSentEmail = Nothing
			
		%>

     		</form>
            </div>
            
            <!-- ********************************************************************************** -->
            <!-- PAGING -->
            <!-- ********************************************************************************** -->
            
            <!--
			    <div class="row">
			        <div class="col-sm-12 col-md-12">                                           
			            <div class="pull-right">
			                <span class="text-muted"><b>1</b>–<b>50</b> of <b>277</b></span>
			                <div class="btn-group btn-group-sm">
			                    <button type="button" class="btn btn-default">
			                        <span class="glyphicon glyphicon-chevron-left"></span>
			                    </button>
			                    <button type="button" class="btn btn-default">
			                        <span class="glyphicon glyphicon-chevron-right"></span>
			                    </button>
			                </div>
			            </div>
			        </div>
			    </div>-->
			    
			 <!-- ********************************************************************************** -->
            
        </div>
    </div>
</div>

							
<!--#include file="../../inc/footer-main.asp"-->