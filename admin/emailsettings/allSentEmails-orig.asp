<!--#include file="../../inc/header.asp"-->

<%

currentEmailCategory1ViewedIDTab = Request.Querystring("cat1ID")
currentEmailCategory2ViewedIDTab = Request.QueryString("tab")

%>


<script language="javascript">

	$(document).ready(function() {
	
		//refresh button functionality - reloads tabs
		
		$(document).on('click', '#refresh', function () {   
			var current_index = $("#emailtabs").tabs("option","active");
			$("#emailtabs").tabs('load',current_index);	    
		});
		
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

	        window.location.href = "allSentEmails.asp?cat1ID=" + cat1 + "#" + cat2;
	    });
	    
	

	    $('#archiveBtn').click(function() {
		
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
					url: "archiveEmailFromTabView.asp",
					success: function (data) {
						//$("#addProspectsToGroup").html("success!");
						swal('Email(s) Successfully Archived');
						window.location.href = "allSentEmails.asp?cat1ID=" + cat1 + "#" + cat2;
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
				//$.ajax({		
					//type:"POST",
					//data: "i="+chkBoxEmailIDArray+"&cat1="+cat1+"&cat2="+cat2,
					//url: "forwardEmailFromTabView.asp",
					//success: function (data) {
						//$("#addProspectsToGroup").html("success!");
						//swal('Email(s) Successfully Forwarded');
						//window.location.href = "allSentEmails.asp?cat1ID=" + cat1 + "#" + cat2;
					//}
				//})	
				
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
						window.location.href = "allSentEmails.asp?cat1ID=" + cat1 + "#" + cat2 + "&cid=" + clientid;
					}
				})	
	    
	    });	
	    


		//Functionality for check/uncheck all checkbox
		$("#checkAll").change(function (activeTab) {
		
			$("#activeChk").children(':first').each(function() {
			    alert($(this).attr("id"));
			});
			
		    $("#emailTabLevel2.active input:checkbox").prop('checked', $(this).prop("checked"));
		    //$("[input:checkbox]").prop('checked', $(this).prop("checked"));
		});
		
		
		$('input[name="chkEmail"]:checked').each(function() {
		   console.log(this.value);
		});		
	
		//function that gets the value of the tab when it is clicked and then
		//updates the value of a hidden form field so when the page posts, it returns
		//back to the tab that was previously opened
		
		//THIS IS SPECIAL TAB CODE FOR EMAILS ONLY - WILL NOT WORK ON OTHER INSIGHT TABS
		$('a[data-toggle="tab"]').on('shown.bs.tab', function (e) {
			var activeTab = $(".tab-content").find(".active");
			var id = activeTab.attr('id');
			$('input[name="txtTab"]').val(id);
		});

	    $('a[data-toggle="tab"]').on('shown.bs.tab', function(e){
	        var currentTab = $(e.target).attr("href"); // get current tab
	        var LastTab = $(e.relatedTarget).attr("href"); // get last tab
	    });

		
		//search tabs functionality - should only display tabs that contain search terms
		var tabLinks = $('a[data-toggle="tab"]'),
		  tabsContent = $('.tab-content'),
		  tabContent = [],
		  string,
		  i,
		  j;
		
		for (i = 0; i < tabsContent.length; i++) {
		  tabContent[i] = tabsContent.eq(i).text().toLowerCase();
		}
		$('#search').on('input', function() {
		  string = $(this).val().toLowerCase();
		  for (j = 0; j < tabsContent.length; j++) {
		    if (tabContent[j].indexOf(string) > -1) {
		      tabLinks.eq(j).show();
		      tabLinks.eq(j).find('a').tab('show');
		    } else {
		      tabLinks.eq(j).hide();
		    }
		  }
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
	.inbox-head .sr-input {
		border: 1px #ccc solid;
	    border-radius: 4px 0 0 4px;
	    box-shadow: none;
	    color: #8a8a8a;
	    float: left;
	    height: 40px;
	    padding: 0 10px;
	}
	.inbox-head .sr-btn {
	    background: none repeat scroll 0 0 #337ab7;
	    border: medium none;
	    border-radius: 0 4px 4px 0;
	    color: #fff;
	    height: 40px;
	    padding: 0 20px;
	}
	.inbox-head .sr-btn:hover {
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
	
	.fa {
		color: #777;
	}
	
	.inbox-head .sr-btn .fa{
		color: #fff;
	}
	
</style>


<h1 class="page-header"><i class="fa fa-envelope-o"></i> All Sent Emails</h1>


<div class="container-full">
    <div class="row">
        <div class="inbox-head">
          <h3>Sent Items</h3>
          <form action="#" class="pull-right position">
              <div class="input-append">
                  <input type="text" id="search" class="sr-input" placeholder="Search Sent Items...">
                  <button class="btn sr-btn" type="button"><i class="fa fa-search"></i></button>
              </div>
          </form>
      </div>

        <div class="col-sm-3 col-md-2">
            <div class="btn-group">
                <button type="button" class="btn btn-primary dropdown-toggle" data-toggle="dropdown">
                    All Mail <span class="caret"></span>
                </button>
                <ul class="dropdown-menu" role="menu">
                    <li><a href="#">Archived Mail</a></li>
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
            <button type="button" class="btn btn-default" data-toggle="tooltip" title="archive" id="archiveBtn">
                   <span class="fa fa-archive"></span>&nbsp;Archive Selected 
            </button>
            <button type="button" class="btn btn-default" data-toggle="tooltip" title="forward" id="forwardBtn">
                   <span class="fa fa-mail-forward"></span>&nbsp;Forward Selected  
            </button>
            <button type="button" class="btn btn-default" data-toggle="tooltip" title="resend" id="resendBtn">
                   <span class="fa fa-retweet"></span>&nbsp;Resend To Original Recipients
            </button>

                                                            
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
				
				SQL_SentEmail = "SELECT COUNT(EmailCategory1) as catCount,EmailCategory1 FROM SC_EmailLog GROUP BY EmailCategory1 ORDER BY EmailCategory1"
				
				Set rsSentEmail = cnnSentEmail.Execute(SQL_SentEmail)
				
				Response.write("<br>currentEmailCategory1ViewedIDTab: " & currentEmailCategory1ViewedIDTab & "<br>")
				Response.write("Session lastEmailCategoryViewed: " & Session("lastEmailCategoryViewed")& "<br>")
				Response.write("currentEmailCategory1ViewedID : " & currentEmailCategory1ViewedID & "<br>")
				
				IF Not rsSentEmail.EOF Then
					
					categoryCount = 0
					
					Do While NOT rsSentEmail.EOF
					
						'********************************************************************************
						'COMPARE CURRENT/ACTIVE TAB ONE WITH TAB ONE FROM SESSION
						'*********************************************************************************
						
						If Request.Querystring("cat1ID") <> "" AND (Session("lastEmailCategoryViewed") <> rsSentEmail("EmailCategory1")) Then


							currentEmailCategory1ViewedID = rsSentEmail("EmailCategory1") 
							currentEmailCategory1ViewedIDTab = Trim(currentEmailCategory1ViewedIDTab)
							currentEmailCategory1ViewedIDTab = Replace(currentEmailCategory1ViewedIDTab," ","")
							currentEmailCategory1ViewedIDTab = LCase(currentEmailCategory1ViewedIDTab)
							Session("lastEmailCategoryViewed") = currentEmailCategory1ViewedID 
							
							Response.write("<br><br><br>currentEmailCategory1ViewedIDTab: " & currentEmailCategory1ViewedIDTab & "<br>")
							Response.write("Session lastEmailCategoryViewed: " & Session("lastEmailCategoryViewed")& "<br>")
							Response.write("currentEmailCategory1ViewedID : " & currentEmailCategory1ViewedID & "<br>")

							Response.Write("111111")
							
						End If
													
						If currentEmailCategory1ViewedIDTab = "" Then
							DefView = currentEmailCategory1ViewedID 
							Response.Write("2222")
						End If
						
						'**********************************************************************
						'TO CREATE TAB NAME, REMOVE ALL SPACES FROM SQL FIELD NAME
						'**********************************************************************
						emailCategory1TabName = Trim(rsSentEmail("EmailCategory1"))
						emailCategory1TabName = Replace(emailCategory1TabName," ","")
						emailCategory1TabName = LCase(emailCategory1TabName)
						
						%>
			                <li id="<%= emailCategory1TabName %>" <% If (currentEmailCategory1ViewedIDTab = emailCategory1TabName) OR (currentEmailCategory1ViewedIDTab = "" AND categoryCount = 0) Then Response.write("class='active'") %>><a href="allSentEmails.asp?cat1ID=<%= rsSentEmail("EmailCategory1") %>"><span class="badge pull-right"><%= rsSentEmail("catCount") %></span> <%= rsSentEmail("EmailCategory1") %></a></li>
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
            <a href="allFailedEmails.asp" class="btn btn-danger btn-sm btn-block" role="button">Emails That Failed To Send</a>
            
            
        </div>
        
        
        <div class="col-sm-9 col-md-10">
        
        	<form method="post" action="allSentEmails.asp" name="frmAllSentEmails" id="frmAllSentEmails">

			<input type="hidden" name="txtTab" id="txtTab" value="">
			<input type="hidden" name="txtClientID" id="txtClientID" value="<%= MUV_Read("ClientID") %>">
			
            <!-- Nav tabs -->
            <ul class="nav nav-tabs" role="tablist" id="emailtabs">
            <%
            
            	Set cnnSentEmail = Server.CreateObject("ADODB.Connection")
				cnnSentEmail.open (Session("ClientCnnString"))
				Set rsSentEmail = Server.CreateObject("ADODB.Recordset")
				rsSentEmail.CursorLocation = 3 
				
				If  Request.Querystring("cat1ID") = "" Then 
					SQL_SentEmail = "SELECT EmailCategory2 FROM SC_EmailLog WHERE EmailCategory1 = '" &  DefView & "' GROUP BY EmailCategory2 ORDER BY EmailCategory2"
				Else
					SQL_SentEmail = "SELECT EmailCategory2 FROM SC_EmailLog WHERE EmailCategory1 = '" & currentEmailCategory1ViewedID & "' GROUP BY EmailCategory2 ORDER BY EmailCategory2"
				End IF
				
				Response.write(SQL_SentEmail)

				Set rsSentEmail = cnnSentEmail.Execute(SQL_SentEmail)
				
				IF Not rsSentEmail.EOF Then
				
				
				categoryTabCount = 0
				
				Do While NOT rsSentEmail.EOF
				
					If categoryTabCount = 0 AND currentEmailCategory2ViewedIDTab = "" Then
						currentEmailCategory2ViewedIDTab = rsSentEmail("EmailCategory2") 
						currentEmailCategory2ViewedIDTabWithSpaces = rsSentEmail("EmailCategory2")
						currentEmailCategory2ViewedIDTab = Trim(currentEmailCategory2ViewedIDTab)
						currentEmailCategory2ViewedIDTab = Replace(currentEmailCategory2ViewedIDTab," ","")
						currentEmailCategory2ViewedIDTab = LCase(currentEmailCategory2ViewedIDTab)
					End If
					
					'**********************************************************************
					'TO CREATE TAB NAME, REMOVE ALL SPACES FROM SQL FIELD NAME
					'**********************************************************************
					emailCategory2TabName = Trim(rsSentEmail("EmailCategory2"))
					emailCategory2TabName = Replace(emailCategory2TabName," ","")
					emailCategory2TabName = LCase(emailCategory2TabName)
					%>
		                
		                <li id="emailTabLevel2" role="presentation" <% If currentEmailCategory2ViewedIDTab = emailCategory2TabName Then Response.write("class='active'") %>>
		                <a href="#<%= emailCategory2TabName %>" role="tab" data-toggle="tab"><span class="glyphicon glyphicon-inbox"></span><%= rsSentEmail("EmailCategory2") %></a>
		                <!--<a href="allSentEmails.asp?cat1ID=<%= currentEmailCategory1ViewedID %>#<%= emailCategory2TabName %>"><span class="glyphicon glyphicon-inbox"></span><%= rsSentEmail("EmailCategory2") %></a>-->
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
				
				If  Request.Querystring("cat1ID") = "" Then 
					SQL_SentEmail = "SELECT * FROM SC_EmailLog WHERE EmailCategory1 = '" & DefView & "' ORDER BY EmailCategory2"
				Else
					SQL_SentEmail = "SELECT * FROM SC_EmailLog WHERE EmailCategory1 = '" & currentEmailCategory1ViewedID & "' ORDER BY EmailCategory2"				
				End IF
				
				
				Set rsSentEmail = cnnSentEmail.Execute(SQL_SentEmail)
				
				IF Not rsSentEmail.EOF Then
				
					oldTab2ID = ""
					currentTab2ID = ""
					
					Do While NOT rsSentEmail.EOF
								
									
						currentTab2ID = rsSentEmail("EmailCategory2")
							
						'**********************************************************************
						'TO CREATE TAB PANLE ID, REMOVE ALL SPACES FROM SQL FIELD NAME
						'**********************************************************************
						emailCategory2LowerTabName = Trim(rsSentEmail("EmailCategory2"))
						emailCategory2LowerTabName = Replace(emailCategory2LowerTabName," ","")
						emailCategory2LowerTabName = LCase(emailCategory2LowerTabName)
		
							If currentTab2ID <> oldTab2ID Then
								%>
								<div role="tabpanel" class="tab-pane fade in <% If currentEmailCategory2ViewedIDTab = emailCategory2LowerTabName Then Response.write("active") %>" id="<%= emailCategory2LowerTabName %>">
			            		<div class="list-group">
								<%
							End If
					
							 InternalRecordNumber = rsSentEmail("InternalRecordNumber")
							 RecordCreationDateTime = rsSentEmail("RecordCreationDateTime")
							 EmailDate = FormatDateTime(rsSentEmail("EmailDate"),2)
							 EmailTime = rsSentEmail("EmailTime")
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
										<label <% If currentEmailCategory2ViewedIDTab = emailCategory2LowerTabName Then Response.write("id='activeChk'") %>>
											<input type="checkbox" name="chkEmail" id="<%= InternalRecordNumber %>">
										</label>
									</div>
									<!--<span class="glyphicon glyphicon-star-empty"></span>-->
									<span class="name" style="min-width: 120px; display: inline-block;">TO: <%= EmailSendTo %></span> 
									<span class=""><strong><a data-toggle="modal" data-show="true" data-target="#myEmailModal<%= InternalRecordNumber %>" href="displayFullEmailModal.asp?i=<%= InternalRecordNumber %>&cat1=<%= currentEmailCategory1ViewedIDTab %>&cat2=<%= currentEmailCategory2ViewedIDTab %>"><%= Subject %></a></strong></span>
									<span class="text-muted" style="font-size: 11px;">- <%= Left(Body,50) %></span> 
									
									<% If dateDiff("d",EmailDate,Now()) <= 1 Then %>
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
		           		
		           		If (currentTab2ID <> oldTab2ID) AND oldTab2ID <> "" Then %>                         
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
        </div>
    </div>
</div>

							
<!--#include file="../../inc/footer-main.asp"-->