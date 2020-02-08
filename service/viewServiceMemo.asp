<!--#include file="../inc/header.asp"-->
 
<!-- Add fancyBox main JS and CSS files -->
<script type="text/javascript" src="<%= BaseURL %>js/jquery-lightbox/jquery.fancybox.js?v=2.1.5"></script>
<link rel="stylesheet" href="<%= BaseURL %>js/jquery-lightbox/jquery.fancybox.css?v=2.1.5" type="text/css" media="screen" />


<% MemoNumber = Request.QueryString("memo") 
If MemoNumber = "" Then Response.Redirect(BaseURL)
%>


	<script type="text/javascript">
		$(document).ready(function() {
			/*
			 *  Simple image gallery. Uses default settings
			 */

			$('.fancybox').fancybox();
			$('.fancybox-signature').fancybox();



		});
	</script>

<link rel="stylesheet" type="text/css" href="<%= BaseURL %>css/tabs.css" />


<style>

	.fancybox-custom .fancybox-skin {
		box-shadow: 0 0 50px #222;
	}

	.thumbnail{
		max-width: 100px;
		max-height: 100px;
		display: inline;
		margin-left:15px;
	}

	.thumbnail-signature{
		max-width: 300px;
		max-height: 200px;
		display: inline;
		margin-left:15px;
		border:1px teal dashed;
	}
	
	.thumbnail-photos{
		float: left;
		margin: 10px;
		margin-top:20px;
	}
	
	#lightbox .modal-content {
    display: inline-block;
    text-align: center;   
	}

	#lightbox .close {
	    opacity: 1;
	    color: rgb(255, 255, 255);
	    background-color: rgb(25, 25, 25);
	    padding: 5px 8px;
	    border-radius: 30px;
	    border: 2px solid rgb(255, 255, 255);
	    position: absolute;
	    top: -15px;
	    right: -55px;
	    
	    z-index:1032;
	}

 	.alert{
 		padding: 6px 12px;
 		margin-bottom: 0px;
	}
	
	.form-control{
		margin-bottom: 20px;
	}
	
	a:hover{
		text-decoration: none;
	}
	
	[class^="col-"]{
	 margin-bottom:25px;
  } 
  
  .custom-hr{
	height: 3px;
	margin-left: auto;
	margin-right: auto;
	background-color:#183049;
	color:#183049;
	border: 0 none;
	}
	
	.control-label{
		padding-top: 5px;
	}
  
	</style>


<h1 class="page-header"><i class="fa fa-wrench"></i> View Service Ticket</h1>

<%
'*************
'OPEN info
SQL = "SELECT * FROM FS_ServiceMemos where MemoNumber = '" & MemoNumber  & "' And RecordSubType='OPEN'"
	
Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3 
Set rs = cnn8.Execute(SQL)
	
If not rs.EOF Then
	OpenServiceMemoRecNumber = rs("ServiceMemoRecNumber")
	OpenCurrentStatus = rs("CurrentStatus")
	OpenRecordSubType = rs("RecordSubType")
	OpenSubmittedByName = rs("SubmittedByName")
	OpenAccountNumber = rs("AccountNumber")
	OpenCompany = rs("Company")
	OpenProblemLocation = rs("ProblemLocation")
	OpenSubmittedByPhone = rs("SubmittedByPhone")
	OpenSubmittedByEmail = rs("SubmittedByEmail")
	OpenSubmissionDateTime = rs("SubmissionDateTime")
	OpenProblemDescription = rs("ProblemDescription")
	OpenMode = rs("Mode")
	OpenSubmissionSource = rs("SubmissionSource")
	OpenUserNoOfServiceTech = rs("UserNoOfServiceTech")
	ReleasedDateTime = rs("ReleasedDateTime")
	ReleasedByUserNo = rs("ReleasedByUserNo")
	ReleasedNotes = rs("ReleasedNotes")

Else
	SQL = "SELECT * FROM FS_ServiceMemos where MemoNumber = '" & MemoNumber  & "' And RecordSubType='HOLD'"
	Set rs = cnn8.Execute(SQL)
	If not rs.EOF Then
		OpenServiceMemoRecNumber = rs("ServiceMemoRecNumber")
		OpenCurrentStatus = rs("CurrentStatus")
		OpenRecordSubType = rs("RecordSubType")
		OpenSubmittedByName = rs("SubmittedByName")
		OpenAccountNumber = rs("AccountNumber")
		OpenCompany = rs("Company")
		OpenProblemLocation = rs("ProblemLocation")
		OpenSubmittedByPhone = rs("SubmittedByPhone")
		OpenSubmittedByEmail = rs("SubmittedByEmail")
		OpenSubmissionDateTime = rs("SubmissionDateTime")
		OpenProblemDescription = rs("ProblemDescription")
		OpenMode = rs("Mode")
		OpenSubmissionSource = rs("SubmissionSource")
		OpenUserNoOfServiceTech = rs("UserNoOfServiceTech")
	End IF
End If
	
set rs = Nothing
cnn8.close
set cnn8 = Nothing

If OpenSubmittedByName = "" Then OpenSubmittedByName = "Not provided"
If OpenSubmittedByPhone = "" Then OpenSubmittedByPhone = "Not provided"
If OpenSubmittedByEmail = "" Then OpenSubmittedByEmail = "Not provided"
If OpenProblemLocation = "" Then OpenProblemLocation = "Not provided"
If OpenProblemDescription = "" Then OpenProblemDescription = "Not provided"

%>


	
<%
'*************
'CloseCancel info
SQL = "SELECT * FROM FS_ServiceMemos where MemoNumber = '" & MemoNumber  & "' And RecordSubType<>'OPEN'"
	
Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3 
Set rs = cnn8.Execute(SQL)
	
If not rs.EOF Then
	CCServiceMemoRecNumber = rs("ServiceMemoRecNumber")
	CCCurrentStatus = rs("CurrentStatus")
	CCRecordSubType = rs("RecordSubType")
	CCSubmittedByName = rs("SubmittedByName")
	CCAccountNumber = rs("AccountNumber")
	CCCompany = rs("Company")
	CCProblemLocation = rs("ProblemLocation")
	CCSubmittedByPhone = rs("SubmittedByPhone")
	CCSubmittedByEmail = rs("SubmittedByEmail")
	CCSubmissionDateTime = rs("SubmissionDateTime")
	CCProblemDescription = rs("ProblemDescription")
	CCMode = rs("Mode")
	CCSubmissionSource = rs("SubmissionSource")
	CCUserNoOfServiceTech = rs("UserNoOfServiceTech")
End If
	
set rs = Nothing
cnn8.close
set cnn8 = Nothing

If CCSubmittedByName = "" Then CCSubmittedByName = "Not provided"
If CCSubmittedByPhone = "" Then CCSubmittedByPhone = "Not provided"
If CCSubmittedByEmail = "" Then CCSubmittedByEmail = "Not provided"
If CCProblemLocation = "" Then CCProblemLocation = "Not provided"
If CCProblemDescription = "" Then CCProblemDescription = "Not provided"

%>

      

        <input type="hidden" id="txtPrintedName" name="txtPrintedName" value="N/A Closed From MDS Insight"  class="form-control last-run-inputs">

        
 	        <!-- row !-->		
	        <div class="row ">
		        

		        <!--account number !-->
		        <div class="col-lg-6 col-md-4 col-sm-12 col-xs-12">
		        	<%SelectedCustomer = OpenAccountNumber %>
					<!--#include file="../inc/commonCustomerDisplay.asp"-->
			    </div>
		        <!-- eof account number !-->
		        
		        <!-- company name !-->
		        <div class="col-lg-4 col-md-4 col-sm-12 col-xs-12">
			        <div class="alert alert-info" role="alert"><strong>Ticket#: <%= MemoNumber %></strong></div>
 		        </div>
		        <!-- eof company name !-->

						        		
		        </div>
 <!-- eof row !-->

 <!-- main row !-->
 <div class="row">
	 
	 <!-- left col !-->
	 <div class="col-lg-4 col-md-4 col-sm-12 col-xs-12">
	
	 
 		        <!-- row !-->			
			    <div class="row">

			    	<!-- Contact Name !-->
			    <div class="col-lg-12">
			        <strong>Contact Name</strong><br>
			        <% =OpenSubmittedByName %>
			        </div>
			    	<!-- Contact Name !-->
	
			    	<!-- Contact Phone !-->
			    	  <div class="col-lg-12">
				    	 <strong>Contact Phone</strong><br>
				    	 <% =OpenSubmittedByPhone %>
			        </div>
			    	<!-- Contact Phone !-->
			    	
			    		<!-- Contact Email !-->
			    	  <div class="col-lg-12">
			        <strong>Contact Email</strong><br>
					<% =OpenSubmittedByEmail %>
			        </div>
			    	<!-- Contact Email !-->
			    	
			    	</div>
			    <!-- eof row !-->
			    	
			    	   
	 </div>
		    	 <!-- eof left col !-->
		    	 
		    	 
 		       
  	
  			<!-- right col !-->
  			<div class="col-lg-4 col-md-4 col-sm-12 col-xs-12">  
	  			<div class="row">
		  					
					<!-- Problem Location !-->
					<div class="col-lg-12">
						<strong>Problem Location</strong><br>
						<%= OpenProblemLocation %>
					</div>
					<!-- Problem Location !-->
					    	
					<!-- Description of problem !-->
					<div class="col-lg-12">
						<strong>Problem Description</strong><br>
						<%= OpenProblemDescription %>
					</div>
					<!-- Description of problem !-->
			

					<!-- Signature !-->
					<div class="col-lg-12">
					
						<% If GetServiceTicketStatus(MemoNumber) = "CLOSE" Then 
	
								'----------------------------
								'Service Signature Check
								'----------------------------
								set fs = CreateObject("Scripting.FileSystemObject")
								Pth =  "../clientfiles/" & trim(MUV_Read("ClientID")) & "/signaturesave/TicketID-" & Trim(MemoNumber) & ".png"
								
								If fs.FileExists(Server.MapPath(Pth)) Then
									hasServiceSignature = True
								Else
									hasServiceSignature = False
								End If
													
								'Response.Write(Pth)
								
								'***************************************************************************************************
								'Display signature file, if any exist in the signaturesave directory
								''Check for the existance of a thumbnail image in the directory, otherwise, size the image with CSS
								'***************************************************************************************************
				
								Pth =  "../clientfiles/" & trim(MUV_Read("ClientID")) & "/signaturesave/TicketID-" & Trim(MemoNumber) & ".png"
								PthThumb =  "../clientfiles/" & trim(MUV_Read("ClientID")) & "/signaturesave/TicketID-" & Trim(MemoNumber) & "-thumb.png"
			
								SignaturePathNameFull = BaseURL & "clientfiles/" & trim(MUV_Read("ClientID")) & "/signaturesave/TicketID-" & Trim(MemoNumber) & ".png"
								SignaturePathNameThumb = BaseURL & "clientfiles/" & trim(MUV_Read("ClientID")) & "/signaturesave/TicketID-" & Trim(MemoNumber) & "-thumb.png"
								
								If hasServiceSignature = True Then
									
									If fs.FileExists(Server.MapPath(PthThumb)) Then
								    	%><strong>Signature</strong><br><a href="<%= SignaturePathNameFull %>" target="_blank" style="border:0px;"><img src="<%= SignaturePathNameThumb %>" alt="Ticket <%= Trim(MemoNumber) %> Signature"></a><%
								    Else
								    	%><strong>Signature</strong><br><a href="<%= SignaturePathNameFull %>" target="_blank" style="border:0px;"><img src="<%= SignaturePathNameFull %>" alt="Ticket <%= Trim(MemoNumber) %> Signature" style="width:200px;"></a><%
								    End If
								    
								Else
									 %><strong>No Signature</strong><%
								End If
								
							End If
							set fs=nothing
						%>
						
					</div>
					<!-- Singature !-->
			
	  			</div>
  			</div>
			<!-- eof right col !-->
			
			<% MDG_MemoNumber = MemoNumber %>
			<!--#include file="memo_details_grid.asp"-->
			
			<!-- rightmost col !-->
			<div class="col-lg-4 col-md-4 col-sm-12 col-xs-12">  
				<div class="row">
					
					

				</div>
			</div>
			<!-- eof rightmost col !-->
	         
	         
		</div>
		<!-- eof main row !-->
		
		<div class="row">
			<div class="col-lg-12">
				<hr class="custom-hr">
			</div>
		</div>
		
		<!-- main row !-->
 		<div class="row">
			
			<!-- Photos !-->
			<div class="col-lg-6">
				<% If GetServiceTicketStatus(MemoNumber) = "CLOSE" Then
						z=0
						set fs = CreateObject("Scripting.FileSystemObject")
						For x = 1 to 20 ' Only have 3 pics but allow for expansion to 20 
							Pth =  "../clientfiles/" & trim(MUV_Read("ClientID")) & "/SvcMemoPics/" & Trim(MemoNumber) & "-" & x & ".png"
							Pth2 =  "../clientfiles/" & trim(MUV_Read("ClientID")) & "/SvcMemoPics/" & Trim(MemoNumber) & "-" & x & ".jpg"
							Pth3 =  "../clientfiles/" & trim(MUV_Read("ClientID")) & "/SvcMemoPics/" & Trim(MemoNumber) & "-" & x & ".jpeg"

							If fs.FileExists(Server.MapPath(Pth)) or fs.FileExists(Server.MapPath(Pth2)) or fs.FileExists(Server.MapPath(Pth3)) Then
								If x = 1 Then%><strong>Photos</strong><p class="thumbnail-photos"><% End If
								If fs.FileExists(Server.MapPath(Pth)) Then %><a class="fancybox" href="<%= Pth %>" data-fancybox-group="gallery" title="Service Ticket #<%= Trim(MemoNumber) %>"><img src="<%= Pth %>" alt="" class="thumbnail"></a><% End If													
								If fs.FileExists(Server.MapPath(Pth2)) Then %><a class="fancybox" href="<%= Pth2 %>" data-fancybox-group="gallery" title="Service Ticket #<%= Trim(MemoNumber) %>"><img src="<%= Pth2 %>" alt="" class="thumbnail"></a><% End If
								If fs.FileExists(Server.MapPath(Pth3)) Then %><a class="fancybox" href="<%= Pth3 %>" data-fancybox-group="gallery" title="Service Ticket #<%= Trim(MemoNumber) %>"><img src="<%= Pth3 %>" alt="" class="thumbnail"></a><% End If													
							End If
							If z = 2 then
								%> <%
								z=0
							Else
								z=z+1 'Three per row
							End If
						Next
						%></p><%	
					End If
					set fs=nothing
				%>
			</div>
			<!-- eof Photos  !-->
			
			 
			 <!-- Signature !-->
			 <div class="col-lg-6">
			 	&nbsp;
			</div>
			<!-- eof rightmost col !-->
			
		</div>
		<!-- eof photos row !-->


			<div class="row">
			
			<div class="col-lg-12">	

			    <% If Instr(ucase(Request.ServerVariables ("HTTP_REFERER")),"CUSTOMERSERVICE") <> 0 Then %>
    			    <a href="<%=Request.ServerVariables("HTTP_REFERER")%>">
			    	<button type="button" class="btn btn-default">&lsaquo; Go Back To <%=GetTerm("Customer")%> Notes</button>
			    <% Else %>
	   			    <a href="<%= BaseURL %>service/main.asp">
			    	<button type="button" class="btn btn-default">&lsaquo; Go Back To Service Screen</button>			    
			    <%End IF%>
				</a>
			
					
			</div>
			
 			
			</div>
			<!-- eof row !-->    


<!-- lightbox JS !-->
<div id="lightbox" class="modal fade" tabindex="-1" role="dialog" aria-labelledby="myLargeModalLabel" aria-hidden="true">
    <div class="modal-dialog">
        <button type="button" class="close hidden" data-dismiss="modal" aria-hidden="true">×</button>
        <div class="modal-content">
            <div class="modal-body">
                <img src="" alt="" />
            </div>
        </div>
    </div>
</div>

						<script>
							$(document).ready(function() {
    var $lightbox = $('#lightbox');
    
    $('[data-target="#lightbox"]').on('click', function(event) {
        var $img = $(this).find('img'), 
            src = $img.attr('src'),
            alt = $img.attr('alt'),
            css = {
                'maxWidth': $(window).width() - 100,
                'maxHeight': $(window).height() - 100
            };
    
        $lightbox.find('.close').addClass('hidden');
        $lightbox.find('img').attr('src', src);
        $lightbox.find('img').attr('alt', alt);
        $lightbox.find('img').css(css);
    });
    
    $lightbox.on('shown.bs.modal', function (e) {
        var $img = $lightbox.find('img');
            
        $lightbox.find('.modal-dialog').css({'width': $img.width()});
        $lightbox.find('.close').removeClass('hidden');
    });
});
							</script>
						<!-- eof lightbox JS !-->

 
 

 
   
<!--#include file="../inc/footer-main.asp"-->
