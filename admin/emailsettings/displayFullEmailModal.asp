<!--#include file="../../inc/settings.asp"-->
<!--#include file="../../inc/InSightFuncs.asp"-->
<!--#include file="../../inc/InSightFuncs_Users.asp"-->
<% 

InternalRecordNumber = Request.QueryString("i") 
currentEmailCategory1ViewedID = Request.QueryString("cat1")
currentEmailCategory2ViewedIDTab = Request.QueryString("cat2")
ClientID = Request.QueryString("cid")

%>

<script language="javascript">

	$(document).ready(function() {   

	    $('#archiveFromModal').click(function() {
	    
	    		emailID = $("input[name='txtInternalRecordNumber']").val();
	    		cat1 = $("input[name='txtCat1']").val();
	    		cat2 = $("input[name='txtCat2']").val();
	    		
				//post all obtained values to processing ASP page
				$.ajax({		
					type:"POST",
					data: "i="+emailID+"&cat1="+cat1+"&cat2="+cat2,
					url: "archiveEmailFromModal.asp",
					success: function (data) {
			              $("#result").html('Email was successfully archived!'); 
			              $("#result").addClass("alert alert-success");			
					}//end success function
				});//end ajax post
	    
	    });//end click function
	    
	    
	    $('#unarchiveFromModal').click(function() {
	    
	    		emailID = $("input[name='txtInternalRecordNumber']").val();
	    		cat1 = $("input[name='txtCat1']").val();
	    		cat2 = $("input[name='txtCat2']").val();
	    
				//post all obtained values to processing ASP page
				$.ajax({		
					type:"POST",
					data: "i="+emailID+"&cat1="+cat1+"&cat2="+cat2,
					url: "unarchiveEmailFromModal.asp",
					success: function (data) {
			              $("#result").html('Email was successfully unarchived!'); 
			              $("#result").addClass("alert alert-success");								
					}//end success function
				});//end ajax post
	    
	    });//end click function
	    
	    
	    
	    $('#forwardFromModal').click(function() {
	    
	    		emailID = $("input[name='txtInternalRecordNumber']").val();
	    		cat1 = $("input[name='txtCat1']").val();
	    		cat2 = $("input[name='txtCat2']").val();
	    		clientid = $("#txtClientID").val();
	    		
	    		window.location.href = "forwardEmailFromModal.asp?i="+emailID+"&cat1ID=" + cat1 + "&cat2=" + cat2 + "&cid=" + clientid;
	    
	    });//end click function
	    
  
	    
	    $('#resendFromModal').click(function() {
	    
	    		emailID = $("input[name='txtInternalRecordNumber']").val();
	    		cat1 = $("input[name='txtCat1']").val();
	    		cat2 = $("input[name='txtCat2']").val();
	    		clientid = $("#txtClientID").val();
	    
				//post all obtained values to processing ASP page
				$.ajax({		
					type:"POST",
					data: "i="+emailID+"&cat1="+cat1+"&cat2="+cat2 + "&cid=" + clientid,
					url: "resendEmailFromModal.asp",
					success: function (data) {
						  //$("#result").html(data);
			              $("#result").html('Email was successfully re-sent!'); 
			              $("#result").addClass("alert alert-success");						
					}//end success function
				});//end ajax post
	    
	    });//end click function
		
});//end document.ready
</script>


<style type="text/css">

.modal-footer{
	margin-top:15px;
}

.modal-header {
    padding: 15px;
    border-bottom:0px;
}

.email .message .btn-group{
	margin-top:20px;
}
.email .message .message-title {
    margin-top: 10px;
    padding-top: 10px;
    font-weight: 700;
    font-size: 14px
}

.email .message .header {
    margin: 20px 0 30px 0;
    padding: 10px 0 10px 0;
    border-top: 1px solid #d1d4d7;
    border-bottom: 1px solid #d1d4d7
}

.email .message .header .avatar {
    -webkit-border-radius: 2px;
    -moz-border-radius: 2px;
    border-radius: 2px;
    height: 34px;
    width: 34px;
    float: left;
    margin-right: 10px
}

.email .message .header i {
    margin-top: 1px
}

.email .message .header .from {
    display: inline-block;
    width: 100%;
    font-size: 12px;
    margin-top: -2px;
    color: #aaa;
}

.email .message .header .from span {
    display: inline-block;
    font-size: 14px;
    font-weight: 700;
    color: #444;
}

.email .message .header .date {
    display: inline-block;
    width: 100%;
    text-align: right;
    float: right;
    font-size: 12px;
    margin-top: 18px;
}

.email .message .content{
    width:400px;
}

.email .message .attachments {
    border-top: 3px solid #e4e5e6;
    border-bottom: 3px solid #e4e5e6;
    padding: 10px 0;
    margin-bottom: 20px;
    font-size: 12px
}

.email .message .attachments ul {
    list-style: none;
    margin: 0 0 0 -40px
}

.email .message .attachments ul li {
    margin: 10px 0
}

.email .message .attachments ul li .label {
    padding: 2px 4px
}

.email .message .attachments ul li span.quickMenu {
    float: right;
    text-align: right
}

.email .message .attachments ul li span.quickMenu .fa {
    padding: 5px 0 5px 25px;
    font-size: 14px;
    margin: -2px 0 0 5px;
    color: #d1d4d7
}

</style>

<div class="col-lg-12">
	<div class="modal-header">
		<button type="button" class="close" data-dismiss="modal" aria-hidden="true">&times;</button>
	</div>
</div>

<%      
	
	Function stripHTML(strHTML)
	'Strips the HTML tags from strHTML
	
	  Dim objRegExp, strOutput
	  Set objRegExp = New Regexp
	
	  objRegExp.IgnoreCase = True
	  objRegExp.Global = True
	  objRegExp.Pattern = "<(.|\n)+?>"
	
	  'Replace all HTML tag matches with the empty string
	  strOutput = objRegExp.Replace(strHTML, "")
	  
	  'Replace all < and > with &lt; and &gt;
	  strOutput = Replace(strOutput, "<", "&lt;")
	  strOutput = Replace(strOutput, ">", "&gt;")
	  
	  stripHTML = strOutput    'Return the value of strOutput
	
	  Set objRegExp = Nothing
	End Function

 
	Set cnnSentEmailModal = Server.CreateObject("ADODB.Connection")
	cnnSentEmailModal.open (Session("ClientCnnString"))
	Set rsSentEmailModal = Server.CreateObject("ADODB.Recordset")
	rsSentEmailModal.CursorLocation = 3 
	
	SQL_SentEmailModal = "SELECT * FROM SC_EmailLog WHERE InternalRecordNumber = '" & InternalRecordNumber & "'"
	
	
	Set rsSentEmailModal = cnnSentEmailModal.Execute(SQL_SentEmailModal)
	
	IF Not rsSentEmailModal.EOF Then
	
		 RecordCreationDateTime = rsSentEmailModal("RecordCreationDateTime")
		 EmailDate = FormatDateTime(rsSentEmailModal("EmailDate"),2)
		 EmailTime = FormatDateTime(rsSentEmailModal("EmailTime"),3)
		 EmailTo = rsSentEmailModal("EmailTo")
		 EmailFrom = rsSentEmailModal("EmailFrom")
		 EmailFromName = rsSentEmailModal("EmailFromName")
		 Subject = rsSentEmailModal("Subject")
		 Body = stripHTML(rsSentEmailModal("Body"))
		 Body = rsSentEmailModal("Body")
		 CCs = rsSentEmailModal("CCs")
		 BCCs = rsSentEmailModal("BCCs")
		 Attachment = rsSentEmailModal("Attachment")
		 ASPMailStatus = rsSentEmailModal("ASPMailStatus")
		 Archived = rsSentEmailModal("Archived")

	End If
	

	set rsSentEmailModal = Nothing
	cnnSentEmailModal.close
	set cnnSentEmailModal = Nothing


%>

<div class="col-lg-12">

	<div class="email">
		
		<div class="panel panel-default">
			
			<div class="panel-body message">

				<div id="result"></div>
				
				<% If Archived = 0 Then %>
					<div class="message-title"><%= Subject %></div>
				<% Else %>
					<div class="message-title">[ARCHIVED] <%= Subject %></div>
				<% End If %>
			
				
				<% If Attachment <> "" Then %>
				
					<div class="date"><span class="fa fa-paper-clip"></span>
						<% If Abs(dateDiff("d",EmailDate,Now())) = 1 Then %>
							Today, <strong><%= EmailTime %></strong> 
						<% Else %>
							<%= EmailDate %>, <%= EmailTime %>
						<% End If %>
					 </div>
					
				<% Else %>
					<div class="date">
						<% If dateDiff("d",EmailDate,Now()) <= 1 Then %>
							Today, <strong><%= EmailTime %></strong> 
						<% Else %>
							<%= EmailDate %>, <%= EmailTime %>
						<% End If %>
					 </div>
				<% End If %>

				
				
				<input type="hidden" name="txtInternalRecordNumber" value="<%= InternalRecordNumber %>">
				<input type="hidden" name="txtCat1" value="<%= currentEmailCategory1ViewedID %>">
				<input type="hidden" name="txtCat2" value="<%= currentEmailCategory2ViewedIDTab %>">
				<input type="hidden" name="txtClientID" id="txtClientID" value="<%= ClientID %>">

				<span class="btn-group">
				  	<!--<button class="btn btn-default"><span class="fa fa-mail-reply"></span></button>
				  	<button class="btn btn-default"><span class="fa fa-mail-reply-all"></span></button>-->
				  	<% If Archived = 0 Then %>
				  		<button class="btn btn-default" id="archiveFromModal" type="button"><span class="fa fa-archive"></span> Archive</button>
				  	<% Else %>
				  		<button class="btn btn-default" id="unarchiveFromModal" type="button"><span class="fa fa-archive"></span> Un-Archive</button>
				  	<% End If %>
				  	<button class="btn btn-default" id="forwardFromModal" type="button"><span class="fa fa-mail-forward"></span> Forward</button>
				  	<button class="btn btn-default" id="resendFromModal" type="button"><span class="fa fa-retweet"></span> Re-Send</button>

				</span>
				
				<div class="header">

					<div class="from">
						<span>From:</span> <%= EmailFromName %> [<%= EmailFrom %>]
					</div>
					<div class="from">
						<span>To:</span> <%= EmailTo %><br>
						<% If CCs <> "" Then %>
							<span>CC:</span> <%= CCs %><br>
						<% End If %>
						<% If BCCs <> "" Then %>
							<span>BCC:</span> <%= BCCs %>
						<% End If %>
					</div>

					<div class="menu"></div>

				</div>

				<div class="content">
					<p>
						<%= body %>
					</p>
				</div>

				
				<% If Attachment <> "" Then %>
				
				<%
					fileName = Mid(Attachment, InStrRev(Attachment, "\") + 1)
					fileExtention = Right(Attachment,3)
					fileDownloadURL = BaseURL & Replace(Attachment,"\","/")
				
				%>
					<div class="attachments">
						<ul>
							<li>
								<span class="label label-success"><%= fileExtention %></span> <b><%= fileName %></b> <!--<i>(984KB)</i>-->
								<span class="quickMenu">
									<a href="#" class="fa fa-share"><i></i></a>
									<a href="<%= fileDownloadURL %>" class="fa fa-cloud-download" target="_blank"><i></i></a>
								</span>
							</li>
						</ul>		
					</div>
				<% End If %>
					
			</div>	
		</div>	
	</div>		
</div><!--/.col-->	
	
<div class="col-lg-12">
	<div class="modal-footer">
		<button type="button" class="btn btn-default" data-dismiss="modal">Close Email View Window</button>
	</div>
</div>






