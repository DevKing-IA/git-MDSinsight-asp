<!--#include file="../inc/settings.asp"-->
<!--#include file="../inc/InSightFuncs.asp"-->
<!--#include file="../inc/InSightFuncs_Prospecting.asp"-->
<!--#include file="../inc/InSightFuncs_Users.asp"-->
<% 

InternalRecordNumber = Request.QueryString("i") 

%>


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
	
	SQL_SentEmailModal = "SELECT * FROM PR_ProspectEmailLog WHERE InternalRecordIdentifier= '" & InternalRecordNumber & "'"
	

	Set rsSentEmailModal = cnnSentEmailModal.Execute(SQL_SentEmailModal)
	
	IF Not rsSentEmailModal.EOF Then
		 
		If not rsSentEmailModal.EOF Then
			EmailTo = rsSentEmailModal("to_addr")
			EmailFrom = rsSentEmailModal("from_addr")
			EmailCC	= rsSentEmailModal("cc_addr")
			EmailBcc = rsSentEmailModal("bcc_addr")
			EmailDateTime = rsSentEmailModal("EmailDateTime")
			EmailSubject = rsSentEmailModal("sub")
			EmailBodyHTML  = rsSentEmailModal("body_html")
			EmailBodyText  = rsSentEmailModal("body_text")
			EmailBodyAttachment  = rsSentEmailModal("attach_count")
		End If
				

	End If

	EmailDateTime = cDate(EmailDateTime)
	EmailSentDate = FormatDateTime(EmailDateTime,2)
	EmailSentTime = FormatDateTime(EmailDateTime,3)


	set rsSentEmailModal = Nothing
	cnnSentEmailModal.close
	set cnnSentEmailModal = Nothing


%>

<div class="col-lg-12">

	<div class="email">
		
		<div class="panel panel-default">
			
			<div class="panel-body message">

				<div id="result"></div>
				

				<div class="message-title">Subject: <%= EmailSubject %></div>
			
				
				<% If EmailBodyAttachment <> 0 Then %>
				
					<div class="date"><span class="fa fa-paper-clip"></span>
						<% If Abs(dateDiff("d",EmailSentDate,Now())) = 1 Then %>
							Today, <strong><%= EmailSentTime %></strong> 
						<% Else %>
							<%= EmailSentDate %>, <%= EmailSentTime %>
						<% End If %>
					 </div>
					
				<% Else %>
					<div class="date">
						<% If dateDiff("d",EmailSentDate,Now()) <= 1 Then %>
							Today, <strong><%= EmailSentTime %></strong> 
						<% Else %>
							<%= EmailSentDate %>, <%= EmailSentTime %>
						<% End If %>
					 </div>
				<% End If %>

	
				<input type="hidden" name="txtInternalRecordNumber" value="<%= InternalRecordNumber %>">
				
				<div class="header">

					<div class="from">
						<span>From:</span> <a href="mailto:<%= EmailFrom %>"><%= EmailFrom %></a>
					</div>
					<div class="from">
						<span>To:</span> <a href="mailto:<%= EmailTo %>"><%= EmailTo %></a><br>
						<% If EmailCC <> "" Then %>
							<span>CC:</span> <%= EmailCC %><br>
						<% End If %>
						<% If EmailBcc <> "" Then %>
							<span>BCC:</span> <%= EmailBcc %>
						<% End If %>
					</div>

					<div class="menu"></div>

				</div>

				<div class="content">
					<p>
						<% If EmailBodyHTML <> "" Then
								Response.Write(EmailBodyHTML)  
							Else
								Response.Write(EmailBodyText)
							End If
						 %>
					</p>
				</div>

				
				<% If EmailBodyAttachment <> 0 Then %>
				
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






