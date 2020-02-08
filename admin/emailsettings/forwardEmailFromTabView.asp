<!--#include file="../../inc/header.asp"-->
<%
	InternalRecordNumber = Request.QueryString("i")
	currentEmailCategory1ViewedIDTab = Request.QueryString("cat1")
	currentEmailCategory2ViewedIDTab = Request.QueryString("cat2")
	ClientID = Request.QueryString("cid")
	
	
	emailReceivedAsArray = false
	
	If InternalRecordNumber <> "" Then

		InternalRecordNumberArray = Split(InternalRecordNumber,",")
		
		If Ubound(InternalRecordNumberArray) = 0 Then
			InternalRecordNumber = InternalRecordNumberArray(0)
		Else
			emailReceivedAsArray = true
		End If
		
	End If

%>


<SCRIPT LANGUAGE="JavaScript">
<!--
	function validateEmail(emailAddress) 
	{
		var multipleAddresses = emailAddress.split(/;/);
		var regularExpression = /^((([a-z]|\d|[!#\$%&'\*\+\-\/=\?\^_`{\|}~]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])+(\.([a-z]|\d|[!#\$%&'\*\+\-\/=\?\^_`{\|}~]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])+)*)|((\x22)((((\x20|\x09)*(\x0d\x0a))?(\x20|\x09)+)?(([\x01-\x08\x0b\x0c\x0e-\x1f\x7f]|\x21|[\x23-\x5b]|[\x5d-\x7e]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(\\([\x01-\x09\x0b\x0c\x0d-\x7f]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF]))))*(((\x20|\x09)*(\x0d\x0a))?(\x20|\x09)+)?(\x22)))@((([a-z]|\d|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(([a-z]|\d|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])*([a-z]|\d|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])))\.)+(([a-z]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(([a-z]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])*([a-z]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF]))){2,6}$/i;

		
		if (multipleAddresses.length <= 1) 
		{
			return regularExpression.test(multipleAddresses[0]);
		}
		else {
			var validEmail = true;
			for (var i=0; i < multipleAddresses.length; i++) 
			{
				if (regularExpression.test(multipleAddresses[i]) == false) {
					validEmail = false;
				}
					
			}
			
			return validEmail;
		}
	} 

   function validateForwardEmailForm()
    {

        if (document.frmForwardEmailAddresses.txtForwardEmailAddresses.value == "") {
            swal("Email address cannot be blank.");
            return false;
        }

        if (validateEmail(document.frmForwardEmailAddresses.txtForwardEmailAddresses.value) == false) {
           swal("One or more email addresses is invalid.");
           return false;
        }

        return true;

    }
// -->
</SCRIPT>   


<!-- password strength meter !-->

<style type="text/css">

.select-line{
	margin-bottom: 15px;
}

.row-line{
	margin-bottom: 25px;
}

.when-col{
	width: 10%;
}

.reference-col{
	width: 45%;
}

.has-more-col{
	width: 12%;
}

.form-control{
	min-width: 300px;
}

.custom-container{
	max-width:800px;
	margin:0 auto;
}

.control-label{
	font-size:12px;
	font-weight:normal;
	padding-top:10px;
}
.control-label-last{
	padding-top:0px;
}

.required{
	border-left:3px solid red;
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
    width:700px;
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
<!-- eof password strength meter !-->

<h1 class="page-header"> Forward Email To Specified Recipients</h1>

<div class="custom-container">

	<form method="POST" action="forwardEmailFromTabView_Submit.asp" name="frmForwardEmailAddresses" id="frmForwardEmailAddresses" onsubmit="return validateForwardEmailForm();">

		<div class="row row-line pull-left">

			<div class="form-group col-lg-12">
				<label for="txtLeadSource" class="col-sm-3 control-label">Foward To:</label>	
    			<div class="col-sm-9">
    				<input type="text" class="form-control required" id="txtForwardEmailAddresses" name="txtForwardEmailAddresses">
    				<input type="hidden" name="txtInternalRecordNumber" id="txtInternalRecordNumber" value="<%= InternalRecordNumber %>">
    				<input type="hidden" name="txtCategory1Active" id="txtCategory1Active" value="<%= currentEmailCategory1ViewedIDTab %>">
    				<input type="hidden" name="txtCategory2Active" id="txtCategory2Active" value="<%= currentEmailCategory2ViewedIDTab %>">
    				<input type="hidden" name="txtCategory2Active" id="txtCategory2Active" value="<%= currentEmailCategory2ViewedIDTab %>">
    				<input type="hidden" name="txtClientID" id="txtClientID" value="<%= ClientID %>">
    			</div>
    			
			</div>
			
			<p><label>To add more than one email address, please separate addresses with a semi-colon (;) character.</label></p>
			
		</div>
		
		
	    <!-- cancel / submit !-->
		<div class="row row-line pull-right">
			<div class="col-lg-12 alertbutton">
				<div class="col-lg-12">
					<a href="<%= BaseURL %>admin/emailsettings/allSentEmails.asp?cat1ID=<%= currentEmailCategory1ViewedID %>&tab=<%= currentEmailCategory2ViewedIDTab %>">
	    				<button type="button" class="btn btn-default">&lsaquo; Cancel &amp; Go Back To All Emails Sent</button>
					</a>
					<button type="submit" class="btn btn-primary"><i class="fa fa-mail-forward"></i> Forward Email</button>
				</div>
		    </div>
		</div>
		
	</form>
	
	<div class="row row-line pull-left">
	<div class="col-lg-12">
	
	
	<% If emailReceivedAsArray = true Then %>
	
	<%
		
		Set cnnSentEmailModal = Server.CreateObject("ADODB.Connection")
		cnnSentEmailModal.open (Session("ClientCnnString"))
		Set rsSentEmailModal = Server.CreateObject("ADODB.Recordset")
		rsSentEmailModal.CursorLocation = 3
		
		%><h1>The following emails will be forwarded:</h1><%
		
		For i = 0 to ubound(InternalRecordNumberArray)

			SQL_SentEmailModal = "SELECT * FROM SC_EmailLog WHERE InternalRecordNumber = " & InternalRecordNumberArray(i)
			
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
		
			End If
		
		
			%>
			
			<div class="email">
			
			<div class="panel panel-default">
				
				<div class="panel-body message">
	
					<div class="message-title"><%= Subject %> sent originally on <%= EmailDate %> at <%= EmailTime %>
						<% If Attachment <> "" Then %>
							with attachments
						<% End If %>
					</div>
				
						
					</div>	<!-- end message -->
				</div>	<!-- end panel -->
			</div>		<!-- end email -->
			
			<%
			
			Next
			
		
			set rsSentEmailModal = Nothing
			cnnSentEmailModal.close
			set cnnSentEmailModal = Nothing
			
		
			%>
		
	<% Else 
		
			Set cnnSentEmailModal = Server.CreateObject("ADODB.Connection")
			cnnSentEmailModal.open (Session("ClientCnnString"))
			Set rsSentEmailModal = Server.CreateObject("ADODB.Recordset")
			rsSentEmailModal.CursorLocation = 3 
	
			SQL_SentEmailModal = "SELECT * FROM SC_EmailLog WHERE InternalRecordNumber = " & InternalRecordNumber
			
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
		
			End If
			
		
			set rsSentEmailModal = Nothing
			cnnSentEmailModal.close
			set cnnSentEmailModal = Nothing
			
			%>

			<h1>Email Preview:</h1>
			<div class="email">
			
			<div class="panel panel-default">
				
				<div class="panel-body message">
	
					<div class="message-title"><%= Subject %></div>
				
					
					<% If Attachment <> "" Then %>
					
						<div class="date"><span class="fa fa-paper-clip"></span>
							<% If dateDiff("d",EmailDate,Now()) <= 1 Then %>
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
									</span>
								</li>
							</ul>		
						</div>
					<% End If %>
						
				</div>	<!-- end message -->
			</div>	<!-- end panel -->
		</div>		<!-- end email -->

	<% End If %>	
		

	
	</div><!-- end col -->
	</div><!-- end row -->
	
	
</div>

<!--#include file="../../inc/footer-main.asp"-->