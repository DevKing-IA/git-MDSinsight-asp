<!--#include file="../../inc/header-edit-email.asp"-->
<!--#include file="../../inc/mail.asp"-->
<script src="<%= baseURL %>js/sweetalert/sweetalert.min.js"></script>
<link rel="stylesheet" type="text/css" href="<%= baseURL %>js/sweetalert/sweetalert.css">

<style type="text/css">
	 
	body{
	 	overflow-x:hidden;
	}
	.page-header{
	 	margin-top: 0px;
	}
	  	  
	h3{
		 margin: 0px;
		 padding: 0px;
		 line-height: 1;
	}
	
	.ui-widget-header{
		background: #193048;
		border: 1px solid #193048;
	}
	
	.custom-row{
	  	margin-top: 10px;
	}
	
	.modal-link{
		cursor: pointer;
	}
	.btnSave{ float:left; width:115px; padding-right:5px; padding-bottom:5px;}
	.btnPreview{ float:left; width:140px; padding-right:20px; padding-bottom:5px;}
	.txtSendTestEmail {width:180px;padding-bottom:5px;float:left;padding-right:5px;vertical-align:bottom}
	.btnSendTestEmail{ float:left; width:140px; padding-right:5px; padding-bottom:5px;}
	.btnCancel{ float:left; width:120px; padding-right:5px; padding-bottom:5px;}
	.btnLoadDefault{ float:left; width:140px; padding-right:20px; padding-bottom:25px;}
	.btnBackToEmails{float:left; width:120px; padding-right:5px; padding-bottom:25px;}
	

/******** Search section ************/

#searchText {
    background-image: url('/css/searchicon.png'); /* Add a search icon to input */
    background-position: 10px 12px; /* Position the search icon */
    background-repeat: no-repeat; /* Do not repeat the icon image */
    width: 400px; /* 400px */
    font-size: 16px; /* Increase font-size */
    padding: 12px 20px 12px 40px; /* Add some padding */
    border: 1px solid #ddd; /* Add a grey border */
    margin-bottom: 12px; /* Add some space below the input */
}

#fieldsList {
    /* Remove default list styling */
    list-style-type: none;
    padding: 0;
    margin: 0;
}

#fieldsList li {
    border: 1px solid #ddd; /* Add a border to all links */
    margin-top: -1px; /* Prevent double borders */
    background-color: #f6f6f6; /* Grey background color */
    padding-top: 5px; /* Add some padding */
    padding-bottom: 25px;
    padding-left:5px;
    font-size: 14px; /* Increase the font-size */
    color: black; /* Add a black text color */
    display: block; /* Make it into a block element to fill the whole list */
}

#fieldsList li div{ width:180px; float:left;}


/******* THE MODAL **********/

/* The Modal (background) */
.modal {
    display: none; /* Hidden by default */
    position: fixed; /* Stay in place */
    z-index: 1; /* Sit on top */
    left: 0;
    top: 0;
    width: 100%; /* Full width */
    height: 100%; /* Full height */
    overflow: auto; /* Enable scroll if needed */
    background-color: rgb(0,0,0); /* Fallback color */
    background-color: rgba(0,0,0,0.4); /* Black w/ opacity */
}

/* Modal Header */
.modal-header {
    padding: 5px 16px;
    margin-bottom:10px;
    background-color: red;
    color: white;
    height:35px;
}

/* Modal Content/Box */
.modal-content {
    background-color: #fefefe;
    margin: 15% auto; /* 15% from the top and centered */
    padding: 20px;
    border: 1px solid #888;
    width: 80%; /* Could be more or less, depending on screen size */
}

/* The Close Button */
.closeModal {
    color: white;
    float: right;
    font-size: 20px;
    font-weight: bold;
}

.closeModal:hover,
.closeModal:focus {
    color: black;
    text-decoration: none;
    cursor: pointer;
}

</style>
<script src="emailDynamicValues.js"></script>
<script>
    function searchFields() {
        // Declare variables
        var input, filter, ul, li, a, i;
        input = document.getElementById('searchText');
        filter = input.value.toUpperCase();
        ul = document.getElementById("fieldsList");
        li = ul.getElementsByTagName('li');

        // Loop through all list items, and hide those who don't match the search query
        for (i = 0; i < li.length; i++) {
            t = li[i];
            if (t.innerText.toUpperCase().indexOf(filter) > -1) {
                li[i].style.display = "";
            } else {
                li[i].style.display = "none";
            }
        }
    }

    function saveEmail() {
        nicEditors.findEditor('editEmailArea').saveContent();
        document.getElementById('frmSaveEmail').submit();
    }
    function sendTestEmail() {
        if (document.getElementById('testEmail').value == "") {
            swal("Please fill in the email address that the test email should be sent to.");
            document.getElementById('testEmail').focus();
            return false;
        }
        if (!validemail(document.getElementById('testEmail').value)) {
            swal("Please type valid email address.");
            document.getElementById('testEmail').focus();
            return false;
        }

        swal({
              title: "",
              text: "Test email sent to " + document.getElementById('testEmail').value,
              type: "success"
            },
            function(isConfirm){
                if (isConfirm) {
                    document.getElementById('action').value = 'send';
                    nicEditors.findEditor('editEmailArea').saveContent();
                    document.getElementById('frmSaveEmail').submit();
                }
            });
            
    }

    function cancelEditEmail() {
        var result = swal("Do you want to leave this page? If you did not save changes - they will be lost", { buttons: true });
        swal({
            title: "Are You Sure?",
            text: "Do you want to leave this page? If you did not save changes - they will be lost",
            type: "warning",
            showCancelButton: true,
            confirmButtonColor: "#DD6B55",
            confirmButtonText: "Yes",
            cancelButtonText: "No",
            closeOnConfirm: true,
            closeOnCancel: true
          },
          function(isConfirm){
            if (isConfirm) {
                location.href='main.asp';
            }
          });

    }

    function previewEmail() {
        var myNicEditor = nicEditors.findEditor('editEmailArea');
        var content = myNicEditor.getContent();

        content = content.replace("{{clientLogo}}", "<img src='<%= baseURL %>clientfilesV/<%= MUV_READ("SERNO")%>/logos/logo_email_header.png' width='650' style='margin-bottom:-5px; margin-left:3px margin-right:3px;'>");

        var newWindow = window.open();
        var styleStr = "<style>.modal-header {padding: 0px;margin-bottom:10px;background-color: red;color: white;height:24px;text-align:center;}</style>";
        var headerStr = "<div class='modal-header'><h3>PLEASE NOTE THAT THIS PREVIEW IS NOT SAVED</h3></div>";
        newWindow.document.write("<html><title>Preview</title><head>" + styleStr + "</head><body>" + headerStr + "<div>" + content + "</div></body></html>");
        newWindow.document.close();
    }


    function validemail(emailVal) {
        var mailformat = /^\w+([\.-]?\w+)*@\w+([\.-]?\w+)*(\.\w{2,3})+$/;
        if (emailVal.match(mailformat)) {
            return true;
        }
        else {
            return false;
        }
    } 
</script>
<br>
<h1 class="page-header"><i class="fa fa-envelope"></i> Edit Email</h1>
<%

Dim emailFileName: emailFileName = Request.QueryString("email")
Dim emailtype: emailtype = Request.QueryString("type") 
dim fs,f, fileContent
set fs=Server.CreateObject("Scripting.FileSystemObject")
Dim action: action = Request.Form("action")
Dim testEmail: testEmail = Request.Form("testEmail") 
fileContent = Request.Form("editEmailArea")
serno = MUV_READ("SERNO") 
Dim forceDefault: forceDefault = False

pathDefault = "C:\home\clientfilesV\default_emails\"
path = "C:\home\clientfilesV\" & serno & "\emails\"


Dim emailFileNameDefault
If InStr(emailFileName,"Default")>0 Then
    emailFileNameDefault = emailFileName
Else
    emailFileNameDefault = Replace(emailFileName,".txt","Default.txt")
End If

If action = "save" Then  
    If InStr(emailFileName,"Default")>0 Then emailFileName = Replace(emailFileName,"Default","") 
    set f=fs.OpenTextFile(path & emailFileName,2,true)
    f.WriteLine(fileContent)
    f.Close
End If

ClientID = MUV_READ("ClientID")
If ucase(right(ClientID,1))="D" then
   ClientID = Left(ClientID,len(ClientID)-1)
End If 

If action = "send" Then
    Dim testEmailBody: testEmailBody = fileContent
    testEmailBody = replaceTestData(testEmailBody, emailFileName, ClientID)
    testEmailFileName = Replace(emailFileName,".txt","")
    If InStr(testEmailFileName,"open")>0 Then testEmailFileName = Replace(testEmailFileName,"open","Open ")
    If InStr(testEmailFileName,"close")>0 Then testEmailFileName = Replace(testEmailFileName,"close","Close ")
    If InStr(testEmailFileName,"cancel")>0 Then testEmailFileName = Replace(testEmailFileName,"cancel","Cancel ")
    If InStr(testEmailFileName,"Service")>0 Then testEmailFileName = Replace(testEmailFileName,"Service","Service ")
    If InStr(testEmailFileName,"Ticket")>0 Then testEmailFileName = Replace(testEmailFileName,"Ticket","Ticket ")
    SendMail "test@ocsaccess.com",testEmail,"Test " & testEmailFileName & " Email",testEmailBody,"",""
Else
    If emailtype = "default" Then
        path = pathDefault
    End If
    If fs.FileExists(path & emailFileName) Then 
       set f=fs.OpenTextFile(path & emailFileName,1)
       fileContent = f.ReadAll()
       f.Close
       set f=Nothing
    Else 'if file doesn't exist
       If InStr(emailFileName,"Default")=0 Then 'If user tried to load custom email and it did not exist
            forceDefault = True
            If fs.FileExists(pathDefault & emailFileNameDefault) Then
               set f=fs.OpenTextFile(pathDefault & emailFileNameDefault,1)
               fileContent = f.ReadAll()
               f.Close
               set f=Nothing
            End If
        End If
    End If
    If emailtype = "default" OR forceDefault = True Then
        SQLSettings = "SELECT * FROM Settings_CompanyID "

        Set cnnSettings = Server.CreateObject("ADODB.Connection")
        cnnSettings.open (Session("ClientCnnString"))
        Set rsSettings = Server.CreateObject("ADODB.Recordset")
        rsSettings.CursorLocation = 3 
        Set rsSettings = cnnSettings.Execute(SQLSettings)

        If not rsSettings.EOF Then
	        CompanyIdentityColor1 = rsSettings("CompanyIdentityColor1")
	        CompanyIdentityColor2 = rsSettings("CompanyIdentityColor2")

            Stmt_CompanyName = rsSettings("Stmt_CompanyName")
            Stmt_Address1 = rsSettings("Stmt_Address1")
            Stmt_Address2 = rsSettings("Stmt_Address2")
            Stmt_City = rsSettings("Stmt_City")
            Stmt_State = rsSettings("Stmt_State")
            Stmt_Zip = rsSettings("Stmt_Zip")

            addressInfo = Stmt_CompanyName & "<br/>"
            addressInfo = addressInfo & Stmt_Address1 & "<br/>"
            If Stmt_Address2<>"" Then  addressInfo = addressInfo & Stmt_Address2 & "<br/>"
            addressInfo = addressInfo & Stmt_City & "," & Stmt_State & "," & Stmt_Zip

            Stmt_Phone1 = rsSettings("Stmt_Phone1")
            Stmt_Phone2 = rsSettings("Stmt_Phone2")
            Stmt_Phone3 = rsSettings("Stmt_Phone3")
            Stmt_Fax = rsSettings("Stmt_Fax")
            Stmt_Email = rsSettings("Stmt_Email")
            phoneEmailInfo = Stmt_Phone1 & "<br/>"
            If Stmt_Phone2<>"" Then  phoneEmailInfo = phoneEmailInfo & Stmt_Phone2 & "<br/>"
            If Stmt_Phone3<>"" Then  phoneEmailInfo = phoneEmailInfo & Stmt_Phone3 & "<br/>"
            If Stmt_Fax<>"" Then  phoneEmailInfo = phoneEmailInfo & Stmt_Fax & "<br/>"
            If Stmt_Email<>"" Then  phoneEmailInfo = phoneEmailInfo & "<a href='mailto:" & Stmt_Email & "' style='color: #fff; text-decoration: none; font-weight: bold;'>" & Stmt_Email & "</a>"
        End If
        set rsSettings = Nothing
        cnnSettings.close
        set cnnSettings = Nothing

        fileContent = Replace(fileContent, "{{companyIdentityColor1}}", CompanyIdentityColor1)
        fileContent = Replace(fileContent, "{{companyIdentityColor2}}", CompanyIdentityColor2)
        fileContent = Replace(fileContent, "{{addressInfo}}", addressInfo)
        fileContent = Replace(fileContent, "{{phoneEmailInfo}}", phoneEmailInfo)
    End If

    If InStr(UCase(emailFileName),"INTERNAL")>0 OR InStr(UCase(emailFileName),"SWAP")>0 Then fileContent = Replace(fileContent, "{{clientLogo}}", "<img src='" & baseURl & "clientfiles/" & ClientID & "/logos/logo.png' style='margin-bottom:10px;'>")
       
    If InStr(UCase(emailFileName),"EXTERNAL")>0 Then fileContent = Replace(fileContent, "{{clientLogo}}", "<img src='" & baseURL & "clientfilesV/" & ClientID & "/logos/logo_email_header.png' width='650' style='margin-bottom:-5px; margin-left:3px margin-right:3px;'>")
End If

set fs=Nothing
%>	
<!-- content starts here !-->
<div class="row">
<form action="edit_email.asp?email=<%= emailFileName %>" method="post" id="frmSaveEmail">
<div style="width:700px;float:left">
    <input type="hidden" name="action" id="action" value="save" />  
    <div class="btnCancel"><button type="button" class="btn btn-primary btn-md btn-block" onclick="cancelEditEmail()">Cancel</button></div>
    <div class="btnSave"><button type="button" class="btn btn-primary btn-md btn-block" onclick="saveEmail()">Save Email</button></div>
    <div class="btnPreview"><button type="button" id="btnPreview" class="btn btn-primary btn-md btn-block" onclick="previewEmail()">Preview Email</button></div>    
    <div class="txtSendTestEmail"><input style="height:33px" type="text" placeholder="Type your email" class="input" name="testEmail" id="testEmail" /></div>
    <div class="btnSendTestEmail"><button type="button" class="btn btn-primary btn-md btn-block" onclick="sendTestEmail()">Send Test Email&nbsp;</button></div>
    <div class="btnBackToEmails"><button type="button" class="btn btn-primary btn-md btn-block" onclick="cancelEditEmail()">Back to Emails</button></div>
    <div class="btnLoadDefault"><button type="button" class="btn btn-primary btn-md btn-block" onclick="location.href='edit_email.asp?type=default&email=<%= emailFileNameDefault %>';" >Load Default</button></div>
    
    <TEXTAREA id="editEmailArea" COLS="92" ROWS="20" name="editEmailArea" class="form_grande"><%= fileContent %></TEXTAREA>
</div>
<div style="width:400px;float:left">
    <input type="text" id="searchText" onkeyup="searchFields()" placeholder="Search..">

    <ul id="fieldsList">

    </ul>
</div>
</form>
</div>
<!-- The Modal -->
<div id="previewModal" class="modal">

  <!-- Modal content -->
  <div class="modal-content">
    <div class="modal-header">
        <span class="closeModal">&times;</span>
        <h3>PLEASE NOTE THAT THIS PREVIEW IS NOT SAVED</h3>
    </div>
    <div id="modalContent"></div>
  </div>

</div>
<script type="text/javascript">
		//<![CDATA[
    bkLib.onDomLoaded(function () {
        displayEmailDynamicValues("fieldsList");
        new nicEditor({ fullPanel: true}).panelInstance('editEmailArea');
    });
		//]]>
</script> 
<!--#include file="../../inc/footer-main.asp"-->

<% Function replaceTestData(emailBody, emailFileName, ClientID)
    emailBody = Replace(emailBody,"{{accountNum}}","1234")
    emailBody = Replace(emailBody,"{{submissionSource}}","MDS Insight")
    If InStr(UCase(emailFileName),"INTERNAL")>0 OR InStr(UCase(emailFileName),"SWAP")>0 Then emailBody = Replace(emailBody, "{{clientLogo}}", "<img src='" & baseURl & "clientfiles/" & ClientID & "/logos/logo.png' style='margin-bottom:10px;'>")
       
    If InStr(UCase(emailFileName),"EXTERNAL")>0  Then emailBody = Replace(emailBody, "{{clientLogo}}", "<img src='" & baseURL & "clientfilesV/" & ClientID & "/logos/logo_email_header.png' width='650' style='margin-bottom:-5px; margin-left:3px margin-right:3px;'>")

    emailBody = Replace(emailBody,"{{companyName}}","ABC Company")
    emailBody = Replace(emailBody,"{{currentDateTime}}",Now())
    emailBody = Replace(emailBody,"{{custInfo}}","123 Main street")
    emailBody = Replace(emailBody,"{{nameByUserNo}}","John P.Tech")
    emailBody = Replace(emailBody,"{{nameForEmail}}","John P.Tech")
    emailBody = Replace(emailBody,"{{problemDescription}}","Replaced faulty relay")
    emailBody = Replace(emailBody,"{{serviceTicketNumber}}","1000")


    replaceTestData = emailBody
   End Function
%>