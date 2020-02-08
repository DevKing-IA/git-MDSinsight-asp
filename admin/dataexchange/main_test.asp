<HTML>
<HEAD>
<META http-equiv="Content-Type" content="text/html; charset=utf-8">

<%
Set UploadProgress = Server.CreateObject("Persits.UploadProgress")
PID = UploadProgress.CreateProgressID()

Session.CodePage = 65001

Set Upload = Server.CreateObject("Persits.Upload")
Upload.CodePage = 65001

' This is needed to enable the progress indicator
Upload.ProgressID = Request.QueryString("PID")
Upload.IgnoreNoPost = True ' to use upload script in the same file as the form.

Upload.SetMaxSize 52428, false

Upload.OverwriteFiles = false

Upload.Save "c:\upload"

If Upload.Files.Count > 0 Then
	Res = "Success! " & Upload.Files.Count & " files have been uploaded."
Else
	Res = ""
End If

%>

<script src="progress_ajax.js"></script> 

</HEAD>
<BODY>
<BASEFONT FACE="Arial" SIZE="2">

	<h3>Ajax-based Progress Bar (requires AspUpload 3.1)</h3>
	
	<P>
	We use HTML5's &lt;INPUT TYPE=FILE NAME="FILE1" <B>multiple="multiple"</B>>
	
	<P>

	<FORM METHOD="POST" ENCTYPE="multipart/form-data"
			ACTION="main_test.asp?pid=<% = PID %>"
			OnSubmit="ShowProgress('<% = PID %>')"> 

	<INPUT TYPE="FILE" NAME="FILE1" SIZE="40" multiple=multiple><BR>
	<!--<INPUT TYPE="FILE" NAME="FILE2" SIZE="40"><BR>
	<INPUT TYPE="FILE" NAME="FILE3" SIZE="40"><P>-->

	<INPUT TYPE="SUBMIT" VALUE="Upload!">
	<INPUT TYPE="BUTTON" VALUE="Stop" OnClick="OnStop()">

	</FORM>

	<div id="txtProgress"></div>

	<% = Res %>


<P>

<A HREF="progress_ajax.zip">Download source files for this live demo</A>

</basefont>
</BODY>
</HTML>
