<!--#include file="../../inc/settings.asp"-->
<!--#include file="../../inc/InSightFuncs.asp"-->
<!--#include file="../../inc/InSightFuncs_Users.asp"-->
<%
Dim emailFileName: emailFileName = Request.QueryString("email")
Dim emailtype: emailtype = Request.QueryString("type")
dim fs,f, fileContent
set fs=Server.CreateObject("Scripting.FileSystemObject")
serno = MUV_READ("SERNO") 
If InStr(emailFileName,"Default")>0 Then
    path = "C:\home\clientfilesV\default_emails\"
Else
    path = "C:\home\clientfilesV\" & serno & "\emails\"
End If

If fs.FileExists(path & emailFileName) Then
   set f=fs.OpenTextFile(path & emailFileName,1)
   fileContent = f.ReadAll()
   f.Close
   set f=Nothing
Else 
    If InStr(emailFileName,"Default")=0 Then
        fileContent = "<b>Custom email template doesn't exist</b>"
    End If
End If

set fs=Nothing

        If emailtype = "default" Then
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

       ClientID = MUV_READ("ClientID")

       If InStr(UCase(emailFileName),"INTERNAL")>0 OR InStr(UCase(emailFileName),"SWAP")>0 Then fileContent = Replace(fileContent, "{{clientLogo}}", "<img src='" & baseURL & "clientfiles/" & ClientID & "/logos/logo.png' style='margin-bottom:10px;'>")
       
       If InStr(UCase(emailFileName),"EXTERNAL")>0 Then fileContent = Replace(fileContent, "{{clientLogo}}", "<img src='" & baseURl & "clientfilesV/" & ClientID & "/logos/logo_email_header.png' width='650' style='margin-bottom:-5px; margin-left:3px margin-right:3px;'>")

Response.Write(fileContent)
%>	

