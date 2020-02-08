<%
'****************************************************
'Create the email that goes to the customer
'****************************************************

emailSubject = "MDS Insight Password Reset - New Password"

emailBody = ""

emailBody =  emailBody & "<table width='650' border='0' cellspacing='5' cellpadding='5' style='font-family:Arial; border:1px solid #1a3049;' align='center'>"

emailBody =  emailBody & "<tr><th scope='col'><img src='" & BaseURL & "emails/img/header.png' ></th></tr>"

emailBody =  emailBody & "<tr><td><br>"

emailBody =  emailBody & "<table width='100%' border='0' cellspacing='5' cellpadding='5'><tr>"

emailBody =  emailBody & "<th scope='col'><img src='" & BaseURL & "emails/img/data.png' ></th>"

emailBody =  emailBody & "<th scope='col' valign='top' align='left' style='font-weight:normal;'>" & Date()  & "  <br><br>Greetings " & userDisplayName & "," & "<br><br>Your <b>MDS Insight</b> password has been reset. Use the information below to login to <b>MDS Insight</b> at: <a href='" & BaseURL & "'>" & BaseURL & "</a>.<br><br><b>Email:</b> " & userEmail  & "<br><b>Password:</b> " & Password & "<br><b>Client Key:</b> " & ClientKey & "<br><br> <a href='" & BaseURL & "'><img src='" & BaseURL & "emails/img/signin.png' border='0' style='border:0px;' ></a> </th>"

emailBody =  emailBody & "</tr></table>"

emailBody =  emailBody & "</td></tr>"

emailBody =  emailBody & " <tr><td><img src='" & BaseURL & "emails/img/emailbody.png' ></td></tr>"

emailBody =  emailBody & "<tr><td><br>"

emailBody =  emailBody & "<table width='100%' border='0' cellspacing='5' cellpadding='5' style='background-color:#1a3049;'><tr>"

emailBody =  emailBody & "<th scope='col'><img src='" & BaseURL & "emails/img/footerlogo.png' ></th>"

emailBody =  emailBody & "<th scope='col' valign='middle' align='left' style='font-weight:normal; color:#fff;'> &copy; " & Year(Now())   &" Metroplex Data Systems. All Rights Reserved. </th>"

emailBody =  emailBody & "</tr></table>"


emailBody =  emailBody & "</table>"



%>