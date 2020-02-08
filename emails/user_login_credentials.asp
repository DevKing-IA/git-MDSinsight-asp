<%
'****************************************************
'Create the email that goes to the customer
'****************************************************

emailSubject = "MDS Insight Login Credentials"

emailBody = ""

emailBody =  emailBody & "<table width='650' border='0' cellspacing='0' cellpadding='5' style='font-family:Arial; border:1px solid #1a3049;' align='center'>"

emailBody =  emailBody & "<tr><td>"

emailBody =  emailBody & "<table width='100%' border='0' cellspacing='0' cellpadding='0'><tr>"

emailBody =  emailBody & "<th scope='col' width='450' valign='middle' align='left' bgcolor='#192f48'><img src='" & BaseURL & "emails/img/header-engineered-450.png' style='width: 100%; height: auto;' ></th>"

emailBody =  emailBody & "<th scope='col'  width='200' valign='middle' align='center' bgcolor='#edf3fa'><img src='" & BaseURL & "clientfiles/"& MUV_Read("ClientID") &"/logos/logo.png' style='width: 100%; height: auto;' ></th>"

emailBody =  emailBody & "</tr></table>"

emailBody =  emailBody & "<br><table width='100%' border='0' cellspacing='5' cellpadding='5'><tr>"

emailBody =  emailBody & "<th scope='col'><img src='" & BaseURL & "emails/img/data.png' ></th>"

emailBody =  emailBody & "<th scope='col' valign='top' align='left' style='font-weight:normal;'><span style='font-family: Arial;'>" & Date()  & "  <br><br>Greetings " & rs("userDisplayName") & "," & "<br><br>Your <b>MDS Insight</b> account has been created. Use the information below to login to <b>MDS Insight</b> at: <a href='" & BaseURL & "'>" & BaseURL & "</a>.<br><br><b>Email:</b> " & rs("userEmail")  & "<br><b>Password:</b> " & rs("userPassword") & "<br><b>Client Key:</b> " &  MUV_Read("ClientID") & "<br><br> <a href='" & BaseURL & "'><img src='" & BaseURL & "emails/img/signin.png' border='0' style='border:0px;' ></a> </span></th>"

emailBody =  emailBody & "</tr></table>"

emailBody =  emailBody & "</td></tr>"

'emailBody =  emailBody & " <tr><td><img src='" & BaseURL & "emails/img/emailbody.png' ></td></tr>"

emailBody =  emailBody & "<tr><td><br>"

emailBody =  emailBody & "<table width='100%' border='0' cellspacing='5' cellpadding='5' style='background-color:#1a3049;'><tr>"

emailBody =  emailBody & "<th scope='col'><img src='" & BaseURL & "emails/img/footerlogo-updated.png' ></th>"

emailBody =  emailBody & "<th scope='col' valign='middle' align='left' style='font-weight:normal; color:#fff;'><span style='font-family: Arial;'> &copy; " & Year(Now())   &" Metroplex Data Systems. All Rights Reserved. </span></th>"

emailBody =  emailBody & "</tr></table>"


emailBody =  emailBody & "</table>"



%>