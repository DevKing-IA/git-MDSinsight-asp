<%
'****************************************************
'Create the email that goes to the customer
'****************************************************

emailSubject = "MDS Insight Quick Login Link"

emailBody = ""

emailBody =  emailBody & "<table width='650' border='0' cellspacing='0' cellpadding='5' style='font-family:Arial; border:1px solid #1a3049;' align='center'>"

emailBody =  emailBody & "<tr><td>"

emailBody =  emailBody & "<table width='100%' border='0' cellspacing='0' cellpadding='0'><tr>"

emailBody =  emailBody & "<th scope='col' width='450' valign='middle' align='left' bgcolor='#192f48'><img src='" & BaseURL & "emails/img/header-engineered-450.png' style='width: 100%; height: auto;' ></th>"

emailBody =  emailBody & "<th scope='col'  width='200' valign='middle' align='center' bgcolor='#edf3fa'><img src='" & BaseURL & "clientfiles/"& MUV_Read("ClientID") &"/logos/logo.png' style='width: 100%; height: auto;' ></th>"

emailBody =  emailBody & "</tr></table>"

emailBody =  emailBody & "<br><table width='100%' border='0' cellspacing='5' cellpadding='5'><tr>"

emailBody =  emailBody & "<th scope='col'><img src='" & BaseURL & "emails/img/data.png' ></th>"

emailBody =  emailBody & "<th scope='col' valign='top' align='left' style='font-weight:normal; font-family: Arial;'>" & Date()  & "  <br><br>Greetings " & rsUsers("userDisplayName") & "," & "<br><br>"

emailBody =  emailBody & "<span style='font-family: Arial;'>Here is your <b>MDS Insight Quick Login</b> link.<br><br> Use this URL to login on your phone: <br><br><b>" & userQuickLoginURL & "?u=" & UserNo & "&c=" & MUV_Read("ClientID") & "</span></b> "

emailBody =	 emailBody & "<br><br><span style='font-family: Arial;'><a href='" & userQuickLoginURL & "?u=" & UserNo & "&c=" & MUV_Read("ClientID") & "'></span>"

emailBody =  emailBody & "<img src='" & BaseURL & "emails/img/signin.png' border='0' style='border:0px;' ></a> </th>"

emailBody =  emailBody & "</tr></table>"

emailBody =  emailBody & "</td></tr>"

emailBody =  emailBody & "<tr><td><br>"

emailBody =  emailBody & "<table width='100%' border='0' cellspacing='5' cellpadding='5' style='background-color:#1a3049;'><tr>"

emailBody =  emailBody & "<th scope='col'><img src='" & BaseURL & "emails/img/footerlogo-updated.png' ></th>"

emailBody =  emailBody & "<th scope='col' valign='middle' align='left' style='font-weight:normal; color:#fff;'> <span style='font-family: Arial;'>&copy; " & Year(Now()) &" Metroplex Data Systems. All Rights Reserved. </span></th>"

emailBody =  emailBody & "</tr></table>"


emailBody =  emailBody & "</table>"



%>