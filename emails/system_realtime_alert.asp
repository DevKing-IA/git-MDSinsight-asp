<%
emailBody = ""

emailBody =  emailBody & "<table width='650' border='0' cellspacing='0' align='center' style='padding:10px; border:1px solid #000000;'>"

emailBody =  emailBody & "<tr><td>"



emailBody =  emailBody & "<table width='100%' border='0' cellspacing='0' cellpadding='15'  >"

emailBody =  emailBody & "<tr><td width='650' style='font-family:Arial, Helvetica, sans-serif; font-size:21px; font-weight:normal; padding-top:15px; padding-bottom:15px; margin-left:3px margin-right:3px;' align='center'>"

If passedNotificationType = "Alert" Then
	emailBody =  emailBody & "<font color='red'>*ALERT*</font><br>"
Else
	emailBody =  emailBody & "*Notification*<br>"
End If
	
emailBody =  emailBody & "<br>" & emailHeadLineText  & "<br>"

emailBody =  emailBody & "</tr></td>"

emailBody =  emailBody & "</table>"




emailBody =  emailBody & "</tr></td>"

emailBody =  emailBody & "<tr><td>"

emailBody =  emailBody & "</td></tr>"

emailBody =  emailBody & "<tr><td>"

emailBody =  emailBody & "(" & rsAlert("AlertName") & ":" & rsAlert("InternalAlertRecNumber") & ":" & ClientKey & ")"

emailBody =  emailBody & "</td></tr>"
emailBody =  emailBody & "</table>"
 %>