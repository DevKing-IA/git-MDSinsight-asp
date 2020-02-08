<%
'****************************************************
'Create the email that goes to the customer
'****************************************************

emailSubject = "MDS Insight New " & GetTerm("Prospecting") & " Prospect Owner Request Rejected"

emailBody = ""

emailBody =  emailBody & "<table width='650' border='0' cellspacing='5' cellpadding='5' style='font-family:Arial; border:1px solid #1a3049;' align='center'>"

emailBody =  emailBody & "<tr><th scope='col'><img src='" & BaseURL & "emails/img/header.png' ></th></tr>"

emailBody =  emailBody & "<tr><td><br>"

emailBody =  emailBody & "<table width='100%' border='0' cellspacing='5' cellpadding='5'><tr>"

emailBody =  emailBody & "<th scope='col'><img src='" & BaseURL & "emails/img/data.png' ></th>"

emailBody =  emailBody & "<th scope='col' valign='top' align='left' style='font-weight:normal;'>" & Date() & "<br><br>Greetings " & userOwnerName & "," & "<br><br>"

emailBody =  emailBody & GetUserDisplayNameByUserNo(userNo) & " has rejected ownership of the following prospect:<br><br>"

emailBody =	 emailBody & "<strong>Name</strong> : " & FirstName & " " & LastName & "<br>"
emailBody =	 emailBody & "<strong>Company</strong> : " & Company & "<br>"
emailBody =	 emailBody & "<strong>Street Address</strong> : " & Street & "<br>"
emailBody =	 emailBody & "<strong>Suite, Floor #, etc.</strong> : " & Address2 & "<br>"
emailBody =	 emailBody & "<strong>City, State Zip</strong> : " & City & ", " & State & " " & PostalCode & "<br><br>"

emailBody =	 emailBody & "<strong># Employees</strong> : " & NumEmployees & "<br>"
emailBody =	 emailBody & "<strong>Primary Competitor</strong> : " & PrimaryCompetitorName & "<br><br>"


If ProspectApptOrMeeting = "Appointment" Then

	emailBody =	 emailBody & "<strong>Next Activity</strong> : " & nextActivity & "<br>"
	emailBody =	 emailBody & "<strong>Activity Type</strong> : Appointment<br>"
	emailBody =	 emailBody & "<strong>Appointment Duration</strong> : " & Duration & "<br>"
	emailBody =	 emailBody & "<strong>Appointment Due Date</strong> : " & nextActivityDueDate & "<br><br>"

ElseIf ProspectApptOrMeeting = "Meeting" Then

	emailBody =	 emailBody & "<strong>Next Activity</strong> : " & nextActivity & "<br>"
	emailBody =	 emailBody & "<strong>Activity Type</strong> : Meeting<br>"
	emailBody =	 emailBody & "<strong>Meeting Location</strong> : " & Location & "<br>"
	emailBody =	 emailBody & "<strong>Meeting Duration</strong> : " & Duration & "<br>"
	emailBody =	 emailBody & "<strong>Meeting Date</strong> : " & nextActivityDueDate & "<br><br>"
	
Else

	emailBody =	 emailBody & "<strong>Next Activity</strong> : " & nextActivity & "<br>"
	emailBody =	 emailBody & "<strong>Next Activity Due Date</strong> : " & nextActivityDueDate & "<br><br>"

End If



emailBody =  emailBody & "</th>"

emailBody =  emailBody & "</tr></table>"

emailBody =  emailBody & "</td></tr>"

emailBody =  emailBody & "<tr><td><br>"

emailBody =  emailBody & "<table width='100%' border='0' cellspacing='5' cellpadding='5' align='center' style='background-color:#1a3049; text-align:center'><tr>"

emailBody =  emailBody & "<th scope='col' colspan='6' valign='middle' align='center' style='background-color:#eea236; font-weight:normal; font-size:22px; color :#fff; height:22px; text-decoration:none;'><a href='" & userQuickLoginURL & "?u=" & OwnerUserNo & "&c=" & userClientID & "&d=viewProspect-" & userProspectIntRecID & "'><font color='white'>Edit Prospect and Select New Owner</font></a></th>"

emailBody =  emailBody & "</tr></table><br>"


emailBody =  emailBody & "<table width='100%' border='0' cellspacing='5' cellpadding='5' style='background-color:#1a3049;'><tr>"

emailBody =  emailBody & "<th scope='col'><img src='" & BaseURL & "emails/img/footerlogo.png' ></th>"

emailBody =  emailBody & "<th scope='col' valign='middle' align='left' style='font-weight:normal; color:#fff;'> &copy; " & Year(Now()) & " Metroplex Data Systems. All Rights Reserved. </th>"

emailBody =  emailBody & "</tr></table>"


emailBody =  emailBody & "</table>"



%>