<%
 
emailBody = ""

emailBody =  emailBody & "<table width='650' border='0' cellspacing='0' align='center' style='padding:10px; border:1px solid #000000;'>"

emailBody =  emailBody & "<tr><td>"

emailBody =  emailBody & "<table width='100%' border='0' cellspacing='0' cellpadding='15'  >"

emailBody =  emailBody & "<tr><td width='650' style='font-family:Arial, Helvetica, sans-serif; font-size:21px; font-weight:normal; padding-top:15px; padding-bottom:15px; margin-left:3px margin-right:3px;' align='center'><font color='red'>*ALERT*</font>"
emailBody =  emailBody & "<br>" & emailHeadLineText  & "<br>"

emailBody =  emailBody & "</tr></td>"

emailBody =  emailBody & "</table>"

emailBody =  emailBody & "</tr></td>"

emailBody =  emailBody & "<tr><td>"

emailBody =  emailBody & "<table width='100%' border='0' cellspacing='0' cellpadding='15'  >"

emailBody =  emailBody & "  <tr><th scope='col' width='40%' style='font-family:Arial, Helvetica, sans-serif; font-size:16px; font-weight:normal;' align='left'><strong>Invoice #</strong> </th> <th scope='col' width='60%' style='font-weight:normal; font-size:16px; font-family:Arial, Helvetica, sans-serif;' align='left'>"
emailBody =  emailBody & PassedInvoiceNumber & "</th></tr>"

emailBody =  emailBody & " <tr style='border-bottom:1px solid #666;' ><td width='40%' style='font-family:Arial, Helvetica, sans-serif; font-size:16px; font-weight:normal;'><strong>" & GetTerm("Customer")& "</strong></td><td width='60%' style='font-weight:normal; font-size:16px; font-family:Arial, Helvetica, sans-serif;' >"
emailBody =  emailBody & GetCustNumberByInvoiceNumDelBoard(passedInvoiceNumber) & "<br>"
emailBody =  emailBody & FormattedCustInfoByCustNum(GetCustNumberByInvoiceNumDelBoard(passedInvoiceNumber))
emailBody =  emailBody & "</td></tr>"

emailBody =  emailBody & " <tr ><td width='40%' style='font-family:Arial, Helvetica, sans-serif; font-size:16px; font-weight:normal;' valign='top'><strong>Driver</strong></td><td width='60%' style='font-weight:normal; font-size:16px; font-family:Arial, Helvetica, sans-serif;' >"
emailBody =  emailBody & GetDriverNameByTruckID(GetTruckByInvoiceNumDelBoard(passedInvoiceNumber)) & "</td></tr>"

emailBody =  emailBody & " <tr ><td width='40%' style='font-family:Arial, Helvetica, sans-serif; font-size:16px; font-weight:normal;' valign='top'><strong>Driver Comments</strong></td><td width='60%' style='font-weight:normal; font-size:16px; font-family:Arial, Helvetica, sans-serif;' >"
emailBody =  emailBody & GetDriverCommentsByInvoiceNumber(passedInvoiceNumber) & "</td></tr>"

emailBody =  emailBody & "</td></tr>"

emailBody =  emailBody & "</table>"

emailBody =  emailBody & "</td></tr>"

emailBody =  emailBody & "<tr><td>"

emailBody =  emailBody & "<small>(Alert Name: " & rsAlert("AlertName") & " -- Record ID: " & rsAlert("InternalAlertRecNumber") & " -- Client Key:" & ClientKey & ")</small>"

emailBody =  emailBody & "</td></tr>"
emailBody =  emailBody & "</table>"
 %>