<!--#include file="../../../inc/InsightFuncs.asp"-->
<!--#include file="../../../inc/mail.asp"-->
<!--#include file="../../../inc/InSightFuncs_routing.asp"-->
<!--#include file="../../../inc/SubsAndFuncs.asp"-->
<%
If right(baseURL,1)="/" Then maildomain = Left(right(baseURL,len(baseURL)-7),len(right(baseURL,len(baseURL)-7))-1) Else maildomain = right(baseURL,len(baseURL)-7)

SQLOutter = "Select IvsNum FROM RT_DeliveryBoard Where CustNum = '" & CustNum & "'"

Set cnn_Alert = Server.CreateObject("ADODB.Connection")
cnn_Alert.open (Session("ClientCnnString"))
Set rsOutter = Server.CreateObject("ADODB.Recordset")
rsOutter.CursorLocation = 3 
Set rsOutter = cnn_Alert.Execute(SQLOutter)

If not rsOutter.Eof Then

	Do until rsOutter.Eof
			
		InvoiceNumberToCheck = rsOutter("IvsNum")
		
		SQLAlerts = "SELECT * FROM SC_Alerts Where AlertType='DeliveryBoard' and ReferenceValue ='" & InvoiceNumberToCheck & "' And Enabled = 1" 
		
		Set rsAlert = Server.CreateObject("ADODB.Recordset")
		rsAlert.CursorLocation = 3 
		
		Set rsAlert = cnn_Alert.Execute(SQLAlerts)
		
		If Not rsAlert.Eof then
		
			Do While Not rsAlert.EOF
			
				'Next stop alerts first
				If rsAlert("Condition") = "This becomes the next stop" Then 
					'See if it is the next stop
					If GetCustNumberByInvoiceNumDelBoard(InvoiceNumberToCheck) = GetNextCustomerStopByTruck(GetTruckByInvoiceNumDelBoard(InvoiceNumberToCheck)) Then
						'Yes, they are the next stop
						emailBody = "The driver has indicated that customer " & GetCustNumberByInvoiceNumDelBoard(InvoiceNumberToCheck) & " - " & GetCustNameByCustNum(GetCustNumberByInvoiceNumDelBoard(InvoiceNumberToCheck)) & "  is the next stop."
						txtMessage = "Customer " & GetCustNumberByInvoiceNumDelBoard(InvoiceNumberToCheck) & " - " & GetCustNameByCustNum(GetCustNumberByInvoiceNumDelBoard(InvoiceNumberToCheck)) & "  is the next stop."
					End If
				End IF
				
				'Next check delievered or skipped
				If rsAlert("Condition") = "Stop is completed or skipped" Then 
					'See is delievered or not
					If GetDeliveryStatusByInvoice(InvoiceNumberToCheck) = "Delivered" or GetDeliveryStatusByInvoice(InvoiceNumberToCheck) = "No Delivery" Then
						'Updated delivery status
						emailBody = "The delivery status of invoice #" & InvoiceNumberToCheck & " for customer " & GetCustNumberByInvoiceNumDelBoard(InvoiceNumberToCheck) & " - " & GetCustNameByCustNum(GetCustNumberByInvoiceNumDelBoard(InvoiceNumberToCheck)) 
						emailBody = emailBody & " was changed to <strong>" & GetDeliveryStatusByInvoice(InvoiceNumberToCheck) & "</strong> at " & FormatDateTime(GetLastDeliveryStatusChangeBYInvoiceNumDelBoard(InvoiceNumberToCheck))
						txtMessage = "Inv#" & InvoiceNumberToCheck & " Cust#" & GetCustNumberByInvoiceNumDelBoard(InvoiceNumberToCheck) & " - " & GetDeliveryStatusByInvoice(InvoiceNumberToCheck)
					End If
				End IF
			
				Send_To = ""
				'*************************
				'First do the email alerts
				'*************************
				'Get user based emails
				If Not IsNull(rsAlert("EmailToUserNos")) Then
					If rsAlert("EmailToUserNos") <> "" And rsAlert("EmailToUserNos") <> "0" Then
						UserNoList = Split(rsAlert("EmailToUserNos"),",")
						For x = 0 To UBound(UserNoList)
							Send_To = Send_To & GetUserEmailByUserNo(UserNoList(x)) & ";"
						Next
					End If
				End If
													
				'Get additional emails if there are any
				If rsAlert("AdditionalEmails") <> "" and not IsNull(rsAlert("AdditionalEmails")) Then
					tmpSendAlertToAdditionalEmails = trim(rsAlert("AdditionalEmails"))		
					If Len(tmpSendAlertToAdditionalEmails) > 1 Then
						If Right(tmpSendAlertToAdditionalEmails,1) <> ";" Then tmpSendAlertToAdditionalEmails = tmpSendAlertToAdditionalEmails & ";"
						Send_To = Send_To & tmpSendAlertToAdditionalEmails
					End If
				End If
				
			
				'Only do this if there are actually emails to send
				If Send_To <> "" Then
					'Now Send the emails
					'Got all the addresses so now break them up
					Send_To_Array = Split(Send_To,";")
			
					For x = 0 to Ubound(Send_To_Array) -1
						Send_To = Send_To_Array(x)
						'Failsafe for dev
						If Instr(ucase(Request.ServerVariables("SERVER_NAME")),"DEV") <> 0 Then Send_To = "rich@ocsaccess.com"
						SendMail "mailsender@" & maildomain,Send_To,"Delivery Alert ",emailBody,GetTerm("Routing"),"Delivery Alert"
					Next 
				End If
			
			
				Send_To=""
				'**********************
				'Now do the text alerts
				'**********************
				'Get user based Texts
				If Not IsNull(rsAlert("TextToUserNos")) Then
					If rsAlert("TextToUserNos") <> "" And rsAlert("TextToUserNos") <> "0" Then
						UserNoList = Split(rsAlert("TextToUserNos"),",")
						For x = 0 To UBound(UserNoList)
							Send_To = Send_To & getUserCellNumber(UserNoList(x)) & ","
						Next
					End If
				End If
			
				'Get additional texts if there are any
				If rsAlert("AdditionalText") <> "" and not IsNull(rsAlert("AdditionalText")) Then
					tmpSendAlertToAdditionalTexts = trim(rsAlert("AdditionalText"))		
					If Len(tmpSendAlertToAdditionalTexts) > 1 Then
						If Right(tmpSendAlertToAdditionalTexts,1) <> "," Then tmpSendAlertToAdditionalTexts = tmpSendAlertToAdditionalTexts & ","
						 Send_To = Send_To & tmpSendAlertToAdditionalTexts
					End If
				End If
			
				'Only do this if there are actually texts to send
				If Send_To <> "" Then
											
					Send_To = Replace(Send_To,"-","") ' EZ Texting doesn't like dashes
			
					txtSubject = "DeliveryAlert"
					
					'*****Text numbers don't get split into an array, the php takes multiple #'s seprated by commas	
						
					If Right(Send_To,1) = "," Then Send_To = Left(Send_To,Len(Send_To)-1)
				
					TEXT_TO = Send_To
					
					'Split Text_To for recording the alerts sent
					TextNumberArray = Split(TEXT_TO&",",",")
					
		
					str_data="txtSubject=" & txtSubject  & "&txtMessage=" & txtMessage & "&txtTEXT_TO=" & TEXT_TO
			
					str_data = str_data & "&txtu1=" & EzTextingUserID() & "&txtu2=" & EzTextingPassword()
					
					str_data = str_data & "&txtCountry=" & GetCompanyCountry()
			
					Set obj_post=Server.CreateObject("Msxml2.SERVERXMLHTTP")
					obj_post.Open "POST", BaseURL & "inc/sendtext_post.php",False
					obj_post.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
					obj_post.Send str_data
					
				End If
			
			
				rsAlert.MoveNext
			Loop			
		
		End If		
		
		Set rsAlert = Nothing
	
		rsOutter.Movenext
	Loop
End If

Set rsOutter = Nothing
cnn_Alert.close
Set cnn_Alert = Nothing

%>