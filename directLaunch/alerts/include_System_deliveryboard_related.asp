<%
Select Case rsAlert("Condition")

	Case "DBoardNotRun"
	
		Response.Write("Check Nightly delivery board did not run <br>")
		
		Set cnnTechInfo = Server.CreateObject("ADODB.Connection")
		cnnTechInfo.open (MUV_Read("ClientCnnString"))
		Set rsTechInfo = Server.CreateObject("ADODB.Recordset")
		rsTechInfo.CursorLocation = 3 
	
		SQL_TechInfo = "SELECT * FROM SC_TechInfo"
		Set rsTechInfo = cnnTechInfo.Execute(SQL_TechInfo)
		
		If rsTechInfo("NightlyDBoard_Start") <> "" AND rsTechInfo("NightlyDBoard_Finish") <> "" Then

			'Construct variable with Today & target time
			DateToCompare = Year(Now()) & "-" & month(Now()) & "-" & Day(Now()) & " "
			DateToCompare = DateToCompare & FormatDateTime(Left(rsAlert("TimeOfDay"), Len(rsAlert("TimeOfDay"))-2 ) & ":" &  Right(rsAlert("TimeOfDay"),2),3)

			DateToCompare = DateAdd ("d",0,DateToCompare) ' This just converts it to a usable date format (we add 0, no change is actually made)
			
			Response.Write ("DateToCompare:" & DateToCompare  & "<br>")

			If Now() > DateToCompare Then ' OK, we are passed the time it should have started
			
				Response.Write ("We have reached the target time, which is:" & DateToCompare  & " start checking<br>")
				
				If DateDiff("d",rsTechInfo("NightlyDBoard_Start"),Now()) > 0 Then ' Didn't run today at all so send the alert, 
				
					If AlertSent(rsAlert("InternalAlertRecNumber"),rsTechInfo("NightlyDBoard_Start")) <> True Then
				
						SendAlert rsAlert("InternalAlertRecNumber"),rsTechInfo("NightlyDBoard_Start"),rsAlert("Condition"),rsAlert("NotificationType")
					
						Response.Write ("SENDING ALERT<br>")
					End If
				
				End If
			
			Else
				Response.Write ("Nothng to do, we haven't reached the target time yet, which is:" & DateToCompare  & "<br>")			
			End If
			
		Else
			Response.Write("<font color='red'><strong>Either NightlyDBoard_Start or NightlyDBoard_Finish is empty in SC_TechInfo - Can't run check</strong></font><br>")
		End If
		
		Set rsTechInfo = Nothing
		cnnTechInfo.Close
		Set cnnTechInfo = Nothing

		
	Case "DBoardSkipped"
	
		Response.Write("Check Nightly delivery board update skipped the update<br>")

		Set cnnTechInfo = Server.CreateObject("ADODB.Connection")
		cnnTechInfo.open (MUV_Read("ClientCnnString"))
		Set rsTechInfo = Server.CreateObject("ADODB.Recordset")
		rsTechInfo.CursorLocation = 3 
	
		SQL_TechInfo = "SELECT * FROM SC_TechInfo"
		Set rsTechInfo = cnnTechInfo.Execute(SQL_TechInfo)
		
		If rsTechInfo("NightlyDBoard_LastAction") = "Skip" Then
		
			If AlertSent(rsAlert("InternalAlertRecNumber"),rsTechInfo("NightlyDBoard_Start")) <> True Then
			
				SendAlert rsAlert("InternalAlertRecNumber"),rsTechInfo("NightlyDBoard_Start"),rsAlert("Condition"),rsAlert("NotificationType")
				
			End If

		End If
		
		Set rsTechInfo = Nothing
		cnnTechInfo.Close
		Set cnnTechInfo = Nothing
	
	Case "DBoardFinished"
	
		Response.Write("Check Nightly delivery board update finished <br>")
	
		Set cnnTechInfo = Server.CreateObject("ADODB.Connection")
		cnnTechInfo.open (MUV_Read("ClientCnnString"))
		Set rsTechInfo = Server.CreateObject("ADODB.Recordset")
		rsTechInfo.CursorLocation = 3 
	
		SQL_TechInfo = "SELECT * FROM SC_TechInfo"
		Set rsTechInfo = cnnTechInfo.Execute(SQL_TechInfo)
		
		If rsTechInfo("NightlyDBoard_LastStatus") = "Finished" Then
		
			'Make sure the finished date is AFTER the last started date
			'That way, we know it's a new entry

			If rsTechInfo("NightlyDBoard_Start") < rsTechInfo("NightlyDBoard_Finish") Then
			
				If AlertSent(rsAlert("InternalAlertRecNumber"),rsTechInfo("NightlyDBoard_Finish")) <> True Then
				
					SendAlert rsAlert("InternalAlertRecNumber"),rsTechInfo("NightlyDBoard_Finish"),rsAlert("Condition"),rsAlert("NotificationType")
					
				End If
			End If
		
		End If
		
		Set rsTechInfo = Nothing
		cnnTechInfo.Close
		Set cnnTechInfo = Nothing

	Case "DBoardOnDemandRun"
	
		Response.Write("Check Delivery board update on demand was run <br>")
	
		Set cnnTechInfo = Server.CreateObject("ADODB.Connection")
		cnnTechInfo.open (MUV_Read("ClientCnnString"))
		Set rsTechInfo = Server.CreateObject("ADODB.Recordset")
		rsTechInfo.CursorLocation = 3 
	
		SQL_TechInfo = "SELECT * FROM SC_TechInfo"
		Set rsTechInfo = cnnTechInfo.Execute(SQL_TechInfo)
		
		If rsTechInfo("OnDemandDBoard_LastStatus") = "Finished" Then
		
			'Make sure the finished date is AFTER the last started date
			'That way, we know it's a new entry

			If rsTechInfo("OnDemandDBoard_Start") < rsTechInfo("OnDemandDBoard_Finish") Then
			
				If AlertSent(rsAlert("InternalAlertRecNumber"),rsTechInfo("OnDemandDBoard_Finish")) <> True Then
				
					SendAlert rsAlert("InternalAlertRecNumber"),rsTechInfo("OnDemandDBoard_Finish"),rsAlert("Condition"),rsAlert("NotificationType")
					
				End If
			End If
		
		End If
		
		Set rsTechInfo = Nothing
		cnnTechInfo.Close
		Set cnnTechInfo = Nothing

	
End Select
%>