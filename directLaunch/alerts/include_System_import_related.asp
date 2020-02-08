<%
Select Case rsAlert("Condition")

	Case "BackendNoStart"
	
		Response.Write("Check Backend data import did not start <br>")
		
		Set cnnTechInfo = Server.CreateObject("ADODB.Connection")
		cnnTechInfo.open (MUV_Read("ClientCnnString"))
		Set rsTechInfo = Server.CreateObject("ADODB.Recordset")
		rsTechInfo.CursorLocation = 3 
	
		SQL_TechInfo = "SELECT * FROM SC_TechInfo"
		Set rsTechInfo = cnnTechInfo.Execute(SQL_TechInfo)
		
		If rsTechInfo("BackendImport_Start") <> "" AND rsTechInfo("BackendImport_Finish") <> "" Then

			'Construct variable with Today & target time
			DateToCompare = Year(Now()) & "-" & month(Now()) & "-" & Day(Now()) & " "
			DateToCompare = DateToCompare & FormatDateTime(Left(rsAlert("TimeOfDay"), Len(rsAlert("TimeOfDay"))-2 ) & ":" &  Right(rsAlert("TimeOfDay"),2),3)

			DateToCompare = DateAdd ("d",0,DateToCompare) ' This just converts it to a usable date format (we add 0, no change is actually made)
			
			Response.Write ("DateToCompare:" & DateToCompare  & "<br>")

			If Now() > DateToCompare Then ' OK, we are passed the time it should have started
			
				Response.Write ("We have reached the target time, which is:" & DateToCompare  & " start checking<br>")
				
				If DateDiff("d",rsTechInfo("BackendImport_Start"),Now()) > 0 Then ' Didn't run today at all so send the alert, 
				
					If AlertSent(rsAlert("InternalAlertRecNumber"),rsTechInfo("BackendImport_Start")) <> True Then
				
						SendAlert rsAlert("InternalAlertRecNumber"),rsTechInfo("BackendImport_Start"),rsAlert("Condition"),rsAlert("NotificationType")
					
					End If
				
				End If
			
			Else
				Response.Write ("Nothng to do, we haven't reached the target time yet, which is:" & DateToCompare  & "<br>")			
			End If
			
		Else
			Response.Write("<font color='red'><strong>Either BackendImport_Start or BackendImport_Finish is empty in SC_TechInfo - Can't run check</strong></font><br>")
		End If
		
		Set rsTechInfo = Nothing
		cnnTechInfo.Close
		Set cnnTechInfo = Nothing

		
	Case "BackendRunTooLong"
	
		Response.Write("Check Backend data import has been running too long<br>")

		Set cnnTechInfo = Server.CreateObject("ADODB.Connection")
		cnnTechInfo.open (MUV_Read("ClientCnnString"))
		Set rsTechInfo = Server.CreateObject("ADODB.Recordset")
		rsTechInfo.CursorLocation = 3 
	
		SQL_TechInfo = "SELECT * FROM SC_TechInfo"
		Set rsTechInfo = cnnTechInfo.Execute(SQL_TechInfo)
		
		If rsTechInfo("BackendImport_LastStatus") = "Started" Then
		
			'Make sure the started date is AFTER the last finished date
			'That way, we know it's a new entry

			If rsTechInfo("BackendImport_Start") > rsTechInfo("BackendImport_Finish") Then
			
				'Figure out how many munites it has been running
				RunningMinutes = DateDiff("n",rsTechInfo("BackendImport_Start"),Now())
				Response.Write("Running minutes is:" & RunningMinutes & "<br>")
				dummy = MUV_WRITE("RunningMinutes",RunningMinutes) ' For this alert, we need this for the email & text
				
				If cint(RunningMinutes) > cint(rsAlert("NBMinutes")) Then

					Response.Write("Running too long<br>")
		
					If AlertSent(rsAlert("InternalAlertRecNumber"),rsTechInfo("BackendImport_Start")) <> True Then
					
						SendAlert rsAlert("InternalAlertRecNumber"),rsTechInfo("BackendImport_Start"),rsAlert("Condition"),rsAlert("NotificationType")
						
					End If
					
				Else
				
					Response.Write("NOT running too long. Runnig for " & RunningMinutes & " minutes. Alert is set for after " & rsAlert("NBMinutes") & " minutes.<br>")
				
				End If
			End If
		
		End If
		
		Set rsTechInfo = Nothing
		cnnTechInfo.Close
		Set cnnTechInfo = Nothing
	
	Case "BackendStarted"
	
		Response.Write("Check Backend data import started <br>")

		Set cnnTechInfo = Server.CreateObject("ADODB.Connection")
		cnnTechInfo.open (MUV_Read("ClientCnnString"))
		Set rsTechInfo = Server.CreateObject("ADODB.Recordset")
		rsTechInfo.CursorLocation = 3 
	
		SQL_TechInfo = "SELECT * FROM SC_TechInfo"
		Set rsTechInfo = cnnTechInfo.Execute(SQL_TechInfo)
		
		If rsTechInfo("BackendImport_LastStatus") = "Started" Then
		
			'Make sure the started date is AFTER the last finished date
			'That way, we know it's a new entry

			If rsTechInfo("BackendImport_Start") > rsTechInfo("BackendImport_Finish") Then
			
				If AlertSent(rsAlert("InternalAlertRecNumber"),rsTechInfo("BackendImport_Start")) <> True Then
				
					SendAlert rsAlert("InternalAlertRecNumber"),rsTechInfo("BackendImport_Start"),rsAlert("Condition"),rsAlert("NotificationType")
					
				End If
			End If
		
		End If
		
		Set rsTechInfo = Nothing
		cnnTechInfo.Close
		Set cnnTechInfo = Nothing
									
	Case "BackendFinished"
	
		Response.Write("Check Backend data import finished <br>")
	
		Set cnnTechInfo = Server.CreateObject("ADODB.Connection")
		cnnTechInfo.open (MUV_Read("ClientCnnString"))
		Set rsTechInfo = Server.CreateObject("ADODB.Recordset")
		rsTechInfo.CursorLocation = 3 
	
		SQL_TechInfo = "SELECT * FROM SC_TechInfo"
		Set rsTechInfo = cnnTechInfo.Execute(SQL_TechInfo)
		
		If rsTechInfo("BackendImport_LastStatus") = "Finished" Then
		
			'Make sure the finished date is AFTER the last started date
			'That way, we know it's a new entry

			If rsTechInfo("BackendImport_Start") < rsTechInfo("BackendImport_Finish") Then
			
				If AlertSent(rsAlert("InternalAlertRecNumber"),rsTechInfo("BackendImport_Finish")) <> True Then
				
					SendAlert rsAlert("InternalAlertRecNumber"),rsTechInfo("BackendImport_Finish"),rsAlert("Condition"),rsAlert("NotificationType")
				Else
				
					Response.Write("This alert previously sent<br>")	
				End If
			End If
		
		End If
		
		Set rsTechInfo = Nothing
		cnnTechInfo.Close
		Set cnnTechInfo = Nothing

	Case "RebuildNotRun"
	
		Response.Write("Check Daily data rebuild did not start<br>")
		
		Set cnnTechInfo = Server.CreateObject("ADODB.Connection")
		cnnTechInfo.open (MUV_Read("ClientCnnString"))
		Set rsTechInfo = Server.CreateObject("ADODB.Recordset")
		rsTechInfo.CursorLocation = 3 
	
		SQL_TechInfo = "SELECT * FROM SC_TechInfo"
		Set rsTechInfo = cnnTechInfo.Execute(SQL_TechInfo)
		
		If rsTechInfo("DailyRebuild_Start") <> "" AND rsTechInfo("DailyRebuild_Finish") <> "" Then

			'Construct variable with Today & target time
			DateToCompare = Year(Now()) & "-" & month(Now()) & "-" & Day(Now()) & " "
			DateToCompare = DateToCompare & FormatDateTime(Left(rsAlert("TimeOfDay"), Len(rsAlert("TimeOfDay"))-2 ) & ":" &  Right(rsAlert("TimeOfDay"),2),3)

			DateToCompare = DateAdd ("d",0,DateToCompare) ' This just converts it to a usable date format (we add 0, no change is actually made)
			
			Response.Write ("DateToCompare:" & DateToCompare  & "<br>")

			If Now() > DateToCompare Then ' OK, we are passed the time it should have started
			
				Response.Write ("We have reached the target time, which is:" & DateToCompare  & " start checking<br>")
				
				If DateDiff("d",rsTechInfo("DailyRebuild_Start"),Now()) > 0 Then ' Didn't run today at all so send the alert, 
				
					If AlertSent(rsAlert("InternalAlertRecNumber"),rsTechInfo("DailyRebuild_Start")) <> True Then
				
						SendAlert rsAlert("InternalAlertRecNumber"),rsTechInfo("DailyRebuild_Start"),rsAlert("Condition"),rsAlert("NotificationType")
					
					End If
				
				End If
			
			Else
				Response.Write ("Nothng to do, we haven't reached the target time yet, which is:" & DateToCompare  & "<br>")			
			End If
			
		Else
			Response.Write("<font color='red'><strong>Either DailyRebuild_Start or DailyRebuild_Finish is empty in SC_TechInfo - Can't run check</strong></font><br>")
		End If
		
		Set rsTechInfo = Nothing
		cnnTechInfo.Close
		Set cnnTechInfo = Nothing
	
	Case "RebuildRunTooLong"
	
		Response.Write("Check Daily data rebuild has been running longer than<br>")
	

		Set cnnTechInfo = Server.CreateObject("ADODB.Connection")
		cnnTechInfo.open (MUV_Read("ClientCnnString"))
		Set rsTechInfo = Server.CreateObject("ADODB.Recordset")
		rsTechInfo.CursorLocation = 3 
	
		SQL_TechInfo = "SELECT * FROM SC_TechInfo"
		Set rsTechInfo = cnnTechInfo.Execute(SQL_TechInfo)
		
		If rsTechInfo("DailyRebuild_LastStatus") = "Started" Then
		
			'Make sure the started date is AFTER the last finished date
			'That way, we know it's a new entry

			If rsTechInfo("DailyRebuild_Start") > rsTechInfo("DailyRebuild_Finish") Then
			
				'Figure out how many munites it has been running
				RunningMinutes = DateDiff("n",rsTechInfo("DailyRebuild_Start"),Now())
				Response.Write("Running minutes is:" & RunningMinutes & "<br>")
				dummy = MUV_WRITE("RunningMinutes",RunningMinutes) ' For this alert, we need this for the email & text
				
				If cint(RunningMinutes) > cint(rsAlert("NBMinutes")) Then

					Response.Write("Running too long<br>")
		
					If AlertSent(rsAlert("InternalAlertRecNumber"),rsTechInfo("DailyRebuild_Start")) <> True Then
					
						SendAlert rsAlert("InternalAlertRecNumber"),rsTechInfo("DailyRebuild_Start"),rsAlert("Condition"),rsAlert("NotificationType")
						
					End If
					
				Else
				
					Response.Write("NOT running too long. Runnig for " & RunningMinutes & " minutes. Alert is set for after " & rsAlert("NBMinutes") & " minutes.<br>")
				
				End If
			End If
		
		End If
		
		Set rsTechInfo = Nothing
		cnnTechInfo.Close
		Set cnnTechInfo = Nothing
	
	Case "RebuildStarted"
	
		Response.Write("Check Daily data rebuild started<br>")

		Set cnnTechInfo = Server.CreateObject("ADODB.Connection")
		cnnTechInfo.open (MUV_Read("ClientCnnString"))
		Set rsTechInfo = Server.CreateObject("ADODB.Recordset")
		rsTechInfo.CursorLocation = 3 
	
		SQL_TechInfo = "SELECT * FROM SC_TechInfo"
		Set rsTechInfo = cnnTechInfo.Execute(SQL_TechInfo)
		
		If rsTechInfo("DailyRebuild_LastStatus") = "Started" Then
		
			'Make sure the started date is AFTER the last finished date
			'That way, we know it's a new entry

			If rsTechInfo("DailyRebuild_Start") > rsTechInfo("DailyRebuild_Finish") Then
			
				If AlertSent(rsAlert("InternalAlertRecNumber"),rsTechInfo("DailyRebuild_Start")) <> True Then
				
					SendAlert rsAlert("InternalAlertRecNumber"),rsTechInfo("DailyRebuild_Start"),rsAlert("Condition"),rsAlert("NotificationType")
					
				End If
			End If
		
		End If
		
		Set rsTechInfo = Nothing
		cnnTechInfo.Close
		Set cnnTechInfo = Nothing
						
	Case "RebuildFinished"
	
		Response.Write("Check Daily data rebuild finished<br>")
	
		Set cnnTechInfo = Server.CreateObject("ADODB.Connection")
		cnnTechInfo.open (MUV_Read("ClientCnnString"))
		Set rsTechInfo = Server.CreateObject("ADODB.Recordset")
		rsTechInfo.CursorLocation = 3 
	
		SQL_TechInfo = "SELECT * FROM SC_TechInfo"
		Set rsTechInfo = cnnTechInfo.Execute(SQL_TechInfo)
		
		If rsTechInfo("DailyRebuild_LastStatus") = "Finished" Then
		
			'Make sure the finished date is AFTER the last started date
			'That way, we know it's a new entry

			If rsTechInfo("DailyRebuild_Start") < rsTechInfo("DailyRebuild_Finish") Then
			
				If AlertSent(rsAlert("InternalAlertRecNumber"),rsTechInfo("DailyRebuild_Finish")) <> True Then
				
					SendAlert rsAlert("InternalAlertRecNumber"),rsTechInfo("DailyRebuild_Finish"),rsAlert("Condition"),rsAlert("NotificationType")
					
				End If
			End If
		
		End If
		
		Set rsTechInfo = Nothing
		cnnTechInfo.Close
		Set cnnTechInfo = Nothing
		
	Case "AutoCompJSONNotRun"
	
		Response.Write("Auto-complete JSON file not rebuilt <br>")
		
		'To begin with, just see if the files are there.
		'If they are missing, no need to check further
		'Just send the alert
		
		Set fs = CreateObject("Scripting.FileSystemObject")
		Pth =  "../../clientfiles/" & ClientKey & "/autocomplete/customer_account_list_" & ClientKey  & ".json"
		customer_account_list = fs.FileExists(Server.MapPath(Pth)) 
		If customer_account_list Then
			Set f1 = fs.GetFile(Server.MapPath(Pth))
			customer_account_list_modified = f1.DateLastModified
		End If
		Set f1 = Nothing
		Set fs = Nothing
	
		Set fs = CreateObject("Scripting.FileSystemObject")
		Pth =  "../../clientfiles/" & ClientKey & "/autocomplete/customer_account_list_CSZ_" & ClientKey  & ".json"
		customer_account_list_CSZ = fs.FileExists(Server.MapPath(Pth)) 
		If customer_account_list_CSZ Then
			Set f2 = fs.GetFile(Server.MapPath(Pth))
			customer_account_list_CSZ_modified  = f2.DateLastModified
		End If
		Set f2 = Nothing
		Set fs = Nothing
				
		If customer_account_list <> True OR customer_account_list_CSZ <> True Then ' At least one doesn't exist

			If AlertSent(rsAlert("InternalAlertRecNumber"),FormatDateTime(Now(),2)) <> True Then
			
				SendAlert rsAlert("InternalAlertRecNumber"),FormatDateTime(Now(),2),"AutoCompJSONNotRun",rsAlert("NotificationType")
					
			End If

		Else
		
			'Construct variable with Today & target time
			DateToCompare = Year(Now()) & "-" & month(Now()) & "-" & Day(Now()) & " "
			DateToCompare = DateToCompare & FormatDateTime(Left(rsAlert("TimeOfDay"), Len(rsAlert("TimeOfDay"))-2 ) & ":" &  Right(rsAlert("TimeOfDay"),2),3)

			DateToCompare = DateAdd ("d",0,DateToCompare) ' This just converts it to a usable date format (we add 0, no change is actually made)
			
			Response.Write ("DateToCompare:" & DateToCompare  & "<br>")

			If Now() > DateToCompare Then ' OK, we are passed the time it should have started
		
				Response.Write ("We have reached the target time, which is:" & DateToCompare  & " start checking<br>")
						
				'See if the file(s) are from today
				If DateDiff("d",customer_account_list_modified,Now()) <> 0 OR DateDiff("d",customer_account_list_CSZ_modified,Now()) <> 0  Then 

					Response.Write ("DateDiff customer_account_list_modified:" & DateDiff("d",customer_account_list_modified,Now()) & "<br>")
					Response.Write ("DateDiff customer_account_list_CSZ_modified:" & DateDiff("d",customer_account_list_CSZ_modified,Now()) & "<br>")
		
					If AlertSent(rsAlert("InternalAlertRecNumber"),FormatDateTime(Now(),2)) <> True Then
					
						SendAlert rsAlert("InternalAlertRecNumber"),FormatDateTime(Now(),2),"AutoCompJSONNotRun",rsAlert("NotificationType")
							
					End If
					
				End If

			Else
				Response.Write ("Nothng to do, we haven't reached the target time yet, which is:" & DateToCompare  & "<br>")			
			End If
		
		End If ' File Exists Endif
		
	Case "HistOldInvoice"
	
		Response.Write("Check most recent history invoice too old<br>")
	
		'Only check this if todays rebuild has completed
		'Otherwise data might be in a transitional state
		
		Response.Write("See if today's rebuild has finished<br>")
		RebuildDone = False
		
		Set cnnTechInfo = Server.CreateObject("ADODB.Connection")
		cnnTechInfo.open (MUV_Read("ClientCnnString"))
		Set rsTechInfo = Server.CreateObject("ADODB.Recordset")
		rsTechInfo.CursorLocation = 3 
	
		SQL_TechInfo = "SELECT * FROM SC_TechInfo"
		Set rsTechInfo = cnnTechInfo.Execute(SQL_TechInfo)
		
		If rsTechInfo("DailyRebuild_LastStatus") = "Finished" Then
		
			'Make sure the finished date is AFTER the last started date
			'That way, we know it's a new entry

			If rsTechInfo("DailyRebuild_Start") < rsTechInfo("DailyRebuild_Finish") Then RebuildDone = True
			
		End If
		
		Set rsTechInfo = Nothing
		cnnTechInfo.Close
		Set cnnTechInfo = Nothing

		If RebuildDone = True Then
			Response.Write("OK to check<br>")
			
			Set cnnInvHist = Server.CreateObject("ADODB.Connection")
			cnnInvHist.open (MUV_Read("ClientCnnString"))
			Set rsInvHist = Server.CreateObject("ADODB.Recordset")
			rsInvHist.CursorLocation = 3 
	
			SQL_InvHist = "SELECT Max(IvsDate) AS Expr1 FROM InvoiceHistory WHERE IvsDate <= getdate()" ' No future invoices
			Set rsInvHist = cnnInvHist.Execute(SQL_InvHist)
			
			If not rsInvHist.EOF Then MostRecentInvoiceDate = rsInvHist("Expr1")
			
			Set rsInvHist = Nothing
			cnnInvHist.Close
			Set cnnInvHist = Nothing
			
			Response.Write("Most recent invoice is " &  DateDiff("d",MostRecentInvoiceDate,Now()) & " days old<br>")
			Response.Write("Alert number of days is " & rsAlert("NumberOfDays") & "<br>")
			
			If DateDiff("d",MostRecentInvoiceDate,Now()) > rsAlert("NumberOfDays")  Then

				If AlertSent(rsAlert("InternalAlertRecNumber"),FormatDateTime(Now(),2)) <> True Then
					
					SendAlert rsAlert("InternalAlertRecNumber"),FormatDateTime(Now(),2),"HistOldInvoice",rsAlert("NotificationType")
				
				End If
				
			End If
			
		End If
		
	Case "RouteFileEmpty"
	
		Response.Write("Check route file empty <br>")
	
		Set cnnTechInfo = Server.CreateObject("ADODB.Connection")
		cnnTechInfo.open (MUV_Read("ClientCnnString"))
		Set rsTechInfo = Server.CreateObject("ADODB.Recordset")
		rsTechInfo.CursorLocation = 3 
	
		SQL_TechInfo = "SELECT Count(*) As RouteCount FROM RT_Truck"
		Set rsTechInfo = cnnTechInfo.Execute(SQL_TechInfo)
		
		If rsTechInfo("RouteCount") = 0 Then
		
			If AlertSent(rsAlert("InternalAlertRecNumber"),cdate(Month(Now()) & "/" &  Day(Now()) & "/" &  Year(Now()))) <> True Then
				
				SendAlert rsAlert("InternalAlertRecNumber"),cdate(Month(Now()) & "/" &  Day(Now()) & "/" &  Year(Now())),rsAlert("Condition"),rsAlert("NotificationType")
					
			End If
		End If
		
		Set rsTechInfo = Nothing
		cnnTechInfo.Close
		Set cnnTechInfo = Nothing
		


	Case "ProspectNoNextActivity"
	
		Response.Write("Prospect with no next activity <br>")
	
		Set cnnProspecting = Server.CreateObject("ADODB.Connection")
		cnnProspecting.open (MUV_Read("ClientCnnString"))
		Set rsProspecting = Server.CreateObject("ADODB.Recordset")
		rsProspecting.CursorLocation = 3 
	
		Set cnnProspectingInner = Server.CreateObject("ADODB.Connection")
		cnnProspectingInner.open (MUV_Read("ClientCnnString"))
		Set rsProspectingInner = Server.CreateObject("ADODB.Recordset")
		rsProspectingInner.CursorLocation = 3 
	
		SQL_Prospecting = "SELECT * FROM PR_Prospects WHERE Pool='Live'"
		Set rsProspecting = cnnProspecting.Execute(SQL_Prospecting)
		
		
		If NOT rsProspecting.EOF Then
		
			Do While NOT rsProspecting.EOF
			
				SQL_ProspectingInner = "SELECT * FROM PR_ProspectActivities WHERE ProspectRecID = " & rsProspecting("InternalRecordIdentifier") & " AND Status IS NULL "
				Set rsProspectingInner = cnnProspecting.Execute(SQL_ProspectingInner)
				
				If rsProspectingInner.EOF Then
				
					If AlertSentProspecting(rsAlert("InternalAlertRecNumber"),cdate(Month(Now()) & "/" &  Day(Now()) & "/" &  Year(Now())),rsProspecting("InternalRecordIdentifier")) <> True Then
						SendAlert rsAlert("InternalAlertRecNumber"),cdate(Month(Now()) & "/" &  Day(Now()) & "/" &  Year(Now())),rsAlert("Condition"),rsAlert("NotificationType")
					End If
			
				End If
			
			rsProspecting.MoveNext
			Loop
			
		End If
		
		
		Set rsProspecting = Nothing
		cnnProspecting.Close
		Set cnnProspecting = Nothing
		
		
		
End Select


%>