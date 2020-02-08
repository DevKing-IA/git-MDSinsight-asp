<%

'*************************************************************************************************
'GET FILTER VALUES FOR STATES
'*************************************************************************************************
	columnNames = ""

	Set cnnColumnFilter  = Server.CreateObject("ADODB.Connection")
	cnnColumnFilter.open Session("ClientCnnString")

	SQLColumnFilter  = "Select Distinct State from PR_Prospects"
	 
	Set rsColumnFilter  = Server.CreateObject("ADODB.Recordset")
	rsColumnFilter.CursorLocation = 3 
	
	Set rsColumnFilter = cnnColumnFilter.Execute(SQLColumnFilter)
				
	If NOT rsColumnFilter.EOF Then
	
		Response.write("<script type='text/javascript'>")
		Response.write("var state_methods = {")
	
		Do While NOT rsColumnFilter.EOF
			StateName = rsColumnFilter("State")
			columnNames = columnNames & "'" & StateName & "':'" & StateName & "',"
			rsColumnFilter.MoveNext
		Loop 
		
		If right(columnNames,1)= "," Then
			columnNames = Left(columnNames,Len(columnNames)-1)
		End If
		
		Response.write(columnNames)
		Response.write("};")
		Response.write("</script>")

	End If
			
	
	rsColumnFilter.Close
	set rsColumnFilter = Nothing
	cnnColumnFilter.Close	
	set cnnColumnFilter = Nothing

'*************************************************************************************************
'*************************************************************************************************	


'*************************************************************************************************
'GET FILTER VALUES FOR LEAD SOURCE 1
'*************************************************************************************************
	columnNames = ""

	Set cnnColumnFilter  = Server.CreateObject("ADODB.Connection")
	cnnColumnFilter.open Session("ClientCnnString")

	SQLColumnFilter  = "Select Distinct LeadSourceNumber from PR_Prospects"
	 
	Set rsColumnFilter  = Server.CreateObject("ADODB.Recordset")
	rsColumnFilter.CursorLocation = 3 
	
	Set rsColumnFilter = cnnColumnFilter.Execute(SQLColumnFilter)
				
	If NOT rsColumnFilter.EOF Then
	
		Response.write("<script type='text/javascript'>")
		Response.write("var lead_source_methods = {")
	
		Do While NOT rsColumnFilter.EOF
			LeadSourceNumber = rsColumnFilter("LeadSourceNumber")
			LeadSourceName = GetLeadSourceByNum(LeadSourceNumber)
			LeadSourceName = Replace(LeadSourceName,"'","")
			LeadSourceName = Replace(LeadSourceName,"/","")
			
			columnNames = columnNames & "'" & LeadSourceName & "':'" & LeadSourceName & "',"
		
			rsColumnFilter.MoveNext
			
		Loop 
		
		If right(columnNames,1)= "," Then
			columnNames = Left(columnNames,Len(columnNames)-1)
		End If
		
		Response.write(columnNames)
		Response.write("};")
		Response.write("</script>")

	End If
			
	
	rsColumnFilter.Close
	set rsColumnFilter = Nothing
	cnnColumnFilter.Close	
	set cnnColumnFilter = Nothing

'*************************************************************************************************
'*************************************************************************************************	



'*************************************************************************************************
'GET FILTER VALUES FOR STAGES
'*************************************************************************************************
	columnNames2 = ""

	Set cnnColumnFilter  = Server.CreateObject("ADODB.Connection")
	cnnColumnFilter.open Session("ClientCnnString")

	SQLColumnFilter  = "Select * FROM PR_Stages ORDER BY SortOrder"
	 
	Set rsColumnFilter  = Server.CreateObject("ADODB.Recordset")
	rsColumnFilter.CursorLocation = 3 
	
	Set rsColumnFilter = cnnColumnFilter.Execute(SQLColumnFilter)
				
	If NOT rsColumnFilter.EOF Then
	
		Response.write(vbcrlf & "<script type='text/javascript'>")
		Response.write("var stage_methods = {")
	
		Do While NOT rsColumnFilter.EOF
		
			StageNumber = rsColumnFilter("InternalRecordIdentifier")
			StageName = rsColumnFilter("Stage")
			StageName = Replace(StageName,"'","")
			StageName = Replace(StageName,"/","")
			
			columnNames2 = columnNames2 & "'" & StageName & "':'" & StageName & "',"
			
		
			rsColumnFilter.MoveNext
			
		Loop 
		
		If right(columnNames2,1)= "," Then
			columnNames2 = Left(columnNames2,Len(columnNames2)-1)
		End If
		
		Response.write(columnNames2)
		Response.write("};")
		Response.write("</script>")

	End If
			
	
	rsColumnFilter.Close
	set rsColumnFilter = Nothing
	cnnColumnFilter.Close	
	set cnnColumnFilter = Nothing

'*************************************************************************************************
'*************************************************************************************************	


'*************************************************************************************************
'GET FILTER VALUES FOR STAGE REASONS
'*************************************************************************************************
	columnNames2 = ""

	Set cnnColumnFilter  = Server.CreateObject("ADODB.Connection")
	cnnColumnFilter.open Session("ClientCnnString")

	SQLColumnFilter  = "Select * FROM PR_Reasons ORDER BY Reason"
	 
	Set rsColumnFilter  = Server.CreateObject("ADODB.Recordset")
	rsColumnFilter.CursorLocation = 3 
	
	Set rsColumnFilter = cnnColumnFilter.Execute(SQLColumnFilter)
				
	If NOT rsColumnFilter.EOF Then
	
		Response.write(vbcrlf & "<script type='text/javascript'>")
		Response.write("var stage_reason_methods = {")
	
		Do While NOT rsColumnFilter.EOF
		
			StageChangeReason = rsColumnFilter("Reason")
			StageChangeReason = Replace(StageChangeReason,"'","")
			StageChangeReason = Replace(StageChangeReason,"/","")
			
			columnNames2 = columnNames2 & "'" & StageChangeReason & "':'" & StageChangeReason & "',"
			
		
			rsColumnFilter.MoveNext
			
		Loop 
		
		If right(columnNames2,1)= "," Then
			columnNames2 = Left(columnNames2,Len(columnNames2)-1)
		End If
		
		Response.write(columnNames2)
		Response.write("};")
		Response.write("</script>")

	End If
			
	
	rsColumnFilter.Close
	set rsColumnFilter = Nothing
	cnnColumnFilter.Close	
	set cnnColumnFilter = Nothing

'*************************************************************************************************
'*************************************************************************************************	



'*************************************************************************************************
'GET FILTER VALUES FOR INDUSTRY
'*************************************************************************************************
	columnNames2 = ""

	Set cnnColumnFilter  = Server.CreateObject("ADODB.Connection")
	cnnColumnFilter.open Session("ClientCnnString")

	SQLColumnFilter  = "Select Distinct IndustryNumber from PR_Prospects"
	 
	Set rsColumnFilter  = Server.CreateObject("ADODB.Recordset")
	rsColumnFilter.CursorLocation = 3 
	
	Set rsColumnFilter = cnnColumnFilter.Execute(SQLColumnFilter)
				
	If NOT rsColumnFilter.EOF Then
	
		Response.write(vbcrlf & "<script type='text/javascript'>")
		Response.write("var industry_methods = {")
	
		Do While NOT rsColumnFilter.EOF
		
			IndustryNumber = rsColumnFilter("IndustryNumber")
			IndustryName = GetIndustryByNum(IndustryNumber)
			IndustryName = Replace(IndustryName,"'","")
			IndustryName = Replace(IndustryName,"/","")
			
			columnNames2 = columnNames2 & "'" & IndustryName & "':'" & IndustryName & "',"
			
		
			rsColumnFilter.MoveNext
			
		Loop 
		
		If right(columnNames2,1)= "," Then
			columnNames2 = Left(columnNames2,Len(columnNames2)-1)
		End If
		
		Response.write(columnNames2)
		Response.write("};")
		Response.write("</script>")

	End If
			
	
	rsColumnFilter.Close
	set rsColumnFilter = Nothing
	cnnColumnFilter.Close	
	set cnnColumnFilter = Nothing

'*************************************************************************************************
'*************************************************************************************************	


'*************************************************************************************************
'GET FILTER VALUES FOR EMPLOYEE RANGE
'*************************************************************************************************
	columnNames2 = ""

	Set cnnColumnFilter  = Server.CreateObject("ADODB.Connection")
	cnnColumnFilter.open Session("ClientCnnString")

	SQLColumnFilter  = "Select Distinct EmployeeRangeNumber from PR_Prospects"
	 
	Set rsColumnFilter  = Server.CreateObject("ADODB.Recordset")
	rsColumnFilter.CursorLocation = 3 
	
	Set rsColumnFilter = cnnColumnFilter.Execute(SQLColumnFilter)
				
	If NOT rsColumnFilter.EOF Then
	
		Response.write(vbcrlf & "<script type='text/javascript'>")
		Response.write("var employee_methods = {")
	
		Do While NOT rsColumnFilter.EOF
		
			EmployeeRangeNumber = rsColumnFilter("EmployeeRangeNumber")
			EmployeeRange = GetEmployeeRangeByNum(EmployeeRangeNumber)
			
			columnNames2 = columnNames2 & "'" & EmployeeRange & "':'" & EmployeeRange & "',"
			
			rsColumnFilter.MoveNext
			
		Loop 
		
		If right(columnNames2,1)= "," Then
			columnNames2 = Left(columnNames2,Len(columnNames2)-1)
		End If
		
		Response.write(columnNames2)
		Response.write("};")
		Response.write("</script>")

	End If
			
	
	rsColumnFilter.Close
	set rsColumnFilter = Nothing
	cnnColumnFilter.Close	
	set cnnColumnFilter = Nothing

'*************************************************************************************************
'*************************************************************************************************	


'*************************************************************************************************
'GET FILTER VALUES FOR RECORD OWNER
'*************************************************************************************************
	columnNames2 = ""

	Set cnnColumnFilter  = Server.CreateObject("ADODB.Connection")
	cnnColumnFilter.open Session("ClientCnnString")

	SQLColumnFilter  = "Select Distinct OwnerUserNo from PR_Prospects"
	 
	Set rsColumnFilter  = Server.CreateObject("ADODB.Recordset")
	rsColumnFilter.CursorLocation = 3 
	
	Set rsColumnFilter = cnnColumnFilter.Execute(SQLColumnFilter)
				
	If NOT rsColumnFilter.EOF Then
	
		Response.write(vbcrlf & "<script type='text/javascript'>")
		Response.write("var owner_methods = {")
	
		Do While NOT rsColumnFilter.EOF
		
			OwnerUserNo = rsColumnFilter("OwnerUserNo")
			OwnerName = GetUserDisplayNameByUserNo(OwnerUserNo)
			OwnerName = Replace(OwnerName,"'","")
			OwnerName = Replace(OwnerName,"/","")
			
			columnNames2 = columnNames2 & "'" & OwnerName & "':'" & OwnerName & "',"
			
		
			rsColumnFilter.MoveNext
			
		Loop 
		
		If right(columnNames2,1)= "," Then
			columnNames2 = Left(columnNames2,Len(columnNames2)-1)
		End If
		
		Response.write(columnNames2)
		Response.write("};")
		Response.write("</script>")

	End If
			
	
	rsColumnFilter.Close
	set rsColumnFilter = Nothing
	cnnColumnFilter.Close	
	set cnnColumnFilter = Nothing

'*************************************************************************************************
'*************************************************************************************************	

'*************************************************************************************************
'GET FILTER VALUES FOR CREATED BY
'*************************************************************************************************
	columnNames2 = ""

	Set cnnColumnFilter  = Server.CreateObject("ADODB.Connection")
	cnnColumnFilter.open Session("ClientCnnString")

	SQLColumnFilter  = "Select Distinct CreatedByUserNo from PR_Prospects"
	 
	Set rsColumnFilter  = Server.CreateObject("ADODB.Recordset")
	rsColumnFilter.CursorLocation = 3 
	
	Set rsColumnFilter = cnnColumnFilter.Execute(SQLColumnFilter)
				
	If NOT rsColumnFilter.EOF Then
	
		Response.write(vbcrlf & "<script type='text/javascript'>")
		Response.write("var createdby_methods = {")
	
		Do While NOT rsColumnFilter.EOF
		
			CreatedByUserNo = rsColumnFilter("CreatedByUserNo")
			CreatedByName = GetUserDisplayNameByUserNo(CreatedByUserNo)
			CreatedByName = Replace(CreatedByName,"'","")
			CreatedByName = Replace(CreatedByName,"/","")
			
			columnNames2 = columnNames2 & "'" & CreatedByName & "':'" & CreatedByName & "',"
			
		
			rsColumnFilter.MoveNext
			
		Loop 
		
		If right(columnNames2,1)= "," Then
			columnNames2 = Left(columnNames2,Len(columnNames2)-1)
		End If
		
		Response.write(columnNames2)
		Response.write("};")
		Response.write("</script>")

	End If
			
	
	rsColumnFilter.Close
	set rsColumnFilter = Nothing
	cnnColumnFilter.Close	
	set cnnColumnFilter = Nothing

'*************************************************************************************************
'*************************************************************************************************	


'*************************************************************************************************
'GET FILTER VALUES FOR TELEMARKETER
'*************************************************************************************************
	columnNames2 = ""

	Set cnnColumnFilter  = Server.CreateObject("ADODB.Connection")
	cnnColumnFilter.open Session("ClientCnnString")

	SQLColumnFilter  = "Select Distinct TelemarketerUserNo from PR_Prospects"
	 
	Set rsColumnFilter  = Server.CreateObject("ADODB.Recordset")
	rsColumnFilter.CursorLocation = 3 
	
	Set rsColumnFilter = cnnColumnFilter.Execute(SQLColumnFilter)
				
	If NOT rsColumnFilter.EOF Then
	
		Response.write(vbcrlf & "<script type='text/javascript'>")
		Response.write("var telemarketer_methods = {")
	
		Do While NOT rsColumnFilter.EOF
		
			TelemarketerUserNo = rsColumnFilter("TelemarketerUserNo")
			TelemarketerName = GetUserDisplayNameByUserNo(TelemarketerUserNo)
			TelemarketerName = Replace(TelemarketerName,"'","")
			TelemarketerName = Replace(TelemarketerName,"/","")
			
			columnNames2 = columnNames2 & "'" & TelemarketerName & "':'" & TelemarketerName & "',"
			
		
			rsColumnFilter.MoveNext
			
		Loop 
		
		If right(columnNames2,1)= "," Then
			columnNames2 = Left(columnNames2,Len(columnNames2)-1)
		End If
		
		Response.write(columnNames2)
		Response.write("};")
		Response.write("</script>")

	End If
			
	
	rsColumnFilter.Close
	set rsColumnFilter = Nothing
	cnnColumnFilter.Close	
	set cnnColumnFilter = Nothing

'*************************************************************************************************
'*************************************************************************************************	

'*************************************************************************************************
'GET FILTER VALUES FOR NEXT ACTIVITY
'*************************************************************************************************
	columnNames2 = ""

	Set cnnColumnFilter  = Server.CreateObject("ADODB.Connection")
	cnnColumnFilter.open Session("ClientCnnString")

	SQLColumnFilter  = "Select Distinct ActivityRecID from PR_Prospects INNER JOIN PR_ProspectActivities ON PR_ProspectActivities.ProspectRecID = PR_Prospects.InternalRecordIdentifier"
	 
	Set rsColumnFilter  = Server.CreateObject("ADODB.Recordset")
	rsColumnFilter.CursorLocation = 3 
	
	Set rsColumnFilter = cnnColumnFilter.Execute(SQLColumnFilter)
				
	If NOT rsColumnFilter.EOF Then
	
		Response.write(vbcrlf & "<script type='text/javascript'>")
		Response.write("var activity_methods = {")
	
		Do While NOT rsColumnFilter.EOF
		
			ActivityRecID = rsColumnFilter("ActivityRecID")
			ActivityName = GetActivityByNum(ActivityRecID)
			ActivityName = Replace(ActivityName,"'","")
			ActivityName = Replace(ActivityName,"/","")
			
			columnNames2 = columnNames2 & "'" & ActivityName & "':'" & ActivityName & "',"
			
		
			rsColumnFilter.MoveNext
			
		Loop 
		
		If right(columnNames2,1)= "," Then
			columnNames2 = Left(columnNames2,Len(columnNames2)-1)
		End If
		
		Response.write(columnNames2)
		Response.write("};")
		Response.write("</script>")

	End If
			
	
	rsColumnFilter.Close
	set rsColumnFilter = Nothing
	cnnColumnFilter.Close	
	set cnnColumnFilter = Nothing

'*************************************************************************************************
'*************************************************************************************************	



	%>
