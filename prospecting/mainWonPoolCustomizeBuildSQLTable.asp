<%
'Response.Write ("OK, you got here")
ActivityDatePopulated = False
	
Set cnnBuildSQL = Server.CreateObject("ADODB.Connection")
cnnBuildSQL.open (Session("ClientCnnString"))
Set rsBuildSQL = Server.CreateObject("ADODB.Recordset")
rsBuildSQL.CursorLocation = 3 
Set rsBuildSQL2 = Server.CreateObject("ADODB.Recordset")
rsBuildSQL2.CursorLocation = 3 

On Error Resume Next ' In caase the table isn't there
SQLBuildSQL = "DROP TABLE zProspectFilter_" & trim(Session("Userno"))
Set rsBuildSQL  = cnnBuildSQL.Execute(SQLBuildSQL)
On Error Goto 0


'Start by moving everything into the temp table
SQLBuildSQL = "SELECT * "
SQLBuildSQL = "SELECT InternalRecordIdentifier, City, State, PostalCode, LeadSourceNumber, IndustryNumber, "
SQLBuildSQL = SQLBuildSQL & "EmployeeRangeNumber, OwnerUserNo, CreatedDate, CreatedByUserNo, "
SQLBuildSQL = SQLBuildSQL & " TelemarketerUserNo , NumberOfPantries "
SQLBuildSQL = SQLBuildSQL & "INTO  zProspectFilter_" & trim(Session("Userno")) & " FROM PR_Prospects WHERE Pool='Won'"
Set rsBuildSQL = cnnBuildSQL.Execute(SQLBuildSQL) ' Maybe, , adAsyncExecute)


'Need to do this all the time to allow for sort order
SQLBuildSQL = "ALTER TABLE zProspectFilter_" & trim(Session("Userno")) & " ADD NextActivityDueDate datetime NULL"
'Response.Write(SQLBuildSQL & "<BR>")
Set rsBuildSQL = cnnBuildSQL.Execute(SQLBuildSQL)

SQLBuildSQL = "CREATE NONCLUSTERED INDEX [IX_zProspectFilter_2] ON [zProspectFilter_" & trim(Session("Userno")) & "] "
SQLBuildSQL = SQLBuildSQL & "( [NextActivityDueDate] DESC "
SQLBuildSQL = SQLBuildSQL & ")WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]"
Set rsBuildSQL = cnnBuildSQL.Execute(SQLBuildSQL)

SQLReportName = Replace(MUV_READ("CRMVIEWSTATEWONPOOL"),"'","''")

'Now get all the report specific data fields
SQLBuildSQL = "SELECT * FROM Settings_Reports WHERE ReportNumber = 1400 AND UserNo = " & Trim(Session("Userno")) & " AND PoolForProspecting = 'Won' AND UserReportName = '" & SQLReportName & "'"
Set rsBuildSQL = cnnBuildSQL.Execute(SQLBuildSQL)
If Not rsBuildSQL.Eof Then
	ReportSpecificData1 = rsBuildSQL("ReportSpecificData1") 
	ReportSpecificData2 = rsBuildSQL("ReportSpecificData2") 
	ReportSpecificData3 = rsBuildSQL("ReportSpecificData3") 
	ReportSpecificData4 = rsBuildSQL("ReportSpecificData4") 
	ReportSpecificData5 = rsBuildSQL("ReportSpecificData5") 	
	ReportSpecificData6 = rsBuildSQL("ReportSpecificData6") 	
	ReportSpecificData7 = rsBuildSQL("ReportSpecificData7") 	
	ReportSpecificData8 = rsBuildSQL("ReportSpecificData8") 	
	ReportSpecificData9 = rsBuildSQL("ReportSpecificData9") 	
	ReportSpecificData10 = rsBuildSQL("ReportSpecificData10") 	
	ReportSpecificData11 = rsBuildSQL("ReportSpecificData11") 	
	ReportSpecificData12 = rsBuildSQL("ReportSpecificData12") 	
	ReportSpecificData13 = rsBuildSQL("ReportSpecificData13") 	
	ReportSpecificData14 = rsBuildSQL("ReportSpecificData14") 	
	ReportSpecificData15 = rsBuildSQL("ReportSpecificData15") 	
	ReportSpecificData16 = rsBuildSQL("ReportSpecificData16") 	
	ReportSpecificData17 = rsBuildSQL("ReportSpecificData17") 	
	ReportSpecificData18 = rsBuildSQL("ReportSpecificData18") 	
	ReportSpecificData19 = rsBuildSQL("ReportSpecificData19") 		
	ReportSpecificData20 = rsBuildSQL("ReportSpecificData20") 
	ReportSpecificData21 = rsBuildSQL("ReportSpecificData21") 
	ReportSpecificData22 = rsBuildSQL("ReportSpecificData22") 		
End If
	

' If STAGE filtering is involved, we open these here
' They stay open until the page is one to assist
' with speed
If Not IsNull(ReportSpecificData18) Then
	If ReportSpecificData18 <> ""  Then
		Set rsStagesTmp = Server.CreateObject("ADODB.Recordset")
		rsStagesTmp.CursorLocation = 3 
	End If
ElseIf Not IsNull(ReportSpecificData2) Then
	If ReportSpecificData2 <> ""  Then
		Set rsStagesTmp = Server.CreateObject("ADODB.Recordset")
		rsStagesTmp.CursorLocation = 3 
	End If
End If

' ReportSpecificData6 - Lead Source Number
If Not IsNull(ReportSpecificData6) Then
	If ReportSpecificData6 <> ""  Then
		SQLBuildSQL = "DELETE FROM zProspectFilter_" & trim(Session("Userno"))
		SQLBuildSQL = SQLBuildSQL & " WHERE LeadSourceNumber NOT IN (" & ReportSpecificData6 & ")"
		'Response.Write(SQLBuildSQL & "<BR>")
		Set rsBuildSQL = cnnBuildSQL.Execute(SQLBuildSQL)
	End If
End IF

' ReportSpecificData7 - Industry Number
If Not IsNull(ReportSpecificData7) Then
	If ReportSpecificData7 <> ""  Then
		SQLBuildSQL = "DELETE FROM zProspectFilter_" & trim(Session("Userno"))
		SQLBuildSQL = SQLBuildSQL & " WHERE IndustryNumber NOT IN (" & ReportSpecificData7 & ")"
		'Response.Write(SQLBuildSQL & "<BR>")
		Set rsBuildSQL = cnnBuildSQL.Execute(SQLBuildSQL)
	End If
End IF

' ReportSpecificData8 - Telemarketer
If Not IsNull(ReportSpecificData8) Then
	If ReportSpecificData8 <> ""  Then
		SQLBuildSQL = "DELETE FROM zProspectFilter_" & trim(Session("Userno"))
		SQLBuildSQL = SQLBuildSQL & " WHERE TelemarketerUserNo NOT IN (" & ReportSpecificData8 & ")"
		'Response.Write(SQLBuildSQL & "<BR>")
		Set rsBuildSQL = cnnBuildSQL.Execute(SQLBuildSQL)
	End If
End IF

' ReportSpecificData9 - Owner
If Not IsNull(ReportSpecificData9) Then
	If ReportSpecificData9 <> ""  Then
		SQLBuildSQL = "DELETE FROM zProspectFilter_" & trim(Session("Userno"))
		SQLBuildSQL = SQLBuildSQL & " WHERE OwnerUserNo NOT IN (" & ReportSpecificData9 & ")"
		'Response.Write(SQLBuildSQL & "<BR>")
		Set rsBuildSQL = cnnBuildSQL.Execute(SQLBuildSQL)
	End If
End IF

' ReportSpecificData10 - CreatedByUserNo
If Not IsNull(ReportSpecificData10) Then
	If ReportSpecificData10 <> ""  Then
		SQLBuildSQL = "DELETE FROM zProspectFilter_" & trim(Session("Userno"))
		SQLBuildSQL = SQLBuildSQL & " WHERE CreatedByUserNo NOT IN (" & ReportSpecificData10 & ")"
		'Response.Write(SQLBuildSQL & "<BR>")
		Set rsBuildSQL = cnnBuildSQL.Execute(SQLBuildSQL)
	End If
End IF

' ReportSpecificData15 - City
If Not IsNull(ReportSpecificData15) Then
	If ReportSpecificData15 <> ""  Then
		SQLBuildSQL = "DELETE FROM zProspectFilter_" & trim(Session("Userno"))
		SQLBuildSQL = SQLBuildSQL & " WHERE City NOT IN (" 
		ReportSpecificData15Arr = Split(ReportSpecificData15,",")
		For x = 0 To UBound(ReportSpecificData15Arr)
			SQLBuildSQL = SQLBuildSQL & "'" & ReportSpecificData15Arr(x) & "',"
		Next
		SQLBuildSQL = Left(SQLBuildSQL,Len(SQLBuildSQL)-1)
		SQLBuildSQL = SQLBuildSQL & ")"
		'Response.Write(SQLBuildSQL & "<BR>")
		Set rsBuildSQL = cnnBuildSQL.Execute(SQLBuildSQL)
	End If
End IF

' ReportSpecificData16 - State
If Not IsNull(ReportSpecificData16) Then
	If ReportSpecificData16 <> ""  Then
		SQLBuildSQL = "DELETE FROM zProspectFilter_" & trim(Session("Userno"))
		SQLBuildSQL = SQLBuildSQL & " WHERE State NOT IN (" 
		ReportSpecificData16Arr = Split(ReportSpecificData16,",")
		For x = 0 To UBound(ReportSpecificData16Arr)
			SQLBuildSQL = SQLBuildSQL & "'" & ReportSpecificData16Arr(x) & "',"
		Next
		SQLBuildSQL = Left(SQLBuildSQL,Len(SQLBuildSQL)-1)
		SQLBuildSQL = SQLBuildSQL & ")"
		'Response.Write(SQLBuildSQL & "<BR>")
		Set rsBuildSQL = cnnBuildSQL.Execute(SQLBuildSQL)
	End If
End IF

' ReportSpecificData17 - Zip Code
If Not IsNull(ReportSpecificData17) Then
	If ReportSpecificData17 <> ""  Then
		SQLBuildSQL = "DELETE FROM zProspectFilter_" & trim(Session("Userno"))
		SQLBuildSQL = SQLBuildSQL & " WHERE PostalCode NOT IN (" 
		ReportSpecificData17Arr = Split(ReportSpecificData17,",")
		For x = 0 To UBound(ReportSpecificData17Arr)
			SQLBuildSQL = SQLBuildSQL & "'" & ReportSpecificData17Arr(x) & "',"
		Next
		SQLBuildSQL = Left(SQLBuildSQL,Len(SQLBuildSQL)-1)
		SQLBuildSQL = SQLBuildSQL & ")"
		'Response.Write(SQLBuildSQL & "<BR>")
		Set rsBuildSQL = cnnBuildSQL.Execute(SQLBuildSQL)
	End If
End IF

' ReportSpecificData11 & ReportSpecificData12 - Created Range
If Not IsNull(ReportSpecificData11) Then
	If ReportSpecificData11 <> ""  Then
		StartDateRange = Year(ReportSpecificData11) 
		If Month(ReportSpecificData11) < 10 Then StartDateRange = StartDateRange & "0"
		StartDateRange = StartDateRange & Month(ReportSpecificData11)
		If Day(ReportSpecificData11) < 10 Then StartDateRange = StartDateRange & "0"
		StartDateRange = StartDateRange & Day(ReportSpecificData11)
		EndDateRange = Year(ReportSpecificData12) 
		If Month(ReportSpecificData12) < 10 Then EndDateRange = EndDateRange & "0"
		EndDateRange = EndDateRange & Month(ReportSpecificData12)
		If Day(ReportSpecificData12) < 10 Then EndDateRange = EndDateRange & "0"
		EndDateRange = EndDateRange & Day(ReportSpecificData12)
		SQLBuildSQL = "DELETE FROM zProspectFilter_" & trim(Session("Userno")) & " WHERE Cast(CreatedDate as date) NOT BETWEEN '" & StartDateRange & "' AND '" & EndDateRange  & "'"
		'Response.Write(SQLBuildSQL & "<BR>")
		Set rsBuildSQL = cnnBuildSQL.Execute(SQLBuildSQL)
	End If
End IF


' ReportSpecificData14 - Number Of Pantries
If Not IsNull(ReportSpecificData14) Then
	If ReportSpecificData14 <> ""  Then
		ReportSpecificData14Arr = Split(ReportSpecificData14,",")
		Select Case ReportSpecificData14Arr(0)
			Case "ByCustomRange"	
				StartRange = ReportSpecificData14Arr(1)
				EndRange = ReportSpecificData14Arr(2)
				SQLBuildSQL = "DELETE FROM zProspectFilter_" & trim(Session("Userno")) & " WHERE NumberOfPantries NOT BETWEEN " & StartRange & " AND " & EndRange 
				'Response.Write(SQLBuildSQL & "<BR>")
				Set rsBuildSQL = cnnBuildSQL.Execute(SQLBuildSQL)
			Case "ByCustomNumber"
				ComparisionOperator = ReportSpecificData14Arr(1)
				ComparisonNumberSingle = ReportSpecificData14Arr(2)
				Select Case ComparisionOperator
					Case "equal to"
						SQLBuildSQL = "DELETE FROM zProspectFilter_" & trim(Session("Userno")) & " WHERE NumberOfPantries <> " & cint(ComparisonNumberSingle)
						'Response.Write(SQLBuildSQL & "<BR>")
						Set rsBuildSQL = cnnBuildSQL.Execute(SQLBuildSQL)
					Case "greater than"
						SQLBuildSQL = "DELETE FROM zProspectFilter_" & trim(Session("Userno")) & " WHERE NumberOfPantries <= " & cint(ComparisonNumberSingle)
						'Response.Write(SQLBuildSQL & "<BR>")
						Set rsBuildSQL = cnnBuildSQL.Execute(SQLBuildSQL)
					Case "greater than or equal to"
						SQLBuildSQL = "DELETE FROM zProspectFilter_" & trim(Session("Userno")) & " WHERE NumberOfPantries < " & cint(ComparisonNumberSingle)
						'Response.Write(SQLBuildSQL & "<BR>")
						Set rsBuildSQL = cnnBuildSQL.Execute(SQLBuildSQL)
					Case "less than"
						SQLBuildSQL = "DELETE FROM zProspectFilter_" & trim(Session("Userno")) & " WHERE NumberOfPantries >= " & cint(ComparisonNumberSingle)
						'Response.Write(SQLBuildSQL & "<BR>")
						Set rsBuildSQL = cnnBuildSQL.Execute(SQLBuildSQL)
					Case "less than or equal to"
						SQLBuildSQL = "DELETE FROM zProspectFilter_" & trim(Session("Userno")) & " WHERE NumberOfPantries > " & cint(ComparisonNumberSingle)
					'Response.Write(SQLBuildSQL & "<BR>")
						Set rsBuildSQL = cnnBuildSQL.Execute(SQLBuildSQL)
				End Select
		End Select
	End If
End IF


' ReportSpecificData18 - Stages to filter
If Not IsNull(ReportSpecificData18) Then
	If ReportSpecificData18 <> ""  Then
		SQLBuildSQL = "ALTER TABLE zProspectFilter_" & trim(Session("Userno")) & " ADD StageNumber int NULL"
		'Response.Write(SQLBuildSQL & "<BR>")
		Set rsBuildSQL = cnnBuildSQL.Execute(SQLBuildSQL)
		' Now populate it
		SQLBuildSQL = "SELECT InternalRecordIdentifier FROM zProspectFilter_" & trim(Session("Userno"))
		'Response.Write(SQLBuildSQL & "<BR>")
		Set rsBuildSQL = cnnBuildSQL.Execute(SQLBuildSQL)
		If Not rsBuildSQL.EOF Then
			Do While Not rsBuildSQL.EOF
				' GetProspectCurrentStageByProspectNumber
				SQLStagesTMP = "Select Top 1 * from PR_ProspectStages Where ProspectRecID = " & rsBuildSQL("InternalRecordIdentifier") & " ORDER BY RecordCReationDateTime DESC"
				Set rsStagesTMP = cnnBuildSQL.Execute(SQLStagesTMP)
				If not rsStagesTMP.EOF Then resultStageNumberTMP = rsStagesTMP("StageRecID") Else resultStageNumberTMP = 0
				SQLBuildSQL2 = "UPDATE zProspectFilter_" & trim(Session("Userno")) & " SET StageNumber = '" & resultStageNumberTMP & "'"
				SQLBuildSQL2 = SQLBuildSQL2 & " WHERE InternalRecordIdentifier = " & rsBuildSQL("InternalRecordIdentifier")
				'Response.Write(SQLBuildSQL2 & "<BR>")
				Set rsBuildSQL2 = cnnBuildSQL.Execute(SQLBuildSQL2)
				rsBuildSQL.movenext
			Loop
		End If
		SQLBuildSQL = "DELETE FROM zProspectFilter_" & trim(Session("Userno"))
		SQLBuildSQL = SQLBuildSQL & " WHERE StageNumber NOT IN (" & ReportSpecificData18 & ")"
		'Response.Write(SQLBuildSQL & "<BR>")
		Set rsBuildSQL = cnnBuildSQL.Execute(SQLBuildSQL)
	End If
End IF


' ReportSpecificData2 - Stage Filter Type
If Not IsNull(ReportSpecificData2) Then
	If ReportSpecificData2 <> ""  Then
		SQLBuildSQL = "ALTER TABLE zProspectFilter_" & trim(Session("Userno")) & " ADD LastStageChangeDate datetime NULL"
		'Response.Write(SQLBuildSQL & "<BR>")
		Set rsBuildSQL = cnnBuildSQL.Execute(SQLBuildSQL)
		' Now populate it
		SQLBuildSQL = "SELECT InternalRecordIdentifier FROM zProspectFilter_" & trim(Session("Userno"))
		'Response.Write(SQLBuildSQL & "<BR>")
		Set rsBuildSQL = cnnBuildSQL.Execute(SQLBuildSQL)
		If Not rsBuildSQL.EOF Then
			Do While Not rsBuildSQL.EOF
				SQLBuildSQL2 = "UPDATE zProspectFilter_" & trim(Session("Userno")) & " SET LastStageChangeDate = '" & GetProspectLastStageChangeDateByProspectNumber(rsBuildSQL("InternalRecordIdentifier")) & "'"
				SQLBuildSQL2 = SQLBuildSQL2 & " WHERE InternalRecordIdentifier = " & rsBuildSQL("InternalRecordIdentifier")
				'Response.Write(SQLBuildSQL2 & "<BR>")
				Set rsBuildSQL2 = cnnBuildSQL.Execute(SQLBuildSQL2)
				rsBuildSQL.movenext
			Loop
		End If
		Select Case ReportSpecificData2
			Case "HasNotChangedInXDays"
				SQLBuildSQL = "DELETE FROM zProspectFilter_" & trim(Session("Userno")) & " WHERE DateDiff(d,LastStageChangeDate,getdate()) < " & ReportSpecificData3 
				'Response.Write(SQLBuildSQL & "<BR>")
				Set rsBuildSQL = cnnBuildSQL.Execute(SQLBuildSQL)
			Case "HasChangedInXDays"
				SQLBuildSQL = "DELETE FROM zProspectFilter_" & trim(Session("Userno")) & " WHERE DateDiff(d,LastStageChangeDate,getdate()) >= " & ReportSpecificData3 
				'Response.Write(SQLBuildSQL & "<BR>")
				Set rsBuildSQL = cnnBuildSQL.Execute(SQLBuildSQL)
			Case "WasUnqualifiedInDateRange"
				StartDateRange = Year(ReportSpecificData4) 
				If Month(ReportSpecificData4) < 10 Then StartDateRange = StartDateRange & "0"
				StartDateRange = StartDateRange & Month(ReportSpecificData4)
				If Day(ReportSpecificData4) < 10 Then StartDateRange = StartDateRange & "0"
				StartDateRange = StartDateRange & Day(ReportSpecificData4)
				EndDateRange = Year(ReportSpecificData5) 
				If Month(ReportSpecificData5) < 10 Then EndDateRange = EndDateRange & "0"
				EndDateRange = EndDateRange & Month(ReportSpecificData5)
				If Day(ReportSpecificData5) < 10 Then EndDateRange = EndDateRange & "0"
				EndDateRange = EndDateRange & Day(ReportSpecificData5)
				SQLBuildSQL = "DELETE FROM zProspectFilter_" & trim(Session("Userno")) & " WHERE LastStageChangeDate BETWEEN '" & StartDateRange & "' AND '" & EndDateRange  & "'"
				'Response.Write(SQLBuildSQL & "<BR>")
				Set rsBuildSQL = cnnBuildSQL.Execute(SQLBuildSQL)
			Case "WasLostInDateRange"
				StartDateRange = Year(ReportSpecificData4) 
				If Month(ReportSpecificData4) < 10 Then StartDateRange = StartDateRange & "0"
				StartDateRange = StartDateRange & Month(ReportSpecificData4)
				If Day(ReportSpecificData4) < 10 Then StartDateRange = StartDateRange & "0"
				StartDateRange = StartDateRange & Day(ReportSpecificData4)
				EndDateRange = Year(ReportSpecificData5) 
				If Month(ReportSpecificData5) < 10 Then EndDateRange = EndDateRange & "0"
				EndDateRange = EndDateRange & Month(ReportSpecificData5)
				If Day(ReportSpecificData5) < 10 Then EndDateRange = EndDateRange & "0"
				EndDateRange = EndDateRange & Day(ReportSpecificData5)
				SQLBuildSQL = "DELETE FROM zProspectFilter_" & trim(Session("Userno")) & " WHERE LastStageChangeDate NOT BETWEEN '" & StartDateRange & "' AND '" & EndDateRange  & "'"
				'Response.Write(SQLBuildSQL & "<BR>")
				Set rsBuildSQL = cnnBuildSQL.Execute(SQLBuildSQL)
		End Select
	End If
End IF


' ReportSpecificData13 - Number Of Emplyees
If Not IsNull(ReportSpecificData13) Then
	If ReportSpecificData13 <> ""  Then
		ReportSpecificData13Arr = Split(ReportSpecificData13,",")

		If ReportSpecificData13Arr(0) <> "ByPredefinedRange" Then
			SQLBuildSQL = "ALTER TABLE zProspectFilter_" & trim(Session("Userno")) & " ADD NumberOfEmployeesStart int NULL, NumberOfEmployeesEnd int NULL"
			'Response.Write(SQLBuildSQL & "<BR>")
			Set rsBuildSQL = cnnBuildSQL.Execute(SQLBuildSQL)
			' Now populate it
			SQLBuildSQL = "SELECT InternalRecordIdentifier,EmployeeRangeNumber FROM zProspectFilter_" & trim(Session("Userno"))
			'Response.Write(SQLBuildSQL & "<BR>")
			Set rsBuildSQL = cnnBuildSQL.Execute(SQLBuildSQL)
			If Not rsBuildSQL.EOF Then
				Do While Not rsBuildSQL.EOF
					SQLBuildSQL2 = "UPDATE zProspectFilter_" & trim(Session("Userno")) 
					SQLBuildSQL2 = SQLBuildSQL2 & " SET NumberOfEmployeesStart = " & GetNumEmployeeStartByEmployeeRangeNo(rsBuildSQL("EmployeeRangeNumber")) 
					SQLBuildSQL2 = SQLBuildSQL2 & ", NumberOfEmployeesEnd = " & GetNumEmployeeEndByEmployeeRangeNo(rsBuildSQL("EmployeeRangeNumber"))
					SQLBuildSQL2 = SQLBuildSQL2 & " WHERE InternalRecordIdentifier = " & rsBuildSQL("InternalRecordIdentifier")
					'Response.Write(SQLBuildSQL2 & "<BR>")
					Set rsBuildSQL2 = cnnBuildSQL.Execute(SQLBuildSQL2)
					rsBuildSQL.movenext
				Loop
			End If
		End If
		Select Case ReportSpecificData13Arr(0)
			Case "ByPredefinedRange"
				SQLBuildSQL = "DELETE FROM zProspectFilter_" & trim(Session("Userno")) & " WHERE EmployeeRangeNumber <> '" & ReportSpecificData13Arr(1) & "'"
				'Response.Write(SQLBuildSQL & "<BR>")
				Set rsBuildSQL = cnnBuildSQL.Execute(SQLBuildSQL)
			Case "ByCustomNumber"
				ComparisionOperator = ReportSpecificData13Arr(1)
				ComparisonNumberSingle = ReportSpecificData13Arr(2)
				Select Case ComparisionOperator
					Case "equal to"
						SQLBuildSQL = "DELETE FROM zProspectFilter_" & trim(Session("Userno")) & " WHERE NumberOfEmployeesStart <> " & cint(ComparisonNumberSingle)
						'Response.Write(SQLBuildSQL & "<BR>")
						Set rsBuildSQL = cnnBuildSQL.Execute(SQLBuildSQL)
					Case "greater than"
						SQLBuildSQL = "DELETE FROM zProspectFilter_" & trim(Session("Userno")) & " WHERE NumberOfEmployeesStart <= " & cint(ComparisonNumberSingle)
						'Response.Write(SQLBuildSQL & "<BR>")
						Set rsBuildSQL = cnnBuildSQL.Execute(SQLBuildSQL)
					Case "greater than or equal to"
						SQLBuildSQL = "DELETE FROM zProspectFilter_" & trim(Session("Userno")) & " WHERE NumberOfEmployeesStart < " & cint(ComparisonNumberSingle)
						'Response.Write(SQLBuildSQL & "<BR>")
						Set rsBuildSQL = cnnBuildSQL.Execute(SQLBuildSQL)
					Case "less than"
						SQLBuildSQL = "DELETE FROM zProspectFilter_" & trim(Session("Userno")) & " WHERE NumberOfEmployeesEnd >= " & cint(ComparisonNumberSingle)
						'Response.Write(SQLBuildSQL & "<BR>")
						Set rsBuildSQL = cnnBuildSQL.Execute(SQLBuildSQL)
					Case "less than or equal to"
						SQLBuildSQL = "DELETE FROM zProspectFilter_" & trim(Session("Userno")) & " WHERE NumberOfEmployeesEnd > " & cint(ComparisonNumberSingle)
						'Response.Write(SQLBuildSQL & "<BR>")
						Set rsBuildSQL = cnnBuildSQL.Execute(SQLBuildSQL)
				End Select
			Case "ByCustomRange"
				ComparisionLowEndOfRange = ReportSpecificData13Arr(1)
				ComparisonHighEndOfRange = ReportSpecificData13Arr(2)
				SQLBuildSQL = "DELETE FROM zProspectFilter_" & trim(Session("Userno")) & " WHERE NumberOfEmployeesStart < " & ComparisionLowEndOfRange & " OR NumberOfEmployeesEnd > " & ComparisonHighEndOfRange 
				'Response.Write(SQLBuildSQL & "<BR>")
				Set rsBuildSQL = cnnBuildSQL.Execute(SQLBuildSQL)
		End Select
	End If
End IF


If Not IsNull(ReportSpecificData18) Then
	If ReportSpecificData18 <> ""  Then
		Set rsStagesTmp = Nothing
	End If
ElseIf Not IsNull(ReportSpecificData2) Then
	If ReportSpecificData2 <> ""  Then
		Set rsStagesTmp = Nothing
	End If
End If

' ReportSpecificData20 & ReportSpecificData21 & ReportSpecificData22 - Next Activity Date Range
If Not IsNull(ReportSpecificData20) Then
	If ReportSpecificData20 <> ""  Then
		' Now populate it
		SQLBuildSQL = "UPDATE zProspectFilter_" & trim(Session("Userno")) & "  "
		SQLBuildSQL = SQLBuildSQL & "SET zProspectFilter_" & trim(Session("Userno")) & ".NextActivityDueDate = PR_ProspectActivities.ActivityDueDate "
		SQLBuildSQL = SQLBuildSQL & "FROM zProspectFilter_" & trim(Session("Userno"))& " INNER JOIN "
		SQLBuildSQL = SQLBuildSQL & "PR_ProspectActivities ON zProspectFilter_" & trim(Session("Userno")) & ".InternalRecordIdentifier = PR_ProspectActivities.ProspectRecID "
		'SQLBuildSQL = SQLBuildSQL & "WHERE PR_ProspectActivities.Status IS NULL"
		'Response.Write(SQLBuildSQL & "<BR>")
		Set rsBuildSQL = cnnBuildSQL.Execute(SQLBuildSQL)
		ActivityDatePopulated = True
		StartDateRange = Year(ReportSpecificData21) 
		If Month(ReportSpecificData21) < 10 Then StartDateRange = StartDateRange & "0"
		StartDateRange = StartDateRange & Month(ReportSpecificData21)
		If Day(ReportSpecificData21) < 10 Then StartDateRange = StartDateRange & "0"
		StartDateRange = StartDateRange & Day(ReportSpecificData21)
		EndDateRange = Year(ReportSpecificData22) 
		If Month(ReportSpecificData22) < 10 Then EndDateRange = EndDateRange & "0"
		EndDateRange = EndDateRange & Month(ReportSpecificData22)
		If Day(ReportSpecificData22) < 10 Then EndDateRange = EndDateRange & "0"
		EndDateRange = EndDateRange & Day(ReportSpecificData22)
		SQLBuildSQL = "DELETE FROM zProspectFilter_" & trim(Session("Userno")) & " WHERE Cast(NextActivityDueDate as date) NOT BETWEEN '" & StartDateRange & "' AND '" & EndDateRange  & "'"
		'Response.Write(SQLBuildSQL & "<BR>")
		Set rsBuildSQL = cnnBuildSQL.Execute(SQLBuildSQL)
	End If
End IF

' ReportSpecificData19 - Activities to filter
If Not IsNull(ReportSpecificData19) Then
	If ReportSpecificData19 <> ""  Then
		SQLBuildSQL = "ALTER TABLE zProspectFilter_" & trim(Session("Userno")) & " ADD NextActivityNumber int NULL"
		'Response.Write(SQLBuildSQL & "<BR>")
		Set rsBuildSQL = cnnBuildSQL.Execute(SQLBuildSQL)
		' Now populate it
		SQLBuildSQL = "UPDATE zProspectFilter_" & trim(Session("Userno")) & "  "
		SQLBuildSQL = SQLBuildSQL & "SET zProspectFilter_" & trim(Session("Userno")) & ".NextActivityNumber = PR_ProspectActivities.ActivityRecID "
		SQLBuildSQL = SQLBuildSQL & "FROM zProspectFilter_" & trim(Session("Userno"))& " INNER JOIN "
		SQLBuildSQL = SQLBuildSQL & "PR_ProspectActivities ON zProspectFilter_" & trim(Session("Userno")) & ".InternalRecordIdentifier = PR_ProspectActivities.ProspectRecID "
		SQLBuildSQL = SQLBuildSQL & "WHERE PR_ProspectActivities.Status IS NULL"
		'Response.Write(SQLBuildSQL & "<BR>")
		Set rsBuildSQL = cnnBuildSQL.Execute(SQLBuildSQL)
		SQLBuildSQL = "DELETE FROM zProspectFilter_" & trim(Session("Userno"))
		SQLBuildSQL = SQLBuildSQL & " WHERE NextActivityNumber NOT IN (" & ReportSpecificData19 & ")"
		'Response.Write(SQLBuildSQL & "<BR>")
		Set rsBuildSQL = cnnBuildSQL.Execute(SQLBuildSQL)
	End If
End IF



'***************************
' Not specific to any filter
' Just needs to be here
'***************************
If ActivityDatePopulated = False Then ' Must populate before exiting
		' Now populate it
		SQLBuildSQL = "UPDATE zProspectFilter_" & trim(Session("Userno")) & "  "
		SQLBuildSQL = SQLBuildSQL & "SET zProspectFilter_" & trim(Session("Userno")) & ".NextActivityDueDate = PR_ProspectActivities.ActivityDueDate "
		SQLBuildSQL = SQLBuildSQL & "FROM zProspectFilter_" & trim(Session("Userno"))& " INNER JOIN "
		SQLBuildSQL = SQLBuildSQL & "PR_ProspectActivities ON zProspectFilter_" & trim(Session("Userno")) & ".InternalRecordIdentifier = PR_ProspectActivities.ProspectRecID "
		'SQLBuildSQL = SQLBuildSQL & "WHERE PR_ProspectActivities.Status IS NULL"
		'Response.Write(SQLBuildSQL & "<BR>")
		Set rsBuildSQL = cnnBuildSQL.Execute(SQLBuildSQL)
End If


	
Set rsBuildSQL = Nothing
cnnBuildSQL.Close
Set cnnBuildSQL= Nothing


'Response.Write("DONE")

Function GetNumEmployeeStartByEmployeeRangeNo(passedEmployeeRangeNo)

	resultGetNumEmployeeStartByIntRecID = 0

	Set cnnGetNumEmployeeStartByIntRecID = Server.CreateObject("ADODB.Connection")
	cnnGetNumEmployeeStartByIntRecID.open Session("ClientCnnString")
		
	SQLGetNumEmployeeStartByIntRecID = "Select * from PR_EmployeeRangeTable Where InternalRecordIdentifier = " & passedEmployeeRangeNo

	Set rsGetNumEmployeeStartByIntRecID = Server.CreateObject("ADODB.Recordset")
	rsGetNumEmployeeStartByIntRecID.CursorLocation = 3 
	Set rsGetNumEmployeeStartByIntRecID = cnnGetNumEmployeeStartByIntRecID.Execute(SQLGetNumEmployeeStartByIntRecID)
			 
	If not rsGetNumEmployeeStartByIntRecID.EOF Then resultGetNumEmployeeStartByIntRecID =  Left(rsGetNumEmployeeStartByIntRecID("Range"),InStr(rsGetNumEmployeeStartByIntRecID("Range"),"-")-1)

	
	rsGetNumEmployeeStartByIntRecID.Close
	set rsGetNumEmployeeStartByIntRecID= Nothing
	cnnGetNumEmployeeStartByIntRecID.Close	
	set cnnGetNumEmployeeStartByIntRecID= Nothing
	
	GetNumEmployeeStartByEmployeeRangeNo = resultGetNumEmployeeStartByIntRecID
	
End Function


Function GetNumEmployeeEndByEmployeeRangeNo(passedEmployeeRangeNo)

	resultGetNumEmployeeEndByIntRecID = 0

	Set cnnGetNumEmployeeEndByIntRecID = Server.CreateObject("ADODB.Connection")
	cnnGetNumEmployeeEndByIntRecID.open Session("ClientCnnString")
		
	SQLGetNumEmployeeEndByIntRecID = "Select * from PR_EmployeeRangeTable Where InternalRecordIdentifier = " & passedEmployeeRangeNo
 
	Set rsGetNumEmployeeEndByIntRecID = Server.CreateObject("ADODB.Recordset")
	rsGetNumEmployeeEndByIntRecID.CursorLocation = 3 
	Set rsGetNumEmployeeEndByIntRecID = cnnGetNumEmployeeEndByIntRecID.Execute(SQLGetNumEmployeeEndByIntRecID)
			 
	If not rsGetNumEmployeeEndByIntRecID.EOF Then resultGetNumEmployeeEndByIntRecID =  Right(rsGetNumEmployeeEndByIntRecID("Range"),Len(rsGetNumEmployeeEndByIntRecID("Range"))-InStr(rsGetNumEmployeeEndByIntRecID("Range"),"-"))
	
	rsGetNumEmployeeEndByIntRecID.Close
	set rsGetNumEmployeeEndByIntRecID= Nothing
	cnnGetNumEmployeeEndByIntRecID.Close	
	set cnnGetNumEmployeeEndByIntRecID= Nothing
	
	GetNumEmployeeEndByEmployeeRangeNo = resultGetNumEmployeeEndByIntRecID
	
End Function
%>