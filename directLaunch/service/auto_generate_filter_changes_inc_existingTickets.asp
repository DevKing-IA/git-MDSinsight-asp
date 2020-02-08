<%
					'The code in this include files looks through existing service tickets & generates filter
					'change tickets where appropriate

					NumberOfTicketsGeneratedSoFar = 0
							
					Set cnnFilterChanges = Server.CreateObject("ADODB.Connection")
					cnnFilterChanges.open (Session("ClientCnnString"))
					Set rsFilterChanges = Server.CreateObject("ADODB.Recordset")
					Set rsFilterChangesForUpdating = Server.CreateObject("ADODB.Recordset")
					Set rsServiceTickets = Server.CreateObject("ADODB.Recordset")
		
					ReDim CustIDsToCheckArray(1)
					ReDim CustIDsToGenerateArray(1) ' This is the final one where they end up
					
					
					'OK, go to the next step
					If AutoFilterChangeGenerationONOFF = 1 Then
					
						'To start with, look at all service tickets which have a status of Awaiting Dispatch,(REDOs) Followup, Swap, Wait For Parts, Unable To Work
						'and if there are any filter changes due for those customers, generate the filter change tickets
						
						SQLServiceTickets = "SELECT * FROM FS_ServiceMemos WHERE CurrentStatus = 'OPEN'"
						
						Set rsServiceTickets = cnnFilterChanges.Execute(SQLServiceTickets)
						
						If NOT rsServiceTickets.EOF Then
						
							ElementNumber = 1
							
							Do While Not rsServiceTickets.EOF
							
								'Check the ticket stage
								GetServiceTicketCurrentStageVar = GetServiceTicketCurrentStage(rsServiceTickets("MemoNumber"))
								
								If rsServiceTickets.Fields("RecordSubType") <> "HOLD"_
									AND (GetServiceTicketCurrentStageVar = "Received"_
									OR GetServiceTicketCurrentStageVar = "Released"_
									OR GetServiceTicketCurrentStageVar = "Declined"_
									OR GetServiceTicketCurrentStageVar = "Unable To Work"_
									OR GetServiceTicketCurrentStageVar = "Swap"_
									OR GetServiceTicketCurrentStageVar = "Wait for parts"_
									OR GetServiceTicketCurrentStageVar = "Follow Up") Then
									
									'Add the customer id to the Custs to check Array
									CustIDsToCheckArray(ElementNumber-1) = rsServiceTickets("AccountNumber")
									ReDim PRESERVE CustIDsToCheckArray(ElementNumber +1) ' add room for another one
									ElementNumber = ElementNumber + 1
								End If

								rsServiceTickets.movenext
							Loop														
						
							msg = "Customer to check from service tickets: "
							For x = 0 to UBOUND(CustIDsToCheckArray)
								msg = msg & CustIDsToCheckArray(x) & "," 		
							Next
							
							WriteResponse msg
								
						End If

						'Now we have the list of customers to check, let's see who is due for a filter change and winnow the list to only valid customers
						ElementNumber = 1
						
						For x = 0 to UBOUND(CustIDsToCheckArray)
		
							If CustHasPendingFilterChange(CustIDsToCheckArray(x)) <> True Then 
								WriteResponse "Removing " & CustIDsToCheckArray(x) & "<br>"
								CustIDsToCheckArray(x)="" ' Just blank out the customer in the array
							Else
								WriteResponse "Leaving " & CustIDsToCheckArray(x) & "<br>"
								'Put this one into the final array
								CustIDsToGenerateArray(ElementNumber-1) = CustIDsToCheckArray(x)
								ReDim PRESERVE CustIDsToGenerateArray(ElementNumber +1) ' add room for another one
								ElementNumber = ElementNumber + 1
							End If
							
						Next
						
						
						
						''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
						''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
						''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''												
						' This is the final step, anything left in the CustIDsToGenerateArray at this point gets the filter change generated
						''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''						
						''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''						
						''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''						
						'Step 1, scrub the duplicate CudtIDs out of the array
						ReDim tmpArray(UBOUND(CustIDsToGenerateArray))
						
						tmpArrayElement = 0
						
						For x = 0 to UBOUND(CustIDsToGenerateArray) 
							
							ExistsInArray = False
							
							For z = 0 to UBOUND(tmpArray)
							
								If tmpArray(z) = CustIDsToGenerateArray(x) Then ExistsInArray = True
								
							Next
							
							If ExistsInArray <> True Then 
								If NOT IsNull(CustIDsToGenerateArray(x)) Then ' strip the empty ones too
									tmpArray(tmpArrayElement) = CustIDsToGenerateArray(x)
									tmpArrayElement = tmpArrayElement + 1
								End If
							End If
							
						Next
						
						If tmpArrayElement = 0 Then ReDim tmpArray(-1) ' Basically destry the array because there is no data
						
						
						ReDim CustIDsToGenerateArray(UBOUND(tmpArray))
						
						
						For x = 0 to UBOUND(tmpArray)
							CustIDsToGenerateArray(x) = tmpArray(x)
						Next

						
						msg = "Final array:"
						msg = msg & "   UBOUND:" & UBOUND(CustIDsToGenerateArray) & ":"
						For x = 0 to UBOUND(CustIDsToGenerateArray) 
							msg = msg & CustIDsToGenerateArray (x) & "," 		
						Next
						msg = msg & "<br>"
						WriteResponse msg
						
						If Ubound(CustIDsToGenerateArray) >= 0 Then 
						
							'Now actually generate the filter change tickets we have to do 1 post only which includes all filters due
							
							SQLFilterChanges = "SELECT * FROM FS_CustomerFilters WHERE CustID IN ("
							For x = 0 to UBOUND(CustIDsToGenerateArray) 
								If CustIDsToGenerateArray(x) <> "" Then SQLFilterChanges = SQLFilterChanges & "'" & CustIDsToGenerateArray(x) & "'," 		
							Next
							
							'Strip trainling comma
							SQLFilterChanges = Left(SQLFilterChanges,LEN(SQLFilterChanges) -1)
							SQLFilterChanges = SQLFilterChanges & ") ORDER BY CustID"
	
							Response.Write("<br>SQLFilterChanges " & SQLFilterChanges& "<br>")						
							
							Set rsFilterChanges = cnnFilterChanges.Execute(SQLFilterChanges)
						
							FiltersToDo = ""
							HeldCustID = ""
							filters = ""
							
							If NOT rsFilterChanges.EOF Then
							
								Do While Not rsFilterChanges.EOF
								
									If HeldCustID = "" Then	HeldCustID = rsFilterChanges("CustID")
									
									If HeldCustID <> rsFilterChanges("CustID") Then ' Customer changed, submit the filters
									
											If Right(FiltersToDo,1) = "," Then FiltersToDo = Left(FiltersToDo,Len(FiltersToDo)-1)
					
											Response.Write("<br>FiltersToDo " & FiltersToDo & "<br>")
											
											NumberOfTicketsGeneratedSoFar  = NumberOfTicketsGeneratedSoFar + 1
											
											Response.Write("<br>NumberOfTicketsGeneratedSoFar  " & NumberOfTicketsGeneratedSoFar  & "<br>")
											
											Response.Write("<br>Call SubmitTicket(" & HeldCustID & "," & FiltersToDo & "," & filters &  "<br>")
											
											Call SubmitTicket(HeldCustID ,FiltersToDo,filters)
									
											HeldCustID = rsFilterChanges("CustID")
											
											FiltersToDo = "" : filters= "" ' Re init
									End If
									
									FiltersToDo = FiltersToDo & rsFilterChanges("InternalRecordIdentifier") & ","
	
									filters = filters & vbcrlf & "Filter: " & GetFilterIDByIntRecID(rsFilterChanges("FilterIntRecID")) & " - " & GetFilterDescByIntRecID(rsFilterChanges("FilterIntRecID"))
							
									rsFilterChanges.Movenext
									
									If rsFilterChanges.EOF Then ' to make sure we catch the last one
									
											If Right(FiltersToDo,1) = "," Then FiltersToDo = Left(FiltersToDo,Len(FiltersToDo)-1)
					
											Response.Write("<br>FiltersToDo " & FiltersToDo & "<br>")
											
											NumberOfTicketsGeneratedSoFar  = NumberOfTicketsGeneratedSoFar + 1
											
											Response.Write("<br>NumberOfTicketsGeneratedSoFar  " & NumberOfTicketsGeneratedSoFar  & "<br>")
											
											Response.Write("<br>Call SubmitTicket(" & HeldCustID & "," & FiltersToDo & "," & filters &  "<br>")
											
											Call SubmitTicket(HeldCustID ,FiltersToDo,filters)
											
											FiltersToDo = "" : filters= "" ' Re init

									End If
									
								Loop  
								
							End If
							
							
						End If
%>