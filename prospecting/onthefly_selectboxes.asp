<!--#include file="../inc/subsandfuncs.asp"-->
<!--#include file="../inc/InSightFuncs.asp"-->
<!--#include file="../inc/InSightFuncs_Prospecting.asp"-->

<%
section= Request("section")
action= Request("action")
selectedvalue = Request("selectedvalue")

If IsEmpty(selectedvalue) OR IsNull(selectedvalue) OR Not IsNumeric(selectedvalue) Then
	selectedvalue = -1
Else
	selectedvalue = Clng(selectedvalue)
End If

UserCanAdd = False

If action="add" Then
	If userCanEditCRMOnTheFly(Session("UserNO")) = True Then
		UserCanAdd =  True
	End If
Else
	If userCanEditCRMOnTheFly(Session("UserNO")) = True AND GetCRMAddEditMenuPermissionLevel(Session("UserNO")) = vbTrue Then
		UserCanAdd =  True
	End If
End If

If section = "txtLeadSource" Then
									SQL9 = "SELECT * FROM PR_LeadSources ORDER BY LeadSource"

									Set cnn9 = Server.CreateObject("ADODB.Connection")
									cnn9.open (Session("ClientCnnString"))
									Set rs9 = Server.CreateObject("ADODB.Recordset")
									rs9.CursorLocation = 3 
									Set rs9 = cnn9.Execute(SQL9)

									If not rs9.EOF Then
									%>
                                    <option value="">Select Lead Source</option>
                                    <%If UserCanAdd Then%>
                            		<option value="-1" style="font-weight:bold"> -- Add a new Lead Source -- </option>
                                    <%End If%>
									<%
										Do																			
											%><option value="<%= rs9("InternalRecordIdentifier") %>" <%If selectedvalue=rs9("InternalRecordIdentifier") Then Response.Write("selected") End If%>><%= rs9("LeadSource") %></option><%
											rs9.movenext
										Loop until rs9.eof
									End If
									set rs9 = Nothing
									cnn9.close
									set cnn9 = Nothing
End If			

If section = "txtIndustry" Then

			  	  			'Get all industries
					      	  	SQL9 = "SELECT * FROM PR_Industries ORDER BY Industry "
			
								Set cnn9 = Server.CreateObject("ADODB.Connection")
								cnn9.open (Session("ClientCnnString"))
								Set rs9 = Server.CreateObject("ADODB.Recordset")
								rs9.CursorLocation = 3 
								Set rs9 = cnn9.Execute(SQL9)
									
								If not rs9.EOF Then
								%>
                                    <option value="">Select Industry</option>
                                    <%If UserCanAdd Then%>
                            		<option value="-1" style="font-weight:bold"> -- Add a new Industry -- </option>
                                    <%End If%>
									<%
									Do
										IndustryNumber = rs9("InternalRecordIdentifier")

										If IndustryNumber = 0 Then
											%><option value="<%= rs9("InternalRecordIdentifier") %>"  <%If selectedvalue=rs9("InternalRecordIdentifier") Then Response.Write("selected") End If%>>-- Not Specified --</option><%
										Else
											%><option value="<%= rs9("InternalRecordIdentifier") %>"  <%If selectedvalue=rs9("InternalRecordIdentifier") Then Response.Write("selected") End If%>><%= rs9("Industry") %></option><%
										End If
										rs9.movenext
									Loop until rs9.eof
								End If
								set rs9 = Nothing
								cnn9.close
								set cnn9 = Nothing
																	
End If		

If section = "txtNumEmployees" Then

				  	  			'Get employee ranges
									SQL9 = "SELECT *, Cast(LEFT(Range,CHARINDEX('-',Range)-1) as int) as Expr1 FROM PR_EmployeeRangeTable "
									SQL9 = SQL9 & "order by Expr1"

									Set cnn9 = Server.CreateObject("ADODB.Connection")
									cnn9.open (Session("ClientCnnString"))
									Set rs9 = Server.CreateObject("ADODB.Recordset")
									rs9.CursorLocation = 3 
									Set rs9 = cnn9.Execute(SQL9)
										
									If not rs9.EOF Then
									%>
                                    <option value="">Select # Employees</option>
                                    <%If UserCanAdd Then%>
                            		<option value="-1" style="font-weight:bold"> -- Add a new Employee Range -- </option>
                                    <%End If%>
									<%
										Do
											%><option value="<%= rs9("InternalRecordIdentifier") %>"  <%If selectedvalue=rs9("InternalRecordIdentifier") Then Response.Write("selected") End If%>><%= rs9("Range") %></option><%
											rs9.movenext
										Loop until rs9.eof
									End If
									set rs9 = Nothing
									cnn9.close
									set cnn9 = Nothing
																	
End If	

If section = "txtPrimaryCompetitor" Then

SQLCompetitorNames = "SELECT * FROM PR_Competitors ORDER BY CompetitorName"
							Set cnnCompetitorNames = Server.CreateObject("ADODB.Connection")
							cnnCompetitorNames.open (Session("ClientCnnString"))
							Set rsCompetitorNames = Server.CreateObject("ADODB.Recordset")
							rsCompetitorNames.CursorLocation = 3 
							Set rsCompetitorNames = cnnCompetitorNames.Execute(SQLCompetitorNames)
							
							If not rsCompetitorNames.EOF Then
							%>
                                    <option value="">Select Primary Competitor</option>
                                    <%If UserCanAdd Then%>
                            		<option value="-1" style="font-weight:bold"> -- Add a new Primary Competitor -- </option>
                                    <%End if%>
									<%
								sep = ""
								Do While Not rsCompetitorNames.EOF
										%><option value="<%= rsCompetitorNames("InternalRecordIdentifier") %>"  <%If selectedvalue=rsCompetitorNames("InternalRecordIdentifier") Then Response.Write("selected") End If%>><%= rsCompetitorNames("CompetitorName") %></option><%
									rsCompetitorNames.MoveNext						
								Loop
							End If
							Set rsCompetitorNames = Nothing
							cnnCompetitorNames.Close
							Set cnnCompetitorNames = Nothing
																	
End If	

If section = "txtTitle" Then

SQLContactTitles = "SELECT *, InternalRecordIdentifier as id FROM PR_ContactTitles ORDER BY ContactTitle"
								Set cnnContactTitles = Server.CreateObject("ADODB.Connection")
								cnnContactTitles.open (Session("ClientCnnString"))
								Set rsContactTitles = Server.CreateObject("ADODB.Recordset")
								rsContactTitles.CursorLocation = 3 
								Set rsContactTitles = cnnContactTitles.Execute(SQLContactTitles)
								%>
                                    <option value="">Select Job Title</option>
                                    <%If UserCanAdd Then%>
                            		<option value="-1" style="font-weight:bold"> -- Add a new Job Title -- </option>
                                    <%End If%>
									<%
								If not rsContactTitles.EOF Then
								
									Do While Not rsContactTitles.EOF
											%><option value="<%= rsContactTitles("id") %>"  <%If selectedvalue=rsContactTitles("InternalRecordIdentifier") Then Response.Write("selected") End If%>><%= rsContactTitles("ContactTitle") %></option><%
										rsContactTitles.MoveNext						
									Loop
								End If
								Set rsContactTitles = Nothing
								cnnContactTitles.Close
								Set cnnContactTitles = Nothing
								
																	
End If	

If section = "txtTitleforTab" Then
Set cnnContactTitles = Server.CreateObject("ADODB.Connection")
cnnContactTitles.open (Session("ClientCnnString"))
								
SQLContactTitles = "SELECT *, InternalRecordIdentifier as id FROM PR_ContactTitles ORDER BY ContactTitle"

Set rsContactTitles = Server.CreateObject("ADODB.Recordset")
rsContactTitles.CursorLocation = 3 
Set rsContactTitles = cnnContactTitles.Execute(SQLContactTitles)

'ContactTitles = ("[")
ContactTitles = ("[{""id"":""0"",""title"":""Select""},")
If userCanEditCRMOnTheFly(Session("UserNO")) = True AND GetCRMAddEditMenuPermissionLevel(Session("UserNO")) = vbTrue Then
ContactTitles  = ContactTitles & ("{""id"":""-1"",""title"":""Add a new Job Title""},")
End If
If not rsContactTitles.EOF Then
	sep = ""
	Do While Not rsContactTitles.EOF
			ContactTitles = ContactTitles & (sep)
			sep = ","
			ContactTitles = ContactTitles & ("{")
			ContactTitles = ContactTitles & ("""id"":""" & Replace(rsContactTitles("id"), """", "\""") & """")
			ContactTitles = ContactTitles & (",""title"":""" & Replace(rsContactTitles("ContactTitle"), """", "\""") & """")
			ContactTitles = ContactTitles & ("}")
		rsContactTitles.MoveNext						
	Loop
End If
ContactTitles = ContactTitles & ("]")
Set rsContactTitles = Nothing

cnnContactTitles.Close
Set cnnContactTitles = Nothing
								
response.Write(ContactTitles)								
																	
End If	

If section = "selProspectNextActivity" Then
	Set cnnNextActivity = Server.CreateObject("ADODB.Connection")
	cnnNextActivity.open (Session("ClientCnnString"))
	
	SQLNextActivity = "SELECT * FROM PR_Activities ORDER BY Activity"	
	Set rsNextActivity = Server.CreateObject("ADODB.Recordset")
	rsNextActivity.CursorLocation = 3 
	Set rsNextActivity = cnnNextActivity.Execute(SQLNextActivity)
	
								%>
                                    <option value="">Select next activity</option>
                                    <%If UserCanAdd Then%>
                            		<option value="-1" style="font-weight:bold"> -- Add a new next activity -- </option>
                                    <%End If%>
									<%
								If not rsNextActivity.EOF Then
								
									Do While Not rsNextActivity.EOF
											%><option value="<%= rsNextActivity("InternalRecordIdentifier") %>"  <%If selectedvalue=rsNextActivity("InternalRecordIdentifier") Then Response.Write("selected") End If%>><%= rsNextActivity("Activity") %></option><%
										rsNextActivity.MoveNext						
									Loop
								End If
								set rsNextActivity = Nothing
								cnnNextActivity.close
								set cnnNextActivity = Nothing					
																		
																	
End If	
					
								%>