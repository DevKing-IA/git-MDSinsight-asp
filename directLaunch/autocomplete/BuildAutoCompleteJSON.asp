<!--#include file="../../inc/SubsAndFuncs.asp"-->
<!--#include file="../../inc/InsightFuncs.asp"-->
<script type="text/javascript">
	function closeme() {
		window.open('', '_parent', '');
		window.close();  }
</script>
<%
Server.ScriptTimeout = 2500
'Delivery Board Alert processing page
'Designed to be launched via a scheduled process (Win Task Scheduler)
'Self contained page will check the alerts db and take the appropriate actions
'Usage = "http://{xxx}.{domain}.com/directLaunch/autocomplete/BuildAutoCompleteJSON.asp?runlevel=run_now

'The runlevel parameter is inconsequential to the operation 
'of the page. It is only used so that the page will not run
'if it is loaded via an unexpected method (spiders, etc)
If Request.QueryString("runlevel") <> "run_now" then response.end

'baseURL should alwats have a trailing /slash, just in case, handle either way
If right(baseURL,1)="/" Then maildomain = Left(right(baseURL,len(baseURL)-7),len(right(baseURL,len(baseURL)-7))-1) Else maildomain = right(baseURL,len(baseURL)-7)

'This single page loops through and handles autocompletes for all databases
SQL = "SELECT * FROM tblServerInfo WHERE Active = 1 and ClientKey='1106d'"

Set TopConnection = Server.CreateObject("ADODB.Connection")
Set TopRecordset = Server.CreateObject("ADODB.Recordset")
TopConnection.Open InsightCnnString
'Open the recordset object executing the SQL statement and return records
TopRecordset.Open SQL,TopConnection,3,3

'First lookup the ClientKey in tblServerInfo
'If there is no record with the entered client key, close connection
'and exit
If Not TopRecordset.Eof Then

	Do While Not TopRecordset.EOF

		PassPhrase = TopRecordset.Fields("directLaunchPassphrase")
		ClientKey = TopRecordset.Fields("clientkey")

		Response.Write("******** Processing " & ClientKey  & "************<br>")
		Call SetClientCnnString
		
		Session("ClientCnnString") = MUV_READ("ClientCnnString") ' Until session vars are gone, then delete this
		
		If MUV_READ("cnnStatus") = "OK" Then ' else it loops and excludes dev client keys

			serverName = Request.ServerVariables("SERVER_NAME")
			
			If serverName = "www.mdsinsight.com" Then serverName = "mdsinsight.com"
			
			Response.Write("serverName: " & serverName & "<br>")
			
			
			If serverName <> "mdsinsight.com" OR (serverName = "mdsinsight.com" AND UCASE(RIGHT(ClientKey,1)) <> "D") Then 
				
					'Each autocomplete is handled individually as every customer has different accounts
					
					'****************************************
					'Begin Build Autocomplete JSON Files
					'****************************************
					 Response.Write("Begin Build Auto Complete JSON<br>")
					 
					 'Response.write(Server.MapPath(".") & "<br><br>")
					
					'******************************************
		
		
					'******************************************************************
					' Begin Auto Complete Customer For CSZ/Non-CSZ, NO POS
					'******************************************************************
					
						SQLAutoComplete = "SELECT custNum,Name,CityStateZip FROM " & MUV_Read("SQL_Owner")  & ".AR_Customer WHERE AcctStatus='A' ORDER BY CustNum"
						
						Response.Write(SQLAutoComplete & "<br>")
						
						Set cnnAutoComplete = Server.CreateObject("ADODB.Connection")
						cnnAutoComplete.open (Session("ClientCnnString"))
						Set rsAutoComplete = Server.CreateObject("ADODB.Recordset")
						rsAutoComplete.CursorLocation = 3 
						
					
						Set rsAutoComplete = cnnAutoComplete.Execute(SQLAutoComplete)
					
						If not rsAutoComplete.EOF Then
						
							CustomerCount = 0
							jsonDataCSZ = ""
							jsonData = ""
							
							Do While Not rsAutoComplete.EOF
							
								CustomerCount = CustomerCount + 1
								
								If CustomerCount = 1 Then
									jsonDataCSZ = "[{""name"":""" & rsAutoComplete("custNum") &" --- "& rsAutoComplete("Name") &" --- "& rsAutoComplete("CityStateZip") & """, ""code"": """ & rsAutoComplete("custNum") & """},"
								Else
									jsonDataCSZ = jsonDataCSZ & "{""name"":""" & rsAutoComplete("custNum") &" --- "& rsAutoComplete("Name") &" --- "& rsAutoComplete("CityStateZip") & """, ""code"": """ & rsAutoComplete("custNum") & """},"
								End If
		
								If CustomerCount = 1 Then
									jsonData = "[{""name"":""" & rsAutoComplete("custNum") &" --- "& rsAutoComplete("Name") & """, ""code"": """ & rsAutoComplete("custNum") & """},"
								Else
									jsonData = jsonData & "{""name"":""" & rsAutoComplete("custNum") &" --- "& rsAutoComplete("Name") & """, ""code"": """ & rsAutoComplete("custNum") & """},"
								End If
														
								rsAutoComplete.MoveNext
								
							Loop
							
							Session("strAccounts") = "GENERATE COMPLETE"
							
							If Len(jsonDataCSZ)>0 Then jsonDataCSZ = Left(jsonDataCSZ,Len(jsonDataCSZ)-1)
							jsonDataCSZ = jsonDataCSZ & "]"
							
							If Len(jsonData)>0 Then jsonData = Left(jsonData,Len(jsonData)-1)
							jsonData = jsonData & "]"
							
							ClientKeyForFileName = ClientKey
						
							set fs=Server.CreateObject("Scripting.FileSystemObject")
							set fs2=Server.CreateObject("Scripting.FileSystemObject")
							
							Response.Write(Server.MapPath("..\..\") & "\clientfiles\"  & ClientKeyForFileName & "\autocomplete\customer_account_list_CSZ_" & ClientKeyForFileName & ".json")
							set tfile=fs.CreateTextFile(Server.MapPath("..\..\") & "\clientfiles\"  & ClientKeyForFileName & "\autocomplete\customer_account_list_CSZ_" & ClientKeyForFileName & ".json")
							tfile.WriteLine(jsonDataCSZ)
							tfile.close
							set tfile=nothing
							set fs=nothing
							
							
							set fs2=Server.CreateObject("Scripting.FileSystemObject")
							set tfile2=fs2.CreateTextFile(Server.MapPath("..\..\") & "\clientfiles\"  & ClientKeyForFileName & "\autocomplete\customer_account_list_" & ClientKeyForFileName & ".json")
							tfile2.WriteLine(jsondata)
							tfile2.close
							set tfile2=nothing
							set fs2=nothing
						
						End If
							
						Set rsAutoComplete = Nothing
						cnnAutoComplete.Close
						Set AutoComplete = nothing
		
		
					'******************************************************************
					' Begin Auto Complete Customer For CSZ/Non-CSZ, WITH POS
					'******************************************************************
					
					SQLAutoCompletePOSCheck = "SELECT PointOfServiceLogicOnOff FROM " & MUV_Read("SQL_Owner")  & ".Settings_CompanyID"
					
					Response.Write(SQLAutoCompletePOSCheck  & "<br>")
					
					Set cnnAutoCompletePOSCheck = Server.CreateObject("ADODB.Connection")
					cnnAutoCompletePOSCheck.open (Session("ClientCnnString"))
					Set rsAutoCompletePOSCheck = Server.CreateObject("ADODB.Recordset")
					rsAutoCompletePOSCheck.CursorLocation = 3 
					
				
					Set rsAutoCompletePOSCheck = cnnAutoCompletePOSCheck.Execute(SQLAutoCompletePOSCheck)
				
					If NOT rsAutoCompletePOSCheck.EOF Then
					
						PointOfServiceLogicOnOff = rsAutoCompletePOSCheck("PointOfServiceLogicOnOff")
						
						If PointOfServiceLogicOnOff = 1 Then
											
							SQLAutoComplete = "SELECT * FROM AR_CustomerPOS "
							SQLAutoComplete = SQLAutoComplete & " INNER JOIN AR_CustomerShipTo ON AR_CustomerPOS.CustID = AR_CustomerShipTo.CustNum AND "
							SQLAutoComplete = SQLAutoComplete & " AR_CustomerPOS.BackendShipToIDIfApplicable = AR_CustomerShipTo.BackendShipToIDIfApplicable  "
							SQLAutoComplete = SQLAutoComplete & " INNER JOIN AR_Customer ON AR_CustomerPOS.CustID = AR_Customer.CustNum "
							SQLAutoComplete = SQLAutoComplete & " WHERE AR_Customer.AcctStatus='A' "
							SQLAutoComplete = SQLAutoComplete & " ORDER BY AR_CustomerPOS.CustID, AR_CustomerPOS.BackendShipToIDIfApplicable, AR_CustomerPOS.POSID, AR_CustomerShipTo.ShipName"
							
							Response.Write(SQLAutoComplete & "<br>")
							
							Set cnnAutoComplete = Server.CreateObject("ADODB.Connection")
							cnnAutoComplete.open (Session("ClientCnnString"))
							Set rsAutoComplete = Server.CreateObject("ADODB.Recordset")
							rsAutoComplete.CursorLocation = 3 
							
						
							Set rsAutoComplete = cnnAutoComplete.Execute(SQLAutoComplete)
						
							If not rsAutoComplete.EOF Then
							
								CustomerCount = 0
								jsonDataCSZ = ""
								jsonData = ""
								jsonDataCSZPOS = ""
								
								Do While Not rsAutoComplete.EOF
								
									CustomerCount = CustomerCount + 1
									
									If CustomerCount = 1 Then
										jsonDataCSZ = "[{""name"":""" & rsAutoComplete("custNum") &" --- "& rsAutoComplete("Name") &" --- "& rsAutoComplete("CityStateZip") & """, ""code"": """ & rsAutoComplete("custNum") & """},"
									Else
										jsonDataCSZ = jsonDataCSZ & "{""name"":""" & rsAutoComplete("custNum") &" --- "& rsAutoComplete("Name") &" --- "& rsAutoComplete("CityStateZip") & """, ""code"": """ & rsAutoComplete("custNum") & """},"
									End If
			
									If CustomerCount = 1 Then
										jsonData = "[{""name"":""" & rsAutoComplete("custNum") &" --- "& rsAutoComplete("Name") & """, ""code"": """ & rsAutoComplete("custNum") & """},"
									Else
										jsonData = jsonData & "{""name"":""" & rsAutoComplete("custNum") &" --- "& rsAutoComplete("Name") & """, ""code"": """ & rsAutoComplete("custNum") & """},"
									End If
													
									If CustomerCount = 1 Then
										jsonDataCSZPOS = "[{""name"":""" & rsAutoComplete("Name")
										jsonDataCSZPOS = jsonDataCSZPOS & " --- " & rsAutoComplete("ShipName") 
										jsonDataCSZPOS = jsonDataCSZPOS & " --- " & rsAutoComplete("POSName") & " --- " & rsAutoComplete("CustNum") & "-" & rsAutoComplete("BackendShipToIDIfApplicable") & "-" & rsAutoComplete("POSID")
										jsonDataCSZPOS = jsonDataCSZPOS & """, ""code"": """ & rsAutoComplete("custNum") & """},"
									Else
										jsonDataCSZPOS = jsonDataCSZPOS & "{""name"":""" & rsAutoComplete("Name") 
										jsonDataCSZPOS = jsonDataCSZPOS & " --- " & rsAutoComplete("ShipName") 
										jsonDataCSZPOS = jsonDataCSZPOS & " --- " & rsAutoComplete("POSName") & " --- " & rsAutoComplete("CustNum") & "-" & rsAutoComplete("BackendShipToIDIfApplicable") & "-" & rsAutoComplete("POSID")
										jsonDataCSZPOS = jsonDataCSZPOS & """, ""code"": """ & rsAutoComplete("custNum") & """},"
									End If
		
									rsAutoComplete.MoveNext
									
								Loop
								
								Session("strAccounts") = "GENERATE COMPLETE"

								If Len(jsonDataCSZPOS)>0 Then jsonDataCSZPOS = Left(jsonDataCSZPOS,Len(jsonDataCSZPOS)-1)
								jsonDataCSZPOS = jsonDataCSZPOS & "]"
								
								If Len(jsonDataCSZ)>0 Then jsonDataCSZ = Left(jsonDataCSZ,Len(jsonDataCSZ)-1)
								jsonDataCSZ = jsonDataCSZ & "]"
								
								If Len(jsonData)>0 Then jsonData = Left(jsonData,Len(jsonData)-1)
								jsonData = jsonData & "]"
								
								jsonData = Replace(jsonData,"'","")
								jsonDataCSZ = Replace(jsonDataCSZ,"'","")
								jsonDataCSZPOS = Replace(jsonDataCSZPOS,"'","")								

								jsonData = Replace(jsonData,"/","-")
								jsonDataCSZ = Replace(jsonDataCSZ,"/","-")
								jsonDataCSZPOS = Replace(jsonDataCSZPOS,"/","-")		
								
								jsonData = Replace(jsonData,"\","-")
								jsonDataCSZ = Replace(jsonDataCSZ,"\","-")
								jsonDataCSZPOS = Replace(jsonDataCSZPOS,"\","-")
								
								ClientKeyForFileName = ClientKey
							
								set fs=Server.CreateObject("Scripting.FileSystemObject")
								set fs2=Server.CreateObject("Scripting.FileSystemObject")
								set fs3=Server.CreateObject("Scripting.FileSystemObject")
								
								Response.Write(Server.MapPath("..\..\") & "\clientfiles\"  & ClientKeyForFileName & "\autocomplete\customer_account_list_" & ClientKeyForFileName & ".json")
								set tfile=fs.CreateTextFile(Server.MapPath("..\..\") & "\clientfiles\"  & ClientKeyForFileName & "\autocomplete\customer_account_list_" & ClientKeyForFileName & ".json")
								tfile.WriteLine(jsonData)
								tfile.close
								
								Response.Write(Server.MapPath("..\..\") & "\clientfiles\"  & ClientKeyForFileName & "\autocomplete\customer_account_list_CSZ_" & ClientKeyForFileName & ".json")
								set tfile2=fs2.CreateTextFile(Server.MapPath("..\..\") & "\clientfiles\"  & ClientKeyForFileName & "\autocomplete\customer_account_list_CSZ_" & ClientKeyForFileName & ".json")
								tfile2.WriteLine(jsonDataCSZ)
								tfile2.close
									
								Response.Write(Server.MapPath("..\..\") & "\clientfiles\"  & ClientKeyForFileName & "\autocomplete\customer_account_list_CSZ_POS_" & ClientKeyForFileName & ".json")
								set tfile3=fs3.CreateTextFile(Server.MapPath("..\..\") & "\clientfiles\"  & ClientKeyForFileName & "\autocomplete\customer_account_list_CSZ_POS_" & ClientKeyForFileName & ".json")
								tfile3.WriteLine(jsonDataCSZPOS)
								tfile3.close
								
								set tfile=nothing
								set fs=nothing

								set tfile2=nothing
								set fs2=nothing

								set tfile3=nothing
								set fs3=nothing
						
							End If
								
							Set rsAutoComplete = Nothing
							cnnAutoComplete.Close
							Set AutoComplete = nothing
						
						End If
					
					End If
					
					Set rsAutoCompletePOSCheck = Nothing
					cnnAutoCompletePOSCheck.Close
					Set AutoCompletePOSCheck = nothing
			
		Response.End
					'**************************************************
					' Begin Auto Complete Customer For Equipment Module
					'**************************************************
		
						Response.Write("Begin Auto Complete Customer For Equipment Module JSON<br>")
					 
					
						SQLAutoComplete = "SELECT custNum,Name,CityStateZip FROM " & MUV_Read("SQL_Owner")  & ".AR_Customer WHERE AcctStatus='A'"
						SQLAutoComplete = SQLAutoComplete & " AND CustNum IN (SELECT DISTINCT CustID FROM EQ_CustomerEquipment) ORDER BY CustNum"
						
						response.write(SQLAutoComplete & "<br>")
						
						Set cnnAutoComplete = Server.CreateObject("ADODB.Connection")
						cnnAutoComplete.open (Session("ClientCnnString"))
						Set rsAutoComplete = Server.CreateObject("ADODB.Recordset")
						rsAutoComplete.CursorLocation = 3 
						
					
						Set rsAutoComplete = cnnAutoComplete.Execute(SQLAutoComplete)
					
						If not rsAutoComplete.EOF Then
						
							response.write("We did not get an EOF for Equipment for CLient " & ClientKeyForFileName & "<br>")
						
							CustomerCount = 0
							jsonDataCSZ = ""
							jsonData = ""
							
							Do While Not rsAutoComplete.EOF
							
								CustomerCount = CustomerCount + 1
								
								If CustomerCount = 1 Then
									jsonDataCSZ = "[{""name"":""" & rsAutoComplete("custNum") &" --- "& rsAutoComplete("Name") &" --- "& rsAutoComplete("CityStateZip") & """, ""code"": """ & rsAutoComplete("custNum") & """},"
								Else
									jsonDataCSZ = jsonDataCSZ & "{""name"":""" & rsAutoComplete("custNum") &" --- "& rsAutoComplete("Name") &" --- "& rsAutoComplete("CityStateZip") & """, ""code"": """ & rsAutoComplete("custNum") & """},"
								End If
		
								If CustomerCount = 1 Then
									jsonData = "[{""name"":""" & rsAutoComplete("custNum") &" --- "& rsAutoComplete("Name") & """, ""code"": """ & rsAutoComplete("custNum") & """},"
								Else
									jsonData = jsonData & "{""name"":""" & rsAutoComplete("custNum") &" --- "& rsAutoComplete("Name") & """, ""code"": """ & rsAutoComplete("custNum") & """},"
								End If
														
								rsAutoComplete.MoveNext
								
							Loop
							
							Session("strAccounts") = "GENERATE COMPLETE"
							
							If Len(jsonDataCSZ)>0 Then jsonDataCSZ = Left(jsonDataCSZ,Len(jsonDataCSZ)-1)
							jsonDataCSZ = jsonDataCSZ & "]"
							
							If Len(jsonData)>0 Then jsonData = Left(jsonData,Len(jsonData)-1)
							jsonData = jsonData & "]"
							
						
							ClientKeyForFileName = ClientKey
		
						
							set fs=Server.CreateObject("Scripting.FileSystemObject")
							set fs2=Server.CreateObject("Scripting.FileSystemObject")
							
							'Response.Write(Server.MapPath("..\..\") & "\clientfiles\"  & ClientKeyForFileName & "\autocomplete\customer_account_list_CSZ_Equipment.json")
							set tfile=fs.CreateTextFile(Server.MapPath("..\..\") & "\clientfiles\"  & ClientKeyForFileName & "\autocomplete\customer_account_list_CSZ_Equipment.json")
							tfile.WriteLine(jsonDataCSZ)
							tfile.close
							set tfile=nothing
							set fs=nothing
							
							
							set fs2=Server.CreateObject("Scripting.FileSystemObject")
							set tfile2=fs2.CreateTextFile(Server.MapPath("..\..\") & "\clientfiles\"  & ClientKeyForFileName & "\autocomplete\customer_account_list_Equipment.json")
							tfile2.WriteLine(jsondata)
							tfile2.close
							set tfile2=nothing
							set fs2=nothing
							
						End If
							
						Set rsAutoComplete = Nothing
						cnnAutoComplete.Close
						Set AutoComplete = nothing
		
					'**************************************************
					' End Auto Complete Customer For Equipment Module				
					'**************************************************
						
						'*********************************************************
						' Begin Auto Complete CHAINs
						'*********************************************************
						 Response.Write("Begin Build Auto Complete JSON Chains <br>")
						'******************************************
		
						DoChains = True
						
						'Only if the backend has a chains table
						On Error Goto 0

						Set cnnAutoComplete = Server.CreateObject("ADODB.Connection")
						cnnAutoComplete.open (Session("ClientCnnString"))
						Set rsAutoComplete = Server.CreateObject("ADODB.Recordset")

						Err.Clear
						on error resume next
						Set rsAutoComplete = cnnAutoComplete.Execute("SELECT TOP 1 * FROM Chain")
						If Err.Description <> "" Then
							If InStr(Ucase(Err.Description),"INVALID OBJECT NAME") <> 0 Then
								DoChains = False
							End If
						End IF

						If DoChains = True Then
						
							Response.Write("[")
							
							SQL = "SELECT Distinct ChainNum FROM AR_Customer where ChainNum <> '' order by ChainNum"
							'response.write(SQL)
							Set cnnAutoComplete = Server.CreateObject("ADODB.Connection")
							cnnAutoComplete.open (Session("ClientCnnString"))
							Set rsAutoComplete = Server.CreateObject("ADODB.Recordset")
							rsAutoComplete.CursorLocation = 3 
							Set rsAutoComplete = cnnAutoComplete.Execute(SQL)
							
							If not rsAutoComplete.EOF Then
							strAuto = "["
							Do While Not rsAutoComplete.EOF
							            
							            strAuto = strAuto & "{""name"":""" & rsAutoComplete("ChainNum") & " --- " & GetChainDescByChainNum(rsAutoComplete("ChainNum")) & """, ""code"":""" & rsAutoComplete("ChainNum") & """},"
							
							            'Response.Write(strAuto)
							            rsAutoComplete.MoveNext
							Loop
							
							End If
							
							If right(strAuto,1)= "," Then strAuto = left(strAuto,len(strAuto)-1) 
							
							strAuto = trim(strAuto) & "]"
							
							Response.Write("]")
						
						End If ' For do Chains
						
						ClientKeyForFileName = ClientKey
					
						set fs=Server.CreateObject("Scripting.FileSystemObject")
						set fs2=Server.CreateObject("Scripting.FileSystemObject")
						
						set tfile=fs.CreateTextFile(Server.MapPath("..\..\") & "\clientfiles\"  & ClientKeyForFileName & "\autocomplete\customer_chain_list_" & ClientKeyForFileName & ".json")
						tfile.WriteLine(strAuto)
						tfile.close
						set tfile=nothing
						set fs=nothing
						
						Set rsAutoComplete = Nothing
						cnnAutoComplete.Close
						Set AutoComplete = nothing
		
						
						'*********************************************************
						' END Auto Complete CHAINs
						'*********************************************************
		
		
						'*********************************************************
						' Begin Auto Complete  Product List
						'*********************************************************
						 Response.Write("Begin Build Auto Complete JSON  Product List <br>")
						'******************************************
		
						
						Response.Write("[")
						SQL = "SELECT Distinct prodSKU,prodDescription FROM IC_Product where prodSKU <> '' order by prodSKU"
						Set cnnAutoComplete = Server.CreateObject("ADODB.Connection")
						cnnAutoComplete.open (Session("ClientCnnString"))
						Set rsAutoComplete = Server.CreateObject("ADODB.Recordset")
						rsAutoComplete.CursorLocation = 3 
						Set rsAutoComplete = cnnAutoComplete.Execute(SQL)
						
						If not rsAutoComplete.EOF Then
						strAuto = "["
						Do While Not rsAutoComplete.EOF
						    strAuto = strAuto & "{""name"":""" & rsAutoComplete("prodSKU") & " --- " & rsAutoComplete("prodDescription") & """, ""code"":""" & rsAutoComplete("prodSKU") & """},"
						    rsAutoComplete.MoveNext
						Loop
						End If
						
						If right(strAuto,1)= "," Then strAuto = left(strAuto,len(strAuto)-1) 
						
						strAuto = trim(strAuto) & "]"
						
						Response.Write("]")
						
						ClientKeyForFileName = ClientKey
		
					
						set fs=Server.CreateObject("Scripting.FileSystemObject")
						set fs2=Server.CreateObject("Scripting.FileSystemObject")
						
						set tfile=fs.CreateTextFile(Server.MapPath("..\..\") & "\clientfiles\"  & ClientKeyForFileName & "\autocomplete\product_list_" & ClientKeyForFileName & ".json")
						tfile.WriteLine(strAuto)
						tfile.close
						set tfile=nothing
						set fs=nothing
						
						Set rsAutoComplete = Nothing
						cnnAutoComplete.Close
						Set AutoComplete = nothing
		
						
						'*********************************************************
						' END Auto Complete Product List
						'*********************************************************
						
						
						'*********************************************************
						' Begin Auto Complete Files Used By The Prospecting Module
						'*********************************************************
						
						 Response.Write("Begin Build Auto Complete JSON CITY <br>")
					 				
						'******************************************
			
		
						
							SQLAutoComplete = "SELECT DISTINCT City FROM PR_Prospects WHERE Len(City) > 1 ORDER BY City"
							
							Set cnnAutoComplete = Server.CreateObject("ADODB.Connection")
							cnnAutoComplete.open (Session("ClientCnnString"))
							Set rsAutoComplete = Server.CreateObject("ADODB.Recordset")
							rsAutoComplete.CursorLocation = 3 
							
						
							Set rsAutoComplete = cnnAutoComplete.Execute(SQLAutoComplete)
						
							If not rsAutoComplete.EOF Then
							
								CustomerCount = 0
								jsonDataCity = ""
								
								Do While Not rsAutoComplete.EOF
								
									CustomerCount = CustomerCount + 1
									
									If CustomerCount = 1 Then
										jsonDataCity = "[{""city"":""" & rsAutoComplete("city") & """},"
									Else
										jsonDataCity = jsonDataCity & "{""city"":""" & rsAutoComplete("city") & """},"
									End If
			
								
									rsAutoComplete.MoveNext
									
								Loop
								
								Session("strAccounts") = "GENERATE COMPLETE"
								
								If Len(jsonDataCity)>0 Then jsonDataCity = Left(jsonDataCity,Len(jsonDataCity)-1)
								jsonDataCity = jsonDataCity & "]"
								
							
								
							End If
							
							ClientKeyForFileName = ClientKey
						
							set fs=Server.CreateObject("Scripting.FileSystemObject")
		
							
							set tfile=fs.CreateTextFile(Server.MapPath("..\..\") & "\clientfiles\"  & ClientKeyForFileName & "\autocomplete\prospect_city_list.json")
							tfile.WriteLine(jsonDataCity)
							tfile.close
							set tfile=nothing
							set fs=nothing
							
							Set rsAutoComplete = Nothing
							cnnAutoComplete.Close
							Set AutoComplete = nothing
		
					'******************************************			 				
		
						 Response.Write("Begin Build Auto Complete JSON STATE <br>")
					 				
						'******************************************
			
		
						
							SQLAutoComplete = "SELECT DISTINCT State FROM PR_Prospects WHERE Len(State) > 1 ORDER BY State"
							
							Set cnnAutoComplete = Server.CreateObject("ADODB.Connection")
							cnnAutoComplete.open (Session("ClientCnnString"))
							Set rsAutoComplete = Server.CreateObject("ADODB.Recordset")
							rsAutoComplete.CursorLocation = 3 
							
						
							Set rsAutoComplete = cnnAutoComplete.Execute(SQLAutoComplete)
						
							If not rsAutoComplete.EOF Then
							
								CustomerCount = 0
								jsonDataState = ""
		
								
								Do While Not rsAutoComplete.EOF
								
									CustomerCount = CustomerCount + 1
			
									If CustomerCount = 1 Then
										jsonDataState = "[{""state"":""" & rsAutoComplete("state") & """},"
									Else
										jsonDataState = jsonDataState & "{""state"":""" & rsAutoComplete("state") & """},"
									End If
								
									rsAutoComplete.MoveNext
									
								Loop
								
								Session("strAccounts") = "GENERATE COMPLETE"
								
								
								If Len(jsonDataState)>0 Then jsonDataState = Left(jsonDataState,Len(jsonDataState)-1)
								jsonDataState = jsonDataState & "]"
								
								
							End If
							
							ClientKeyForFileName = ClientKey
		
						
							set fs=Server.CreateObject("Scripting.FileSystemObject")
		
							set tfile=fs.CreateTextFile(Server.MapPath("..\..\") & "\clientfiles\"  & ClientKeyForFileName & "\autocomplete\prospect_state_list.json")
							tfile.WriteLine(jsonDataState)
							tfile.close
							set tfile=nothing
							set fs=nothing
							
							Set rsAutoComplete = Nothing
							cnnAutoComplete.Close
							Set AutoComplete = nothing
		
					'******************************************
					
						 Response.Write("Begin Build Auto Complete JSON ZIP CODE <br>")
					 				
						'******************************************
			
		
						
							SQLAutoComplete = "SELECT DISTINCT PostalCode FROM PR_Prospects WHERE Len(PostalCode) > 1 ORDER BY PostalCode"
							
							Set cnnAutoComplete = Server.CreateObject("ADODB.Connection")
							cnnAutoComplete.open (Session("ClientCnnString"))
							Set rsAutoComplete = Server.CreateObject("ADODB.Recordset")
							rsAutoComplete.CursorLocation = 3 
							
						
							Set rsAutoComplete = cnnAutoComplete.Execute(SQLAutoComplete)
						
							If not rsAutoComplete.EOF Then
							
								CustomerCount = 0
								jsonDataZip = ""
								
								Do While Not rsAutoComplete.EOF
								
									CustomerCount = CustomerCount + 1
									
									If CustomerCount = 1 Then
										jsonDataZip = "[{""zip"":""" & rsAutoComplete("PostalCode") & """},"
									Else
										jsonDataZip = jsonDataZip & "{""zip"":""" & rsAutoComplete("PostalCode") & """},"
									End If
								
									rsAutoComplete.MoveNext
									
								Loop
								
								Session("strAccounts") = "GENERATE COMPLETE"
								
								
								If Len(jsonDataZip)>0 Then jsonDataZip = Left(jsonDataZip,Len(jsonDataZip)-1)
								jsonDataZip = jsonDataZip & "]"
								
								
							End If
							
							ClientKeyForFileName = ClientKey
						
							set fs=Server.CreateObject("Scripting.FileSystemObject")
							
							set tfile=fs.CreateTextFile(Server.MapPath("..\..\") & "\clientfiles\"  & ClientKeyForFileName & "\autocomplete\prospect_zip_list.json")
							tfile.WriteLine(jsonDataZip)
							tfile.close
							set tfile=nothing
							set fs=nothing
							
							Set rsAutoComplete = Nothing
							cnnAutoComplete.Close
							Set AutoComplete = nothing
		
					'******************************************	
				Response.Write("******** DONE Processing " & ClientKey  & "************<br>")
				
		End If
	End If				
	TopRecordset.movenext
	
	Loop

	TopRecordset.Close
	Set TopRecordset = Nothing
	TopConnection.Close
	Set TopConnection = Nothing
	
End If

Response.write("<script type=""text/javascript"">closeme();</script>")	
'Response.End
'*************************
'*************************
'Subs and funcs begin here


Sub SetClientCnnString

	dummy=MUV_WRITE("cnnStatus","")

	SQL = "SELECT * FROM tblServerInfo where clientKey='"& ClientKey &"'"

	Set Connection = Server.CreateObject("ADODB.Connection")
	Set Recordset = Server.CreateObject("ADODB.Recordset")
	Connection.Open InsightCnnString
	
	'Open the recordset object executing the SQL statement and return records
	Recordset.Open SQL,Connection,3,3

	
	'First lookup the ClientKey in tblServerInfo
	'If there is no record with the entered client key, close connection
	'and exit
	If Recordset.recordcount <= 0 then
		Recordset.close
		Connection.close
		set Recordset=nothing
		set Connection=nothing
	Else
		ClientCnnString = "Driver={SQL Server};Server=" & Recordset.Fields("dbServer")
		ClientCnnString = ClientCnnString & ";Database=" & Recordset.Fields("dbCatalog")
		ClientCnnString = ClientCnnString & ";Uid=" & Recordset.Fields("dbLogin")
		ClientCnnString = ClientCnnString & ";Pwd=" & Recordset.Fields("dbPassword") & ";"
		dummy = MUV_Write("ClientCnnString",ClientCnnString)
		dummy = MUV_Write("SQL_Owner",Recordset.Fields("dbLogin"))
		dummy = MUV_Write("ClientID",Recordset.Fields("clientkey"))
		Recordset.close
		Connection.close
		dummy=MUV_WRITE("cnnStatus","OK")
	End If
End Sub



%>