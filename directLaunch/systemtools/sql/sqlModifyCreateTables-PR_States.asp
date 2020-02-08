<%	
	Set cnnCheckPRStates = Server.CreateObject("ADODB.Connection")
	cnnCheckPRStates.open (Session("ClientCnnString"))
	Set rsCheckPRStates = Server.CreateObject("ADODB.Recordset")

	Err.Clear
	on error resume next
	Set rsCheckPRStates = cnnCheckPRStates.Execute("SELECT TOP 1 * FROM PR_States")
	If Err.Description <> "" Then
		If InStr(Ucase(Err.Description),"INVALID OBJECT NAME") <> 0 Then
			On error goto 0		
			'The table is not there, we need to create it
			
		    SQLCheckPRStates = "CREATE TABLE [PR_States]("
			SQLCheckPRStates = SQLCheckPRStates & " [stateID] [int] IDENTITY(1,1) NOT NULL,"
			SQLCheckPRStates = SQLCheckPRStates & " [stateCode] [nchar](2) NOT NULL,"
			SQLCheckPRStates = SQLCheckPRStates & " [stateName] [nvarchar](128) NOT NULL,"
			SQLCheckPRStates = SQLCheckPRStates & " CONSTRAINT [PK_state] PRIMARY KEY CLUSTERED ("
			SQLCheckPRStates = SQLCheckPRStates & " [stateID] ASC"
			SQLCheckPRStates = SQLCheckPRStates & " )WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]"
			SQLCheckPRStates = SQLCheckPRStates & " ) ON [PRIMARY] "
			
			Set rsCheckPRStates = cnnCheckPRStates.Execute(SQLCheckPRStates)

		
			SQLCheckPRStates = "INSERT [PR_States] ( [stateCode], [stateName]) VALUES (N'AL', N'Alabama')"
			Set rsCheckPRStates = cnnCheckPRStates.Execute(SQLCheckPRStates)
			
			SQLCheckPRStates = "INSERT [PR_States] ( [stateCode], [stateName]) VALUES (N'AK', N'Alaska')"
			Set rsCheckPRStates = cnnCheckPRStates.Execute(SQLCheckPRStates)
			
			SQLCheckPRStates = "INSERT [PR_States] ( [stateCode], [stateName]) VALUES (N'AZ', N'Arizona')"
			Set rsCheckPRStates = cnnCheckPRStates.Execute(SQLCheckPRStates)
			
			SQLCheckPRStates = "INSERT [PR_States] ( [stateCode], [stateName]) VALUES (N'AR', N'Arkansas')"
			Set rsCheckPRStates = cnnCheckPRStates.Execute(SQLCheckPRStates)
			
			SQLCheckPRStates = "INSERT [PR_States] ( [stateCode], [stateName]) VALUES (N'CA', N'California')"
			Set rsCheckPRStates = cnnCheckPRStates.Execute(SQLCheckPRStates)
			
			SQLCheckPRStates = "INSERT [PR_States] ( [stateCode], [stateName]) VALUES (N'CO', N'Colorado')"
			Set rsCheckPRStates = cnnCheckPRStates.Execute(SQLCheckPRStates)
			
			SQLCheckPRStates = "INSERT [PR_States] ( [stateCode], [stateName]) VALUES (N'CT', N'Connecticut')"
			Set rsCheckPRStates = cnnCheckPRStates.Execute(SQLCheckPRStates)
			
			SQLCheckPRStates = "INSERT [PR_States] ( [stateCode], [stateName]) VALUES (N'DE', N'Delaware')"
			Set rsCheckPRStates = cnnCheckPRStates.Execute(SQLCheckPRStates)
			
			SQLCheckPRStates = "INSERT [PR_States] ( [stateCode], [stateName]) VALUES (N'DC', N'District of Columbia')"
			Set rsCheckPRStates = cnnCheckPRStates.Execute(SQLCheckPRStates)
			
			SQLCheckPRStates = "INSERT [PR_States] ( [stateCode], [stateName]) VALUES (N'FL', N'Florida')"
			Set rsCheckPRStates = cnnCheckPRStates.Execute(SQLCheckPRStates)
			
			SQLCheckPRStates = "INSERT [PR_States] ( [stateCode], [stateName]) VALUES (N'GA', N'Georgia')"
			Set rsCheckPRStates = cnnCheckPRStates.Execute(SQLCheckPRStates)
			
			SQLCheckPRStates = "INSERT [PR_States] ( [stateCode], [stateName]) VALUES (N'HI', N'Hawaii')"
			Set rsCheckPRStates = cnnCheckPRStates.Execute(SQLCheckPRStates)
			
			SQLCheckPRStates = "INSERT [PR_States] ( [stateCode], [stateName]) VALUES (N'ID', N'Idaho')"
			Set rsCheckPRStates = cnnCheckPRStates.Execute(SQLCheckPRStates)
			
			SQLCheckPRStates = "INSERT [PR_States] ( [stateCode], [stateName]) VALUES (N'IL', N'Illinois')"
			Set rsCheckPRStates = cnnCheckPRStates.Execute(SQLCheckPRStates)
			
			SQLCheckPRStates = "INSERT [PR_States] ( [stateCode], [stateName]) VALUES (N'IN', N'Indiana')"
			Set rsCheckPRStates = cnnCheckPRStates.Execute(SQLCheckPRStates)
			
			SQLCheckPRStates = "INSERT [PR_States] ( [stateCode], [stateName]) VALUES (N'IA', N'Iowa')"
			Set rsCheckPRStates = cnnCheckPRStates.Execute(SQLCheckPRStates)
			
			SQLCheckPRStates = "INSERT [PR_States] ( [stateCode], [stateName]) VALUES (N'KS', N'Kansas')"
			Set rsCheckPRStates = cnnCheckPRStates.Execute(SQLCheckPRStates)
			
			SQLCheckPRStates = "INSERT [PR_States] ( [stateCode], [stateName]) VALUES (N'KY', N'Kentucky')"
			Set rsCheckPRStates = cnnCheckPRStates.Execute(SQLCheckPRStates)
			
			SQLCheckPRStates = "INSERT [PR_States] ( [stateCode], [stateName]) VALUES (N'LA', N'Louisiana')"
			Set rsCheckPRStates = cnnCheckPRStates.Execute(SQLCheckPRStates)
			
			SQLCheckPRStates = "INSERT [PR_States] ( [stateCode], [stateName]) VALUES (N'ME', N'Maine')"
			Set rsCheckPRStates = cnnCheckPRStates.Execute(SQLCheckPRStates)
			
			SQLCheckPRStates = "INSERT [PR_States] ( [stateCode], [stateName]) VALUES (N'MD', N'Maryland')"
			Set rsCheckPRStates = cnnCheckPRStates.Execute(SQLCheckPRStates)
			
			SQLCheckPRStates = "INSERT [PR_States] ( [stateCode], [stateName]) VALUES (N'MA', N'Massachusetts')"
			Set rsCheckPRStates = cnnCheckPRStates.Execute(SQLCheckPRStates)
			
			SQLCheckPRStates = "INSERT [PR_States] ( [stateCode], [stateName]) VALUES (N'MI', N'Michigan')"
			Set rsCheckPRStates = cnnCheckPRStates.Execute(SQLCheckPRStates)
			
			SQLCheckPRStates = "INSERT [PR_States] ( [stateCode], [stateName]) VALUES (N'MN', N'Minnesota')"
			Set rsCheckPRStates = cnnCheckPRStates.Execute(SQLCheckPRStates)
			
			SQLCheckPRStates = "INSERT [PR_States] ( [stateCode], [stateName]) VALUES (N'MS', N'Mississippi')"
			Set rsCheckPRStates = cnnCheckPRStates.Execute(SQLCheckPRStates)
			
			SQLCheckPRStates = "INSERT [PR_States] ( [stateCode], [stateName]) VALUES (N'MO', N'Missouri')"
			Set rsCheckPRStates = cnnCheckPRStates.Execute(SQLCheckPRStates)
			
			SQLCheckPRStates = "INSERT [PR_States] ( [stateCode], [stateName]) VALUES (N'MT', N'Montana')"
			Set rsCheckPRStates = cnnCheckPRStates.Execute(SQLCheckPRStates)
			
			SQLCheckPRStates = "INSERT [PR_States] ( [stateCode], [stateName]) VALUES (N'NE', N'Nebraska')"
			Set rsCheckPRStates = cnnCheckPRStates.Execute(SQLCheckPRStates)
			
			SQLCheckPRStates = "INSERT [PR_States] ( [stateCode], [stateName]) VALUES (N'NV', N'Nevada')"
			Set rsCheckPRStates = cnnCheckPRStates.Execute(SQLCheckPRStates)
			
			SQLCheckPRStates = "INSERT [PR_States] ( [stateCode], [stateName]) VALUES (N'NH', N'New Hampshire')"
			Set rsCheckPRStates = cnnCheckPRStates.Execute(SQLCheckPRStates)
			
			SQLCheckPRStates = "INSERT [PR_States] ( [stateCode], [stateName]) VALUES (N'NJ', N'New Jersey')"
			Set rsCheckPRStates = cnnCheckPRStates.Execute(SQLCheckPRStates)
			
			SQLCheckPRStates = "INSERT [PR_States] ( [stateCode], [stateName]) VALUES (N'NM', N'New Mexico')"
			Set rsCheckPRStates = cnnCheckPRStates.Execute(SQLCheckPRStates)
			
			SQLCheckPRStates = "INSERT [PR_States] ( [stateCode], [stateName]) VALUES (N'NY', N'New York')"
			Set rsCheckPRStates = cnnCheckPRStates.Execute(SQLCheckPRStates)
			
			SQLCheckPRStates = "INSERT [PR_States] ( [stateCode], [stateName]) VALUES (N'NC', N'North Carolina')"
			Set rsCheckPRStates = cnnCheckPRStates.Execute(SQLCheckPRStates)
			
			SQLCheckPRStates = "INSERT [PR_States] ( [stateCode], [stateName]) VALUES (N'ND', N'North Dakota')"
			Set rsCheckPRStates = cnnCheckPRStates.Execute(SQLCheckPRStates)
			
			SQLCheckPRStates = "INSERT [PR_States] ( [stateCode], [stateName]) VALUES (N'OH', N'Ohio')"
			Set rsCheckPRStates = cnnCheckPRStates.Execute(SQLCheckPRStates)
			
			SQLCheckPRStates = "INSERT [PR_States] ( [stateCode], [stateName]) VALUES (N'OK', N'Oklahoma')"
			Set rsCheckPRStates = cnnCheckPRStates.Execute(SQLCheckPRStates)
			
			SQLCheckPRStates = "INSERT [PR_States] ( [stateCode], [stateName]) VALUES (N'OR', N'Oregon')"
			Set rsCheckPRStates = cnnCheckPRStates.Execute(SQLCheckPRStates)
			
			SQLCheckPRStates = "INSERT [PR_States] ( [stateCode], [stateName]) VALUES (N'PA', N'Pennsylvania')"
			Set rsCheckPRStates = cnnCheckPRStates.Execute(SQLCheckPRStates)
			
			SQLCheckPRStates = "INSERT [PR_States] ( [stateCode], [stateName]) VALUES (N'PR', N'Puerto Rico')"
			Set rsCheckPRStates = cnnCheckPRStates.Execute(SQLCheckPRStates)
			
			SQLCheckPRStates = "INSERT [PR_States] ( [stateCode], [stateName]) VALUES (N'RI', N'Rhode Island')"
			Set rsCheckPRStates = cnnCheckPRStates.Execute(SQLCheckPRStates)
			
			SQLCheckPRStates = "INSERT [PR_States] ( [stateCode], [stateName]) VALUES (N'SC', N'South Carolina')"
			Set rsCheckPRStates = cnnCheckPRStates.Execute(SQLCheckPRStates)
			
			SQLCheckPRStates = "INSERT [PR_States] ( [stateCode], [stateName]) VALUES (N'SD', N'South Dakota')"
			Set rsCheckPRStates = cnnCheckPRStates.Execute(SQLCheckPRStates)
			
			SQLCheckPRStates = "INSERT [PR_States] ( [stateCode], [stateName]) VALUES (N'TN', N'Tennessee')"
			Set rsCheckPRStates = cnnCheckPRStates.Execute(SQLCheckPRStates)
			
			SQLCheckPRStates = "INSERT [PR_States] ( [stateCode], [stateName]) VALUES (N'TX', N'Texas')"
			Set rsCheckPRStates = cnnCheckPRStates.Execute(SQLCheckPRStates)
			
			SQLCheckPRStates = "INSERT [PR_States] ( [stateCode], [stateName]) VALUES (N'UT', N'Utah')"
			Set rsCheckPRStates = cnnCheckPRStates.Execute(SQLCheckPRStates)
			
			SQLCheckPRStates = "INSERT [PR_States] ( [stateCode], [stateName]) VALUES (N'VT', N'Vermont')"
			Set rsCheckPRStates = cnnCheckPRStates.Execute(SQLCheckPRStates)
			
			SQLCheckPRStates = "INSERT [PR_States] ( [stateCode], [stateName]) VALUES (N'VA', N'Virginia')"
			Set rsCheckPRStates = cnnCheckPRStates.Execute(SQLCheckPRStates)
			
			SQLCheckPRStates = "INSERT [PR_States] ( [stateCode], [stateName]) VALUES (N'WA', N'Washington')"
			Set rsCheckPRStates = cnnCheckPRStates.Execute(SQLCheckPRStates)
			
			SQLCheckPRStates = "INSERT [PR_States] ( [stateCode], [stateName]) VALUES (N'WV', N'West Virginia')"
			Set rsCheckPRStates = cnnCheckPRStates.Execute(SQLCheckPRStates)
			
			SQLCheckPRStates = "INSERT [PR_States] ( [stateCode], [stateName]) VALUES (N'WI', N'Wisconsin')"
			Set rsCheckPRStates = cnnCheckPRStates.Execute(SQLCheckPRStates)
			
			SQLCheckPRStates = "INSERT [PR_States] ( [stateCode], [stateName]) VALUES (N'WY', N'Wyoming')"
			Set rsCheckPRStates = cnnCheckPRStates.Execute(SQLCheckPRStates)
			
		   
		End If
	End If

	
	set rsCheckPRStates = nothing
	cnnCheckPRStates.close
	set cnnCheckPRStates = nothing
				
%>