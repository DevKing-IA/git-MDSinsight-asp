<%	
	Set cnnUser_Teams = Server.CreateObject("ADODB.Connection")
	cnnUser_Teams.open (Session("ClientCnnString"))
	Set rsUser_Teams = Server.CreateObject("ADODB.Recordset")
	rsUser_Teams.CursorLocation = 3 

	Err.Clear
	on error resume next
	Set rsUser_Teams = cnnUser_Teams.Execute("SELECT * FROM USER_Teams")
	If Err.Description <> "" Then
		If InStr(Ucase(Err.Description),"INVALID OBJECT NAME") <> 0 Then
			On error goto 0		
			'The table is not there, we need to create it
			SQLUser_Teams = "CREATE TABLE [USER_Teams]( "
			SQLUser_Teams = SQLUser_Teams & "InternalRecordIdentifier [int] IDENTITY(1,1) NOT NULL, "
			SQLUser_Teams = SQLUser_Teams & "RecordCreationDateTime [datetime] NULL CONSTRAINT [DF_User_Teams_RecordCreationDateTime] DEFAULT (getdate()), "
			SQLUser_Teams = SQLUser_Teams & " [TeamName] [varchar](1000) NULL,"
			SQLUser_Teams = SQLUser_Teams & " [TeamUserNos] [varchar](8000) NULL "
			SQLUser_Teams = SQLUser_Teams & ") ON [PRIMARY]"
						
			Set rsUser_Teams = cnnUser_Teams.Execute(SQLUser_Teams)
			
		End If
	End If

	set rsUser_Teams = nothing
	cnnUser_Teams.close
	set cnnUser_Teams = nothing
%>