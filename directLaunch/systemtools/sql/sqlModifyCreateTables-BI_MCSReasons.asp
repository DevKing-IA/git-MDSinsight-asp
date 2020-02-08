<%	
	Set cnnBI_MCSReasons = Server.CreateObject("ADODB.Connection")
	cnnBI_MCSReasons.open (Session("ClientCnnString"))
	Set rsBI_MCSReasons = Server.CreateObject("ADODB.Recordset")
	rsBI_MCSReasons.CursorLocation = 3 


	Err.Clear
	on error resume next
	Set rsBI_MCSReasons = cnnBI_MCSReasons.Execute("SELECT TOP 1 * FROM BI_MCSReasons ORDER BY InternalRecordIdentifier DESC")
	If Err.Description <> "" Then
		If InStr(Ucase(Err.Description),"INVALID OBJECT NAME") <> 0 Then
			On error goto 0		
			'The table is not there, we need to create it
			SQLBI_MCSReasons = "CREATE TABLE [BI_MCSReasons]( "
			SQLBI_MCSReasons = SQLBI_MCSReasons & " [InternalRecordIdentifier] [int] IDENTITY(2,1) NOT NULL, "
			SQLBI_MCSReasons = SQLBI_MCSReasons & " [RecordCreationDateTime] [datetime] NULL CONSTRAINT [DF_biMCSReasons]  DEFAULT (getdate()), "
			SQLBI_MCSReasons = SQLBI_MCSReasons & " [Reason] [varchar](255) NULL "
			SQLBI_MCSReasons = SQLBI_MCSReasons & ") ON [PRIMARY]"
		
			Set rsBI_MCSReasons = cnnBI_MCSReasons.Execute(SQLBI_MCSReasons)
		End If
	End If
	
	set rsBI_MCSReasons = nothing
	cnnBI_MCSReasons.close
	set cnnBI_MCSReasons = nothing
				
%>