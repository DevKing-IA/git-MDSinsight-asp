<%	
	Response.Write("sqlModifyCreateTables-AR_Chain.asp" & "<br>")
	On Error Goto 0

	Set cnnCheckARChain = Server.CreateObject("ADODB.Connection")
	cnnCheckARChain.open (Session("ClientCnnString"))
	Set rsCheckARChain = Server.CreateObject("ADODB.Recordset")

	Err.Clear
	on error resume next
	Set rsCheckARChain = cnnCheckARChain.Execute("SELECT TOP 1 * FROM AR_Chain")
	If Err.Description <> "" Then
		If InStr(Ucase(Err.Description),"INVALID OBJECT NAME") <> 0 Then
			On error goto 0		
			'The table is not there, we need to create it
			
		    SQLCheckARChain = "CREATE TABLE [AR_Chain]("
			SQLCheckARChain = SQLCheckARChain & " [InternalRecordIdentifier] [int] IDENTITY(1,1) NOT NULL,"
			SQLCheckARChain = SQLCheckARChain & " [RecordCreationDateTime] [datetime] NULL CONSTRAINT [DF_AR_Chain_RecordCreationDateTime]  DEFAULT (getdate()),"
			SQLCheckARChain = SQLCheckARChain & "ChainID varchar (255) NULL, "
			SQLCheckARChain = SQLCheckARChain & "Description varchar (1000) NULL, "
			SQLCheckARChain = SQLCheckARChain & "mtdQty0 int NULL, "
			SQLCheckARChain = SQLCheckARChain & "mtdAmt0 money NULL, "
			SQLCheckARChain = SQLCheckARChain & "mtdQty1 int NULL, "
			SQLCheckARChain = SQLCheckARChain & "mtdAmt1 money NULL, "
			SQLCheckARChain = SQLCheckARChain & "mtdQty2 int NULL, "
			SQLCheckARChain = SQLCheckARChain & "mtdAmt2 money NULL, "
			SQLCheckARChain = SQLCheckARChain & "mtdQty3 int NULL, "
			SQLCheckARChain = SQLCheckARChain & "mtdAmt3 money NULL, "
			SQLCheckARChain = SQLCheckARChain & "mtdQty4 int NULL, "
			SQLCheckARChain = SQLCheckARChain & "mtdAmt4 money NULL, "
			SQLCheckARChain = SQLCheckARChain & "mtdQty5 int NULL, "
			SQLCheckARChain = SQLCheckARChain & "mtdAmt5 money NULL, "
			SQLCheckARChain = SQLCheckARChain & "mtdQty6 int NULL, "
			SQLCheckARChain = SQLCheckARChain & "mtdAmt6 money NULL, "
			SQLCheckARChain = SQLCheckARChain & "mtdQty7 int NULL, "
			SQLCheckARChain = SQLCheckARChain & "mtdAmt7 money NULL, "
			SQLCheckARChain = SQLCheckARChain & "mtdQty8 int NULL, "
			SQLCheckARChain = SQLCheckARChain & "mtdAmt8 money NULL, "
			SQLCheckARChain = SQLCheckARChain & "mtdQty9 int NULL, "
			SQLCheckARChain = SQLCheckARChain & "mtdAmt9 money NULL, "
			SQLCheckARChain = SQLCheckARChain & "mtdQty10 int NULL, "
			SQLCheckARChain = SQLCheckARChain & "mtdAmt10 money NULL, "
			SQLCheckARChain = SQLCheckARChain & "mtdQty11 int NULL, "
			SQLCheckARChain = SQLCheckARChain & "mtdAmt11 money NULL, "
			SQLCheckARChain = SQLCheckARChain & "updateDiscount varchar (255) NULL, "
			SQLCheckARChain = SQLCheckARChain & "SellOnlyQuoted varchar (255) NULL, "
			SQLCheckARChain = SQLCheckARChain & "qcd0 int NULL, "
			SQLCheckARChain = SQLCheckARChain & "qcd1 int NULL, "
			SQLCheckARChain = SQLCheckARChain & "qcd2 int NULL, "
			SQLCheckARChain = SQLCheckARChain & "qcd3 int NULL, "
			SQLCheckARChain = SQLCheckARChain & "qcd4 int NULL, "
			SQLCheckARChain = SQLCheckARChain & "qcd5 int NULL, "
			SQLCheckARChain = SQLCheckARChain & "qcd6 int NULL, "
			SQLCheckARChain = SQLCheckARChain & "qcd7 int NULL, "
			SQLCheckARChain = SQLCheckARChain & "qcd8 int NULL, "
			SQLCheckARChain = SQLCheckARChain & "qcd9 int NULL, "
			SQLCheckARChain = SQLCheckARChain & "qcd10 int NULL, "
			SQLCheckARChain = SQLCheckARChain & "qcd11 int NULL, "
			SQLCheckARChain = SQLCheckARChain & "qcd12 int NULL, "
			SQLCheckARChain = SQLCheckARChain & "qcd13 int NULL, "
			SQLCheckARChain = SQLCheckARChain & "qcd14 int NULL, "
			SQLCheckARChain = SQLCheckARChain & "qcd15 int NULL, "
			SQLCheckARChain = SQLCheckARChain & "qcd16 int NULL, "
			SQLCheckARChain = SQLCheckARChain & "qcd17 int NULL, "
			SQLCheckARChain = SQLCheckARChain & "qcd18 int NULL, "
			SQLCheckARChain = SQLCheckARChain & "qcd19 int NULL, "
			SQLCheckARChain = SQLCheckARChain & "qcd20 int NULL, "			
			SQLCheckARChain = SQLCheckARChain & "chainPrice varchar (255) NULL, "
			SQLCheckARChain = SQLCheckARChain & "poFlag varchar (255) NULL, "
			SQLCheckARChain = SQLCheckARChain & "purchaseOrder varchar (255) NULL, "
			SQLCheckARChain = SQLCheckARChain & "programType int NULL, "
			SQLCheckARChain = SQLCheckARChain & "primarySalesman int NULL, "
			SQLCheckARChain = SQLCheckARChain & "webRequiredFields varchar (255) NULL, "
			SQLCheckARChain = SQLCheckARChain & "defQuoteValidDate date NULL "
			SQLCheckARChain = SQLCheckARChain & " ) ON [PRIMARY]"      
		   Set rsCheckARChain = cnnCheckARChain.Execute(SQLCheckARChain)
		   
		End If
	End If


	
	set rsCheckARChain = nothing
	cnnCheckARChain.close
	set cnnCheckARChain = nothing
				
%>