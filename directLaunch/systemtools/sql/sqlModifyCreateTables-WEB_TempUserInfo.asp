<%	

	Set cnnWEB_TempUserInfo = Server.CreateObject("ADODB.Connection")
	cnnWEB_TempUserInfo.open (Session("ClientCnnString"))
	Set rsWEB_TempUserInfo = Server.CreateObject("ADODB.Recordset")
	rsWEB_TempUserInfo.CursorLocation = 3 


	Err.Clear
	on error resume next
	Set rsWEB_TempUserInfo = cnnWEB_TempUserInfo.Execute("SELECT TOP 1 * FROM WEB_TempUserInfo")
	If Err.Description <> "" Then
		If InStr(Ucase(Err.Description),"INVALID OBJECT NAME") <> 0 Then
			On error goto 0		
			'The table is not there, we need to create it
			
		   SQLWEB_TempUserInfo = "CREATE TABLE [WEB_TempUserInfo]("
		   SQLWEB_TempUserInfo = SQLWEB_TempUserInfo & " [InternalRecordIdentifier] [int] IDENTITY(2,1) NOT NULL, "
		   SQLWEB_TempUserInfo = SQLWEB_TempUserInfo & " [RecordCreationDateTime] [datetime] NULL CONSTRAINT [DF_WEB_TempUserInfo]  DEFAULT (getdate()), "
	       SQLWEB_TempUserInfo = SQLWEB_TempUserInfo & " [tmpOrderID] [int] NOT NULL,"
	       SQLWEB_TempUserInfo = SQLWEB_TempUserInfo & " [tmpUserNo] [int] NULL,"
	       SQLWEB_TempUserInfo = SQLWEB_TempUserInfo & " [tmpCustID] [varchar](50) NULL,"
	       SQLWEB_TempUserInfo = SQLWEB_TempUserInfo & " [tmpOCSList] [int] NULL,"
	       SQLWEB_TempUserInfo = SQLWEB_TempUserInfo & " [tmpPromoCode] [varchar](20) NULL,"
	       SQLWEB_TempUserInfo = SQLWEB_TempUserInfo & " [tmpBillToName] [varchar](75) NULL,"
	       SQLWEB_TempUserInfo = SQLWEB_TempUserInfo & " [tmpBillToAddress1] [varchar](50) NULL,"
	       SQLWEB_TempUserInfo = SQLWEB_TempUserInfo & " [tmpBillToAddress2] [varchar](50) NULL,"
	       SQLWEB_TempUserInfo = SQLWEB_TempUserInfo & " [tmpBillToCity] [varchar](50) NULL,"
	       SQLWEB_TempUserInfo = SQLWEB_TempUserInfo & " [tmpBillToState] [varchar](25) NULL,"
	       SQLWEB_TempUserInfo = SQLWEB_TempUserInfo & " [tmpBillToZip] [varchar](10) NULL,"
	       SQLWEB_TempUserInfo = SQLWEB_TempUserInfo & " [tmpBillToPhone] [varchar](25) NULL,"
	       SQLWEB_TempUserInfo = SQLWEB_TempUserInfo & " [tmpBillToPhoneExt] [varchar](10) NULL,"
	       SQLWEB_TempUserInfo = SQLWEB_TempUserInfo & " [tmpFax] [varchar](25) NULL,"
	       SQLWEB_TempUserInfo = SQLWEB_TempUserInfo & " [tmpEmail] [varchar](100) NULL,"
	       SQLWEB_TempUserInfo = SQLWEB_TempUserInfo & " [tmpPassword] [varchar](10) NULL,"
	       SQLWEB_TempUserInfo = SQLWEB_TempUserInfo & " [tmpDevDate] [varchar](50) NULL,"
	       SQLWEB_TempUserInfo = SQLWEB_TempUserInfo & " [tmpTerms] [varchar](50) NULL,"
	       SQLWEB_TempUserInfo = SQLWEB_TempUserInfo & " [tmpBillToCompany] [varchar](50) NULL,"
	       SQLWEB_TempUserInfo = SQLWEB_TempUserInfo & " [tmpDepartment] [varchar](50) NULL,"
	       SQLWEB_TempUserInfo = SQLWEB_TempUserInfo & " [tmpCostCenter] [varchar](50) NULL,"
	       SQLWEB_TempUserInfo = SQLWEB_TempUserInfo & " [tmpCCDescription] [varchar](50) NULL,"
	       SQLWEB_TempUserInfo = SQLWEB_TempUserInfo & " [tmpPO] [varchar](50) NULL,"
	       SQLWEB_TempUserInfo = SQLWEB_TempUserInfo & " [tmpShipToName] [varchar](50) NULL,"
	       SQLWEB_TempUserInfo = SQLWEB_TempUserInfo & " [tmpShipToCompany] [varchar](50) NULL,"
	       SQLWEB_TempUserInfo = SQLWEB_TempUserInfo & " [tmpShipToAddress1] [varchar](50) NULL,"
	       SQLWEB_TempUserInfo = SQLWEB_TempUserInfo & " [tmpShipToAddress2] [varchar](50) NULL,"
	       SQLWEB_TempUserInfo = SQLWEB_TempUserInfo & " [tmpShipToCity] [varchar](50) NULL,"
	       SQLWEB_TempUserInfo = SQLWEB_TempUserInfo & " [tmpShipToState] [varchar](25) NULL,"
	       SQLWEB_TempUserInfo = SQLWEB_TempUserInfo & " [tmpShipToZip] [varchar](10) NULL,"
	       SQLWEB_TempUserInfo = SQLWEB_TempUserInfo & " [tmpShipToPhone] [varchar](25) NULL,"
	       SQLWEB_TempUserInfo = SQLWEB_TempUserInfo & " [tmpShipToPhoneExt] [varchar](10) NULL,"
	       SQLWEB_TempUserInfo = SQLWEB_TempUserInfo & " [tmpNameOnCard] [varchar](50) NULL,"
	       SQLWEB_TempUserInfo = SQLWEB_TempUserInfo & " [tmpCCType] [varchar](50) NULL,"
	       SQLWEB_TempUserInfo = SQLWEB_TempUserInfo & " [tmpCCNumber] [varchar](100) NULL,"
	       SQLWEB_TempUserInfo = SQLWEB_TempUserInfo & " [tmpCCExpMonth] [varchar](15) NULL,"
	       SQLWEB_TempUserInfo = SQLWEB_TempUserInfo & " [tmpCCExpYear] [varchar](10) NULL,"
	       SQLWEB_TempUserInfo = SQLWEB_TempUserInfo & " [tmpGiftMessage] [varchar](255) NULL,"
	       SQLWEB_TempUserInfo = SQLWEB_TempUserInfo & " [tmpShipMethod] [varchar](50) NULL,"
	       SQLWEB_TempUserInfo = SQLWEB_TempUserInfo & " [tmpRecurrOrdName] [varchar](50) NULL,"
	       SQLWEB_TempUserInfo = SQLWEB_TempUserInfo & " [tmpRecurrOrdDay] [int] NULL,"
	       SQLWEB_TempUserInfo = SQLWEB_TempUserInfo & " [tmpBillToCheckName] [varchar](255) NULL,"
	       SQLWEB_TempUserInfo = SQLWEB_TempUserInfo & " [tmpCheckNumber] [int] NULL,"
	       SQLWEB_TempUserInfo = SQLWEB_TempUserInfo & " [tmpMICRNumber] [varchar](100) NULL,"
	       SQLWEB_TempUserInfo = SQLWEB_TempUserInfo & " [tmpBillToStateID] [varchar](50) NULL,"
	       SQLWEB_TempUserInfo = SQLWEB_TempUserInfo & " [tmpCheckType] [varchar](50) NULL,"
	       SQLWEB_TempUserInfo = SQLWEB_TempUserInfo & " [tmpCheckDLSID] [varchar](100) NULL,"
	       SQLWEB_TempUserInfo = SQLWEB_TempUserInfo & " [tmpPaymentMethod] [varchar](50) NULL,"
	       SQLWEB_TempUserInfo = SQLWEB_TempUserInfo & " [tmpShipToAddressType] [varchar](50) NULL,"
	       SQLWEB_TempUserInfo = SQLWEB_TempUserInfo & " [tmpComments] [varchar](5000) NULL"
		   SQLWEB_TempUserInfo = SQLWEB_TempUserInfo & " ) ON [PRIMARY]"

			
			Set rsWEB_TempUserInfo = cnnWEB_TempUserInfo.Execute(SQLWEB_TempUserInfo)
		End If
	End If
	
	
	set rsWEB_TempUserInfo = nothing
	cnnWEB_TempUserInfo.close
	set cnnWEB_TempUserInfo = nothing


%>