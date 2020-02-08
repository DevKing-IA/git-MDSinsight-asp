<!--#include file="../../../inc/subsandfuncs.asp"-->
<!--#include file="../../../inc/Insightfuncs.asp"-->
<%
'baseURL should always have a trailing /slash, just in case, handle either way
If right(baseURL,1)="/" Then maildomain = Left(right(baseURL,len(baseURL)-7),len(right(baseURL,len(baseURL)-7))-1) Else maildomain = right(baseURL,len(baseURL)-7)

Dim objConn, strFile
Dim intCampaignRecipientID
DIM buffer
DIM APBU, VedorID,Acct,DistributionDescr
buffer=array()
    
Set cnnTmpTable = Server.CreateObject("ADODB.Connection")
cnnTmpTable.open (Session("ClientCnnString"))
Set rsTmpTable = Server.CreateObject("ADODB.Recordset")
rsTmpTable.CursorLocation = 3 
SQLTmpTable = "DELETE FROM zExportPeopleSoftInclude_" & Trim(Session("userNo")) 
Set rsTmpTable = cnnTmpTable.Execute(SQLTmpTable)
set rsTmpTable = Nothing

SQLTmpTable = "SELECT * FROM settings_Reports WHERE reportNumber=8001" 
Set rsTmpTable = cnnTmpTable.Execute(SQLTmpTable)
IF NOT rsTmpTable.EOF THEN
    APBU=rsTmpTable("reportspecificdata1")
    VedorID=rsTmpTable("reportspecificdata2")
    Acct=rsTmpTable("reportspecificdata2a")
    DistributionDescr=rsTmpTable("reportspecificdata2b")
Else
		Response.Write("ERROR - Unable to read data for report number 8001 from settings_Reports table")
		Response.End
END IF
rsTmpTable.close
set rsTmpTable = Nothing
cnnTmpTable.close
set cnnTmpTable = Nothing




StartDate = MUV_READ("PSoftStartDate")
EndDate = MUV_READ("PSoftEndDate")
SelectedPeriod = MUV_READ("PSoftSelectedPeriod")
If MUV_READ("PSoftSkipZeroDollar") = "True" Then SkipZeroDollar = True Else SkipZeroDollar = False
If MUV_READ("PSoftSkipLessThanZero") = "True" Then SkipLessThanZero = True Else SkipLessThanZero = False
If MUV_READ("PSoftSkipLessThanZeroLines") = "True" Then SkipLessThanZeroLines = True Else SkipLessThanZeroLines = False

IncludedType = MUV_READ("PSoftIncludedType")
CustomOrPredefined = MUV_READ("PSoftCustomOrPredefined")
Account = MUV_READ("PSoftAccount")
DueDateDays = MUV_READ("PSoftDueDateDays")
DueDateSingleDate = MUV_READ("PSoftDueDateSingleDate") 
DoNotShowDueDate = MUV_READ("PSoftDoNotShowDueDate") 
typeOfAccounts = MUV_READ("PSofttypeOfAccounts") 
Chain = MUV_READ("PSoftChain")



'Now get the actual invoice data
SQLInvoices = "SELECT * FROM " & MUV_Read("SQL_Owner") & ".InvoiceHistory Where CustNum"
SELECT CASE  typeOfAccounts
    CASE "Account"
        SQLInvoices = SQLInvoices& "= '" & Account &"'"
    CASE "Chain"
        SQLInvoices = SQLInvoices&" IN (SELECT CustNum FROM AR_Customer WHERE ChainNum = "&Chain&")"
END SELECT
SQLInvoices = SQLInvoices & " AND IvsDate >= '" & StartDate & "' AND IvsDate <= '" & EndDate & "' "
If SkipZeroDollar = True Then SQLInvoices = SQLInvoices & "AND IvsTotalAmt <> 0 "
If SkipLessThanZero = True Then SQLInvoices = SQLInvoices & "AND IvsTotalAmt > 0 "
If IncludedType <> "" Then SQLInvoices = SQLInvoices & "AND CHARINDEX(IvsType,'" & IncludedType & "') <> 0 "
SQLInvoices = SQLInvoices & "AND IvsHistSequence NOT IN (Select IvsHistSequence from zExportPeopleSoftInvoiceOmit_" & Trim(Session("userNo")) & ") "
SQLInvoices = SQLInvoices & " order by IvsNum"

Set cnnInvoices = Server.CreateObject("ADODB.Connection")
cnnInvoices.open (Session("ClientCnnString"))
Set rsInvoices = Server.CreateObject("ADODB.Recordset")
rsInvoices.CursorLocation = 3

  
Set rsInvoices = cnnInvoices.Execute(SQLInvoices)
If not rsInvoices.Eof Then
    
    
	Set cnnTmpTable = Server.CreateObject("ADODB.Connection")
	cnnTmpTable.open (Session("ClientCnnString"))
	Set rsTmpTable = Server.CreateObject("ADODB.Recordset")
	rsTmpTable.CursorLocation = 3 

	TotalAmt = 0
    TotalInvoicesQty=0
	Do While not rsInvoices.Eof
        TotalInvoicesQty=TotalInvoicesQty+1
        TotalAmt = TotalAmt +rsInvoices("IvsTotalAmt")
        SQLInvoiceDetails =  "Select * from InvoiceHistoryDetail WHERE "
		SQLInvoiceDetails = SQLInvoiceDetails & "InvoiceHistoryDetail.IvsHistSequence = " & rsInvoices("IvsHistSequence")
			
		If SkipLessThanZeroLines = True Then SQLInvoiceDetails = SQLInvoiceDetails & "AND InvoiceHistoryDetail.itemPrice <> 0 " 
			
		SQLInvoiceDetails = SQLInvoiceDetails & " order by IvsHistDetSequence"
			
		Set cnnInvoiceDetails = Server.CreateObject("ADODB.Connection")
		cnnInvoiceDetails.open (Session("ClientCnnString"))
		Set rsInvoiceDetails = Server.CreateObject("ADODB.Recordset")
		rsInvoiceDetails.CursorLocation = 3 
		Set rsInvoiceDetails = cnnInvoiceDetails.Execute(SQLInvoiceDetails)

		If not rsInvoiceDetails.Eof Then
		    SubTot = 0
			Do While Not rsInvoiceDetails.eof
			    SubTot = SubTot + rsInvoiceDetails("itemQuantity") * rsInvoiceDetails("itemPrice")										
				rsInvoiceDetails.movenext
			Loop
            rsInvoiceDetails.Close
                
		End If
        
        rsInvoices.MoveNext
    Loop
    
    'Make file header
    buffer=AddItem(buffer,"C"&APBU&PadNumber(TotalAmt,16,1,"0",1)&PadNumber(TotalInvoicesQty, 3,2,"0",1))
    
    Dim buffData
    rsInvoices.MoveFirst
    Do While not rsInvoices.Eof
    
        'Make H record 
        buffer=AddItem(buffer,"H"&PadNumber(VedorID,10,2,"0",1)&PadNumber(rsInvoices("IvsNum"),16,2," ",1)&PadNumber(Month(rsInvoices("IvsDate")),2,2,"0",1) & PadNumber(Day(rsInvoices("IvsDate")),2,2,"0",1)&PadNumber(Year(rsInvoices("IvsDate")),4,2,"0",1)&PadNumber(rsInvoices("IvsTotalAmt"),16,1,"0",1)&"CORPH")

        
        SQLInvoiceDetails =  "SELECT * FROM InvoiceHistoryDetail WHERE "
		SQLInvoiceDetails = SQLInvoiceDetails & "InvoiceHistoryDetail.IvsHistSequence = " & rsInvoices("IvsHistSequence")
			
		If SkipLessThanZeroLines = True Then SQLInvoiceDetails = SQLInvoiceDetails & "AND InvoiceHistoryDetail.itemPrice <> 0 " 
			
		SQLInvoiceDetails = SQLInvoiceDetails & " order by IvsHistDetSequence"
			
		Set cnnInvoiceDetails = Server.CreateObject("ADODB.Connection")
		cnnInvoiceDetails.open (Session("ClientCnnString"))
		Set rsInvoiceDetails = Server.CreateObject("ADODB.Recordset")
		rsInvoiceDetails.CursorLocation = 3 
		Set rsInvoiceDetails = cnnInvoiceDetails.Execute(SQLInvoiceDetails)

		If not rsInvoiceDetails.Eof Then
		    SubTot = 0
           
			Do While Not rsInvoiceDetails.eof
			    SubTot = rsInvoiceDetails("itemQuantity") * rsInvoiceDetails("itemPrice")	
			    
                buffer=AddItem(buffer,"L"&PadNumber(Replace(Replace(Replace(Replace(rsInvoiceDetails("prodDescription"),"<",""),")",""),"(",""),"&",""),30,3," ",2)&PadNumber(SubTot, 16,1,"0",1))

                
				dummyprodvar= " " ' Unitl LIJ tells us what to do                
                buffer=AddItem(buffer,"D"&Acct&PadNumber(getSpecialData(rsInvoices("CustNum"),"DeptID"),8,3," ",2)&PadNumber(getSpecialData(rsInvoices("CustNum"),"GLBU"),5,3," ",2)&PadNumber(DistributionDescr,29,3," ",2)&PadNumber(SubTot, 16,1,"0",1)&PadNumber(dummyprodvar,5,3," ",2)&PadNumber(getSpecialData(rsInvoices("CustNum"),"Project"),6,3," ",2))

				rsInvoiceDetails.movenext
			Loop
            
                
		End If
        rsInvoiceDetails.Close

        rsInvoices.MoveNext
    LOOP
    

END IF





'-- the filename you give it will be the one that is shown
' to the users by default when they download

strFile = "invoice_psoft"&Right("0" & Day(Now),2)&+Right("0" & Month(Now),2)&YEAR(Now)&".txt"

Response.Clear
IF Request.Form("DownloadFile") = "on" THEN
    
' Download to user
Response.AddHeader "Content-Disposition", "attachment; filename=" & strFile
Response.AddHeader "Content-Length", LEN(JOIN(buffer,CHR(13)&CHR(10)))
Response.ContentType = "application/octet-stream"
Response.CharSet = "UTF-8"
'-- send the stream in the response
Response.BinaryWrite(JOIN(buffer,CHR(13)&CHR(10)))

END IF
    
    
Function AddItem(arr, val)
    ReDim Preserve arr(UBound(arr) + 1)
    arr(UBound(arr)) = val
    AddItem = arr
End Function

function PadNumber(number, width,typeNumber,padSymbol,typePad)
    dim padded
    SELECT CASE typeNumber
        CASE 1
             padded = ""&FormatNumber(number,2,0,0,0)
        CASE 2
            padded = cStr(number)
        CASE 3
            padded = number
    END SELECT
   

   while (len(padded) < width)
        IF typePAd=1 THEN
            padded = padSymbol & padded
            ELSE
                padded = padded&padSymbol
        END IF
   wend

   PadNumber = padded
end function

    FUNCTION getSpecialData(custID,specialFileldName) 
        DIM retValue
        retValue=" "
        Set SpecialDataConn = Server.CreateObject("ADODB.Connection")
        SpecialDataConn.open (Session("ClientCnnString"))
        Set SpecialDataTable = Server.CreateObject("ADODB.Recordset")
        SpecialDataTable.CursorLocation = 3 
        SpecialDataSql = "SELECT * FROM AR_CustomerBillinfo WHERE CustID="&custID&" AND IncludeOnInvoices=1 AND BillInfoFieldTitle='"& specialFileldName &"'"
        Set SpecialDataTable = SpecialDataConn.Execute(SpecialDataSql)
        IF NOT SpecialDataTable.EOF THEN
            retValue=SpecialDataTable("BillInfoFieldData")
           
        END IF
        SpecialDataTable.Close
        SET SpecialDataTable=Nothing

        SpecialDataConn.Close
        SET SpecialDataConn=Nothing
        getSpecialData=retValue
    END FUNCTION

  
%>