<% Server.ScriptTimeout = 360 %>

<!--#include file="../../../inc/InSightFuncs_BizIntel.asp"--> 
<!--#include file="../../../inc/InSightFuncs_Equipment.asp"--> 
<!--#include file="../../../inc/InSightFuncs_InventoryControl.asp"--> 
<!--#include file="../../../inc/InSightFuncs.asp"--> 
<!--#include file="../../../inc/InSightFuncs_Users.asp"-->
<!--#include file="../../../inc/InSightFuncs_AR_AP.asp"-->		

<%		
	DIM JSONdata
	JSONdata=""
	IF Request.QueryString("length")="-1" THEN
		PageSize=0
		ELSE
		
			PageSize=CLng(Request.QueryString("length"))
	END IF
	IF Request.QueryString("start")="NaN" THEN
		rowStart=1
		ELSE
			rowStart=CLng(Request.QueryString("start"))
	END IF
	IF PageSize=0 THEN
		nPage=1
		ELSE
			nPage=1+rowStart/PageSize
			
	END IF
	
	regionID=Request.QueryString("regionID")
	searchValue=Request.QueryString("search[value]")
	orderValue=Request.QueryString("order[0][column]")
	orderType=Request.QueryString("order[0][dir]")
	
	SQL = "SELECT FS_CustomerFilters.*,IC_Filters.Description AS filterDescription,IC_Filters.FilterID AS FilterID,AR_Customer.name AS custName, "
	
	SQL=SQL & " CASE WHEN FS_CustomerFilters.FrequencyType='D' THEN DATEADD(day, FS_CustomerFilters.FrequencyTime, FS_CustomerFilters.LastChangeDateTime) "

	SQL=SQL & " WHEN FS_CustomerFilters.FrequencyType='M' THEN DATEADD(day, 28*FS_CustomerFilters.FrequencyTime, FS_CustomerFilters.LastChangeDateTime) "
	SQL=SQL & " WHEN FS_CustomerFilters.FrequencyType='W' THEN DATEADD(day, 7*FS_CustomerFilters.FrequencyTime, FS_CustomerFilters.LastChangeDateTime) "
	SQL=SQL & " ELSE FS_CustomerFilters.LastChangeDateTime END AS NextChangeDateTime"
	
	
	SQL=SQL & ", CASE WHEN FS_CustomerFilters.FrequencyType='D' THEN DATEDIFF(day,GETDATE(),DATEADD(day, FS_CustomerFilters.FrequencyTime, FS_CustomerFilters.LastChangeDateTime)) "

	SQL=SQL & " WHEN FS_CustomerFilters.FrequencyType='M' THEN DATEDIFF(day,GETDATE(),DATEADD(day, 28*FS_CustomerFilters.FrequencyTime, FS_CustomerFilters.LastChangeDateTime)) "
	SQL=SQL & " WHEN FS_CustomerFilters.FrequencyType='W' THEN DATEDIFF(day,GETDATE(),DATEADD(day, 7*FS_CustomerFilters.FrequencyTime, FS_CustomerFilters.LastChangeDateTime)) "
	SQL=SQL & " ELSE DATEDIFF(day,GETDATE(),FS_CustomerFilters.LastChangeDateTime) END AS tillDays"
	
	SQL=SQL & ", FS_CustomerFilters.qty*FS_CustomerFilters.price AS linetotal"
	
	SQL=SQL & ", IsNULL(AR_Regions.Region,'Undefined') As Region"
	
	SQL=SQL & " FROM FS_CustomerFilters,IC_Filters,AR_Customer "
	
	SQL=SQL & " LEFT OUTER JOIN AR_Regions ON "
	SQL=SQL & " (AR_Customer.City Is NOT NULL AND AR_Customer.[State] Is NOT NULL AND AR_Regions.StateForCities=AR_Customer.[State] AND "
	SQL=SQL & " (CHARINDEX(AR_Customer.City, { fn CONCAT({ fn CONCAT({ fn CONCAT({ fn CONCAT(REPLACE(REPLACE(AR_Regions.Cities1, ' ,', ','), ', ', ','), ',') },REPLACE(REPLACE(AR_Regions.Cities2, ' ,', ','), ', ', ',')) }, ',') }, REPLACE(REPLACE(AR_Regions.Cities3, ' ,', ','), ', ', ','))})>0 "
	SQL=SQL & " OR "
	SQL=SQL & " CHARINDEX(AR_Customer.City, ','+{ fn CONCAT({ fn CONCAT({ fn CONCAT({ fn CONCAT(REPLACE(REPLACE(AR_Regions.Cities1, ' ,', ','), ', ', ','), ',') },REPLACE(REPLACE(AR_Regions.Cities2, ' ,', ','), ', ', ',')) }, ',') }, REPLACE(REPLACE(AR_Regions.Cities3, ' ,', ','), ', ', ','))})>0)) "
	SQL=SQL & " OR "
	SQL=SQL & " (AR_Customer.Zip Is NOT NULL AND ({ fn CONCAT({ fn CONCAT(REPLACE(REPLACE(AR_Regions.ZipOrPostalCodes1, ' ,', ','), ', ', ','), ',') },REPLACE(REPLACE(AR_Regions.ZipOrPostalCodes2, ' ,', ','), ', ', ',')) } =  AR_Customer.Zip)) "
	SQL=SQL & " OR (AR_Customer.[State] IN (SELECT StatesOrProvinces FROM AR_Regions))"
	
	SQL=SQL & " WHERE IC_Filters.InternalRecordIdentifier=FS_CustomerFilters.FilterIntRecID AND AR_Customer.CustNum=FS_CustomerFilters.CustID"
	
	IF LEN(regionID)>0 THEN
		IF regionID="0" THEN
			SQL=SQL & " AND (AR_Regions.Region IS NULL OR AR_Regions.InternalRecordIdentifier=0)"
			ELSE
				SQL=SQL & " AND AR_Regions.InternalRecordIdentifier=" & regionID
		END IF
	END IF
	IF LEN(searchValue)>0 THEN
		SQL=SQL & " AND (FS_CustomerFilters.CustID LIKE '%" & searchValue & "%' OR IC_Filters.FilterID LIKE '%" & searchValue & "%' OR AR_Customer.name LIKE '%" & searchValue & "%' OR IC_Filters.Description LIKE '%" & searchValue & "%')"
	END IF
	SELECT CASE orderValue
		CASE "0"
			SQL=SQL & " ORDER BY FS_CustomerFilters.CustID " & orderType
		CASE "1"
			SQL=SQL & " ORDER BY AR_Customer.name " & orderType
		CASE "2"
			SQL=SQL & " ORDER BY IsNULL(AR_Regions.Region,'Undefined') " & orderType
		CASE "3"
			SQL=SQL & " ORDER BY IC_Filters.FilterID  " & orderType
		CASE "4"
			SQL=SQL & " ORDER BY IC_Filters.Description  " & orderType
		CASE "5"
			SQL=SQL & " ORDER BY FS_CustomerFilters.notes  " & orderType
		CASE "6"
			SQL=SQL & " ORDER BY FS_CustomerFilters.FrequencyType  " & orderType
		CASE "7"
			SQL=SQL & " ORDER BY FS_CustomerFilters.FrequencyTime  " & orderType
		CASE "8"
			SQL=SQL & " ORDER BY FS_CustomerFilters.LastChangeDateTime  " & orderType
		CASE "9"
			
			SQL=SQL & " ORDER BY 14 " & orderType
		CASE ""
			
			SQL=SQL & " ORDER BY 15 " & orderType
		CASE "10"
			
			SQL=SQL & " ORDER BY FS_CustomerFilters.Qty " & orderType
		CASE "11"
			
			SQL=SQL & " ORDER BY FS_CustomerFilters.Price " & orderType
		CASE "12"
			
			SQL=SQL & " ORDER BY 16 " & orderType
		CASE "13"
			
			'SQL=SQL & " ORDER BY 17 " & orderType
	
			
	END SELECT
	
	Set cnnCustFilters = Server.CreateObject("ADODB.Connection")
	cnnCustFilters.open (Session("ClientCnnString"))
	
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3
	
	'Set rs = cnnCustFilters.Execute(SQL)
	rs.open SQL,cnnCustFilters,1
	IF PageSize=0 THEN
		rs.PageSize=rs.recordCount
		
		ELSE
			rs.PageSize = PageSize
	END IF
	nPageCount = rs.PageCount
	nRecordCount=rs.recordCount
	
	If Not rs.Eof Then
		rs.AbsolutePage = nPage
		Do While Not ( rs.Eof Or rs.AbsolutePage <> nPage )
			IF LEN(JSONdata)>0 THEN
				JSONdata=JSONdata & ","
			END IF
			JSONdata=JSONdata & "{"
			JSONdata=JSONdata & """id"":""" & rs("CustID") & """"
			JSONdata=JSONdata & ","
			JSONdata=JSONdata & """CustName"":""" & removeUnusualForJSON(rs("custName")) & """"
			JSONdata=JSONdata & ","
			'JSONdata=JSONdata & """CustRegion"":""" & removeUnusualForJSON(GetCustRegionByCustID(rs("CustID"))) & """"
			
			JSONdata=JSONdata & """CustRegion"":""" & removeUnusualForJSON(rs("region")) & """"
			JSONdata=JSONdata & ","
			JSONdata=JSONdata & """FilterID"":""" & rs("FilterID") & """"
			JSONdata=JSONdata & ","
			JSONdata=JSONdata & """FilterIntRecID"":""" & rs("InternalRecordIdentifier") & """"
			JSONdata=JSONdata & ","
			JSONdata=JSONdata & """filterDescription"":""" & removeUnusualForJSON(rs("filterDescription")) & """"
			JSONdata=JSONdata & ","
			JSONdata=JSONdata & """notes"":""" & removeUnusualForJSON(rs("notes")) & """"
			JSONdata=JSONdata & ","
			JSONdata=JSONdata & """FrequencyType"":""" & rs("FrequencyType") & """"
			JSONdata=JSONdata & ","	
			JSONdata=JSONdata & """FrequencyTime"":""" & rs("FrequencyTime") & """"
			JSONdata=JSONdata & ","	
			
			'EnrollmentDate Date
									
			JSONdata=JSONdata & """LastChangeDateTime"":""" & rs("LastChangeDateTime") & """"
			JSONdata=JSONdata & ","	
	

			JSONdata=JSONdata & """NextChangeDateTime"":""" & rs("NextChangeDateTime") & """"
			JSONdata=JSONdata & ","	
			
				
				

				
			
			JSONdata=JSONdata & """dayTill"":""" & rs("tillDays") & """"
			JSONdata=JSONdata & ","	
			
			JSONdata=JSONdata & """Qty"":""" & rs("Qty") & """"
			JSONdata=JSONdata & ","	
				
			JSONdata=JSONdata & """Price"":""" & FormatCurrency(rs("Price"),2) & """"
			JSONdata=JSONdata & ","	
			
			JSONdata=JSONdata & """TotalCost"":""" & FormatCurrency(rs("linetotal"),2) & """"
			JSONdata=JSONdata & ","	
			
			LCPGP = 0 
			IF Request.QueryString("maxID")="1" THEN
				TotalEquipmentValue = GetTotalValueOfEquipmentForCustomer(rs("CustID"))
				ELSE
					TotalEquipmentValue = 0
			END IF

			If TotalEquipmentValue <> 0 Then
				JSONdata=JSONdata & """CheckEquipment"":""<a data-toggle='modal' data-show='true' href='#' data-cust-id='" & rs("CustID") & "' data-lcp-gp='" & LCPGP & "' data-target='#modalEquipmentVPC' data-tooltip='true' data-title='View Customer Equipment'>" & FormatCurrency(TotalEquipmentValue,0) & "</a>"""    
			   
			   
					ELSE
						JSONdata=JSONdata & """CheckEquipment"":""No Equipment"""
						
					
			 End If 
			 JSONdata=JSONdata & ","	
			 
			ActiveFilterTicketNumber = ""
			Set rsActiveTicket = Server.CreateObject("ADODB.Recordset")
			rsActiveTicket.CursorLocation = 3
			
			SQActiveTicket = "SELECT TOP 1 ServiceTicketID FROM FS_ServiceMemosFilterInfo WHERE CustID = '" &  rs("CustID") & "' AND "
			SQActiveTicket = SQActiveTicket & " ServiceTicketID IN "
			SQActiveTicket = SQActiveTicket & "(SELECT MemoNumber FROM FS_ServiceMemos WHERE AccountNumber = '" & rs("CustID") & "' AND CurrentStatus='OPEN')"
			
			SQActiveTicket = SQActiveTicket & " AND CustFilterIntRecID = '" & rs("InternalRecordIdentifier") & "' "
						
			Set rsActiveTicket = cnnCustFilters.Execute(SQActiveTicket )
			If NOT rsActiveTicket.EOF Then ActiveFilterTicketNumber = rsActiveTicket("ServiceTicketID")
			 
			 JSONdata=JSONdata & """ActiveTicketNumber"":""" & ActiveFilterTicketNumber  & """"			 
			 JSONdata=JSONdata & ","
			 
			 
			 JSONdata=JSONdata & """Action"":"""""
			JSONdata=JSONdata & ","	
			JSONdata=JSONdata & """AdditionalInfo"":"""""
			JSONdata=JSONdata & ","	
			JSONdata=JSONdata & """AdditionalInfo2"":""" & rs("CustID") & """"
			
			JSONdata=JSONdata & "}"		
			
			rs.movenext
				
		Loop
		
			
End If
retData="{""orderby"":""" & orderValue & """,""draw"": " & CLng(Request.QueryString("draw")) & ",""recordsTotal"": " & nRecordCount & ",""recordsFiltered"": " & nRecordCount & ",""data"": [" & JSONdata & "],""byRegionData"":"+GetQtyCustByRegion()+"}"

  
Response.AddHeader "Content-Type", "application/json"
Response.Write retData

function removeUnusualForJSON(value)

removeUnusualForJSON=REPLACE(value,"""","&quot;")
	
END FUNCTION
%>

