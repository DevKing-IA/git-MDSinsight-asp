<!--#include file="../../inc/subsandfuncs.asp"-->
<!--#include file="../../inc/InSightFuncs.asp"-->

<%
InternalRecordIdentifier = Request.Form("txtInternalRecordIdentifier")

'*******************************************************************
'Lookup the record as it exists now so we can fillin the audit trail

SQL = "SELECT * FROM AR_Regions where InternalRecordIdentifier = " & InternalRecordIdentifier
	
Set cnnregions = Server.CreateObject("ADODB.Connection")
cnnregions.open (Session("ClientCnnString"))
Set rsregions = Server.CreateObject("ADODB.Recordset")
rsregions.CursorLocation = 3 
Set rsregions = cnnregions.Execute(SQL)
	
If not rsregions.EOF Then
	Orig_Region = rsregions("Region")
	'Orig_PartDescription = rsregions("PartDescription")	
End If

'***********************************************************************
'End Lookup the record as it exists now so we can fillin the audit trail
'***********************************************************************

'PartNumber = Request.Form("txtPartNumber")
'PartDescription = Request.Form("txtPartDescription")
'DisplayOrder = Request.Form("txtPartDisplayOrder")

Region = Request.Form("txtRegion")
Cities1 = Request.Form("txtCities")
ZipPostalCodes = Request.Form("txtZipPostalCodes")
StatesProvinces = Request.Form("txtStatesProvinces")

If Request.Form("chkAutomaticFilter") = "on" then AutomaticFilter = 1 Else AutomaticFilter = 0
If Request.Form("chkSuggestedFilter") = "on" then SuggestedFilter = 1 Else SuggestedFilter = 0
AutoTicket = Request.Form("selAutoTicket")
SuggestedTicket = Request.Form("selSuggestedTicket")
StateForCities = Request.Form("selStateOrProvince")
CatchAllRegionIntRecIDs = Request.Form("lstSelectedRegionList")

Cities = split(Cities1,",")
CitiesCount = UBound(Cities) + 1
CitiesLen = len(Cities1)
ZipLen = len(ZipPostalCodes)

'Response.Write CitiesLen

if CitiesLen > 8000 Then
	CitiesTest1 = mid(Cities1,1,8000)
	firstPosition = InStrRev(CitiesTest1, ",", -1, vbTextCompare) 
	CitiesArray1 = mid(Cities1,1,firstPosition-1)
Else
	CitiesArray1 = Cities1
End if	

if CitiesLen > 8000 Then
	CitiesTest2 = mid(Cities1,firstPosition+1,8000)
	secondPosition = InStrRev(CitiesTest2, ",", -1, vbTextCompare) 
	'Response.Write secondPosition
	If secondPosition > 0 Then
		CitiesArray2 = mid(Cities1,firstPosition+1,secondPosition-1)
	else
		CitiesArray2 = mid(Cities1,firstPosition+1,CitiesLen)
	End IF	
End if

if CitiesLen > 16000 Then
	CitiesTest3 = mid(Cities1,secondPosition+1,8000)
	thirdPosition = InStrRev(CitiesTest3, ",", -1, vbTextCompare) 
	CitiesArray3 = mid(Cities1,firstPosition+secondPosition+1,thirdPosition-1)
End If

if ZipLen > 8000 Then
	ZipTest1 = mid(ZipPostalCodes,1,8000)
	firstPos = InStrRev(ZipTest1, ",", -1, vbTextCompare) 
	ZipArray1 = mid(ZipPostalCodes,1,firstPos-1)
Else
	ZipArray1 = ZipPostalCodes
End If	

if ZipLen > 8000 Then
	ZipTest2 = mid(ZipPostalCodes,firstPos+1,8000)
	secondPos = InStrRev(ZipTest2, ",", -1, vbTextCompare) 
	ZipArray2 = mid(ZipPostalCodes,firstPos+1,secondPos-1)
End If

If Request.Form("chkUseRegionForServiceTickets") = "on" then UseRegionForServiceTickets = 1 Else UseRegionForServiceTickets = 0

SQL = "UPDATE AR_Regions SET "
SQL = SQL &  "Region = '" & Region & "' "
SQL = SQL &  ", Cities1 = '" & CitiesArray1 & "' "
SQL = SQL &  ", Cities2 = '" & CitiesArray2 & "' "
SQL = SQL &  ", Cities3 = '" & CitiesArray3 & "' "
SQL = SQL &  ", StatesOrProvinces = '" & StatesProvinces & "' "
SQL = SQL &  ", ZipOrPostalCodes1 = '" & ZipArray1 & "' "
SQL = SQL &  ", ZipOrPostalCodes2 = '" & ZipArray2 & "' "
SQL = SQL &  ", IncludeInAutoFilterTickets = '" & AutomaticFilter & "' "
SQL = SQL &  ", IncludeInSuggestedFilterTickets = '" & SuggestedFilter & "' "
SQL = SQL &  ", AutoFilterChangeMaxNumTicketsPerDay = '" & AutoTicket & "' "
SQL = SQL &  ", SuggestedFilterChangeMaxNumTicketsPerDay = '" & SuggestedTicket & "' "
SQL = SQL &  ", StateForCities = '" & StateForCities & "' "
SQL = SQL &  ", CatchAllRegionIntRecIDs = '" & CatchAllRegionIntRecIDs & "' "
SQL = SQL &  ", UseRegionForServiceTickets = " & UseRegionForServiceTickets & " "
SQL = SQL &  " WHERE InternalRecordIdentifier = " & InternalRecordIdentifier

Response.Write("<br>" & SQL & "<br>")

Set rsregions = cnnregions.Execute(SQL)
set rsregions = Nothing


Description = ""
If Orig_Region  <> Region  Then
	Description = Description & "AccountsReceivable module region changed from " & Orig_Region  & " to " & Region  
End If

CreateAuditLogEntry "AccountsReceivable module region edited","AccountsReceivable module region edited","Minor",0,Description

Response.Redirect("main.asp")

%>















