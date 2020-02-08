<!--#include file="../../inc/subsandfuncs.asp"-->
<!--#include file="../../inc/InSightFuncs.asp"-->
<%

Region = Request.Form("txtRegion")
Cities1 = Request.Form("txtCities")
ZipPostalCodes = Request.Form("txtZipPostalCodes")
StatesProvinces = Request.Form("txtStatesProvinces")
CatchAllRegionIntRecIDs = Request.Form("lstSelectedRegionList")

If Request.Form("chkAutomaticFilter") = "on" then AutomaticFilter = 1 Else AutomaticFilter = 0
If Request.Form("chkSuggestedFilter") = "on" then SuggestedFilter = 1 Else SuggestedFilter = 0
If Request.Form("chkUseRegionForServiceTickets") = "on" then UseRegionForServiceTickets = 1 Else UseRegionForServiceTickets = 0

AutoTicket = Request.Form("selAutoTicket")
SuggestedTicket = Request.Form("selSuggestedTicket")
StateForCities = Request.Form("selStateOrProvince")

Cities = split(Cities1,",")
CitiesCount = UBound(Cities) + 1
CitiesLen = len(Cities1)
ZipLen = len(ZipPostalCodes)



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



'ZipTest3 = mid(ZipPostalCodes,secondPos+1,25)
'thirdPos = InStrRev(ZipTest3, ",", -1, vbTextCompare) 
'ZipArray3 = mid(ZipPostalCodes,firstPos+secondPos+1,thirdPos)


'Response.Write("<br>" & Cities1 & "<br>")
'Response.Write("<br>" & CitiesCount & ":" & CitiesLen & "<br>")
'Response.Write("<br>" & firstPosition & "<br>")
'Response.Write("<br>" & CitiesArray1 & "<br>")
'Response.Write("<br>" & secondPosition & "<br>")
'Response.Write("<br>" & CitiesArray2 & "<br>")
'Response.Write("<br>" & thirdPosition & "<br>")
'Response.Write("<br>" & CitiesArray3 & "<br>")
'Response.Write("<br>" & ZipArray1 & "<br>")
'Response.Write("<br>" & ZipArray2 & "<br>")

'If DisplayOrder = "" Then
'	DisplayOrder = 0
'End If	

SQL = "INSERT INTO AR_Regions (Region, Cities1, Cities2, Cities3 , StatesOrProvinces, ZipOrPostalCodes1, ZipOrPostalCodes2, IncludeInAutoFilterTickets, IncludeInSuggestedFilterTickets, AutoFilterChangeMaxNumTicketsPerDay, SuggestedFilterChangeMaxNumTicketsPerDay,StateForCities,CatchAllRegionIntRecIDs,UseRegionForServiceTickets)"
SQL = SQL &  " VALUES (" 
SQL = SQL & "'"  & Region & "', '"  & CitiesArray1 & "', '"  & CitiesArray2 & "', '"  & CitiesArray3 & "', '"  & StatesProvinces & "', '"  & ZipArray1 & "', '"  & ZipArray2 & "'," & AutomaticFilter & "," & SuggestedFilter & "," & AutoTicket & "," & SuggestedTicket & ",'" & StateForCities  & "','" & CatchAllRegionIntRecIDs & "'," & UseRegionForServiceTickets & ")"


Response.Write("<br>" & SQL & "<br>")

Set cnnregions = Server.CreateObject("ADODB.Connection")
cnnregions.open (Session("ClientCnnString"))

Set cnnparts = Server.CreateObject("ADODB.Recordset")
cnnparts.CursorLocation = 3 

Set cnnparts = cnnregions.Execute(SQL)
set cnnparts = Nothing

Description = GetUserDisplayNameByUserNo(Session("Userno")) & " added the region: " & Region 
CreateAuditLogEntry "Accounts Recivable module" & " region added","Accounts Recivable module region added","Minor",0,Description

Response.Redirect("main.asp")

%>