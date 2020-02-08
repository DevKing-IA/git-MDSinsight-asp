<!--#include file="../../../inc/SubsAndFuncs.asp"-->
<!--#include file="../../../inc/mail.asp"-->
<!--#include file="../../../inc/InsightFuncs.asp"-->

<script type="text/javascript">
	function closeme() {
		window.open('', '_parent', '');
		window.close();  }
</script>
 
<%
'Designed to be launched via a scheduled process

'Usage = "http://{xxx}.{domain}.com/directLaunch/systemtools/gps/populateLatAndLong_launch.asp?runlevel=run_now
Server.ScriptTimeout = 250000

Dim EntryThread

'The runlevel parameter is inconsequential to the operation 
'of the page. It is only used so that the page will not run
'if it is loaded via an unexpected method (spiders, etc)

If Request.QueryString("runlevel") <> "run_now" then
	Response.Write("Improper usage, no run level was specified in the query string")	
	response.end
End IF 


'This single page loops through and handles alerts for ALL clients
SQL = "SELECT * FROM tblServerInfo WHERE Active = 1 and ClientKey <> '1999d'  and ClientKey <> '1999'"

Set TopConnection = Server.CreateObject("ADODB.Connection")
Set TopRecordset = Server.CreateObject("ADODB.Recordset")
TopConnection.Open InsightCnnString
	
'Open the recordset object executing the SQL statement and return records
TopRecordset.Open SQL,TopConnection,3,3


If Not TopRecordset.Eof Then

	Do While Not TopRecordset.EOF
	
		ClientKey = TopRecordset.Fields("clientkey")
	

		'The IF statement below makes sure that when run from DEV it only deos client keys with a d
		'and when run from LIVE it only does client keys without a d
		'Pretty smart, huh
		
		If (Instr(ucase(Request.ServerVariables("SERVER_NAME")),"DEV") = 0 AND Instr(ucase(ClientKey),"D") = 0)_
		or (Instr(ucase(Request.ServerVariables("SERVER_NAME")),"DEV") <> 0 AND Instr(ucase(ClientKey),"D") <> 0) Then

			Call SetClientCnnString
					
			Session("ClientCnnString") = MUV_READ("ClientCnnString") ' Until session vars are gone, then delete this


			'**************************************************************
			'Get next Entry Thread for use in the SC_AuditLogDLaunch table
			'**************************************************************
			On Error Goto 0
			Set cnnAuditLog = Server.CreateObject("ADODB.Connection")
			cnnAuditLog.open MUV_READ("ClientCnnString") 
			Set rsAuditLog = Server.CreateObject("ADODB.Recordset")
			rsAuditLog.CursorLocation = 3 
			Set rsAuditLog = cnnAuditLog.Execute("Select TOP 1 * from SC_AuditLogDLaunch order by EntryThread desc")
			If Not rsAuditLog.EOF Then 
				If IsNull(rsAuditLog("EntryThread")) Then EntryThread =1 Else EntryThread = rsAuditLog("EntryThread") + 1
			Else
				EntryThread = 1
			End If
			set rsAuditLog = nothing
			cnnAuditLog.close
			set cnnAuditLog = nothing
					

			WriteResponse "<font color='purple' size='24'>Start processing " & ClientKey  & "</font><br>"

			
			If MUV_READ("cnnStatus") = "OK" Then ' else it loops
			%>
					
					<!--#include file="populateGPSCache.asp"-->

					
			<%	
					Response.Write("~~~~~~~~~~~~ DONE Processing " & ClientKey  & " ~~~~~~~~~~~~<br>")
					
			End If						

									
			WriteResponse ("******** DONE Processing " & ClientKey  & "************<br>")
				
		End If
			
		TopRecordset.movenext
	
	Loop
	
	TopRecordset.Close
	Set TopRecordset = Nothing
	TopConnection.Close
	Set TopConnection = Nothing
	
End If

Response.write("<script type='text/javascript'>closeme();</script>")

'************************************************************************************
'************************************************************************************
'Subs and funcs begin here
'************************************************************************************

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
		Session("SQL_Owner") = Recordset.Fields("dbLogin")
		dummy = MUV_Write("ClientID",Recordset.Fields("clientkey"))
		Recordset.close
		Connection.close
		dummy=MUV_WRITE("cnnStatus","OK")
	End If
End Sub


Sub WriteResponse(passedLogEntry)

	response.write(Now() & "&nbsp;&nbsp;&nbsp;" & passedLogEntry)
	
	passedLogEntry = Replace(passedLogEntry,"'","''")
	
	SQL = "INSERT INTO SC_AuditLogDLaunch (EntryThread, DirectLaunchName, DirectLaunchFile, LogEntry)"
	SQL = SQL &  " VALUES (" & EntryThread & ""
	SQL = SQL & ",'MCS Rebuild Helper'"
	SQL = SQL & ",'/directlaunch/bizintel/mcs_rebuild_helper_launch.asp'"
	SQL = SQL & ",'"  & passedLogEntry & "'"
	SQL = SQL & ")"
	
	'Response.write("<BR>" & SQL & "<BR>")
	
	Set cnnAuditLog = Server.CreateObject("ADODB.Connection")
	cnnAuditLog.open Session("ClientCnnString") 
	Set rsAuditLog = Server.CreateObject("ADODB.Recordset")
	rsAuditLog.CursorLocation = 3 
	
	Set rsAuditLog = cnnAuditLog.Execute(SQL)

	set rsAuditLog = nothing
	cnnAuditLog.close
	set cnnAuditLog = nothing

End Sub

Sub DelayResponse(numberOfseconds)
 Dim WshShell
 Set WshShell=Server.CreateObject("WScript.Shell")
 WshShell.Run "waitfor /T " & numberOfSecond & "SignalThatWontHappen", , True
End Sub

Function CheckRemoteURL(fileURL)
    ON ERROR RESUME NEXT
    Dim xmlhttp

    Set xmlhttp = Server.CreateObject("MSXML2.ServerXMLHTTP")

    xmlhttp.open "GET", fileURL, False
    xmlhttp.send
    If(Err.Number<>0) then
        Response.Write "Could not connect to remote server"
    else
        Select Case Cint(xmlhttp.status)
            Case 200, 202, 302
                Set xmlhttp = Nothing
                CheckRemoteURL = True
            Case Else
                Set xmlhttp = Nothing
                CheckRemoteURL = False
        End Select
    end if
    ON ERROR GOTO 0
End Function

Sub ArrayCull(ByRef arr)
  Dim i, dict
  If IsArray(arr) Then
    Set dict = CreateObject("Scripting.Dictionary")
    For i = 0 To UBound(arr)
      If Not dict.Exists(arr(i)) Then
        Call dict.Add(arr(i), arr(i))
      End If
    Next
    arr = dict.Items
  End If
End Sub


Function stateCleanUp(state)
   ' Fix common abbreviations
   tmpState = UCase(Trim(state))
   If  Right(tmpState,1) = "." Then tmpState = Left(tmpState,Len(tmpState) - 1)
   If Len(tmpState) = 2 Then
      stateCleanUp = tmpState
   Else
      Select Case tmpState
         ' United States
         Case "ALABAMA","ALA"
             stateCleanUp = "AL"
         Case "ALASKA"
             stateCleanUp = "AK"
         Case "ARIZONA","ARIZ"
            stateCleanUp = "AZ"
         Case "ARKANSAS","ARK"
            stateCleanUp = "AR"
         Case "CALIFORNIA","CAL","CALIF"
            stateCleanUp = "CA"
         Case "COLORADO","COLO","COL"
            stateCleanUp = "CO"
         Case "CONNECTICUT","CONN"
            stateCleanUp = "CT"
         Case "DELAWARE","DEL"
            stateCleanUp = "DE"
         Case "D.C.", "DISTRICT OF COLUMBIA"
            stateCleanUp = "DC"
         Case "FLORIDA","FLA"
            stateCleanUp = "FL"
         Case "GEORGIA"
            stateCleanUp = "GA"
         Case "HAWAII"
            stateCleanUp = "HI"
         Case "IDAHO","IDA"
            stateCleanUp = "ID"
         Case "ILLINOIS","ILL"
            stateCleanUp = "IL"
         Case "INDIANA","IND"
            stateCleanUp = "IN"
         Case "IOWA"
            stateCleanUp = "IA"
         Case "KANSAS","KAN"
            stateCleanUp = "KS"
         Case "KENTUCKY"
            stateCleanUp = "KY"
         Case "LOUISIANA"
            stateCleanUp = "LA"
         Case "MAINE"
            stateCleanUp = "ME"
         Case "MARYLAND"
            stateCleanUp = "MD"
         Case "MASSACHUSETTS","MASS"
            stateCleanUp = "MA"
         Case "MICHIGAN","MICH"
            stateCleanUp="MI"
         Case "MINNESOTA","MINN"
            stateCleanUp = "MN"
         Case "MISSISSIPPI", "MISS"
            stateCleanUp = "MS"
         Case "MISSOURI"
            stateCleanUp = "MO"
         Case "MONTANA","MONT"
            stateCleanUp = "MT"
         Case "NEBRASKA","NEBR","NEB"
            stateCleanUp = "NE"
         Case "NEVADA","NEV"
            stateCleanUp = "NV"
         Case "NEW HAMPSHIRE","N.H."
            stateCleanUp = "NH"
         Case "N.J.", "NEW JERSEY", "N.JERSEY", "N. JERSEY"
            stateCleanUp = "NJ"
         Case "NEW MEXICO","N.M.","N.MEX","N. MEX"
            stateCleanUp = "NM"
         Case "N.Y.", "NEW YORK","N. YORK"
            stateCleanUp = "NY"
         Case "NORTH CAROLINA","N.C."
            stateCleanUp = "NC"
         Case "NORTH DAKOTA","N.D."
            stateCleanUp = "ND"
         Case "OHIO"
            stateCleanUp = "OH"
         Case "OKLAHOMA","OKLA"
            stateCleanUp = "OK"
         Case "OREGON","ORE"
            stateCleanUp = "OR"
         Case "PENNSYLVANIA", "PENN"
            stateCleanUp = "PA"
         Case "PUERTO RICO","P.R."
            stateCleanUp = "PR"
         Case "RHODE ISLAND","R.I."
            stateCleanUp = "RI"
         Case "SOUTH CAROLINA","S.C."
            stateCleanUp = "SC"
         Case "SOUTH DAKOTA","S.D."
            stateCleanUp = "SD"
         Case "TENNESSEE","TENN"
            stateCleanUp = "TN"
         Case "TEX", "TEXAS"
            stateCleanUp = "TX"
         Case "UTAH"
            stateCleanUp = "UT"
         Case "VERMONT"
            stateCleanUp = "VT"
         Case "VIRGINIA"
            stateCleanUp = "VA"
         Case "WASHINGTON","WASH"
            stateCleanUp = "WA"
         Case "WEST VIRGINIA","W.V."
            stateCleanUp = "WV"
         Case "WISCONSIN","WIS","WISC"
            stateCleanUp = "WI"
         Case "WYOMING","WYO"
            stateCleanUp = "WY"

         ' Canada
         Case "ONT", "ONT.", "ONTARIO"
            stateCleanUp = "ON"
         Case "B.C.", "BC.", "BRITISH COLUMBIA"
            stateCleanUp = "BC"
         Case "ALBERTA"
            stateCleanUp = "AB"
         Case "QUEBEC"
            stateCleanUp = "QC"
         Case "MANITOBA"
            stateCleanUp = "MB"
         Case "SASKATCHEWAN"
            stateCleanUp = "SK"
         Case "NOVA SCOTIA"
            stateCleanUp = "NS"
         Case "NEW BRUNSWICK"
            stateCleanUp = "NB"
         Case "YUKON", "YUKON TERRITORY"
            stateCleanUp = "YT"
         Case "NUNAVUT"
            stateCleanUp = "NU"
         Case "NEWFOUNDLAND AND LABRADOR", "NEWFOUNDLAND"
            stateCleanUp = "NL"
         Case "PRINCE EDWARD ISLAND", "PRINCE EDWARD"
            stateCleanUp = "PE"
         Case "NORTHWEST TERRITORIES"
            stateCleanUp = "NT"
         Case Else
            stateCleanUp = Trim(state)
      End Select
   End If
End Function

Function isCanadian(state)
   Select Case stateCleanUp(state)
      Case "ON", "QC", "NS", "NB", "MB", "BC", "SK", "AB", "NL", "NT", "YT", "PE", "NU"
         isCanadian = True
      Case Else
         isCanadian = False
   End Select
End Function

Function censusGPSLookup(streetAddress, csz)
   Dim xmlhttp, url, censusUrl, censusParams, latitude, longitude, myJson

   longitude = ""
   latitude = ""
   censusUrl = "https://geocoding.geo.census.gov/geocoder/locations/onelineaddress?address="
   censusParams = "&benchmark=Public_AR_Current&format=json&returntype=locations"
   url = censusUrl & Replace(Trim(streetAddress), " ", "+") & ",+" & Replace(csz, " ", "+") & censusParams

	'Response.Write url & "<BR>"

   set xmlhttp = CreateObject("WinHTTP.WinHTTPRequest.5.1")
   xmlhttp.open "GET",url,false
   xmlhttp.setRequestHeader "User-Agent", "Mozilla/4.0"
   xmlhttp.setRequestHeader "Accept", "application/json"
   xmlhttp.send
   If CInt(xmlhttp.Status) = 200 Then

	Response.Write xmlhttp.responseText & "<BR>"

      set myJson = JSON.parse(xmlhttp.responseText)
      set xmlhttp = Nothing
      On Error Resume Next
      longitude = myJson.result.addressMatches.get(0).coordinates.x
      If Err.Number <> 0 Then
         On Error Goto 0
         longitude = ""
      Else
         On Error Goto 0
         latitude = myJson.result.addressMatches.get(0).coordinates.y
      End If
      set myJson = Nothing
   End If

   censusGPSLookup = latitude & ":" & longitude
End Function

Function bingGPSLookup(streetAddress, csz, canadianAddress)
   Dim xmlhttp, url, bingUrl, bingKey, latitude, longitude, myJson
   longitude = ""
   latitude = ""
   bingUrl = "http://dev.virtualearth.net/REST/v1/Locations?CountryRegionIso2="
   bingKey = "&o=json&userIp=127.0.0.1&key=Au_r8O94UF0tqnu2nEOPCj4IKfwDio6Jl3IcLKwEPVrUfusmCGGQ7vbNpCDWxD2q"
   set xmlhttp = CreateObject("WinHTTP.WinHTTPRequest.5.1")
   if canadianAddress Then
      url = bingUrl & "CA"
   Else
      url = bingUrl & "US"
   End If
   url = url & "&q=" & Replace(Trim(streetAddress)," ","%20") & ",%20" & Replace(csz," ","%20") & bingKey

	Response.Write url & "<BR>"

   xmlhttp.open "GET",url,false
   xmlhttp.setRequestHeader "User-Agent", "Mozilla/4.0"
   xmlhttp.setRequestHeader "Accept", "application/json"
   xmlhttp.send
   If CInt(xmlhttp.Status) = 200 Then

	Response.Write xmlhttp.responseText & "<BR>"

      set myJson = JSON.parse(xmlhttp.responseText)
      set xmlhttp = Nothing
      If myJson.resourceSets.length >= 1 and CInt(myJson.resourceSets.get(0).estimatedTotal) >= 1 Then
         Dim confidence, matchCode, goodMatchCodeFound
         confidence = myJson.resourceSets.get(0).resources.get(0).confidence
         goodMatchCodeFound = False
         For Each matchCode In myJson.resourceSets.get(0).resources.get(0).matchCodes
            If matchCode = "Good" Then goodMatchCodeFound = True
         Next
         If confidence = "High" and goodMatchCodeFound Then
            Dim geoPoint, routingUsage, usageType, tmpLat, tmpLon
            For Each geoPoint In myJson.resourceSets.get(0).resources.get(0).geocodePoints
               If geoPoint.hasOwnProperty("coordinates") Then
                  tmpLat = geoPoint.coordinates.get(0)
                  tmpLon = geoPoint.coordinates.get(1)
                  routingUsage = False
                  For Each usage In geoPoint.usageTypes
                     If usage = "Route" Then routingUsage = True
                  Next
                  If routingUsage Then
                     latitude = geoPoint.coordinates.get(0)
                     longitude = geoPoint.coordinates.get(1)
                     Exit For
                  End If
               End If
            Next
            If Len(latitude) = 0 Then
               ' This handles the case when only result is a "Display" usage Type
               latitude = tmpLat
               longitude = tmpLon
            End If
         End If
      End If
   End If

   bingGPSLookup = latitude & ":" & longitude
End Function

Function lookupGPS(passedstreetAddress, passedcity, passedstate, passedzip,clientKey)

	Dim resultlookupGPS, OKToProceed, csv
	resultlookupGPS = ""
	OKToProceed = True
		
   ' make sure function arguments have sane and easy to use values
   ' for successful GPS lookup we must have a street address and one or more of passedcity, passedstate, and passedzip

   If isEmpty(passedstreetAddress) or isNull(passedstreetAddress) or len(passedstreetAddress) = 0 Then
      OKToProceed = False
   End If
   
   If OKToProceed = True Then
   
	   If isEmpty(passedcity) or len(passedcity) = 0 Then passedcity = Null
	   If isEmpty(passedstate) or len(passedstate) = 0 Then passedstate = Null
	   If isEmpty(passedzip) or len(passedzip) = 0 Then passedzip = Null
	
	   If Not isNull(passedcity) Then csz = trim(passedcity)
	
	   If Not isNull(passedstate) Then
	      If isEmpty(csz) Then csz = trim(passedstate) Else csz = csz & ", " & trim(passedstate)
	   End If
	
	   If Not isNull(passedzip) Then
	      If isEmpty(csz) Then csz = trim(passedzip) else csz = csz & " " & trim(passedzip)
	   End If

	   If isEmpty(csz) Then
	      OKToProceed = False
	   End If

	End If
	
   If OKToProceed = True Then
	
	   ' check for cached lat/lon
	   SQLLookupGPS = "SELECT latitude, longitude FROM mdsinsight.GPS_CACHE WHERE streetAddress='" & Replace(passedstreetAddress, "'", "''") & "' AND city"
	   
	   If isNull(passedcity) Then
	      SQLLookupGPS = SQLLookupGPS & " IS NULL "
	   Else
	      SQLLookupGPS = SQLLookupGPS & "='" & Replace(passedcity, "'", "''") & "'"
	   End If
	   
	   SQLLookupGPS = SQLLookupGPS & " AND state"
	   
	   If isNull(passedstate) Then
	      SQLLookupGPS = SQLLookupGPS & " IS NULL "
	   Else
	      SQLLookupGPS = SQLLookupGPS & "='" & passedstate & "'"
	   End If
	   
	   SQLLookupGPS = SQLLookupGPS & " AND zip"
	   
	   If isNull(passedzip) Then
	      SQLLookupGPS = SQLLookupGPS & " IS NULL"
	   Else
	      SQLLookupGPS = SQLLookupGPS & "='" & passedzip & "'"
	   End If
	   
		Response.Write ("<font color='green'>" & ClientKey  & "    " & Now() &  "</font>" & "Query: " & SQLLookupGPS & "<BR>")

		Set rslookupGPS = Server.CreateObject("ADODB.Recordset")
		rslookupGPS.CursorLocation = 3

		rslookupGPS.Open SQLLookupGPS ,InsightCnnString,0,3
		
		If NOT rslookupGPS.EOF Then
			resultlookupGPS = rslookupGPS("latitude") & ":" & rslookupGPS("longitude")
			OKToProceed = False
		End If
			   
	   rslookupGPS.Close
	   Set rslookupGPS = Nothing
	   
	End If

   If OKToProceed = True Then

	   Dim lat, lon, latlon
	   if (isNull(state) and (Left(clientKey,4) = "1128" or Left(clientKey,4) = "1190")) or isCanadian(state) Then
	      latlon = bingGPSLookup(passedstreetAddress,csz,True)
	   Else
	      latlon = censusGPSLookup(passedstreetAddress,csz)
	      If Len(latlon) = 1 Then
	         latlon = bingGPSLookup(passedstreetAddress,csz,False)
	      End If
	   End If
	   If Len(latlon) = 1 Then
	      lat = Null
	      lon = Null
	   Else
	      latlonSplit = Split(latlon,":")
	      lat = latlonSplit(0)
	      lon = latlonSplit(1)
	      resultlookupGPS = lat & ":" & lon
	   End If

		Set rslookupGPS = Server.CreateObject("ADODB.Recordset")
		rslookupGPS.CursorLocation = 3

	   ' update cache
	   SQLLookupGPS = "INSERT INTO GPS_CACHE (streetAddress, city, state, zip, latitude, longitude) VALUES ("
	   SQLLookupGPS = SQLLookupGPS & "'" & Replace(passedstreetAddress, "'", "''") & "',"
	   If isNull(passedcity) Then SQLLookupGPS = SQLLookupGPS & "NULL," Else SQLLookupGPS = SQLLookupGPS & "'" & Replace(passedcity, "'", "''") & "',"
	   If isNull(passedstate) Then SQLLookupGPS = SQLLookupGPS & "NULL," Else SQLLookupGPS = SQLLookupGPS & "'" & passedstate & "',"
	   If isNull(passedzip) Then SQLLookupGPS = SQLLookupGPS & "NULL," Else SQLLookupGPS = SQLLookupGPS & "'" & passedzip & "',"
	   if isNull(lat) Then
	      SQLLookupGPS = SQLLookupGPS & "NULL,NULL)"
	   Else
	      SQLLookupGPS = SQLLookupGPS & "'" & lat & "',"
	      SQLLookupGPS = SQLLookupGPS & "'" & lon & "')"
	   End If
	   
		Response.Write "Insert SQL: " & SQLLookupGPS & "<BR>"

		Set cnnInsight = Server.createObject("ADODB.Connection")
		cnnInsight.open InsightCnnString

		Set rslookupGPS = Server.CreateObject("ADODB.Recordset")
		rslookupGPS.CursorLocation = 3
		
		Set rslookupGPS = cnnInsight.Execute(SQLLookupGPS)

		Set rslookupGPS = Nothing
		cnnInsight.Close
		Set cnnInsight = Nothing
		
	End IF	   

   lookupGPS = resultlookupGPS 
	   
End Function


'************************************************************************************
'************************************************************************************
'Subs and funcs end here
'************************************************************************************

%>