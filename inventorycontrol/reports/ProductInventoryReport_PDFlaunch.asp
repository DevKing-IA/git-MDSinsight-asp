<!--#include file="../../inc/header.asp"-->
<style>
#PleaseWaitPanel{
position: fixed;
left: 470px;
top: 275px;
width: 975px;
height: 300px;
z-index: 9999;
background-color: #fff;
opacity:1.0;
text-align:center;
}    
</style>

<div id="PleaseWaitPanel">
	<br><br>Processing, please wait...<br><br>
	<img src="../../img/loading.gif"/>
</div>

<script type="text/javascript">
$(document).ready(function() {
    $("#PleaseWaitPanel").hide();
});
</script>

<script type="text/javascript">
function HideIt()
{
	$("#PleaseWaitPanel").hide();
}
</script>

<%
'baseURL should always have a trailing /slash, just in case, handle either way
If right(baseURL,1)="/" Then maildomain = Left(right(baseURL,len(baseURL)-7),len(right(baseURL,len(baseURL)-7))-1) Else maildomain = right(baseURL,len(baseURL)-7)
sURL = Request.ServerVariables("SERVER_NAME")

DebugMessages = False ' Set to true to turn om Response.Writes

'Generate a unique number to be used for all pdfs throughout this page
Randomize
UniqueNum = int((9999999-1111111+1)*rnd+1111111)


'************************
'Read Settings_Reports
'************************
SQL = "SELECT * FROM Settings_Reports WHERE ReportNumber = 1600 AND UserNo = " & Session("userNo")
Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rs = Server.CreateObject("ADODB.Recordset")
Set rs= cnn8.Execute(SQL)
If NOT rs.EOF Then
	UnitUPCData = rs("ReportSpecificData1")
	CaseUPCData = rs("ReportSpecificData2")
	InventoriedItem = rs("ReportSpecificData3")
	PickableItem = rs("ReportSpecificData4")
	ProductCategoriesForInventoryReport = rs("ReportSpecificData5")
	If IsNull(UnitUPCData) Then UnitUPCData = ""
	If IsNull(CaseUPCData) Then CaseUPCData  = ""
	If IsNull(InventoriedItem) Then InventoriedItem = ""
	If IsNull(PickableItem) Then PickableItem = ""
	If IsNull(ProductCategoriesForInventoryReport) Then ProductCategoriesForInventoryReport = ""
Else
	UnitUPCData = ""
	CaseUPCData = ""
	InventoriedItem = ""
	PickableItem = ""
	ProductCategoriesForInventoryReport = ""
End If										
'****************************
'End Read Settings_Reports
'****************************

Set Pdf = Server.CreateObject("Persits.Pdf")
Set Doc = Pdf.CreateDocument

ImpVar = baseURL & "inventorycontrol/reports/ProductInventoryReport_PDFgen.asp?"
ImpVar = ImpVar & "cl=" & MUV_Read("ClientID")
ImpVar = ImpVar & "&u=" & MUV_Read("SQL_Owner")
ImpVar = ImpVar & "&un=" & Session("UserNo")
ImpVar = ImpVar & "&uupc=" & UnitUPCData
ImpVar = ImpVar & "&cupc=" & CaseUPCData
ImpVar = ImpVar & "&i=" & InventoriedItem
ImpVar = ImpVar & "&p=" & PickableItem
ImpVar = ImpVar & "&c=" & ProductCategoriesForInventoryReport


If DebugMessages = True Then Response.Write("<br><br><br><br>" & ImpVar & "<br>")

Doc.ImportFromUrl ImpVar, "scale=0.75; hyperlinks=false; drawbackground=true; landscape=true"

 
fn = "\clientfiles\" & Left(MUV_Read("ClientID"),4) & "\z_pdfs\ProdInvReport" & Trim(UniqueNum) & "_Main.pdf"
fn = Replace(fn,"/","-")
fn = Replace(fn,":","-")
'response.write(fn & "<br>")
fn2 = Left(baseURL,Len(baseURL)-1) & fn
fn2 = Replace(fn2,"\","/")
'response.write(fn2 & "<br>")
Main_PDF_Filename = fn
If DebugMessages = True Then response.write("Main_PDF_Filename:" & Main_PDF_Filename & "<br>")
Filename = Doc.Save(Server.MapPath(fn), False)



'Now wait until the file exists on the server before we try to mail it
TimeoutSecs = 60
TimeoutCounter=0
FOundFile = False
Do While TimeoutCounter < TimeoutSecs 
	If CheckRemoteURL(fn2) = True Then
		FoundFile = True
		Exit Do ' The file is there
	End If
	DelayResponse(1) ' wait 1 sec & try again
	TimeoutCounter = TimeoutCounter + 1
Loop

If FoundFile <> True Then 
	Response.Write ("NO FILE FOUND")
	Response.End ' Could not fine the pdf, so just bail
End If


'Now open the PDF in a new window

Response.Write("<SCRIPT language='javascript'>window.open('" & fn2 & "');</SCRIPT>")

Response.Write("<script language=javascript>HideIt();</script>")
%>	

<br><br><br>
<a href="ProductInventoryReport.asp">
	<button type="button" class="btn btn-default">&lsaquo; Back To Product UPC Report</button>
</a>
<p><br>*If the Product UPC Report PDF has not opened, please check your popup blocker.</p>
<%
'*******************************************************************************************************************************************************************
'*******************************
' SUBs and FUNCTIONs Start Here
'*******************************
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

%>