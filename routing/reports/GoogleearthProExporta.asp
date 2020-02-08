
<%
	'On error resume next
   Session.Timeout=60
   Server.ScriptTimeout=60


optioncolors = split("#FFFFFF,#87D8C7,#D0C76F,#D9BCB5,#E24932,#D13EA2,#DD754C,#3FE337,#8F6FDA,#BD4C73,#509FD3,#796230,#FFFF00,#FF00FF,#00ffff,#66819D,#2C6FD3,#ff0000,#00FF00,#4444FF,#495CB9,#B47C34.#36AF9C.#E17A75,#AFC5A2,#726E69,#8B4BDF,#4EA16B,#ffb6c1,#FFA500", ",")

Set conn = Server.CreateObject("ADODB.Connection")
conn.open (Session("ClientCnnString"))
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3 

Set rs1 = Server.CreateObject("ADODB.Recordset")
rs1.CursorLocation = 3 

'SQL = "UPDATE AR_Customer set Latitude=null,Longitude=null"
'conn.execute(SQL)

DateToRetrieve = Request.QueryString("date")
If DateToRetrieve = "" Then
	checked1 = "checked"
	checked2 = ""
else
	checked2 = "checked"
	checked1 = ""
end if


SQL = "SELECT DISTINCT RT_DeliveryBoardHistory.DeliveryDate FROM RT_DeliveryBoardHistory ORDER BY RT_DeliveryBoardHistory.DeliveryDate"
	Set rs = conn.Execute(SQL)
	i = 0
	'olddate = ""
	disabledates = ""
	do while not rs.Eof
		if i = 0 then
			startdate = Right("0" & month(rs("DeliveryDate")), 2) & "/" & Right("0" & Day(rs("DeliveryDate")), 2)  & "/" & Year(rs("DeliveryDate"))
		end if
		k = 0
		newdate = olddate + 1
		do while ((olddate <> "") and (cdate(rs("DeliveryDate")) > newdate)) 
			k = k + 1
			if disabledates = "" then
				disabledates = """" & Right("0" & month(newdate), 2) & "/" & Right("0" & Day(newdate), 2)  & "/" & Year(newdate) & """"
			else 
				disabledates = disabledates & "," & """" & Right("0" & month(newdate), 2) & "/" & Right("0" & Day(newdate), 2)  & "/" & Year(newdate) & """" 
			end if
			olddate = newdate
			newdate = olddate + 1
			if k > 20 then
				exit do 
			end if
		loop
		i = i + 1
		enddate = Right("0" & month(rs("DeliveryDate")), 2) & "/" & Right("0" & Day(rs("DeliveryDate")), 2)  & "/" & Year(rs("DeliveryDate"))
		olddate = cdate(rs("DeliveryDate"))
		rs.movenext
	loop
	rs.close

%> 
<link href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css" rel="stylesheet">
<script src="http://code.jquery.com/jquery-1.11.2.min.js"></script>
<link href="http://ajax.googleapis.com/ajax/libs/jqueryui/1.9.2/themes/ui-darkness/jquery-ui.css" rel="stylesheet">
<script src="http://ajax.googleapis.com/ajax/libs/jqueryui/1.9.2/jquery-ui.min.js"></script>
<script src="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/js/bootstrap.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/jquery.inputmask/3.1.62/jquery.inputmask.bundle.js"></script>
<style  type="text/css">
	table.sortable th:not(.sorttable_sorted):not(.sorttable_sorted_reverse):not(.sorttable_nosort):after { 
	    content: " \25B4\25BE" 
	}
	table.sortable thead {
	    color:#222;
	    font-weight: bold;
	    cursor: pointer;
	}
	
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
	
	.container-center{
		max-width:1200px;
		margin:0 auto;
	}
	
	.page-header{
		width:100%;
		text-align:left;
	}
	
	.sortable-right{
		text-align:right;
	}
	
	.sortable-center{
		text-align:center;
	}
#Largebox {
  	width: 1200px;
	height: 250px;
	padding-top: 25px;
	padding-left: 10px;
	padding-right: 10px;
}
#box1 {
	width: 280px;
	height: 250px;
	border: 1px solid grey;
	float: left;
	padding: 10px 10px 10px 10px;
}
#box2 {
	width: 700px;
	height: 250px;
	border: 1px solid grey;
	float: left;
	margin-left:20px;
}
#box3 {
	width: 150px;
	height: 250px;
	float: right;
	
}
#box4 {
	width: 150px;
	height: 70px;
	border: 1px solid grey;
	float: left;
	margin-top: 20px;
	padding: 10px 10px 10px 10px;
}
#reportfrm {
	margin-left:20px;
}
    .colorbox {
	float: left;
	width: 100px;
	height: 50px;
	color: black;
    }
    .subject-info-box-1,
    .subject-info-box-2 {
        float: left;
        width: 250px;
        height: 400px;
	padding: 5px 5px 5px 5px;

        select {
            height: 250px;
            option {
                padding: 4px 10px 4px 10px;
            }
            option:hover {
                background: #EEEEEE;
            }
        }
    }
    .subject-info-arrows {
        float: left;
        width: 150px;
	padding: 10px 10px 10px 10px;
	margin-bottom: 5px;
        input {
            margin-bottom: 15px;
            margin-top: 15px;
            padding: 15px 15px 15px 15px;
        }
    }
#btnRight {
	margin-top:7px;
	margin-bottom:7px;
}

#btnAllRight {
	margin-bottom:7px;
}

</style>

<!-- datepicker for historical delivery board !-->
<script src="http://fl2.mdsinsight.com/js/moment.min.js" type="text/javascript"></script>
<link href="http://fl2.mdsinsight.com/js/bootstrap-datepicker/css/bootstrap-datepicker.css" rel="stylesheet" type="text/css">
<script src="http://fl2.mdsinsight.com/js/bootstrap-datepicker/js/bootstrap-datepicker.js" type="text/javascript"></script>
<!-- end datepicker for historical delivery board !-->

<script type="text/javascript">
	$(document).ready(function() {
	    $("#PleaseWaitPanel").hide();
	});
</script>


<h3 class="page-header"><i class="fa fa-globe"></i> Export Deliveries To Google Earth &nbsp;&nbsp;
	<a href="<%= BaseURL %>routing/reports/main.asp"><button type="button" class="btn btn-primary">Back To Routing Reports List</button></a></h3>

<div class="row">
	<div class="container-center">
		
<script>
   $(document).ready(function(){	
  
        $('#datepicker1').datepicker({
		autoclose: true,
		format: 'mm/dd/yyyy',
		useCurrent: false,
		minDate: '<%=startdate%>',
		startDate: '<%=startdate%>',
		endDate: '<%=enddate%>',
		defaultDate: '<%=enddate%>',
		maxDate: '<%=enddate%>',
		datesDisabled: [<%=disabledates%>]
        });
        
	$("#datepicker1").datepicker()
		.on('hide', function(e) {
	        // `e` here contains the extra attributes
	    	selectedDate = $("#datepicker1").find("input").val();
	        location.href = 'GoogleearthProExporta.asp?date=' + selectedDate;

	});

	$('#report_1').click(function (e) {
		location.href = 'GoogleearthProExporta.asp';
	});
        $('#btnRight').click(function (e) {
            var selectedOpts = $('#lstBox1 option:selected');
            if (selectedOpts.length == 0) {
                alert("Nothing to move.");
                e.preventDefault();
            }
            $('#lstBox2').append($(selectedOpts).clone());
            $(selectedOpts).remove();
	    movemarkers(selectedOpts, "map1");
            e.preventDefault();
        });
        $('#btnAllRight').click(function (e) {
            var selectedOpts = $('#lstBox1 option');
            if (selectedOpts.length == 0) {
                alert("Nothing to move.");
                e.preventDefault();
            }
            $('#lstBox2').append($(selectedOpts).clone());
            $(selectedOpts).remove();
	    movemarkers(selectedOpts, "map1");
            e.preventDefault();
        });
        $('#btnLeft').click(function (e) {
            var selectedOpts = $('#lstBox2 option:selected');
            if (selectedOpts.length == 0) {
                alert("Nothing to move.");
                e.preventDefault();
            }
            $('#lstBox1').append($(selectedOpts).clone());
            $(selectedOpts).remove();
	    movemarkers(selectedOpts, "map");
            e.preventDefault();
        });
        $('#btnAllLeft').click(function (e) {
            var selectedOpts = $('#lstBox2 option');
            if (selectedOpts.length == 0) {
                alert("Nothing to move.");
                e.preventDefault();
            }
            $('#lstBox1').append($(selectedOpts).clone());
            $(selectedOpts).remove();
	    movemarkers(selectedOpts, "map");
            e.preventDefault();
        });   
	$('#btnExport').click(function (e) {
		if ($("#radio_csv").is(":checked")) {
			var selectedOpts = $('#lstBox2 option');
		        if (selectedOpts.length == 0) {
                		alert("No User Selected");
		                e.preventDefault();
		        } else {
				csvtext = getcsvtext(selectedOpts);
				
				download(csvfilename, csvtext);
			}
		}
		if ($("#radio_kml").is(":checked")) {
			var selectedOpts = $('#lstBox2 option');
		        if (selectedOpts.length == 0) {
                		alert("No User Selected");
		                e.preventDefault();
		        }
			kmltext = getkmltext(selectedOpts);
			download(kmlfilename, kmltext);
		}

	});
      });     
//download(this['name'].value, this['text'].value)

function download(filename, text) {
  var element = document.createElement('a');
  element.setAttribute('href', 'data:text/plain;charset=utf-8,' + encodeURIComponent(text));
  element.setAttribute('download', filename);

  element.style.display = 'none';
  document.body.appendChild(element);

  element.click();

  document.body.removeChild(element);
}

  function getcsvtext(objs) {
	var h1 = "";
	var csvtext = "";
	for (h=0; h < objs.length; h++) {
		h1 += objs[h].value + "\r\n";
		for (i = 0; i < locations.length; i++) {
		   if (locations[i][0] == objs[h].value) {
			for(j = 0; j < locations[i][3].length; j++) {
			    if (typeof(locations[i][3][j][16]) != "undefined") {
				csvtext += "\"" + locations[i][3][j][4] + "\",\"" + locations[i][3][j][3] + "\",\"" + locations[i][0] + "\",\"" + locations[i][1] + "\"";
				for(k = 5; k < locations[i][3][j].length - 1; k++) {
					if (typeof(locations[i][3][j][k]) != "undefined") {
						csvtext += ",\"" + locations[i][3][j][k] + "\"";
					}
				}
				csvtext += "\r\n";
			    }
			}
		   }
		}
	}
	return csvtext;
  }
  function getkmltext(objs) {
	var h1 = "";
	var kmltext = "";
	kmltext += "<?xml version=\"1.0\" encoding=\"UTF-8\"\?>\r\n";
	kmltext += "<kml xmlns=\"http://www.opengis.net/kml/2.2\" \r\nxmlns:gx=\"http://www.google.com/kml/ext/2.2\" \r\n";
	kmltext += "xmlns:kml=\"http://www.opengis.net/kml/2.2\" \r\nxmlns:atom=\"http://www.w3.org/2005/Atom\">\r\n";
	kmltext += "<Document>\r\n";
	kmltext += "\t<name>" + kmlfilename + "</name>\r\n";
	var style1 = "";
	var placemark = "";
	for (h=0; h < objs.length; h++) {
		h1 += objs[h].value + "\r\n";
		for (i = 0; i < locations.length; i++) {
		   if (locations[i][0] == objs[h].value) {
			st = "\t<Style id=\"sh_wht-blank" + locations[i][0] + "\">\r\n";
			st1 = "\t<Style id=\"sn_wht-blank" + locations[i][0] + "\">\r\n";
			st += "\t\t<IconStyle>\r\n";
			st1 += "\t\t<IconStyle>\r\n";
			s = locations[i][2].substr(1);
			s = s.slice(4) + s.slice(2,4) + s.slice(0,2);
			st += "\t\t\t<color>ff"+ s +"</color>\r\n";
			st1 += "\t\t\t<color>ff"+ s +"</color>\r\n";
			st += "\t\t\t<scale>1.3</scale>\r\n";
			st1 += "\t\t\t<scale>1.1</scale>\r\n";
			st += "\t\t\t<Icon>\r\n";
			st1 += "\t\t\t<Icon>\r\n";
			st += "\t\t\t\t<href>http://maps.google.com/mapfiles/kml/paddle/wht-blank.png</href>\r\n";
			st1 += "\t\t\t\t<href>http://maps.google.com/mapfiles/kml/paddle/wht-blank.png</href>\r\n";
			st += "\t\t\t</Icon>\r\n";
			st1 += "\t\t\t</Icon>\r\n";
			st += "\t\t\t<hotSpot x=\"32\" y=\"1\" xunits=\"pixels\" yunits=\"pixels\"/>\r\n";
			st1 += "\t\t\t<hotSpot x=\"32\" y=\"1\" xunits=\"pixels\" yunits=\"pixels\"/>\r\n";
			st += "\t\t</IconStyle>\r\n";
			st1 += "\t\t</IconStyle>\r\n";
			st += "\t\t<LabelStyle>\r\n";
			st1 += "\t\t<LabelStyle>\r\n";
			st += "\t\t\t<color>ff" + s + "</color>\r\n";
			st1 += "\t\t\t<color>ff" + s + "</color>\r\n";
			st += "\t\t</LabelStyle>\r\n";
			st1 += "\t\t</LabelStyle>\r\n";
			st += "\t\t<BalloonStyle>\r\n";
			st1 += "\t\t<BalloonStyle>\r\n";
			st += "\t\t</BalloonStyle>\r\n";
			st1 += "\t\t</BalloonStyle>\r\n";
			st += "\t\t<ListStyle>\r\n";
			st1 += "\t\t<ListStyle>\r\n";
			st += "\t\t\t<bgColor>ff" + s + "</bgColor>\r\n";
			st1 += "\t\t\t<bgColor>ff" + s + "</bgColor>\r\n";
			st += "\t\t\t<ItemIcon>\r\n";
			st1 += "\t\t\t<ItemIcon>\r\n";
			st += "\t\t\t\t<href>http://maps.google.com/mapfiles/kml/paddle/wht-blank-lv.png</href>\r\n";
			st1 += "\t\t\t\t<href>http://maps.google.com/mapfiles/kml/paddle/wht-blank-lv.png</href>\r\n";
			st += "\t\t\t</ItemIcon>\r\n";
			st1 += "\t\t\t</ItemIcon>\r\n";
			st += "\t\t</ListStyle>\r\n";
			st1 += "\t\t</ListStyle>\r\n";
			st += "\t</Style>\r\n";
			st1 += "\t</Style>\r\n";
			stfinal = st+st1;
			stfinal += "\t<StyleMap id=\"msn_wht-blank" + locations[i][0] + "\">\r\n";
			stfinal += "\t\t<Pair>\r\n";
			stfinal += "\t\t\t<key>normal</key>\r\n";
			stfinal += "\t\t\t<styleUrl>#sn_wht-blank" + locations[i][0] + "</styleUrl>\r\n";
			stfinal += "\t\t</Pair>\r\n";
			stfinal += "\t\t<Pair>\r\n";
			stfinal += "\t\t\t<key>highlight</key>\r\n";
			stfinal += "\t\t\t<styleUrl>#sh_wht-blank" + locations[i][0] + "</styleUrl>\r\n";
			stfinal += "\t\t</Pair>\r\n";
			stfinal += "\t</StyleMap>\r\n";
			style1 += stfinal;

			for(j = 0; j < locations[i][3].length; j++) {
			    if (typeof(locations[i][3][j][16]) != "undefined") {
				//kmltext += "\"" + locations[i][3][j][4] + "\",\"" + locations[i][3][j][3] + "\",\"" + locations[i][0] + "\",\"" + locations[i][1] + "\"";
				placemark += "\t\t<Placemark>\r\n";
				placemark += "\t\t\t<name>" + locations[i][3][j][5].replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&apos;') + "</name>\r\n";
				var desc = locations[i][3][j][5] + "\r\n" + locations[i][3][j][6] + "\r\n" + locations[i][3][j][7] + "\r\n" + locations[i][3][j][8];
				desc = desc.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&apos;');
				placemark += "\t\t\t<description>" + desc + "</description>\r\n";
				placemark += "\t\t\t<LookAt>\r\n";
				placemark += "\t\t\t\t<longitude>" + locations[i][3][j][3] + "</longitude>\r\n";
				placemark += "\t\t\t\t<latitude>" + locations[i][3][j][4] + "</latitude>\r\n";
				placemark += "\t\t\t\t<altitude>0</altitude>\r\n";
				placemark += "\t\t\t\t<heading>0</heading>\r\n";
				placemark += "\t\t\t\t<tilt>0</tilt>\r\n";
				placemark += "\t\t\t\t<range>999.9999990774612</range>\r\n";
				placemark += "\t\t\t\t<gx:altitudeMode>relativeToSeaFloor</gx:altitudeMode>\r\n";
				placemark += "\t\t\t</LookAt>\r\n";
				placemark += "\t\t\t<styleUrl>#msn_wht-blank" + locations[i][0] + "</styleUrl>\r\n";
				placemark += "\t\t\t<Point>\r\n";
				placemark += "\t\t\t\t<gx:drawOrder>1</gx:drawOrder>\r\n";
				placemark += "\t\t\t\t<coordinates>" + locations[i][3][j][3] + "," + locations[i][3][j][4] + ",0</coordinates>\r\n";
				placemark += "\t\t\t</Point>\r\n";
				placemark += "\t\t</Placemark>\r\n";

				//for(k = 5; k < locations[i][3][j].length - 1; k++) {
				//	if (typeof(locations[i][3][j][k]) != "undefined") {
				//		//kmltext += ",\"" + locations[i][3][j][k] + "\"";
				//	}
				//}
			    }
			}
		   }
		}
	}
	kmltext += style1
	kmltext += "\t<Folder>\r\n";
	kmltext += "\t\t<name>" + curdate + "</name>\r\n";
	kmltext += "\t\t<open>1</open>\r\n";
	kmltext += "\t\t<Style>\r\n";
	kmltext += "\t\t\t<ListStyle>\r\n";
	kmltext += "\t\t\t\t<listItemType>check</listItemType>\r\n";
	kmltext += "\t\t\t\t<ItemIcon>\r\n";
	kmltext += "\t\t\t\t\t<state>open</state>\r\n";
	kmltext += "\t\t\t\t\t<href>:/mysavedplaces_open.png</href>\r\n";
	kmltext += "\t\t\t\t</ItemIcon>\r\n";
	kmltext += "\t\t\t\t<ItemIcon>\r\n";
	kmltext += "\t\t\t\t\t<state>closed</state>\r\n";
	kmltext += "\t\t\t\t\t<href>:/mysavedplaces_closed.png</href>\r\n";
	kmltext += "\t\t\t\t</ItemIcon>\r\n";
	kmltext += "\t\t\t\t<bgColor>00ffffff</bgColor>\r\n";
	kmltext += "\t\t\t\t<maxSnippetLines>2</maxSnippetLines>\r\n";
	kmltext += "\t\t\t</ListStyle>\r\n";
	kmltext += "\t\t</Style>\r\n";
	kmltext += placemark;
	kmltext += "\t</Folder>\r\n";
	kmltext += "</Document>\r\n"
	kmltext += "</kml>\r\n";

	return kmltext;
  }
</script>

<div id="Largebox">
<div id="box1">
<p> <form action="" id="reportfrm">
  <input type="radio" name="report" id="report_1" value="export_deliveries" <%=checked1%>> Export today's deliveries<br><br>
  <input type="radio" name="report" id="report_2" value="date_export" <%=checked2%>> Select a date to export<BR>&nbsp;<BR>
 		<!-- Bootstrap datepicker for filtering leads by date -->
        <div class="form-group">
            <div class="input-group date datepicker" id="datepicker1">
                <input type="text" class="form-control" name="txtDriverDate" id="txtDriverDate" value="<% if DateToRetrieve = "" then response.write enddate else response.write DateToRetrieve end if %>">
                <span class="input-group-addon">
                    <span class="glyphicon glyphicon-calendar"></span>
                </span>
            </div>
      </div>
    <!-- eof datepicker !-->
<BR>
 
</p>
</div>
<div id="box2">
    <div class="subject-info-box-1">
      <select multiple="multiple" id='lstBox1' size="12" class="form-control">
<%
If DateToRetrieve = "" Then
	SQL = "SELECT DISTINCT tblUsers.userDisplayName, tblUsers.userTruckNumber FROM RT_DeliveryBoard INNER JOIN tblUsers ON tblUsers.userTruckNumber = RT_DeliveryBoard.TruckNumber ORDER BY tblUsers.userDisplayName"
ELSE
	SQL = "SELECT DISTINCT tblUsers.userDisplayName, tblUsers.userTruckNumber FROM RT_DeliveryBoardHistory INNER JOIN tblUsers ON tblUsers.userTruckNumber = RT_DeliveryBoardHistory.TruckNumber WHERE (RT_DeliveryBoardHistory.DeliveryDate='" & DateToRetrieve & "') ORDER BY tblUsers.userDisplayName"
End If
	Set rs = conn.Execute(SQL)
	i = 0
	randomize()
	jsarray = "  <script type=""text/javascript""> " & vbcrlf & "    var locations = [" & vbcrlf
	csvtext = ""
	kmltext = ""
	do while not rs.Eof
		'response.write i & vbcrlf
		c1 = optioncolors(i)
		if i = 0 then

		else
			jsarray = jsarray & ","
		end if
		jsarray = jsarray & vbcrlf & "[""" & rs("userTruckNumber") & """,""" & rs("userDisplayName") & """,""" & c1 & """,[" & vbcrlf
		Response.Write vbcrlf & "<option value=""" &  rs("userTruckNumber") & """ style=""background-color: " & c1 & ";color: black;"">" &  rs("userDisplayName") & " (" & rs("userTruckNumber") & ")</option>" & vbcrlf
		If DateToRetrieve = "" Then
			SQL1 = "SELECT RT_DeliveryBoard.CustNum,RT_DeliveryBoard.TruckNumber,RT_DeliveryBoard.SequenceNumber,AR_Customer.Longitude,AR_Customer.Latitude,AR_Customer.Name,AR_Customer.Addr1,AR_Customer.Addr2,AR_Customer.CityStateZip,AR_Customer.Phone,AR_Customer.Contact,AR_Customer.ContactFirstName,AR_Customer.ContactLastName,AR_Customer.City,AR_Customer.State,AR_Customer.Zip,RT_DeliveryBoard.DeliveryDate FROM RT_DeliveryBoard LEFT JOIN AR_Customer ON RT_DeliveryBoard.CustNum = AR_Customer.CustNum WHERE RT_DeliveryBoard.TruckNumber='" & rs("userTruckNumber") & "'"
		ELSE
			SQL1 = "SELECT RT_DeliveryBoardHistory.CustNum,RT_DeliveryBoardHistory.TruckNumber,RT_DeliveryBoardHistory.SequenceNumber,AR_Customer.Longitude,AR_Customer.Latitude,AR_Customer.Name,AR_Customer.Addr1,AR_Customer.Addr2,AR_Customer.CityStateZip,AR_Customer.Phone,AR_Customer.Contact,AR_Customer.ContactFirstName,AR_Customer.ContactLastName,AR_Customer.City,AR_Customer.State,AR_Customer.Zip,RT_DeliveryBoardHistory.DeliveryDate FROM RT_DeliveryBoardHistory LEFT JOIN AR_Customer ON RT_DeliveryBoardHistory.CustNum = AR_Customer.CustNum WHERE RT_DeliveryBoardHistory.TruckNumber='" & rs("userTruckNumber") & "' and (RT_DeliveryBoardHistory.DeliveryDate='" & DateToRetrieve & "')" 
		END IF
		'err.clear
		'response.write SQL1 & "<BR>"  & vbcrlf
		Set rs1 = conn.Execute(SQL1)
		'response.write err.number & "--" & err.description & "<BR>"  & vbcrlf
		'response.write rs1.bof & "<BR>" & vbcrlf
		k = 0
		do while not rs1.Eof
			k = k + 1
			'response.write "\nk: " & k & " " & rs1("3") & "<BR>" & vbcrlf
			'response.write "CustNum:" & rs1("CustNum") & " ,TruckNumber:" & rs1("TruckNumber") & " ,SequenceNumber:" & rs1("SequenceNumber") & " ,Longitude:" & rs1("Longitude") & " ,Latitude:" & rs1("Latitude") & " ,Name:" & rs1("Name") & " ,Addr1:" & rs1("Addr1") & " ,Addr2:" & rs1("Addr2") & " ,CityStateZip:" & rs1("CityStateZip") & " ,Phone:" & rs1("Phone") & " ,Contact:" & rs1("Contact") & " ,ContactFirstName:" & rs1("ContactFirstName") & " ,ContactLastName:" & rs1("ContactLastName") & " ,City:" & rs1("City") & " ,State:" & rs1("State") & " ,Zip:" & rs1("Zip") & vbcrlf
			csvtext1 = "\""" & rs("userTruckNumber") & "\"",\""" & rs("userDisplayName") & "\"",\""" & rs1("Latitude") & "\"",\""" & rs1("Longitude") & "\"",\""" & rs1("Name") & "\"",\""" & rs1("Addr1") & "\"",\""" & rs1("Addr2") & "\"",\""" & rs1("City") & "\"",\""" & rs1("State") & "\"",\""" & rs1("Zip") & "\"",\""" & rs1("CityStateZip") & "\"",\""" & rs1("Phone") & "\"",\""" & rs1("Contact") & "\"",\""" & rs1("ContactFirstName") & "\"",\""" & rs1("ContactLastName") & "\""\r\n"
			if (not isnull(rs1("Addr1"))) OR (not isnull(rs1("Addr2"))) then
				if isnull(rs1("Longitude")) OR isnull(rs1("Latitude")) then
					if (instr(1, rs1("Addr1"), "ATTN") > 0) then
						address = rs1("Addr2") & " " & rs1("CityStateZip")
					else 
						address = rs1("Addr1") & " " & rs1("Addr2") & " " & rs1("CityStateZip")
					end if
					coords = GetXML(address)
					'Do we have a valid array?
					If IsArray(coords) Then
					  'Response.Write "The geo-coded coordinates are: " & Join(coords, ",") & vbcrlf
					   SQL = "UPDATE AR_Customer set Latitude='" & coords(0) & "',Longitude='" & coords(1) & "' WHERE CustNum='" & rs1("CustNum") & "'"
					   'response.write SQL & vbcrlf
					   set rs2 = conn.Execute(SQL)
					Else
					  'No coordinates were returned
					  'Response.Write "The address could not be geocoded."  & vbcrlf
					End If
				end if
			end if
			if (k = 1) then

			else 
				jsarray = jsarray & "," & vbcrlf
			end if
			jsarray = jsarray & "[""" & rs1("CustNum") & """,""" & rs1("TruckNumber") & """,""" & rs1("SequenceNumber") & """,""" & rs1("Longitude") & """,""" & rs1("Latitude") & """,""" & rs1("Name") & """,""" & rs1("Addr1") & """,""" & rs1("Addr2") & """,""" & rs1("CityStateZip") & """,""" & rs1("Phone") & """,""" & rs1("Contact") & """,""" & rs1("ContactFirstName") & """,""" & rs1("ContactLastName") & """,""" & rs1("City") & """,""" & rs1("State") & """,""" & rs1("Zip") & """]"
			'     For l = 0 To rs1.Fields.Count - 1
			'         Response.Write l & " -- " & rs1(l) & "<BR>" & vbcrlf
			'     Next
			curtdate = Right("0" & month(rs1("DeliveryDate")), 2) & "-" & Right("0" & Day(rs1("DeliveryDate")), 2)  & "-" & Year(rs1("DeliveryDate"))
			rs1.movenext
			if (k > 20) then
				exit do
			end if
		loop
		jsarray = jsarray & "]]"	& vbcrlf
		i = i + 1

		rs.movenext
	loop
	jsarray = jsarray & "]</script>"
%>
      </select>
    </div>
    <div class="subject-info-arrows text-center">
      <input type="button" id="btnRight" value="Add Route->" class="btn btn-success" /><br />
      <input type="button" id="btnLeft" value="<-Remove Route" class="btn btn-danger" /><br />
	<br />
<br />&nbsp; 

      <input type="button" id="btnAllRight" value="Add All->" class="btn btn-success" /><br />
      <input type="button" id="btnAllLeft" value="<-Remove All" class="btn btn-danger" />
    </div>
    <div class="subject-info-box-2">
      <select multiple="multiple" id='lstBox2' size="12" class="form-control">

      </select>
    </div>
    <div class="clearfix"></div>
</div>
<div id="box3">
<h4>Export File Type</h4>
<p> 
<div id="box4">

  <input type="radio" name="createFile" id="radio_kml" value="kml" checked> create.kml file<br>
  <input type="radio" name="createFile" id="radio_csv" value="csv"> create.csv file<br>  
</div>
<br />&nbsp; 
<br />&nbsp; 
  <input type="button" id="btnCancel" value="Cancel" class="btn btn-default" />
  <input type="button" id="btnExport" value="Export" class="btn btn-primary" />
</form> 
</p>
</div>
</div>
<div style="width: 1200;">
<div id="map" style="width: 48%; height: 500px;float: left;"></div><div style="width: 1%; height: 500px;float:left"></div><div id="map1" style="width: 48%; height: 500px; float: right;"></div>
  <script src="http://maps.google.com/maps/api/js?sensor=false&key=AIzaSyBoiohQSvfpeqB_5nlzHxXTuw6fBvBJTaw" 
          type="text/javascript"></script>
<% response.write jsarray %>
<script>
    var csvtext = "<% response.write csvtext %>";
    var csvfilename = "<% response.write curtdate %>.csv";
    //download(csvfilename, csvtext);
    var kmltext = "<% response.write kmltext %>";
    var kmlfilename = "<% response.write curtdate %>.kml";
    var curdate = "<% response.write curtdate %>";

    var bound = new google.maps.LatLngBounds();

    var map = new google.maps.Map(document.getElementById('map'), {
      zoom: 9,
      center: new google.maps.LatLng(locations[0][3][0][4], locations[0][3][0][3]),
      mapTypeId: google.maps.MapTypeId.ROADMAP
    });

    var map1 = new google.maps.Map(document.getElementById('map1'), {
      zoom: 9,
      center: new google.maps.LatLng(locations[0][3][0][4], locations[0][3][0][3]),
      mapTypeId: google.maps.MapTypeId.ROADMAP
    });

    var infowindow = new google.maps.InfoWindow();
    var marker, i, j;

    for (i = 0; i < locations.length; i++) {
	for(j = 0; j < locations[i][3].length; j++) {
	   if (typeof(locations[i][3][j]) != "undefined") {
	     if (locations[i][3][j][4] == "" || locations[i][3][j][3] == "") {

	     } else {
		bound.extend( new google.maps.LatLng(locations[i][3][j][4], locations[i][3][j][3]) );
      		marker = new google.maps.Marker({
        	position: new google.maps.LatLng(locations[i][3][j][4], locations[i][3][j][3]),
	        draggable: true,
	        map: map,
        	labelContent: locations[i][3][j][5],
	        labelAnchor: new google.maps.Point(22,0),
        	labelClass: "labels",
	        labelStyle: {opacity: 0.85},
		title : locations[i][3][j][5],
		zIndex: locations[i][3][j][2],
		icon: 'http://chart.apis.google.com/chart?chst=d_map_pin_letter&chld='+locations[i][3][j][2]+'|'+locations[i][2].substr(1)+'|000000'  
	      });
		locations[i][3][j][16] = marker;
	      google.maps.event.addListener(marker, 'click', (function(marker, i, j) {
        	return function() {
	          infowindow.setContent(locations[i][0] + " " + locations[i][1] + "<BR>" + locations[i][3][j][5] + "<BR>" + locations[i][3][j][6] + "<BR>" + locations[i][3][j][7] + "<BR>" + locations[i][3][j][8] + "<BR>");
        	  infowindow.open(marker.getMap, marker);
	        }
	      })(marker, i, j));
	     }
	   }
	}
    }
	mapcenter = bound.getCenter();
	map.setCenter(new window.google.maps.LatLng(mapcenter.lat(), mapcenter.lng()));
	map1.setCenter(new window.google.maps.LatLng(mapcenter.lat(), mapcenter.lng()));

  function movemarkers(objs, movetomap) {
	var h1 = "";
	for (h=0; h < objs.length; h++) {
		h1 += objs[h].value + "\r\n";
		for (i = 0; i < locations.length; i++) {
		   if (locations[i][0] == objs[h].value) {	
			for(j = 0; j < locations[i][3].length; j++) {
			    if (typeof(locations[i][3][j][16]) != "undefined") {
				if (movetomap == "map") {
					locations[i][3][j][16].setMap(map);
				} else {
					locations[i][3][j][16].setMap(map1);
				}
			    }
			}
		   }
		}
	}
  }
  </script>
</div>
<div class="clearfix"></div>
<div style="height: 20px;float:left"></div>
<div class="clearfix"></div>
<div style="width: 1200;">
<div class="colorbox" style="background-color: #FFFFFF;">#FFFFFF</div>
<div class="colorbox" style="background-color: #87D8C7;">#87D8C7</div>
<div class="colorbox" style="background-color: #D0C76F;">#D0C76F</div>
<div class="colorbox" style="background-color: #D9BCB5;">#D9BCB5</div>
<div class="colorbox" style="background-color: #E24932;">#E24932</div>
<div class="colorbox" style="background-color: #D13EA2;">#D13EA2</div>
<div class="clearfix"></div>
<div class="colorbox" style="background-color: #DD754C;">#DD754C</div>
<div class="colorbox" style="background-color: #3FE337;">#3FE337</div>
<div class="colorbox" style="background-color: #8F6FDA;">#8F6FDA</div>
<div class="colorbox" style="background-color: #BD4C73;">#BD4C73</div>
<div class="colorbox" style="background-color: #509FD3;">#509FD3</div>
<div class="colorbox" style="background-color: #796230;">#796230</div>
<div class="clearfix"></div>
<div class="colorbox" style="background-color: #FFFF00;">#FFFF00</div>
<div class="colorbox" style="background-color: #FF00FF;">#FF00FF</div>
<div class="colorbox" style="background-color: #00ffff;">#00ffff</div>
<div class="colorbox" style="background-color: #66819D;">#66819D</div>
<div class="colorbox" style="background-color: #2C6FD3;">#2C6FD3</div>
<div class="colorbox" style="background-color: #ff0000;">#ff0000</div>
<div class="clearfix"></div>
<div class="colorbox" style="background-color: #00FF00;">#00FF00</div>
<div class="colorbox" style="background-color: #4444FF;">#4444FF</div>
<div class="colorbox" style="background-color: #495CB9;">#495CB9</div>
<div class="colorbox" style="background-color: #B47C34;">#B47C34</div>
<div class="colorbox" style="background-color: #36AF9C;">#36AF9C</div>
<div class="colorbox" style="background-color: #E17A75;">#E17A75</div>
<div class="clearfix"></div>
<div class="colorbox" style="background-color: #AFC5A2;">#AFC5A2</div>
<div class="colorbox" style="background-color: #726E69;">#726E69</div>
<div class="colorbox" style="background-color: #8B4BDF;">#8B4BDF</div>
<div class="colorbox" style="background-color: #4EA16B;">#4EA16B</div>
<div class="colorbox" style="background-color: #ffb6c1;">#ffb6c1</div>
<div class="colorbox" style="background-color: #FFA500;">#FFA500</div>
<div class="clearfix"></div>
</div>

<%
Function GetXML(addr)
  'on error goto 0
  Dim objXMLDoc, url, docXML, lat, lng, mapref

  'URL for Google Maps API - Doesn't need to stay here could be stored in a 
  'config include file or passed in as a function parameter.
  url = "https://maps.googleapis.com/maps/api/geocode/xml?address={addr}&sensor=false&key=AIzaSyDfgyquAQg0QzNnaxMTUdxH5CJNuoSxcwY"
  'Inject address into the URL
  url = Replace(url, "{addr}", Server.URLEncode(addr))

  Set objXMLDoc = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
  objXMLDoc.setTimeouts 30000, 30000, 30000, 30000
  objXMLDoc.Open "GET", url, False
  objXMLDoc.send()

  'response.write objXMLDoc.status & vbcrlf 
  'response.write objXMLDoc.responseText & vbcrlf
  'response.wrte objXMLDoc.responseXML & vbcrlf 
  If objXMLDoc.status = 200 Then
    'Check the response for a valid status
     okpos = instr(1, objXMLDoc.responseText, "<status>OK</status>")
    if (okpos > 0) then
    'If UCase(docXML.documentElement.selectSingleNode("/GeocodeResponse/status").Text) = "OK" Then
	locationpos =  instr(1, objXMLDoc.responseText, "<location>")
	locationendpos = instr(1, objXMLDoc.responseText, "</location>")
	'response.write "locationpos: " & locationpos & " , locationendpos: " & locationendpos & vbcrlf
	lat = mid(objXMLDoc.responseText, instr(locationpos, objXMLDoc.responseText, "<lat>") + 5, instr(locationpos, objXMLDoc.responseText, "</lat>") - (instr(locationpos, objXMLDoc.responseText, "<lat>") + 5))
	lng = mid(objXMLDoc.responseText, instr(locationpos, objXMLDoc.responseText, "<lng>") + 5, instr(locationpos, objXMLDoc.responseText, "</lng>") - (instr(locationpos, objXMLDoc.responseText, "<lng>") + 5))
	'response.write "locationpos: " & locationpos & " , locationendpos: " & locationendpos & ", lat:" & lat & ", lng:" & lng & vbcrlf
      'lat = docXML.documentElement.selectSingleNode("/GeocodeResponse/result/geometry/location/lat").Text
      'lng = docXML.documentElement.selectSingleNode("/GeocodeResponse/result/geometry/location/lng").Text
      'Create array containing lat and long
      mapref = Array(lat, lng)
    Else
      mapref = Empty
    End If
  Else
    mapref = Empty
  End If

  'Return array
  GetXML = mapref
  'on error resume next
End Function

%>
		
