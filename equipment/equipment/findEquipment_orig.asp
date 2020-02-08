<!--#include file="../../inc/header.asp"-->
<!--#include file="../../inc/InsightFuncs_Equipment.asp"-->
<%
Server.ScriptTimeout = 900000 'Default value

ClientKeyForFileNames = MUV_READ("ClientKeyForFileNames")

EquipIDForDetail = Request.QueryString("CID")

If EquipIDForDetail = "" Then 
	EquipIDForDetail = Request.Form("txtEquipIDToPass")
End If

%>
<!---------------------------------------------------------------------------------------------------------->
<!----------THIS IS A CUSTOM STYLESHEET ADDED FOR THE AUTOCOMPLETE SEARCH FOR CATEGORY ANALYSIS ONLY-------->
<!-----------IT OVERRIDES THE STYLES THAT ARE STILL LOADED IN HEADER.ASP------------------------------------>

<!---------------------------------------------------------------------------------------------------------->
<link rel="stylesheet" href="<%= BaseURL %>js/easyautocomplete/EasyAutocomplete-1.4.0/easy-autocomplete-cat-analysis.css"> 

<!---------------------------------------------------------------------------------------------------------->
<!---------------------------------------------------------------------------------------------------------->
<!--
	js/easyautocomplete/EasyAutocomplete-1.4.0/easy-autocomplete.themes.css ALSO CONTAINS A CUSTOM STYLE
	SET CALLED "easy-autocomplete.eac-cat-analysis"
-->
<!---------------------------------------------------------------------------------------------------------->
<!---------------------------------------------------------------------------------------------------------->

<script type="text/javascript">


$(document).ready(function(){	

	$("#PleaseWaitPanel").hide();
	
	var randomNumberBetween0and100 = Math.floor(Math.random() * 100);
	
	var autocompleteJSONFileURLAccount = "../../clientfiles/<%= ClientKeyForFileNames %>/autocomplete/equipment_list_<%= ClientKeyForFileNames %>.json?v=" + randomNumberBetween0and100;

	var optionsEquipment = {
	  url: autocompleteJSONFileURLAccount,
	  placeholder: "Search for equipment by serial number or asset tag",
	  getValue: "name",
	  list: {	
        onChooseEvent: function() {
        
            var EquipIntRecID = $("#txtEquipID").getSelectedItemData().code;
            $("#txtEquipIDToPass").val(EquipIntRecID);
            window.location.href = "editEquipment.asp?i=" + EquipIntRecID;
            
    	},		  
	    match: {
	      enabled: true
		},
		maxNumberOfElements: 30		
	  },
	  theme: "cat-analysis"
	};
	
	$("#txtEquipID").easyAutocomplete(optionsEquipment);
	
  
});



</script>

<%

		'*********************************************************
		' Begin Auto Complete Equipment List
		'*********************************************************

		'SQL = "SELECT InternalRecordIdentifier, ModelIntRecID, SerialNumber, AssetTag1 FROM EQ_Equipment WHERE ModelIntRecID <> '' ORDER BY SerialNumber"
		'Set cnnAutoComplete = Server.CreateObject("ADODB.Connection")
		'cnnAutoComplete.open (Session("ClientCnnString"))
		'Set rsAutoComplete = Server.CreateObject("ADODB.Recordset")
		'rsAutoComplete.CursorLocation = 3 
		'Set rsAutoComplete = cnnAutoComplete.Execute(SQL)
		
		'If not rsAutoComplete.EOF Then
		'strAuto = "["
		'Do While Not rsAutoComplete.EOF
		   ' strAuto = strAuto & "{""name"":""" & GetModelNameByIntRecID(rsAutoComplete("ModelIntRecID")) & " --- " & rsAutoComplete("SerialNumber") & " --- " & rsAutoComplete("AssetTag1") & """, ""code"":""" & rsAutoComplete("InternalRecordIdentifier") & """},"
		    'rsAutoComplete.MoveNext
		'Loop
		'End If
		
		'If right(strAuto,1)= "," Then strAuto = left(strAuto,len(strAuto)-1) 
		
		'strAuto = trim(strAuto) & "]"
		
		'set fs=Server.CreateObject("Scripting.FileSystemObject")
		'set fs2=Server.CreateObject("Scripting.FileSystemObject")
		
		'set tfile=fs.CreateTextFile(Server.MapPath("..\..\..\") & "\clientfiles\"  & ClientKeyForFileNames & "\autocomplete\equipment_list_" & ClientKeyForFileNames & ".json")
		'tfile.WriteLine(strAuto)
		'tfile.close
		'set tfile=nothing
		'set fs=nothing
		
		'Set rsAutoComplete = Nothing
		'cnnAutoComplete.Close
		'Set AutoComplete = nothing

		
		'*********************************************************
		' END Auto Complete Equipment List
		'*********************************************************


%>

 
<style type="text/css">
 	.email-table{
		width:46%;
	}
	
	table.sortable th:not(.sorttable_sorted):not(.sorttable_sorted_reverse):not(.sorttable_nosort):after { 
    	content: " \25B4\25BE" 
	}
	
	.nav-tabs>li>a{
		background: #f5f5f5;
		border: 1px solid #ccc;
		color: #000;
	}
	
	.nav-tabs>li>a:hover{
		border: 1px solid #ccc;
	}
	
	.nav-tabs>li.active>a, .nav-tabs>li.active>a:focus, .nav-tabs>li.active>a:hover{
		color: #000;
		border: 1px solid #ccc;
	}
	
	.container{
		max-width:1200px;
		margin:0 auto;
	}

	.narrow-results{
		margin:0px 0px 20px 0px;
	}
	
	#filter{
		width:40%;
	}
	
	.modal-link{
		cursor:pointer;
	}
	
	.modal-content{
		max-height:360px;
		overflow-y:auto;
	}

	 .modal-content .row{
		 padding-bottom:20px;
	 }
	
	 .modal-content p{
		 margin-bottom:20px;
		 white-space:normal;
	 }
</style>

<!--- eof on/off scripts !-->

<h1 class="page-header">Find / Edit <%= GetTerm("Equipment") %></h1>

	<div class="row">
	 	<div class="col-lg-12">
		 	<p><a href="addEquipment.asp"><button type="button" class="btn btn-success">Add New Piece of Equipment</button></a></p>
	 	</div>
	</div>
	
	<br>	


	<div class="row">
		<div class="col-lg-3">
    		<!-- select equipment record !-->
				<input id="txtEquipID" name="txtEquipID">
				<input type="hidden" id="txtEquipIDToPass" name="txtEquipIDToPass" value="<%= EquipIDForDetail %>" >
				<i id="searchIcon" class="fa fa-search fa-2x"></i>
			<!-- eof select equipment record !-->
		</div>
	</div>
	<!-- eof row !-->    


								

<!--#include file="../../inc/footer-main.asp"-->