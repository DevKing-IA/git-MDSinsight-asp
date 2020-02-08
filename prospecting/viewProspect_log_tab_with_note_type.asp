<%'******************
' **** Notes Tab ****
'********************
%>
<%
SQLLogNoteTypes = "SELECT *, InternalRecordIdentifier as id FROM PR_NoteTypes ORDER BY NoteType"
Set cnnLogNoteTypes = Server.CreateObject("ADODB.Connection")
cnnLogNoteTypes.open (Session("ClientCnnString"))
Set rsLogNoteTypes = Server.CreateObject("ADODB.Recordset")
rsLogNoteTypes.CursorLocation = 3 
Set rsLogNoteTypes = cnnLogNoteTypes.Execute(SQLLogNoteTypes)

LogNoteTypes = ("[")
If not rsLogNoteTypes.EOF Then
	sep = ""
	Do While Not rsLogNoteTypes.EOF
			LogNoteTypes = LogNoteTypes & (sep)
			sep = ","
			LogNoteTypes = LogNoteTypes & ("{")
			LogNoteTypes = LogNoteTypes & ("""id"":""" & Replace(rsLogNoteTypes("id"), """", "\""") & """")
			LogNoteTypes = LogNoteTypes & (",""title"":""" & Replace(rsLogNoteTypes("NoteType"), """", "\""") & """")
			LogNoteTypes = LogNoteTypes & ("}")
		rsLogNoteTypes.MoveNext						
	Loop
End If
LogNoteTypes = LogNoteTypes & ("]")
Set rsLogNoteTypes = Nothing
cnnLogNoteTypes.Close
Set cnnLogNoteTypes = Nothing


%> 
<input type="hidden" name="txtProspectID" id="txtProspectID" value="<%= InternalRecordIdentifier %>">				
<script language="javascript">

	$(document).ready(function() {
		
		//we need to reload the page when the full email modal is closed because the hash tag
		//it leaves in the URL makes the page throw jQuery errors
	
	    $("[id^='myEmailModal']").on('hidden.bs.modal', function () {
	    	$(this).removeData('bs.modal');
	    	prospectID = $("#txtProspectID").val();
	        window.location.href = "viewProspectDetail.asp?id=" + prospectID;
	    });
	});
		
	
</script>
							
<div role="tabpanel" class="tab-pane fade in active" id="log">

	<p> <button type="button" class="btn btn-success" onclick="ajaxRowNewLogNotes();">New Log Note Entry</button> </p>

		<div class="input-group narrow-results"> <span class="input-group-addon">Narrow Results</span>
		    <input id="filter-notes" type="text" class="form-control filter-search-width" placeholder="Type here...">
		</div>
	  
		<div class="table-responsive">
            <table    class="table table-striped table-condensed table-hover table-bordered sortable standard-font">
              <thead  >
                <tr>
                  <th width="8%">Log Type</th>
                  <th width="5%">Date</th>
				  <th width="5%">Time</th>
				  <th width="8%">Entered By</th>
				  <th style="width: 180px;">Note Type</th>
				  <th>Notes</th>
  				  <th width="3%">Sticky</th>
                  <th class="sorttable_nosort text-center" style="width: 80px;">Actions</th>
                </tr>
              </thead>

			<tbody id="ajaxContainerLogNotes" class='searchable-notes ajax-loading'></tbody>
		</table>
	</div>
</div>
<%'**********************
' **** eof Notes Tab ****
'************************
%>

<script>
	var LogNoteTypes = <%= LogNoteTypes %>;
	$(document).ready(function () { ajaxLoadLogNotes(); });
	function ajaxRowNewLogNotes() {
		var value = {};
		value.id = 0;
		value.LogDetailType = "-";
		value.Date = "-";
		value.Time = "-";
		value.User = "-";
		value.LogNoteTypeNumber = "";
		value.LogNote = "";
		value.Sticky = "0";
		$('#ajaxRowLogNotes-' + 0 + '').remove();		
		$("#ajaxContainerLogNotes").prepend(ajaxRowHtmlNotes(value));
	}
	function ajaxRowHtmlNotes(value) {
		var LogNoteTypesSelect = '<select class="form-control" data-type="LogNoteTypeNumber">';
		$.each(LogNoteTypes, function (key, LogNoteType) {
			LogNoteTypesSelect+='<option value="'+LogNoteType.id+'" ' + (value.LogNoteTypeNumber+""==LogNoteType.id+""?'selected':'') + '>'+LogNoteType.title+'</option>';
		});
		LogNoteTypesSelect+='</select>';
		var btns = '\
					<div class="visibleRowView btn-group btn-group-sm"><a class="btn btn-primary" onclick="ajaxRowMode(\'LogNotes\', ' + value.id + ', \'Edit\');"><i class="fa fa-edit"></i></a><a class="btn btn-danger" onclick="ajaxLoadLogNotes(\'delete\', ' + value.id + ',\'' + value.LogDetailType + '\');"><i class="fas fa-trash-alt"></i></a></div>\
					<div class="visibleRowEdit btn-group btn-group-sm"><a class="btn btn-success" onclick="ajaxLoadLogNotes(\'save\', ' + value.id + ',\'' + value.LogDetailType + '\');"><i class="fa fa-save"></i></a><a class="btn btn-default" onclick="ajaxRowMode(\'LogNotes\', ' + value.id + ', \'View\');"><i class="fa fa-times"></i></a></div>\
				';
		if(value.id==0)
			btns = '\
					<div class="visibleRowEdit btn-group btn-group-sm"><a class="btn btn-success" onclick="ajaxLoadLogNotes(\'insert\', ' + value.id + ',\'' + value.LogDetailType + '\');"><i class="fa fa-save"></i></a><a class="btn btn-default" onclick="ajaxRowMode(\'LogNotes\', ' + value.id + ', \'View\');"><i class="fa fa-times"></i></a></div>\
				';
		
		//If Log note type is note
		if (value.LogDetailType == '-') {
			var html = '\
				<tr id="ajaxRowLogNotes-' + value.id + '" class="' + (value.id==0?'ajaxRowEdit':'ajaxRowView') + ' note">\
					<td><i class="fa fa-sticky-note" aria-hidden="true"></i>&nbsp;Note</td>\
					<td>' + value.Date + '</td>\
					<td>' + value.Time + '</td>\
					<td>' + value.User + '</td>\
					<td>\
						<div class="visibleRowView">' + value.LogNoteType + '</div>\
						<div class="visibleRowEdit">'+ LogNoteTypesSelect + '</div>\
					</td>\
					<td>\
						<div class="visibleRowView">' + value.LogNote + '</div>\
						<div class="visibleRowEdit"><input class="form-control" data-type="LogNote" value="' + value.LogNote.replace(/"/g, '&quot;') + '" /></div>\
					</td>\
					<td>\
						<div class="visibleRowView"><a onclick="ajaxLoadLogNotes(\'Sticky-' + (value.Sticky==1?'0':'1') + '\', ' + value.id + ',\'' + value.LogDetailType + '\');" class="label label-' + (value.Sticky==1?'success':'danger') + '">' + (value.Sticky==1?'Yes':'No') + '</a></div>\
						<div class="visibleRowEdit"><input type="checkbox" data-type="Sticky" ' + (value.Sticky==1?'checked':'') + '></div>\
					</td>\
					<td class="text-center">'+btns+'</td>\
			   </tr>\
				';
		}
		
		else if (value.LogDetailType == 'Note') {
			var html = '\
				<tr id="ajaxRowLogNotes-' + value.id + '" class="' + (value.id==0?'ajaxRowEdit':'ajaxRowView') + ' note">\
					<td><i class="fa fa-sticky-note" aria-hidden="true"></i>&nbsp;' + value.LogDetailType + '</td>\
					<td>' + value.Date + '</td>\
					<td>' + value.Time + '</td>\
					<td>' + value.User + '</td>\
					<td>\
						<div class="visibleRowView">' + value.LogNoteType + '</div>\
						<div class="visibleRowEdit">'+ LogNoteTypesSelect + '</div>\
					</td>\
					<td>\
						<div class="visibleRowView">' + value.LogNote + '</div>\
						<div class="visibleRowEdit"><input class="form-control" data-type="LogNote" value="' + value.LogNote.replace(/"/g, '&quot;') + '" /></div>\
					</td>\
					<td>\
						<div class="visibleRowView"><a onclick="ajaxLoadLogNotes(\'Sticky-' + (value.Sticky==1?'0':'1') + '\', ' + value.id + ',\'' + value.LogDetailType + '\');" class="label label-' + (value.Sticky==1?'success':'danger') + '">' + (value.Sticky==1?'Yes':'No') + '</a></div>\
						<div class="visibleRowEdit"><input type="checkbox" data-type="Sticky" ' + (value.Sticky==1?'checked':'') + '></div>\
					</td>\
					<td class="text-center">'+btns+'</td>\
			   </tr>\
				';
		}

		//If Log note type is Email
		else if (value.LogDetailType == 'Email') {
			value.LogNoteType = 'Email';
			var html = '\
				<tr id="ajaxRowLogNotes-' + value.id + '" class="' + (value.id==0?'ajaxRowEdit':'ajaxRowView') + ' email">\
					<td><i class="fa fa-envelope" aria-hidden="true"></i>&nbsp;' + value.LogDetailType + '</td>\
					<td>' + value.Date + '</td>\
					<td>' + value.Time + '</td>\
					<td>' + value.User + '</td>\
					<td>\
						<div class="visibleRowView">' + value.LogNoteType + '</div>\
						<div class="visibleRowEdit">' + value.LogNoteType + '</div>\
					</td>\
					<td>\
						<div class="visibleRowView">' + value.LogNote + '</div>\
						<div class="visibleRowEdit"' + value.LogNote + '</div>\
					</td>\
					<td>\
						<div class="visibleRowView"><a onclick="ajaxLoadLogNotes(\'Sticky-' + (value.Sticky==1?'0':'1') + '\', ' + value.id + ',\'' + value.LogDetailType + '\');" class="label label-' + (value.Sticky==1?'success':'danger') + '">' + (value.Sticky==1?'Yes':'No') + '</a></div>\
						<div class="visibleRowEdit"><input type="checkbox" data-type="Sticky" ' + (value.Sticky==1?'checked':'') + '></div>\
					</td>\
					<td class="text-center"><a class="btn btn-primary btn-sm" data-toggle="modal" data-show="true" data-target="#myEmailModal' + value.id + '" href="viewProspect_displayFullEmailModal.asp?i=' + value.id + '"><i class="fa fa-envelope-open" aria-hidden="true"></i></a></td>\
			   </tr>\
				';
		}
		//If Log note type is Stage Change
		else if (value.LogDetailType == 'Stage Change') {
			value.LogNoteType = 'Stage Change';
			var html = '\
				<tr id="ajaxRowLogNotes-' + value.id + '" class="' + (value.id==0?'ajaxRowEdit':'ajaxRowView') + ' stagechange">\
					<td><i class="fa fa-tasks" aria-hidden="true"></i>&nbsp;' + value.LogDetailType + '</td>\
					<td>' + value.Date + '</td>\
					<td>' + value.Time + '</td>\
					<td>' + value.User + '</td>\
					<td>\
						<div class="visibleRowView">' + value.LogNoteType + '</div>\
						<div class="visibleRowEdit">' + value.LogNoteType + '</div>\
					</td>\
					<td>\
						<div class="visibleRowView">' + value.LogNote + '</div>\
						<div class="visibleRowEdit"' + value.LogNote + '</div>\
					</td>\
					<td>\
						<div class="visibleRowView">-</div>\
						<div class="visibleRowEdit">-</div>\
					</td>\
					<td class="text-center">&nbsp;</td>\
			   </tr>\
				';
		}
			
		//else log note type is activity 
		else {
			value.LogNoteType = 'Activity';
			var html = '\
				<tr id="ajaxRowLogNotes-' + value.id + '" class="' + (value.id==0?'ajaxRowEdit':'ajaxRowView') + ' activity">\
					<td><i class="fa fa-check-square" aria-hidden="true"></i>&nbsp;' + value.LogDetailType + '</td>\
					<td>' + value.Date + '</td>\
					<td>' + value.Time + '</td>\
					<td>' + value.User + '</td>\
					<td>\
						<div class="visibleRowView">' + value.LogNoteType + '</div>\
						<div class="visibleRowEdit">' + value.LogNoteType + '</div>\
					</td>\
					<td>\
						<div class="visibleRowView">' + value.LogNote + '</div>\
						<div class="visibleRowEdit"' + value.LogNote + '</div>\
					</td>\
					<td>\
						<div class="visibleRowView">-</div>\
						<div class="visibleRowEdit">-</div>\
					</td>\
					<td class="text-center">&nbsp;</td>\
			   </tr>\
				';
		}			
		
		return html;
	}
	function ajaxLoadLogNotes(updateAction, updateActionId, updateLogNoteType) {

		if (updateAction == "delete" && !confirm("Are you sure you want to delete this log note?")) return;
		$("#ajaxContainerLogNotes").addClass("ajax-loading");
		var url = "ajax/pr_log.asp?i=<%= InternalRecordIdentifier %>";
		var jsondata = {};
		jsondata.updateAction = updateAction;
		jsondata.updateActionId = updateActionId;
		jsondata.updateLogNoteType = updateLogNoteType;
		
		//alert(updateLogNoteType);
		
		if(updateAction=="save" || updateAction=="insert"){
			jsondata.LogNoteTypeNumber = $('#ajaxRowLogNotes-' + updateActionId + ' [data-type="LogNoteTypeNumber"]').val();
			jsondata.LogNote= $('#ajaxRowLogNotes-' + updateActionId + ' [data-type="LogNote"]').val();
			jsondata.Sticky = $('#ajaxRowLogNotes-' + updateActionId + ' [data-type="Sticky"]').is(':checked')?1:0;
		}
		$.ajax({
			type: "POST",
			url: url,
			dataType: "json",
			data: jsondata,
			success: function (data) {
				//if (!data || data+""=="") { window.location.href = window.location.href + ""; return; }				
				var html = "";
				$.each(data, function (key, value) {
					html += ajaxRowHtmlNotes(value);
				});
				$("#ajaxContainerLogNotes").html(html);
				setTimeout(function(){
					$("#ajaxContainerLogNotes").removeClass("ajax-loading");
				}, 0);
				
			}
		});
	}
</script>


<%

SQLEmailLog = "SELECT * FROM PR_ProspectEmailLog WHERE ProspectRecID='"& InternalRecordIdentifier &"'"
Set cnnEmailLog = Server.CreateObject("ADODB.Connection")
cnnEmailLog.open (Session("ClientCnnString"))
Set rsEmailLog = Server.CreateObject("ADODB.Recordset")
rsEmailLog.CursorLocation = 3 
Set rsEmailLog = cnnEmailLog.Execute(SQLEmailLog)

If not rsEmailLog.EOF Then

	Do While Not rsEmailLog.EOF

%>

	<!-- modal  starts here !-->
	 <!-- Modal -->
	<div class="modal fade" id="myEmailModal<%= rsEmailLog("InternalRecordIdentifier") %>" tabindex="-1" role="dialog" aria-labelledby="myEmailModalLabel<%= rsEmailLog("InternalRecordIdentifier") %>" aria-hidden="true">
	    <div class="modal-dialog">
	        <div class="modal-content">
	            <div class="modal-body"></div>
	        </div>
	        <!-- /.modal-content -->
	    </div>
	    <!-- /.modal-dialog -->
	</div>
	<!-- /.modal -->
	<!-- modal  ends here !-->
<%							
		rsEmailLog.MoveNext						
	Loop
End If

Set rsEmailLog = Nothing
cnnEmailLog.Close
Set cnnEmailLog = Nothing


%> 

