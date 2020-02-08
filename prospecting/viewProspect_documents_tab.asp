<div role="tabpanel" class="tab-pane fade in" id="documents">

	<p> <button type="button" class="btn btn-success" onclick="ajaxRowNewDocumentAttachment();">New Document Attachment</button> </p>

		<div class="input-group narrow-results"> <span class="input-group-addon">Narrow Results</span>
		    <input id="filter-documents" type="text" class="form-control filter-search-width" placeholder="Type here...">
		</div>
	  <form>
		<div class="table-responsive">
            <table class="table table-striped table-condensed table-hover table-bordered sortable standard-font">
              <thead>
                <tr>
                  <th width="5%">&nbsp;</th>
                  <th width="5%">Date</th>
				  <th width="5%">Time</th>
				  <th width="8%">Attached By</th>
				  <th>Notes</th>
				  <th>File Attachment</th>
  				  <th width="3%">Sticky</th>
                  <th class="sorttable_nosort text-center" style="width: 80px;">Actions</th>
                </tr>
              </thead>
			<tbody id="ajaxContainerDocumentAttachment" class='searchable-documents ajax-loading'></tbody>
		</table>
	</div>
    </form>
</div>
<%'**********************
' **** eof Notes Tab ****
'************************
%>

<script>
	$(document).ready(function () { 
		ajaxLoadDocumentAttachment(); 	
	});
	
	function ajaxRowNewDocumentAttachment() {
		var value = {};
		value.id = 0;
		value.Date = "-";
		value.Time = "-";
		value.User = "-";
		value.DocumentNotes = "";
		value.DocumentAttachment = "";
		value.DocumentPath = "";
		value.DocumentExt = "";
		value.Sticky = "0";
		$('#ajaxRowDocumentAttachment-' + 0 + '').remove();		
		$("#ajaxContainerDocumentAttachment").prepend(ajaxRowHtmlDocuments(value));
	}
	function ajaxRowHtmlDocuments(value) {
		var btns = '\
					<div class="visibleRowView btn-group btn-group-sm"><a class="btn btn-primary" onclick="ajaxRowMode(\'DocumentAttachment\', ' + value.id + ', \'Edit\');"><i class="fa fa-edit"></i></a><a class="btn btn-danger" onclick="ajaxLoadDocumentAttachment(\'delete\', ' + value.id + ');"><i class="fas fa-trash-alt"></i></a></div>\
					<div class="visibleRowEdit btn-group btn-group-sm"><a class="btn btn-success" onclick="ajaxLoadDocumentAttachment(\'save\', ' + value.id + ');"><i class="fa fa-save"></i></a><a class="btn btn-default" onclick="ajaxRowMode(\'DocumentAttachment\', ' + value.id + ', \'View\');"><i class="fa fa-times"></i></a></div>\
				';
		if(value.id==0)
			btns = '\
					<div class="visibleRowEdit btn-group btn-group-sm"><a class="btn btn-success" onclick="ajaxLoadDocumentAttachment(\'insert\', ' + value.id + ');"><i class="fa fa-save"></i></a><a class="btn btn-default" onclick="ajaxRowMode(\'DocumentAttachment\', ' + value.id + ', \'View\');"><i class="fa fa-times"></i></a></div>\
				';	

		if (value.DocumentExt !== "") {				
			var html = '\
				<tr id="ajaxRowDocumentAttachment-' + value.id + '" class="' + (value.id==0?'ajaxRowEdit':'ajaxRowView') + '">\
					<td><img src="../../img/file-icons/' + value.DocumentExt + '.png" class="fileicon">&nbsp;' + value.DocumentExt + '</td>\
					<td>' + value.Date + '</td>\
					<td>' + value.Time + '</td>\
					<td>' + value.User + '</td>\
					<td>\
						<div class="visibleRowView">' + value.DocumentNotes + '</div>\
						<div class="visibleRowEdit"><input class="form-control" data-type="DocumentNotes" value="' + value.DocumentNotes + '" /></div>\
					</td>\
					<td>\
						<div class="visibleRowView"><a href="' + value.DocumentPath + '" target="_blank">' + value.DocumentAttachment + '</a></div>\
						<div class="visibleRowEdit"><input id="file'+value.id+'" type="file" class="form-control" data-type="DocumentAttachment" value="' + value.DocumentAttachment + '" /></div>\
					</td>\
					<td>\
						<div class="visibleRowView"><a onclick="ajaxLoadDocumentAttachment(\'Sticky-' + (value.Sticky==1?'0':'1') + '\', ' + value.id + ');" class="label label-' + (value.Sticky==1?'success':'danger') + '">' + (value.Sticky==1?'Yes':'No') + '</a></div>\
						<div class="visibleRowEdit"><input type="checkbox" data-type="Sticky" ' + (value.Sticky==1?'checked':'') + '></div>\
					</td>\
					<td class="text-center">'+btns+'</td>\
			   </tr>\
				';	
			return html;
		}
		else {
			var html = '\
			<tr id="ajaxRowDocumentAttachment-' + value.id + '" class="' + (value.id==0?'ajaxRowEdit':'ajaxRowView') + ' note">\
				<td><i class="fa fa-sticky-note" aria-hidden="true"></i>&nbsp;Note</td>\
				<td>' + value.Date + '</td>\
				<td>' + value.Time + '</td>\
				<td>' + value.User + '</td>\
				<td>\
					<div class="visibleRowView">' + value.DocumentNotes + '</div>\
					<div class="visibleRowEdit"><input class="form-control" data-type="DocumentNotes" value="' + value.DocumentNotes + '" /></div>\
				</td>\
				<td>\
					<div class="visibleRowView"><a href="' + value.DocumentPath + '" target="_blank">' + value.DocumentAttachment + '</a></div>\
					<div class="visibleRowEdit"><input id="file'+value.id+'" type="file" class="form-control" data-type="DocumentAttachment" value="' + value.DocumentAttachment + '" /></div>\
				</td>\
				<td>\
					<div class="visibleRowView"><a onclick="ajaxLoadDocumentAttachment(\'Sticky-' + (value.Sticky==1?'0':'1') + '\', ' + value.id + ');" class="label label-' + (value.Sticky==1?'success':'danger') + '">' + (value.Sticky==1?'Yes':'No') + '</a></div>\
					<div class="visibleRowEdit"><input type="checkbox" data-type="Sticky" ' + (value.Sticky==1?'checked':'') + '></div>\
				</td>\
				<td class="text-center">'+btns+'</td>\
		   </tr>\
			';	
		return html;
		}
	}
	function ajaxLoadDocumentAttachment(updateAction, updateActionId) {
		
		
		
		var formData = new FormData();
		
		if (updateAction == "delete" && !confirm("Are you sure you want to delete this document?")) return;
		$("#ajaxContainerDocumentAttachment").addClass("ajax-loading");
		var url = "ajax/pr_documents.asp?i=<%= InternalRecordIdentifier %>";
		
		//var url = "ajax/upload_test.asp";
		
		/*
		var jsondata = {};		
		jsondata.updateAction = updateAction;
		jsondata.updateActionId = updateActionId;
		
		$("#ajaxContainerDocumentAttachment").removeClass("ajax-loading");
		*/
		
		formData.append("updateAction",updateAction);
		formData.append("updateActionId",updateActionId);

		
		if(updateAction=="save" || updateAction=="insert"){
			//jsondata.DocumentNotes= $('#ajaxRowDocumentAttachment-' + updateActionId + ' [data-type="DocumentNotes"]').val();
			//jsondata.DocumentAttachment= $('#ajaxRowDocumentAttachment-' + updateActionId + ' [data-type="DocumentAttachment"]').val();
			//jsondata.Sticky = $('#ajaxRowDocumentAttachment-' + updateActionId + ' [data-type="Sticky"]').is(':checked')?1:0;
			//alert("111"+ jsondata.DocumentNotes+ " " + jsondata.DocumentAttachment);
			var Sticky = $('#ajaxRowDocumentAttachment-' + updateActionId + ' [data-type="Sticky"]').is(':checked')?1:0;
			formData.append("DocumentNotes",$('#ajaxRowDocumentAttachment-' + updateActionId + ' [data-type="DocumentNotes"]').val());
			formData.append("Sticky",Sticky);
			
			if ($('#ajaxRowDocumentAttachment-' + updateActionId + ' [data-type="DocumentAttachment"]').val()!=""){
				var filefield = $('#ajaxRowDocumentAttachment-' + updateActionId + ' [data-type="DocumentAttachment"]').prop("files")[0];		
				formData.append("DocumentAttachment", filefield, filefield.name);
			}
			
		
		//alert(updateAction + " - " + updateActionId + " = " +$('#ajaxRowDocumentAttachment-' + updateActionId + ' [data-type="Sticky"]').is(':checked')?1:0);
		//$("#avatar").prop("files")[0];
		//alert($('#ajaxRowDocumentAttachment-' + updateActionId + ' [data-type="Sticky"]').is(':checked')?1:0)
		
		}
		
		
		$.ajax({
			type: "POST",
			url: url,
			//dataType: "json",
			//data: jsondata,
			async: true,
       	 	data: formData,
        	cache: false,
        	contentType: false,
        	processData: false,
		
			success: function (data) {

				//if (!data || data+""=="") { window.location.href = window.location.href + ""; return; }		
				var html = "";
				$.each(data, function (key, value) {
					html += ajaxRowHtmlDocuments(value);
				});
				$("#ajaxContainerDocumentAttachment").html(html);
				setTimeout(function(){
					$("#ajaxContainerDocumentAttachment").removeClass("ajax-loading");
				}, 0);
				
			},
			error: function (error) {
				// handle error
				alert("error:" +error.responseText);
			}
			
		});
		
	}
</script>
