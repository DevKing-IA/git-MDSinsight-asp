<div role="tabpanel" class="tab-pane fade in active" id="documents">

	<p> <button type="button" class="btn btn-success" onclick="ajaxRowNewDocumentAttachment();">New Document Attachment</button> </p>

		<div class="input-group narrow-results"> <span class="input-group-addon">Narrow Results</span>
		    <input id="filter-documents" type="text" class="form-control filter-search-width" placeholder="Type here...">
		</div>
	  
		<div class="table-responsive">
        <form id="fileuploadform" action="ajax/eq_documents_upload.asp" method="POST" ENCTYPE="multipart/form-data">
        
	        <input type="hidden" name="i" value="<%= ModelIntRecID %>" />
	        <input type="hidden" id="fileuploadstatus" />
	        <input type="hidden" name="updateAction" id="updateAction" value="" />
	        <input type="hidden" name="updateActionId" id="updateActionId" value="" />
	        
            <table class="table table-striped table-condensed table-hover table-bordered sortable standard-font">
              <thead>
                <tr>
                  <th width="5%">File Type</th>
                  <th width="5%">Date</th>
				  <th width="5%">Time</th>
				  <th width="8%">Attached By</th>
				  <th>Document Notes</th>
				  <th>File Attachment</th>
  				  <th width="3%">Sticky</th>
                  <th class="sorttable_nosort text-center" style="width: 80px;">Actions</th>
                </tr>
              </thead>
			<tbody id="ajaxContainerDocumentAttachment" class='searchable-documents ajax-loading'></tbody>
		</table>
        </form> 
	</div>
</div>
<%'**********************
' **** eof Notes Tab ****
'************************
%>

<script>
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
        value.Sticky = 0;
        $('#ajaxRowDocumentAttachment-' + value.id + '').remove();
        $("#ajaxContainerDocumentAttachment").prepend(ajaxRowHtmlDocuments(value));
        fileUploadInit("insert", value.id);
    }

    function fileUploadInit(action, id){
        if ($('#fileuploadstatus').val() == 'active') {  
            $('#fileuploadform').fileupload('destroy'); 
        }
        $("#updateAction").val(action);
        $("#updateActionId").val(id);
                
        
        $('#fileuploadform').fileupload({
            dataType: "json",
            paramName: "DocumentAttachment",
            replaceFileInput: false,
            add: function (e, data) {
                $('.lnkUpload').click(function () {           
                    $("#ajaxContainerDocumentAttachment").addClass("ajax-loading");
                    $('div.visibleRowEdit:hidden > input[name=DocumentNotes]').attr("disabled",true);
                    $('div.visibleRowEdit:hidden > input[name=Sticky]').attr("disabled",true);
                    data.submit();
                });
            },
            done: function (e, data) {
                ajaxLoadDocumentAttachment();
                $("#ajaxContainerDocumentAttachment").removeClass("ajax-loading");
                $('div.visibleRowEdit:hidden > input[name=DocumentNotes]').attr("disabled",false);
                $('div.visibleRowEdit:hidden > input[name=Sticky]').attr("disabled",false);                
            }
        });
        $('#fileuploadstatus').val('active');
    }
		
	function ajaxRowHtmlDocuments(value) {
		var btns = '\
					<div class="visibleRowView btn-group btn-group-sm"><a class="btn btn-primary" onclick="ajaxRowMode(\'DocumentAttachment\', ' + value.id + ', \'Edit\');fileUploadInit(\'save\',' + value.id + ');"><i class="fa fa-edit"></i></a><a class="btn btn-danger" onclick="ajaxLoadDocumentAttachment(\'delete\', ' + value.id + ');"><i class="fas fa-trash-alt"></i></a></div>\
					<div class="visibleRowEdit btn-group btn-group-sm"><a class="btn btn-success lnkUpload"><i class="fa fa-save"></i></a><a class="btn btn-default" onclick="ajaxRowMode(\'DocumentAttachment\', ' + value.id + ', \'View\');"><i class="fa fa-times"></i></a></div>\
				';
		if(value.id==0)
			btns = '\
					<div class="visibleRowEdit btn-group btn-group-sm"><a class="btn btn-success lnkUpload"><i class="fa fa-save"></i></a><a class="btn btn-default" onclick="ajaxRowMode(\'DocumentAttachment\', ' + value.id + ', \'View\');"><i class="fa fa-times"></i></a></div>\
				';	

		if (value.DocumentExt !== "") {				
			var html = '\
				<tr id="ajaxRowDocumentAttachment-' + value.id + '" class="' + (value.id==0?'ajaxRowEdit':'ajaxRowView') + '">\
					<td><img src="../../../img/file-icons/' + value.DocumentExt + '.png" class="fileicon">&nbsp;' + value.DocumentExt + '</td>\
					<td>' + value.Date + '</td>\
					<td>' + value.Time + '</td>\
					<td>' + value.User + '</td>\
					<td>\
						<div class="visibleRowView">' + value.DocumentNotes + '</div>\
						<div class="visibleRowEdit"><input class="form-control" data-type="DocumentNotes" id="DocumentNotes" name="DocumentNotes" value="' + value.DocumentNotes + '" /></div>\
					</td>\
					<td>\
						<div class="visibleRowView"><a href="' + value.DocumentPath + '" target="_blank">' + value.DocumentAttachment + '</a></div>\
						<div class="visibleRowEdit"><a href="' + value.DocumentPath + '" target="_blank">' + value.DocumentAttachment + '</a></div>\
					</td>\
					<td>\
						<div class="visibleRowView"><a onclick="ajaxLoadDocumentAttachment(\'Sticky-' + (value.Sticky == 1 ? '0' : '1') + '\', ' + value.id + ');" class="label label-' + (value.Sticky == 1 ? 'success' : 'danger') + '">' + (value.Sticky == 1 ? 'Yes' : 'No') + '</a></div>\
						<div class="visibleRowEdit"><input type="checkbox" name="Sticky" data-type="Sticky" ' + (value.Sticky==1?'checked':'') + '></div>\
					</td>\
					<td class="text-center cell">'+btns+'</td>\
			   </tr>\
				';	
			return html;
		}
		else {
			var html = '\
			<tr id="ajaxRowDocumentAttachment-' + value.id + '" class="' + (value.id==0?'ajaxRowEdit':'ajaxRowView') + ' note">\
				<td><i class="fa fa-sticky-note" aria-hidden="true"></i>&nbsp;Document</td>\
				<td>' + value.Date + '</td>\
				<td>' + value.Time + '</td>\
				<td>' + value.User + '</td>\
				<td>\
					<div class="visibleRowView">' + value.DocumentNotes + '</div>\
					<div class="visibleRowEdit"><input class="form-control" data-type="DocumentNotes"  id="DocumentNotes" name="DocumentNotes" value="' + value.DocumentNotes + '" /></div>\
				</td>\
				<td>\
					<div class="visibleRowView"><a href="' + value.DocumentPath + '" target="_blank">' + value.DocumentAttachment + '</a></div>\
					<div class="visibleRowEdit"><input type="file" class="form-control" name="DocumentAttachment" id="DocumentAttachment" data-type="DocumentAttachment" value="' + value.DocumentAttachment + '" /></div>\
				</td>\
				<td>\
					<div class="visibleRowView"><a onclick="ajaxLoadDocumentAttachment(\'Sticky-' + (value.Sticky == 1 ? '0' : '1') + '\', ' + value.id + ');" class="label label-' + (value.Sticky == 1 ? 'success' : 'danger') + '">' + (value.Sticky == 1 ? 'Yes' : 'No') + '</a></div>\
					<div class="visibleRowEdit"><input type="checkbox" name="Sticky" data-type="Sticky" ' + (value.Sticky==1?'checked':'') + '></div>\
				</td>\
				<td class="text-center">'+btns+'</td>\
		   </tr>\
			';	
		return html;
		}
}


$(document).ready(function () { 

	ajaxLoadDocumentAttachment(); 
	
	$("#ajaxContainerDocumentAttachment").delegate('#DocumentAttachment', 'change', function(){
	
		var ext = $('#DocumentAttachment').val().split('.').pop().toLowerCase();

		if($.inArray(ext, ['gif','png','jpg','jpeg','bmp','tiff']) >= 0) {
		   swal('Please Add Images Under The Images Tab');
		   $('#DocumentAttachment').val('');
		   ajaxRowMode("DocumentAttachment","0","View");
		}
	    
	});	
	

});


function ajaxLoadDocumentAttachment(updateAction, updateActionId) {

    if (updateAction == "delete" && !confirm("Are you sure you want to delete this document?")) return;
    $("#ajaxContainerDocumentAttachment").addClass("ajax-loading");
    var url = "ajax/eq_documents.asp?i=<%= ModelIntRecID %>";
    var jsondata = {};
    jsondata.updateAction = updateAction;
    jsondata.updateActionId = updateActionId;
    if (updateAction == "save" || updateAction == "insert") {
        jsondata.DocumentNotes = $('#ajaxRowDocumentAttachment-' + updateActionId + ' [data-type="DocumentNotes"]').val();
        jsondata.DocumentAttachment = $('#ajaxRowDocumentAttachment-' + updateActionId + ' [data-type="DocumentAttachment"]').val();
        jsondata.Sticky = $('#ajaxRowDocumentAttachment-' + updateActionId + ' [data-type="Sticky"]').is(':checked') ? 1 : 0;
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
                html += ajaxRowHtmlDocuments(value);
            });
            $("#ajaxContainerDocumentAttachment").html(html);
            setTimeout(function () {
                $("#ajaxContainerDocumentAttachment").removeClass("ajax-loading");
            }, 0);
            getNumberOfDocuments();
        }
    });
}

function getNumberOfDocuments() {
    var url = "ajax/eq_get_docs_number.asp?i=<%= ModelIntRecID %>";
    $.ajax({
        type: "POST",
        url: url,
        success: function (data) {
            $("#docsNum").html(data);
        }
    });
}


</script>
