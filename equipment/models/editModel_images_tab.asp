<div role="tabpanel" class="tab-pane fade in" id="images">

	<p> <button type="button" class="btn btn-success" onclick="ajaxRowNewImageAttachment();">New Model Image</button> </p>

		<div class="input-group narrow-results"> <span class="input-group-addon">Narrow Results</span>
		    <input id="filter-images" type="text" class="form-control filter-search-width" placeholder="Type here...">
		</div>
	  
		<div class="table-responsive">
        <form id="imageuploadform" action="ajax/eq_images_upload.asp" method="POST" ENCTYPE="multipart/form-data">
        
	        <input type="hidden" name="i" value="<%= ModelIntRecID %>" />
	        <input type="hidden" id="fileuploadstatusimages" />
	        <input type="hidden" name="updateActionImages" id="updateActionImages" value="" />
	        <input type="hidden" name="updateActionIdImages" id="updateActionIdImages" value="" />
	        
            <table class="table table-striped table-condensed table-hover table-bordered sortable standard-font">
              <thead>
                <tr>
                  <th width="5%">Image</th>
                  <th width="5%">Image Type</th>
                  <th width="5%">Date</th>
				  <th width="5%">Time</th>
				  <th width="8%">Attached By</th>
				  <th>Image Notes</th>
				  <th>Image Attachment</th>
  				  <th width="3%">Sticky</th>
                  <th class="sorttable_nosort text-center" style="width: 80px;">Actions</th>
                </tr>
              </thead>
			<tbody id="ajaxContainerImageAttachment" class='searchable-images ajax-loading'></tbody>
		</table>
        </form> 
	</div>
</div>
<%'**********************
' **** eof Notes Tab ****
'************************
%>

<script>
	
    function ajaxRowNewImageAttachment() {
        var value = {};
        value.id = 0;
        value.Date = "-";
        value.Time = "-";
        value.User = "-";
        value.ImageNotes = "";
        value.ImageAttachment = "";
        value.ImagePath = "";
        value.ImageExt = "";
        value.Sticky = 0;
        $('#ajaxRowImageAttachment-' + value.id + '').remove();
        $("#ajaxContainerImageAttachment").prepend(ajaxRowHtmlImages(value));
        fileUploadInitImage("insert", value.id);
    }

    function fileUploadInitImage(action, id){
        if ($('#fileuploadstatusimages').val() == 'active') {  
            $('#imageuploadform').fileupload('destroy'); 
        }
        $("#updateActionImages").val(action);
        $("#updateActionIdImages").val(id);
      
        $('#imageuploadform').fileupload({
            dataType: "json",
            paramName: "ImageAttachment",
            replaceFileInput: false,
            add: function (e, data) {
                $('.lnkUploadImage').click(function () {
                    $("#ajaxContainerImageAttachment").addClass("ajax-loading");
                    $('div.visibleRowEdit:hidden > input[name=ImageNotes]').attr("disabled",true);
                    $('div.visibleRowEdit:hidden > input[name=Sticky]').attr("disabled",true);
                    data.submit();
                });
            },
            done: function (e, data) {
                ajaxLoadImageAttachment();
                $("#ajaxContainerImageAttachment").removeClass("ajax-loading");
                $('div.visibleRowEdit:hidden > input[name=ImageNotes]').attr("disabled",false);
                $('div.visibleRowEdit:hidden > input[name=Sticky]').attr("disabled",false);                
            }
        });
        $('#fileuploadstatusimages').val('active');
    }
		
	function ajaxRowHtmlImages(value) {
		var btns = '\
					<div class="visibleRowView btn-group btn-group-sm"><a class="btn btn-primary" onclick="ajaxRowMode(\'ImageAttachment\', ' + value.id + ', \'Edit\');fileUploadInitImage(\'save\',' + value.id + ');"><i class="fa fa-edit"></i></a><a class="btn btn-danger" onclick="ajaxLoadImageAttachment(\'delete\', ' + value.id + ');"><i class="fas fa-trash-alt"></i></a></div>\
					<div class="visibleRowEdit btn-group btn-group-sm"><a class="btn btn-success lnkUploadImage"><i class="fa fa-save"></i></a><a class="btn btn-default" onclick="ajaxRowMode(\'ImageAttachment\', ' + value.id + ', \'View\');"><i class="fa fa-times"></i></a></div>\
				';
		if(value.id==0)
			btns = '\
					<div class="visibleRowEdit btn-group btn-group-sm"><a class="btn btn-success lnkUploadImage"><i class="fa fa-save"></i></a><a class="btn btn-default" onclick="ajaxRowMode(\'ImageAttachment\', ' + value.id + ', \'View\');"><i class="fa fa-times"></i></a></div>\
				';	

		if (value.ImageExt !== "") {				
			var htmlImages = '\
				<tr id="ajaxRowImageAttachment-' + value.id + '" class="' + (value.id==0?'ajaxRowEdit':'ajaxRowView') + '">\
					<td><img src="ResizeImage.aspx?ImgHt=50&amp;IptFl=' + value.ImagePath + '"/></td>\
					<td><img src="../../../img/file-icons/' + value.ImageExt + '.png" class="fileicon">&nbsp;' + value.ImageExt + '</td>\
					<td>' + value.Date + '</td>\
					<td>' + value.Time + '</td>\
					<td>' + value.User + '</td>\
					<td>\
						<div class="visibleRowView">' + value.ImageNotes + '</div>\
						<div class="visibleRowEdit"><input class="form-control" data-type="ImageNotes" id="ImageNotes" name="ImageNotes" value="' + value.ImageNotes + '" /></div>\
					</td>\
					<td>\
						<div class="visibleRowView"><a href="' + value.ImagePath + '" target="_blank">' + value.ImageAttachment + '</a></div>\
						<div class="visibleRowEdit"><a href="' + value.ImagePath + '" target="_blank">' + value.ImageAttachment + '</a></div>\
					</td>\
					<td>\
						<div class="visibleRowView"><a onclick="ajaxLoadImageAttachment(\'Sticky-' + (value.Sticky == 1 ? '0' : '1') + '\', ' + value.id + ');" class="label label-' + (value.Sticky == 1 ? 'success' : 'danger') + '">' + (value.Sticky == 1 ? 'Yes' : 'No') + '</a></div>\
						<div class="visibleRowEdit"><input type="checkbox" name="Sticky" data-type="Sticky" ' + (value.Sticky==1?'checked':'') + '></div>\
					</td>\
					<td class="text-center">'+btns+'</td>\
			   </tr>\
				';	
			return htmlImages;
		}
		else {
			var htmlImages = '\
			<tr id="ajaxRowImageAttachment-' + value.id + '" class="' + (value.id==0?'ajaxRowEdit':'ajaxRowView') + ' note">\
				<td><i class="fa fa-image" aria-hidden="true"></i>&nbsp;Image</td>\
				<td>&nbsp;</td>\
				<td>' + value.Date + '</td>\
				<td>' + value.Time + '</td>\
				<td>' + value.User + '</td>\
				<td>\
					<div class="visibleRowView">' + value.ImageNotes + '</div>\
					<div class="visibleRowEdit"><input class="form-control" data-type="ImageNotes"  id="ImageNotes" name="ImageNotes" value="' + value.ImageNotes + '" /></div>\
				</td>\
				<td>\
					<div class="visibleRowView"><a href="' + value.ImagePath + '" target="_blank">' + value.ImageAttachment + '</a></div>\
					<div class="visibleRowEdit"><input type="file" class="form-control" name="ImageAttachment" id="ImageAttachment" data-type="ImageAttachment" value="' + value.ImageAttachment + '" /></div>\
				</td>\
				<td>\
					<div class="visibleRowView"><a onclick="ajaxLoadImageAttachment(\'Sticky-' + (value.Sticky == 1 ? '0' : '1') + '\', ' + value.id + ');" class="label label-' + (value.Sticky == 1 ? 'success' : 'danger') + '">' + (value.Sticky == 1 ? 'Yes' : 'No') + '</a></div>\
					<div class="visibleRowEdit"><input type="checkbox" name="Sticky" data-type="Sticky" ' + (value.Sticky==1?'checked':'') + '></div>\
				</td>\
				<td class="text-center">'+btns+'</td>\
		   </tr>\
			';	
		return htmlImages;
		}
	}

	$(document).ready(function () { 
	
		ajaxLoadImageAttachment(); 
		
		$("#ajaxContainerImageAttachment").delegate('#ImageAttachment', 'change', function(){
		
			var ext = $('#ImageAttachment').val().split('.').pop().toLowerCase();
						
			if($.inArray(ext, ['gif','png','jpg','jpeg','bmp','tiff']) < 0) {
			   swal('Please Add Images Only');
			   $('#ImageAttachment').val('');
			   ajaxRowMode("ImageAttachment","0","View");
			}
			
		    
		});	
	
	});

	function ajaxLoadImageAttachment(updateActionImages, updateActionIdImages) {
	
	    if (updateActionImages == "delete" && !confirm("Are you sure you want to delete this image?")) return;
	    $("#ajaxContainerImageAttachment").addClass("ajax-loading");
	    var url = "ajax/eq_images.asp?i=<%= ModelIntRecID %>";
	    var jsondata = {};
	    jsondata.updateActionImages = updateActionImages;
	    jsondata.updateActionIdImages = updateActionIdImages;
	    if (updateActionImages == "save" || updateActionImages == "insert") {
	        jsondata.ImageNotes = $('#ajaxRowImageAttachment-' + updateActionIdImages + ' [data-type="ImageNotes"]').val();
	        jsondata.ImageAttachment = $('#ajaxRowImageAttachment-' + updateActionIdImages + ' [data-type="ImageAttachment"]').val();
	        jsondata.Sticky = $('#ajaxRowImageAttachment-' + updateActionIdImages + ' [data-type="Sticky"]').is(':checked') ? 1 : 0;
	        //alert(jsondata.ImageNotes + " " + jsondata.ImageAttachment);
	    }
	
	    $.ajax({
	        type: "POST",
	        url: url,
	        dataType: "json",
	        data: jsondata,
	        success: function (data) {		
	            var htmlImages = "";
	            $.each(data, function (key, value) {
	                htmlImages += ajaxRowHtmlImages(value);
	            });
	            $("#ajaxContainerImageAttachment").html(htmlImages);
	            setTimeout(function () {
	                $("#ajaxContainerImageAttachment").removeClass("ajax-loading");
	            }, 0);
	            getNumberOfImages();
	        }
	    });
	}

	function getNumberOfImages() {
	    var url = "ajax/eq_get_images_number.asp?i=<%= ModelIntRecID %>";
	    $.ajax({
	        type: "POST",
	        url: url,
	        success: function (data) {
	            $("#imagesNum").html(data);
	        }
	    });
	}


</script>
