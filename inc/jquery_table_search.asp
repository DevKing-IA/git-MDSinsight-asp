<!-- custom table search !-->

<script>
	$(document).ready(function () {

    (function ($) {

        $('#filter').keyup(function () {

            var rex = new RegExp($(this).val(), 'i');
            $('.searchable tr').hide();
            $('.searchable tr').filter(function () {
                return rex.test($(this).text());
            }).show();

        })

    }(jQuery));
    
    
    (function ($) {

       $('#filterProspects').keyup(function () {

            var rex = new RegExp($(this).val(), 'i');
            $('.searchable tr').hide();
            
            $('.searchable tr').filter(function () {
                //return rex.test($(this).text());
            }).show();
            
            var numOfVisibleRows = $('tr:visible').length - 1;
            document.getElementById("TotalNumberOfProspects").innerHTML = "<strong>Currently Viewing " + numOfVisibleRows + " Total Prospects</strong>";


        })

   }(jQuery));
  

  


});
</script>
<!-- eof custom table search !-->