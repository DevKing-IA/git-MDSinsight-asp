
<!-- Resources -->
<script src="https://www.amcharts.com/lib/3/amcharts.js"></script>
<script src="https://www.amcharts.com/lib/3/pie.js"></script>
<script src="https://www.amcharts.com/lib/3/serial.js"></script>
<script src="https://www.amcharts.com/lib/3/plugins/export/export.min.js"></script>

<head>
<link rel="stylesheet" href="https://www.amcharts.com/lib/3/plugins/export/export.css" type="text/css" media="all" />
<script src="https://www.amcharts.com/lib/3/themes/light.js"></script>

<link rel="stylesheet" href="https://cdn.datatables.net/1.10.16/css/jquery.dataTables.min.css" />
<script type="text/javascript" src="https://cdn.datatables.net/1.10.16/js/jquery.dataTables.min.js"></script>
<script type="text/javascript" src="https://cdn.datatables.net/plug-ins/1.10.18/sorting/currency.js"></script>

<script type="text/javascript">

$(document).ready(function() {

    $("#PleaseWaitPanel").hide();
    
 
    $('#tableSuperSumCustType').DataTable({
	    searching: false,
	    scrollY: 500,
	    scrollCollapse: true,
	    paging: false,
	    info: false,
        order: [ 1, 'desc'],
		columnDefs: [
		        { type: 'currency', targets: 3}
		    ]	
    });
   
    $('#tableSuperSumPrimarySlsmn').DataTable({
		"searching": false,
	  	"scrollY": 500,
	    "scrollCollapse": true,
	    "paging": false,
	    "info": false,		    
        "order": [ 1, 'desc'],
        "orderCellsTop": false,
        "autoWidth":true,
		"columnDefs": [
	    	{"type": "currency", "targets": [1,2,3,4,5,6,7,8,9,10,11,12,13,14,15]},
	    	{"orderable": true, "targets": [0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15]},
	    	{"className": "dt-center pct10", "targets": [0]},
	    	{"className": "dt-center pct6", "targets": [1,2,3,4,5,6,7,8,9,10,11,12,13,14,15]}
		  ]	        
    });
   
    
    $('#tableSuperSumSecondarySlsmn').DataTable({
		"searching": false,
	  	"scrollY": 500,
	    "scrollCollapse": true,
	    "paging": false,
	    "info": false,		    
        "order": [ 1, 'desc'],
        "orderCellsTop": false,
        "autoWidth":false,
		"columnDefs": [
	    	{"type": "currency", "targets": [1,2,3,4,5,6,7,8,9,10,11,12,13,14,15]},
	    	{"orderable": true, "targets": [0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15]},
	    	{"className": "dt-center pct10", "targets": [0]},
	    	{"className": "dt-center pct6", "targets": [1,2,3,4,5,6,7,8,9,10,11,12,13,14,15]}
		  ]	        
    });
    

    
    $('#tableSuperSumReferral').DataTable({
		"searching": false,
	  	"scrollY": 500,
	    "scrollCollapse": true,
	    "paging": false,
	    "info": false,		    
        "order": [ 1, 'desc'],
        "orderCellsTop": false,
        "autoWidth":true,
		"columnDefs": [
	    	{"type": "currency", "targets": [1,2,3,4,5,6,7,8,9,10,11,12,13,14,15]},
	    	{"orderable": true, "targets": [0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15]},
	    	{"className": "dt-center pct10", "targets": [0]},
	    	{"className": "dt-center pct6", "targets": [1,2,3,4,5,6,7,8,9,10,11,12,13,14,15]}
		  ]	        
    });
    
	$('a[data-toggle="tab"]').on('shown.bs.tab', function (e) {
	
	  var target = $(e.target).attr("href") // activated tab
	  //alert(target);
	  
	  if (target == "#referral") {
	  	$("#tableSuperSumReferral").resize();
	  }

	  if (target == "#custtype") {
	  	$("#tableSuperSumCustType").resize();
	  }

	  if (target == "#sls1") {
	  	$("#tableSuperSumPrimarySlsmn").resize();
	  }

	  if (target == "#sls2") {
	  	$("#tableSuperSumSecondarySlsmn").resize();
	  }
	  
	});    
	    
    
});
</script>


<style type="text/css">
.auto-style1 {
	text-align: right;
}
</style>
</head>


<!-- Chart code -->
<script>

	
	var chart = AmCharts.makeChart( "chartdivRef", {
	
		"titles": [
				{
					"text": "Referral Code Leakage for <%=LCP_Display_Var%>",
					"size": 18,
					"bold": "false"
				}				
			],
        "numberFormatter": {
            "precision": 2,
            "decimalSeparator": ".",
            "thousandsSeparator": ","
        },			
		"creditsPosition":"bottom-left",
		"labelRadius": -50,
		"labelText": "[[referral]]",
		"type": "pie",
		"theme": "light",	
		"dataProvider": [ <%=amChartDataReferral%> ],
		"valueField": "contribPercent",
		"titleField": "referral",
		"maxLabelWidth" : "100",
		"autoMargins": false,
		"marginTop": 0,
		"marginBottom": 0,
		"marginLeft": 0,
		"marginRight": 0,
		"startEffect": "easeInSine",
		"pullOutRadius": 0,
		"balloon":{
			"fixedPosition":true			
		},
		"balloonText": "[[referral]]" + ' ' + "[[contribPercent]]" + '% ($' + "[[contribDollars]]" + ')',
		"export": {
		"enabled": false
		}
	  
	});
	
	var chart = AmCharts.makeChart( "chartdivCustType", {
	
		"titles": [
				{
					"text": "Cust Type Leakage for <%=LCP_Display_Var%>",
					"size": 18,
					"bold": "false"
				}				
			],
		"creditsPosition":"bottom-left",
		"labelRadius": -50,
		"labelText": "[[custtype]]",
		"type": "pie",
		"theme": "light",
		"dataProvider": [ <%=amChartDataCustType%> ],
		"valueField": "contribPercent",
		"titleField": "custtype",
		"maxLabelWidth" : "100",
		"autoMargins": false,
		"marginTop": 0,
		"marginBottom": 0,
		"marginLeft": 0,
		"marginRight": 0,
		"startEffect": "easeInSine",		
		"pullOutRadius": 0,
		"balloonText": "[[custtype]]" + ' ' + "[[contribPercent]]" + '% ($' + "[[contribDollars]]" + ')',
		"balloon":{
			"fixedPosition":true
		},
		"export": {
		"enabled": false
		}
	  
	});
	
	var chart = AmCharts.makeChart( "chartdivSls1", {
	
		"titles": [
				{
					"text": "Primary Leakage for <%=LCP_Display_Var%>",
					"size": 18,
					"bold": "false"
				}				
			],
		"creditsPosition":"bottom-left",
		"labelRadius": -50,
		"labelText": "[[primary]]",
		"type": "pie",
		"theme": "light",
		"dataProvider": [ <%=amChartDataSls1%> ],
		"valueField": "contribPercent",
		"titleField": "primary",
		"maxLabelWidth" : "100",
		"autoMargins": false,
		"marginTop": 0,
		"marginBottom": 0,
		"marginLeft": 0,
		"marginRight": 0,
		"startEffect": "easeInSine",		
		"pullOutRadius": 0,
		"balloonText": "[[primary]]" + ' ' + "[[contribPercent]]" + '% ($' + "[[contribDollars]]" + ')',
		"balloon":{
			"fixedPosition":true
		},
		"export": {
		"enabled": false
		}
	  
	});
	
	
	var chart = AmCharts.makeChart( "chartdivSls2", {
	
		"titles": [
				{
					"text": "Secondary Leakage for <%=LCP_Display_Var%>",
					"size": 18,
					"bold": "false"
				}				
			],
		"creditsPosition":"bottom-left",
		"labelRadius": -50,
		"labelText": "[[secondary]]",
		"type": "pie",
		"theme": "light",
		"dataProvider": [ <%=amChartDataSls2%> ],
		"valueField": "contribPercent",
		"titleField": "secondary",
		"maxLabelWidth" : "100",
		"autoMargins": false,
		"marginTop": 0,
		"marginBottom": 0,
		"marginLeft": 0,
		"marginRight": 0,
		"startEffect": "easeInSine",		
		"pullOutRadius": 0,
		"balloonText": "[[secondary]]" + ' ' + "[[contribPercent]]" + '% ($' + "[[contribDollars]]" + ')',
		"balloon":{
			"fixedPosition":true
		},
		"export": {
		"enabled": false
		}
	  
	});
	

var chart = AmCharts.makeChart("chartdivSAL", {
	"titles": [
			{
				"text": "Period Sales",
				"size": 18,
				"bold": "false"
			}				
		],
	"type": "serial",
    "theme": "light",
    "marginRight":30,
    "legend": {
        "equalWidths": false,
        "periodValueText": "total: [[value.sum]]",
        "position": "top",
        "valueAlign": "left",
        "valueWidth": 100
    },
    "dataProvider": [<%= amChartDataSAL %>],
    "valueAxes": [{
        "gridAlpha": 0.07,
        "position": "left",
        "title": "Sales Dollars"
    }],
    "graphs": [{
        "balloonText": "<span style='font-size:14px; color:#000000;'><b>Sales:$[[value]]</b></span>",
        "fillAlphas": 0.6,
        "lineAlpha": 0.4,
        "title": "Sales",
        "valueField": "sales"
    }, {
        "balloonText": "<span style='font-size:14px; color:#000000;'><b>PP Sales:$[[value]]</b></span>",
        "fillAlphas": 0.6,
        "lineAlpha": 0.4,
        "hidden": true,
        "title": "Prior period sales",
        "valueField": "ppsales"
    }, {
        "balloonText": "<span style='font-size:14px; color:#000000;'><b>P3P Avg:$[[value]]</b></span>",
        "fillAlphas": 0.6,
        "lineAlpha": 0.4,
        "title": "Prior 3 periods average sales",
        "valueField": "p3pavgsales"
    }],
    "plotAreaBorderAlpha": 0,
    "marginTop": 10,
    "marginLeft": 0,
    "marginBottom": 0,

    "chartCursor": {
        "cursorAlpha": 0
    },
    "categoryField": "year",
    "categoryAxis": {
        "startOnAxis": true,
        "axisColor": "#DADADA",
        "gridAlpha": 0.07,
        "title": "Period",
        "guides": [{
            category: "2001",
            toCategory: "2003",
            lineColor: "#CC0000",
            lineAlpha: 1,
            fillAlpha: 0.2,
            fillColor: "#CC0000",
            dashLength: 2,
            inside: true,
            labelRotation: 90,
            label: "fines for speeding increased"
        }, {
            category: "2007",
            lineColor: "#CC0000",
            lineAlpha: 1,
            dashLength: 2,
            inside: true,
            labelRotation: 90,
            label: "motorcycle fee introduced"
        }]
    },
    "export": {
    	"enabled": false
     }
});



var chart = AmCharts.makeChart( "chartdivRefSALES", {

	"titles": [
			{
				"text": "Referral Code Sales for <%=LCP_Display_Var%>",
				"size": 18,
				"bold": "false"
			}				
		],
	"creditsPosition":"bottom-left",
	"labelRadius": -50,
	"labelText": "[[referral]]",
	"type": "pie",
	"theme": "light",
	"dataProvider": [ <%=amChartDataReferralSALES%> ],
	"valueField": "contribPercent",
	"titleField": "referral",
	"maxLabelWidth" : "100",
	"autoMargins": false,
	"marginTop": 0,
	"marginBottom": 0,
	"marginLeft": 0,
	"marginRight": 0,
	"startEffect": "easeInSine",
	"pullOutRadius": 0,
	"balloonText": "[[referral]]" + ' ' + "[[contribPercent]]" + '% ($' + "[[contribDollars]]" + ')',
	"balloon":{
		"fixedPosition":true
	},
	"export": {
	"enabled": false
	}
  
});
	

var chart = AmCharts.makeChart( "chartdivCustTypeSALES", {

	"titles": [
			{
				"text": "Cust Type Sales for <%=LCP_Display_Var%>",
				"size": 18,
				"bold": "false"
			}				
		],
	"creditsPosition":"bottom-left",
	"labelRadius": -50,
	"labelText": "[[custtype]]",
	"type": "pie",
	"theme": "light",
	"dataProvider": [ <%=amChartDataCustTypeSALES%> ],
	"valueField": "contribPercent",
	"titleField": "custtype",
	"maxLabelWidth" : "100",
	"autoMargins": false,
	"marginTop": 0,
	"marginBottom": 0,
	"marginLeft": 0,
	"marginRight": 0,
	"startEffect": "easeInSine",		
	"pullOutRadius": 0,
	"balloonText": "[[custtype]]" + ' ' + "[[contribPercent]]" + '% ($' + "[[contribDollars]]" + ')',
	"balloon":{
		"fixedPosition":true
	},
	"export": {
	"enabled": false
	}
  
});
	
	
	
	var chart = AmCharts.makeChart( "chartdivSls1SALES", {
	
		"titles": [
				{
					"text": "Primary Sales  for <%=LCP_Display_Var%>",
					"size": 18,
					"bold": "false"
				}				
			],
		"creditsPosition":"bottom-left",
		"labelRadius": -50,
		"labelText": "[[primary]]",
		"type": "pie",
		"theme": "light",
		"dataProvider": [ <%=amChartDataSls1SALES%> ],
		"valueField": "contribPercent",
		"titleField": "primary",
		"maxLabelWidth" : "100",
		"autoMargins": false,
		"marginTop": 0,
		"marginBottom": 0,
		"marginLeft": 0,
		"marginRight": 0,
		"startEffect": "easeInSine",		
		"pullOutRadius": 0,
		"balloonText": "[[primary]]" + ' ' + "[[contribPercent]]" + '% ($' + "[[contribDollars]]" + ')',
		"balloon":{
			"fixedPosition":true
		},
		"export": {
		"enabled": false
		}
	  
	});
	




var chart = AmCharts.makeChart( "chartdivSls2SALES", {

	"titles": [
			{
				"text": "Secondary Sales for <%=LCP_Display_Var%>",
				"size": 18,
				"bold": "false"
			}				
		],
	"creditsPosition":"bottom-left",
	"labelRadius": -50,
	"labelText": "[[secondary]]",
	"type": "pie",
	"theme": "light",
	"dataProvider": [ <%=amChartDataSls2SALES%> ],
	"valueField": "contribPercent",
	"titleField": "secondary",
	"maxLabelWidth" : "100",
	"autoMargins": false,
	"marginTop": 0,
	"marginBottom": 0,
	"marginLeft": 0,
	"marginRight": 0,
	"startEffect": "easeInSine",		
	"pullOutRadius": 0,
	"balloonText": "[[secondary]]" + ' ' + "[[contribPercent]]" + '% ($' + "[[contribDollars]]" + ')',
	"balloon":{
		"fixedPosition":true
	},
	"export": {
	"enabled": false
	}
  
});
	

	
	
	var chart = AmCharts.makeChart("chartdivInvoiceHistoryCounts", {
		"titles": [
				{
					"text": "Order Counts",
					"size": 18,
					"bold": "false"
				}				
			],
		"type": "serial",
	    "theme": "light",
	    "marginRight":30,   
		"legend": {
		    "markerType": "circle",
	        "equalWidths": false,
	        "periodValueText": "[[value.sum]]",
	        "position": "top",
	        "valueAlign": "left",
	        "valueWidth": 100,
		    "valueFunction": function(legendData, valueText) {
		      //values is available on mouseover
		      if (legendData.values) {
		        var id = legendData.graph.id;
		
		        if (id === "totalorders") {
		          return valueText === "0" ? " " : valueText;
		        }
		      }
		      //initial value when no mouse cursor is present or on mouseout
		      else if (legendData.id === "totalorders") { 
		        return " ";
		      }
		
		      return valueText;
		    }
		},    
	    "dataProvider": [<%= amChartDataInvHistCounts %>],
	    "valueAxes": [{
	        "gridAlpha": 0.07,
	        "position": "left",
	        "title": "Num Orders"
	    }],
	    "balloon": {
	   		"hideBalloonTime": 1000, // 1 second
	    	"disableMouseEvents": false, // allow click
	    	"fixedPosition": true,
		    "cornerRadius": 6,
		    "adjustBorderColor": false,
		    "horizontalPadding": 10,
		    "verticalPadding": 10	,
		    "maxWidth": 10000 
	  	},	    
		"chartCursor": {
		    "cursorAlpha": 0,
		    "oneBalloonOnly": true
		},	 	    
	    "graphs": [{
	    	"id": "totalorders",
	        "balloonText": "<span style='font-size:14px; color:#000000;'><b>Total # Orders: [[value]]</b></span>",
	        "fillAlphas": 0.6,
	        "lineAlpha": 0.4,
	        "hidden": false,
	        "title": "Total # Orders",
	        "valueField": "totalorders",
	        "lineColor": "#79c0df",
		    "bullet": "round",
		    "bulletBorderThickness": 1
	        
	    }, {
	    	"id": "websiteorders",
	        "balloonText": "<span style='font-size:14px; color:#000000;'><b># Website Orders: [[value]] ([[websiteorderspercent]]%)</b></span>",	        
	        "fillAlphas": 0.6,
	        "lineAlpha": 0.4,
	        "title": "# Website Orders",
	        "valueField": "websiteorders",
			"lineColor": "#83b871",
			"bullet": "square",
			"bulletBorderThickness": 1 	        	   	             
	    }, {	    
	    	"id": "telselorders",
	        "balloonText": "<span style='font-size:14px; color:#000000;'><b># Telsel Orders: [[value]] ([[telselorderspercent]]%)</b></span>",
	        "fillAlphas": 0.6,
	        "lineAlpha": 0.4,
	        "title": "# Telsel Orders",
	        "valueField": "telselorders",
		    "lineColor": "#bf6059",
		    "bullet": "triangleUp",
		    "bulletBorderThickness": 1 	        	   
	    }, {
	    	"id": "apiorders",
	        "balloonText": "<span style='font-size:14px; color:#000000;'><b># API Orders: [[value]] ([[apiorderspercent]]%)</b></span>",
	        "fillAlphas": 0.6,
	        "lineAlpha": 0.4,
	        "title": "# API Orders",
	        "valueField": "apiorders",
		    "lineColor": "#f7d85a",
		    "bullet": "bubble",
		    "bulletBorderThickness": 1 	        
	    }],
	    "plotAreaBorderAlpha": 0,
	    "marginTop": 10,
	    "marginLeft": 0,
	    "marginBottom": 0,
	
	    "chartCursor": {
	        "cursorAlpha": 0
	    },
	    "categoryField": "month",
	    "categoryAxis": {
	        "startOnAxis": true,
	        "axisColor": "#DADADA",
	        "gridAlpha": 0.07,
	        "title": "Month/Year"
	    },
	    "synchronizeGrid" : true,
	    "export": {
	    	"enabled": false
	     }
	});
	


	
	
	var chart = AmCharts.makeChart("chartdivInvoiceHistoryDollars", {
		"titles": [
				{
					"text": "Order $ Totals",
					"size": 18,
					"bold": "false"
				}				
			],
		"type": "serial",
	    "theme": "light",
	    "marginRight":30,   
		"legend": {
		    "markerType": "circle",
	        "equalWidths": false,
	        "periodValueText": "$[[value.sum]]",
	        "position": "top",
	        "valueAlign": "left",
	        "valueWidth": 150,
	        "labelWidth": 150,
		},    
	    "dataProvider": [<%= amChartDataInvHistDollars %>],
	    "valueAxes": [{
	        "gridAlpha": 0.07,
	        "position": "left",
	        "title": "$ Total"
	    }],
	    "balloon": {
	   		"hideBalloonTime": 1000, // 1 second
	    	"disableMouseEvents": false, // allow click
	    	"fixedPosition": true,
		    "cornerRadius": 6,
		    "adjustBorderColor": false,
		    "horizontalPadding": 10,
		    "verticalPadding": 10,
		    "maxWidth": 10000 
	  	},	    
		"chartCursor": {
		    "cursorAlpha": 0,
		    "oneBalloonOnly": true
		},	 	    
	    "graphs": [{
	    	"id": "totalordersdollars",
	        "balloonText": "<span style='font-size:14px; color:#000000;'><b>Total Orders: $[[value]]</b></span>",
	        "fillAlphas": 0.6,
	        "lineAlpha": 0.4,
	        "hidden": false,
	        "title": "Total Orders",
	        "valueField": "totalordersdollars",
	        "lineColor": "#79c0df",
		    "bullet": "round",
		    "bulletBorderThickness": 1
	        
	    }, {
	    	"id": "websiteordersdollars",
	        "balloonText": "<span style='font-size:14px; color:#000000;'><b>Website Orders: $[[value]]</b></span>",	        
	        "fillAlphas": 0.6,
	        "lineAlpha": 0.4,
	        "title": "Website Orders",
	        "valueField": "websiteordersdollars",
			"lineColor": "#83b871",
			"bullet": "square",
			"bulletBorderThickness": 1 	        	   	             
	    }, {	    
	    	"id": "telselordersdollars",
	        "balloonText": "<span style='font-size:14px; color:#000000;'><b>Telsel Orders: $[[value]]</b></span>",
	        "fillAlphas": 0.6,
	        "lineAlpha": 0.4,
	        "title": "Telsel Orders",
	        "valueField": "telselordersdollars",
		    "lineColor": "#bf6059",
		    "bullet": "triangleUp",
		    "bulletBorderThickness": 1 	        	   
	    }, {
	    	"id": "apiordersdollars",
	        "balloonText": "<span style='font-size:14px; color:#000000;'><b>API Orders: $[[value]]</b></span>",
	        "fillAlphas": 0.6,
	        "lineAlpha": 0.4,
	        "title": "API Orders",
	        "valueField": "apiordersdollars",
		    "lineColor": "#f7d85a",
		    "bullet": "bubble",
		    "bulletBorderThickness": 1 	        
	    }],
	    "plotAreaBorderAlpha": 0,
	    "marginTop": 10,
	    "marginLeft": 0,
	    "marginBottom": 0,
	
	    "chartCursor": {
	        "cursorAlpha": 0
	    },
	    "categoryField": "month",
	    "categoryAxis": {
	        "startOnAxis": true,
	        "axisColor": "#DADADA",
	        "gridAlpha": 0.07,
	        "title": "Month/Year"
	    },
	    "synchronizeGrid" : true,
	    "export": {
	    	"enabled": false
	     }
	});

</script>



<!-- am chart js !-->  
<p class="auto-style1">&nbsp;</p>
  
