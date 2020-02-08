
<!-- Resources -->
<script src="https://www.amcharts.com/lib/3/amcharts.js"></script>
<script src="https://www.amcharts.com/lib/3/pie.js"></script>
<script src="https://www.amcharts.com/lib/3/serial.js"></script>
<script src="https://www.amcharts.com/lib/3/plugins/export/export.min.js"></script>
<link rel="stylesheet" href="https://www.amcharts.com/lib/3/plugins/export/export.css" type="text/css" media="all" />
<script src="https://www.amcharts.com/lib/3/themes/light.js"></script>


<!-- Chart code -->
<script>
	
	var chart = AmCharts.makeChart("chartdivARCustomerCounts", {
		"titles": [
				{
					"text": "Customer Counts",
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
	        "periodValueText": "total: [[value.sum]]",
	        "position": "top",
	        "valueAlign": "left",
	        "valueWidth": 100,
		    "valueFunction": function(legendData, valueText) {
		      //values is available on mouseover
		      if (legendData.values) {
		        var id = legendData.graph.id;
		
		        if (id === "totalaccounts") {
		          return valueText === "0" ? " " : valueText;
		        }
		      }
		      //initial value when no mouse cursor is present or on mouseout
		      else if (legendData.id === "totalaccounts") { 
		        return " ";
		      }
		
		      return valueText;
		    }
		},    
	    "dataProvider": [<%= amChartDataARCustCounts %>],
	    "valueAxes": [{
	        "gridAlpha": 0.07,
	        "position": "left",
	        "title": "Num Customer Accounts"
	    }],
	    "balloon": {
	   		"hideBalloonTime": 1000, // 1 second
	    	"disableMouseEvents": false, // allow click
		    "fixedPosition": true,
			"cornerRadius": 6,
		    "adjustBorderColor": false,
		    "horizontalPadding": 10,
		    "verticalPadding": 10,
		    "maxWidth": 10000 ,
		    "offsetX" : 0,
		    "offsetY" : 10,
	  	},	    
		"chartCursor": {
		    "cursorAlpha": 0,
		    "oneBalloonOnly": true
		},	  	
	    "graphs": [{
	    	"id": "totalaccounts",
			"balloonText": "",
	        "fillAlphas": 0.6,
	        "lineAlpha": 0.4,
	        "hidden": true,
	        "title": "Total # Accounts",
	        "valueField": "totalaccounts",
	        "lineColor": "#FF6600",
		    "bullet": "round",
		    "bulletBorderThickness": 1
	    }, {
	    	"id": "activeaccounts",
	    	"balloonText": "<p style='text-align: left;'>" +
	    					"<i class='fa fa-lg fa-users' aria-hidden='true'></i>" +
	    					"<span style='font-size:14px; color:#000000;'>&nbsp;Total Accounts <strong>[[totalaccounts]]</strong></span></p><p style='text-align: left;'>" +
	    					"<a href='dashboard/graph_ARCustomerCounts_ActiveDrilldown.asp?m=[[monthsingle]]&y=[[year]]'><i class='fa fa-lg fa-user-plus' aria-hidden='true'></i></a>" +
	    					"<span style='font-size:14px; color:#000000;'>&nbsp;<a href='dashboard/graph_ARCustomerCounts_ActiveDrilldown.asp?m=[[monthsingle]]&y=[[year]]'>New Accounts <strong>[[value]]</strong></span></a></p><p style='text-align: left;'>" +
	    					"<a href='dashboard/graph_ARCustomerCounts_InactiveDrilldown.asp?m=[[monthsingle]]&y=[[year]]'><i class='fa fa-lg fa-user-times' aria-hidden='true'></i></a>" +
	    					"<span style='font-size:14px; color:#000000;'>&nbsp;<a href='dashboard/graph_ARCustomerCounts_InactiveDrilldown.asp?m=[[monthsingle]]&y=[[year]]'>Inactive Accounts <strong>[[inactiveaccounts]]</strong></span></a></p>",	        
	        //"balloonText": "<span style='font-size:14px; color:#000000;'><b># New Active Accounts: [[value]]</b><a href='https://google.com/'>Google</a></span>",
	        "fillAlphas": 0.6,
	        "lineAlpha": 0.4,
	        "title": "# New Active Accounts",
	        "valueField": "activeaccounts",
			"lineColor": "#FCD202",
			"bullet": "square",
			"bulletBorderThickness": 1             
	    }, {
	    	"id": "inactiveaccounts",
	    	"balloonText": "",
	        //"balloonText": "<span style='font-size:14px; color:#000000;'><b># Accounts Moved to Inactive: [[value]]</b><br><a href='https://google.com/'>Google</a></span>",
	        "fillAlphas": 0.6,
	        "lineAlpha": 0.4,
	        "title": "# Accounts Moved to Inactive",
	        "valueField": "inactiveaccounts",
		    "lineColor": "#B0DE09",
		    "bullet": "triangleUp",
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
