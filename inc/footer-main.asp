  </div>
  <!-- eof content area !-->
  
</div>
<!-- dashboard ends here !-->

 
	<script src="https://cdnjs.cloudflare.com/ajax/libs/highlight.js/8.3/highlight.min.js"></script>		

    <!-- IE10 viewport hack for Surface/desktop Windows 8 bug -->
    <script src="<%= BaseURL %>js/ie10-viewport-bug-workaround.js"></script>
    
    <!-- charts !-->
	<script src="<%= BaseURL %>js/charts/highcharts.js"></script>
	<script src="<%= BaseURL %>js/charts/exporting.js"></script>
	<!-- eof charts !-->

<!-- tooltip JS !-->
<script type="text/javascript">
$(function () {
  $('[data-toggle="tooltip"]').tooltip()
})
 </script>
<!-- eof tooltip JS !-->

<!-- chart js !-->
<script type="text/javascript">
	$(function () {
	
	try{
        $('#container').highcharts({
            chart: {
                type: 'column'
            },
            title: {
                text: 'Service Call Activity'
            },
            subtitle: {
                text: '<%=DateRangeTitleForGraph%>'
            },
            xAxis: {
                categories: [
                    'Sun',
                    'Mon',
                    'Tue',
                    'Wed',
                    'Thu',
                    'Fri',
                    'Sat'
                    
                ]
            },
            yAxis: {
                min: 0,
                title: {
                    text: 'Number of tickets'
                }
            },
            tooltip: {
                headerFormat: '<span style="font-size:10px">{point.key}</span><table>',
                pointFormat: '<tr><td style="color:{series.color};padding:0">{series.name}: </td>' +
                    '<td style="padding:0"><b>{point.y}</b></td></tr>',
                footerFormat: '</table>',
                shared: true,
                useHTML: true
            },
            plotOptions: {
                column: {
                    pointPadding: 0.2,
                    borderWidth: 0,
                    maxPointWidth: 30
                }
            },
            series: [{
				 name: 'Open',
                data: [<%=SundayOpen%>, <%=MondayOpen%>, <%=TuesdayOpen%>, <%=WednesdayOpen%>, <%=ThursdayOpen%>, <%=FridayOpen%>, <%=SaturdayOpen%>]

            }, {
                name: 'Close',
                data: [<%=SundayClose%>, <%=MondayClose%>, <%=TuesdayClose%>, <%=WednesdayClose%>, <%=ThursdayClose%>, <%=FridayClose%>, <%=SaturdayClose%>]

            }, {
                name: 'Cancel',
                data: [<%=SundayCancel%>,  <%=MondayCancel%>, <%=TuesdayCancel%>, <%=WednesdayCancel%>, <%=ThursdayCancel%>, <%=FridayCancel%>, <%=SaturdayCancel%>]
			
			}, {
                name: 'Awaiting Dispatch',
                data: [<%=SundayDisp%>, <%=MondayDisp%>, <%=TuesdayDisp%>, <%=WednesdayDisp%>, <%=ThursdayDisp%>, <%=FridayDisp%>, <%=SaturdayDisp%>]

            }]
        });
}catch(ex){}
    });
    
    
    $(function () {
	try{
    $('#container2').highcharts({
        chart: {
            type: 'column'
        },
        title: {
            text: 'User Activity'
        },
        subtitle: {
	            text: '<%=ActivityChartSubTitle%>' 
        },
        xAxis: {
            type: 'category',
            labels: {
                rotation: -45,
                style: {
                    fontSize: '10px',
                    fontFamily: 'Verdana, sans-serif'
                }
            }
        },
        yAxis: {
            min: 0,
            title: {
                text: 'Audit Trail Events'
            }
        },
        legend: {
            enabled: false
        },
        tooltip: {
            pointFormat: 'Audit Trail Events: <b>{point.y}</b>'
        },
        
        series: [{
            name: 'User',
				data: [
				
						<%=aspDataVar%>  
				
				       ],
				
            dataLabels: {
                enabled: false,
                rotation: -90,
                color: '#FFFFFF',
                align: 'right',
                format: '{point.y:.1f}', // one decimal
                y: 10, // 10 pixels down from the top
                style: {
                    fontSize: '10px',
                    fontFamily: 'Verdana, sans-serif'
                }
            }
        }]
    });
}catch(ex){}
});	


	$(function () {
	try{
			
        $('#containerAPI').highcharts({
            chart: {
                type: 'column'            
            },
            title: {
                text: 'Orders API Activity'
            },
            subtitle: {
                text: '<%=DateRangeTitleForAPIGraph %>'
            },
            xAxis: {
                categories: ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'],
		        title: {
		            text: null
		        }                
            },
            yAxis: {
                min: 0,
                title: {
                    text: 'Number of Records'
                },
		        labels: {
		            overflow: 'justify'
		        }                
            },
            tooltip: {
                headerFormat: '<span style="font-size:10px">{point.key}</span><table>',
                pointFormat: '<tr><td style="color:{series.color};padding:0">{series.name}: </td>' +
                    '<td style="padding:0"><b>{point.y}</b></td></tr>',
                footerFormat: '</table>',
                shared: true,
                useHTML: true
            },
            plotOptions: {
				bar: {
		            dataLabels: {
		                enabled: true
		            }
				}	
             
            },
            series: [{
				 name: 'Orders',
                data: [<%=SundayNumOrders%>, <%=MondayNumOrders%>, <%=TuesdayNumOrders%>, <%=WednesdayNumOrders%>, <%=ThursdayNumOrders%>, <%=FridayNumOrders%>, <%=SaturdayNumOrders%>]

            }, {
                name: 'Invoices',
                data: [<%=SundayNumInvoices%>, <%=MondayNumInvoices%>, <%=TuesdayNumInvoices%>, <%=WednesdayNumInvoices%>, <%=ThursdayNumInvoices%>, <%=FridayNumInvoices%>, <%=SaturdayNumInvoices%>]

            }, {
                name: 'RAs',
                data: [<%=SundayNumRAs%>,  <%=MondayNumRAs%>, <%=TuesdayNumRAs%>, <%=WednesdayNumRAs%>, <%=ThursdayNumRAs%>, <%=FridayNumRAs%>, <%=SaturdayNumRAs%>]
			
			}, {
			
                name: 'Credit Memos',
                data: [<%=SundayNumCMs%>,  <%=MondayNumCMs%>, <%=TuesdayNumCMs%>, <%=WednesdayNumCMs%>, <%=ThursdayNumCMs%>, <%=FridayNumCMs%>, <%=SaturdayNumCMs%>]
			
			}, {
			
                name: 'Summary Inv',
                data: [<%=SundayNumSummaryInvoices%>, <%=MondayNumSummaryInvoices%>, <%=TuesdayNumSummaryInvoices%>, <%=WednesdayNumSummaryInvoices%>, <%=ThursdayNumSummaryInvoices%>, <%=FridayNumSummaryInvoices%>, <%=SaturdayNumSummaryInvoices%>]

            }]
        });
}catch(ex){}
    });

</script>
<!-- eof chart js !-->

<!--#include file="sessionTimeout.asp"-->
 
  </body>
</html>