<%@ Language="VBScript" %>
<!DOCTYPE html>
<html lang="en">
    <head>
        <meta charset="utf-8"/>
        <title>Classic ASP Image Resizer</title>
        <style type="text/css">
            img {border:5px solid #ffffff}
            body {font-family:Verdana;font-size:10pt;color:#ffffff;background-color:#333333}
			a {color:#dddddd}
        </style>        
    </head>
    <body>
		<h1>Classic ASP Image Resizer</h1>
        <p>Classic ASP script which uses a .NET Image resizer. JPG, PNG, GIF and BMP images are resized on the fly by the ResizeImage.aspx file using the following QueryString values:</p>
		<ul>
			<li><b>ImgWd</b> - IMaGe WiDth - <i>The required width of the image. Optional: Exclude to scale with height.</i></li>
			<li><b>ImgHt</b> - IMaGe HeighT - <i>The required height of the image. Optional: Exclude to scale with width.</i></li>
			<li><b>CrpYN</b> - CRoP Yes / No - <i>Should the image be cropped or aspect ratio maintained (even if padding is required)? Optional: If excluded, default value is "N".</i></li>
			<li><b>PadCl</b> - PAD CoLour - <i>Pass a HTML hex value (minus the hash #) to be used as the background colour when padding images. Optional: If excluded and padding is required, default value is "ffffff" i.e. white.</i></li>
			<li><b>IptFl</b> - InPut FiLe - <i>The relative path and filename of the image to be resized e.g. image.jpg, folder/image.jpg, /folder/image.jpg.</i></li>
			<li><b>OptFl</b> - OutPuT FiLe - <i>The relative path and filename to save the resulting image to. Optional: Exclude to just display on screen.</i></li>
			<li><b>OptSc</b> - OutPuT to SCreen - <i>Should the resulting image be returned to the screen. Optional: If excluded, default value is "Y".</i></li>
			<li><b>SplEf</b> - SPeciaL EfFect - <i>Add a special effect to your image using SplEf=X. Optional: Options as follows: "1" Black and White "2" GreyScale "3" Sepia</i></li>
		</ul>
		<p style="color:red"><b>IMPORTANT</b> - If you wish to use the OptFl option, you will need to uncomment the relevant lines in the aspx file. See the security advice at the top of the file for details.</p>
		<h2>Relative Resize (Width / Height)</h2>
        <div style="margin:20px;float:left"><img src="ResizeImage.aspx?ImgWd=800&amp;IptFl=TicketID-1383090.png"/><br/>ImgWd=800&amp;IptFl=TicketID-1383090.png</div>
        <div style="margin:20px;float:left"><img src="ResizeImage.aspx?ImgHt=225&amp;IptFl=TicketID-1383090.png"/><br/>ImgHt=225&amp;IptFl=TicketID-1383090.png</div>
        <div style="clear:both"></div>

		<h2>Fixed Resize (Cropped / Padded / Coloured Padding)</h2>
        <div style="margin:20px;float:left"><img src="ResizeImage.aspx?ImgWd=400&amp;ImgHt=400&amp;IptFl=TicketID-1383090.png&amp;CrpYN=Y"/><br/>ImgWd=400&amp;ImgHt=400&amp;IptFl=TicketID-1383090.png&amp;CrpYN=Y</div>
        <div style="margin:20px;float:left"><img src="ResizeImage.aspx?ImgWd=400&amp;ImgHt=400&amp;IptFl=TicketID-1383090.png"/><br/>ImgWd=400&amp;ImgHt=400&amp;IptFl=TicketID-1383090.png</div>
        <div style="margin:20px;float:left"><img src="ResizeImage.aspx?ImgWd=400&amp;ImgHt=400&amp;IptFl=TicketID-1383090.png&amp;PadCl=ff0000"/><br/>ImgWd=400&amp;ImgHt=400&amp;IptFl=TicketID-1383090.png&amp;PadCl=ff0000</div>
        <div style="clear:both"></div>

		<h2>Special Effects (Black &amp; White / Greyscale / Sepia)</h2>
        <div style="margin:20px;float:left"><img src="ResizeImage.aspx?ImgWd=300&amp;IptFl=TicketID-1383090.png&amp;SplEf=1"/><br/>ImgWd=300&amp;IptFl=TicketID-1383090.png&amp;SplEf=1</div>
        <div style="margin:20px;float:left"><img src="ResizeImage.aspx?ImgWd=300&amp;IptFl=TicketID-1383090.png&amp;SplEf=2"/><br/>ImgWd=300&amp;IptFl=TicketID-1383090.png&amp;SplEf=2</div>
        <div style="margin:20px;float:left"><img src="ResizeImage.aspx?ImgWd=300&amp;IptFl=TicketID-1383090.png&amp;SplEf=3"/><br/>ImgWd=300&amp;IptFl=TicketID-1383090.png&amp;SplEf=3</div>
        <div style="clear:both"></div>

		<h2>Output (Screen Only / Screen &amp; File / File Only)</h2>
        <div style="margin:20px;float:left"><img src="ResizeImage.aspx?ImgWd=400&amp;IptFl=TicketID-1383090.png"/><br/>ImgWd=400&amp;IptFl=TicketID-1383090.png</div>
        <div style="margin:20px;float:left"><img src="ResizeImage.aspx?ImgWd=300&amp;IptFl=TicketID-1383090.png&amp;OptFl=Thumb.jpg"/><br/>ImgWd=300&amp;IptFl=TicketID-1383090.png&amp;OptFl=Thumb.jpg</div>
        <div style="margin:20px;float:left"><img src="ResizeImage.aspx?ImgWd=1024&amp;IptFl=TicketID-1383090.png&amp;OptFl=Gallery.jpg&amp;OptSc=N" style="display:none"/><br/>ImgWd=1024&amp;IptFl=TicketID-1383090.png&amp;OptFl=Gallery.jpg&amp;OptSc=N</div>
        <div style="clear:both"></div>

		<h2>Credits</h2>
		<p>This solution was originally developed by Pieter Cooreman, <a href="http://www.quickersite.com">QuickerSite (Classic ASP Website CMS)</a>. It was adapted and republished by David Barton, <a href="http://www.easierthan.co.uk">EasierThan Website Design</a> via <a href="http://easierthan.blogspot.co.uk/2013/02/code-tip-3-classic-asp-image-resizer.html">EasierThan (Official Blog)</a>.</p>
		<p>The script is offered freely and no liability is accepted by either party for any issues arising from the downloading, installing or use of it.</p>
		<p>If you find the script useful, backlinks and / or a beer (see original <a href="http://www.imageresizing.quickersite.com">ASP Image Resizer</a> page) would be much appreciated, but are not necessary.</p>
    </body>
</html>
