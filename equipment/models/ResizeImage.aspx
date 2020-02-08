<%@ OutputCache Duration="600" VaryByParam="*" %>
<%@ Page Debug="false" %>
<%@ Import Namespace="System.Drawing" %>
<%@ Import Namespace="System.Drawing.Imaging" %>
<%@ Import Namespace="System.IO" %>
<script language="VB" runat="server">
	
	' PASSED VARIABLES
	' ================

	' ImgWd - IMaGe WiDth - The required width of the image. Optional: Exclude to scale with height.
	' ImgHt - IMaGe HeighT - The required height of the image. Optional: Exclude to scale with width.
	' CrpYN - CRoP Yes / No - Should the image be cropped or aspect ratio maintained (even if padding is required)? Optional: If excluded, default value is "N".
	' PadCl - PAD CoLour - Pass a HTML hex value (minus the hash #) to be used as the background colour when padding images. Optional: If excluded and padding is required, default value is "ffffff" i.e. white.
	' IptFl - InPut FiLe - The relative path and filename of the image to be resized e.g. image.jpg, folder/image.jpg, /folder/image.jpg.
	' OptFl - OutPuT FiLe - The relative path and filename to save the resulting image to. Optional: Exclude to just display on screen.
	' OptSc - OutPuT to SCreen - Should the resulting image be returned to the screen. Optional: If excluded, default value is "Y".
	' SplEf - SPeciaL EfFect - Add a special effect to your image using SplEf=X. Optional: Options as follows: "1" Black and White "2" GreyScale "3" Sepia

	' SECURITY ADVICE
	' ===============

	' If you wish to use the OptFl option, you will need to uncomment the relevant lines below, but be cautious in using it on a public website. Having the OptFl option available could allow a determined hacker
	' to place / replace images on your website.
	' We have included a simple check on the HTTP_REFERER value to verify that the script was called from the same website, but note that HTTP_REFERER itself can be faked using readily available software. Plus
	' in some instances HTTP_REFERER may be blocked by browsers, firewalls or anti-virus.
	' For more security and reliability, you could try passing session state between your ASP page and this one or writing a line to a database table (or file) within ASP and reading it out from this one.
	' If the filename of the image to be saved is always the same (e.g. Temp_Upload.jpg), it would be more advisable to hard code it, rather than using OptFl.
	' As this script is freely available, it would also be sensible to change the name of the aspx file and the titles of the query string values i.e. Find and replace all ImgWd with IW or similar.

	' Run when the page loads
	Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs)
		
		On Error Resume Next
			 
		' Attempt to check that the script is being called from its own server
' --- UNCOMMENT IF YOU WISH TO USE OptFl OPTION ----------------------------------------------------------
'		If InStr(Request.ServerVariables("HTTP_REFERER"), Request.ServerVariables("HTTP_HOST")) > 0 Then
' --------------------------------------------------------------------------------------------------------

			' Retrieve relative path to image
			Dim sInputURL As String = HttpUtility.UrlDecode(Request.QueryString("IptFl"))
			
			Response.ContentType = "image/jpeg"
			
			' Prepare image
			Dim imgFullSize As System.Drawing.Image
			imgFullSize = System.Drawing.Image.FromFile(Server.MapPath(sInputURL))		
			
			' Discard if image file not found
			If imgFullSize Is Nothing Then Response.End()			
			
			' Not sure what this does? Necessary though!
			Dim cbDummyCallBack As System.Drawing.Image.GetThumbnailImageAbort
			cbDummyCallBack = New System.Drawing.Image.GetThumbnailImageAbort(AddressOf ThumbnailCallback)
			
			Dim imgThumbnail As System.Drawing.Image
			Dim iOrigWidth As Integer = imgFullSize.Width
			Dim iOrigHeight As Integer = imgFullSize.Height
			Dim iResizeWidth As Integer
			Dim iResizeHeight As Integer
			Dim iThumbWidth As Integer = IIf(IsNumeric(Request.QueryString("ImgWd")) And Request.QueryString("ImgWd") <> "", Request.QueryString("ImgWd"), 0)
			Dim iThumbHeight As Integer = IIf(IsNumeric(Request.QueryString("ImgHt")) And Request.QueryString("ImgHt") <> "", Request.QueryString("ImgHt"), 0)
			Dim biCropYN As Boolean = IIf(Request.QueryString("CrpYN") = "Y", True, False)
			
			' Calculate new width / height, if any  
			If (iOrigWidth <> iThumbWidth Or iOrigHeight <> iThumbHeight) And (iThumbWidth + iThumbHeight <> 0) Then
			
				' For better quality resize?!?
				imgFullSize.RotateFlip(System.Drawing.RotateFlipType.Rotate90FlipX)
				imgFullSize.RotateFlip(System.Drawing.RotateFlipType.Rotate90FlipX)
			
				' If not specified, make the thumb width or height relative
				If iThumbWidth = 0 Then
					iThumbWidth = (iOrigWidth / iOrigHeight) * iThumbHeight
				End If
				If iThumbHeight = 0 Then
					iThumbHeight = (iOrigHeight / iOrigWidth) * iThumbWidth
				End If

				If biCropYN = False Then

					' Maintain aspect ratio. Padding may be required.
					If (iOrigWidth / iOrigHeight) = (iThumbWidth / iThumbHeight) Then

						' Exact aspect ratio match. No padding required.
						iResizeWidth = iThumbWidth
						iResizeHeight = iThumbHeight
						imgThumbnail = imgFullSize.GetThumbnailImage(iResizeWidth, iResizeHeight, cbDummyCallBack, IntPtr.Zero)
					
					Else

						' Different aspect ratio. Padding required.
						If (iOrigWidth / iOrigHeight) > (iThumbWidth / iThumbHeight) Then
							
							' Landscape. Resize maintaining aspect ratio, then pad.
							iResizeWidth = iThumbWidth
							iResizeHeight =	(iOrigHeight / iOrigWidth) * iThumbWidth
							imgThumbnail = imgFullSize.GetThumbnailImage(iResizeWidth, iResizeHeight, cbDummyCallBack, IntPtr.Zero)
							imgThumbnail = fCropPadImage(imgThumbnail, 0, -((iThumbHeight - iResizeHeight) / 2), iThumbWidth, iThumbHeight)

						Else

							' Portrait. Resize maintaining aspect ratio, then pad.
							iResizeHeight = iThumbHeight
							iResizeWidth = (iOrigWidth / iOrigHeight) * iThumbHeight
							imgThumbnail = imgFullSize.GetThumbnailImage(iResizeWidth, iResizeHeight, cbDummyCallBack, IntPtr.Zero)
							imgThumbnail = fCropPadImage(imgThumbnail, -((iThumbWidth - iResizeWidth) / 2), 0, iThumbWidth, iThumbHeight)

						End If
					
					End If
					
				Else
		
					' Cropping required
					If (iOrigWidth / iOrigHeight) > (iThumbWidth / iThumbHeight) Then
						
						' Landscape. Resize maintaining aspect ratio, then crop.
						iResizeHeight = iThumbHeight
						iResizeWidth = (iOrigWidth / iOrigHeight) * iThumbHeight
						imgThumbnail = imgFullSize.GetThumbnailImage(iResizeWidth, iResizeHeight, cbDummyCallBack, IntPtr.Zero)
						imgThumbnail = fCropPadImage(imgThumbnail, -((iThumbWidth - iResizeWidth) / 2), 0, iThumbWidth, iThumbHeight)

					Else

						' Portrait. Resize maintaining aspect ratio, then crop.
						iResizeWidth = iThumbWidth
						iResizeHeight =	(iOrigHeight / iOrigWidth) * iThumbWidth
						imgThumbnail = imgFullSize.GetThumbnailImage(iResizeWidth, iResizeHeight, cbDummyCallBack, IntPtr.Zero)
						imgThumbnail = fCropPadImage(imgThumbnail, 0, -((iThumbHeight - iResizeHeight) / 2), iThumbWidth, iThumbHeight)

					End If

				End If
			Else
				
				' No resize required
				imgThumbnail = imgFullSize

			End If
			
			fOutputImage(imgThumbnail)				   
			
			' Clean up / Dispose...
			imgThumbnail.Dispose()
			
			' Clean up / Dispose...
			imgFullSize.Dispose()

' --- UNCOMMENT IF YOU WISH TO USE OptFl OPTION ----------------------------------------------------------
'		Else
'
			' If script being called from another server, it may be a hack. Redirect the user to the root of your domain (or change as appropriate)
'			Response.Redirect("/")
'
'		End If
' --------------------------------------------------------------------------------------------------------
		
		On Error GoTo 0
	End Sub

	' Write image to screen and / or file
	Sub fOutputImage(ByRef imgOutput As System.Drawing.Image)
		
		On Error Resume Next
		
		' Add special effect if requested
		Dim sSpecialEffect as String = HttpUtility.UrlDecode(Request.QueryString("SplEf"))
		Select Case sSpecialEffect
			Case "1"
				fEffectPureBW (imgOutput)
			Case "2"
				GrayScale (imgOutput)
			Case "3"
				fEffectSepia (imgOutput)
		End Select
			
		' Send the resulting thumbnail back to the screen. When combined with a style of "display:none" in the calling HTML, this option can be used to hide the image, but save the result to a file (as below).
		Dim bOutputToScreenYN As Boolean = IIf(Request.QueryString("OptSc") = "N", False, True)
		If bOutputToScreenYN = True Then
			imgOutput.Save(Response.OutputStream, ImageFormat.Jpeg)
		End If

		' Save the resulting thumbnail to an output file if one is specified
' --- UNCOMMENT IF YOU WISH TO USE OptFl OPTION ----------------------------------------------------------
'		Dim sOutputURL As String = IIf(Request.QueryString("OptFl") <> "", HttpUtility.UrlDecode(Request.QueryString("OptFl")), "")
'		If sOutputURL <> "" And sOutputURL.Length > 5 Then
'			imgOutput.Save(Server.MapPath(sOutputURL), ImageFormat.Jpeg)
'		End If
' --------------------------------------------------------------------------------------------------------

		On Error GoTo 0
	End Sub
	
	' Crop or pad a passed image
	Private Function fCropPadImage(ByVal bmpOriginal As Bitmap, ByVal iCropX As Integer, ByVal iCropY As Integer, ByVal iCropWidth As Integer, ByVal iCropHeight As Integer) As Bitmap

		' Create the new bitmap and associated graphics object
		Dim bmpCropped As New Bitmap(iCropWidth, iCropHeight)
		Dim g As Graphics = Graphics.FromImage(bmpCropped)

		' Paint the canvas white or as per the passed HTML hex colour
		Dim sPadColour As String = IIf(Request.QueryString("PadCl") <> "", "#" & HttpUtility.UrlDecode(Request.QueryString("PadCl")), "#ffffff")
		g.Clear(ColorTranslator.FromHtml(sPadColour))

		' Draw the specified section of the source bitmap to the new one
		g.DrawImage(bmpOriginal, New Rectangle(0, 0, iCropWidth, iCropHeight), iCropX, iCropY, iCropWidth, iCropHeight, GraphicsUnit.Pixel)
		' Clean up
		g.Dispose()

		' Return the finished bitmap
		Return bmpCropped

	End Function
	
	' Set up callback
	Function ThumbnailCallback() As Boolean
		
		Return False
	
	End Function
	
	' Convert an image to pure black & white
	Public Function fEffectPureBW(ByVal bmpImage As System.Drawing.Bitmap, Optional ByVal mMode As BWMode = BWMode.By_Lightness, Optional ByVal dTolerance As Single = 0) As System.Drawing.Bitmap
		Dim x As Integer
		Dim y As Integer
		If dTolerance > 1 Or dTolerance < -1 Then
			Throw New ArgumentOutOfRangeException
			Exit Function
		End If
		For x = 0 To bmpImage.Width - 1 Step 1
			For y = 0 To bmpImage.Height - 1 Step 1
				Dim cColour As Color = bmpImage.GetPixel(x, y)
				If mMode = BWMode.By_RGB_Value Then
					If (CInt(cColour.R) + CInt(cColour.G) + CInt(cColour.B)) > 383 - (dTolerance * 383) Then
						bmpImage.SetPixel(x, y, Color.White)
					Else
						bmpImage.SetPixel(x, y, Color.Black)
					End If
				Else
					If cColour.GetBrightness > 0.5 - (dTolerance / 2) Then
						bmpImage.SetPixel(x, y, Color.White)
					Else
						bmpImage.SetPixel(x, y, Color.Black)
					End If
				End If
			Next
		Next
		Return bmpImage
	End Function
	
	' Set up BWMode
	Enum BWMode
		By_Lightness
		By_RGB_Value
	End Enum
	
	' Convert an image to greyscale
	Public Function GrayScale (ByVal bmpImage As System.Drawing.Bitmap)
		
		Dim X As Integer
		Dim Y As Integer
		Dim iColour As Integer

		For X = 0 To bmpImage.Width - 1
			For Y = 0 To bmpImage.Height - 1
				iColour = (CInt(bmpImage.GetPixel(X, Y).R) + bmpImage.GetPixel(X, Y).G + bmpImage.GetPixel(X, Y).B) \ 3
				bmpImage.SetPixel(X, Y, Color.FromArgb(iColour, iColour, iColour))
			Next Y
		Next X
		
		Return bmpImage
		
	End Function
	
	' Convert an image to sepia
	Public Function fEffectSepia (ByVal bmpImage As System.Drawing.Bitmap)	
   
		For i As Integer = 0 To bmpImage.Width - 1
			For j As Integer = 0 To bmpImage.Height - 1
				Dim iRed As Integer = bmpImage.GetPixel(i, j).R
				Dim iGreen As Integer = bmpImage.GetPixel(i, j).G
				Dim iBlue As Integer = bmpImage.GetPixel(i, j).B

				Dim iSepiaRed As Integer = Math.Min(Convert.ToInt32(iRed * 0.393 + iGreen * 0.769 + iBlue * 0.189), 255)
				Dim iSepiaGreen As Integer = Math.Min(Convert.ToInt32(iRed * 0.349 + iGreen * 0.686 + iBlue * 0.168), 255)
				Dim iSepiaBlue As Integer = Math.Min(Convert.ToInt32(iRed * 0.272 + iGreen * 0.534 + iBlue * 0.131), 255)

				bmpImage.SetPixel(i, j, Color.FromArgb(iSepiaRed, iSepiaGreen, iSepiaBlue))
			Next
		Next
	
		Return bmpImage

	End Function

</script>