Sub Main()

Dim IPLive As FNCImageProcess
	Set IPLive = IPEditor.GetIPFromName("Live")
	'save the current image in TIFF:
	IPLive.SaveImage("C:\WWTEMP\Test.tif")
	'save the current image as a bitmap:
	IPLive.SaveImage("C:\WWTEMP\Test.bmp")
	'save the current image in JPEG format:
	IPLive.SaveImage("C:\WWTEMP\Test.jpg", 50)
End Sub
