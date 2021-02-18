
Sub Main
	SetXRay False, "Microfocus", 120, False, 20, 1, fncBVSmall, True, True
	'Set Manipulator to position
	SetManipulator "XM", 443394, "YM", 14912, "ZT", -200000, "ZD", 300000, "TiltD", 0, fncPUMicrometerMillidegree, False
	If(FNCIPCleanUp) Then
		CloseIP(FNCLastIP)
	End If
	SetIP "Easy1", FNCImageProcessPath & "Easy1.ipr"
	If(FNCIPCleanUp) Then
		FNCLastIP="Easy1"
	End If
	XRay.Wait(10, fncWFMicroAmp)
	IPEditor.GetIPFromName(IPEditor.ActiveIP).RunOnce(FNCResultPath)
	XRay.Wait(0, fncWFXRayOn)
	Wait 2

	IPEditor.GetIPFromName(IPEditor.ActiveIP).SaveImage(FNCResultPath & "123.jpg")

End Sub


'============ Internal Subroutines, do not touch anything below this line ===============
Sub SetXRay(ByVal OpenDoor As Boolean, ByVal OperatingMode As String, ByVal KiloVolt As Double, ByVal IsoWatt As Boolean, ByVal MicroAmpere As Double, ByVal Watt As Double, ByVal ImgIntZoom As FNCBVZoom, ByVal IntensityControl As Boolean, ByVal XRayOn As Boolean)
	If OpenDoor = False Then
		Manipulator.CloseDoor(True)
	End If
	XRay.Parameter(VID_ImgIntZoom) = ImgIntZoom
	XRay.Parameter(VID_IntensityControl) = IntensityControl
	XRay.Parameter(VID_IsoWatt) = IsoWatt
	XRay.Parameter(VID_KiloVolt) = KiloVolt
	If IsoWatt = True Then
		XRay.Parameter(VID_Watt) = Watt
	End If
	If IsoWatt = False Then
		XRay.Parameter(VID_MicroAmpere) = MicroAmpere
	End If
	XRay.OperatingMode = OperatingMode
	XRAY.Parameter(VID_XRayOn) = XRayOn
	If OpenDoor = True Then
		Manipulator.OpenDoor(True)
	End If
End Sub



'============ Internal Subroutines, do not touch anything below this line ===============
Sub SetManipulator(ByVal AxisName1 As String, ByVal AxisVal1#, ByVal AxisName2 As String, ByVal AxisVal2#, ByVal AxisName3 As String, ByVal AxisVal3#, ByVal AxisName4 As String, ByVal AxisVal4#, ByVal AxisName5 As String, ByVal AxisVal5#, ByVal Unit As FNCPosUnit, ByVal MoveRelative As Boolean)
	Dim Axis As FNCAxis

	Manipulator.Unit = Unit
	Set Axis = Manipulator.GetAxisFromName(AxisName1)
	Axis.MoveRelative = MoveRelative
	Axis.Position = AxisVal1

	Set Axis = Manipulator.GetAxisFromName(AxisName2)
	Axis.MoveRelative = MoveRelative
	Axis.Position = AxisVal2

	Set Axis = Manipulator.GetAxisFromName(AxisName3)
	Axis.MoveRelative = MoveRelative
	Axis.Position = AxisVal3

	Set Axis = Manipulator.GetAxisFromName(AxisName4)
	Axis.MoveRelative = MoveRelative
	Axis.Position = AxisVal4

	Set Axis = Manipulator.GetAxisFromName(AxisName5)
	Axis.MoveRelative = MoveRelative
	Axis.Position = AxisVal5

	Manipulator.Go(True)
End Sub



'============ Internal Subroutines, do not touch anything below this line ===============
Sub CloseIP(ByVal IPName$)
	If IPEditor.IsIPAvailable(IPName) Then
		IPEditor.CloseIP(IPName)
	End If
End Sub



'============ Internal Subroutines, do not touch anything below this line ===============
Sub SetIP(ByVal IPName$, ByVal IPFileName$)
	If IPEditor.IsIPAvailable(IPName) Then
		IPEditor.CloseIP(IPName)
	End If
	If InStr(IPFileName, "/")=0 And InStr(IPFileName, "\")=0  Then
		IPFileName = MacroDir + "/" + IPFileName
	End If
	IPEditor.LoadIP(IPFileName)
End Sub

