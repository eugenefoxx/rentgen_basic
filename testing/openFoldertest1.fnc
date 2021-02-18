Sub Main()'  saveImage()

Dim L As String
Dim i, p As Integer
Dim IPLive As FNCImageProcess

'Dim strFolder As String
'  L = InputBox$("Enter some text:", _
 '   "Input Box Example", "")
 '   If L = "" Then Exit Sub
  '  NString = "C:\wwtemp\" + L
    NString = "C:\WWTEMP\"

'    L = InputBox$("Enter some text:", _
'    "Input select folders", NString)
'     ActiveWorkbook.FollowHyperlink Address:=NString, NewWindow:=True
	Shell "Explorer.exe C:\WWTEMP\", vbNormalFocus
    F$ = Dir$(L)
    NStringS = InputBox$("Enter some text:", _
    "Input selected folder of Date", F$)
	Shell "Explorer.exe C:\WWTEMP\" + NStringS, vbNormalFocus
'	D$ = NString & NStringS & "\"
'    MsgBox D
    'strFolder = a??C:\exceltrainingvideos\a??

 '   Shell a??Explorer.exe C:\wwtemp\a??, vbNormalFocus
 '   Shell (a??Explorer.exe""C:\wwtemp\", vbNormalFocus)
	Line1:
 For i = 0 To 100
 If MsgBox("Perform image - OK, if not save - Cancel", vbOkCancel) = vbCancel Then
    Debug.Print "OK was pressed"
    Exit Sub
 End If
'	N = N + 1

' edit
	F$ = Dir$
    NStringP = InputBox$("Enter some text:", _
    "Input selected folder PCB", F$)
	For p = 1 To 100
    If (MsgBox("Save image?",vbYesNo) = vbNo) Then
    	index = 0
    	GoTo Line1
    ElseIf Dir$("C:\WWTEMP\" & NStringS & "\" & NStringP & "\", vbDirectory) = "" Then
        index = index + 1
    	Set IPLive = IPEditor.GetIPFromName("Live")
    	IPLive.SaveImage ("C:\WWTEMP\" & NStringS & "\" & NStringP & "\" & NStringP & "_foto_" & index &  ".tif")
    Else
    	If Dir$("C:\WWTEMP\" & NStringS & "\" & NStringP & "\" & NStringP, vbDirectory) = "" Then
    	index = index + 1
        Set IPLive = IPEditor.GetIPFromName("Live")
        IPLive.SaveImage ("C:\WWTEMP\" & NStringS & "\" & NStringP & "\" & NStringP & "_foto_" & index  & ".tif")
        Else
	'		If Dir$("C:\WWTEMP\" & NStringS & "\" & NStringP & "\" & NStringP, vbDirectory) = "" Then
            postfix = getNextPostfix("C:\WWTEMP\" & NStringS & "\" & NStringP & "\" & NStringP & "_foto_" & index  & ".tif")
            If postfix > 0 And postfix < 13 Then
                Set IPLive = IPEditor.GetIPFromName("Live")
                IPLive.SaveImage ("C:\WWTEMP\" & NStringS & "\" & NStringP & "\" & NStringP & "_fotosec_" & postfix &  ".tif")

            End If
        End If
 	End If
 '   Folders$ = Get_DirS(D)
 '   MsgBox Join(Folders, vbCrLf)
 '   D = "savefile.xlsm"
 '   Workbooks(D).Save

'    Folders = CurDir$(D)
'    MsgBox Folders
'    Folders = NStringS & "\"
'    N = N + 1
'    Word = Mid(NStringS & "\" & N & "1_", 1, 6)  '& N & "_" & S
'        MsgBox Word

 '   Set IPLive = IPEditor.GetIPFromName("Live")
 '   IPLive.SaveImage (L)
  Next p
 Next i
	'Set Manipulator to position
	SetManipulator "XM", 0, "YM", 0, "ZT", -200000, "ZD", 300000, "TiltD", 0, fncPUMicrometerMillidegree, False
	SetXRay False, "Microfocus", 130, False, 20, 1, fncBVSmall, False, True
	'Set Manipulator to position
	SetManipulator "XM", 0, "YM", 0, "ZT", -200000, "ZD", 300000, "TiltD", 0, fncPUMicrometerMillidegree, False
	If(FNCIPCleanUp) Then
		CloseIP(FNCLastIP)
	End If
	SetIP "Easy", FNCImageProcessPath & "Easy.ipr"
	If(FNCIPCleanUp) Then
		FNCLastIP="Easy"
	End If
	XRay.Wait(10, fncWFMicroAmp)
	IPEditor.GetIPFromName(IPEditor.ActiveIP).RunOnce(FNCResultPath)
	XRay.Wait(0, fncWFXRayOn)
	Wait 2

End Sub

Function getNextPostfix(ByVal pathDir As String) As Integer


Dim result As Integer, i As Integer
For i = 1 To 15
    If Dir$(pathDir & CStr(i), vbDirectory) = "" Then
           result = i
        Exit For
    End If
Next i
getNextPostfix = result
End Function




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
	XRay.Parameter(VID_XRayOn) = XRayOn
	If OpenDoor = True Then
		Manipulator.OpenDoor(True)
	End If
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

