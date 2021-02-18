Sub Main()
Dim L As String
Dim IPLive As FNCImageProcess

'Dim strFolder As String
  L = InputBox$("Enter some text:", _
    "Input Box Example", "")
    If L = "" Then Exit Sub
    NString = "C:\WWTEMP\" + L

'strFolder = “C:\exceltrainingvideos\”
' ActiveWorkbook.FollowHyperlink Address:=NString, NewWindow:=True
    Shell "Explorer.exe C:\WWTEMP\" + L, vbNormalFocus
 '   Shell (“Explorer.exe""C:\wwtemp\", vbNormalFocus)
 If MsgBox("Perform image - OK, if not save - Cancel", vbOkCancel) = vbCancel Then
    Debug.Print "OK was pressed"
    Exit Sub
 End If
    Set IPLive = IPEditor.GetIPFromName("Live")
    IPLive.SaveImage ("C:\WWTEMP\" & L & "\" & "test.tif")
	'Set Manipulator to position
	SetManipulator "XM", 0, "YM", 0, "ZT", -200000, "ZD", 300000, "TiltD", 0, fncPUMicrometerMillidegree, False
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
