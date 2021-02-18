'Main program
Sub Main()

'Declare variables:
Dim axis As FNCAxis
Dim index As Integer
Dim FileName As String
Dim FileNameStem As String
Dim ip As FNCImageProcess
Dim bestImageNo As Integer

'Activate error handling:
On Error GoTo errorhandler

'User input: 
FileNameStem = InputBox("Enter file name stem","Stem input","C:\WWTEMP\sample_image")
index = 0

'Assign object references:
Set ip = IPEditor.GetIPFromName(IPEditor.ActiveIP)
Set axis = Manipulator.GetAxis(fncAIX)

'Access to object methods and properties:
ip.Start
axis.MoveRelative =True
axis.Position = 1000

'Do loop (is executed at least once):
Do

axis.Go True

'Procedure call:
	DoImageProcessing(ip)
	'User interaction and If branching:
	If (MsgBox("Save image?",vbYesNo) = vbYes) Then
		index = index + 1

'String operations:
			'Concatenation with & and convertion of a number into a string with Str()

FileName = FileNameStem & Str(index) & ".tif"
		'File output:

ip.Start
			ip.SaveImage(FileName, 0)

End If
'loop exit condition: user clicks on "No" in the message box
Loop While MsgBox("Seen enough?", vbYesNo Or vbDefaultButton2) = vbNo

'Function call:

bestImageNo = ChooseBestImage(index,FileNameStem)

'Output into message window:

Debug.Print "Selected image: " & Str(bestImageNo)

'Jump (exit program):

Exit Sub

'Error handling code:

'label:

errorhandler:

'Data output via message box:

'Inserting line feeds with vbCrLf in strings:

MsgBox "An error occurred!" & vbCrLf & "Error code: " & Err.Number & vbCrLf & "Description: " & Err.Description

End Sub

'Procedure definition:

'Parameter hand-over "by reference":

Sub DoImageProcessing(ByRef ip As FNCImageProcess)
	ip.RunOnce
End Sub

'Function definition:
'Parameter hand-over "by value":
Function ChooseBestImage (ByVal number As Integer, ByVal nameStem As String) As Integer
	Dim FileName As String
	Dim index As Integer
	Dim prompt As String
	Dim title As String
	Dim originalIP As String

	'store the currently active image process in originalIP:

originalIP=IPEditor.ActiveIP
	'activate the (necessarily empty) image process "New"
	IPEditor.ActiveIP = "New"

	'Initialise the function's return value:
	ChooseBestImage = 0

	'For loop
	'is executed as long as index is not greater than number:
	For index = 1 To number
		FileName = nameStem & Str(index) & ".tif"
		'Read image data from an image file:
		IPEditor.GetIPFromName(IPEditor.ActiveIP).LoadImage(FileName)

'Querying the user:
		prompt = "Is this image good enough?"
		title = "Image selection"
		decision=MsgBox (prompt,vbYesNo,title)
		If decision = vbYes Then
			'Set the function's return value:
			ChooseBestImage=index
			'Jump (exit "For" loop):
			Exit For
		End If
	Next
	'restore the original state of the image processes:
	IPEditor.GetIPFromName("New").ActivateCamera

IPEditor.ActiveIP = originalIP
End Function
