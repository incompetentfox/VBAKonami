Option Explicit

Dim order As Variant
Dim index As Integer

Private Sub UserForm_Initialize()
' Initialise the array of KeyDown values (up up down down left right left right a b)
    order = Array(38, 38, 40, 40, 37, 39, 37, 39, 65, 66)
End Sub


Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
' listen for keypresses that match the order in the array

    If KeyCode = order(index) Then
        kcode
    Else
        index = 0
    End If
End Sub

Private Sub kcode()
	'routine to increase the array index until it reaches the end of the sequence.
    index = index + 1
    If index = 10 Then
        index = 0
        ' do something.
    End If
    
End Sub

