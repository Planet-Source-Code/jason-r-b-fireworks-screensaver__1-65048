Attribute VB_Name = "StartUp"
Public Declare Function InitCommonControls Lib "comctl32.dll" () As Long

Sub Main()
    If App.PrevInstance = True Then
        End
    End If
    Call InitCommonControls
    Select Case Left(UCase(Command$), 2)
        Case "/C"
            Form2.Show
        Case "/S"
            Form1.Show
    End Select
End Sub

