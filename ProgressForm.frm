Private Sub UserForm_Initialize()
    ' Initialize the progress label
    Me.ProgressLabel.Caption = "Processing: 0%"
End Sub

Public Sub UpdateProgress(Current As Long, Total As Long)
    ' Update the progress percentage
    Dim Percentage As Double
    Percentage = (Current / Total) * 100
    Me.ProgressLabel.Caption = "Processing: " & Format(Percentage, "0") & "%"
    Me.Repaint
End Sub

