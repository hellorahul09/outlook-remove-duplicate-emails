VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ProgressForm 
   Caption         =   "UserForm1"
   ClientHeight    =   3040
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   4580
   OleObjectBlob   =   "ProgressForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ProgressForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

