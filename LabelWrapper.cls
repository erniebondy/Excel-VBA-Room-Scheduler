VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LabelWrapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents MyLabel As MSForms.Label
Attribute MyLabel.VB_VarHelpID = -1

Private Sub MyLabel_Click()
    MsgBox "Clicked! " & MyLabel.Name
End Sub
