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

Private WithEvents pMyLabel As MSForms.Label
Attribute pMyLabel.VB_VarHelpID = -1
Private pDateValue As Date

Property Let DateValue(RHS As Date)
    pDateValue = RHS
End Property

Property Get DateValue() As Date
    DateValue = pDateValue
End Property

Property Set MyLabel(RHS As MSForms.Label)
    Set pMyLabel = RHS
End Property

Property Get MyLabel() As MSForms.Label
    Set MyLabel = pMyLabel
End Property

Private Sub pMyLabel_Click()
    'MsgBox MyLabel.Name & ": " & pDateValue
    Call LoadRoomDetails(pDateValue)
End Sub
