VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ControlWrapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mMyControl As Object
Private WithEvents mMyForm As MSForms.UserForm
Attribute mMyForm.VB_VarHelpID = -1
Private WithEvents mMyListBox As MSForms.ListBox
Attribute mMyListBox.VB_VarHelpID = -1
Private WithEvents mMyLabel As MSForms.Label
Attribute mMyLabel.VB_VarHelpID = -1
Private WithEvents mMyButton As MSForms.CommandButton
Attribute mMyButton.VB_VarHelpID = -1

' PARENT PROCEDURES

Private Sub ControlWrapper_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Debug.Print mMyControl.Name & " MouseMove "; X; Y
End Sub

' PROPERTIES

Public Property Set MyForm(RHS As Object)
    Set mMyForm = RHS
    Set mMyControl = RHS
End Property

Public Property Get MyForm() As MSForms.UserForm
    Set MyForm = mMyForm
End Property

Public Property Set MyLabel(RHS As MSForms.Label)
    Set mMyLabel = RHS
    Set mMyControl = RHS
End Property

Public Property Get MyLabel() As MSForms.Label
    Set MyLabel = mMyLabel
End Property

Public Property Set MyListBox(RHS As MSForms.ListBox)
    Set mMyListBox = RHS
    Set mMyControl = RHS
End Property

Public Property Get MyListBox() As MSForms.ListBox
    Set MyListBox = mMyListBox
End Property

' PROCEDURES

Private Sub mMyForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call ControlWrapper_MouseMove(Button, Shift, X, Y)
End Sub

Private Sub mMyLabel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call ControlWrapper_MouseMove(Button, Shift, X, Y)
End Sub

Private Sub mMyListBox_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call ControlWrapper_MouseMove(Button, Shift, X, Y)
End Sub
