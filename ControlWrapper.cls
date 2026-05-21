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

Private mCollection As ControlWrappers
Attribute mCollection.VB_VarHelpID = -1
Private mMyControl As Object
Private WithEvents mMyForm As MSForms.UserForm
Attribute mMyForm.VB_VarHelpID = -1
Private WithEvents mMyListBox As MSForms.ListBox
Attribute mMyListBox.VB_VarHelpID = -1
Private WithEvents mMyLabel As MSForms.Label
Attribute mMyLabel.VB_VarHelpID = -1
Private WithEvents mMyButton As MSForms.CommandButton
Attribute mMyButton.VB_VarHelpID = -1

Private mDefaultBackColor As Long
Private mBackColor As Long

Event MyMouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Event MyClick()


' PARENT PROCEDURES



Private Sub MouseMoveHandler(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Debug.Print mMyControl.Name & " MouseMove "; X; Y
    Set mCollection.ActiveControl = Me
    RaiseEvent MyMouseMove(Button, Shift, X, Y)
End Sub

Private Sub ClickHandler()
    RaiseEvent MyClick
End Sub

Sub ResetBackColor()
    BackColor = mDefaultBackColor
End Sub

' PROPERTIES

Public Property Set MyForm(RHS As Object)
    Set mMyForm = RHS
    Set MyControl = RHS
End Property

Public Property Get MyForm() As MSForms.UserForm
    Set MyForm = mMyForm
End Property

Public Property Set MyLabel(RHS As MSForms.Label)
    Set mMyLabel = RHS
    Set MyControl = RHS
End Property

Public Property Get MyLabel() As MSForms.Label
    Set MyLabel = mMyLabel
End Property

Public Property Set MyListBox(RHS As MSForms.ListBox)
    Set mMyListBox = RHS
    Set MyControl = RHS
End Property

Public Property Get MyListBox() As MSForms.ListBox
    Set MyListBox = mMyListBox
End Property

Public Property Set MyControl(RHS As Object)
    Set mMyControl = RHS
    mDefaultBackColor = RHS.BackColor
End Property

Public Property Get MyControl() As Object
    Set MyControl = mMyControl
End Property

Public Property Set Collection(RHS As ControlWrappers)
    Set mCollection = RHS
End Property

Public Property Get Collection() As ControlWrapper
    Set Collection = mCollection
End Property

Public Property Let BackColor(RHS As Long)
    mBackColor = RHS
    mMyControl.BackColor = RHS
End Property

Public Property Get BackColor() As Long
    BackColor = mBackColor
End Property

' PROCEDURES

Private Sub mMyForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call MouseMoveHandler(Button, Shift, X, Y)
End Sub

Private Sub mMyLabel_Click()
    Call ClickHandler
End Sub

Private Sub mMyLabel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call MouseMoveHandler(Button, Shift, X, Y)
End Sub

Private Sub mMyListBox_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call MouseMoveHandler(Button, Shift, X, Y)
End Sub
