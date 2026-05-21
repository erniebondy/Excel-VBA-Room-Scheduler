VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ControlWrappers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mControls As Collection
Private WithEvents mActiveControl As ControlWrapper

Event ActiveControlMouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Event ActiveControlClick()

Sub Add(Control As ControlWrapper)
    Set Control.Collection = Me
    mControls.Add Control
End Sub

Property Set ActiveControl(RHS As ControlWrapper)
    Set mActiveControl = RHS
End Property

Property Get ActiveControl() As ControlWrapper
    Set ActiveControl = mActiveControl
End Property

Public Property Get Controls() As Collection
    Set Controls = mControls
End Property

Private Sub Class_Initialize()
    Set mControls = New Collection
End Sub

Private Sub mActiveControl_MyClick()
    RaiseEvent ActiveControlClick
End Sub

Private Sub mActiveControl_MyMouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    RaiseEvent ActiveControlMouseMove(Button, Shift, X, Y)
End Sub
