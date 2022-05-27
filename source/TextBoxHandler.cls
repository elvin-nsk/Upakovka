VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TextBoxHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
Option Explicit

'===============================================================================

Private WithEvents ThisControl As MSForms.TextBox
Attribute ThisControl.VB_VarHelpID = -1

'===============================================================================

Public Function Create(ByVal Control As MSForms.TextBox) As TextBoxHandler
  With New TextBoxHandler
    Set Create = .Init(Control)
  End With
End Function

Private Sub Class_Terminate()
  '
  ' код деструктора
  '
End Sub

Friend Function Init(ByVal Control As MSForms.TextBox) As TextBoxHandler
  Set ThisControl = Control
  Set Init = Me
End Function

'===============================================================================

Public Property Get SomeProperty() As Variant
  SomeProperty = This.SomeValue
End Property
Public Property Let SomeProperty(ByVal RHS As Variant)
  This.SomeValue = RHS
End Property
Public Property Set SomeProperty(ByVal RHS As Variant)
  Set This.SomeValue = RHS
End Property

Public Sub SomeSub(ByVal Value As Variant)
End Sub

Public Function SomeFunction(ByVal Value As Variant) As Variant
End Function

Public Property Get Self() As TextBoxHandler
  Set Self = Me
End Property

'===============================================================================

Private Sub ThisControl_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
  OnlyNum KeyAscii
End Sub
Private Sub ThisControl_AfterUpdate()
  GuardRangeDbl ThisControl, 0.1
  CalcFromCover
End Sub
Private Sub ThisControl_Change()
  GuardRangeDbl ThisControl, 0.1
  CalcFromCover
End Sub





'===============================================================================

Private Sub OnlyInt(ByVal KeyAscii As MSForms.ReturnInteger)
  Select Case KeyAscii
    Case VBA.Asc("0") To VBA.Asc("9")
    Case Else
      KeyAscii = 0
  End Select
End Sub

Private Sub OnlyNum(ByVal KeyAscii As MSForms.ReturnInteger)
  Select Case KeyAscii
    Case VBA.Asc("0") To VBA.Asc("9")
    Case VBA.Asc(",")
    Case Else
      KeyAscii = 0
  End Select
End Sub

Private Sub GuardRangeDbl(TextBox As MSForms.TextBox, ByVal Min As Double, Optional ByVal Max As Double = 2147483647)
  With TextBox
    If .Value = "" Then .Value = VBA.CStr(Min)
    If VBA.CDbl(.Value) > Max Then .Value = VBA.CStr(Max)
    If VBA.CDbl(.Value) < Min Then .Value = VBA.CStr(Min)
  End With
End Sub

Private Sub GuardRangeLng(TextBox As MSForms.TextBox, ByVal Min As Long, Optional ByVal Max As Long = 2147483647)
  With TextBox
    If .Value = "" Then .Value = VBA.CStr(Min)
    If VBA.CLng(.Value) > Max Then .Value = VBA.CStr(Max)
    If VBA.CLng(.Value) < Min Then .Value = VBA.CStr(Min)
  End With
End Sub
