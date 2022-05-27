VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} BuilderView 
   Caption         =   "Ïîñòðîåíèå ðàçâ¸ðòîê"
   ClientHeight    =   7470
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3765
   OleObjectBlob   =   "BuilderView.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "BuilderView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'===============================================================================

Public IsOk As Boolean
Public IsCancelled As Boolean

'===============================================================================

Private Sub UserForm_Initialize()
  tbTolerance = 5#
  tbBedWidth = 170#
  tbBedHeight = 150#
  tbBedFlap = 60#
  tbCoverFlap = 30#
  CalcFromBed
End Sub

Private Sub UserForm_Activate()
  '
End Sub

Private Sub tbBedWidth_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
  GuardNum KeyAscii
End Sub
Private Sub tbBedWidth_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
  tbBedWidth_AfterUpdate
End Sub
Private Sub tbBedWidth_AfterUpdate()
  GuardRangeDbl tbBedWidth, 0.1
  CalcFromBed
End Sub

Private Sub tbBedHeight_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
  GuardNum KeyAscii
End Sub
Private Sub tbBedHeight_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
  tbBedHeight_AfterUpdate
End Sub
Private Sub tbBedHeight_AfterUpdate()
  GuardRangeDbl tbBedHeight, 0.1
  CalcFromBed
End Sub

Private Sub tbBedFlap_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
  GuardNum KeyAscii
End Sub
Private Sub tbBedFlap_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
  tbBedFlap_AfterUpdate
End Sub
Private Sub tbBedFlap_AfterUpdate()
  GuardRangeDbl tbBedFlap, 0.1
End Sub

Private Sub tbTolerance_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
  GuardNum KeyAscii
End Sub
Private Sub tbTolerance_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
  tbTolerance_AfterUpdate
End Sub
Private Sub tbTolerance_AfterUpdate()
  GuardRangeDbl tbTolerance, 0
  CalcFromBed
  CalcFromCover
End Sub

Private Sub tbCoverWidth_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
  GuardNum KeyAscii
End Sub
Private Sub tbCoverWidth_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
  tbCoverWidth_AfterUpdate
End Sub
Private Sub tbCoverWidth_AfterUpdate()
  GuardRangeDbl tbCoverWidth, 0.1
  CalcFromCover
End Sub

Private Sub tbCoverHeight_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
  GuardNum KeyAscii
End Sub
Private Sub tbCoverHeight_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
  tbCoverHeight_AfterUpdate
End Sub
Private Sub tbCoverHeight_AfterUpdate()
  GuardRangeDbl tbCoverHeight, 0.1
  CalcFromCover
End Sub

Private Sub tbCoverFlap_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
  GuardNum KeyAscii
End Sub
Private Sub tbCoverFlap_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
  tbCoverFlap_AfterUpdate
End Sub
Private Sub tbCoverFlap_AfterUpdate()
  GuardRangeDbl tbCoverFlap, 0.1
End Sub

Private Sub tbLodgmentWidth_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
  GuardNum KeyAscii
End Sub
Private Sub tbLodgmentWidth_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
  tbLodgmentWidth_AfterUpdate
End Sub
Private Sub tbLodgmentWidth_AfterUpdate()
  GuardRangeDbl tbLodgmentWidth, 0.1
  CalcFromLodgment
End Sub

Private Sub tbLodgmentHeight_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
  GuardNum KeyAscii
End Sub
Private Sub tbLodgmentHeight_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
  tbLodgmentHeight_AfterUpdate
End Sub
Private Sub tbLodgmentHeight_AfterUpdate()
  GuardRangeDbl tbLodgmentHeight, 0.1
  CalcFromLodgment
End Sub

Private Sub btnCancel_Click()
  FormCancel
End Sub

Private Sub btnOK_Click()
  FormÎÊ
End Sub

'===============================================================================

Private Sub FormÎÊ()
  Me.Hide
  IsOk = True
End Sub

Private Sub FormCancel()
  Me.Hide
  IsCancelled = True
End Sub

Private Sub CalcFromBed()
  tbCoverWidth.Value = VBA.CStr(VBA.CDbl(tbBedWidth.Value) + VBA.CDbl(tbTolerance.Value))
  tbCoverHeight.Value = VBA.CStr(VBA.CDbl(tbBedHeight.Value) + VBA.CDbl(tbTolerance.Value))
  tbLodgmentWidth.Value = VBA.CStr(VBA.CDbl(tbBedWidth.Value) - VBA.CDbl(tbTolerance.Value))
  tbLodgmentHeight.Value = VBA.CStr(VBA.CDbl(tbBedHeight.Value) - VBA.CDbl(tbTolerance.Value))
End Sub

Private Sub CalcFromCover()
  tbBedWidth.Value = VBA.CStr(VBA.CDbl(tbCoverWidth.Value) - VBA.CDbl(tbTolerance.Value))
  tbBedHeight.Value = VBA.CStr(VBA.CDbl(tbCoverHeight.Value) - VBA.CDbl(tbTolerance.Value))
  tbLodgmentWidth.Value = VBA.CStr(VBA.CDbl(tbBedWidth.Value) - VBA.CDbl(tbTolerance.Value))
  tbLodgmentHeight.Value = VBA.CStr(VBA.CDbl(tbBedHeight.Value) - VBA.CDbl(tbTolerance.Value))
End Sub

Private Sub CalcFromLodgment()
  tbBedWidth.Value = VBA.CStr(VBA.CDbl(tbLodgmentWidth.Value) + VBA.CDbl(tbTolerance.Value))
  tbBedHeight.Value = VBA.CStr(VBA.CDbl(tbLodgmentHeight.Value) + VBA.CDbl(tbTolerance.Value))
  tbCoverWidth.Value = VBA.CStr(VBA.CDbl(tbBedWidth.Value) + VBA.CDbl(tbTolerance.Value))
  tbCoverHeight.Value = VBA.CStr(VBA.CDbl(tbBedHeight.Value) + VBA.CDbl(tbTolerance.Value))
End Sub

'===============================================================================

Private Sub GuardInt(ByVal KeyAscii As MSForms.ReturnInteger)
  Select Case KeyAscii
    Case VBA.Asc("0") To VBA.Asc("9")
    Case Else
      KeyAscii = 0
  End Select
End Sub

Private Sub GuardNum(ByVal KeyAscii As MSForms.ReturnInteger)
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

Private Sub UserForm_QueryClose(Ñancel As Integer, CloseMode As Integer)
  If CloseMode = VbQueryClose.vbFormControlMenu Then
    Ñancel = True
    FormCancel
  End If
End Sub
