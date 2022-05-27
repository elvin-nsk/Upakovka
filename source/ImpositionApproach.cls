VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ImpositionApproach"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
Option Explicit

'===============================================================================

Private Type typeThis
  ItemWidth As Double
  ItemHeight As Double
  LayoutWidth As Double
  LayoutHeight As Double
  RowsAlongLayoutWidth As Boolean
  Rows As Collection
  Width As Double
  Height As Double
  Count As Long
End Type
Private This As typeThis

'===============================================================================

Public Function Create(ByVal ItemWidth As Double, _
                       ByVal ItemHeight As Double, _
                       ByVal LayoutWidth As Double, _
                       ByVal LayoutHeight As Double, _
                       ByVal RowsAlongLayoutWidth As Boolean _
                       ) As ImpositionApproach
  With New ImpositionApproach
    Set Create = .Init(ItemWidth, ItemHeight, LayoutWidth, LayoutHeight, RowsAlongLayoutWidth)
  End With
End Function

Friend Function Init(ByVal ItemWidth As Double, _
                     ByVal ItemHeight As Double, _
                     ByVal LayoutWidth As Double, _
                     ByVal LayoutHeight As Double, _
                     ByVal RowsAlongLayoutWidth As Boolean _
                     ) As ImpositionApproach
  With This
    .ItemWidth = ItemWidth
    .ItemHeight = ItemHeight
    .LayoutWidth = LayoutWidth
    .LayoutHeight = LayoutHeight
    .RowsAlongLayoutWidth = RowsAlongLayoutWidth
    Set .Rows = New Collection
  End With
  Set Init = Me
End Function

'===============================================================================

Public Property Get Count() As Long
  Count = This.Count
End Property

Public Property Get Row(ByVal RowNum As Long) As structRow
  Set Row = This.Rows(RowNum)
End Property

Public Property Get RowsCount() As Long
  RowsCount = This.Rows.Count
End Property

Public Property Get ImpositionWidth() As Long
  ImpositionWidth = This.Width
End Property

Public Property Get ImpositionHeight() As Long
  ImpositionHeight = This.Height
End Property

Public Property Get RowsAlongLayoutWidth() As Boolean
  RowsAlongLayoutWidth = This.RowsAlongLayoutWidth
End Property

Public Function tryAddRow(ByVal ItemsRotated As Boolean) As Boolean
  With This
    Dim Row As New structRow
    Dim ItemWidth As Double
    Row.Height = VBA.IIf(ItemsRotated, .ItemWidth, .ItemHeight)
    If .Height + Row.Height <= .LayoutHeight Then
      ItemWidth = VBA.IIf(ItemsRotated, .ItemHeight, .ItemWidth)
      Row.Count = VBA.Fix(.LayoutWidth / ItemWidth)
      Row.Width = Row.Count * ItemWidth
      Row.ItemsRotated = ItemsRotated
      .Rows.Add Row
      tryAddRow = True
      .Height = .Height + Row.Height
      If .Width < Row.Width Then .Width = Row.Width
      .Count = .Count + Row.Count
    End If
  End With
End Function

Public Property Get Self() As ImpositionApproach
  Set Self = Me
End Property

'===============================================================================
