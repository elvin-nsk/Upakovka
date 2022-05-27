VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ImpositionCalculator"
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
  SuccessfulApproach As ImpositionApproach
End Type
Private This As typeThis

'===============================================================================

Public Function Create(ByVal ItemWidth As Double, _
                       ByVal ItemHeight As Double, _
                       ByVal LayoutWidth As Double, _
                       ByVal LayoutHeight As Double _
                       ) As ImpositionCalculator
  With New ImpositionCalculator
    Set Create = .Init(ItemWidth, ItemHeight, LayoutWidth, LayoutHeight)
  End With
End Function

Friend Function Init(ByVal ItemWidth As Double, _
                     ByVal ItemHeight As Double, _
                     ByVal LayoutWidth As Double, _
                     ByVal LayoutHeight As Double _
                     ) As ImpositionCalculator
  With This
    .ItemWidth = ItemWidth
    .ItemHeight = ItemHeight
    .LayoutWidth = LayoutWidth
    .LayoutHeight = LayoutHeight
  End With
  Calculate
  Set Init = Me
End Function

'===============================================================================

Public Property Get Row(ByVal RowNum As Long) As structRow
  Set Row = This.SuccessfulApproach.Row(RowNum)
End Property

Public Property Get RowsCount() As Long
  RowsCount = This.SuccessfulApproach.RowsCount
End Property

Public Property Get RowsAlongLayoutWidth() As Boolean
  RowsAlongLayoutWidth = This.SuccessfulApproach.RowsAlongLayoutWidth
End Property

Public Property Get Count() As Long
  Count = This.SuccessfulApproach.Count
End Property

Public Property Get Self() As ImpositionCalculator
  Set Self = Me
End Property

'===============================================================================

Private Sub Calculate()
  
  Dim Approaches As New Collection
  Dim AproachesPlaces As New Collection
  Dim MinItemSize As Double
  Dim Row As Long
  Dim Column As Long
  
  With This
  
    MinItemSize = lib_elvin.MinOfTwo(.ItemWidth, .ItemHeight)
    
    For Row = 1 To VBA.Fix(.LayoutHeight / MinItemSize) + 1
      With PopulateApproach(Row, False)
        Approaches.Add .Self
        AproachesPlaces.Add .Count
      End With
      With PopulateApproach(Row, True)
        Approaches.Add .Self
        AproachesPlaces.Add .Count
      End With
    Next Row
    
    Set This.SuccessfulApproach = Approaches(lib_elvin.FindMaxItemNum(AproachesPlaces))
    
  End With
  
End Sub

Private Function PopulateApproach(ByVal tryRotatedRows As Long, _
                                  ByVal RunRowsAlongLayoutWidth As Boolean _
                                  ) As ImpositionApproach
  With This
    Dim Row As Long
    Dim LayoutWidth As Double, LayoutHeight As Double
    LayoutWidth = VBA.IIf(RunRowsAlongLayoutWidth, .LayoutHeight, .LayoutWidth)
    LayoutHeight = VBA.IIf(RunRowsAlongLayoutWidth, .LayoutWidth, .LayoutHeight)
    Set PopulateApproach = _
      ImpositionApproach.Create(.ItemWidth, .ItemHeight, _
                                LayoutWidth, LayoutHeight, RunRowsAlongLayoutWidth)
    Do
      Row = Row + 1
    Loop Until Not PopulateApproach.tryAddRow(Row <= tryRotatedRows)
  End With
End Function
