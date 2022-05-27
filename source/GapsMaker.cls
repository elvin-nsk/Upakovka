VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GapsMaker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
Option Explicit

'===============================================================================

Private Type typeThis
  Pool As ShapeRange
  InvalidShapesCount As Long
  GapSize As Double
  MinGapsDistance As Double
  MaxGapsDistance As Double
End Type
Private This As typeThis

'===============================================================================

Public Function Create(ByVal Lines As ShapeRange, _
                       ByVal GapSize As Double, _
                       ByVal MinGapsDistance As Double, _
                       ByVal MaxGapsDistance As Double _
                       ) As GapsMaker
  With New GapsMaker
    Set Create = .Init(Lines, GapSize, MinGapsDistance, MaxGapsDistance)
  End With
End Function

Friend Function Init(ByVal Lines As ShapeRange, _
                     ByVal GapSize As Double, _
                     ByVal MinGapsDistance As Double, _
                     ByVal MaxGapsDistance As Double _
                     ) As GapsMaker
  With This
    .GapSize = GapSize
    .MinGapsDistance = MinGapsDistance
    .MaxGapsDistance = MaxGapsDistance
    Set .Pool = New ShapeRange
    Helpers.CreatePool Lines, .Pool, .InvalidShapesCount
  End With
  Set Init = Me
End Function

'===============================================================================

Public Property Get InvalidShapesCount() As Long
  InvalidShapesCount = This.InvalidShapesCount
End Property

Public Sub MakeGaps()
  Dim Line As Shape
  For Each Line In This.Pool
    ProcessLine Line
  Next Line
End Sub

Public Property Get Self() As GapsMaker
  Set Self = Me
End Property

'===============================================================================

Private Sub ProcessLine(ByVal Line As Shape)
  With This
    Dim LineLength As Double
    LineLength = Line.Curve.Length
    If LineLength < .MinGapsDistance Then Exit Sub
    Dim GapsNum As Long
    Dim FirstGap As Double
    Dim GapsDistance As Double
    Dim Gaps As New Collection
    Dim i As Long
    GapsNum = VBA.Fix(LineLength \ .MaxGapsDistance)
    FirstGap = (LineLength - GapsNum * .MaxGapsDistance) / 2
    If FirstGap < .MinGapsDistance And GapsNum > 0 Then
      GapsDistance = LineLength \ (GapsNum + 1)
      For i = 1 To GapsNum Step 1
        Gaps.Add i * GapsDistance
      Next i
    Else
      GapsDistance = .MaxGapsDistance
      For i = 0 To GapsNum Step 1
        Gaps.Add FirstGap + i * GapsDistance
      Next i
    End If
    
    SetGaps Line, Gaps
  End With
End Sub

Private Sub SetGaps(ByVal Line As Shape, ByVal Gaps As Collection)
  Dim Knives As New ShapeRange
  Dim X As Double, Y As Double
  Dim i As Long
  For i = 1 To Gaps.Count
    Line.Curve.Segments.First.GetPointPositionAt X, Y, Gaps(i), cdrAbsoluteSegmentOffset
    Knives.Add Line.Layer.CreateEllipse2(X, Y, This.GapSize / 2)
  Next i
  Knives.Combine.Trim Line, False, False
End Sub
