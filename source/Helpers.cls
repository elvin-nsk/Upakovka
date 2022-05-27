VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Helpers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Singleton
'@PredeclaredId
Option Explicit

'===============================================================================

Private Sub Class_Initialize()
  If Not Me Is Helpers Then Err.Raise 425 'Invalid object use
End Sub

'===============================================================================

Public Sub CenterShapesOnPage(ByVal Shapes As ShapeRange, Page As Page)
  Shapes.CenterX = Page.CenterX
  Shapes.CenterY = Page.CenterY
End Sub

Public Function BreakBySegments(ByVal Shape As Shape) As ShapeRange
  Set BreakBySegments = New ShapeRange
  Dim Segment As Segment
  For Each Segment In Shape.DisplayCurve.Segments
    BreakBySegments.Add Shape.Layer.CreateCurve(Segment.GetCopy)
  Next Segment
  Shape.Delete
End Function

Public Sub Impose(ByVal Shape As Shape, ByVal Width As Double, ByVal Height As Double)
  Dim Row As Long
  Dim Place As Long
  Dim Offset As Double
  Dim TempShape As Shape
  With ImpositionCalculator.Create(Shape.SizeWidth, Shape.SizeHeight, Width, Height)
    'Debug.Print .RowsAlongLayoutWidth
    For Row = 1 To .RowsCount
      For Place = 1 To .Row(Row).Count
        Set TempShape = Shape.Duplicate
        If .RowsAlongLayoutWidth Then
          If Not .Row(Row).ItemsRotated Then TempShape.Rotate 90
          TempShape.LeftX = Offset
          TempShape.BottomY = (Place - 1) * TempShape.SizeHeight
        Else
          If .Row(Row).ItemsRotated Then TempShape.Rotate 90
          TempShape.LeftX = (Place - 1) * TempShape.SizeWidth
          TempShape.BottomY = Offset
        End If
      Next Place
      Offset = Offset + .Row(Row).Height
    Next Row
    Shape.Delete
  End With
End Sub

Public Function HaveDisplayCurve(ByVal Shape As Shape) As Boolean
  On Error GoTo Finally
  If Not Shape.DisplayCurve Is Nothing Then HaveDisplayCurve = True
Finally:
  On Error GoTo 0
End Function

Public Sub CreatePool(ByVal RawLines As ShapeRange, _
                      ByVal Pool As ShapeRange, _
                      ByRef ioShapesCounter As Long)
  Dim Shape As Shape
  For Each Shape In RawLines
    If IsValid(Shape) Then
      Pool.Add Shape
    Else
      ioShapesCounter = ioShapesCounter + 1
    End If
  Next Shape
End Sub

Public Function IsValid(ByVal Shape As Shape) As Boolean
  On Error GoTo Finally
  If Shape.Curve.Nodes.Count = 2 Then IsValid = True
Finally:
  On Error GoTo 0
End Function

Public Sub SetOutlineColor(ByVal Shapes As ShapeRange, ByVal Color As Color)
  Dim Shape As Shape
  For Each Shape In Shapes
    Shape.Outline.Color.CopyAssign Color
  Next Shape
End Sub

'===============================================================================
