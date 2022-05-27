VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BoxBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
Option Explicit

'===============================================================================

Private Type typeThis
  Layer As Layer
  BoxData As structBoxData
End Type
Private This As typeThis

'===============================================================================

Public Function Create(ByVal Layer As Layer, _
                       ByVal BoxData As structBoxData _
                       ) As BoxBuilder
  With New BoxBuilder
    Set Create = .Init(Layer, BoxData)
  End With
End Function

Friend Function Init(ByVal Layer As Layer, _
                     ByVal BoxData As structBoxData _
                     ) As BoxBuilder
  With This
    Set .Layer = Layer
    Set .BoxData = BoxData
  End With
  Set Init = Me
End Function

'===============================================================================

Public Function BuildBox(Optional ByVal LeftX As Double = 0, _
                         Optional ByVal BottomY As Double = 0 _
                         ) As ShapeRange
  
  Set BuildBox = New ShapeRange
  Dim TempShapes As ShapeRange
  Dim TempShapesToAdd As ShapeRange
  
  Set TempShapesToAdd = _
    Helpers.BreakBySegments(BuildBase(LeftX, BottomY))
  SetColoredOutline TempShapesToAdd, This.BoxData.BigColor
  BuildBox.AddRange TempShapesToAdd
  
  If This.BoxData.Flap = 0 Then Exit Function
  
  Set TempShapes = New ShapeRange
  With TempShapes
    .Add BuildBase(LeftX, BottomY)
    .Add BuildBottomFlap(LeftX, BottomY)
    .Add BuildLeftFlap(LeftX, BottomY)
    .Add BuildTopFlap(LeftX, BottomY)
    .Add BuildRightFlap(LeftX, BottomY)
    Set TempShapesToAdd = _
      Helpers.BreakBySegments(lib_elvin.CreateBoundary(TempShapes))
    .Delete
  End With
  SetColoredOutline TempShapesToAdd, This.BoxData.CutColor
  BuildBox.AddRange TempShapesToAdd
  
End Function

Private Function BuildBase(ByVal LeftX As Double, _
                           ByVal BottomY As Double _
                           ) As Shape
 With This
    Set BuildBase = _
      .Layer.CreateRectangle(LeftX + .BoxData.Flap, _
                             BottomY + .BoxData.Flap + .BoxData.Height, _
                             LeftX + .BoxData.Flap + .BoxData.Width, _
                             BottomY + .BoxData.Flap)
  End With
End Function

Private Function BuildBottomFlap(ByVal LeftX As Double, _
                                 ByVal BottomY As Double _
                                 ) As Shape
 With This
    Set BuildBottomFlap = _
      .Layer.CreateRectangle(LeftX + .BoxData.Flap, _
                             BottomY + .BoxData.Flap, _
                             LeftX + .BoxData.Flap + .BoxData.Width, _
                             BottomY)
  End With
End Function

Private Function BuildLeftFlap(ByVal LeftX As Double, _
                               ByVal BottomY As Double _
                               ) As Shape
 With This
    Set BuildLeftFlap = _
      .Layer.CreateRectangle(LeftX, _
                             BottomY + .BoxData.Flap + .BoxData.Height, _
                             LeftX + .BoxData.Flap, _
                             BottomY + .BoxData.Flap)
  End With
End Function

Private Function BuildTopFlap(ByVal LeftX As Double, _
                              ByVal BottomY As Double _
                              ) As Shape
 With This
    Set BuildTopFlap = _
      .Layer.CreateRectangle(LeftX + .BoxData.Flap, _
                             BottomY + .BoxData.Flap * 2 + .BoxData.Height, _
                             LeftX + .BoxData.Flap + .BoxData.Width, _
                             BottomY + .BoxData.Flap + .BoxData.Height)
  End With
End Function

Private Function BuildRightFlap(ByVal LeftX As Double, _
                                ByVal BottomY As Double _
                                ) As Shape
 With This
    Set BuildRightFlap = _
      .Layer.CreateRectangle(LeftX + .BoxData.Flap + .BoxData.Width, _
                             BottomY + .BoxData.Flap + .BoxData.Height, _
                             LeftX + .BoxData.Flap * 2 + .BoxData.Width, _
                             BottomY + .BoxData.Flap)
  End With
End Function

Private Sub SetColoredOutline(ByVal Shapes As ShapeRange, _
                              ByVal ColorString As String)
  Dim Shape As Shape
  For Each Shape In Shapes
    Shape.Outline.Color.CopyAssign CreateColor(ColorString)
  Next Shape
End Sub

Public Property Get Self() As BoxBuilder
  Set Self = Me
End Property

'===============================================================================
