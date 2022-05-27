VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LinesDeduplicator"
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
End Type
Private This As typeThis

'===============================================================================

Public Function Create(ByVal Lines As ShapeRange) As LinesDeduplicator
  With New LinesDeduplicator
    Set Create = .Init(Lines)
  End With
End Function

Friend Function Init(ByVal Lines As ShapeRange) As LinesDeduplicator
  With This
    Set .Pool = New ShapeRange
    Helpers.CreatePool Lines, .Pool, .InvalidShapesCount
  End With
  Set Init = Me
End Function

'===============================================================================

Public Property Get InvalidShapesCount() As Long
  InvalidShapesCount = This.InvalidShapesCount
End Property

Public Sub Deduplicate()
  With This
    If .Pool.Count < 2 Then Exit Sub
    Do Until .Pool.Count < 2
      CheckLine .Pool(1)
      .Pool.Remove 1
    Loop
  End With
End Sub

Public Property Get Self() As LinesDeduplicator
  Set Self = Me
End Property

'===============================================================================

Private Sub CheckLine(ByVal Line As Shape)
  With This
    Dim LocalPool As New ShapeRange
    LocalPool.AddRange .Pool
    Do Until LocalPool.Count = 0
      CheckTwoLines Line, LocalPool(1)
      LocalPool.Remove 1
    Loop
  End With
End Sub

Private Sub CheckTwoLines(ByVal Line1 As Shape, ByVal Line2 As Shape)
  If lib_elvin.IsNothing(Line1) Or lib_elvin.IsNothing(Line2) Then Exit Sub
  If IsSameShape(Line1, Line2) Then Exit Sub
  If IsFirstFullOnSecond(Line1, Line2) Then
    Line1.Delete
    Exit Sub
  End If
  If IsSecondFullOnFirst(Line1, Line2) Then
    Line2.Delete
    Exit Sub
  End If
  With NodesOfPartiallyIntersect(Line1, Line2)
    If .Count > 0 Then
      .Item(1).PositionX = .Item(2).PositionX
      .Item(1).PositionY = .Item(2).PositionY
    End If
  End With
End Sub

Private Function IsFirstFullOnSecond(ByVal Line1 As Shape, _
                                     ByVal Line2 As Shape _
                                     ) As Boolean
  With Line1.Curve.Nodes
    If Line2.Curve.IsOnCurve(.First.PositionX, .First.PositionY) > 0 And _
       Line2.Curve.IsOnCurve(.Last.PositionX, .Last.PositionY) > 0 Then _
      IsFirstFullOnSecond = True
  End With
End Function

Private Function IsSecondFullOnFirst(ByVal Line1 As Shape, _
                                     ByVal Line2 As Shape _
                                     ) As Boolean
  With Line2.Curve.Nodes
    If Line1.Curve.IsOnCurve(.First.PositionX, .First.PositionY) > 0 And _
       Line1.Curve.IsOnCurve(.Last.PositionX, .Last.PositionY) > 0 Then _
      IsSecondFullOnFirst = True
  End With
End Function

Private Function NodesOfPartiallyIntersect(ByVal Line1 As Shape, _
                                            ByVal Line2 As Shape _
                                            ) As Collection
  Set NodesOfPartiallyIntersect = New Collection
  Dim Node1 As Node
  Dim Node2 As Node
  Set Node1 = FindNodeOnLine(Line1.Curve.Nodes, Line2)
  Set Node2 = FindNodeOnLine(Line2.Curve.Nodes, Line1)
  If (Not Node1 Is Nothing) And (Not Node2 Is Nothing) Then
    NodesOfPartiallyIntersect.Add Node1
    NodesOfPartiallyIntersect.Add Node2
  End If
End Function

Private Function FindNodeOnLine(ByVal Nodes As Nodes, _
                                ByVal Line As Shape _
                                ) As Node
  Dim Node As Node
  For Each Node In Nodes
    If Line.Curve.IsOnCurve(Node.PositionX, Node.PositionY) Then
      Set FindNodeOnLine = Node
      Exit Function
    End If
  Next Node
End Function

Private Function IsSameShape(ByVal Shape1 As Shape, Shape2 As Shape) As Boolean
  IsSameShape = (Shape1.StaticID = Shape2.StaticID)
End Function
