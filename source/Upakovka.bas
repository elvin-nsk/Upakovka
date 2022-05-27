Attribute VB_Name = "Upakovka"
'===============================================================================
' Макрос           : Upakovka
' Версия           : 2021.11.11
' Сайт             : https://github.com/elvin-nsk/Upakovka
' Автор            : elvin-nsk (me@elvin.nsk.ru, https://vk.com/elvin_macro)
'===============================================================================

Option Explicit

Public Const RELEASE As Boolean = True

'===============================================================================

Public Const CutColor As String = "RGB255,USER,0,0,0"
Public Const BigColor As String = "RGB255,USER,0,0,255"
Public Const DefaultDocWidth As Double = 1050
Public Const DefaultDocHeight As Double = 900
Public Const GapSize As Double = 0.6
Public Const MinGapsDistance As Double = 30
Public Const MaxGapsDistance As Double = 100

'===============================================================================

Sub Build()

  If RELEASE Then On Error GoTo Catch
  
  Dim BuilderDataSource As BuilderView
  Set BuilderDataSource = New BuilderView
  Dim BoxesData As structBoxesData
  BuilderDataSource.Show
  If Not BuilderDataSource.IsOk Then Exit Sub
  Set BoxesData = GetBoxesData(BuilderDataSource)
  
  BuildCoverAndBed BoxesData
  BuildLodgment BoxesData

Finally:
  Exit Sub

Catch:
  VBA.MsgBox VBA.Err.Description, vbCritical, "Ошибка"
  lib_elvin.BoostFinish
  Resume Finally

End Sub

Sub Impose()

  If RELEASE Then On Error GoTo Catch
  
  Dim Shape As Shape
  With ActivePage.Shapes.All
    If .Count = 0 Then
      Exit Sub
    ElseIf .Count > 1 Then
      Set Shape = .Group
    Else
      Set Shape = .FirstShape
    End If
  End With
  
  lib_elvin.BoostStart "Раскладка", RELEASE
  
  Helpers.Impose _
    Shape, ActivePage.SizeWidth, ActivePage.SizeHeight
    
  Helpers.CenterShapesOnPage ActivePage.Shapes.All, ActivePage
  ActivePage.Shapes.All.UngroupAll

Finally:
  lib_elvin.BoostFinish
  Exit Sub

Catch:
  VBA.MsgBox VBA.Err.Description, vbCritical, "Ошибка"
  Resume Finally

End Sub

Sub Refine()

  If RELEASE Then On Error GoTo Catch
  
  If ActivePage.Shapes.Count = 0 Then Exit Sub
  
  lib_elvin.BoostStart "Дедупликация и расстановка разрывов", RELEASE
    
  ActiveDocument.Unit = cdrMillimeter
  Dim Shape As Shape
  Dim Color As Color
  For Each Shape In ActivePage.Shapes.All
    If Not Helpers.IsValid(Shape) Then
      If Helpers.HaveDisplayCurve(Shape) Then
        Set Color = New Color
        Color.CopyAssign Shape.Outline.Color
        Helpers.SetOutlineColor Helpers.BreakBySegments(Shape), Color
      End If
    End If
  Next Shape
  
  LinesDeduplicator.Create(ActivePage.Shapes.All).Deduplicate
  Dim CutShapes As New ShapeRange
  For Each Shape In ActivePage.Shapes.All
    If Shape.Outline.Color.IsSame(CreateColor(CutColor)) Then _
      CutShapes.Add Shape
  Next Shape
  GapsMaker.Create(CutShapes, GapSize, MinGapsDistance, MaxGapsDistance).MakeGaps

Finally:
  lib_elvin.BoostFinish
  Exit Sub

Catch:
  VBA.MsgBox VBA.Err.Description, vbCritical, "Ошибка"
  Resume Finally

End Sub

'===============================================================================

Private Sub BuildCoverAndBed(ByVal BoxesData As structBoxesData)

  Dim Doc As Document
  Dim Box As ShapeRange
   
  Set Doc = CreateDocument
  SetDocDefaults Doc, "Крышка + дно"
  
  lib_elvin.BoostStart "Крышка + дно", RELEASE
  
  With New structBoxData
    .Width = BoxesData.CoverWidth
    .Height = BoxesData.CoverHeight
    .Flap = BoxesData.CoverFlap
    .CutColor = CutColor
    .BigColor = BigColor
    BoxBuilder.Create(Doc.ActiveLayer, .Self).BuildBox
  End With
  
  With New structBoxData
    .Width = BoxesData.BedWidth
    .Height = BoxesData.BedHeight
    .Flap = BoxesData.BedFlap
    .CutColor = Upakovka.CutColor
    .BigColor = Upakovka.BigColor
    Set Box = BoxBuilder.Create(Doc.ActiveLayer, .Self).BuildBox
  End With
  
  Box.Move Box.SizeWidth + Box.SizeWidth / 5, 0
  Helpers.CenterShapesOnPage Doc.ActivePage.Shapes.All, Doc.ActivePage
  
  lib_elvin.BoostFinish
  
End Sub

Private Sub BuildLodgment(ByVal BoxesData As structBoxesData)

  Dim Doc As Document
  Dim Box As ShapeRange

  Set Doc = CreateDocument
  SetDocDefaults Doc, "Ложемент"
  
  lib_elvin.BoostStart "Ложемент", RELEASE
  
  With New structBoxData
    .Width = BoxesData.LodgmentWidth
    .Height = BoxesData.LodgmentHeight
    .Flap = 0
    .CutColor = CutColor
    .BigColor = CutColor
    Set Box = BoxBuilder.Create(Doc.ActiveLayer, .Self).BuildBox
  End With
  
  Helpers.CenterShapesOnPage Box, Doc.ActivePage
  
  lib_elvin.BoostFinish

End Sub

Private Sub SetDocDefaults(ByVal Doc As Document, ByVal Name As String)
  With Doc
    .Unit = cdrMillimeter
    .Name = Name
    .ActivePage.SizeWidth = DefaultDocWidth
    .MasterPage.SizeWidth = DefaultDocWidth
    .ActivePage.SizeHeight = DefaultDocHeight
    .MasterPage.SizeHeight = DefaultDocHeight
  End With
End Sub

Private Function GetBoxesData(ByVal BuilderDataSource As BuilderView) As structBoxesData
  Set GetBoxesData = New structBoxesData
  With BuilderDataSource
    GetBoxesData.CoverWidth = .tbCoverWidth
    GetBoxesData.CoverHeight = .tbCoverHeight
    GetBoxesData.CoverFlap = .tbCoverFlap
    GetBoxesData.BedWidth = .tbBedWidth
    GetBoxesData.BedHeight = .tbBedHeight
    GetBoxesData.BedFlap = .tbBedFlap
    GetBoxesData.LodgmentWidth = .tbLodgmentWidth
    GetBoxesData.LodgmentHeight = .tbLodgmentHeight
  End With
End Function

'===============================================================================
' тесты
'===============================================================================

Private Sub testBuilderView()
  With New BuilderView
    .Show
  End With
End Sub

Private Sub testDeduplicator()
  With LinesDeduplicator.Create(ActivePage.Shapes.All)
    .Deduplicate
    'Debug.Print .InvalidShapesCount
  End With
End Sub

Private Sub testGapsMaker()
  ActiveDocument.Unit = cdrMillimeter
  With GapsMaker.Create(ActivePage.Shapes.All, _
                        GapSize, MinGapsDistance, MaxGapsDistance)
    .MakeGaps
    'Debug.Print .InvalidShapesCount
  End With
End Sub

Private Sub testApproach1()
  With ImpositionApproach.Create(10, 5, 15, 5, False)
    Debug.Print .tryAddRow(False)
    Debug.Print .tryAddRow(False)
    Debug.Print .tryAddRow(False)
    Debug.Print .Count '1
  End With
End Sub

Private Sub testApproach2()
  With ImpositionApproach.Create(10, 10, 20, 20, False)
    Debug.Print .tryAddRow(False)
    Debug.Print .tryAddRow(False)
    Debug.Print .tryAddRow(False)
    Debug.Print .Count '4
  End With
End Sub

Private Sub testApproach3()
  With ImpositionApproach.Create(3, 2, 7, 12, False)
    Debug.Print .tryAddRow(False)
    Debug.Print .tryAddRow(True)
    Debug.Print .tryAddRow(True)
    Debug.Print .tryAddRow(True)
    Debug.Print .Count '11
    Debug.Print .ImpositionWidth '6
    Debug.Print .ImpositionHeight '11
  End With
End Sub

Private Sub testCalculate()
  Dim Row As Long
  With ImpositionCalculator.Create(3, 2, 7, 12)
    Debug.Print .Count
    If .RowsAlongLayoutWidth Then Debug.Print "Rows along width"
    For Row = 1 To .RowsCount
      Debug.Print "Row(" & Row & ").Count=" & .Row(Row).Count
    Next Row
  End With
End Sub
