VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "structBoxesData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'===============================================================================

Public CoverWidth As Double
Public CoverHeight As Double
Public CoverFlap As Double
Public BedWidth As Double
Public BedHeight As Double
Public BedFlap As Double
Public LodgmentWidth As Double
Public LodgmentHeight As Double
Public Tolerance As Double

'===============================================================================

Public Property Get Self() As structBoxesData
  Set Self = Me
End Property
