VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "structBoxData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'===============================================================================

Public Width As Double
Public Height As Double
Public Flap As Double
Public CutColor As String
Public BigColor As String

'===============================================================================

Public Property Get Self() As structBoxData
  Set Self = Me
End Property
