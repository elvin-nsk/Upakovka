VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "structRow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'===============================================================================

Public ItemsRotated As Boolean
Public Width As Double
Public Height As Double
Public Count As Long

'===============================================================================

Public Property Get Self() As structRow
  Set Self = Me
End Property
