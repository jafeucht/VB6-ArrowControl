VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "csArrow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Sub Class_Initialize()
    ArrowStr = "[0,4]-[-2,-3]-[0,-2]-[2,-3]-[0,4]"
    ObjMult = 1
End Sub

Property Get ArrowString() As String
    ArrowString = ArrowStr
End Property

Property Let ArrowString(NewValue As String)
    ArrowStr = NewValue
End Property

Property Get ArrowPosX() As Double
    ArrowPosX = ObjPos.X
End Property

Property Let ArrowPosX(NewValue As Double)
    ObjPos.X = NewValue
End Property

Property Get ArrowPosY() As Double
    ArrowPosY = ObjPos.Y
End Property

Property Let ArrowPosY(NewValue As Double)
    ObjPos.Y = NewValue
End Property

Public Sub SetArrowPos(X As Double, Y As Double)
    ObjPos = GetPoint(X, Y)
End Sub

Property Get ArrowZoom() As Double
    ArrowZoom = ObjMult
End Property

Property Let ArrowZoom(NewValue As Double)
    ObjMult = NewValue
End Property

Property Get ArrowRotate() As Integer
    ArrowRotate = Rotate
End Property

Property Let ArrowRotate(NewValue As Integer)
    Rotate = NewValue
End Property

Property Get ArrowContainer() As Object
    Set ArrowContainer = LinkObject
End Property

Property Set ArrowContainer(NewObject As Object)
    Set LinkObject = NewObject
End Property

Property Get ArrowDisabled() As Boolean
    ArrowDisabled = Disabled
End Property

Property Let ArrowDisabled(NewValue As Boolean)
    Disabled = NewValue
End Property

Public Sub Refresh()
    DrawArrow
End Sub
