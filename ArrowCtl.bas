Attribute VB_Name = "mdArrow"
Option Explicit

' Arrow:
' [0,3]-[3,0]-[1,0]-[1,-4]-[-1,-4]-[-1,0]-[-3,0]-[0,3]
' Odd designs:
' [0,2]-[2,-2]-[-2,-2]-[0,2]:[-4,4]-[4,4]-[0,-4]-[-4,4]
' Gauge needle:
' [0,8]-[.25,7.75]-[1,1]-[1,0]-[.75,-.75]-[0,-1]-[-.75,-.75]-[-1,0]-[-1,1]-[-.25,7.75]-[0,8]
' Airplane:
' [0,10]-[2,7]-[2,5]-[0,4]-[-2,5]-[-2,7]-[0,10]:[0,17]-[3,13]-[4,0]-[3,-5]-[3,-12]-[-3,-12]-[-3,-5]-[-4,0]-[-3,13]-[0,17]:[0,-5]-[0,-10]:[4,0]-[10,-5]-[3,-5]:[-4,0]-[-10,-5]-[-3,-5]:[3,-7]-[7,-12]-[3,-12]:[-3,-7]-[-7,-12]-[-3,-12]
' Missile:
' [-.25,10]-[.25,10]-[1.25,9]-[1.5,8]-[1.5,-3]-[.5,-4]-[-.5,-4]-[-1.5,-3]-[-1.5,8]-[-1.25,9]-[-.25,10]:[-1.5,6.5]-1.5,6.5]:[-1.5,6]-[1.5,6]:[-1.5,6.5]-[1.5,6.5]:[.5,3]-[-.25,3]-[-.5,3.25]-[-.5,3.75]-[-.25,4]-[.5,4]:[0,4]-[0,3]:[-.5,2.5]-[-.5,2]-[-.25,1.75]-[0,2]-[0,2.25]-[.25,2.5]-[.5,2.25]-[.5,1.75]:[-.5,1.25]-[.25,1.25]-[.5,1]-[.5,.5]-[.25,.25]-[-.5,.25]:[-1.5,-1]-[-3.5,-1.75]-[-3.5,-3]-[-1.5,-3]:[1.5,-1]-[3.5,-1.75]-[3.5,-3]-[1.5,-3]:[-1.5,-1.75]-[1.5,-1.75]:[-1.5,-2.25]-[1.5,-2.25]:[0,-1]-[0,-3]:[-.5,-4]-[-1,-4.5]-[1,-4.5]-[.5,-4]
' Thermas Top:
' [-.25,3]-[.25,3]-[.5,2.75]-[.5,-.25]-[.25,-.5]-[-.25,-.5]-[-.5,-.25]-[-.5,2.75]-[-.25,3]-[-.25,3]:[.5,1.75]-[1,1.75]-[1.75,1]-[1.75,.75]-[2,.5]-[2,-.5]-[1.75,-.75]-[1.75,-1]-[1,-1.75]-[.5,-1.75]-[-.5,-1.75]-[-1,-1.75]-[-1.75,-1]-[-1.75,-.75]-[-2,-.5]-[-2,.5]-[-1.75,.75]-[-1.75,1]-[-1,1.75]-[-.5,1.75]

' Distinguishes return value in GetPart function
Enum Points
    X
    Y
End Enum

' Ordered pair
Type PointAPI
    X As Double
    Y As Double
End Type

'  A circle's circumference devided by it's diameter
Const PI = 3.14159265358979

' Example: "-[5,0]-[-2,-2]-[0,0]-[2,-2]-[5,0]"

' Distinguish how to link two ordered pairs
Const LinkChar = "-"
Const BreakChar = ":"
Const SepChar = ","

' Distinguish the beginning and the ending of an ordered pair
Const BegPos = "["
Const EndPos = "]"

' Variables of the DrawArrow proceedure
Public ArrowStr As String ' String containing drawing instructions
Public ObjPos As PointAPI ' Defined position of [0,0]
Public ObjMult As Double ' Zoom power
Public Rotate As Integer
Public LinkObject As Object ' Object to print arrow on
Public Disabled As Boolean

' Graphs a shape with custom form, width, position and angle
Public Sub DrawArrow()
Dim i As Integer, SelChar As String, j As Integer
Dim Pos1 As PointAPI, Pos2 As PointAPI
Dim PosArg As String, SepChar As String
    On Error GoTo DrawError
    SepChar = BreakChar
    ' Go through each character in ArrowStr
    For i = 1 To Len(ArrowStr)
        ' Return the next character in the loop
        SelChar = Mid$(ArrowStr, i, 1)
        Select Case SelChar
            Case LinkChar ' Don't skip this line
                SepChar = LinkChar
            Case BreakChar ' Skip this line
                SepChar = BreakChar
            Case BegPos ' Entering an ordered pair
                ' Get the ordered pair
                PosArg = Mid$(ArrowStr, i + 1, InStr(i, ArrowStr, EndPos) - i - 1)
                ' Break the ordered pair into the Pos2 variable
                Pos2 = GetPoint(GetPart(X, PosArg), GetPart(Y, PosArg))
                'If Rotate = 0 Then Stop
                If Rotate <> 0 Then
                Dim PDist As Double, Angle As Double, XMult As Integer
                    ' Distance from Pos2 to ObjPos
                    PDist = Sqr(Pos2.X ^ 2 + Pos2.Y ^ 2)
                    Angle = 0
                    ' Find the angle of the current line, add the rotation amount
                    If PDist <> 0 Then Angle = InverseSine(Pos2.X / PDist) - Rotate * GetSign(Pos2.Y)
                    ' Find the new values for Pos2 with changed angle
                    Pos2.X = Sine(Angle) * PDist
                    Pos2.Y = Cosine(Angle) * PDist * GetSign(Pos2.Y)
                End If
                
                ' If line is linked, draw line from Pos1 to Pos2
                If SepChar = LinkChar And Not Disabled Then
                    LinkObject.Line (-Pos1.X * ObjMult + ObjPos.X, Pos1.Y * ObjMult + ObjPos.Y)-(-Pos2.X * ObjMult + ObjPos.X, Pos2.Y * ObjMult + ObjPos.Y)
                End If
                Pos1 = Pos2
        End Select
    Next i
    Exit Sub
    
DrawError:
    If Err.Number = 91 Then
        MsgBox "Runtime error #91:" & vbCrLf & vbCrLf & "The LinkObject variable is not set.", vbCritical, "Class Arrow"
    End If
    Exit Sub
End Sub

' Find if a number is positive or negative
Function GetSign(Number As Double) As Double
    ' Without this line, evident error
    If Number = 0 Then GetSign = 1: Exit Function
    ' Return positive or negative 1
    GetSign = Number / Abs(Number)
End Function

' Disects a string of an ordered pair (i.e. "3,6") into two parts
Function GetPart(Part As Points, PosStr As String) As String
Dim i As Integer
    ' Find the comma in PosStr
    i = InStr(1, PosStr, SepChar)
    Select Case Part
        Case X
            ' Get all before the comma
            GetPart = Left(PosStr, i - 1)
        Case Y
            ' Get all after the comma
            GetPart = Right(PosStr, Len(PosStr) - i)
    End Select
End Function

' Find the cosine of i in degrees
Function Cosine(ByVal i As Double) As Double
    Cosine = Cos(i * (PI / 180))
End Function

' Find the sine of i in degrees
Function Sine(ByVal i As Double) As Double
    Sine = Sin(i * (PI / 180))
End Function

' Find the inverse sine of i in degrees
Function InverseSine(ByVal i As Double) As Double
    On Error GoTo MathErr
    InverseSine = Atn(i / Sqr(-i * i + 1)) * (180 / PI)
    Exit Function
MathErr:
    If i = 1 Then InverseSine = 90
    If i = -1 Then InverseSine = -90
End Function

' Function used only once, sets values for a PointAPI variable
Function GetPoint(X As Double, Y As Double) As PointAPI
    GetPoint.X = X
    GetPoint.Y = Y
End Function

