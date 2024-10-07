VERSION 5.00
Begin VB.Form frmTest 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Arrow Control Test"
   ClientHeight    =   4380
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmTest.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   292
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cboImage 
      Height          =   315
      Left            =   75
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   3975
      Width           =   4515
   End
   Begin VB.CommandButton cmdAccept 
      Cancel          =   -1  'True
      Caption         =   "&Accept"
      Height          =   390
      Left            =   3675
      TabIndex        =   2
      Top             =   3525
      Width           =   840
   End
   Begin VB.TextBox txtArrowString 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      HideSelection   =   0   'False
      Left            =   75
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   3525
      Width           =   3615
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "&Stop"
      Height          =   315
      Left            =   3750
      TabIndex        =   0
      Top             =   150
      Width           =   840
   End
   Begin VB.Label lblDescription 
      BackStyle       =   0  'Transparent
      Caption         =   "Shape Properties:"
      Height          =   195
      Left            =   75
      TabIndex        =   3
      Top             =   3300
      Width           =   1260
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TestPoint As New csArrow
Dim StopArrow As Boolean

Private Sub cboImage_Click()
    txtArrowString.Locked = Not cboImage.ListIndex = 7
    cmdAccept.Enabled = Not txtArrowString.Locked
    Select Case cboImage.ListIndex
        Case 0
            TestPoint.ArrowString = "[0,4]-[-2,-3]-[0,-2]-[2,-3]-[0,4]"
        Case 1
            TestPoint.ArrowString = "[0,3]-[3,0]-[1,0]-[1,-4]-[-1,-4]-[-1,0]-[-3,0]-[0,3]"
        Case 2
            TestPoint.ArrowString = "[0,2]-[2,-2]-[-2,-2]-[0,2]:[-4,4]-[4,4]-[0,-4]-[-4,4]"
        Case 3
            TestPoint.ArrowString = "[0,8]-[.25,7.75]-[1,1]-[1,0]-[.75,-.75]-[0,-1]-[-.75,-.75]-[-1,0]-[-1,1]-[-.25,7.75]-[0,8]"
        Case 4
            TestPoint.ArrowString = "[0,10]-[2,7]-[2,5]-[0,4]-[-2,5]-[-2,7]-[0,10]:[0,17]-[3,13]-[4,0]-[3,-5]-[3,-12]-[-3,-12]-[-3,-5]-[-4,0]-[-3,13]-[0,17]:[0,-5]-[0,-10]:[4,0]-[10,-5]-[3,-5]:[-4,0]-[-10,-5]-[-3,-5]:[3,-7]-[7,-12]-[3,-12]:[-3,-7]-[-7,-12]-[-3,-12]"
        Case 5
            TestPoint.ArrowString = "[-.25,10]-[.25,10]-[1.25,9]-[1.5,8]-[1.5,-3]-[.5,-4]-[-.5,-4]-[-1.5,-3]-[-1.5,8]-[-1.25,9]-[-.25,10]:[-1.5,6.5]-1.5,6.5]:[-1.5,6]-[1.5,6]:[-1.5,6.5]-[1.5,6.5]:[.5,3]-[-.25,3]-[-.5,3.25]-[-.5,3.75]-[-.25,4]-[.5,4]:[0,4]-[0,3]:[-.5,2.5]-[-.5,2]-[-.25,1.75]-[0,2]-[0,2.25]-[.25,2.5]-[.5,2.25]-[.5,1.75]:[-.5,1.25]-[.25,1.25]-[.5,1]-[.5,.5]-[.25,.25]-[-.5,.25]:[-1.5,-1]-[-3.5,-1.75]-[-3.5,-3]-[-1.5,-3]:[1.5,-1]-[3.5,-1.75]-[3.5,-3]-[1.5,-3]:[-1.5,-1.75]-[1.5,-1.75]:[-1.5,-2.25]-[1.5,-2.25]:[0,-1]-[0,-3]:[-.5,-4]-[-1,-4.5]-[1,-4.5]-[.5,-4]"
        Case 6
            TestPoint.ArrowString = "[-.25,3]-[.25,3]-[.5,2.75]-[.5,-.25]-[.25,-.5]-[-.25,-.5]-[-.5,-.25]-[-.5,2.75]-[-.25,3]-[-.25,3]:[.5,1.75]-[1,1.75]-[1.75,1]-[1.75,.75]-[2,.5]-[2,-.5]-[1.75,-.75]-[1.75,-1]-[1,-1.75]-[.5,-1.75]-[-.5,-1.75]-[-1,-1.75]-[-1.75,-1]-[-1.75,-.75]-[-2,-.5]-[-2,.5]-[-1.75,.75]-[-1.75,1]-[-1,1.75]-[-.5,1.75]"
    End Select
    txtArrowString = TestPoint.ArrowString
End Sub

Private Sub cmdStop_Click()
    StopArrow = True
End Sub

Private Sub cmdAccept_Click()
    TestPoint.ArrowString = txtArrowString
End Sub

Private Sub Form_Click()
Dim ZChange As Currency, RChange As Integer
    ZChange = 0.25
    RChange = 1
    StopArrow = False
    Do
        DoEvents
        If TestPoint.arrowzoom >= 30 Then
            TestPoint.arrowzoom = 30
            ZChange = -Abs(ZChange)
        ElseIf TestPoint.arrowzoom <= 1 Then
            TestPoint.arrowzoom = 1
            ZChange = Abs(ZChange)
        End If
        If RChange >= 360 Then RChange = RChange - 720
        TestPoint.SetArrowPos 100, 100
        Cls
        TestPoint.Refresh
        PSet (100, 100), 255
        TestPoint.arrowzoom = TestPoint.arrowzoom + ZChange
        TestPoint.ArrowRotate = ((TestPoint.ArrowRotate + RChange) Mod 360)
        If StopArrow Then
            StopArrow = False
            Exit Sub
        End If
    Loop
End Sub

Private Sub Form_Load()
    Set TestPoint.ArrowContainer = Me
    txtArrowString = TestPoint.ArrowString
    With cboImage
        .AddItem "Normal Arrow"
        .AddItem "Pointer"
        .AddItem "Odd Design"
        .AddItem "Gauge Needle"
        .AddItem "Airplane"
        .AddItem "Nuclear Missile"
        .AddItem "Something"
        .AddItem "Custom"
    End With
End Sub

Private Sub lblDescription_Click()
    TestPoint.ArrowDisabled = Not TestPoint.ArrowDisabled
End Sub

