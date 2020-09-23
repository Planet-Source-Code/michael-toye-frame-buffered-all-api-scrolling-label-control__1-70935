VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   2295
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   3555
   LinkTopic       =   "Form1"
   ScaleHeight     =   2295
   ScaleWidth      =   3555
   StartUpPosition =   3  'Windows Default
   Begin Project1.ctlScroller ctlScroller1 
      Height          =   255
      Index           =   0
      Left            =   420
      TabIndex        =   0
      Top             =   120
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   450
      Caption         =   "ctlScroller"
      Style           =   1
      Face            =   "Tahoma"
      fntSize         =   10
      BackColor       =   16777215
   End
   Begin Project1.ctlScroller ctlScroller1 
      Height          =   255
      Index           =   1
      Left            =   420
      TabIndex        =   1
      Top             =   420
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   450
      Caption         =   "ctlScroller"
      Style           =   1
      Face            =   "Tahoma"
      fntSize         =   10
      BackColor       =   16777215
   End
   Begin Project1.ctlScroller ctlScroller1 
      Height          =   255
      Index           =   2
      Left            =   420
      TabIndex        =   2
      Top             =   720
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   450
      Caption         =   "ctlScroller"
      Style           =   1
      Face            =   "Tahoma"
      fntSize         =   10
      BackColor       =   16777215
   End
   Begin Project1.ctlScroller ctlScroller1 
      Height          =   255
      Index           =   3
      Left            =   420
      TabIndex        =   3
      Top             =   1020
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   450
      Caption         =   "ctlScroller"
      Style           =   1
      Face            =   "Tahoma"
      fntSize         =   10
      BackColor       =   16777215
   End
   Begin Project1.ctlScroller ctlScroller1 
      Height          =   255
      Index           =   4
      Left            =   420
      TabIndex        =   4
      Top             =   1320
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   450
      Caption         =   "ctlScroller"
      Style           =   1
      Face            =   "Tahoma"
      fntSize         =   10
      BackColor       =   16777215
   End
   Begin Project1.ctlScroller ctlScroller1 
      Height          =   435
      Index           =   5
      Left            =   420
      TabIndex        =   5
      Top             =   1620
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   767
      Caption         =   "ctlScroller"
      Style           =   1
      Face            =   "Tahoma"
      fntSize         =   10
      BackColor       =   16777215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim n&
For n = 0 To ctlScroller1.Count - 1
    ctlScroller1(n).Caption = "ABCDEFGHIJ abcdefghijkl 0123456789"
Next
ctlScroller1(0).BackStyle = 0
ctlScroller1(0).BackColor = Me.BackColor

ctlScroller1(1).BackStyle = 0
ctlScroller1(1).BackColor = vbWhite

ctlScroller1(2).BackStyle = 1
ctlScroller1(3).BackStyle = 2
ctlScroller1(4).BackStyle = 3
ctlScroller1(5).BackStyle = 2
ctlScroller1(5).FontSize = 16
End Sub
