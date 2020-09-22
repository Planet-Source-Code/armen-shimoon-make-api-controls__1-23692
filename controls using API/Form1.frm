VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "API Controls"
   ClientHeight    =   3240
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9825
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3240
   ScaleWidth      =   9825
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "Make Static"
      Height          =   2355
      Left            =   6390
      TabIndex        =   26
      Top             =   720
      Width           =   3075
      Begin VB.CommandButton Command3 
         Caption         =   "Make Static"
         Height          =   465
         Left            =   900
         TabIndex        =   37
         Top             =   1710
         Width           =   1545
      End
      Begin VB.TextBox lblw1 
         Height          =   375
         Left            =   810
         TabIndex        =   36
         Top             =   1170
         Width           =   735
      End
      Begin VB.TextBox lblh1 
         Height          =   375
         Left            =   1980
         TabIndex        =   35
         Top             =   1170
         Width           =   735
      End
      Begin VB.TextBox lbly1 
         Height          =   375
         Left            =   1980
         TabIndex        =   34
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox lblx1 
         Height          =   375
         Left            =   810
         TabIndex        =   33
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox lblCaption 
         Height          =   375
         Left            =   900
         TabIndex        =   28
         Top             =   270
         Width           =   1815
      End
      Begin VB.Label Label16 
         Caption         =   "h1"
         Height          =   285
         Left            =   1710
         TabIndex        =   32
         Top             =   1260
         Width           =   285
      End
      Begin VB.Label Label15 
         Caption         =   "w1"
         Height          =   375
         Left            =   540
         TabIndex        =   31
         Top             =   1260
         Width           =   375
      End
      Begin VB.Label Label14 
         Caption         =   "y1"
         Height          =   285
         Left            =   1710
         TabIndex        =   30
         Top             =   810
         Width           =   375
      End
      Begin VB.Label Label13 
         Caption         =   "x1"
         Height          =   285
         Left            =   540
         TabIndex        =   29
         Top             =   810
         Width           =   375
      End
      Begin VB.Label Label12 
         Caption         =   "Caption"
         Height          =   285
         Left            =   270
         TabIndex        =   27
         Top             =   360
         Width           =   915
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Make Edit"
      Height          =   2355
      Left            =   3240
      TabIndex        =   14
      Top             =   720
      Width           =   3075
      Begin VB.CommandButton Command2 
         Caption         =   "Make Edit"
         Height          =   465
         Left            =   900
         TabIndex        =   25
         Top             =   1710
         Width           =   1545
      End
      Begin VB.TextBox edth1 
         Height          =   375
         Left            =   1980
         TabIndex        =   24
         Top             =   1170
         Width           =   735
      End
      Begin VB.TextBox edtw1 
         Height          =   375
         Left            =   810
         TabIndex        =   23
         Top             =   1170
         Width           =   735
      End
      Begin VB.TextBox edty1 
         Height          =   375
         Left            =   1980
         TabIndex        =   20
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox edtx1 
         Height          =   375
         Left            =   810
         TabIndex        =   18
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox edtCaption 
         Height          =   375
         Left            =   900
         TabIndex        =   16
         Top             =   270
         Width           =   1815
      End
      Begin VB.Label Label11 
         Caption         =   "h1"
         Height          =   465
         Left            =   1710
         TabIndex        =   22
         Top             =   1260
         Width           =   375
      End
      Begin VB.Label Label10 
         Caption         =   "w1"
         Height          =   285
         Left            =   450
         TabIndex        =   21
         Top             =   1260
         Width           =   375
      End
      Begin VB.Label Label9 
         Caption         =   "y1"
         Height          =   195
         Left            =   1710
         TabIndex        =   19
         Top             =   810
         Width           =   285
      End
      Begin VB.Label Label8 
         Caption         =   "x1"
         Height          =   285
         Left            =   450
         TabIndex        =   17
         Top             =   810
         Width           =   375
      End
      Begin VB.Label Label7 
         Caption         =   "Caption"
         Height          =   285
         Left            =   270
         TabIndex        =   15
         Top             =   360
         Width           =   645
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Make Button"
      Height          =   2355
      Left            =   90
      TabIndex        =   2
      Top             =   720
      Width           =   3075
      Begin VB.CommandButton Command1 
         Caption         =   "Make Button"
         Height          =   465
         Left            =   900
         TabIndex        =   13
         Top             =   1710
         Width           =   1545
      End
      Begin VB.TextBox txth1 
         Height          =   375
         Left            =   1980
         TabIndex        =   12
         Top             =   1170
         Width           =   735
      End
      Begin VB.TextBox txtw1 
         Height          =   375
         Left            =   810
         TabIndex        =   10
         Top             =   1170
         Width           =   735
      End
      Begin VB.TextBox txty1 
         Height          =   375
         Left            =   1980
         TabIndex        =   8
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox txtx1 
         Height          =   375
         Left            =   810
         TabIndex        =   6
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox txtCaption 
         Height          =   375
         Left            =   900
         TabIndex        =   4
         Top             =   270
         Width           =   1815
      End
      Begin VB.Label Label6 
         Caption         =   "h1"
         Height          =   285
         Left            =   1710
         TabIndex        =   11
         Top             =   1260
         Width           =   285
      End
      Begin VB.Label Label5 
         Caption         =   "w1"
         Height          =   285
         Left            =   540
         TabIndex        =   9
         Top             =   1260
         Width           =   285
      End
      Begin VB.Label Label4 
         Caption         =   "y1"
         Height          =   285
         Left            =   1710
         TabIndex        =   7
         Top             =   810
         Width           =   555
      End
      Begin VB.Label Label3 
         Caption         =   "x1"
         Height          =   285
         Left            =   540
         TabIndex        =   5
         Top             =   810
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "Caption"
         Height          =   195
         Left            =   270
         TabIndex        =   3
         Top             =   360
         Width           =   645
      End
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1350
      TabIndex        =   0
      Top             =   90
      Width           =   2085
   End
   Begin VB.Label Label17 
      Caption         =   "Made by Armen Shimoon.  You can use this code in your programs as long as you give me credit. 2001©"
      Height          =   465
      Left            =   4860
      TabIndex        =   38
      Top             =   90
      Width           =   4515
   End
   Begin VB.Label Label1 
      Caption         =   "Window name"
      Height          =   285
      Left            =   180
      TabIndex        =   1
      Top             =   180
      Width           =   1185
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Call MakeButton(Text1, txtx1, txty1, txtw1, txth1, txtCaption)
End Sub

Private Sub Command2_Click()
Call MakeEdit(Text1, edtx1, edty1, edtw1, edth1, edtCaption)
End Sub

Private Sub Command3_Click()
Call MakeStatic(Text1, lblx1, lbly1, lblw1, lblh1, lblCaption)
End Sub
