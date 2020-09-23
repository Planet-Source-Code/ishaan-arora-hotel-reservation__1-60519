VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7110
   ClientLeft      =   1155
   ClientTop       =   -420
   ClientWidth     =   11550
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7.04100e5
   ScaleLeft       =   444
   ScaleMode       =   0  'User
   ScaleTop        =   44
   ScaleWidth      =   1.15505e5
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   6930
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   10320
      Begin VB.PictureBox Picture2 
         Height          =   855
         Left            =   600
         Picture         =   "frmSplash.frx":000C
         ScaleHeight     =   795
         ScaleWidth      =   1635
         TabIndex        =   9
         Top             =   4080
         Width           =   1695
      End
      Begin VB.PictureBox Picture1 
         Height          =   1215
         Left            =   360
         Picture         =   "frmSplash.frx":21B2
         ScaleHeight     =   1155
         ScaleWidth      =   1395
         TabIndex        =   8
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "click anywhere on the form or press enter to continue"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1335
         Left            =   600
         TabIndex        =   12
         Top             =   5520
         Width           =   6615
      End
      Begin VB.Label Label3 
         Caption         =   "FLY TO GOA AND STAY IN OUR HOTEL"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   735
         Left            =   2040
         TabIndex        =   11
         Top             =   2280
         Width           =   4455
      End
      Begin VB.Label Label2 
         Caption         =   "CALL NOW AND RESERVE. PH.NO- 2410657"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2760
         TabIndex        =   10
         Top             =   4200
         Width           =   3855
      End
      Begin VB.Label Label1 
         Caption         =   "THE SPARKLES"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   2160
         TabIndex        =   7
         Top             =   720
         Width           =   3015
      End
      Begin VB.Label lblCopyright 
         Caption         =   "WE ACCEPT VISA ELECTRON AND OTHER CREDIT CARDS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   3600
         Width           =   6615
      End
      Begin VB.Label lblWarning 
         Caption         =   "COOPERATE WITH US FOR BEST RESULTS"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   435
         Left            =   600
         TabIndex        =   1
         Top             =   5040
         Width           =   5415
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "SINCE 1985"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   5400
         TabIndex        =   3
         Top             =   2040
         Width           =   1350
      End
      Begin VB.Label lblPlatform 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "SERVICE AT ITS BEST"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   360
         Left            =   3600
         TabIndex        =   4
         Top             =   1560
         Width           =   3360
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         Caption         =   "FIVE STAR MEGA DELUXE HOTEL"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1560
         TabIndex        =   6
         Top             =   1200
         Width           =   3465
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         Caption         =   "ISHAAN ARORA INTRODUCES"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   720
         TabIndex        =   5
         Top             =   240
         Width           =   5040
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
MDIForm1.Hide
Form3.Show
frmSplash.Hide
End Sub

Private Sub Frame1_Click()
MDIForm1.Hide
Form3.Show
frmSplash.Hide
End Sub

