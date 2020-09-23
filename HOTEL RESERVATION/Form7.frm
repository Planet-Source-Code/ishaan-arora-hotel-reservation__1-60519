VERSION 5.00
Begin VB.Form Form7 
   Caption         =   "HOTEL RESERVATION MANAGEMENT SYSTEM "
   ClientHeight    =   5880
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7260
   LinkTopic       =   "Form7"
   MDIChild        =   -1  'True
   ScaleHeight     =   5880
   ScaleWidth      =   7260
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   6
      Top             =   2880
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "CHECK"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8640
      TabIndex        =   5
      Top             =   1680
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   855
      IMEMode         =   3  'DISABLE
      Left            =   4800
      PasswordChar    =   "*"
      TabIndex        =   3
      Tag             =   "ISHAAN"
      Top             =   1560
      Width           =   3495
   End
   Begin VB.PictureBox Picture1 
      Height          =   3975
      Left            =   2640
      Picture         =   "Form7.frx":0000
      ScaleHeight     =   3915
      ScaleWidth      =   7155
      TabIndex        =   2
      Top             =   2640
      Width           =   7215
   End
   Begin VB.Label Label4 
      Caption         =   "PASSWORD IS ""ISHAAN"".WRITE IN CAPITALS AND CLICK ON CKECK COMMAND BUTTON"
      Height          =   495
      Left            =   4440
      TabIndex        =   7
      Top             =   1200
      Width           =   3975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0FF&
      Caption         =   "WELCOME TO HOTEL RESERVATION              MANAGEMENT SYSTEM"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1095
      Left            =   1560
      TabIndex        =   4
      Top             =   0
      Width           =   6735
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFC0FF&
      Caption         =   " TO ENABLE THE MENU ITEMS ENTER PASSWORD"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   855
      Left            =   0
      TabIndex        =   1
      Top             =   1680
      Width           =   4455
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFC0FF&
      Caption         =   "THIS SOFTWARE IS COPYRIGHT UNDER SECTION 453-C .ANY ILLEGAL COPY OF IT WILL BE A PUNISHABLE ACT."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   975
      Left            =   480
      TabIndex        =   0
      Top             =   6600
      Width           =   7335
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'********HOTEL RESERVATION MANAGEMENT SYSTEM************
'MADE BY ===ISHAAN ARORA*************8
'CONTACT AT ARORAISHAAN666@YAHOO.COM




  'THIS IS MY FIRST REALISTIC PROGRAM .I GOT INSPIRATION
  ' FROM MANY CODERS ON THIS SITE.
  ' *****************************************************
  'THIS IS A PROGRAM FOR MAINTAINING THE RECORDS OF PEOPLE
  'COMING AND GOING IN THE HOTEL...
  
  '***********************************************
  'THIS IS A VERY FLEXIBLE PROGRAM AND HAS A VERY USER
  'FRIENDLY INTERFACE.
  'IT HAS TWO PARTS............
  '1.RECORDS OF SINGLE IN THE ROOM OF A HOTEL AND....
  '2.RECORDS OF COUPLES IN THE HOTEL
  ' TWO RUN THIS PROGRAM YOU MUST HAVE ORACLE INSTALLED ON YOUR PC
  'AND MUST HAVE THE TWO TABLES MADE WITHOUT ANT RELATION BETWEEN THEM
  '*******************************************************
  '*******************THE STRUCTURE OF TWO TABLES IS AS FOLLOWS
  '1.TABLE NAME-HOTEL
  'FIELDS AND DATA TYPES===
  
 '------------------------------- -------- ----
' NAME                                     VARCHAR2(20)
' ADDRESS                                  VARCHAR2(30)
' PHONE_N0                                  Number(10)
' DAYS_OF_STAYING                           Number(10)
 'ROOMS_ON_RENT                             Number(10)
 'Class                                     VARCHAR2(10)
 'TOTAL                                     Number(20)
 '********************************************************
 
'2.TABLE NAME==COUPLE


'FIELDS AND DATATYPES
 'NAME                                      VARCHAR2(20)
 'ADDRESS                                   VARCHAR2(30)
 'PHONE_N0                                  Number(10)
 'DAYS_OF_STAYING                           Number(10)
 'ROOMS_ON_RENT                             Number(10)
 'Class                                     VARCHAR2(10)
 'TOTAL                                     Number(20)
 
 'WHILE INSERTING THE RECORDS 'TOTAL' COLUMN MUST BE KEPT EMPTY
 ' AS THE PROGRAM HAS A COMMAND BUTTON NAMED"CALCULATE TOTAL"
 'WHICH AUTOMATICALLY CALCULATES THE TOTAL OF A CUSTOMER ACCORDING TO HIS DAYS OF STAYING
 'AND ROOMS ONRENT
 
 
 '**************************************************
'IMPORTANT====="DO NOT CHANGE THE FIELD NAME AS IT MAY
'CAUSE THE PROGRAM TO RUN CORECTLY.
'PLEASE VOTE THIS PROGRAM

Private Sub Command1_Click()
A = MsgBox("ARE YOU SURE DO YOU REALLY WANT TO EXIT", vbInformation + vbYesNo)
If A = vbYes Then
End
Else: Form7.Show
End If
End Sub

Private Sub Command2_Click()
If Text1.Text = Text1.Tag Then

A = MsgBox("RECORDS OF ALLTHE CUSTOMERS OF OUR HOTEL JUST SECONDS AWAY,ENJOY THE TOUR", vbExclamation)
MsgBox "MENU ITEMS ARE ENABLED NOW"
frmSplash.Show
MDIForm1.VIEW.Enabled = True
MDIForm1.REP.Enabled = True
MDIForm1.ADD.Enabled = True
MDIForm1.SEARCH.Enabled = True
MDIForm1.OPT = True
End If
If Text1.Text <> Text1.Tag Then
b = MsgBox("NO NO  NO! INVALID PASSWORD. YOU CANNOT CONTINUE", vbCritical, "PASSWORD VALIDATION")

  c = MsgBox("TRY ONCE MORE OR EXIT", vbRetryCancel)
  If c = vbRetry Then
  Text1.SetFocus
  Text1.Text = ""
  ElseIf c = vbCancel Then
  
  U = MsgBox("BYE,BYE", vbOKOnly)
  End
  End If
  End If
End Sub

Private Sub Form_Load()
MsgBox "IMPORTANT INSTRUCTIONS-PLEASE REFER FORM7 AND THE NOTEPAD FILE ATTACHED TO THIS PROGRAM FOR CORRECT RUNNING OF THIS PROGRAM"
MsgBox "MY EMAIL ADDRESS IS ALSO ON THE NOTEPAD FILE-PLEASE GIVE YOUR SUGGESTIONS"

End Sub
