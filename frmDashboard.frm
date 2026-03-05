VERSION 5.00
Begin VB.Form frmDashboard 
   BackColor       =   &H0000FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LPG Gas Managment System - Dashboard"
   ClientHeight    =   7800
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14400
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7800
   ScaleWidth      =   14400
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrDate 
      Interval        =   1000
      Left            =   12840
      Top             =   6840
   End
   Begin VB.CommandButton cmdLogout 
      BackColor       =   &H000000FF&
      Caption         =   "Log out"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4320
      Width           =   2655
   End
   Begin VB.CommandButton cmdStock 
      BackColor       =   &H00C0C000&
      Caption         =   "Stock  Managment"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1800
      Left            =   10800
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1680
      Width           =   3000
   End
   Begin VB.CommandButton cmdBill 
      BackColor       =   &H00C0C000&
      Caption         =   "Bill - Generation"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1800
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1680
      Width           =   3000
   End
   Begin VB.CommandButton cmdBooking 
      BackColor       =   &H00C0C000&
      Caption         =   "Gas -  Booking"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1680
      Width           =   3000
   End
   Begin VB.CommandButton cmdCustomer 
      BackColor       =   &H00C0C000&
      Caption         =   "Customer Registration"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1800
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1680
      Width           =   3000
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   7920
      TabIndex        =   7
      Top             =   6600
      Width           =   2895
   End
   Begin VB.Label lblWelcome 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      Caption         =   "Welcome, Admin !"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   3960
      TabIndex        =   6
      Top             =   6720
      Width           =   3015
   End
   Begin VB.Label lblHeader 
      Alignment       =   2  'Center
      BackColor       =   &H00404000&
      Caption         =   "LPG Gas Managment System"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14415
   End
End
Attribute VB_Name = "frmDashboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBill_Click()
frmBill.Show 1
End Sub

Private Sub cmdBooking_Click()
 frmBooking.Show 1
End Sub

Private Sub cmdCustomer_Click()
frmCustomer.Show 1
End Sub

Private Sub cmdLogout_Click()
If MsgBox("Are you sure you want to logout?", vbYesNo + vbQuestion, "Logout") = vbYes Then
        frmLogin.Show
        Unload Me
    End If
End Sub

Private Sub cmdStock_Click()
frmStock.Show 1
End Sub

Private Sub Form_Load()
lblDate.Caption = "Today : " & Format(Now, "DD/MM/YYYY hh:mm AM/PM")
End Sub

Private Sub tmrDate_Timer()
lblDate.Caption = "Today : " & Format(Now, "DD/MM/YYYY hh:mm AM/PM")
End Sub
