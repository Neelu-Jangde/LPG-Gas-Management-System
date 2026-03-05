VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "LPG Gas Managment System - Login"
   ClientHeight    =   5385
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8220
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   13.5
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   8220
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H000000FF&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton cmdLogin 
      BackColor       =   &H0000FF00&
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3360
      Width           =   1455
   End
   Begin VB.TextBox txtPassword 
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      TabIndex        =   4
      Top             =   2280
      Width           =   2895
   End
   Begin VB.TextBox txtUsername 
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   3840
      TabIndex        =   3
      Top             =   1440
      Width           =   2895
   End
   Begin VB.Label lblCopywrite 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "© 2026 LPG Gas Management System"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   4920
      Width           =   8295
   End
   Begin VB.Label lblPassword 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Password:-"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   495
      Left            =   1320
      TabIndex        =   2
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label lblUsername 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Username:-"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   495
      Left            =   1320
      TabIndex        =   1
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      Caption         =   "                                LPG Gas Managment System                               Your Trusted Gas Partner."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8175
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Change()
If MsgBox("Are you sure you want to exit?", vbYesNo + vbQuestion, "Exit") = vbYes Then
        End
    End If
End Sub

Private Sub cmdExit_Click()
If MsgBox("Are you sure you want to exit?", vbYesNo + vbQuestion, "Exit") = vbYes Then
        End
    End If
End Sub

Private Sub cmdLogin_Click()
If txtUsername.Text = "admin" And txtPassword.Text = "admin123" Then
        MsgBox "Login Successful! Welcome!", vbInformation, "Success"
        frmDashboard.Show
        Unload Me
    Else
        MsgBox "Invalid Username or Password!", vbCritical, "Error"
        txtPassword.Text = ""
        txtUsername.SetFocus
    End If
End Sub

