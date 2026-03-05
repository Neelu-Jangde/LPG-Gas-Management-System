VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSplash 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "LPG Gas Managment System"
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9405
   ForeColor       =   &H00E0E0E0&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   9405
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrSplash 
      Interval        =   150
      Left            =   8640
      Top             =   4080
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   2760
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   873
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label lblPercent 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0%"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   3480
      TabIndex        =   4
      Top             =   3360
      Width           =   495
   End
   Begin VB.Label lblLoading 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Loading.......Please Wait ."
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   4080
      TabIndex        =   3
      Top             =   3360
      Width           =   2415
   End
   Begin VB.Label lblTagline 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Your Trusted Gas Partner......"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   1440
      Width           =   4575
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "LPG Gas Managment System"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9375
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
ProgressBar1.Min = 0
    ProgressBar1.Max = 100
    ProgressBar1.Value = 0
End Sub

Private Sub tmrSplash_Timer()
 If ProgressBar1.Value < 100 Then
        ProgressBar1.Value = ProgressBar1.Value + 1
        lblPercent.Caption = ProgressBar1.Value & "%"
    Else
        tmrSplash.Enabled = False
        frmLogin.Show
        Unload Me
    End If
End Sub
