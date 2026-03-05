VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmCustomer 
   BackColor       =   &H00FFFF80&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "LPG Gas Managment System - Customer Registration."
   ClientHeight    =   7725
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14475
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7725
   ScaleWidth      =   14475
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSDataGridLib.DataGrid dgCustomer 
      Bindings        =   "frmCustomer.frx":0000
      Height          =   7695
      Left            =   6480
      TabIndex        =   20
      Top             =   0
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   13573
      _Version        =   393216
      BackColor       =   16777215
      HeadLines       =   1
      RowHeight       =   17
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Rockwell"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   6
      BeginProperty Column00 
         DataField       =   "CustomerID"
         Caption         =   "CustomerID"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "CustomerName"
         Caption         =   "CustomerName"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "Address"
         Caption         =   "Address"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "PhoneNumber"
         Caption         =   "PhoneNumber"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "ConnectionType"
         Caption         =   "ConnectionType"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "ConnectionDate"
         Caption         =   "ConnectionDate"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   1110.047
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1814.74
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   2085.166
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc adoCustomer 
      Height          =   330
      Left            =   0
      Top             =   7440
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\durgesh computer\Documents\LPGGasSystem.accdb"
      OLEDBString     =   "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\durgesh computer\Documents\LPGGasSystem.accdb"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "tblCustomer"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00FF80FF&
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   6720
      Width           =   1455
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00808080&
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   6720
      Width           =   1455
   End
   Begin VB.CommandButton cmdSearch 
      BackColor       =   &H000080FF&
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   6720
      Width           =   1455
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H000000FF&
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   6120
      Width           =   1455
   End
   Begin VB.CommandButton cmdUpdate 
      BackColor       =   &H00FF8080&
      Caption         =   "Update"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6120
      Width           =   1455
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H0000FF00&
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   6120
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Customer Details--"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4935
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   6255
      Begin VB.TextBox txtSearch 
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3360
         TabIndex        =   21
         Top             =   4440
         Width           =   975
      End
      Begin VB.TextBox txtConnDate 
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3000
         TabIndex        =   13
         Top             =   3960
         Width           =   2895
      End
      Begin VB.TextBox txtConnType 
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3000
         TabIndex        =   12
         Top             =   3240
         Width           =   2895
      End
      Begin VB.TextBox txtPhone 
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3000
         TabIndex        =   11
         Top             =   2400
         Width           =   2895
      End
      Begin VB.TextBox txtAddress 
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3000
         TabIndex        =   10
         Top             =   1800
         Width           =   2895
      End
      Begin VB.TextBox txtCustName 
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3000
         TabIndex        =   9
         Top             =   1200
         Width           =   2895
      End
      Begin VB.TextBox txtCustID 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   600
         Width           =   2895
      End
      Begin VB.Label Label7 
         Caption         =   "Search By ID :"
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         TabIndex        =   22
         Top             =   4440
         Width           =   1815
      End
      Begin VB.Label Label6 
         Caption         =   "Connection Date:"
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   7
         Top             =   3960
         Width           =   1815
      End
      Begin VB.Label Label5 
         Caption         =   "Connection Type:"
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   6
         Top             =   3240
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Phone Number:"
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   5
         Top             =   2520
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "Address:"
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   4
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Custoner Name:"
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   3
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Customer ID:"
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   2
         Top             =   720
         Width           =   1455
      End
   End
   Begin VB.Label lblHeader 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "Customer Registration"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6375
   End
End
Attribute VB_Name = "frmCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    adoCustomer.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
        "Data Source=" & App.Path & "\LPGGasSystem.accdb"
    adoCustomer.RecordSource = "tblCustomer"
    adoCustomer.Refresh
End Sub

Private Sub cmdSave_Click()
    If txtCustName.Text = "" Or txtPhone.Text = "" Then
        MsgBox "Please fill all fields!", vbExclamation, "Warning"
        Exit Sub
    End If
    adoCustomer.Recordset.AddNew
    adoCustomer.Recordset("CustomerName") = txtCustName.Text
    adoCustomer.Recordset("Address") = txtAddress.Text
    adoCustomer.Recordset("PhoneNumber") = txtPhone.Text
    adoCustomer.Recordset("ConnectionType") = txtConnType.Text
    adoCustomer.Recordset("ConnectionDate") = txtConnDate.Text
    adoCustomer.Recordset.Update
    MsgBox "Customer Saved Successfully!", vbInformation, "Success"
    adoCustomer.Refresh
    Call cmdClear_Click
End Sub

Private Sub cmdSearch_Click()
    If txtSearch.Text = "" Then
        MsgBox "Please enter Customer ID!", vbExclamation, "Warning"
        Exit Sub
    End If
    Dim rs As New ADODB.Recordset
    rs.Open "SELECT * FROM [tblCustomer] WHERE CustomerID = " & txtSearch.Text, adoCustomer.ConnectionString
    If rs.EOF Then
        MsgBox "Record Not Found!", vbExclamation, "Not Found"
    Else
        txtCustID.Text = rs("CustomerID")
        txtCustName.Text = rs("CustomerName")
        txtAddress.Text = rs("Address")
        txtPhone.Text = rs("PhoneNumber")
        txtConnType.Text = rs("ConnectionType")
        txtConnDate.Text = rs("ConnectionDate")
    End If
    rs.Close
End Sub

Private Sub cmdUpdate_Click()
    If txtCustID.Text = "" Then
        MsgBox "Please Search First!", vbExclamation, "Warning"
        Exit Sub
    End If
    Dim rs As New ADODB.Recordset
    rs.Open "SELECT * FROM [tblCustomer] WHERE CustomerID = " & txtCustID.Text, adoCustomer.ConnectionString, 1, 3
    If rs.EOF Then
        MsgBox "Record Not Found!", vbExclamation, "Not Found"
    Else
        rs("CustomerName") = txtCustName.Text
        rs("Address") = txtAddress.Text
        rs("PhoneNumber") = txtPhone.Text
        rs("ConnectionType") = txtConnType.Text
        rs("ConnectionDate") = txtConnDate.Text
        rs.Update
        MsgBox "Customer Updated Successfully!", vbInformation, "Success"
        adoCustomer.Refresh
    End If
    rs.Close
End Sub

Private Sub cmdDelete_Click()
    If txtCustID.Text = "" Then
        MsgBox "Please Search First!", vbExclamation, "Warning"
        Exit Sub
    End If
    If MsgBox("Are you sure?", vbYesNo + vbQuestion, "Delete") = vbYes Then
        Dim rs As New ADODB.Recordset
        rs.Open "SELECT * FROM [tblCustomer] WHERE CustomerID = " & txtCustID.Text, adoCustomer.ConnectionString, 1, 3
        If Not rs.EOF Then
            rs.Delete
            MsgBox "Customer Deleted Successfully!", vbInformation, "Success"
            adoCustomer.Refresh
            Call cmdClear_Click
        End If
        rs.Close
    End If
End Sub

Private Sub cmdClear_Click()
    txtCustID.Text = ""
    txtCustName.Text = ""
    txtAddress.Text = ""
    txtPhone.Text = ""
    txtConnType.Text = ""
    txtConnDate.Text = ""
    txtSearch.Text = ""
    txtCustName.SetFocus
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub
