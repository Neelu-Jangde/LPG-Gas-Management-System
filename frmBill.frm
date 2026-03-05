VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmBill 
   BackColor       =   &H00FFFF80&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "LPG Gas Managment System - Bill Generation"
   ClientHeight    =   7890
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14655
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7890
   ScaleWidth      =   14655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSDataGridLib.DataGrid dgBill 
      Bindings        =   "frmBill.frx":0000
      Height          =   7575
      Left            =   7320
      TabIndex        =   20
      Top             =   0
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   13361
      _Version        =   393216
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
         DataField       =   "BillID"
         Caption         =   "BillID"
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
      BeginProperty Column02 
         DataField       =   "BookingID"
         Caption         =   "BookingID"
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
         DataField       =   "BillDate"
         Caption         =   "BillDate"
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
         DataField       =   "Amount"
         Caption         =   "Amount"
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
         DataField       =   "PaymentStatus"
         Caption         =   "PaymentStatus"
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
            ColumnWidth     =   1094.74
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1094.74
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1094.74
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   2085.166
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc adoBill 
      Height          =   330
      Left            =   0
      Top             =   7320
      Width           =   7215
      _ExtentX        =   12726
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
      RecordSource    =   "tblBill"
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
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   6720
      Width           =   1335
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Clear"
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
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   6720
      Width           =   1335
   End
   Begin VB.CommandButton cmdSearch 
      BackColor       =   &H0080C0FF&
      Caption         =   "Search"
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
      Left            =   600
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
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   6120
      Width           =   1335
   End
   Begin VB.CommandButton cmdUpdate 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Update"
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
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6120
      Width           =   1335
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H0000FF00&
      Caption         =   "save"
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
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   6120
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Bill Details:"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5055
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   6975
      Begin VB.TextBox txtSearch 
         Height          =   375
         Left            =   3600
         TabIndex        =   21
         Top             =   4560
         Width           =   1095
      End
      Begin VB.TextBox txtPayStatus 
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
         Left            =   2880
         TabIndex        =   13
         Top             =   4080
         Width           =   3255
      End
      Begin VB.TextBox txtAmount 
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
         Left            =   2880
         TabIndex        =   12
         Top             =   3480
         Width           =   3255
      End
      Begin VB.TextBox txtBillDate 
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
         Left            =   2880
         TabIndex        =   11
         Top             =   2640
         Width           =   3255
      End
      Begin VB.TextBox txtBookingID 
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
         Left            =   2880
         TabIndex        =   10
         Top             =   1920
         Width           =   3255
      End
      Begin VB.TextBox txtCustID 
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
         Left            =   2880
         TabIndex        =   9
         Top             =   1200
         Width           =   3255
      End
      Begin VB.TextBox txtBillID 
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
         Height          =   375
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   480
         Width           =   3255
      End
      Begin VB.Label Label7 
         Caption         =   "Search By ID :"
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         TabIndex        =   22
         Top             =   4560
         Width           =   1815
      End
      Begin VB.Label Label6 
         Caption         =   "Payment Status:"
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
         Left            =   360
         TabIndex        =   7
         Top             =   4200
         Width           =   1815
      End
      Begin VB.Label Label5 
         Caption         =   "Amount:"
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   6
         Top             =   3600
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Bill Date:"
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   5
         Top             =   2760
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "Booking ID:"
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   4
         Top             =   1920
         Width           =   1815
      End
      Begin VB.Label Label2 
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
         Height          =   495
         Left            =   360
         TabIndex        =   3
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Bill ID:"
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
         Left            =   360
         TabIndex        =   2
         Top             =   480
         Width           =   1815
      End
   End
   Begin VB.Label lblHeader 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "Bill Generation"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   24
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
      Width           =   7095
   End
End
Attribute VB_Name = "frmBill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    adoBill.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
        "Data Source=" & App.Path & "\LPGGasSystem.accdb"
    adoBill.RecordSource = "tblBill"
    adoBill.Refresh
    txtBillDate.Text = Format(Now, "DD/MM/YYYY")
End Sub

Private Sub cmdSave_Click()
    If txtCustID.Text = "" Or txtAmount.Text = "" Then
        MsgBox "Please fill all fields!", vbExclamation, "Warning"
        Exit Sub
    End If
    adoBill.Recordset.AddNew
    adoBill.Recordset("CustomerID") = txtCustID.Text
    adoBill.Recordset("BookingID") = txtBookingID.Text
    adoBill.Recordset("BillDate") = txtBillDate.Text
    adoBill.Recordset("Amount") = txtAmount.Text
    adoBill.Recordset("PaymentStatus") = txtPayStatus.Text
    adoBill.Recordset.Update
    MsgBox "Bill Saved Successfully!", vbInformation, "Success"
    adoBill.Refresh
    Call cmdClear_Click
End Sub

Private Sub cmdSearch_Click()
    If txtSearch.Text = "" Then
        MsgBox "Please enter Bill ID!", vbExclamation, "Warning"
        Exit Sub
    End If
    Dim rs As New ADODB.Recordset
    rs.Open "SELECT * FROM [tblBill] WHERE BillID = " & txtSearch.Text, adoBill.ConnectionString
    If rs.EOF Then
        MsgBox "Record Not Found!", vbExclamation, "Not Found"
    Else
        txtBillID.Text = rs("BillID")
        txtCustID.Text = rs("CustomerID")
        txtBookingID.Text = rs("BookingID")
        txtBillDate.Text = rs("BillDate")
        txtAmount.Text = rs("Amount")
        txtPayStatus.Text = rs("PaymentStatus")
    End If
    rs.Close
End Sub

Private Sub cmdUpdate_Click()
    If txtBillID.Text = "" Then
        MsgBox "Please Search First!", vbExclamation, "Warning"
        Exit Sub
    End If
    Dim rs As New ADODB.Recordset
    rs.Open "SELECT * FROM [tblBill] WHERE BillID = " & txtBillID.Text, adoBill.ConnectionString, 1, 3
    If rs.EOF Then
        MsgBox "Record Not Found!", vbExclamation, "Not Found"
    Else
        rs("CustomerID") = txtCustID.Text
        rs("BookingID") = txtBookingID.Text
        rs("BillDate") = txtBillDate.Text
        rs("Amount") = txtAmount.Text
        rs("PaymentStatus") = txtPayStatus.Text
        rs.Update
        MsgBox "Bill Updated Successfully!", vbInformation, "Success"
        adoBill.Refresh
    End If
    rs.Close
End Sub

Private Sub cmdDelete_Click()
    If txtBillID.Text = "" Then
        MsgBox "Please Search First!", vbExclamation, "Warning"
        Exit Sub
    End If
    If MsgBox("Are you sure?", vbYesNo + vbQuestion, "Delete") = vbYes Then
        Dim rs As New ADODB.Recordset
        rs.Open "SELECT * FROM [tblBill] WHERE BillID = " & txtBillID.Text, adoBill.ConnectionString, 1, 3
        If Not rs.EOF Then
            rs.Delete
            MsgBox "Bill Deleted Successfully!", vbInformation, "Success"
            adoBill.Refresh
            Call cmdClear_Click
        End If
        rs.Close
    End If
End Sub

Private Sub cmdClear_Click()
    txtBillID.Text = ""
    txtCustID.Text = ""
    txtBookingID.Text = ""
    txtBillDate.Text = Format(Now, "DD/MM/YYYY")
    txtAmount.Text = ""
    txtPayStatus.Text = ""
    txtSearch.Text = ""
    txtCustID.SetFocus
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub
