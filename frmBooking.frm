VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmBooking 
   BackColor       =   &H00FFFF80&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "LPG Gas Managment System - Gas Booking"
   ClientHeight    =   7815
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14310
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7815
   ScaleWidth      =   14310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSDataGridLib.DataGrid dgBooking 
      Bindings        =   "frmBooking.frx":0000
      Height          =   7815
      Left            =   6960
      TabIndex        =   20
      Top             =   0
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   13785
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
         DataField       =   "BookingDate"
         Caption         =   "BookingDate"
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
         DataField       =   "CylinderType"
         Caption         =   "CylinderType"
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
         DataField       =   "Quantity"
         Caption         =   "Quantity"
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
         DataField       =   "Status"
         Caption         =   "Status"
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
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1739.906
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc adoBooking 
      Height          =   330
      Left            =   0
      Top             =   7440
      Width           =   6855
      _ExtentX        =   12091
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
      RecordSource    =   "tblBooking"
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
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   6480
      Width           =   1575
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00808080&
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
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   6480
      Width           =   1695
   End
   Begin VB.CommandButton cmdSearch 
      BackColor       =   &H000080FF&
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
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   6480
      Width           =   1695
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
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   5880
      Width           =   1575
   End
   Begin VB.CommandButton cmdUpdate 
      BackColor       =   &H00808000&
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
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   5880
      Width           =   1695
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H0080FF80&
      Caption         =   "Save"
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
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5880
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Booking Details:"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4935
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   6735
      Begin VB.TextBox txtSearch 
         Height          =   375
         Left            =   3120
         TabIndex        =   21
         Top             =   4440
         Width           =   1095
      End
      Begin VB.TextBox txtStatus 
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
         Left            =   2640
         TabIndex        =   13
         Top             =   3960
         Width           =   3015
      End
      Begin VB.TextBox txtQuantity 
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
         Left            =   2640
         TabIndex        =   12
         Top             =   3360
         Width           =   3015
      End
      Begin VB.TextBox txtCylType 
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
         Left            =   2640
         TabIndex        =   11
         Top             =   2640
         Width           =   3015
      End
      Begin VB.TextBox txtBookingDate 
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
         Left            =   2640
         TabIndex        =   10
         Top             =   1920
         Width           =   3015
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
         Left            =   2640
         TabIndex        =   9
         Top             =   1080
         Width           =   3015
      End
      Begin VB.TextBox txtBookingID 
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
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   480
         Width           =   3015
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
         Left            =   960
         TabIndex        =   22
         Top             =   4440
         Width           =   1695
      End
      Begin VB.Label Label6 
         Caption         =   "Status:"
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
         Caption         =   "Quantity:"
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
         Top             =   3360
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "Cylinder Type:"
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
         Top             =   2640
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Booking Date:"
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
         Height          =   375
         Left            =   480
         TabIndex        =   3
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label1 
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
         Height          =   375
         Left            =   480
         TabIndex        =   2
         Top             =   600
         Width           =   1575
      End
   End
   Begin VB.Label lblHeader 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      Caption         =   "Gas Booking System"
      BeginProperty Font 
         Name            =   "Rockwell"
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
      Width           =   6855
   End
End
Attribute VB_Name = "frmBooking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    adoBooking.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\durgesh computer\Documents\LPGGasSystem.accdb"
    adoBooking.RecordSource = "tblBooking"
    adoBooking.Refresh
    txtBookingDate.Text = Format(Now, "DD/MM/YYYY")
End Sub

Private Sub cmdSave_Click()
    If txtCustID.Text = "" Or txtCylType.Text = "" Then
        MsgBox "Please fill all fields!", vbExclamation, "Warning"
        Exit Sub
    End If
    adoBooking.Recordset.AddNew
    adoBooking.Recordset("CustomerID") = txtCustID.Text
    adoBooking.Recordset("BookingDate") = txtBookingDate.Text
    adoBooking.Recordset("CylinderType") = txtCylType.Text
    adoBooking.Recordset("Quantity") = txtQuantity.Text
    adoBooking.Recordset("Status") = txtStatus.Text
    adoBooking.Recordset.Update
    MsgBox "Booking Saved Successfully!", vbInformation, "Success"
    adoBooking.Refresh
    Call cmdClear_Click
End Sub

Private Sub cmdSearch_Click()
    If txtSearch.Text = "" Then
        MsgBox "Please enter Booking ID!", vbExclamation, "Warning"
        Exit Sub
    End If
    Dim rs As New ADODB.Recordset
    rs.Open "SELECT * FROM [tblBooking] WHERE BookingID = " & txtSearch.Text, adoBooking.ConnectionString
    If rs.EOF Then
        MsgBox "Record Not Found!", vbExclamation, "Not Found"
    Else
        txtBookingID.Text = rs("BookingID")
        txtCustID.Text = rs("CustomerID")
        txtBookingDate.Text = rs("BookingDate")
        txtCylType.Text = rs("CylinderType")
        txtQuantity.Text = rs("Quantity")
        txtStatus.Text = rs("Status")
    End If
    rs.Close
End Sub

Private Sub cmdUpdate_Click()
    If txtBookingID.Text = "" Then
        MsgBox "Please Search First!", vbExclamation, "Warning"
        Exit Sub
    End If
    Dim rs As New ADODB.Recordset
    rs.Open "SELECT * FROM [tblBooking] WHERE BookingID = " & txtBookingID.Text, adoBooking.ConnectionString, 1, 3
    If rs.EOF Then
        MsgBox "Record Not Found!", vbExclamation, "Not Found"
    Else
        rs("CustomerID") = txtCustID.Text
        rs("BookingDate") = txtBookingDate.Text
        rs("CylinderType") = txtCylType.Text
        rs("Quantity") = txtQuantity.Text
        rs("Status") = txtStatus.Text
        rs.Update
        MsgBox "Booking Updated Successfully!", vbInformation, "Success"
        adoBooking.Refresh
    End If
    rs.Close
End Sub

Private Sub cmdDelete_Click()
    If txtBookingID.Text = "" Then
        MsgBox "Please Search First!", vbExclamation, "Warning"
        Exit Sub
    End If
    If MsgBox("Are you sure?", vbYesNo + vbQuestion, "Delete") = vbYes Then
        Dim rs As New ADODB.Recordset
        rs.Open "SELECT * FROM [tblBooking] WHERE BookingID = " & txtBookingID.Text, adoBooking.ConnectionString, 1, 3
        If Not rs.EOF Then
            rs.Delete
            MsgBox "Booking Deleted Successfully!", vbInformation, "Success"
            adoBooking.Refresh
            Call cmdClear_Click
        End If
        rs.Close
    End If
End Sub

Private Sub cmdClear_Click()
    txtBookingID.Text = ""
    txtCustID.Text = ""
    txtBookingDate.Text = Format(Now, "DD/MM/YYYY")
    txtCylType.Text = ""
    txtQuantity.Text = ""
    txtStatus.Text = ""
    txtSearch.Text = ""
    txtCustID.SetFocus
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub
