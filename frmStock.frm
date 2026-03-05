VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmStock 
   BackColor       =   &H00FFFF80&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "LPG Gas Managment System - Stock Managment"
   ClientHeight    =   7830
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14700
   BeginProperty Font 
      Name            =   "Rockwell"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7830
   ScaleWidth      =   14700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSDataGridLib.DataGrid dgStock 
      Bindings        =   "frmStock.frx":0000
      Height          =   7815
      Left            =   7200
      TabIndex        =   18
      Top             =   0
      Width           =   7215
      _ExtentX        =   12726
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
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "StockID"
         Caption         =   "StockID"
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
      BeginProperty Column02 
         DataField       =   "TotalStock"
         Caption         =   "TotalStock"
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
         DataField       =   "AvailableStock"
         Caption         =   "AvailableStock"
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
         DataField       =   "LastUpdated"
         Caption         =   "LastUpdated"
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
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1094.74
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1230.236
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   2085.166
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc adoStock 
      Height          =   330
      Left            =   0
      Top             =   7440
      Width           =   6975
      _ExtentX        =   12303
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
      RecordSource    =   "tblStock"
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
      TabIndex        =   17
      Top             =   6720
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
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   6720
      Width           =   1575
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
      TabIndex        =   15
      Top             =   6720
      Width           =   1575
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
      TabIndex        =   14
      Top             =   6120
      Width           =   1575
   End
   Begin VB.CommandButton cmdUpdate 
      BackColor       =   &H00FF8080&
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
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   6120
      Width           =   1575
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H0000FF00&
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
      TabIndex        =   12
      Top             =   6120
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "Stock Details:"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5175
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   6855
      Begin VB.TextBox txtSearch 
         Height          =   375
         Left            =   3840
         TabIndex        =   19
         Top             =   4440
         Width           =   1095
      End
      Begin VB.TextBox txtLastUpdated 
         Height          =   375
         Left            =   3120
         TabIndex        =   11
         Top             =   3840
         Width           =   3135
      End
      Begin VB.TextBox txtAvailStock 
         Height          =   375
         Left            =   3120
         TabIndex        =   10
         Top             =   3000
         Width           =   3135
      End
      Begin VB.TextBox txtTotalStock 
         Height          =   375
         Left            =   3120
         TabIndex        =   9
         Top             =   2160
         Width           =   3135
      End
      Begin VB.TextBox txtCylType 
         Height          =   375
         Left            =   3120
         TabIndex        =   8
         Top             =   1320
         Width           =   3135
      End
      Begin VB.TextBox txtStockID 
         Enabled         =   0   'False
         Height          =   375
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   480
         Width           =   3135
      End
      Begin VB.Label Label6 
         Caption         =   "Search By ID:"
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1920
         TabIndex        =   20
         Top             =   4440
         Width           =   1815
      End
      Begin VB.Label Label5 
         Caption         =   "Last Updated"
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
         Left            =   480
         TabIndex        =   6
         Top             =   3960
         Width           =   1935
      End
      Begin VB.Label Label4 
         Caption         =   "Available Stock:"
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
         Left            =   480
         TabIndex        =   5
         Top             =   3120
         Width           =   1935
      End
      Begin VB.Label Label3 
         Caption         =   "Total Stock:"
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
         Left            =   480
         TabIndex        =   4
         Top             =   2280
         Width           =   1935
      End
      Begin VB.Label Label2 
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
         Height          =   495
         Left            =   480
         TabIndex        =   3
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Stock ID:"
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
         Left            =   480
         TabIndex        =   2
         Top             =   480
         Width           =   1935
      End
   End
   Begin VB.Label lblHeader 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "Stock Managment"
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
      Height          =   615
      Left            =   -120
      TabIndex        =   0
      Top             =   0
      Width           =   7095
   End
End
Attribute VB_Name = "frmStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    adoStock.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\durgesh computer\Documents\LPGGasSystem.accdb"
    adoStock.RecordSource = "tblStock"
    adoStock.Refresh
    txtLastUpdated.Text = Format(Now, "DD/MM/YYYY")
End Sub

Private Sub cmdSave_Click()
    If txtCylType.Text = "" Or txtTotalStock.Text = "" Then
        MsgBox "Please fill all fields!", vbExclamation, "Warning"
        Exit Sub
    End If
    adoStock.Recordset.AddNew
    adoStock.Recordset("CylinderType") = txtCylType.Text
    adoStock.Recordset("TotalStock") = txtTotalStock.Text
    adoStock.Recordset("AvailableStock") = txtAvailStock.Text
    adoStock.Recordset("LastUpdated") = txtLastUpdated.Text
    adoStock.Recordset.Update
    MsgBox "Stock Saved Successfully!", vbInformation, "Success"
    adoStock.Refresh
    Call cmdClear_Click
End Sub

Private Sub cmdSearch_Click()
    If txtSearch.Text = "" Then
        MsgBox "Please enter Stock ID!", vbExclamation, "Warning"
        Exit Sub
    End If
    Dim rs As New ADODB.Recordset
    rs.Open "SELECT * FROM [tblStock] WHERE StockID = " & txtSearch.Text, adoStock.ConnectionString
    If rs.EOF Then
        MsgBox "Record Not Found!", vbExclamation, "Not Found"
    Else
        txtStockID.Text = rs("StockID")
        txtCylType.Text = rs("CylinderType")
        txtTotalStock.Text = rs("TotalStock")
        txtAvailStock.Text = rs("AvailableStock")
        txtLastUpdated.Text = rs("LastUpdated")
    End If
    rs.Close
End Sub

Private Sub cmdUpdate_Click()
    If txtStockID.Text = "" Then
        MsgBox "Please Search First!", vbExclamation, "Warning"
        Exit Sub
    End If
    Dim rs As New ADODB.Recordset
    rs.Open "SELECT * FROM [tblStock] WHERE StockID = " & txtStockID.Text, adoStock.ConnectionString, 1, 3
    If rs.EOF Then
        MsgBox "Record Not Found!", vbExclamation, "Not Found"
    Else
        rs("CylinderType") = txtCylType.Text
        rs("TotalStock") = txtTotalStock.Text
        rs("AvailableStock") = txtAvailStock.Text
        rs("LastUpdated") = txtLastUpdated.Text
        rs.Update
        MsgBox "Stock Updated Successfully!", vbInformation, "Success"
        adoStock.Refresh
    End If
    rs.Close
End Sub

Private Sub cmdDelete_Click()
    If txtStockID.Text = "" Then
        MsgBox "Please Search First!", vbExclamation, "Warning"
        Exit Sub
    End If
    If MsgBox("Are you sure?", vbYesNo + vbQuestion, "Delete") = vbYes Then
        Dim rs As New ADODB.Recordset
        rs.Open "SELECT * FROM [tblStock] WHERE StockID = " & txtStockID.Text, adoStock.ConnectionString, 1, 3
        If Not rs.EOF Then
            rs.Delete
            MsgBox "Stock Deleted Successfully!", vbInformation, "Success"
            adoStock.Refresh
            Call cmdClear_Click
        End If
        rs.Close
    End If
End Sub

Private Sub cmdClear_Click()
    txtStockID.Text = ""
    txtCylType.Text = ""
    txtTotalStock.Text = ""
    txtAvailStock.Text = ""
    txtLastUpdated.Text = Format(Now, "DD/MM/YYYY")
    txtSearch.Text = ""
    txtCylType.SetFocus
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub
