VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form LMSDeleteaccountfrm 
   BackColor       =   &H00FF80FF&
   Caption         =   "LMSDelete account"
   ClientHeight    =   8850
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17445
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   8850
   ScaleWidth      =   17445
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   240
      Top             =   120
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
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
   Begin VB.TextBox txtStudID 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   615
      Left            =   9840
      TabIndex        =   1
      Top             =   2160
      Width           =   2895
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Delete"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7560
      Width           =   1695
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11760
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7560
      Width           =   1695
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "LMSDeleteaccountfrm.frx":0000
      Height          =   4335
      Left            =   4440
      TabIndex        =   3
      Top             =   3000
      Width           =   12975
      _ExtentX        =   22886
      _ExtentY        =   7646
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   128
      ForeColor       =   65535
      HeadLines       =   1
      RowHeight       =   20
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bookman Old Style"
         Size            =   9
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Student ID:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10200
      TabIndex        =   6
      Top             =   1680
      Width           =   2415
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00000000&
      BorderColor     =   &H00000000&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      Height          =   735
      Left            =   11640
      Top             =   7440
      Width           =   1935
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BorderColor     =   &H00000000&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      Height          =   735
      Left            =   8640
      Top             =   7440
      Width           =   1935
   End
   Begin VB.Image Image2 
      Height          =   7275
      Left            =   4440
      Picture         =   "LMSDeleteaccountfrm.frx":0015
      Top             =   1560
      Width           =   12945
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Student ID:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10320
      TabIndex        =   5
      Top             =   1680
      Width           =   2415
   End
   Begin VB.Shape Shape2 
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      Height          =   735
      Left            =   8520
      Top             =   7440
      Width           =   1935
   End
   Begin VB.Shape Shape3 
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      Height          =   735
      Left            =   11640
      Top             =   7440
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   3525
      Left            =   600
      Picture         =   "LMSDeleteaccountfrm.frx":132EF7
      Top             =   3360
      Width           =   3405
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Delete Account"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6360
      TabIndex        =   4
      Top             =   120
      Width           =   6855
   End
End
Attribute VB_Name = "LMSDeleteaccountfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim warning$

Private Sub cmdCancel_Click()
LMSMainfrm.Show
    Unload Me
End Sub

Private Sub cmdDelete_Click()
    SetCon
    sql = "SELECT * FROM tblStudents WHERE StudentID='" + txtStudID.text + "'"
    SetRs
    If rs.RecordCount > 0 Then
        warning = MsgBox("Are you sure you want to delete this record?", vbQuestion + vbYesNo, "Deleting...")
        If warning = vbYes Then
            cn.Execute "DELETE * FROM tblStudents WHERE StudentID='" + txtStudID.text + "'"
            Call MsgBox("Account has been successfully deleted", vbInformation, "Deleted!")
            Adodc1.Refresh
            CloseRS
            CloseCon
        End If
    Else
        Call MsgBox("The account you want to delete does exist on the database!", vbCritical, "Unable to delete!")
        txtStudID.SetFocus
    End If
End Sub

Private Sub Form_Load()
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Database\LibrarySys.mdb;Persist Security Info=False"
Adodc1.RecordSource = "Select*from tblStudents"
Set DataGrid1.DataSource = Adodc1
End Sub

Private Sub txtStudID_Change()
    If Len(txtStudID.text) > 0 Then
        cmdDelete.Enabled = True
    Else
        cmdDelete.Enabled = False
    End If
End Sub

Private Sub txtStudID_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call cmdDelete_Click
    End If
End Sub


