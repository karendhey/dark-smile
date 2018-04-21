VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form LMSEditaccountfrm 
   BackColor       =   &H00FF80FF&
   Caption         =   "LMSEdit account"
   ClientHeight    =   10935
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   20250
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   600
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
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
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   11160
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   9480
      Width           =   1935
   End
   Begin VB.TextBox txtFirstname 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      BorderStyle     =   0  'None
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
      Height          =   435
      Left            =   5880
      TabIndex        =   2
      Top             =   2760
      Width           =   4815
   End
   Begin VB.TextBox txtSchoolname 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      Height          =   405
      Left            =   5880
      TabIndex        =   1
      Top             =   2160
      Width           =   4815
   End
   Begin VB.TextBox txtMiddlename 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      Height          =   405
      Left            =   5880
      TabIndex        =   3
      Top             =   3360
      Width           =   4815
   End
   Begin VB.TextBox txtLastname 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      Height          =   405
      Left            =   14760
      TabIndex        =   4
      Top             =   1560
      Width           =   4695
   End
   Begin VB.TextBox txtContact 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      Height          =   405
      Left            =   14760
      TabIndex        =   6
      Top             =   3120
      Width           =   4695
   End
   Begin VB.TextBox txtAddress 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      Height          =   405
      Left            =   14760
      TabIndex        =   5
      Top             =   2400
      Width           =   4695
   End
   Begin VB.TextBox txtStudID 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      Height          =   405
      Left            =   5880
      TabIndex        =   0
      Top             =   1560
      Width           =   4815
   End
   Begin VB.CommandButton cmdClose 
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
      Height          =   855
      Left            =   15240
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   9480
      Width           =   1935
   End
   Begin VB.CommandButton cmdEdit 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Edit"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   9480
      Width           =   2055
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "LMSEditaccountfrm.frx":0000
      Height          =   5295
      Left            =   4200
      TabIndex        =   10
      Top             =   3960
      Width           =   15495
      _ExtentX        =   27331
      _ExtentY        =   9340
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
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Contact number:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   11520
      TabIndex        =   25
      Top             =   3240
      Width           =   2295
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Address:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   12240
      TabIndex        =   24
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Last name:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   12120
      TabIndex        =   23
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "School name:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   3840
      TabIndex        =   22
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Middle name:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   3840
      TabIndex        =   21
      Top             =   3360
      Width           =   1815
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "First name:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   3720
      TabIndex        =   20
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Shape Shape10 
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      Height          =   1095
      Left            =   11040
      Top             =   9360
      Width           =   2175
   End
   Begin VB.Shape Shape9 
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      Height          =   1095
      Left            =   6600
      Top             =   9360
      Width           =   2295
   End
   Begin VB.Shape Shape5 
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      Height          =   1095
      Left            =   15120
      Top             =   9360
      Width           =   2175
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H8000000E&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      Height          =   1095
      Left            =   5280
      Top             =   6720
      Width           =   2175
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H8000000E&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      Height          =   855
      Left            =   5400
      Top             =   6720
      Width           =   2175
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Student ID:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   3840
      TabIndex        =   19
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00000000&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      Height          =   1095
      Left            =   4800
      Top             =   6480
      Width           =   2175
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00000000&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      Height          =   1095
      Left            =   4800
      Top             =   7800
      Width           =   2175
   End
   Begin VB.Image Image2 
      Height          =   10305
      Left            =   3480
      Picture         =   "LMSEditaccountfrm.frx":0015
      Top             =   1200
      Width           =   17055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "School name:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   4080
      TabIndex        =   18
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Contact number:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   12000
      TabIndex        =   17
      Top             =   3120
      Width           =   2295
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Address:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   12720
      TabIndex        =   16
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Last name:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   12600
      TabIndex        =   15
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Middle name:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   4200
      TabIndex        =   14
      Top             =   3360
      Width           =   1815
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "First name:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   4080
      TabIndex        =   13
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Student ID:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   3960
      TabIndex        =   12
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Shape Shape6 
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      Height          =   1095
      Left            =   11040
      Top             =   9360
      Width           =   2175
   End
   Begin VB.Shape Shape7 
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      Height          =   1095
      Left            =   6600
      Top             =   9360
      Width           =   2295
   End
   Begin VB.Shape Shape8 
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      Height          =   1095
      Left            =   15120
      Top             =   9360
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   3525
      Left            =   0
      Picture         =   "LMSEditaccountfrm.frx":23C4C3
      Top             =   3600
      Width           =   3405
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Edit Account"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   6000
      TabIndex        =   11
      Top             =   -120
      Width           =   9015
   End
End
Attribute VB_Name = "LMSEditaccountfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim warning As String
Private Sub cmdSave_Click()
    If txtSchoolname.text = "" Or txtFirstname.text = "" Or txtMiddlename.text = "" Or txtLastname.text = "" Or txtAddress.text = "" Or txtContact.text = "" Then
        Call MsgBox("Unable to update, One of the field(s) is empty", vbCritical + vbOKOnly, "Error!")
        Exit Sub
    ElseIf Not IsNumeric(txtContact.text) Then
        Call MsgBox("Contact number contain a non numeric value!", vbExclamation, "Unable to Add!")
        txtContact.SetFocus
        Exit Sub
    End If
    SetCon
    sql = "SELECT * FROM tblStudents WHERE StudentID='" + txtStudID.text + "'"
    SetRs
    If rs.RecordCount > 0 Then
        cn.Execute "UPDATE tblStudents SET SchoolName='" + txtSchoolname.text + "',FirstName='" + txtFirstname.text + "',MiddleInitial='" + txtMiddlename.text + "',LastName='" + txtLastname.text + "',Address='" + txtAddress.text + "' ,ContactNumber='" + txtContact.text + "' WHERE StudentID='" + txtStudID.text + "'"
        Call MsgBox("Account has been successfully updated!", vbInformation, "Updated!")
        Adodc1.Refresh
        Call txtL
        Call ClearText(LMSEditaccountfrm)
    End If
    CloseRS
    CloseCon
End Sub

Private Sub Form_Load()
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Database\LibrarySys.mdb;Persist Security Info=False"
Adodc1.RecordSource = "Select*from tblStudents"
Set DataGrid1.DataSource = Adodc1
End Sub

Private Sub cmdClose_Click()
LMSMainfrm.Show
    Unload Me
End Sub

Private Function txtL()
    With Me
        .txtContact.Enabled = False
        .txtStudID.Enabled = False
        .txtSchoolname.Enabled = False
        .txtFirstname.Enabled = False
        .txtMiddlename.Enabled = False
        .txtLastname.Enabled = False
        .txtAddress.Enabled = False
        .cmdsave.Enabled = False
        cmdEdit.Enabled = True
    End With
End Function

Private Function txtUL()
    With Me
        .txtContact.Enabled = True
        .txtStudID.Enabled = True
        .txtSchoolname.Enabled = True
        .txtFirstname.Enabled = True
        .txtMiddlename.Enabled = True
        .txtLastname.Enabled = True
        .txtAddress.Enabled = True
        .cmdsave.Enabled = True
        .cmdEdit.Enabled = False
    End With
End Function


Private Sub cmdEdit_Click()
Dim sInput$
re:
    sInput = InputBox("Type here...", "Enter ID of the student!")
    If sInput <> "" Then
        SetCon
        sql = "SELECT * FROM tblStudents WHERE StudentID='" + Trim(sInput) + "'"
        SetRs
        If rs.RecordCount > 0 Then
    
            Call txtUL
            With Me
                .txtStudID.text = rs.Fields("StudentID")
                .txtSchoolname.text = rs.Fields("SchoolName")
                .txtFirstname.text = rs.Fields("FirstName")
                .txtMiddlename.text = rs.Fields("MiddleInitial")
                .txtLastname.text = rs.Fields("LastName")
                .txtAddress.text = rs.Fields("Address")
                .txtContact.text = rs.Fields("ContactNumber")
                
            End With
        Else
            warning = MsgBox("Student ID not exist on the database!", vbCritical + vbRetryCancel, "Unable to Edit")
            If warning = vbRetry Then
                GoTo re
            End If
        End If
        CloseRS
        CloseCon
    Else
        warning = MsgBox("Empty!", vbCritical + vbRetryCancel, "Oppss!")
        If warning = vbRetry Then
            GoTo re
        End If
End If

   
End Sub


