VERSION 5.00
Begin VB.Form LMSsearchaccountfrm 
   BackColor       =   &H00FF80FF&
   Caption         =   "LMSSearch Account"
   ClientHeight    =   10380
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16620
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   10380
   ScaleWidth      =   16620
   StartUpPosition =   2  'CenterScreen
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
      Left            =   7440
      TabIndex        =   2
      Top             =   3360
      Width           =   3135
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
      Left            =   7440
      TabIndex        =   3
      Top             =   3960
      Width           =   3135
   End
   Begin VB.TextBox txtFirstname 
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
      Left            =   7440
      TabIndex        =   4
      Top             =   4560
      Width           =   3135
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
      Left            =   7440
      TabIndex        =   5
      Top             =   5160
      Width           =   3135
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
      Left            =   7440
      TabIndex        =   6
      Top             =   5760
      Width           =   3135
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
      Left            =   7440
      TabIndex        =   7
      Top             =   6360
      Width           =   3135
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
      Left            =   7680
      TabIndex        =   8
      Top             =   6960
      Width           =   2895
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00C0FFFF&
      Caption         =   "&Clear"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   8400
      Width           =   1935
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C0FFFF&
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
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   8400
      Width           =   1935
   End
   Begin VB.ListBox lstSearch 
      BackColor       =   &H00000080&
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   15.75
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   6960
      Left            =   10800
      TabIndex        =   11
      Top             =   2040
      Width           =   4695
   End
   Begin VB.ComboBox cboCategory 
      BackColor       =   &H00000080&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   480
      ItemData        =   "LMSsearchaccountfrm.frx":0000
      Left            =   7680
      List            =   "LMSsearchaccountfrm.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Tag             =   "      "
      Top             =   2040
      Width           =   3015
   End
   Begin VB.TextBox txtSearch 
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
      Height          =   510
      Left            =   7440
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   2640
      Width           =   3135
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "Student ID:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5880
      TabIndex        =   30
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Last Name:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5760
      TabIndex        =   29
      Top             =   5760
      Width           =   1695
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Address:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5880
      TabIndex        =   28
      Top             =   6360
      Width           =   1695
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Contact Number:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      TabIndex        =   27
      Top             =   6960
      Width           =   1935
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Middle Name:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5760
      TabIndex        =   26
      Top             =   5160
      Width           =   1695
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "First Name:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5760
      TabIndex        =   25
      Top             =   4560
      Width           =   1695
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "School Name:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5760
      TabIndex        =   24
      Top             =   3960
      Width           =   1695
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Category : "
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
      Height          =   375
      Left            =   6000
      TabIndex        =   23
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Type Here : "
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
      Height          =   375
      Left            =   5880
      TabIndex        =   22
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Shape Shape4 
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      Height          =   1095
      Left            =   6240
      Top             =   8280
      Width           =   2175
   End
   Begin VB.Shape Shape3 
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      Height          =   1095
      Left            =   8400
      Top             =   8280
      Width           =   2175
   End
   Begin VB.Image Image2 
      Height          =   8985
      Left            =   5160
      Picture         =   "LMSsearchaccountfrm.frx":0021
      Top             =   1320
      Width           =   10935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Student ID:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5880
      TabIndex        =   21
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Last Name:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6000
      TabIndex        =   20
      Top             =   5760
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Address:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6240
      TabIndex        =   19
      Top             =   6360
      Width           =   1695
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Contact Number:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      TabIndex        =   18
      Top             =   6960
      Width           =   1935
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Middle Name:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5760
      TabIndex        =   17
      Top             =   5160
      Width           =   1695
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "First Name:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6000
      TabIndex        =   16
      Top             =   4560
      Width           =   1695
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "School Name:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5760
      TabIndex        =   15
      Top             =   3960
      Width           =   1695
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      Height          =   1095
      Left            =   6240
      Top             =   8280
      Width           =   2175
   End
   Begin VB.Shape Shape2 
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      Height          =   1095
      Left            =   8400
      Top             =   8280
      Width           =   2175
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Category : "
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
      Height          =   375
      Left            =   6000
      TabIndex        =   14
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Type Here : "
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
      Height          =   375
      Left            =   5880
      TabIndex        =   13
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   3525
      Left            =   1200
      Picture         =   "LMSsearchaccountfrm.frx":13FFF7
      Top             =   3720
      Width           =   3405
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Search Account"
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
      Left            =   5280
      TabIndex        =   12
      Top             =   120
      Width           =   9015
   End
End
Attribute VB_Name = "LMSsearchaccountfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cboCategory_Click()
    clear
End Sub

Private Sub cboCategory_DropDown()
    With cboCategory
        .clear
        .AddItem "Student ID"
        .AddItem "School Name"
        .AddItem "First Name"
        .AddItem "Last Name"
    End With
End Sub

Private Sub cmdClose_Click()
LMSMainfrm.Show
    Unload Me
End Sub

Private Sub cmdCancel_Click()
LMSMainfrm.Show
Unload Me
End Sub

Private Sub cmdClear_Click()
txtStudID.text = ""
txtSchoolname.text = ""
txtFirstname.text = ""
txtMiddlename.text = ""
txtLastname.text = ""
txtAddress.text = ""
txtContact.text = ""
End Sub

Private Sub lstSearch_Click()
    SetCon
    sql = "SELECT * FROM tblStudents WHERE StudentID='" + lstSearch.text + "'"
    SetRs
    txtStudID.text = rs.Fields("StudentID")
    txtSchoolname.text = rs.Fields("SchoolName")
    txtFirstname.text = rs.Fields("FirstName")
    txtMiddlename.text = rs.Fields("MiddleInitial")
    txtLastname.text = rs.Fields("LastName")
    txtAddress.text = rs.Fields("Address")
    txtContact.text = rs.Fields("ContactNumber")
    CloseRS
    CloseCon
End Sub

Private Sub Text2_Change()

End Sub

Private Sub txtSearch_Change()
lstSearch.clear
 If Len(txtSearch.text) > 0 Then
    If Not cboCategory = vbNullString Then
        SetCon
        If cboCategory.text = "Student ID" Then
            sql = "SELECT * FROM tblStudents WHERE StudentID LIKE '" + txtSearch.text + "' & '%'"
            SetRs
            If rs.RecordCount > 0 Then
                rs.MoveFirst
                Do While Not rs.EOF
                    lstSearch.AddItem rs.Fields("StudentID")
                    rs.MoveNext
                Loop
            End If
        ElseIf cboCategory.text = "School Name" Then
            sql = "SELECT * FROM tblStudents WHERE SchoolName LIKE '" + txtSearch.text + "' & '%'"
            SetRs
            If rs.RecordCount > 0 Then
                rs.MoveFirst
                Do While Not rs.EOF
                    lstSearch.AddItem rs.Fields("StudentID")
                    rs.MoveNext
                Loop
            End If
        ElseIf cboCategory.text = "First Name" Then
            sql = "SELECT * FROM tblStudents WHERE FirstName LIKE '" + txtSearch.text + "' & '%'"
            SetRs
            If rs.RecordCount > 0 Then
                rs.MoveFirst
                Do While Not rs.EOF
                    lstSearch.AddItem rs.Fields("StudentID")
                    rs.MoveNext
                Loop
            End If
        ElseIf cboCategory.text = "Last Name" Then
            sql = "SELECT * FROM tblStudents WHERE LastName LIKE '" + txtSearch.text + "' & '%'"
            SetRs
            If rs.RecordCount > 0 Then
                rs.MoveFirst
                Do While Not rs.EOF
                    lstSearch.AddItem rs.Fields("StudentID")
                    rs.MoveNext
                Loop
            End If
        End If
        CloseRS
        CloseCon
    End If
 Else
    lstSearch.clear
    clear
 End If
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
    If Not cboCategory.text = vbNullString Then
        txtSearch.Locked = False
    Else
        txtSearch.Locked = True
    End If
End Sub

Private Function clear()
    With Me
        .txtStudID.text = ""
        .txtSchoolname.text = ""
        .txtFirstname.text = ""
        .txtMiddlename.text = ""
        .txtLastname.text = ""
        .txtAddress.text = ""
        .txtContact.text = ""
        .lstSearch.clear
    End With
End Function

