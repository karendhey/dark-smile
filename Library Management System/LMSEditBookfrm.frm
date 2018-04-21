VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form LMSEditBookfrm 
   BackColor       =   &H00FF80FF&
   Caption         =   "Form1"
   ClientHeight    =   10935
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   19350
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   10935
   ScaleWidth      =   19350
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   240
      Top             =   240
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
   Begin VB.TextBox TxtVolume 
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
      Height          =   495
      Left            =   13200
      TabIndex        =   7
      Top             =   3840
      Width           =   4575
   End
   Begin VB.TextBox TxtDescription 
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
      Height          =   840
      Left            =   4080
      TabIndex        =   8
      Top             =   4935
      Width           =   13815
   End
   Begin VB.TextBox txtCopyYear 
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
      Height          =   495
      Left            =   13200
      TabIndex        =   4
      Top             =   1680
      Width           =   4575
   End
   Begin VB.TextBox txtCopy 
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
      Height          =   495
      Left            =   13200
      TabIndex        =   5
      Top             =   2400
      Width           =   4575
   End
   Begin VB.TextBox txtCallID 
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
      Height          =   495
      Left            =   5520
      TabIndex        =   0
      Top             =   1680
      Width           =   5175
   End
   Begin VB.TextBox txtISBN 
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
      Height          =   495
      Left            =   5520
      TabIndex        =   3
      Top             =   3960
      Width           =   5175
   End
   Begin VB.TextBox txtTitle 
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
      Height          =   495
      Left            =   5520
      TabIndex        =   1
      Top             =   2400
      Width           =   5175
   End
   Begin VB.TextBox txtAuthor 
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
      Height          =   495
      Left            =   5520
      TabIndex        =   2
      Top             =   3240
      Width           =   5175
   End
   Begin VB.TextBox txtAccession 
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
      Height          =   495
      Left            =   13200
      TabIndex        =   6
      Top             =   3120
      Width           =   4575
   End
   Begin VB.CommandButton cmdEdit 
      BackColor       =   &H00C0FFFF&
      Caption         =   "&Edit"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   11520
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   9960
      Width           =   1935
   End
   Begin VB.CommandButton cmdsave 
      BackColor       =   &H00C0FFFF&
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   13680
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   9960
      Width           =   1935
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C0FFFF&
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   15840
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   9960
      Width           =   1935
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3975
      Left            =   4080
      TabIndex        =   12
      Top             =   5880
      Width           =   13935
      _ExtentX        =   24580
      _ExtentY        =   7011
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   128
      ForeColor       =   65535
      HeadLines       =   1
      RowHeight       =   18
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
         Weight          =   600
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
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright Year:"
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
      Left            =   11160
      TabIndex        =   33
      Top             =   1920
      Width           =   2175
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "No. of Copies : "
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
      Left            =   11160
      TabIndex        =   32
      Top             =   2640
      Width           =   1935
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   " Accession No. : "
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
      Left            =   11040
      TabIndex        =   31
      Top             =   3240
      Width           =   2055
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Description: "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   9600
      TabIndex        =   30
      Top             =   4440
      Width           =   2055
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   " Volume : "
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
      Left            =   11040
      TabIndex        =   29
      Top             =   3960
      Width           =   1575
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Call ID:"
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
      Left            =   4200
      TabIndex        =   28
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "ISBN #: "
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
      Left            =   4320
      TabIndex        =   27
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   " Title : "
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
      Left            =   4200
      TabIndex        =   26
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Author : "
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
      Left            =   4200
      TabIndex        =   25
      Top             =   3240
      Width           =   2055
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Edit Book"
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
      Left            =   5760
      TabIndex        =   24
      Top             =   -120
      Width           =   9015
   End
   Begin VB.Image Image2 
      Height          =   3525
      Left            =   -480
      Picture         =   "LMSEditBookfrm.frx":0000
      Top             =   4080
      Width           =   3405
   End
   Begin VB.Image Image1 
      Height          =   10035
      Left            =   3000
      Picture         =   "LMSEditBookfrm.frx":27426
      Top             =   960
      Width           =   15825
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Shelves : "
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
      Left            =   10200
      TabIndex        =   23
      Top             =   6360
      Width           =   1575
   End
   Begin VB.Label lblTime 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9360
      TabIndex        =   22
      Top             =   5760
      Width           =   3255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Date and Time:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   10080
      TabIndex        =   21
      Top             =   6720
      Width           =   2775
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Description: "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   9840
      TabIndex        =   20
      Top             =   4440
      Width           =   2055
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   " Volume : "
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
      Left            =   11400
      TabIndex        =   19
      Top             =   3840
      Width           =   1575
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Quantity : "
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
      Left            =   11400
      TabIndex        =   18
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Year Published : "
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
      Left            =   11040
      TabIndex        =   17
      Top             =   1920
      Width           =   2175
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Book Author : "
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
      Left            =   4320
      TabIndex        =   16
      Top             =   3840
      Width           =   2055
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Book Title : "
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
      Left            =   4320
      TabIndex        =   15
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "ISBN #: "
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
      Left            =   4440
      TabIndex        =   14
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Book ID:"
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
      Left            =   4320
      TabIndex        =   13
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H00000000&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      Height          =   855
      Left            =   15720
      Top             =   9840
      Width           =   2175
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      Height          =   855
      Left            =   13560
      Top             =   9840
      Width           =   2175
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00000000&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      Height          =   855
      Left            =   11400
      Top             =   9840
      Width           =   2175
   End
End
Attribute VB_Name = "LMSEditBookfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim warning As String

Private Sub cmdClose_Click()
LMSMainfrm.Show
    Unload Me
End Sub

Private Function txtL()
    With Me
        .txtCopy.Enabled = False
        .txtCallID.Enabled = False
        .txtTitle.Enabled = False
        .txtAuthor.Enabled = False
        .txtISBN.Enabled = False
        .txtCopyYear.Enabled = False
        .txtAccession.Enabled = False
        .TxtVolume.Enabled = False
        .TxtDescription.Enabled = False
        .cmdsave.Enabled = False
        cmdEdit.Enabled = True
    End With
End Function

Private Function txtUL()
    With Me
        .txtCopy.Enabled = True
        .txtCallID.Enabled = True
        .txtTitle.Enabled = True
        .txtAuthor.Enabled = True
        .txtISBN.Enabled = True
        .txtCopyYear.Enabled = True
        .txtAccession.Enabled = True
        .TxtVolume.Enabled = True
        .TxtDescription.Enabled = True
        .cmdsave.Enabled = True
        cmdEdit.Enabled = False
    End With
End Function
Private Sub cmdCancel_Click()
LMSMainfrm.Show
    Unload Me
End Sub

Private Sub cmdEdit_Click()
Dim sInput$
re:
    sInput = InputBox("Type here...", "Enter ID of the book!")
    If sInput <> "" Then
        SetCon
        sql = "SELECT * FROM tblBooks WHERE CallID='" + Trim(sInput) + "'"
        SetRs
        If rs.RecordCount > 0 Then
            Call txtUL
            With Me
                .txtCallID.text = rs.Fields("CallID")
                .txtTitle.text = rs.Fields("Title")
                .txtAuthor.text = rs.Fields("Author")
                .txtISBN.text = rs.Fields("ISBN")
                .txtCopyYear.text = rs.Fields("CopyrightYear")
                .txtCopy.text = rs.Fields("NumofCopies")
                .txtAccession.text = rs.Fields("AccessionNum")
                .TxtVolume.text = rs.Fields("Volume")
                .TxtDescription.text = rs.Fields("Description")
                
            End With
        Else
            warning = MsgBox("Call ID not exist on the database!", vbCritical + vbRetryCancel, "Unable to Edit")
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

Private Sub cmdSave_Click()
    If txtISBN.text = "" Or txtAuthor.text = "" Or txtTitle.text = "" Or txtCopyYear.text = "" Or txtCopy.text = "" Or txtAccession.text = "" Or TxtVolume.text = "" Or TxtDescription.text = "" Then
        Call MsgBox("Unable to update, One of the field(s) is empty", vbCritical + vbOKOnly, "Error!")
        Exit Sub
    ElseIf Not IsNumeric(txtISBN.text) Then
        Call MsgBox("ISBN contain a non numeric value!", vbExclamation, "Unable to Edit!")
        txtISBN.SetFocus
        Exit Sub
    ElseIf Not IsNumeric(txtCopyYear.text) Then
        Call MsgBox("Copyright Year contain a non numeric value!", vbExclamation, "Unable to Edit!")
        txtCopyYear.SetFocus
        Exit Sub
    ElseIf Not IsNumeric(txtCopy.text) Then
        Call MsgBox("Number of copies contain a non numeric value!", vbExclamation, "Unable to Edit!")
        txtCopy.SetFocus
        Exit Sub
    ElseIf Not IsNumeric(txtAccession.text) Then
        Call MsgBox("Accession no. contain a non numeric value!", vbExclamation, "Unable to Edit!")
        txtAccession.SetFocus
        Exit Sub
    ElseIf Not IsNumeric(TxtVolume.text) Then
        Call MsgBox("Volume contain a non numeric value!", vbExclamation, "Unable to Edit!")
        TxtVolume.SetFocus
        Exit Sub
    End If
    SetCon
    sql = "SELECT * FROM tblBooks WHERE CallID='" + txtCallID.text + "'"
    SetRs
    If rs.RecordCount > 0 Then
        cn.Execute "UPDATE tblBooks SET ISBN='" + txtISBN.text + "',Title='" + txtTitle.text + "',Author='" + txtAuthor.text + "',CopyrightYear='" + txtCopyYear.text + "',NumofCopies='" + txtCopy.text + "' ,AccessionNum='" + txtAccession.text + "' ,Volume='" + TxtVolume.text + "' ,Description='" + TxtDescription.text + "' WHERE CallID='" + txtCallID.text + "'"
        Call MsgBox("Book has been successfully updated!", vbInformation, "Updated!")
        Adodc1.Refresh
        Call txtL
        Call ClearText(LMSEditBookfrm)
    End If
    CloseRS
    CloseCon
End Sub

Private Sub Form_Load()
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Database\LibrarySys.mdb;Persist Security Info=False"
Adodc1.RecordSource = "Select*from tblBooks"
Set DataGrid1.DataSource = Adodc1
End Sub

Private Sub txtBookID_Change()
If KeyAscii = 8 Or KeyAscii = vbKeyDecimal Then
KeyAscii = 8
Else

If (KeyAscii > 47 And KeyAscii < 58) Then
KeyAscii = KeyAscii

Else
KeyAscii = 0
End If
End If
End Sub

Private Sub txtBookID_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Or KeyAscii = vbKeyDecimal Then
KeyAscii = 8
Else

If (KeyAscii > 47 And KeyAscii < 58) Then
KeyAscii = KeyAscii

Else
KeyAscii = 0
End If
End If
End Sub

Private Sub txtISBN_Change()
If KeyAscii = 8 Or KeyAscii = vbKeyDecimal Then
KeyAscii = 8
Else

If (KeyAscii > 47 And KeyAscii < 58) Then
KeyAscii = KeyAscii

Else
KeyAscii = 0
End If
End If
End Sub

Private Sub txtISBN_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Or KeyAscii = vbKeyDecimal Then
KeyAscii = 8
Else

If (KeyAscii > 47 And KeyAscii < 58) Then
KeyAscii = KeyAscii

Else
KeyAscii = 0
End If
End If
End Sub


Private Sub txtQuantity_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Or KeyAscii = vbKeyDecimal Then
KeyAscii = 8
Else

If (KeyAscii > 47 And KeyAscii < 58) Then
KeyAscii = KeyAscii

Else
KeyAscii = 0
End If
End If
End Sub

Private Sub txtYearPub_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Or KeyAscii = vbKeyDecimal Then
KeyAscii = 8
Else

If (KeyAscii > 47 And KeyAscii < 58) Then
KeyAscii = KeyAscii

Else
KeyAscii = 0
End If
End If
End Sub

