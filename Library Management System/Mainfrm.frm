VERSION 5.00
Begin VB.Form LMSMainfrm 
   BackColor       =   &H00404040&
   Caption         =   "LMSLibrary Management System"
   ClientHeight    =   8595
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   15270
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   15270
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image Image2 
      Height          =   3525
      Left            =   11880
      Picture         =   "Mainfrm.frx":0000
      Top             =   0
      Width           =   3405
   End
   Begin VB.Image Image1 
      Height          =   8610
      Left            =   0
      Picture         =   "Mainfrm.frx":27426
      Top             =   0
      Width           =   15285
   End
   Begin VB.Menu managebook 
      Caption         =   "Manage books"
      Begin VB.Menu addmenu 
         Caption         =   "Add Book"
         Shortcut        =   ^A
      End
      Begin VB.Menu editmenu 
         Caption         =   "Edit Book"
         Shortcut        =   ^E
      End
      Begin VB.Menu deletemenu 
         Caption         =   "Delete Book"
         Shortcut        =   ^D
      End
      Begin VB.Menu searchmenu 
         Caption         =   "Search Book"
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu transactionmenu 
      Caption         =   "Transaction Menu"
      Begin VB.Menu borrowmenu 
         Caption         =   "Borrow Book"
         Shortcut        =   ^B
      End
      Begin VB.Menu returnmenu 
         Caption         =   "Return Book"
         Shortcut        =   ^R
      End
   End
   Begin VB.Menu memberacc 
      Caption         =   "Member Account"
      Begin VB.Menu addaccount 
         Caption         =   "Add Account"
         Shortcut        =   ^C
      End
      Begin VB.Menu editaccount 
         Caption         =   "Edit Account"
         Shortcut        =   ^F
      End
      Begin VB.Menu deleteaccount 
         Caption         =   "Delete Account"
         Shortcut        =   ^G
      End
      Begin VB.Menu searchaccount 
         Caption         =   "Search Account"
         Shortcut        =   ^H
      End
   End
   Begin VB.Menu records 
      Caption         =   "Records"
      Begin VB.Menu bookr 
         Caption         =   "Book Records"
         Shortcut        =   ^I
      End
      Begin VB.Menu membr 
         Caption         =   "Member Records"
         Shortcut        =   ^M
      End
   End
   Begin VB.Menu acces 
      Caption         =   "Accessories"
      Begin VB.Menu calcu 
         Caption         =   "calculator"
      End
      Begin VB.Menu note 
         Caption         =   "Notepad"
      End
   End
   Begin VB.Menu reg 
      Caption         =   "Register"
   End
   Begin VB.Menu about 
      Caption         =   "About"
   End
   Begin VB.Menu logout 
      Caption         =   "Log Out"
   End
   Begin VB.Menu close 
      Caption         =   "Close"
   End
End
Attribute VB_Name = "LMSMainfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub about_Click()
LMSAboutfrm.Show
End Sub

Private Sub addaccount_Click()
LMSAddaccountfrm.Show
End Sub

Private Sub addmenu_Click()
LMSAddBookfrm.Show
End Sub

Private Sub bookr_Click()
LMSBookrecordfrm.Show
End Sub

Private Sub borrowmenu_Click()
LMSrentfrm.Show
End Sub

Private Sub calcu_Click()
On Error Resume Next
Shell ("calc.exe"), vbNormalFocus
End Sub

Private Sub calendar_Click()
On Error Resume Next
Shell ("cal.exe"), vbNormalFocus
End Sub

Private Sub change_Click()
LMSChangefrm.Show
End Sub

Private Sub close_Click()
   warning = MsgBox("Are you sure you want to Exit?", vbQuestion + vbYesNo, "Exit?")
    If warning = vbYes Then
        End
    End If
End Sub

Private Sub deleteaccount_Click()
LMSDeleteaccountfrm.Show
End Sub

Private Sub deletemenu_Click()
LMSDeletebookfrm.Show
End Sub

Private Sub editaccount_Click()
LMSEditaccountfrm.Show
End Sub

Private Sub editmenu_Click()
LMSEditBookfrm.Show
End Sub

Private Sub logout_Click()
   warning = MsgBox("Are you sure you want to log-out?", vbQuestion + vbYesNo, "Log-Out?")
    If warning = vbYes Then
    Unload Me
        LMSloginfrm.Show
    End If
End Sub

Private Sub memberinfo_Click()
LMSStudentFrm.Show
End Sub

Private Sub membr_Click()
LMSMemberrecordfrm.Show
End Sub

Private Sub note_Click()
On Error Resume Next
Shell ("notepad.exe"), vbNormalFocus
End Sub

Private Sub reg_Click()
LMSRegisterfrm.Show
End Sub

Private Sub returnmenu_Click()
LMSreturnfrm.Show
End Sub

Private Sub searchaccount_Click()
LMSsearchaccountfrm.Show
End Sub

Private Sub searchmenu_Click()
LMSSearchbookFrm.Show
End Sub

Private Sub zoomin_Click()

End Sub
