VERSION 5.00
Begin VB.Form frmRegister 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Product Registration"
   ClientHeight    =   4110
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4695
   Icon            =   "frmRegister.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   4695
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtHowMany 
      Height          =   285
      Left            =   2040
      TabIndex        =   12
      Top             =   3720
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtRegistered 
      Height          =   285
      Left            =   960
      TabIndex        =   11
      Top             =   3720
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Picture1 
      Height          =   1935
      Left            =   120
      Picture         =   "frmRegister.frx":0442
      ScaleHeight     =   1875
      ScaleWidth      =   4395
      TabIndex        =   10
      Top             =   0
      Width           =   4455
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Ver 1.0.1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   495
         Left            =   2640
         TabIndex        =   14
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Product Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   13
         Top             =   600
         Width           =   2295
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Registration Information"
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   2040
      Width           =   4455
      Begin VB.CommandButton Command1 
         Caption         =   "Register"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3400
         TabIndex        =   6
         Top             =   1140
         Width           =   855
      End
      Begin VB.TextBox Text3 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2760
         MaxLength       =   4
         TabIndex        =   5
         Top             =   1200
         Width           =   495
      End
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1800
         MaxLength       =   8
         TabIndex        =   4
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1200
         MaxLength       =   4
         TabIndex        =   3
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Enter Registration Number"
         Height          =   255
         Left            =   1200
         TabIndex        =   9
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   "usssssy@yahoo.com"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   2640
         TabIndex        =   2
         Top             =   630
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   $"frmRegister.frx":1C600
         Height          =   735
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   4215
      End
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Continue Un-Registered"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   2760
      TabIndex        =   8
      Top             =   3720
      Width           =   1815
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   3720
      Width           =   615
   End
End
Attribute VB_Name = "frmRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const SW_SHOW = 5
Private Sub Command1_Click()
Dim Registration As String
Registration = Text1.Text & Text2.Text & Text3.Text
'Debug.Print Registration

If ENCRYPT(Registration, 16) <> ">?@A>?@ABCDEBCDE" Then
    MsgBox "You Have Entered An Invalid Registration Number.", vbCritical, "REGISTRATION FAILED"
    Text1.Text = ""
    Text2.Text = ""
    Text2.Enabled = False
    Text3.Text = ""
    Text3.Enabled = False
    Command1.Enabled = False
    Text1.SetFocus
    Exit Sub
End If
'This is where we write the information to the registry
Dim strNg As String
strNg = DECRYPT(">?@A>?@ABCDEBCDE", 16)
CreateKey "HKEY_LOCAL_MACHINE\SOFTWARE\TrialUse"
SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\TrialUse", "TU1", strNg
MsgBox "Thank you for registering Trial Use", vbOKOnly, "REGISTRATION ACKNOWLEDGEMENT"
'load frmMain  <- put the name of the form you want to load here
Unload Me
End Sub

Private Sub Form_Load()
'This "IF/THEN" only works when compiled
If App.PrevInstance = True Then
    MsgBox "This Program Is Already Running", vbCritical, "PREVIOUS INSTANCE DETECTED"
    End
End If
On Error GoTo Err
'Unless the "dll" file exist, this will cause an
'error and cause the product to finish loading
Open "c:\" & "TrailUse.dll" For Input As #1
    Line Input #1, ajt
Close #1
'if the "dll" exist, this means that the user
'has used all the trial uses.  remember to delete
'the "dll" after registration
MsgBox "You have used your trial uses already! You may not re-install this program to continue using it. To obtain a registration number contact the programmer.", vbCritical, "EXPIRED TRIAL USES"
Label4.Visible = False
Label3.Visible = True
frmRegister.Caption = "YOU MUST REGISTER TO CONTINUE USING"
Exit Sub
Err:
FinishLoading
End Sub
Private Sub FinishLoading()
'This is where it reads the registry to see
'if it's been registered yet and how many times
'it's been ran
txtRegistered.Text = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\TrialUse", "TU1")
If txtRegistered.Text = DECRYPT(">?@A>?@ABCDEBCDE", 16) Then
    'load frmMain   <-put the name of the form you want to load here
    Unload Me
    Exit Sub
Else
    retvalue = GetSetting("A", "0", "Runcount")
    gd$ = Val(retvalue) + 1
    SaveSetting "A", "0", "runcount", gd$
    txtHowMany.Text = gd$
End If
    'this is where you determine how many uses the
    'user has before they must register.
If gd$ > 30 Then
    Open "c:\" & "TrialUse.dll" For Output As #1
        Print #1, ajt
    Close #1
    Label4.Visible = False
    Me.Caption = "EXPIRED TRIAL USE"
    Exit Sub
End If
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.ForeColor = &H80&
Label4.ForeColor = &H80&
Label2.ForeColor = &H800000

End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.ForeColor = &H80&
Label4.ForeColor = &H80&
Label2.ForeColor = &H800000
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = &H800000

End Sub

Private Sub Label2_Click()
ShellExecute hwnd, "open", "mailto:usssssy@yahoo.com", vbNullString, vbNullString, SW_SHOW

End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = vbBlue

End Sub

Private Sub Label3_Click()
MsgBox "Thank You For Using Trial Use.  Contact USSSSSY@YAHOO.COM For A Registration Number.", vbOKOnly, "THANK YOU"
End

End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.ForeColor = vbRed

End Sub

Private Sub Label4_Click()
MsgBox "Some Options May Not Be Available Until Registered.", vbOKOnly, "TRIAL-USE"
'load frmMain
Unload Me

End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.ForeColor = vbRed

End Sub

Private Sub Text1_Change()
If Len(Text1.Text) >= 4 Then
    Text2.Enabled = True
    Text2.SetFocus
Else
    Command1.Enabled = False
End If
End Sub

Private Sub Text2_Change()
If Len(Text2.Text) >= 8 Then
    Text3.Enabled = True
    Text3.SetFocus
Else
    Command1.Enabled = False
End If
End Sub

Private Sub Text3_Change()
If Len(Text3.Text) >= 4 Then
    Command1.Enabled = True
Else
    Command1.Enabled = False
End If
End Sub

Private Sub Text3_GotFocus()
If Len(Text3.Text) >= 4 Then Command1.Enabled = True
End Sub
