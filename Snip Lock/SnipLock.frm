VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3255
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2775
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   2775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   1680
      TabIndex        =   6
      Top             =   2760
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   2760
      Width           =   855
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Left            =   240
      TabIndex        =   4
      Top             =   1800
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   2280
      Width           =   2295
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Login:"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   495
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00808000&
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   0
      Top             =   0
      Width           =   255
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00808000&
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   2520
      Top             =   0
      Width           =   255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Snip Lock v. 1.0.1"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2775
   End
   Begin VB.Shape Shape2 
      Height          =   3015
      Left            =   0
      Top             =   240
      Width           =   2775
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Left            =   0
      Top             =   0
      Width           =   2775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim verifyLogin As String
Dim verifyPass As String

    verifyLogin = GetFromINI("User Info:", "Login: ", App.Path & "\" & "info.ini")
    verifyPass = GetFromINI("User Info:", "Password: ", App.Path & "\" & "info.ini")

        If Text1.Text = verifyLogin Then
            If Text2.Text = verifyPass Then
            
            Desktop_IconsShow
            Taskbar_Show
            Enable_CtrlAltDel
            Minimize_AllWin True
            
            End
            
        Else
            Label4.Caption = "Invalid Login/Password!"
                Pause 2
            Label4.Caption = ""
            End If
        Else
            Label4.Caption = "Invalid Login/Password!"
                Pause 2
            Label4.Caption = ""
        End If
End Sub

Private Sub Form_Load()
Dim tCnt As String
Dim tLoad As String
Dim tUser As String
Dim tPass As String


tLoad = GetFromINI("Load", "Loaded: ", App.Path & "\" & "info.ini")

tCnt = tLoad + 1


If tCnt = 1 Then
    tUser = InputBox("Welcome to Snip Lock!" & vbCrLf & vbCrLf & "Please enter a Login name:" & vbCrLf & vbCrLf & "Hint:  If you close this dialog, the Login name will                       be set to """".", "Snip Lock v. 1.0.1")
    tPass = InputBox("Welcome to Snip Lock!" & vbCrLf & vbCrLf & "Please enter a Password:" & vbCrLf & vbCrLf & "Hint:  If you close this dialog, the Password will be set to           """".")

    Call WriteToINI("User Info:", "Login:", tUser, App.Path & "\" & "info.ini")
    Call WriteToINI("User Info:", "Password:", tPass, App.Path & "\" & "info.ini")
End If

Call WriteToINI("Load", "Loaded:", tCnt, App.Path & "\" & "info.ini")

Desktop_IconsHide
Taskbar_Hide
Disable_CtrlAltDel
Minimize_AllWin
End Sub
