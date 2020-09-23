VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "Local Machine Info"
   ClientHeight    =   1890
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2820
   LinkTopic       =   "Form1"
   ScaleHeight     =   1890
   ScaleWidth      =   2820
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.TextBox User 
      Height          =   285
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox Port 
      Height          =   285
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox Host 
      Height          =   285
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
   Begin MSWinsockLib.Winsock Winsock 
      Left            =   2160
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   327681
   End
   Begin VB.TextBox IP 
      Height          =   285
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "UserID"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   240
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "Port"
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "Host"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   960
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "IP"
      Height          =   255
      Left            =   600
      TabIndex        =   3
      Top             =   600
      Width           =   255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim UserName As String

Sub Get_User_Name()

     ' Dimension variables
     Dim lpBuff As String * 25
     

     ' Get the user name minus any trailing spaces found in the name.
     ret = GetUserName(lpBuff, 25)
     UserName = Left(lpBuff, InStr(lpBuff, Chr(0)) - 1)

   
End Sub

Private Sub Form_Load()
Get_User_Name
IP.Text = Winsock.LocalIP
Host.Text = Winsock.LocalHostName
Port.Text = Winsock.LocalPort
User.Text = UserName
End Sub
