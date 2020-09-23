VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Alert Box Demo"
   ClientHeight    =   2115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2115
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Alert message"
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.CommandButton Command2 
         Caption         =   "About"
         Height          =   375
         Left            =   3360
         TabIndex        =   9
         Top             =   1320
         Width           =   975
      End
      Begin VB.OptionButton Option5 
         Caption         =   "Msn Alert"
         Height          =   255
         Left            =   1200
         TabIndex        =   8
         Top             =   1560
         Width           =   1095
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Msn Type"
         Height          =   255
         Left            =   1200
         TabIndex        =   7
         Top             =   1200
         Width           =   1095
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Msn Email"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   1560
         Width           =   1095
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Original"
         Height          =   255
         Left            =   2280
         TabIndex        =   4
         Top             =   1560
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Msn Online"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         Caption         =   "&Show alert"
         Default         =   -1  'True
         Height          =   375
         Left            =   2760
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Text            =   "You have recived an e-mail from playxcube"
         Top             =   360
         Width           =   4215
      End
      Begin VB.Label Label1 
         Caption         =   "Sound:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   615
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Alert(Text As String)
    ' Display a new alertbox
    Dim AlertBox As frmAlert
    Set AlertBox = New frmAlert
    AlertBox.DisplayAlert Text, 3000
    Me.SetFocus
End Sub

Private Sub Command1_Click()
    Alert Text1.Text
End Sub

Private Sub Command2_Click()
frmSplash.Show
End Sub
