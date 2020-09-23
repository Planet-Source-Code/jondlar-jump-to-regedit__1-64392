VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "GreenEye2oo4 // Open Key in RegEdit [20th Feb 2006]"
   ClientHeight    =   2400
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7035
   LinkTopic       =   "Form1"
   ScaleHeight     =   2400
   ScaleWidth      =   7035
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtValueName 
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   6855
   End
   Begin VB.ComboBox comboRegKey 
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   360
      Width           =   6855
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Open in RegEdit"
      Height          =   375
      Left            =   5400
      TabIndex        =   0
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Value to select :"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   1140
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Key to open in RegEdit :"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1725
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOpen_Click()
    Call modRegEditJump.RegEditJump(comboRegKey.Text, txtValueName.Text)
End Sub


Private Sub Form_Load()
    With comboRegKey
        .AddItem "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Internet Explorer\Main"
        .AddItem "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Applets\Regedit"
        .AddItem "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Fonts"
        .AddItem "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run"
        .AddItem "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
        .Text = .List(0)
    End With
    txtValueName.Text = "Start Page"
End Sub
