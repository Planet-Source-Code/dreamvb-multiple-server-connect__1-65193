VERSION 5.00
Begin VB.Form FrmCfg 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Config"
   ClientHeight    =   3330
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   5415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      Height          =   345
      Left            =   1695
      ScaleHeight     =   285
      ScaleWidth      =   3540
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   120
      Width           =   3600
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   60
         TabIndex        =   9
         Top             =   30
         Width           =   60
      End
   End
   Begin VB.PictureBox pBottom 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   5415
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2835
      Width           =   5415
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   350
         Left            =   3060
         TabIndex        =   4
         Top             =   75
         Width           =   1035
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update"
         Height          =   350
         Left            =   4245
         TabIndex        =   5
         Top             =   75
         Width           =   1035
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   1
         X1              =   -45
         X2              =   420
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         Index           =   0
         X1              =   -30
         X2              =   435
         Y1              =   0
         Y2              =   0
      End
   End
   Begin VB.PictureBox PicTab 
      BackColor       =   &H00FFFFFF&
      Height          =   2295
      Left            =   1695
      ScaleHeight     =   2235
      ScaleWidth      =   3540
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   450
      Width           =   3600
      Begin VB.TextBox TxtServName 
         Height          =   285
         Left            =   105
         MaxLength       =   30
         TabIndex        =   3
         Text            =   "~::Some Test Server::~"
         Top             =   1245
         Width           =   2640
      End
      Begin VB.TextBox TxtMaxConn 
         Height          =   300
         Left            =   1410
         MaxLength       =   5
         TabIndex        =   2
         Text            =   "20"
         Top             =   525
         Width           =   705
      End
      Begin VB.TextBox txtPort 
         Height          =   300
         Left            =   1020
         MaxLength       =   5
         TabIndex        =   1
         Text            =   "90"
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label lblServname 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Server name:"
         Height          =   195
         Left            =   105
         TabIndex        =   12
         Top             =   1005
         Width           =   945
      End
      Begin VB.Label lblMax 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Max Connections:"
         Height          =   195
         Left            =   105
         TabIndex        =   11
         Top             =   585
         Width           =   1275
      End
      Begin VB.Label lblport 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Listen Port:"
         Height          =   195
         Left            =   105
         TabIndex        =   10
         Top             =   180
         Width           =   795
      End
   End
   Begin VB.ListBox LstCfg 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2610
      IntegralHeight  =   0   'False
      Left            =   75
      TabIndex        =   0
      Top             =   135
      Width           =   1590
   End
End
Attribute VB_Name = "FrmCfg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub DoDefault(mTextBox As TextBox, iDefaultVal As Variant)
    If Not IsNumeric(mTextBox.Text) Then
        mTextBox.Text = iDefaultVal
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload FrmCfg
End Sub

Private Sub cmdUpdate_Click()
    SaveSetting "VbServTest", "Cfg", "Port", txtPort.Text
    SaveSetting "VbServTest", "Cfg", "MaxUsers", TxtMaxConn.Text
    SaveSetting "VbServTest", "Cfg", "ServName", TxtServName.Text
    'Reload the server info
    Call LoadCfg
    Unload FrmCfg
End Sub

Private Sub Form_Load()
    Set FrmCfg.Icon = Nothing
    LstCfg.AddItem "General"
    LstCfg.ListIndex = 0
    
    'Display server config
    'Server port
    txtPort.Text = TServCfg.ServPort
    'Max connections
    TxtMaxConn.Text = TServCfg.ServMaxConn
    'Serv name
    TxtServName.Text = TServCfg.ServName
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set FrmCfg = Nothing
End Sub

Private Sub LstCfg_Click()
    lblCaption.Caption = LstCfg.Text
End Sub

Private Sub pBottom_Resize()
    Line1(0).X2 = pBottom.ScaleWidth
    Line1(1).X2 = Line1(0).X2
End Sub

Private Sub TxtMaxConn_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then Exit Sub
    If Not IsNumeric(Chr(KeyAscii)) Then KeyAscii = 0
End Sub

Private Sub TxtMaxConn_LostFocus()
    DoDefault TxtMaxConn, 20
End Sub

Private Sub txtPort_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then Exit Sub
    If Not IsNumeric(Chr(KeyAscii)) Then KeyAscii = 0
End Sub

Private Sub txtPort_LostFocus()
    DoDefault txtPort, 90
End Sub
