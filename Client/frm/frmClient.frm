VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmClient 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Client"
   ClientHeight    =   3285
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6075
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   6075
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAbout 
      Caption         =   "About"
      Height          =   350
      Left            =   3945
      TabIndex        =   11
      Top             =   2790
      Width           =   915
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send Client Msg"
      Enabled         =   0   'False
      Height          =   350
      Left            =   2040
      TabIndex        =   9
      Top             =   2790
      Width           =   1830
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Get Client List"
      Height          =   350
      Left            =   135
      TabIndex        =   8
      Top             =   2790
      Width           =   1830
   End
   Begin VB.ListBox LstClients 
      Height          =   1230
      Left            =   135
      TabIndex        =   7
      Top             =   1500
      Width           =   5880
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   330
      Left            =   5055
      TabIndex        =   6
      Top             =   285
      Width           =   930
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Connect"
      Enabled         =   0   'False
      Height          =   330
      Left            =   3735
      TabIndex        =   4
      Top             =   285
      Width           =   1215
   End
   Begin MSWinsockLib.Winsock WskConnect 
      Left            =   4380
      Top             =   675
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox txtPort 
      Height          =   300
      Left            =   2745
      TabIndex        =   3
      Text            =   "90"
      Top             =   300
      Width           =   885
   End
   Begin VB.TextBox txtHost 
      Height          =   300
      Left            =   165
      TabIndex        =   1
      Text            =   "127.0.0.1"
      Top             =   300
      Width           =   2445
   End
   Begin VB.Timer TmrConn 
      Enabled         =   0   'False
      Interval        =   30
      Left            =   3870
      Top             =   675
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Connected Clients:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   210
      TabIndex        =   10
      Top             =   1200
      Width           =   1620
   End
   Begin VB.Label lblServname 
      AutoSize        =   -1  'True
      Caption         =   "#"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   210
      TabIndex        =   5
      Top             =   855
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   210
      X2              =   1500
      Y1              =   765
      Y2              =   765
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   0
      X1              =   210
      X2              =   1470
      Y1              =   750
      Y2              =   750
   End
   Begin VB.Label lblport 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Port:"
      Height          =   195
      Left            =   2745
      TabIndex        =   2
      Top             =   90
      Width           =   330
   End
   Begin VB.Label lblhost 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Host:"
      Height          =   195
      Left            =   165
      TabIndex        =   0
      Top             =   90
      Width           =   375
   End
End
Attribute VB_Name = "frmClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public HostServ As String
Public ServPort As Long
Public iConnect As Boolean
Public iClientName As String

Private Sub cmdAbout_Click()
    MsgBox "Client and server example." _
    & vbCrLf & vbTab & "By Ben Jones", vbInformation, "About.."
End Sub

Private Sub cmdConnect_Click()
On Error GoTo ErrFlag:
    iConnect = Not iConnect
    lblServname.Caption = ""
    
    If (cmdConnect.Caption = "Connect") Then
        cmdConnect.Caption = "Disconnect"
    ElseIf (cmdConnect.Caption = "Disconnect") Then
        cmdConnect.Caption = "Connect"
    End If
    
    TmrConn.Enabled = iConnect
    
    If (Not iConnect) Then
        WskConnect.Close
        Exit Sub
    Else
        HostServ = Trim(txtHost.Text)
        ServPort = Val(txtPort.Text)
        WskConnect.Protocol = sckTCPProtocol
        WskConnect.Connect HostServ, ServPort
    End If
    
    Exit Sub
ErrFlag:
    If Err Then MsgBox Err.Description, vbCritical, "Error_" & Err.Number
    
End Sub

Private Sub cmdExit_Click()
    If (WskConnect.State = sckConnected) Then
        MsgBox "Your still connected to the server." _
        & vbCrLf & "Please disconnect first.", vbInformation, frmClient.Caption
        Exit Sub
    Else
        HostServ = ""
        ServPort = 0
        Unload frmClient
    End If
    
End Sub

Private Sub cmdRefresh_Click()
    If (WskConnect.State = sckConnected) Then
        'Send the server a request for the connected client list
        WskConnect.SendData "CCL " & vbCrLf
    End If
End Sub

Private Sub cmdSend_Click()
Dim Client_Idx As Integer, e_pos As Integer, sText As String
    
    If (LstClients.ListCount = 0) Or (WskConnect.State <> sckConnected) Then
        Exit Sub
    Else
        sText = LstClients.Text
        'Send the select client a message
        'Extract the Client ID
        e_pos = InStrRev(sText, ":", Len(sText), vbBinaryCompare)
        If (e_pos <> 0) Then
            Client_Idx = Val(Mid(sText, e_pos + 1, Len(sText)))
            sText = Trim(InputBox("Enter a message for the client: ", "Send Client Message"))
            'Don't send the message if empty
            If Len(sText) = 0 Then
                Exit Sub
            Else
                'Send the server the message to send to the client
                WskConnect.SendData "CSM " & Client_Idx & " Message from: " & iClientName & vbCrLf _
                & sText & vbCrLf
            End If
        End If
    End If
End Sub

Private Sub Form_Load()
    Call txtHost_Change
    iClientName = Environ("USERNAME")
    If Len(iClientName) = 0 Then iClientName = "Client User"
End Sub

Private Sub Form_Resize()
    Line1(0).X2 = (frmClient.ScaleWidth - Line1(0).X1)
    Line1(1).X2 = Line1(0).X2
End Sub

Private Sub LstClients_Click()
    cmdSend.Enabled = True
End Sub

Private Sub TmrConn_Timer()
    'Small timer to see if we are connected to the server

    If (WskConnect.State = sckConnected) Then
        'Send the server this PC's name
        WskConnect.SendData "CPN " & Environ("COMPUTERNAME") & vbCrLf
        TmrConn.Enabled = False
        Exit Sub
    ElseIf (WskConnect.State = sckError) Then
        'Opps seems like an error or timeout
        MsgBox "Error or time out while connecting to host.", vbCritical, "Error_" & Err.Number
        TmrConn.Enabled = False
        cmdConnect_Click
        Exit Sub
    ElseIf (WskConnect.State = sckClosing) Then
        MsgBox "There was an error or the server maybe full." _
        & vbCrLf & "Please try agian latter.", vbInformation, "Server Error"
        TmrConn.Enabled = False
        cmdConnect_Click
        Exit Sub
    End If
    
End Sub

Private Sub txtHost_Change()
    cmdConnect.Enabled = Len(Trim(txtHost.Text)) <> 0 And Len(Trim(txtPort.Text)) <> 0
End Sub

Private Sub txtHost_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 32) Then KeyAscii = 0
End Sub

Private Sub txtPort_Change()
    txtHost_Change
End Sub

Private Sub txtPort_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 32) Then KeyAscii = 0
End Sub

Private Sub FillClientList(sData As String)
Dim vLst As Variant, x As Integer, sLine As String
    LstClients.Clear
    vLst = Split(sData, vbCrLf)
    'This will split the string sData by crlf and add the clients to the listbox
    For x = 0 To UBound(vLst)
        sLine = Trim(vLst(x))
        'Get and trim current line
        If Len(sLine) Then
            'Add to the listbox
            LstClients.AddItem sLine
        End If
    Next x
    
    'Clear up
    sLine = ""
    x = 0
    If Not IsEmpty(vLst) Then Erase vLst
    
End Sub

Private Sub WskConnect_DataArrival(ByVal bytesTotal As Long)
Dim n_pos As Integer, e_pos As Integer, Counter As Integer
Dim InData As String, InList As String

    'Process all the commands from the server
    WskConnect.GetData InData, , bytesTotal
    'Locate first chr(32)
    n_pos = InStr(InData, " ")
    
    If (n_pos <> 0) Then
        'Extract command from string
        sCmd = UCase$(Trim$(Left(InData, n_pos - 1)))
    End If
    
    'Process commands
    Select Case sCmd
        Case "RSN"
            'Remote server name
            e_pos = InStr(n_pos + 1, InData, vbCrLf, vbTextCompare)
            If (e_pos <> 0) Then
                lblServname.Caption = "Connected to host: " _
                & Trim(Mid(InData, n_pos + 1, e_pos - n_pos - 1))
                lblServname.Visible = True
            End If
            
        Case "SHD"
            'Server Shutdown
            e_pos = InStr(n_pos + 1, InData, vbCrLf, vbTextCompare)
            If (e_pos <> 0) Then
                lblServname.Caption = Trim(Mid(InData, n_pos + 1, e_pos - n_pos - 1))
                Call cmdConnect_Click
            End If
            
        Case "SMG"
            'Message from server
            e_pos = InStrRev(InData, vbCrLf, Len(InData), vbTextCompare)
            If (e_pos <> 0) Then
                MsgBox Trim(Mid(InData, n_pos + 1, e_pos - n_pos - 1)), vbInformation, "Message from Server"
            End If
            
        Case "CCL"
            'Client list sent from server
            e_pos = InStrRev(InData, vbCrLf, Len(InData), vbTextCompare)
            
            If (e_pos <> 0) Then
                InData = Mid(InData, n_pos + 1, e_pos - n_pos - 1)
                Call FillClientList(InData)
                InData = ""
            End If
        
        Case "CSM"
            'Message sent from server , that was sent from a client
            e_pos = InStrRev(InData, vbCrLf, Len(InData), vbTextCompare)
            If (e_pos <> 0) Then
                InData = Mid(InData, n_pos + 1, e_pos - n_pos - 1)
                MsgBox InData, vbInformation, "Message Recived"
            End If
            
        Case "UCL"
            'Update User List
            Call cmdRefresh_Click
            
        Case "CLOSE"
            'Close Client
            Call cmdConnect_Click
    End Select
    
    'Clear Buffers
    e_pos = 0
    n_pos = 0
    InData = ""
    sCmd = ""
End Sub

