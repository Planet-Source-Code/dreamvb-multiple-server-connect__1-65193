VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmmain 
   Caption         =   "Multiple Connections"
   ClientHeight    =   4320
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   6150
   LinkTopic       =   "Form1"
   ScaleHeight     =   4320
   ScaleWidth      =   6150
   StartUpPosition =   2  'CenterScreen
   Begin MSWinsockLib.Winsock WskServ 
      Index           =   0
      Left            =   225
      Top             =   1590
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar StatBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   3
      Top             =   4050
      Width           =   6150
      _ExtentX        =   10848
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2700
      Top             =   180
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView LstClients 
      Height          =   750
      Left            =   30
      TabIndex        =   0
      Top             =   465
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   1323
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Remote Addr"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Computer Name"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label lblClients 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   1500
      TabIndex        =   2
      Top             =   180
      Width           =   90
   End
   Begin VB.Label lblConnections 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Connected Clients"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   105
      TabIndex        =   1
      Top             =   180
      Width           =   1305
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000D&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00800000&
      Height          =   330
      Left            =   45
      Top             =   120
      Width           =   2400
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   1
      X1              =   -60
      X2              =   840
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   0
      X2              =   900
      Y1              =   15
      Y2              =   15
   End
   Begin VB.Menu mnuServ 
      Caption         =   "&Server"
      Begin VB.Menu mnuStart 
         Caption         =   "&Start"
      End
      Begin VB.Menu mnuStop 
         Caption         =   "Stop"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnubalnk2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuconfig 
         Caption         =   "&Config"
      End
      Begin VB.Menu mnublank1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuClients 
      Caption         =   "&Clients"
      Begin VB.Menu msnuClientMsg 
         Caption         =   "&Send Client Message"
      End
      Begin VB.Menu mnuMsgAll 
         Caption         =   "Send Message to all Clients"
      End
      Begin VB.Menu mnubalnk3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCloeClient 
         Caption         =   "Close Client"
      End
      Begin VB.Menu mnuCloseAll 
         Caption         =   "Close All Clients"
      End
   End
   Begin VB.Menu mnuabout 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private uSocks As Long          'Total number of sockets created
Private uClients As Long        'Total number of connected clients
Private ClientName As String    'Client connected name
Private Client_Idx As Integer   'Index of selected client in listview control

Sub SendClientMsg(vMsg As String)
On Error Resume Next
Dim x As Integer
    'This sub sends out a message to all connected clients to the server
    For x = 1 To WskServ.Count
        'Check the state of each winsock control
        If WskServ(x).State = sckConnected Then
            'If a client is connected send the message in vMsg
            WskServ(x).SendData vMsg & vbCrLf
        End If
        'Allow other things to process
        DoEvents
    Next
End Sub

Sub ProcExit()
    'Clean up program
    'Send a message to all clients informaing the server is shutting down
    Call SendClientMsg("SHD Server closed Shutdown.")
    'Close and unload all winsock controls
    Call DestroyConnections
    'Reset variables
    Client_Idx = 0
    DoEvents
    'Unload main form
    Unload frmmain
End Sub

Function FindItem(sIp As String) As Integer
On Error Resume Next
Dim mItem As ListItem
    If (LstClients.ListItems.Count = 0) Or Len(sIp) = 0 Then Exit Function
    Set mItem = LstClients.FindItem(sIp)
    FindItem = mItem.Index
End Function

Private Sub RemoveClientFromLst(sIp As String)
Dim Idx As Integer
    Idx = FindItem(sIp)
    If (Idx <> 0) Then LstClients.ListItems.Remove Idx
End Sub

Private Sub AddClient(Index)
On Error Resume Next
Dim sIp As String, Idx As Integer
    'Get and store remote IP
    sIp = WskServ(Index).RemoteHostIP
    'Add the Client addr
    LstClients.ListItems.Add , sIp & ":" & Index, sIp, 1, 1
    'Add the clients computer name
    Idx = FindItem(sIp)
    If (Idx = 0) Then Exit Sub
    LstClients.ListItems(Idx).SubItems(1) = ClientName
    ClientName = ""
End Sub

Private Sub UpdateStat()
    lblClients.Caption = uClients
End Sub

Private Sub Abort(ab_code)
    If (ab_code = sckAddressInUse) Then
        MsgBox "The selected address is already in use.", vbCritical, "Error_" & Err.Number
    Else
        MsgBox Err.Description, vbCritical, "Error_" & Err.Number
    End If
End Sub

Private Sub DestroyConnections()
Dim x As Integer
    For x = 1 To WskServ.Count - 1
        If WskServ(x).State = sckConnected Then
            'Send a message to inform the user the server is closeing down
            WskServ(x).SendData "SHD Server is closeing down." & vbCrLf
        End If
        WskServ(x).Close
        Unload WskServ(x)
    Next x
    
    uSocks = 0
    uClients = 0
    
    WskServ(0).Close
    x = 0
End Sub

Private Sub ServerControl(iStart As Boolean)
On Error GoTo ErrFlag:
    
    mnuStart.Enabled = Not iStart
    mnuStop.Enabled = iStart
    
    If (iStart) Then
        Call DestroyConnections
        'Tell the server witch port to listen on
        Call UpdateStat
        'Above just Resets the client counter
        WskServ(0).LocalPort = TServCfg.ServPort
        'Tell the server to start listening for connections
        WskServ(0).Listen
        StatBar.Panels(1).Text = "Status: Started"
    Else
        'Clean up
        StatBar.Panels(1).Text = "Status: Stopped"
        Call DestroyConnections
        Call UpdateStat
    End If
    
    Exit Sub
ErrFlag:
    'Looks like an error call error handler
   Abort Err.Number
   'Rset server control
   Call ServerControl(False)
End Sub

Function GetClientList(ReqIp As String) As String
Dim sBuffer As String, x As Integer
    'Build a list of all clients connected to the server
    'ReqIp is the client requesting the list, so we don;t need to send
    'ReqIp own IP address
    Do While (x < LstClients.ListItems.Count)
        x = x + 1
        If LstClients.ListItems(x).Text <> ReqIp Then
            sBuffer = sBuffer & LstClients.ListItems(x).Key & vbCrLf
        End If
    Loop
    GetClientList = sBuffer
    x = 0
    sBuffer = ""
End Function

Private Sub Form_Load()
    'Load server config
    Call LoadCfg
    'Set up the protocol for this server TCP
    WskServ(0).Protocol = sckTCPProtocol
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    ProcExit
End Sub

Private Sub Form_Resize()
On Error Resume Next
    Line1(0).X2 = frmmain.ScaleWidth
    Line1(1).X2 = Line1(0).X2
    
    LstClients.Width = (frmmain.ScaleWidth - LstClients.Left)
    LstClients.Height = (frmmain.ScaleHeight - LstClients.Top) - StatBar.Height
    
    If Err Then
        frmmain.Width = 3240
        frmmain.Height = 3150
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set FrmCfg = Nothing
    Set frmmain = Nothing
End Sub

Private Sub LstClients_ItemClick(ByVal Item As MSComctlLib.ListItem)
Dim cKey As String, e_pos As Integer
    If LstClients.ListItems.Count = 0 Then Exit Sub
    'Get the Clients Key
    cKey = LstClients.SelectedItem.Key
    'Locate : in the key, after : is found the Client index will follow
    e_pos = InStr(1, cKey, ":", vbBinaryCompare)
    'Extact the Clients Index of the Winsock control
    If (e_pos <> 0) Then Client_Idx = Val(Mid(cKey, e_pos + 1, Len(cKey)))
    
End Sub

Private Sub mnuabout_Click()
    MsgBox "Client and server example." _
    & vbCrLf & vbTab & "By Ben Jones", vbInformation, "About.."
End Sub

Private Sub mnuCloeClient_Click()
Dim sMsg As Integer
    If (Client_Idx <> 0) Then
        sMsg = MsgBox("Are you sure you want to close the selected client.", vbYesNo Or vbQuestion, "Close Client")
        If (sMsg = vbNo) Then Exit Sub
        'Send the command to close down the client
        WskServ(Client_Idx).SendData "CLOSE " & vbCrLf
        DoEvents
    End If
End Sub

Private Sub mnuCloseAll_Click()
Dim sMsg As Integer
    sMsg = MsgBox("Are you sure you want to close all the clients", vbYesNo Or vbQuestion, "Close All Clients")
    If (sMsg = vbNo) Then Exit Sub
    'Send the command to close down all the clients
    Call SendClientMsg("CLOSE ")
End Sub

Private Sub mnuconfig_Click()
    FrmCfg.Show vbModal, frmmain
End Sub

Private Sub mnuexit_Click()
    ProcExit
End Sub

Private Sub mnuMsgAll_Click()
Dim sMsg As String
    'Promp user to enter a message
    sMsg = Trim(InputBox("Enter a message to send to the client: ", "Message Client"))
    'Do not send the message if it empty
    If Len(sMsg) = 0 Then Exit Sub
    'Send the message to all clients connected
    Call SendClientMsg("SMG Message From Server: " & vbCrLf & sMsg)
    'Clear up
    sMsg = ""
End Sub

Private Sub mnuStart_Click()
    ServerControl True
End Sub

Private Sub mnuStop_Click()
    ServerControl False
End Sub

Private Sub msnuClientMsg_Click()
Dim sMsg As String
    If (Client_Idx <> 0) Then
        'Promp user to enter a message
        sMsg = Trim(InputBox("Enter a message to send to the client: ", "Message Client"))
        'Do not send the message if it empty
        If Len(sMsg) = 0 Then Exit Sub
        'Send the message to the client
        WskServ(Client_Idx).SendData "SMG Message From Server: " & vbCrLf & sMsg & vbCrLf
        DoEvents
        sMsg = ""
    End If
End Sub

Private Sub WskServ_Close(Index As Integer)
    'Remove the Client from the listview
    Call RemoveClientFromLst(WskServ(Index).RemoteHostIP)
    'Clsoe the Clients connection
    WskServ(Index).Close
    'Update Client counter
    uClients = uClients - 1
    'Update Client lables
    Call UpdateStat
    'Now inform all clients to update the client lists
    SendClientMsg "UCL " & vbCrLf
End Sub

Function FreeSocket() As Integer
Dim x As Integer
    'This function checks for a free socket
    'If one is found to be closed then we will return the index
    'The point of this was for only allow the server to
    'Create new winsock controls when needed
    FreeSocket = 0
    For x = 0 To WskServ.Count - 1
        'Check the state of each control
        If WskServ(x).State = 0 Then
            'Return the free socket controls index
            FreeSocket = x
            Exit Function
        End If
    Next x
End Function

Private Sub WskServ_ConnectionRequest(Index As Integer, ByVal requestID As Long)
Dim mFreeSock As Integer

    'Check if we have exceeded our max connections
    'If we have do nothing
    If (uClients >= TServCfg.ServMaxConn) Then Exit Sub
    
    'See if we have a FreeSocket
    mFreeSock = FreeSocket
    
    If (mFreeSock <> 0) Then
        'Ok we have a free socket so we use it
        WskServ(mFreeSock).Accept requestID
        'Just jump to update client count
        'As we have no need to load a new winsock
        GoTo Here:
    End If

    'Update socket counter
    uSocks = uSocks + 1
    'Load up new Winsock control
    Load WskServ(uSocks)
    'Accept the remote request
    WskServ(uSocks).Accept requestID
    
Here:
    'Used to update the client counter label
    'If <= 0 reset uClients to zero
    If (uClients <= 0) Then uClients = 0
    'Update uClients counter
    uClients = uClients + 1
    'Update Client counter label
    Call UpdateStat
    
End Sub

Private Sub WskServ_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim InData As String, Temp As String
Dim e_pos As Integer, n_pos As Integer, d_pos As Integer
    
    'Here is were we process all the messages sent to the server
    'Get the incomming data
    WskServ(Index).GetData InData, , bytesTotal
    
    'Locate the first chr(32) in the string
    n_pos = InStr(1, InData, " ", vbBinaryCompare)
    'If above is a match extract the command commands look like Command Param
    If (n_pos <> 0) Then
        'Get command
        sCmd = UCase$(Trim$(Left(InData, n_pos - 1)))
    End If
    
    Select Case sCmd
        Case "CPN" 'Client Computername
            'Extract and store client name
            e_pos = InStr(n_pos + 1, InData, vbCrLf, vbTextCompare)
            If (e_pos <> 0) Then
                ClientName = Trim(Mid(InData, n_pos + 1, e_pos - n_pos - 1))
            End If
            'Now slow the client in the listview
            Call AddClient(Index)
            'Now send the Client the servers name
            WskServ(Index).SendData "RSN " & TServCfg.ServName & vbCrLf
            DoEvents
        'Client has requested client connected list
        Case "CCL"
            e_pos = InStr(n_pos + 1, InData, vbCrLf, vbTextCompare)
            If (e_pos <> 0) Then
                'Send the client the requested list
                WskServ(Index).SendData "CCL " & GetClientList(WskServ(Index).RemoteHostIP) & vbCrLf
                DoEvents
            End If
        Case "CSM"
            'Client send message
            n_pos = InStr(n_pos + 1, InData, " ", vbBinaryCompare)
            If (n_pos <> 0) Then Temp = Trim(Left(InData, n_pos - 1))
            d_pos = InStrRev(Temp, " ", Len(Temp), vbBinaryCompare)
            'Extract Clients Index
            If (d_pos <> 0) Then Temp = Trim(Mid(Temp, d_pos, Len(Temp)))
            'Now Extract the message
            e_pos = InStrRev(InData, vbCrLf, Len(InData), vbTextCompare)
            If (e_pos <> 0) Then
                InData = Trim(Mid(InData, n_pos + 1, e_pos - n_pos - 1))
                'Send the message to the client using Temp as the index
                WskServ(Val(Temp)).SendData "CSM " & InData & vbCrLf
            End If

    End Select
    
    'Clear Buffers
    e_pos = 0
    n_pos = 0
    InData = ""
    sCmd = ""
    Temp = ""
End Sub

