VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form SCMxfr 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SCMxfr"
   ClientHeight    =   6945
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   9000
   Icon            =   "SCMxfr.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6945
   ScaleWidth      =   9000
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   5280
      Top             =   4950
   End
   Begin VB.PictureBox ctlProgress1 
      FillColor       =   &H00FF0000&
      Height          =   225
      Left            =   3840
      ScaleHeight     =   165
      ScaleWidth      =   2445
      TabIndex        =   13
      Top             =   6660
      Visible         =   0   'False
      Width           =   2505
   End
   Begin MSWinsockLib.Winsock Winsock2 
      Left            =   4530
      Top             =   5010
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.DirListBox Dir1 
      BackColor       =   &H00E0E0E0&
      Height          =   5940
      Left            =   30
      TabIndex        =   12
      Top             =   60
      Visible         =   0   'False
      Width           =   6345
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4290
      Top             =   3450
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5925
      Left            =   6480
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      Top             =   60
      Width           =   2505
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      Height          =   315
      Left            =   60
      TabIndex        =   8
      ToolTipText     =   "Type the text to send  -  Then press Send -or- the <ENTER> key"
      Top             =   6030
      Width           =   6285
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Command1"
      Height          =   315
      Left            =   6510
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6030
      Width           =   2205
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   5070
      Top             =   1230
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   15000
      Left            =   5010
      Top             =   1890
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   3120
      Top             =   2250
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      RequestTimeout  =   5
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00E0E0E0&
      Height          =   5910
      Left            =   30
      TabIndex        =   7
      Top             =   60
      Width           =   6375
   End
   Begin VB.Label Label7 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000C0&
      Height          =   225
      Left            =   150
      TabIndex        =   9
      Top             =   6660
      Width           =   3645
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3210
      TabIndex        =   5
      ToolTipText     =   "The Port communication uses"
      Top             =   6360
      Width           =   555
   End
   Begin VB.Label Label5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Port:"
      Height          =   225
      Left            =   2850
      TabIndex        =   4
      Top             =   6390
      Width           =   345
   End
   Begin VB.Label Label4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Host Address:"
      Height          =   225
      Left            =   210
      TabIndex        =   3
      Top             =   6390
      Width           =   1005
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1230
      TabIndex        =   2
      ToolTipText     =   "The address the Master will communicate with."
      Top             =   6360
      Width           =   1605
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   4020
      TabIndex        =   1
      ToolTipText     =   "Your current mode.  (server/client)"
      Top             =   6360
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "12:12"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   8310
      TabIndex        =   0
      Top             =   6690
      Width           =   615
   End
   Begin VB.Label Label8 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   6510
      TabIndex        =   11
      Top             =   6390
      Width           =   2445
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuLoad 
         Caption         =   "Load"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuOpt 
      Caption         =   "Options"
   End
   Begin VB.Menu mnuTest 
      Caption         =   "Test"
   End
End
Attribute VB_Name = "SCMxfr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public tp$
Public gFileName As String
Public gfilesize As Long
Public tempo As Boolean
Public gSize As Double
Public gFile As File
Public gFSO As New Scripting.FileSystemObject
Public glfn As Integer
Public gcursize As Long





Private Sub Command1_Click()

If listening = True Then ' Were listening and user has pressed "Close" command
    If Winsock1.State = 2 Then
        Winsock1.Close: Winsock2.Close
        listening = False: connected = False
        Command1.Caption = "Listen"
        UpDate ("Socket Closed")
        Exit Sub
    End If
End If
'--------------------------------------------------------------------------------------
If (db0.flag1 And 1) = 1 Then ' were in master mode
    If (Command1.Caption = "Close") Then
        Winsock1.Close: Winsock2.Close
        Command1.Caption = "Connect"
        connected = False
        UpDate ("Closed as requested")
        Exit Sub
    End If
Winsock1.Close: Winsock2.Close
DoEvents
UpDate ("Attempting 2 Connect")
On Error GoTo ziper
Winsock1.Connect Trim(db0.dot), db0.port
DoEvents
Winsock2.Connect Trim(db0.dot), db0.port + 1
On Error GoTo 0
Command1.Caption = "Close"
Else ' were in slave mode
    If Winsock1.State = 0 Then ' The socket is closed
        Winsock1.LocalPort = db0.port: Winsock2.LocalPort = db0.port + 1
        Winsock1.Listen: Winsock2.Listen
        UpDate ("Listening-Base Port:" & Str(db0.port))
        Command1.Caption = "Close"
        listening = True: tempo = True
    Else ' This tells us that the socket isn't closed
        If Winsock1.State = 7 Then ' connected
        Winsock1.Close: Winsock2.Close
        UpDate ("Socket Closed")
        Command1.Caption = "Listen"
        listening = False: connected = False
        Exit Sub
        End If
        UpDate ("Socket State =" & Str(Winsock1.State))
        Winsock1.Close: Winsock2.Close
        DoEvents
        UpDate ("Socket State =" & Str(Winsock1.State))
    End If
End If
Exit Sub
ziper:
g& = MsgBox("Winsock has returned an error" + vbCr + Err.Description, vbCritical, "Error Reported")

End Sub

Private Sub Command2_Click()
End Sub

Private Sub Dir1_LostFocus()

Dir1.Visible = False
db0.defsave = Dir1.Path
Label7 = Trim(db0.defsave)

End Sub

Private Sub Dir1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
Dir1.Visible = False
db0.defsave = Dir1.Path
Label7 = Trim(db0.defsave)
End If
End Sub

Private Sub Form_Load()
Dim filenum As Integer: Dim filesize As Long: Dim siteurl As String

SCMxfr.Caption = "    INet Chat / Xfr v" + Trim(Str(App.Major)) + "." + Trim(Str(App.Minor)) + "         " & Date$ & "          Start: " & Left$(Time$, 5)
Label1 = Left$(Time$, 5)
On Error GoTo nofile
filesize = FileLen(App.Path + "\xfr.db")
filenum = FreeFile
Open App.Path + "\xfr.db" For Random As filenum Len = datasize
Get #filenum, 1, db0
Close #filenum

If (db0.flag1 And 1) = 1 Then
    Label2 = "Master"
    Command1.Caption = "Connect"
Else
    Label2 = "Slave"
    Command1.Caption = "Listen"
    Label3.Enabled = False
End If
Label3 = Trim(db0.dot)
Label6 = Trim(Str(db0.port))
Text2 = "            " & Date$ & " " & Time$
Inet1.RequestTimeout = db0.tmeout
If (db0.flag1 And 2) = 2 Then ' Check for external IP address upon starting program
End If

'connected = True '        For test purposes
If (db0.flag1 And 2) = 2 Then ' obtain external IP adr is checked

If db0.urlidx > 0 Then 'There is a selected item in the URL combo box
siteurl = Trim(db0.ipadr(db0.urlidx)) ' selected URL
On Error GoTo weberr
tp$ = Inet1.OpenURL(siteurl) ' download the page into tp$
DoEvents
On Error GoTo 0
tp$ = "  " & tp$ & "  " 'pad the buffer to prevent errors
ParseURL
Label8 = Label8 + vbCr + siteurl + vbCr + " Responce Code:" + Str(Inet1.ResponseCode)
End If
End If
'App.Title = "SCMxfr       V" + Str(App.Major) + "      build" + Str(App.Minor)
Label7 = Trim(db0.defsave)
Exit Sub '                 This is the normal exit if db file exists
nofile:
filenum = FreeFile
Open App.Path + "\xfr.db" For Random As filenum Len = datasize
Put #filenum, 1, db0
Close #filenum
Label7 = "**********"
Exit Sub
weberr:
Resume
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim filenum As Integer
If Winsock1.State > 0 Then
Cancel = 1
j& = MsgBox("Socket is currently open !" & vbCrLf & "     Socket State = " & Str(Winsock1.State), vbInformation, "   Close Socket B4 Exiting Program !")
Exit Sub
End If

filenum = FreeFile
Open App.Path + "\xfr.db" For Random As filenum Len = datasize
Put #filenum, 1, db0
Close #filenum


End Sub

Private Sub Label1_Click()
'Static r%
'ctlProgress1.Min = 0: ctlProgress1.Max = 20
'r% = r% + 1
'ctlProgress1.Value = r%
'ctlProgress1.Visible = True
End Sub

Private Sub Label4_Click()

glfn = FreeFile
Open gFileName For Binary As #glfn
'ctlProgress1.Visible = True
End Sub

Private Sub mnuExit_Click()
Unload Me
End Sub

Private Sub mnuLoad_Click()
On Error GoTo Cancel
If connected = True Then
CommonDialog1.Flags = &H4 Or &H200 Or &H1000 Or &H200000
CommonDialog1.DialogTitle = "Select File To Transfer"
CommonDialog1.CancelError = True
CommonDialog1.Filter = "All Files (*.*)|*.*"
CommonDialog1.ShowOpen
gFileName = CommonDialog1.FileName 'User has selected a download file
Set gFile = gFSO.GetFile(gFileName)
gfln = FreeFile
Open gFileName For Append As #gfln
gSize = LOF(1) \ 1024
Close #gfln
'ctlProgress1.Min = 0
'ctlProgress1.Max = gSize
'ctlProgress1.Value = 0
End If

Cancel:
End Sub

Private Sub mnuOpt_Click()

Dim frmChild As New Options

frmChild.Show vbModeless, Me
SCMxfr.Enabled = False

End Sub

Private Sub mnuSave_Click()
Dir1.Path = Trim(db0.defsave)
Dir1.Visible = True
Label7 = "Select download directory and then right-click to close"
End Sub

Private Sub mnuTest_Click()

Form1.Show

End Sub

Private Sub Text1_GotFocus()
Text1.BackColor = &HFFFF00
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
Dim tmp$

If KeyAscii = 13 And connected = True Then
KeyAscii = 0
tmp$ = "(" & Trim(db0.nicknam) & "): " & Text1
List1.AddItem tmp$
Winsock1.SendData tmp$
Text1 = ""
End If

End Sub

Private Sub Text1_LostFocus()
Text1.BackColor = &HE0E0E0
End Sub

Private Sub Timer1_Timer()
Label1 = Left$(Time$, 5)
End Sub

Public Sub UpDate(txt As String)
Text2 = Text2 & vbCrLf & " " & Left$(Time$, 5) & "   " & txt
End Sub

Public Sub ParseURL()
Dim t%:  Dim w%: Dim x%: Dim y%: Dim z%: Dim ck As Boolean

x% = 1
deloop:
x% = InStr(x%, tp$, ".") ' If x%>0 then pointing to 1st occurence of "."
If x% > 0 Then 'we have detected a period
GoSub ckdot: If ck = False Then GoTo continue
Label8 = "External IP Address:  " & Mid$(tp$, y%, w% - y%)
Exit Sub
End If
Label8 = "External IP Dot Adr. Not Found"
Exit Sub

ckdot:
ck = False
For y% = x% - 1 To x% - 4 Step -1
If y% < 1 Then Return
z% = Asc(Mid$(tp$, y%, 1))
If Not (z% > 47 And z% < 58) Then Exit For
Next y%
If y% = x% - 1 Then Return
If y% = x% - 5 Then Return
y% = y% + 1 '--------------------> y% now points to start of adr. string
z% = x% + 1 '--------------------> Point z% to char past 1st period
GoSub ckit: If ck = False Then Return
ck = False
GoSub sample: If t% <> 46 Then Return
z% = w% + 1 ' Point z% to char past 2nd period
GoSub ckit: If ck = False Then Return
ck = False
GoSub sample: If t% <> 46 Then Return
z% = w% + 1
GoSub ckit
GoSub sample: If (t% > 47 And t% < 58) Then ck = False ' last set of digits is too long
Return ' w%=last char of address +1

ckit: 'parse the string to see if it is a valid dot address  ie   nnn.nnn.nnn.nnn
For w% = z% To z% + 2
If w% > Len(tp$) Then Return
t% = Asc(Mid$(tp$, w%, 1))
If Not (t% > 47 And t% < 58) Then Exit For
ck = True
Next w%
Return

sample:
t% = Asc(Mid$(tp$, w%, 1))
Return

continue:
If x% < Len(tp$) - 7 Then
    x% = x% + 1
    GoTo deloop
End If
Label8 = "Adr Not Found"

End Sub

Private Sub Winsock1_Close()

UpDate ("Socket Closed")
Winsock1.Close
listening = False
connected = False
If (db0.flag1 And 1) = 1 Then 'master
Command1.Caption = "Connect"
Else ' slave
Command1.Caption = "Listen"
End If

End Sub

Private Sub Winsock1_Connect()
UpDate ("Socket Connected")
connected = True
End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
If listening = True Then
If Winsock1.State <> 0 Then Winsock1.Close
DoEvents
Winsock1.Accept (requestID)
DoEvents
    If Winsock1.State = 7 Then
        UpDate ("Socket Connected:" + Str(requestID))
        connected = True
    End If
End If
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim chatstuff As String

Winsock1.GetData chatstuff
DoEvents
List1.AddItem chatstuff

End Sub

Private Sub Winsock2_ConnectionRequest(ByVal requestID As Long)

If listening = True Then
If Winsock2.State <> 0 Then Winsock1.Close
DoEvents
Winsock2.Accept (requestID)
End If

End Sub

