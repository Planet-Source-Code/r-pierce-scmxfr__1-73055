VERSION 5.00
Begin VB.Form Options 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4860
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4620
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4860
   ScaleWidth      =   4620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00C0C0C0&
      Height          =   315
      ItemData        =   "Options.frx":0000
      Left            =   1230
      List            =   "Options.frx":0002
      TabIndex        =   13
      ToolTipText     =   "The ext. IP site the program uses"
      Top             =   2280
      Width           =   2325
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00808080&
      Caption         =   "User Selectable Options "
      ForeColor       =   &H0000FFFF&
      Height          =   1005
      Left            =   1230
      TabIndex        =   11
      Top             =   3750
      Width           =   2025
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   60
         MaxLength       =   2
         TabIndex        =   15
         Text            =   "05"
         Top             =   600
         Width           =   285
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00808080&
         Caption         =   "Display Ext. IP"
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   150
         TabIndex        =   12
         ToolTipText     =   "Obtain External IP upon program startup"
         Top             =   270
         Width           =   1335
      End
      Begin VB.Label Label6 
         BackColor       =   &H00808080&
         Caption         =   "Timeout in seconds"
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Left            =   420
         TabIndex        =   16
         ToolTipText     =   "Time External IP function will wait for URL to return data before canceling attempt"
         Top             =   630
         Width           =   1425
      End
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Height          =   315
      Left            =   1230
      MaxLength       =   4
      TabIndex        =   9
      ToolTipText     =   "The base port number"
      Top             =   1740
      Width           =   2025
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Height          =   315
      Left            =   1230
      TabIndex        =   7
      ToolTipText     =   "Your Handle"
      Top             =   1200
      Width           =   2025
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Height          =   315
      Left            =   1230
      TabIndex        =   5
      ToolTipText     =   "The dot address to connect to if your the master"
      Top             =   660
      Width           =   2025
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      Caption         =   "Chat Mode "
      ForeColor       =   &H0080FFFF&
      Height          =   1005
      Left            =   1230
      TabIndex        =   2
      Top             =   2640
      Width           =   2025
      Begin VB.OptionButton Option2 
         BackColor       =   &H00808080&
         Caption         =   "Slave"
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Top             =   570
         Width           =   795
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00808080&
         Caption         =   "Master"
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Top             =   270
         Width           =   825
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4290
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   30
      Width           =   315
   End
   Begin VB.Label Label5 
      BackColor       =   &H00808080&
      Caption         =   "Ext IP URL"
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Left            =   390
      TabIndex        =   14
      Top             =   2340
      Width           =   855
   End
   Begin VB.Label Label4 
      BackColor       =   &H00808080&
      Caption         =   "Port         :"
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Left            =   420
      TabIndex        =   10
      Top             =   1770
      Width           =   735
   End
   Begin VB.Label Label3 
      BackColor       =   &H00808080&
      Caption         =   "Nickname:"
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Left            =   390
      TabIndex        =   8
      Top             =   1230
      Width           =   765
   End
   Begin VB.Label Label2 
      BackColor       =   &H00808080&
      Caption         =   "Host Adr  :"
      ForeColor       =   &H00FFFF00&
      Height          =   225
      Left            =   390
      TabIndex        =   6
      Top             =   720
      Width           =   765
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Caption         =   "SCM's   INet   Settings"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   555
      Left            =   30
      TabIndex        =   1
      Top             =   60
      Width           =   4545
   End
End
Attribute VB_Name = "Options"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Check1_Click()
If Check1.Value = 1 Then
    Combo1.Enabled = True
Else
    Combo1.Enabled = False
End If
End Sub

Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim tmp0$: Dim x%
'46 is the delete key REMOVE ITEM DISPLAYED IN THE TEXT BOX of Combo Box
If KeyCode = 46 Then
tmp0$ = Combo1.Text
    For x% = 0 To Combo1.ListCount - 1
        If Combo1.List(x%) = tmp0$ Then
        Combo1.RemoveItem (x%)
        Exit For
        End If
    Next x%
End If
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
Dim x%: Dim Ret&
If KeyAscii = 13 Then
    If Combo1.ListCount = 10 Then
    Ret& = MsgBox("This exceeds the 10 entry maximum storage" + vbCr + vbCr + "Do you wish to overwrite the 1st entry ?", vbOKCancel, "10 entries maximum allowed")
    '1=ok 2=Cancel
    If Ret& = 1 Then
    Combo1.RemoveItem (0)
    Combo1.AddItem Combo1.Text
    End If
    Else
        If Combo1.Text <> "" Then
            For x% = 0 To Combo1.ListCount - 1
                If Combo1.List(x%) = Combo1.Text Then
                db0.urlidx = x%
                Exit Sub
                End If
            Next x%
            Combo1.AddItem Combo1.Text, Combo1.ListCount 'Add item to next available combo box location
        End If
    End If
End If


End Sub

Private Sub Command1_Click()
SCMxfr.Enabled = True
Unload Me
End Sub

Private Sub Form_Load()

If (db0.flag1 And 1) = 1 Then
    Option1.Value = True
Else
    Option2.Value = True
End If

If (db0.flag1 And 2) = 2 Then
    Check1.Value = 1 ' Obtain ext IP adr upon power-up
    Combo1.Enabled = True
Else
    Check1.Value = 0
    Combo1.Enabled = False
End If

Text1 = Trim(db0.dot) ' Host Adr
Text2 = Trim(db0.nicknam)
Text3 = Trim(Str(db0.port))
Text4 = Trim(Str(db0.tmeout))

For x% = 1 To 10
    If (Asc(db0.ipadr(x%)) <> 0 And Asc(db0.ipadr(x%)) <> 32) Then
        Combo1.AddItem Trim(db0.ipadr(x%)) ' Populate the combo box
    End If
Next x%

If db0.urlidx > 0 Then 'There was a selected item in the combo box
Combo1.Text = Trim(db0.ipadr(db0.urlidx)) ' display selected item
End If


End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim filenum As Integer: Dim x%: Dim y%
y% = 1

If Option1.Value = True Then
    If listening = False Then
    db0.flag1 = (db0.flag1 Or 1)
    SCMxfr.Label2 = "Master"
    SCMxfr.Command1.Caption = "Connect"
    SCMxfr.Label3.Enabled = True
    End If
Else
    If listening = False Then
    db0.flag1 = (db0.flag1 And 254)
    SCMxfr.Label2 = "Slave"
    SCMxfr.Command1.Caption = "Listen"
    SCMxfr.Label3.Enabled = False
    End If
End If

If Check1.Value = 1 Then
db0.flag1 = db0.flag1 Or 2
Else
db0.flag1 = (db0.flag1 And 253)
End If

db0.dot = Text1
SCMxfr.Label3 = Text1
db0.nicknam = Text2
db0.port = Val(Text3)
db0.tmeout = Val(Text4)
SCMxfr.Inet1.RequestTimeout = db0.tmeout
SCMxfr.Label6 = Trim(Str(db0.port))

If Combo1.ListIndex <> -1 Then
db0.urlidx = Combo1.ListIndex + 1 ' Save the selected combo box item (or -1 if not selected)
Else
    For x% = 0 To Combo1.ListCount - 1
        If Combo1.List(x%) = Combo1.Text Then db0.urlidx = x% + 1
    Next x%
End If

For x% = 0 To Combo1.ListCount - 1 ' Store the combo box entries into our array
db0.ipadr(x% + 1) = Combo1.List(x%)
Next x%
'x% = 1
clrlp: ' Clear out the remaining entries from our array
If x% < 10 Then
db0.ipadr(x% + 1) = "": x% = x% + 1
GoTo clrlp
End If

filenum = FreeFile
Open App.Path + "\xfr.db" For Random As filenum Len = datasize
Put #filenum, 1, db0
Close #filenum

End Sub

Private Sub Option1_Click()
If connected = True Then
    If db0.flag1 And 1 = 1 Then
        Option1.Value = True
    Else
        Option2.Value = True
    End If
End If
End Sub

Private Sub Option2_Click()
If connected = True Then
    If db0.flag1 And 1 = 1 Then
        Option1.Value = True
    Else
        Option2.Value = True
    End If
End If
End Sub

Private Sub Text1_Change()
If connected = True Then
Text1 = Trim(db0.dot)
End If
End Sub

Private Sub Text3_Change()
If (connected = True Or listening = True) Then
Text3 = Trim(Str(db0.port))
End If
End Sub

Public Sub fge()

End Sub
