VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2190
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3780
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   3780
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "CANCEL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1020
      Width           =   1245
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "YES"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   450
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1020
      Width           =   1245
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00C0C0C0&
      X1              =   2490
      X2              =   3750
      Y1              =   1770
      Y2              =   1770
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      X1              =   1260
      X2              =   0
      Y1              =   1770
      Y2              =   1770
   End
   Begin VB.Label Label4 
      BackColor       =   &H00808080&
      Caption         =   "Download Path"
      ForeColor       =   &H00FFFFC0&
      Height          =   195
      Left            =   1320
      TabIndex        =   5
      Top             =   1680
      Width           =   1125
   End
   Begin VB.Label Label3 
      BackColor       =   &H00808080&
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   225
      Left            =   30
      TabIndex        =   4
      Top             =   1920
      Width           =   3705
   End
   Begin VB.Label Label2 
      BackColor       =   &H00808080&
      Caption         =   "Start The File Download?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   255
      Left            =   540
      TabIndex        =   3
      Top             =   510
      Width           =   2745
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF0000&
      Caption         =   " File Download Request"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   30
      Width           =   3765
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()

Unload Me
End Sub

Private Sub Form_Load()
Label3 = Trim(db0.defsave) 'display the save file directory
End Sub

