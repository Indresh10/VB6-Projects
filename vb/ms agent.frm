VERSION 5.00
Object = "{1147E550-A208-11DE-ABF2-002421116FB2}#1.1#0"; "DaControl.dll"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7965
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   7965
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Do it"
      Height          =   615
      Left            =   1320
      TabIndex        =   4
      Top             =   1920
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   1320
      Width           =   2295
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1920
      TabIndex        =   0
      Top             =   840
      Width           =   2295
   End
   Begin DoubleAgentCtl.DaControl DaControl1 
      Left            =   4800
      Top             =   840
      _Version        =   257
      _ExtentX        =   1508
      _ExtentY        =   1508
      RaiseRequestErrors=   -1  'True
      AutoConnect     =   32
      AutoSize        =   0   'False
      BackColor       =   -2147483643
      BorderColor     =   -2147483640
      BorderStyle     =   1
      BorderVisible   =   -1  'True
      BorderWidth     =   1
      MousePointer    =   0
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Enter Words to speak"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   1545
   End
   Begin VB.Label Label1 
      Caption         =   "Select Animation"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Set merlin = DaControl1.Characters("Merlin")
   If Combo1.Text = "" Then                ' If the combo1 ComboBox has
                                              ' nothing selected
      p1 = MsgBox("No Animation Selected", vbExclamation)       ' Give an error
   Else  ' If something has been chosen in the ComboBox
          merlin.Play Combo1.Text
   End If

   If Text1.Text = "" Then    ' If no text has been entered into the TextBox
      p2 = MsgBox("No Text Entered to speak", vbExclamation) ' Give an error
   Else                        ' If text has been entered
      merlin.Speak Text1.Text ' Have Merlin say whatever is written in the TextBox
   End If
End Sub

Private Sub Form_Load()
Dim CharPath As String

' Activate Merlin
CharPath = "C:\Windows\Agent" ' The path is the root directory for your project

' Load the characters specified in path
DaControl1.Characters.Load "Merlin", CharPath & "\Merlin.acs"
Set merlin = DaControl1.Characters("Merlin")

' Display Merlin on the screen
merlin.Show
' Add the Character combo1s to the ComboBox
Combo1.AddItem ("Acknowledge")
Combo1.AddItem ("Announce")
Combo1.AddItem ("Blink")
Combo1.AddItem ("Congratulate")
Combo1.AddItem ("DoMagic1")
Combo1.AddItem ("DoMagic2")
Combo1.AddItem ("Explain")
Combo1.AddItem ("GestureDown")
Combo1.AddItem ("GestureLeft")
Combo1.AddItem ("GestureRight")
Combo1.AddItem ("GetAttention")
Combo1.AddItem ("LookUpBlink")
Combo1.AddItem ("MoveDown")
Combo1.AddItem ("MoveLeft")
Combo1.AddItem ("MoveRight")
Combo1.AddItem ("MoveUp")
Combo1.AddItem ("Pleased")
Combo1.AddItem ("Process")
Combo1.AddItem ("Read")
Combo1.AddItem ("Sad")
Combo1.AddItem ("Search")
Combo1.AddItem ("Show")
Combo1.AddItem ("Think")
Combo1.AddItem ("Wave")
Combo1.AddItem ("Write")
End Sub
