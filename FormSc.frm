VERSION 5.00
Begin VB.Form FormSc 
   BackColor       =   &H00C00000&
   Caption         =   "ECG  Simulation Test (32-bit program) "
   ClientHeight    =   5715
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12600
   Icon            =   "FormSc.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5715
   ScaleWidth      =   12600
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdExit 
      BackColor       =   &H00FF8080&
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2760
      Width           =   3255
   End
   Begin VB.CommandButton CmdNew 
      BackColor       =   &H00FF8080&
      Caption         =   "NEW"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2760
      Width           =   3255
   End
   Begin VB.CommandButton CmdScore 
      BackColor       =   &H00FF8080&
      Caption         =   "SCORE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2760
      Width           =   3255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C00000&
      BackStyle       =   0  'Transparent
      Caption         =   $"FormSc.frx":08CA
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   975
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   11175
   End
End
Attribute VB_Name = "FormSc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub CmdExit_Click()
'Reinitlaize all score variables

cnt = 0
Score = 0
extra = 0
converta = 0
convertb = 0

'Unload all forms and close the program
Unload FormQ1
Unload FormQ2
Unload FormQ3
Unload FormQ4
Unload FormQ5
Unload FormQ6
Unload FormQ7
Unload FormQ8
Unload FormQ9
Unload FormQ10
Unload FormQ11
Unload FormQ12
Unload FormQ13
Unload FormQ14
Unload FormQ15
Unload FormQ16
Unload FormQ17
Unload FormQ18
Unload FormQ19
Unload FormQ20
Unload FormQ21
Unload FormQ22
Unload FormQ23
Unload FormQ24
Unload FormQ25
Unload FormQ26
Unload FormQ27
Unload FormStart
Unload FormSc
End
End Sub
Private Sub CmdNew_Click()
'Reinitialize all score variables
cnt = 0
Score = 0
extra = 0
converta = 0
convertb = 0

'Unload Quetion forms and This form, and display FormStart
Unload FormQ1
Unload FormQ2
Unload FormQ3
Unload FormQ4
Unload FormQ5
Unload FormQ6
Unload FormQ7
Unload FormQ8
Unload FormQ9
Unload FormQ10
Unload FormQ11
Unload FormQ12
Unload FormQ13
Unload FormQ14
Unload FormQ15
Unload FormQ16
Unload FormQ17
Unload FormQ18
Unload FormQ19
Unload FormQ20
Unload FormQ21
Unload FormQ22
Unload FormQ23
Unload FormQ24
Unload FormQ25
Unload FormQ26
Unload FormQ27
FormStart.Show
Unload FormSc
End Sub

Private Sub CmdScore_Click()

'Score the test. cnt questions are full value regualar questions. extra questions are the two extra-credit questions
'worth half-value. Score Variable will be in terms of per cent.

converta = cnt * 4
convertb = extra * 2
If converta = 100 Then
Score = converta
Else
Score = converta + convertb
End If


'Use IF/Then expressions to display custom score in a message box.

'Validate Score as a valid percent. If the value of score is Not between 0 and 100, then use
'message box to display error message.

If Score > 100 Then
MsgBox "There has been an error, and your test could not be scored. Please take the test again.", vbOKOnly, Error
End If
If Score < 0 Then
MsgBox "There has been an error, and your test could not be scored. Please take the test again.", vbOKOnly, Error
End If

' Display one version of score message if score is LESS than 100 percent.
If Score < 100 Then
MsgBox "Your test score is" & Score & " %. You have room to improve your score, and study more and try again!", vbOKOnly, Score
Else
'Display Second version if score is 100 percent.
MsgBox "Your test score is" & Score & " %. You have mastered this test. Great Job!!!", vbOKOnly, Score
End If
End Sub

Private Sub Form_Load()
With FormSc
        .Top = (Screen.Height - .Height) / 2
        .Left = (Screen.Width - .Width) / 2
    End With
End Sub
