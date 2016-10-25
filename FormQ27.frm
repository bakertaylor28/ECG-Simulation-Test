VERSION 5.00
Begin VB.Form FormQ27 
   BackColor       =   &H00C00000&
   Caption         =   "ECG  Simulation Test (32-bit program) "
   ClientHeight    =   8235
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14790
   Icon            =   "FormQ27.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8235
   ScaleWidth      =   14790
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdNext 
      BackColor       =   &H00FF8080&
      Caption         =   "NEXT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   11400
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5400
      Width           =   2775
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00800000&
      Height          =   3015
      Left            =   600
      TabIndex        =   1
      Top             =   4440
      Width           =   10335
      Begin VB.OptionButton Option3 
         BackColor       =   &H00800000&
         Caption         =   "C. Paced Ventricular"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   615
         Left            =   360
         TabIndex        =   5
         Top             =   1920
         Width           =   9495
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00800000&
         Caption         =   "B. Idioventricular"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   360
         TabIndex        =   4
         Top             =   1200
         Width           =   8895
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00800000&
         Caption         =   "A. Acellerated Idioventricular "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   615
         Left            =   360
         TabIndex        =   2
         Top             =   240
         Width           =   6255
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   2415
      Left            =   360
      Picture         =   "FormQ27.frx":08CA
      ScaleHeight     =   2355
      ScaleWidth      =   13875
      TabIndex        =   0
      Top             =   1560
      Width           =   13935
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C00000&
      Caption         =   $"FormQ27.frx":8C37
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1095
      Left            =   720
      TabIndex        =   3
      Top             =   240
      Width           =   12495
   End
End
Attribute VB_Name = "FormQ27"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Public score As Integer and cnt As Intger declared in ModuleVars and are program-wide variables.

Private Sub CmdNext_Click()
'evaluate answer. Replace option1 with the correct choice (options 1,2 or 3)
If Option3.Value = True Then
cnt = cnt + 1
End If
'Advance to next question
 'modify the next two lines and remove the comment markup
 FormSc.Show
 FormQ27.Hide

End Sub

Private Sub Form_Load()
'Uncomment and fill in the variable to center this form
With FormQ27
        .Top = (Screen.Height - .Height) / 2
        .Left = (Screen.Width - .Width) / 2
    End With
End Sub

Private Sub Option2_Click()
'For option 2


End Sub

Private Sub Option3_Click()
'Validation


End Sub

