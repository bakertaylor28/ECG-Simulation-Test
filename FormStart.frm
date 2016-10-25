VERSION 5.00
Begin VB.Form FormStart 
   BackColor       =   &H00800000&
   Caption         =   "ECG  Simulation Test (32-bit program) "
   ClientHeight    =   6060
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11085
   FillColor       =   &H00800000&
   FillStyle       =   0  'Solid
   Icon            =   "FormStart.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "FormStart.frx":08CA
   ScaleHeight     =   6060
   ScaleWidth      =   11085
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdExit 
      BackColor       =   &H00FF8080&
      Caption         =   "EXIT "
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
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4560
      Width           =   2655
   End
   Begin VB.CommandButton CmdAbout 
      BackColor       =   &H00FF8080&
      Caption         =   "ABOUT "
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
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4560
      Width           =   2775
   End
   Begin VB.CommandButton CmdStart 
      BackColor       =   &H00FF8080&
      Caption         =   "START TEST "
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
      Left            =   360
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4560
      Width           =   2775
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   $"FormStart.frx":8BBC
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
      Height          =   3135
      Left            =   6240
      TabIndex        =   4
      Top             =   1080
      Width           =   4695
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "ECG Simulation Test "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   855
      Left            =   2520
      TabIndex        =   0
      Top             =   240
      Width           =   5295
   End
End
Attribute VB_Name = "FormStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Public score As Integer and cnt As Intger declared in ModuleVars and are program-wide variables.

Private Sub CmdAbout_Click()
'Program Name, Author, and Copyright Information
MsgBox "ECG Simulation Test. Written by A. Scott Fulkerson Powered by SkillStat. Copyright (C) 2016, All Rights Reserved.", vbOKOnly, "About ECG Simulation Test"

End Sub

Private Sub CmdExit_Click()
'Close the program
Unload FormStart
End
End Sub

Private Sub CmdStart_Click()
'initialise correct answer count,extra-credit correct answer count,  and score to 0

Score = 0
cnt = 0
extra = 0
converta = 0
convertb = 0

'advance to First Question. Subsequent questions linked by the next button on each question form.

FormQ1.Show
FormStart.Hide

End Sub

Private Sub Form_Load()
    With FormStart
        .Top = (Screen.Height - .Height) / 2
        .Left = (Screen.Width - .Width) / 2
    End With
End Sub

