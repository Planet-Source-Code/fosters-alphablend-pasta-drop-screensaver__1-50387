VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mellon Screensaver"
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3885
   ControlBox      =   0   'False
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   3885
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check2 
      Caption         =   "Can I have trails please"
      Height          =   195
      Left            =   214
      TabIndex        =   9
      Top             =   1260
      Width           =   2055
   End
   Begin VB.CheckBox Check1 
      Caption         =   "My PC is slow"
      Height          =   195
      Left            =   214
      TabIndex        =   8
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   2816
      TabIndex        =   6
      Top             =   2700
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   315
      Left            =   1860
      TabIndex        =   5
      Top             =   2700
      Width           =   855
   End
   Begin VB.HScrollBar HS 
      Height          =   255
      LargeChange     =   10
      Left            =   214
      Max             =   500
      Min             =   5
      TabIndex        =   1
      Top             =   420
      Value           =   5
      Width           =   3435
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Mike Toye, 2003"
      ForeColor       =   &H00808080&
      Height          =   195
      Left            =   2460
      TabIndex        =   7
      Top             =   2220
      Width           =   1215
   End
   Begin VB.Label lblNumber 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Left            =   1661
      TabIndex        =   4
      Top             =   720
      Width           =   495
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "500"
      Height          =   195
      Left            =   3401
      TabIndex        =   3
      Top             =   720
      Width           =   255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "5"
      Height          =   195
      Left            =   221
      TabIndex        =   2
      Top             =   720
      Width           =   255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Number of Logos"
      Height          =   195
      Left            =   221
      TabIndex        =   0
      Top             =   180
      Width           =   1335
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    SaveSetting "PastaDrop", "Settings", "NumLogos", HS.Value
    SaveSetting "PastaDrop", "Settings", "Slow", Check1.Value
    SaveSetting "PastaDrop", "Settings", "Alpha", Check2.Value
    Unload Me
    End
End Sub

Private Sub Command2_Click()
    Unload Me
    End
End Sub

Private Sub Form_Load()
    If Len(GetSetting("PastaDrop", "Settings", "NumLogos")) = 0 Then
        HS.Value = 50
    Else
        HS.Value = CInt(GetSetting("PastaDrop", "Settings", "NumLogos"))
    End If
    If Len(GetSetting("PastaDrop", "Settings", "Slow")) = 0 Then
        Check1.Value = vbUnchecked
    Else
        Check1.Value = CInt(GetSetting("PastaDrop", "Settings", "Slow"))
    End If
    If Len(GetSetting("PastaDrop", "Settings", "Alpha")) = 0 Then
        Check2.Value = vbUnchecked
    Else
        Check2.Value = CInt(GetSetting("PastaDrop", "Settings", "Alpha"))
    End If

End Sub

Private Sub HS_Change()
    lblNumber = HS.Value
End Sub
