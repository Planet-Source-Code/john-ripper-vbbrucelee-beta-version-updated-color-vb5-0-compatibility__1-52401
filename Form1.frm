VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VBBruceLee"
   ClientHeight    =   3780
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2925
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3780
   ScaleWidth      =   2925
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   375
      Left            =   840
      TabIndex        =   3
      Top             =   3240
      Width           =   1275
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Colour CPC 464 Mode"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   1740
      Value           =   -1  'True
      Width           =   2775
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Green CPC 464 Mode"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   420
      Width           =   2775
   End
   Begin VB.Image Image1 
      Height          =   915
      Index           =   1
      Left            =   420
      Top             =   2040
      Width           =   1515
   End
   Begin VB.Image Image1 
      Height          =   915
      Index           =   0
      Left            =   420
      Top             =   660
      Width           =   1515
   End
   Begin VB.Label Label1 
      Caption         =   "Choose Display mode"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   2115
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    If Option1(0).Value = True Then
        ColorMode = False
    Else
        ColorMode = True
    End If
    frmMain.Show
End Sub

Private Sub Form_Load()
    Image1(0).Picture = LoadPicture(App.Path & "\pants\logotiny.gif")
    Image1(1).Picture = LoadPicture(App.Path & "\pants\logotinyc.gif")
End Sub
