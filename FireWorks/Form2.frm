VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Settings"
   ClientHeight    =   2685
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3390
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2685
   ScaleWidth      =   3390
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "Rocket Speed"
      Height          =   615
      Left            =   0
      TabIndex        =   5
      Top             =   1440
      Width           =   3375
      Begin VB.HScrollBar RocketSpeed 
         Height          =   255
         Left            =   120
         Max             =   50
         Min             =   1
         TabIndex        =   6
         Top             =   240
         Value           =   10
         Width           =   3135
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Firework Lanuch Rate"
      Height          =   615
      Left            =   0
      TabIndex        =   3
      Top             =   720
      Width           =   3375
      Begin VB.HScrollBar LaunchRate 
         Height          =   255
         Left            =   120
         Max             =   1
         Min             =   1000
         TabIndex        =   4
         Top             =   240
         Value           =   50
         Width           =   3135
      End
   End
   Begin VB.CommandButton SaveSettings 
      Caption         =   "Ok"
      Height          =   495
      Left            =   2160
      TabIndex        =   2
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tail Length"
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3375
      Begin VB.HScrollBar TailLength 
         Height          =   255
         Left            =   120
         Max             =   1
         Min             =   255
         TabIndex        =   1
         Top             =   240
         Value           =   8
         Width           =   3135
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    TailLength.Value = GetSetting(App.EXEName, "SETTINGS", "TailLength", 8)
    LaunchRate.Value = GetSetting(App.EXEName, "SETTINGS", "LaunchRate", 50)
    RocketSpeed.Value = GetSetting(App.EXEName, "SETTINGS", "RocketSpeed", 10)
End Sub

Private Sub SaveSettings_Click()
    SaveSetting App.EXEName, "SETTINGS", "TailLength", TailLength.Value
    SaveSetting App.EXEName, "SETTINGS", "LaunchRate", LaunchRate.Value
    SaveSetting App.EXEName, "SETTINGS", "RocketSpeed", RocketSpeed.Value
    
    End
End Sub
