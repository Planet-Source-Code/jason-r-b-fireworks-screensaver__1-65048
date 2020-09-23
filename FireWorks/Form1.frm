VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer2 
      Interval        =   50
      Left            =   2040
      Top             =   2040
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   81
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   2520
      Top             =   2040
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type Particle
    X As Single
    Y As Single
    Xv As Single
    Yv As Single
    Life As Integer
    Dead As Boolean
    Color As Long
End Type

Private Type FireWork
    X As Single
    Y As Single
    Height As Integer
    Color As Long
    Exploded As Boolean
    P() As Particle
End Type

Const AC_SRC_OVER = &H0

Private Type BLENDFUNCTION
  BlendOp As Byte
  BlendFlags As Byte
  SourceConstantAlpha As Byte
  AlphaFormat As Byte
End Type

Private Declare Function AlphaBlend Lib "msimg32.dll" (ByVal hDC As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal hDC As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal BLENDFUNCT As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32.dll" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function SetPixelV Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long

Dim BF As BLENDFUNCTION
Dim lBF As Long

Dim FW() As FireWork
Dim FWCount As Integer
Dim RocketSpeed As Integer

Private Sub StartFireWork()
    For i = 0 To FWCount
        If FW(i).Y = -1 Then
            GoTo MAKEFIREWORK
        End If
    Next i
    
    FWCount = FWCount + 1
    
    ReDim Preserve FW(FWCount)
    i = FWCount
    
MAKEFIREWORK:
    
    With FW(i)
        .X = Int(Rnd * Me.ScaleWidth)
        .Y = Me.ScaleHeight
        .Height = Rnd * Me.ScaleHeight ' - ((Me.ScaleHeight / 2) + Int((Me.ScaleHeight / 2) * Rnd))
        .Color = Int(Rnd * vbWhite)
        .Exploded = False
        ReDim .P(10)
    End With
End Sub

Private Sub DrawFireWork(tFW As FireWork)
    Dim DeadCount As Integer
    Dim RndSpeed As Single
    Dim RndDeg As Single

    With tFW
        If .Exploded Then
            For i = 0 To UBound(.P)
                If .P(i).Life > 0 Then
                    .P(i).Life = .P(i).Life - 1
                    .P(i).X = .P(i).X + .P(i).Xv
                    .P(i).Y = .P(i).Y + .P(i).Yv
                    .P(i).Xv = .P(i).Xv / 1.05
                    .P(i).Yv = .P(i).Yv / 1.05 + 0.05
                    PSet (.P(i).X, .P(i).Y), .P(i).Color
                ElseIf .P(i).Life > -40 Then
                    .P(i).Life = .P(i).Life - 1
                    .P(i).X = .P(i).X + .P(i).Xv + (0.5 - Rnd)
                    .P(i).Y = .P(i).Y + .P(i).Yv + 0.1
                    .P(i).Xv = .P(i).Xv / 1.05
                    .P(i).Yv = .P(i).Yv
                    SetPixelV Me.hDC, .P(i).X, .P(i).Y, .P(i).Color
                Else
                    .P(i).Dead = True
                    DeadCount = DeadCount + 1
                End If
            Next i
            
            If DeadCount >= UBound(.P) Then
                .Y = -1
            End If
        Else
            .Y = .Y - RocketSpeed
            If .Y < .Height Then
                Dim ExplosionShape As Integer
                
                ExplosionShape = Int(Rnd * 6)
                
                Select Case ExplosionShape
                    Case 0 'Regular
                        ReDim .P(Int(Rnd * 100) + 100)
                        
                        For i = 0 To UBound(.P)
                            .P(i).X = .X
                            .P(i).Y = .Y
                            .P(i).Life = Int(Rnd * 20) + 20
                            
                            RndSpeed = (Rnd * 5)
                            RndDeg = (Rnd * 360) / 57.3
                            
                            .P(i).Xv = RndSpeed * Cos(RndDeg)
                            .P(i).Yv = RndSpeed * Sin(RndDeg)
                            .P(i).Color = .Color
                        Next i
                        
                        .Exploded = True
                    Case 1 'Smilely
                        ReDim .P(35)
                        ReDim .P(50)
                        ReDim .P(52)
                        
                        For i = 0 To 35
                            .P(i).X = .X
                            .P(i).Y = .Y
                            .P(i).Life = 50
                            
                            .P(i).Xv = 3 * Cos(((360 / 35) * (i + 1)) / 57.3)
                            .P(i).Yv = 3 * Sin(((360 / 35) * (i + 1)) / 57.3)
                            .P(i).Color = .Color
                        Next i
                        
                        For i = 36 To 50
                            .P(i).X = .X
                            .P(i).Y = .Y
                            .P(i).Life = 50
                            
                            .P(i).Xv = 2 * Cos(((360 / 35) * i + 15) / 57.3)
                            .P(i).Yv = 2 * Sin(((360 / 35) * i + 15) / 57.3)
                            .P(i).Color = .Color
                        Next i
                        
                        With .P(51)
                            .X = tFW.X
                            .Y = tFW.Y
                            .Life = 50
                            .Xv = 2 * Cos(-55 / 57.3)
                            .Yv = 2 * Sin(-55 / 57.3)
                            .Color = tFW.Color
                        End With
                        
                        With .P(52)
                            .X = tFW.X
                            .Y = tFW.Y
                            .Life = 50
                            .Xv = 2 * Cos(-125 / 57.3)
                            .Yv = 2 * Sin(-125 / 57.3)
                            .Color = tFW.Color
                        End With
                        
                        .Exploded = True
                    Case 2 'Star
                        ReDim .P(50)
                        
                        RndDeg = Int(360 * Rnd)
                        
                        For i = 0 To UBound(.P)
                            .P(i).X = .X
                            .P(i).Y = .Y
                            .P(i).Life = 50
                            
                            .P(i).Xv = (i * 0.1) * Cos(((360 / 5) * (i + 1) + RndDeg) / 57.3)
                            .P(i).Yv = (i * 0.1) * Sin(((360 / 5) * (i + 1) + RndDeg) / 57.3)
                            .P(i).Color = .Color
                        Next i
                        
                        .Exploded = True
                    Case 3 'Spiral
                        ReDim .P(50)
                        
                        RndDeg = (360 * Rnd)
                        
                        For i = 0 To UBound(.P)
                            .P(i).X = .X
                            .P(i).Y = .Y
                            .P(i).Life = 50
                            
                            .P(i).Xv = (i * 0.1) * Cos(((360 / 25) * (i + 1) + RndDeg) / 57.3)
                            .P(i).Yv = (i * 0.1) * Sin(((360 / 25) * (i + 1) + RndDeg) / 57.3)
                            .P(i).Color = .Color
                        Next i
                        
                        .Exploded = True
                    Case 4 'Regular Random
                        
                        
                        ReDim .P(Int(Rnd * 100) + 100)
                        
                        For i = 0 To UBound(.P)
                            .P(i).X = .X
                            .P(i).Y = .Y
                            .P(i).Life = Int(Rnd * 20) + 20
                            
                            RndSpeed = (Rnd * 5)
                            RndDeg = (Rnd * 360) / 57.3
                            
                            .P(i).Xv = RndSpeed * Cos(RndDeg)
                            .P(i).Yv = RndSpeed * Sin(RndDeg)
                            .P(i).Color = Int(Rnd * vbWhite)
                        Next i
                        
                        .Exploded = True
                End Select
            Else
                SetPixelV Me.hDC, .X, .Y, vbWhite
            End If
        End If
    End With
End Sub

Private Sub Form_Click()
    End
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    End
End Sub

Private Sub Form_Load()
    Randomize

    RocketSpeed = GetSetting(App.EXEName, "SETTINGS", "RocketSpeed", 10)
    Timer2.Interval = GetSetting(App.EXEName, "SETTINGS", "LaunchRate", 50)

    FWCount = -1
    
    Picture1.Width = Me.ScaleWidth
    Picture1.Height = Me.ScaleHeight
    
    With BF
        .BlendOp = AC_SRC_OVER
        .BlendFlags = 0
        .SourceConstantAlpha = GetSetting(App.EXEName, "SETTINGS", "TailLength", 8)
        .AlphaFormat = 0
    End With
End Sub

Private Sub Timer1_Timer()
    For i = 0 To FWCount
        DrawFireWork FW(i)
    Next i

    
    RtlMoveMemory lBF, BF, 4
    AlphaBlend Me.hDC, 0, 0, Me.ScaleWidth, Me.ScaleHeight, Picture1.hDC, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, lBF
    Me.Refresh
End Sub

Private Sub Timer2_Timer()
    StartFireWork
End Sub

