VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9900
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   281
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   660
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox PicBlank 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   840
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   2
      Top             =   1200
      Width           =   330
   End
   Begin VB.PictureBox PicGround 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   1080
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   1
      Top             =   600
      Width           =   330
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   600
      Picture         =   "Plane.frx":0000
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   0
      Top             =   600
      Width           =   330
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   120
      Top             =   5280
   End
   Begin VB.Shape Bullet 
      DrawMode        =   6  'Mask Pen Not
      Height          =   75
      Index           =   20
      Left            =   600
      Shape           =   3  'Circle
      Top             =   360
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Shape Bullet 
      DrawMode        =   6  'Mask Pen Not
      Height          =   75
      Index           =   19
      Left            =   600
      Shape           =   3  'Circle
      Top             =   360
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Shape Bullet 
      DrawMode        =   6  'Mask Pen Not
      Height          =   75
      Index           =   18
      Left            =   600
      Shape           =   3  'Circle
      Top             =   360
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Shape Bullet 
      DrawMode        =   6  'Mask Pen Not
      Height          =   75
      Index           =   17
      Left            =   600
      Shape           =   3  'Circle
      Top             =   360
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Shape Bullet 
      DrawMode        =   6  'Mask Pen Not
      Height          =   75
      Index           =   16
      Left            =   600
      Shape           =   3  'Circle
      Top             =   360
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Shape Bullet 
      DrawMode        =   6  'Mask Pen Not
      Height          =   75
      Index           =   15
      Left            =   600
      Shape           =   3  'Circle
      Top             =   360
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Shape Bullet 
      DrawMode        =   6  'Mask Pen Not
      Height          =   75
      Index           =   14
      Left            =   600
      Shape           =   3  'Circle
      Top             =   360
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Shape Bullet 
      DrawMode        =   6  'Mask Pen Not
      Height          =   75
      Index           =   13
      Left            =   600
      Shape           =   3  'Circle
      Top             =   360
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Shape Bullet 
      DrawMode        =   6  'Mask Pen Not
      Height          =   75
      Index           =   12
      Left            =   600
      Shape           =   3  'Circle
      Top             =   360
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Shape Bullet 
      DrawMode        =   6  'Mask Pen Not
      Height          =   75
      Index           =   11
      Left            =   600
      Shape           =   3  'Circle
      Top             =   360
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Shape Bullet 
      DrawMode        =   6  'Mask Pen Not
      Height          =   75
      Index           =   10
      Left            =   600
      Shape           =   3  'Circle
      Top             =   360
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Shape Bullet 
      DrawMode        =   6  'Mask Pen Not
      Height          =   75
      Index           =   9
      Left            =   600
      Shape           =   3  'Circle
      Top             =   360
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Shape Bullet 
      DrawMode        =   6  'Mask Pen Not
      Height          =   75
      Index           =   8
      Left            =   600
      Shape           =   3  'Circle
      Top             =   360
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Shape Bullet 
      DrawMode        =   6  'Mask Pen Not
      Height          =   75
      Index           =   7
      Left            =   600
      Shape           =   3  'Circle
      Top             =   360
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Shape Bullet 
      DrawMode        =   6  'Mask Pen Not
      Height          =   75
      Index           =   6
      Left            =   600
      Shape           =   3  'Circle
      Top             =   360
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Shape Bullet 
      DrawMode        =   6  'Mask Pen Not
      Height          =   75
      Index           =   5
      Left            =   600
      Shape           =   3  'Circle
      Top             =   360
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Shape Bullet 
      DrawMode        =   6  'Mask Pen Not
      Height          =   75
      Index           =   4
      Left            =   600
      Shape           =   3  'Circle
      Top             =   360
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Shape Bullet 
      DrawMode        =   6  'Mask Pen Not
      Height          =   75
      Index           =   3
      Left            =   600
      Shape           =   3  'Circle
      Top             =   360
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Shape Bullet 
      DrawMode        =   6  'Mask Pen Not
      Height          =   75
      Index           =   2
      Left            =   600
      Shape           =   3  'Circle
      Top             =   360
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Shape Bullet 
      DrawMode        =   6  'Mask Pen Not
      Height          =   75
      Index           =   1
      Left            =   600
      Shape           =   3  'Circle
      Top             =   360
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Shape Bullet 
      DrawMode        =   6  'Mask Pen Not
      Height          =   75
      Index           =   0
      Left            =   600
      Shape           =   3  'Circle
      Top             =   360
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Shape Sh 
      BorderStyle     =   3  'Dot
      DrawMode        =   6  'Mask Pen Not
      FillColor       =   &H00808000&
      FillStyle       =   7  'Diagonal Cross
      Height          =   135
      Left            =   7320
      Shape           =   2  'Oval
      Top             =   2400
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
Private Sub Form_Load()
    While PlaneW <= 0
        PlaneW = GenerateDC(App.Path & "\planew.bmp")
    Wend
    While PlaneB <= 0
        PlaneB = GenerateDC(App.Path & "\planeB.bmp")
    Wend
    While BarrelW <= 0
        BarrelW = GenerateDC(App.Path & "\BarrelW.bmp")
    Wend
    While BarrelB <= 0
        BarrelB = GenerateDC(App.Path & "\BarrelB.bmp")
    Wend
    While BoomW <= 0
        BoomW = GenerateDC(App.Path & "\BoomW.bmp")
    Wend
    While BoomB <= 0
        BoomB = GenerateDC(App.Path & "\BoomB.bmp")
    Wend
    While GuyB <= 0
        GuyB = GenerateDC(App.Path & "\GuyB.bmp")
    Wend
    While GuyW <= 0
        GuyW = GenerateDC(App.Path & "\GuyW.bmp")
    Wend
    While GraveBW <= 0
        GraveBW = GenerateDC(App.Path & "\Grave.bmp")
    Wend
    While groundW <= 0
        groundW = GenerateDC(App.Path & "\bottomWhite.bmp")
    Wend
    While groundB <= 0
        groundB = GenerateDC(App.Path & "\bottomBlack.bmp")
    Wend
    While BG <= 0
        BG = GenerateDC(App.Path & "\bgbig.bmp")
    Wend
    PX = 200
    x = 500
    y = 500
    For g = 0 To 10
        GuyX(g) = Form1.ScaleWidth + Int(Rnd * 1150)
    Next
    Sh.Top = Int(Rnd * (Form1.ScaleHeight - 9)) + 9
    Sh.Left = Form1.ScaleWidth
    GroundTop = 780
End Sub

Public Function Run()
    Do
        Form1.Refresh
        Up = GetAsyncKeyState(vbKeyUp)
        Down = GetAsyncKeyState(vbKeyDown)
        Right = GetAsyncKeyState(vbKeyRight)
        Lef = GetAsyncKeyState(vbKeyLeft)
        STP = GetAsyncKeyState(vbKeyEscape)
        Space = GetAsyncKeyState(vbKeyControl)
        Shift = GetAsyncKeyState(vbKeyShift)
        If Shift <= -32767 Then
            Shoot = True
        Else
            Shoot = False
        End If
        If Up <= -32767 Then
            PX = 0
            y = y - 5
        ElseIf Down > -32767 Then
            PX = 200
        End If
        If Right <= -32767 Then
            x = x + 10
        End If
        If Down <= -32767 Then
            y = y + 5
            PX = 400
        ElseIf Up > -32767 Then
            PX = 200
        End If
        If Lef <= -32767 Then
            x = x - 10
        End If
        If Space <= -32767 And slowdown = 5 Then
            slowdown = 0
            While ind < 5
                If Bomb(ind) = False Then
                    BarrelX(ind) = x + 100
                    BarrelY(ind) = y + 100
                    Bomb(ind) = True
                    If Right <= -32767 Then
                        Righty(ind) = True
                        BarrelV(ind) = 0.8
                        BarrelMove(ind) = 30
                        If Up <= -32767 Then
                            BarrelV(ind) = 0.85
                            BarrelMove(ind) = 40
                        End If
                    ElseIf Lef <= -32767 Then
                        Lefty(ind) = True
                        BarrelV(ind) = 0.85
                        BarrelMove(ind) = 30
                        If Up <= -32767 Then
                            BarrelV(ind) = 0.85
                            BarrelMove(ind) = 40
                        End If
                    End If
                    ind = 4
                End If
                ind = ind + 1
            Wend
            ind = 0
        End If
        slowdown = slowdown + 1
        If slowdown = 6 Then slowdown = 5
        If STP = -32767 Then
            DeleteGeneratedDC PlaneB
            DeleteGeneratedDC PlaneW
            DeleteGeneratedDC BarrelB
            DeleteGeneratedDC BarrelW
            DeleteGeneratedDC BoomB
            DeleteGeneratedDC BoomW
            DeleteGeneratedDC GuyB
            DeleteGeneratedDC GuyW
            DeleteGeneratedDC BG
            
            Unload Me
            Set frmMemoryDC = Nothing
            End
        End If
        
        BitBlt Form1.hdc, 0, 0, Form1.ScaleWidth, Form1.ScaleHeight, BG, BGPosition, 0, vbSrcCopy
        BGPosition = BGPosition + 5
        If BGPosition >= (10301 - 1144) Then BGPosition = 1144
        For f = 0 To 10
            GuyX(f) = GuyX(f) - 7
            If GuyX(f) < Form1.ScaleWidth And GuyX(f) >= 0 Then
                'BitBlt PicGround.hdc, 0, 0, 20, 30, ground, BGPosition + GuyX(f), 0, vbSrcAnd
                'BitBlt PicGround.hdc, 0, 0, 20, 30, ground, BGPosition + GuyX(f), 0, vbSrcPaint
                'While Collide = False
                '    Collide = CollisionDetect(0, GroundTop, Pic, 0, 800, PicGround, PicBlank)
                '    GroundTop = GroundTop + 1
                '    If GroundTop > 820 Then Collide = True
                'Wend
                'GroundTop = 780
                'Collide = False
                'GuyY(f) = GroundTop
                BitBlt Form1.hdc, GuyX(f), 785, 20, 30, GuyW, GuyP(f), 0, vbSrcAnd
                BitBlt Form1.hdc, GuyX(f), 785, 20, 30, GuyB, GuyP(f), 0, vbSrcPaint
                'BitBlt PicGround.hdc, 0, 0, 20, 30, ground, BGPosition + GuyX(f), 0, vbSrcCopy
                GuySlowdown(f) = GuySlowdown(f) + 1
                If GuySlowdown(f) = 5 Then
                    GuyP(f) = GuyP(f) + 20
                    GuySlowdown(f) = 0
                End If
                If GuyP(f) = 40 Then GuyP(f) = 0
            ElseIf GuyX(f) < 0 Then
                GuyX(f) = 1145 + Int(Rnd * 500)
            End If
            If Grave(f) = True Then
                BitBlt Form1.hdc, GraveX(f), 790, 15, 12, GraveBW, 15, 0, vbSrcAnd
                BitBlt Form1.hdc, GraveX(f), 790, 15, 12, GraveBW, 0, 0, vbSrcPaint
                GraveX(f) = GraveX(f) - 5
                If GraveEnd(f) < BGPosition Then Grave(f) = False
            End If
        Next
        For t = 0 To 4
                If Bomb(t) = True Then
                    BitBlt Form1.hdc, BarrelX(t), BarrelY(t), 20, 20, BarrelW, BarrelP(t), 0, vbSrcAnd
                    BitBlt Form1.hdc, BarrelX(t), BarrelY(t), 20, 20, BarrelB, BarrelP(t), 0, vbSrcPaint
                    BarrelMove(t) = BarrelV(t) * BarrelMove(t)
                    BarrelY(t) = BarrelY(t) * 1.06
                    If Lefty(t) = True Then
                        BarrelX(t) = BarrelX(t) - BarrelMove(t)
                    Else
                        BarrelX(t) = BarrelX(t) + BarrelMove(t)
                    End If
                    If BarrelY(t) > Form1.ScaleHeight - 10 Then
                        Bomb(t) = False
                        Righty(t) = False
                        Lefty(t) = False
                        BarrelV(t) = 0
                        Boom(t) = True
                        BoomX(t) = BarrelX(t)
                    End If
                    If SlowDownBarrel(t) = 2 Then
                        BarrelP(t) = BarrelP(t) + 20
                        SlowDownBarrel(t) = 0
                    End If
                    SlowDownBarrel(t) = SlowDownBarrel(t) + 1
                    If SlowDownBarrel(t) = 3 Then SlowDownBarrel(t) = 2
                    If BarrelP(t) = 80 Then
                        BarrelP(t) = 0
                        SlowDownBarrel(t) = 0
                    End If
                End If
                If Boom(t) = True Then
                    BitBlt Form1.hdc, BoomX(t), 790, 20, 20, BoomW, BoomP(t), 0, vbSrcAnd
                    BitBlt Form1.hdc, BoomX(t), 790, 20, 20, BoomB, BoomP(t), 0, vbSrcPaint
                    If SlowDownBoom(t) = 2 Then BoomP(t) = BoomP(t) + 20
                    SlowDownBoom(t) = SlowDownBoom(t) + 1
                    If SlowDownBoom(t) = 3 Then SlowDownBoom(t) = 2
                    If BoomP(t) = 80 Then
                        Boom(t) = False
                        BoomP(t) = 0
                        SlowDownBoom(t) = 0
                    End If
                    For S = 0 To 10
                        If BoomP(t) <= 20 And Boom(t) = True And BoomX(t) <= GuyX(S) + 20 And BoomX(t) + 20 >= GuyX(S) Then
                            While ind < 21
                                If Grave(ind) = False Then
                                    Grave(ind) = True
                                    GraveX(ind) = GuyX(S)
                                    GraveEnd(ind) = BGPosition + GuyX(S)
                                    ind = 20
                                End If
                                ind = ind + 1
                            Wend
                            GuyX(S) = Int(Rnd * 800) + Form1.ScaleWidth
                            ind = 0
                        End If
                    Next
                End If
        Next
        BitBlt Form1.hdc, x, y, 200, 130, PlaneW, PX, 0, vbSrcAnd
        BitBlt Form1.hdc, x, y, 200, 130, PlaneB, PX, 0, vbSrcPaint
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            While indb < 21
                If Bullet(indb).Visible = False And Shoot = True Then
                    Bullet(indb).Visible = True
                    Bullet(indb).Left = x + 178
                    Bullet(indb).Top = y + 60
                    angle(indb) = 0
                    If Down <= -32767 Then angle(indb) = 13
                    If Up <= -32767 Then angle(indb) = -13
                    If angle(indb) < 0 Then
                        Bullet(indb).Top = y + 38
                        If Right <= -32767 Then angle(indb) = -10
                    End If
                    If angle(indb) > 0 Then
                        Bullet(indb).Top = y + 86
                        If Right <= -32767 Then angle(indb) = 10
                    End If
                    indb = 20
                End If
                If Bullet(indb).Visible = True Then
                    Bullet(indb).Left = Bullet(indb).Left + 21
                    Bullet(indb).Top = Bullet(indb).Top + angle(indb)
                    If Bullet(indb).Top < 0 Or Bullet(indb).Top > Form1.ScaleHeight Then Bullet(indb).Visible = False
                    If Bullet(indb).Left > Form1.ScaleWidth Then Bullet(indb).Visible = False
                End If
                indb = indb + 1
            Wend
            indb = 0
                
        'Sh.Left = Sh.Left - 20
        'If Sh.Left < 0 Then
        '    Sh.Left = Form1.ScaleWidth
        '    Sh.Top = Int(Rnd * (Form1.ScaleHeight - 9)) + 9
        'End If
    Loop
End Function

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    Run
End Sub
