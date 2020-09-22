VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   12585
   ClientLeft      =   -315
   ClientTop       =   0
   ClientWidth     =   17190
   DrawMode        =   7  'Invert
   BeginProperty Font 
      Name            =   "Arial Black"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   839
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1146
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picBlank 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   1800
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   2
      Top             =   -1080
      Width           =   330
   End
   Begin VB.PictureBox picRope 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1080
      Left            =   1440
      Picture         =   "Plane2.frx":0000
      ScaleHeight     =   70
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   0
      Top             =   -1080
      Width           =   330
   End
   Begin VB.PictureBox picGuy 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   1080
      Picture         =   "Plane2.frx":0162
      ScaleHeight     =   450
      ScaleWidth      =   300
      TabIndex        =   1
      Top             =   -1080
      Width           =   330
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   120
      Top             =   5280
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   99
      Visible         =   0   'False
      X1              =   104
      X2              =   144
      Y1              =   520
      Y2              =   536
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   98
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   97
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   96
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   95
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   94
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   93
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   92
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   91
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   90
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   89
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   88
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   87
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   86
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   85
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   84
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   83
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   82
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   81
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   80
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   79
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   78
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   77
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   76
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   75
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   74
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   73
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   72
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   71
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   70
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   69
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   68
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   67
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   66
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   65
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   64
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   63
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   62
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   61
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   60
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   59
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   58
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   57
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   56
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   55
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   54
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   53
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   52
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   51
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   50
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   49
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   48
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   47
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   46
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   45
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   44
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   43
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   42
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   41
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   40
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   39
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   38
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   37
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   36
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   35
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   34
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   33
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   32
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   31
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   30
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   29
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   28
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   27
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   26
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   25
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   24
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   23
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   22
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   21
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   20
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   19
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   18
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   17
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   16
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   15
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   14
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   13
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   12
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   11
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   10
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   9
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   8
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   7
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   6
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   5
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   4
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   3
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   2
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   1
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   0
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   399
      Visible         =   0   'False
      X1              =   32
      X2              =   32
      Y1              =   712
      Y2              =   728
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   398
      Visible         =   0   'False
      X1              =   40
      X2              =   40
      Y1              =   712
      Y2              =   728
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   397
      Visible         =   0   'False
      X1              =   48
      X2              =   48
      Y1              =   712
      Y2              =   728
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   396
      Visible         =   0   'False
      X1              =   112
      X2              =   112
      Y1              =   688
      Y2              =   704
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   395
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   394
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   393
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   392
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   391
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   390
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   389
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   388
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   387
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   386
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   385
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   384
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   383
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   382
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   381
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   380
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   379
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   378
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   377
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   376
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   375
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   374
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   373
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   372
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   371
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   370
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   369
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   368
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   367
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   366
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   365
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   364
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   363
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   362
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   361
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   360
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   359
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   358
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   357
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   356
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   355
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   354
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   353
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   352
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   351
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   350
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   349
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   348
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   347
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   346
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   345
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   344
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   343
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   342
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   341
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   340
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   339
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   338
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   337
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   336
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   335
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   334
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   333
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   332
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   331
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   330
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   329
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   328
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   327
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   326
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   325
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   324
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   323
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   322
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   321
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   320
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   319
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   318
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   317
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   316
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   315
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   314
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   313
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   312
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   311
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   310
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   309
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   308
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   307
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   306
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   305
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   304
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   303
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   302
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   301
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   300
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   299
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   298
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   297
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   296
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   295
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   294
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   293
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   292
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   291
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   290
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   289
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   288
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   287
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   286
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   285
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   284
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   283
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   282
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   281
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   280
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   279
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   278
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   277
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   276
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   275
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   274
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   273
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   272
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   271
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   270
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   269
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   268
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   267
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   266
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   265
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   264
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   263
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   262
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   261
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   260
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   259
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   258
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   257
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   256
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   255
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   254
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   253
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   252
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   251
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   250
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   249
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   248
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   247
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   246
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   245
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   244
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   243
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   242
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   241
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   240
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   239
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   238
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   237
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   236
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   235
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   234
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   233
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   232
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   231
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   230
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   229
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   228
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   227
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   226
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   225
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   224
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   223
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   222
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   221
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   220
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   219
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   218
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   217
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   216
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   215
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   214
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   213
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   212
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   211
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   210
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   209
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   208
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   207
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   206
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   205
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   204
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   203
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   202
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   201
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   200
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   199
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   198
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   197
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   196
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   195
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   194
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   193
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   192
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   191
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   190
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   189
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   188
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   187
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   186
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   185
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   184
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   183
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   182
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   181
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   180
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   179
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   178
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   177
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   176
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   175
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   174
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   173
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   172
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   171
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   170
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   169
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   168
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   167
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   166
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   165
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   164
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   163
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   162
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   161
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   160
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   159
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   158
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   157
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   156
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   155
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   154
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   153
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   152
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   151
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   150
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   149
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   148
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   147
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   146
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   145
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   144
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   143
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   142
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   141
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   140
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   139
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   138
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   137
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   136
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   135
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   134
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   133
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   132
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   131
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   130
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   129
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   128
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   127
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   126
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   125
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   124
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   123
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   122
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   121
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   120
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   119
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   118
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   117
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   116
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   115
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   114
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   113
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   112
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   111
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   110
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   109
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   108
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   107
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   106
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   105
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   104
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   103
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   102
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   101
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   100
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   99
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   98
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   97
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   96
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   95
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   94
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   93
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   92
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   91
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   90
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   89
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   88
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   87
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   86
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   85
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   84
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   83
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   82
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   81
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   80
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   79
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   78
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   77
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   76
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   75
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   74
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   73
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   72
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   71
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   70
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   69
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   68
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   67
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   66
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   65
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   64
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   63
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   62
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   61
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   60
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   59
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   58
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   57
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   56
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   55
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   54
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   53
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   52
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   51
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   50
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   49
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   48
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   47
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   46
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   45
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   44
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   43
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   42
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   41
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   40
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   39
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   38
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   37
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   36
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   35
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   34
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   33
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   32
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   31
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   30
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   29
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   28
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   27
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   26
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   25
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   24
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   23
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   22
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   21
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   20
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   19
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   18
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   17
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   16
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   15
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   14
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   13
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   12
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   11
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   10
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   9
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   8
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   7
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   6
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   5
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   4
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   3
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   2
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   1
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   0
      Visible         =   0   'False
      X1              =   24
      X2              =   24
      Y1              =   712
      Y2              =   728
   End
   Begin VB.Shape Bullet 
      BackColor       =   &H80000003&
      BorderStyle     =   3  'Dot
      DrawMode        =   2  'Blackness
      FillColor       =   &H00404040&
      FillStyle       =   0  'Solid
      Height          =   45
      Index           =   20
      Left            =   600
      Shape           =   2  'Oval
      Top             =   360
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Shape Bullet 
      BackColor       =   &H80000003&
      BorderStyle     =   3  'Dot
      DrawMode        =   2  'Blackness
      FillColor       =   &H00404040&
      FillStyle       =   0  'Solid
      Height          =   45
      Index           =   19
      Left            =   600
      Shape           =   2  'Oval
      Top             =   360
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Shape Bullet 
      BackColor       =   &H80000003&
      BorderStyle     =   3  'Dot
      DrawMode        =   2  'Blackness
      FillColor       =   &H00404040&
      FillStyle       =   0  'Solid
      Height          =   45
      Index           =   18
      Left            =   600
      Shape           =   2  'Oval
      Top             =   360
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Shape Bullet 
      BackColor       =   &H80000003&
      BorderStyle     =   3  'Dot
      DrawMode        =   2  'Blackness
      FillColor       =   &H00404040&
      FillStyle       =   0  'Solid
      Height          =   45
      Index           =   17
      Left            =   600
      Shape           =   2  'Oval
      Top             =   360
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Shape Bullet 
      BackColor       =   &H80000003&
      BorderStyle     =   3  'Dot
      DrawMode        =   2  'Blackness
      FillColor       =   &H00404040&
      FillStyle       =   0  'Solid
      Height          =   45
      Index           =   16
      Left            =   600
      Shape           =   2  'Oval
      Top             =   360
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Shape Bullet 
      BackColor       =   &H80000003&
      BorderStyle     =   3  'Dot
      DrawMode        =   2  'Blackness
      FillColor       =   &H00404040&
      FillStyle       =   0  'Solid
      Height          =   45
      Index           =   15
      Left            =   600
      Shape           =   2  'Oval
      Top             =   360
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Shape Bullet 
      BackColor       =   &H80000003&
      BorderStyle     =   3  'Dot
      DrawMode        =   2  'Blackness
      FillColor       =   &H00404040&
      FillStyle       =   0  'Solid
      Height          =   45
      Index           =   14
      Left            =   600
      Shape           =   2  'Oval
      Top             =   360
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Shape Bullet 
      BackColor       =   &H80000003&
      BorderStyle     =   3  'Dot
      DrawMode        =   2  'Blackness
      FillColor       =   &H00404040&
      FillStyle       =   0  'Solid
      Height          =   45
      Index           =   13
      Left            =   600
      Shape           =   2  'Oval
      Top             =   360
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Shape Bullet 
      BackColor       =   &H80000003&
      BorderStyle     =   3  'Dot
      DrawMode        =   2  'Blackness
      FillColor       =   &H00404040&
      FillStyle       =   0  'Solid
      Height          =   45
      Index           =   12
      Left            =   600
      Shape           =   2  'Oval
      Top             =   360
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Shape Bullet 
      BackColor       =   &H80000003&
      BorderStyle     =   3  'Dot
      DrawMode        =   2  'Blackness
      FillColor       =   &H00404040&
      FillStyle       =   0  'Solid
      Height          =   45
      Index           =   11
      Left            =   600
      Shape           =   2  'Oval
      Top             =   360
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Shape Bullet 
      BackColor       =   &H80000003&
      BorderStyle     =   3  'Dot
      DrawMode        =   2  'Blackness
      FillColor       =   &H00404040&
      FillStyle       =   0  'Solid
      Height          =   45
      Index           =   10
      Left            =   600
      Shape           =   2  'Oval
      Top             =   360
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Shape Bullet 
      BackColor       =   &H80000003&
      BorderStyle     =   3  'Dot
      DrawMode        =   2  'Blackness
      FillColor       =   &H00404040&
      FillStyle       =   0  'Solid
      Height          =   45
      Index           =   9
      Left            =   600
      Shape           =   2  'Oval
      Top             =   360
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Shape Bullet 
      BackColor       =   &H80000003&
      BorderStyle     =   3  'Dot
      DrawMode        =   2  'Blackness
      FillColor       =   &H00404040&
      FillStyle       =   0  'Solid
      Height          =   45
      Index           =   8
      Left            =   600
      Shape           =   2  'Oval
      Top             =   360
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Shape Bullet 
      BackColor       =   &H80000003&
      BorderStyle     =   3  'Dot
      DrawMode        =   2  'Blackness
      FillColor       =   &H00404040&
      FillStyle       =   0  'Solid
      Height          =   45
      Index           =   7
      Left            =   600
      Shape           =   2  'Oval
      Top             =   360
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Shape Bullet 
      BackColor       =   &H80000003&
      BorderStyle     =   3  'Dot
      DrawMode        =   2  'Blackness
      FillColor       =   &H00404040&
      FillStyle       =   0  'Solid
      Height          =   45
      Index           =   6
      Left            =   600
      Shape           =   2  'Oval
      Top             =   360
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Shape Bullet 
      BackColor       =   &H80000003&
      BorderStyle     =   3  'Dot
      DrawMode        =   2  'Blackness
      FillColor       =   &H00404040&
      FillStyle       =   0  'Solid
      Height          =   45
      Index           =   5
      Left            =   600
      Shape           =   2  'Oval
      Top             =   360
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Shape Bullet 
      BackColor       =   &H80000003&
      BorderStyle     =   3  'Dot
      DrawMode        =   2  'Blackness
      FillColor       =   &H00404040&
      FillStyle       =   0  'Solid
      Height          =   45
      Index           =   4
      Left            =   600
      Shape           =   2  'Oval
      Top             =   360
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Shape Bullet 
      BackColor       =   &H80000003&
      BorderStyle     =   3  'Dot
      DrawMode        =   2  'Blackness
      FillColor       =   &H00404040&
      FillStyle       =   0  'Solid
      Height          =   45
      Index           =   3
      Left            =   600
      Shape           =   2  'Oval
      Top             =   360
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Shape Bullet 
      BackColor       =   &H80000003&
      BorderStyle     =   3  'Dot
      DrawMode        =   2  'Blackness
      FillColor       =   &H00404040&
      FillStyle       =   0  'Solid
      Height          =   45
      Index           =   2
      Left            =   600
      Shape           =   2  'Oval
      Top             =   360
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Shape Bullet 
      BackColor       =   &H80000003&
      BorderStyle     =   3  'Dot
      DrawMode        =   2  'Blackness
      FillColor       =   &H00404040&
      FillStyle       =   0  'Solid
      Height          =   45
      Index           =   1
      Left            =   600
      Shape           =   2  'Oval
      Top             =   360
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Shape Bullet 
      BackColor       =   &H80000003&
      BorderStyle     =   3  'Dot
      DrawMode        =   2  'Blackness
      FillColor       =   &H00404040&
      FillStyle       =   0  'Solid
      Height          =   45
      Index           =   0
      Left            =   600
      Shape           =   2  'Oval
      Top             =   360
      Visible         =   0   'False
      Width           =   75
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
Private Sub Form_Load()
    While RPlaneW <= 0
        RPlaneW = GenerateDC(App.Path & "\planew.bmp")
    Wend
    While RPlaneB <= 0
        RPlaneB = GenerateDC(App.Path & "\planeB.bmp")
    Wend
    While LPlaneW <= 0
        LPlaneW = GenerateDC(App.Path & "\planew2.bmp")
    Wend
    While LPlaneB <= 0
        LPlaneB = GenerateDC(App.Path & "\planeB2.bmp")
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
    While RopeW <= 0
        RopeW = GenerateDC(App.Path & "\RopeW.bmp")
    Wend
    While RopeB <= 0
        RopeB = GenerateDC(App.Path & "\RopeB.bmp")
    Wend
    While BG <= 0
        BG = GenerateDC(App.Path & "\bgbig.bmp")
    Wend
    PlaneW = RPlaneW
    PlaneB = RPlaneB
    PX = 200
    x = 500
    y = 500
    Direction = 5
    For G = 0 To 10
        GuyX(G) = Form1.ScaleWidth + Int(Rnd * 1150)
        GuyY(G) = 785
    Next
    
    TalleyY = 10
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
            For m = 0 To 10
                If GuyClimb(m) = True Then
                    GuyY(m) = GuyY(m) - 5
                End If
            Next
        ElseIf Down <= -32767 Then
            PX = 200
            For n = 0 To 10
                If GuyClimb(n) = True Then
                    GuyY(n) = GuyY(n) + 5
                End If
            Next
        End If
        If Right <= -32767 Then
            'x = x + 10
            If Direction > -4 And Direction <= 4 And LSwitch = False Then
                RSwitch = True
            ElseIf Direction < 26 Then
                Direction = Direction + 1
            End If
        End If
        If Down <= -32767 Then
            y = y + 5
            PX = 400
        ElseIf Up > -32767 Then
            PX = 200
        End If
        If Lef <= -32767 Then
            'x = x - 10
            If Direction > -4 And Direction <= 4 And RSwitch = False Then
                LSwitch = True
            ElseIf Direction > -26 Then
                Direction = Direction - 1
            End If
        End If
        If Space <= -32767 And slowdown = 5 Then
            slowdown = 0
            While ind < 5
                If Bomb(ind) = False Then
                    BarrelX(ind) = x + 100
                    BarrelY(ind) = y + 100
                    Bomb(ind) = True
                    If Lef <= -32767 Then
                        Righty(ind) = True
                        BarrelV(ind) = 0.8
                        BarrelMove(ind) = 30
                        If Up <= -32767 Then
                            BarrelV(ind) = 0.85
                            BarrelMove(ind) = 40
                        End If
                    ElseIf Right <= -32767 Then
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
        If RSwitch = True Then Direction = Direction + 1
        If LSwitch = True Then Direction = Direction - 1
        If RSwitch = True And Direction >= 4 Then RSwitch = False
        If LSwitch = True And Direction <= -4 Then LSwitch = False
        If Direction = 0 And LSwitch = True Then
            PlaneW = LPlaneW
            PlaneB = LPlaneB
        End If
        If Direction = 0 And RSwitch = True Then
            PlaneW = RPlaneW
            PlaneB = RPlaneB
        End If
        BGPosition = BGPosition + Direction
        If BGPosition >= (10301 - 1144) Then BGPosition = 1144
        For f = 0 To 10
            If GuyClimb(f) = True Then
                GuyY(f) = GuyY(f) - 1
                If Direction < 0 Then GuyX(f) = x + 130
                If Direction > 0 Then GuyX(f) = x + 70
                If y + RopeP + 70 < GuyY(f) Then
                    Fall(f) = True
                    GuyClimb(f) = False
                End If
            Else
                GuyX(f) = GuyX(f) - 2 - Direction
            End If
            If Fall(f) = True Then
                GuyY(f) = GuyY(f) + 8
                If GuyY(f) > 780 Then Fall(f) = False
            End If
            If GuyY(f) < y + 70 And GuyClimb(f) = True Then
                GuyClimb(f) = False
                GuyX(f) = -5
                TalleyX = TalleyX + 8
                If TalleyX = 1096 Then
                    TalleyX = 8
                    TalleyY = TalleyY + 24
                End If
                TalleyCount = TalleyCount + 1
                If TalleyCount = 5 Then
                    Fifth((TalleyIndex / 4) - 1).X1 = TalleyX - 40
                    Fifth((TalleyIndex / 4) - 1).X2 = TalleyX
                    Fifth((TalleyIndex / 4) - 1).Y1 = TalleyY
                    Fifth((TalleyIndex / 4) - 1).Y2 = TalleyY + 12
                    Fifth((TalleyIndex / 4) - 1).Visible = True
                    TalleyCount = 0
                Else
                    Tally(TalleyIndex).X1 = TalleyX
                    Tally(TalleyIndex).X2 = TalleyX
                    Tally(TalleyIndex).Y1 = TalleyY
                    Tally(TalleyIndex).Y2 = TalleyY + 12
                    Tally(TalleyIndex).Visible = True
                End If
                TalleyIndex = TalleyIndex + 1
            End If
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
                BitBlt Form1.hdc, GuyX(f), GuyY(f), 20, 30, GuyW, GuyP(f), 0, vbSrcAnd
                BitBlt Form1.hdc, GuyX(f), GuyY(f), 20, 30, GuyB, GuyP(f), 0, vbSrcPaint
                'BitBlt PicGround.hdc, 0, 0, 20, 30, ground, BGPosition + GuyX(f), 0, vbSrcCopy
                GuySlowdown(f) = GuySlowdown(f) + 1
                If GuySlowdown(f) = 5 Then
                    GuyP(f) = GuyP(f) + 20
                    GuySlowdown(f) = 0
                End If
                If GuyP(f) = 40 Then GuyP(f) = 0
            ElseIf GuyX(f) < 0 Then
                GuyX(f) = 1145 + Int(Rnd * 500)
                GuyY(f) = 785
            End If
            If Grave(f) = True Then
                BitBlt Form1.hdc, GraveX(f), 785, 15, 12, GraveBW, 15, 0, vbSrcAnd
                BitBlt Form1.hdc, GraveX(f), 785, 15, 12, GraveBW, 0, 0, vbSrcPaint
                GraveX(f) = GraveX(f) - Direction
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
                    For s = 0 To 10
                        If (BoomP(t) <= 20 And Boom(t) = True And BoomX(t) <= GuyX(s) + 20 And BoomX(t) + 20 >= GuyX(s)) Or Death(s) = True Then
                            While ind < 21
                                If Grave(ind) = False Then
                                    Death(s) = False
                                    Grave(ind) = True
                                    GraveX(ind) = GuyX(s)
                                    GraveEnd(ind) = BGPosition + GuyX(s)
                                    ind = 20
                                    TalleyX = TalleyX + 8
                                    If TalleyX = 1096 Then
                                        TalleyX = 8
                                        TalleyY = TalleyY + 24
                                    End If
                                    TalleyCount = TalleyCount + 1
                                    If TalleyCount = 5 Then
                                        Fifth((TalleyIndex / 4) - 1).X1 = TalleyX - 40
                                        Fifth((TalleyIndex / 4) - 1).X2 = TalleyX
                                        Fifth((TalleyIndex / 4) - 1).Y1 = TalleyY
                                        Fifth((TalleyIndex / 4) - 1).Y2 = TalleyY + 12
                                        Fifth((TalleyIndex / 4) - 1).BorderColor = &HFF&
                                        Fifth((TalleyIndex / 4) - 1).Visible = True
                                        TalleyCount = 0
                                    Else
                                        Tally(TalleyIndex).X1 = TalleyX
                                        Tally(TalleyIndex).X2 = TalleyX
                                        Tally(TalleyIndex).Y1 = TalleyY
                                        Tally(TalleyIndex).Y2 = TalleyY + 12
                                        Tally(TalleyIndex).BorderColor = &HFF&
                                        Tally(TalleyIndex).Visible = True
                                    End If
                                    TalleyIndex = TalleyIndex + 1
                                End If
                                ind = ind + 1
                            Wend
                            GuyX(s) = Int(Rnd * 800) + Form1.ScaleWidth
                            GuyY(s) = 785
                            ind = 0
                        End If
                    Next
                End If
        Next
        If y >= 670 Then
            If RopeP < 70 Then RopeP = RopeP + 10
        ElseIf RopeP > 0 Then
            RopeP = RopeP - 10
        End If
        If RopeP > 0 Then
            If Direction <= 0 Then
                BitBlt Form1.hdc, x + 130, y + 70, 20, RopeP, RopeW, 0, 0, vbSrcAnd
                BitBlt Form1.hdc, x + 130, y + 70, 20, RopeP, RopeB, 0, 0, vbSrcPaint
                If Direction >= -4 Then
                    For G = 0 To 10
                        If CollisionDetect(x + 130, y + 60, picRope, GuyX(G), 785, picGuy, picBlank) = True Then
                            GuyClimb(G) = True
                            Exit For
                        End If
                    Next
                End If
            Else
                BitBlt Form1.hdc, x + 70, y + 70, 20, RopeP, RopeW, 0, 0, vbSrcAnd
                BitBlt Form1.hdc, x + 70, y + 70, 20, RopeP, RopeB, 0, 0, vbSrcPaint
                If Direction <= 4 Then
                    For G = 0 To 10
                        If CollisionDetect(x + 70, y + 60, picRope, GuyX(G), 785, picGuy, picBlank) = True Then
                            GuyClimb(G) = True
                            Exit For
                        End If
                    Next
                End If
            End If
        End If
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
