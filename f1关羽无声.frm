VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "����˧��ս-���𹥻�����"
   ClientHeight    =   10560
   ClientLeft      =   11745
   ClientTop       =   3495
   ClientWidth     =   15015
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10560
   ScaleWidth      =   15015
   ShowInTaskbar   =   0   'False
   Begin VB.Timer p2���� 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   2280
      Top             =   840
   End
   Begin VB.CommandButton Command2 
      Caption         =   "˫����Ϸ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   42
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   2520
      TabIndex        =   26
      Top             =   3720
      Width           =   8175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "������Ϸ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   42
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   2400
      TabIndex        =   25
      Top             =   720
      Width           =   8415
   End
   Begin VB.Timer ��� 
      Interval        =   5000
      Left            =   960
      Top             =   4440
   End
   Begin VB.Timer Timer6 
      Interval        =   18000
      Left            =   840
      Top             =   3240
   End
   Begin VB.Timer heliu 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   240
      Top             =   8760
   End
   Begin VB.Timer youhao 
      Interval        =   88
      Left            =   1200
      Top             =   9720
   End
   Begin VB.PictureBox MMControl7 
      Height          =   330
      Left            =   2280
      ScaleHeight     =   270
      ScaleWidth      =   3480
      TabIndex        =   22
      Top             =   8040
      Visible         =   0   'False
      Width           =   3540
   End
   Begin VB.PictureBox MMControl6 
      Height          =   375
      Left            =   2280
      ScaleHeight     =   315
      ScaleWidth      =   3480
      TabIndex        =   21
      Top             =   7440
      Visible         =   0   'False
      Width           =   3540
   End
   Begin VB.Timer �ؽ��� 
      Interval        =   500
      Left            =   240
      Top             =   8160
   End
   Begin VB.Timer Timer13 
      Interval        =   35000
      Left            =   1680
      Top             =   6720
   End
   Begin VB.PictureBox MMControl5 
      Height          =   495
      Left            =   2160
      ScaleHeight     =   435
      ScaleWidth      =   3480
      TabIndex        =   20
      Top             =   6720
      Visible         =   0   'False
      Width           =   3540
   End
   Begin VB.Timer Timer12 
      Interval        =   5000
      Left            =   1440
      Top             =   6720
   End
   Begin VB.Timer dimai 
      Interval        =   3000
      Left            =   480
      Top             =   9720
   End
   Begin VB.Timer Timer7 
      Interval        =   10000
      Left            =   840
      Top             =   3840
   End
   Begin VB.Timer Timer5 
      Interval        =   400
      Left            =   840
      Top             =   2640
   End
   Begin VB.Timer Timer4 
      Interval        =   4
      Left            =   840
      Top             =   2040
   End
   Begin VB.Timer Timer3 
      Interval        =   1001
      Left            =   840
      Top             =   1440
   End
   Begin VB.Timer Timer2 
      Interval        =   2999
      Left            =   720
      Top             =   840
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   720
      Top             =   240
   End
   Begin VB.Timer ��ʾ 
      Interval        =   5000
      Left            =   240
      Top             =   7440
   End
   Begin VB.Timer ���Ʊ߽� 
      Interval        =   20
      Left            =   240
      Top             =   6840
   End
   Begin VB.Timer ���ָ���ƣ 
      Interval        =   1000
      Left            =   240
      Top             =   6240
   End
   Begin VB.Timer ����ǰ 
      Interval        =   1000
      Left            =   240
      Top             =   5640
   End
   Begin VB.Timer ��ɱ�� 
      Interval        =   900
      Left            =   240
      Top             =   5040
   End
   Begin VB.Timer �о�����ǰ 
      Left            =   240
      Top             =   4440
   End
   Begin VB.Timer �о������� 
      Interval        =   777
      Left            =   240
      Top             =   3840
   End
   Begin VB.Timer �о������� 
      Left            =   240
      Top             =   3240
   End
   Begin VB.Timer �г��� 
      Interval        =   40000
      Left            =   240
      Top             =   2640
   End
   Begin VB.Timer �������� 
      Left            =   240
      Top             =   2040
   End
   Begin VB.Timer ��ɱ�� 
      Interval        =   900
      Left            =   240
      Top             =   1440
   End
   Begin VB.Timer man 
      Interval        =   11000
      Left            =   240
      Top             =   840
   End
   Begin VB.Timer kuai 
      Interval        =   87
      Left            =   240
      Top             =   240
   End
   Begin VB.PictureBox MMControl4 
      Height          =   495
      Left            =   2040
      ScaleHeight     =   435
      ScaleWidth      =   3555
      TabIndex        =   19
      Top             =   6000
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.PictureBox MMControl3 
      Height          =   495
      Left            =   2040
      ScaleHeight     =   435
      ScaleWidth      =   3555
      TabIndex        =   18
      Top             =   5400
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.PictureBox MMControl2 
      Height          =   495
      Left            =   2040
      ScaleHeight     =   435
      ScaleWidth      =   3555
      TabIndex        =   17
      Top             =   4800
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.PictureBox MMControl1 
      Height          =   495
      Left            =   2040
      ScaleHeight     =   435
      ScaleWidth      =   3555
      TabIndex        =   16
      Top             =   4200
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.Image Image16 
      Height          =   1755
      Left            =   6600
      Picture         =   "f1��������.frx":0000
      Top             =   1920
      Visible         =   0   'False
      Width           =   1830
   End
   Begin VB.Image Image15 
      Height          =   450
      Left            =   13320
      Picture         =   "f1��������.frx":08DE
      Top             =   1560
      Width           =   450
   End
   Begin VB.Image Image14 
      Height          =   405
      Left            =   12840
      Picture         =   "f1��������.frx":1A24
      Top             =   1560
      Width           =   420
   End
   Begin VB.Image Image13 
      Height          =   4800
      Left            =   1320
      Picture         =   "f1��������.frx":4B0C
      Top             =   1920
      Visible         =   0   'False
      Width           =   11640
   End
   Begin VB.Image Image12 
      Height          =   1425
      Left            =   7560
      Picture         =   "f1��������.frx":2C9DE
      Top             =   3480
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.Image Image11 
      Height          =   1500
      Left            =   9480
      Picture         =   "f1��������.frx":36F4A
      Top             =   2400
      Visible         =   0   'False
      Width           =   2400
   End
   Begin VB.Image Image9 
      Height          =   450
      Left            =   13200
      Picture         =   "f1��������.frx":3E4BE
      Top             =   120
      Width           =   450
   End
   Begin VB.Image Image8 
      Height          =   720
      Left            =   10920
      Picture         =   "f1��������.frx":3F1C0
      Top             =   480
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image Image2 
      Height          =   450
      Index           =   116
      Left            =   13320
      Picture         =   "f1��������.frx":4008A
      Top             =   1080
      Width           =   450
   End
   Begin VB.Image Image2 
      Height          =   465
      Index           =   115
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Image Image2 
      Height          =   465
      Index           =   114
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Image Image2 
      Height          =   465
      Index           =   113
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Image Image2 
      Height          =   465
      Index           =   112
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Image Image2 
      Height          =   465
      Index           =   111
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Image Image2 
      Height          =   465
      Index           =   110
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Image Image2 
      Height          =   465
      Index           =   109
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Image Image2 
      Height          =   465
      Index           =   108
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Image Image2 
      Height          =   465
      Index           =   107
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Image Image2 
      Height          =   465
      Index           =   106
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Image Image2 
      Height          =   465
      Index           =   105
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Image Image2 
      Height          =   465
      Index           =   104
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Image Image2 
      Height          =   465
      Index           =   103
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Image Image2 
      Height          =   465
      Index           =   102
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Image Image2 
      Height          =   465
      Index           =   101
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Image Image2 
      Height          =   465
      Index           =   100
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Image Image2 
      Height          =   465
      Index           =   99
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Image Image2 
      Height          =   465
      Index           =   98
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Image Image2 
      Height          =   465
      Index           =   97
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Image Image2 
      Height          =   465
      Index           =   96
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Image Image2 
      Height          =   465
      Index           =   95
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Image Image2 
      Height          =   465
      Index           =   94
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Image Image2 
      Height          =   465
      Index           =   93
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Image Image2 
      Height          =   465
      Index           =   92
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Image Image2 
      Height          =   465
      Index           =   91
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Image Image2 
      Height          =   465
      Index           =   90
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Image Image2 
      Height          =   465
      Index           =   89
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Image Image2 
      Height          =   465
      Index           =   88
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Image Image2 
      Height          =   465
      Index           =   87
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Image Image2 
      Height          =   465
      Index           =   86
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Image Image2 
      Height          =   465
      Index           =   85
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Image Image2 
      Height          =   465
      Index           =   84
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Image Image2 
      Height          =   465
      Index           =   83
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Image Image2 
      Height          =   465
      Index           =   82
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Image Image2 
      Height          =   465
      Index           =   81
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Image Image2 
      Height          =   465
      Index           =   80
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Image Image2 
      Height          =   465
      Index           =   79
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Image Image2 
      Height          =   465
      Index           =   78
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Image Image2 
      Height          =   465
      Index           =   77
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Image Image2 
      Height          =   465
      Index           =   76
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Image Image7 
      Height          =   7905
      Left            =   0
      Picture         =   "f1��������.frx":411D0
      Top             =   2280
      Visible         =   0   'False
      Width           =   15315
   End
   Begin VB.Image Image6 
      Height          =   450
      Left            =   13200
      Picture         =   "f1��������.frx":B5468
      Top             =   600
      Width           =   450
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "f��ǹ����4��ǹ���������ǵ�;�Ͻ�����150���ǳ�����1500�����2�����Ͽ�������¿���С��"
      Height          =   1215
      Left            =   720
      TabIndex        =   24
      Top             =   8520
      Width           =   3135
   End
   Begin VB.Image Image5 
      Height          =   405
      Left            =   12840
      Picture         =   "f1��������.frx":B74DA
      ToolTipText     =   "�׶�ģʽ"
      Top             =   1080
      Width           =   420
   End
   Begin VB.Image Image2 
      Height          =   390
      Index           =   75
      Left            =   2280
      Top             =   240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   30
      Index           =   74
      Left            =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Image Image2 
      Height          =   30
      Index           =   73
      Left            =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Image Image2 
      Height          =   30
      Index           =   72
      Left            =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Image Image2 
      Height          =   30
      Index           =   71
      Left            =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Image Image2 
      Height          =   30
      Index           =   70
      Left            =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Image Image2 
      Height          =   30
      Index           =   69
      Left            =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Image Image2 
      Height          =   30
      Index           =   68
      Left            =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Image Image2 
      Height          =   30
      Index           =   67
      Left            =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Image Image2 
      Height          =   615
      Index           =   66
      Left            =   0
      Picture         =   "f1��������.frx":BA5C2
      Top             =   120
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Image Image2 
      Height          =   615
      Index           =   65
      Left            =   0
      Picture         =   "f1��������.frx":BBC8A
      Top             =   120
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Image Image2 
      Height          =   810
      Index           =   64
      Left            =   0
      Picture         =   "f1��������.frx":BD352
      Top             =   120
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Image Image2 
      Height          =   570
      Index           =   63
      Left            =   0
      Picture         =   "f1��������.frx":BEA1A
      Top             =   120
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Image Image2 
      Height          =   615
      Index           =   62
      Left            =   0
      Picture         =   "f1��������.frx":C00E2
      Top             =   120
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.Image Image2 
      Height          =   615
      Index           =   61
      Left            =   0
      Picture         =   "f1��������.frx":C17AA
      Top             =   120
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Image Image2 
      Height          =   795
      Index           =   60
      Left            =   0
      Picture         =   "f1��������.frx":C2E72
      Top             =   120
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Image Image2 
      Height          =   555
      Index           =   59
      Left            =   0
      Picture         =   "f1��������.frx":C453A
      Top             =   120
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Image Image2 
      Height          =   600
      Index           =   58
      Left            =   0
      Picture         =   "f1��������.frx":C5C02
      Top             =   120
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Image Image2 
      Height          =   945
      Index           =   57
      Left            =   0
      Picture         =   "f1��������.frx":C6AC2
      Top             =   120
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Image Image2 
      Height          =   600
      Index           =   56
      Left            =   0
      Picture         =   "f1��������.frx":C79A2
      Top             =   120
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Image Image2 
      Height          =   915
      Index           =   55
      Left            =   0
      Picture         =   "f1��������.frx":C8862
      Top             =   120
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Image Image2 
      Height          =   600
      Index           =   54
      Left            =   0
      Picture         =   "f1��������.frx":C9722
      Top             =   120
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.Image Image2 
      Height          =   675
      Index           =   53
      Left            =   0
      Picture         =   "f1��������.frx":CA3E0
      Top             =   120
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Image Image2 
      Height          =   675
      Index           =   52
      Left            =   0
      Picture         =   "f1��������.frx":CB0BE
      Top             =   120
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Image Image2 
      Height          =   30
      Index           =   51
      Left            =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Image Image2 
      Height          =   30
      Index           =   50
      Left            =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Image Image2 
      Height          =   600
      Index           =   49
      Left            =   0
      Picture         =   "f1��������.frx":CBD7C
      Top             =   120
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.Image Image2 
      Height          =   930
      Index           =   48
      Left            =   3240
      Picture         =   "f1��������.frx":CCA3A
      Top             =   1080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image Image2 
      Height          =   690
      Index           =   47
      Left            =   3960
      Picture         =   "f1��������.frx":CDC3C
      Top             =   1080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image Image2 
      Height          =   495
      Index           =   46
      Left            =   5520
      Picture         =   "f1��������.frx":CEE3E
      Top             =   1680
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Image Image2 
      Height          =   495
      Index           =   45
      Left            =   3960
      Picture         =   "f1��������.frx":D0040
      Top             =   2040
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Image Image2 
      Height          =   795
      Index           =   44
      Left            =   3840
      Picture         =   "f1��������.frx":D1242
      Top             =   840
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Image Image2 
      Height          =   555
      Index           =   43
      Left            =   4440
      Picture         =   "f1��������.frx":D2444
      Top             =   1080
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Image Image2 
      Height          =   495
      Index           =   42
      Left            =   4680
      Picture         =   "f1��������.frx":D3646
      Top             =   1080
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.Image Image2 
      Height          =   495
      Index           =   41
      Left            =   5280
      Picture         =   "f1��������.frx":D4848
      Top             =   1320
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.Image Image2 
      Height          =   30
      Index           =   40
      Left            =   5520
      Top             =   1560
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Image Image2 
      Height          =   30
      Index           =   39
      Left            =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Image Image2 
      Height          =   30
      Index           =   38
      Left            =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Image Image2 
      Height          =   30
      Index           =   37
      Left            =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Image Image2 
      Height          =   30
      Index           =   36
      Left            =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Image Image4 
      Height          =   1275
      Index           =   0
      Left            =   6480
      Picture         =   "f1��������.frx":D5A4A
      Top             =   360
      Visible         =   0   'False
      Width           =   1650
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   3360
      TabIndex        =   23
      Top             =   360
      Width           =   7815
   End
   Begin VB.Image Image4 
      Height          =   360
      Index           =   8
      Left            =   12960
      Top             =   8280
      Width           =   135
   End
   Begin VB.Image Image4 
      Height          =   360
      Index           =   7
      Left            =   13200
      Top             =   8880
      Width           =   135
   End
   Begin VB.Image Image4 
      Height          =   360
      Index           =   5
      Left            =   12960
      Top             =   8880
      Width           =   135
   End
   Begin VB.Image Image3 
      Height          =   3345
      Left            =   11280
      Picture         =   "f1��������.frx":D6328
      Top             =   6600
      Width           =   3375
   End
   Begin VB.Image Image4 
      Height          =   360
      Index           =   2
      Left            =   600
      Top             =   390
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   345
      Index           =   200
      Left            =   960
      Top             =   720
      Width           =   345
   End
   Begin VB.Image Image1 
      Height          =   345
      Index           =   0
      Left            =   1560
      Top             =   960
      Width           =   345
   End
   Begin VB.Image Image2 
      Height          =   465
      Index           =   35
      Left            =   2640
      Picture         =   "f1��������.frx":1D9D14
      Top             =   2400
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Image Image2 
      Height          =   345
      Index           =   34
      Left            =   3000
      Picture         =   "f1��������.frx":1DC042
      Top             =   1800
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image Image2 
      Height          =   345
      Index           =   33
      Left            =   3000
      Picture         =   "f1��������.frx":1DC5B0
      Top             =   1800
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image Image2 
      Height          =   345
      Index           =   32
      Left            =   3000
      Picture         =   "f1��������.frx":1DCA9E
      Top             =   1800
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image Image2 
      Height          =   345
      Index           =   31
      Left            =   3000
      Picture         =   "f1��������.frx":1DD00C
      Top             =   1800
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image Image2 
      Height          =   570
      Index           =   30
      Left            =   3000
      Picture         =   "f1��������.frx":1ED7C4
      Top             =   1800
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.Image Image2 
      Height          =   570
      Index           =   29
      Left            =   3000
      Picture         =   "f1��������.frx":201254
      Top             =   1800
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Image Image2 
      Height          =   570
      Index           =   28
      Left            =   3000
      Picture         =   "f1��������.frx":211D64
      Top             =   1800
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image2 
      Height          =   810
      Index           =   27
      Left            =   3000
      Picture         =   "f1��������.frx":2225F4
      Top             =   1800
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Image Image2 
      Height          =   345
      Index           =   26
      Left            =   3000
      Picture         =   "f1��������.frx":246EE0
      Top             =   1800
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image Image2 
      Height          =   570
      Index           =   25
      Left            =   3000
      Picture         =   "f1��������.frx":24744E
      Top             =   1800
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.Image Image2 
      Height          =   345
      Index           =   24
      Left            =   3000
      Picture         =   "f1��������.frx":25AEDE
      Top             =   1800
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image Image2 
      Height          =   570
      Index           =   23
      Left            =   3000
      Picture         =   "f1��������.frx":25B3CC
      Top             =   1800
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Image Image2 
      Height          =   345
      Index           =   22
      Left            =   3000
      Picture         =   "f1��������.frx":26BEDC
      Top             =   1800
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image Image2 
      Height          =   570
      Index           =   21
      Left            =   3000
      Picture         =   "f1��������.frx":26C44A
      Top             =   1800
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image2 
      Height          =   345
      Index           =   20
      Left            =   3000
      Picture         =   "f1��������.frx":27CCDA
      Top             =   1800
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image Image2 
      Height          =   810
      Index           =   19
      Left            =   3000
      Picture         =   "f1��������.frx":282356
      Top             =   1800
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Image Image2 
      Height          =   570
      Index           =   18
      Left            =   3000
      Picture         =   "f1��������.frx":2A6C42
      Top             =   1800
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Image Image2 
      Height          =   360
      Index           =   17
      Left            =   3000
      Picture         =   "f1��������.frx":2A960A
      Top             =   1800
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Image Image2 
      Height          =   570
      Index           =   16
      Left            =   3000
      Picture         =   "f1��������.frx":2ADEBA
      Top             =   1800
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Image Image2 
      Height          =   315
      Index           =   15
      Left            =   3000
      Picture         =   "f1��������.frx":2B00DA
      Top             =   1800
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Image Image2 
      Height          =   345
      Index           =   14
      Left            =   3000
      Picture         =   "f1��������.frx":2B44CE
      Top             =   1800
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image Image2 
      Height          =   345
      Index           =   13
      Left            =   3000
      Picture         =   "f1��������.frx":2B4A3C
      Top             =   1800
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image Image2 
      Height          =   345
      Index           =   12
      Left            =   3240
      Picture         =   "f1��������.frx":2B4F2A
      Top             =   2400
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image Image2 
      Height          =   270
      Index           =   11
      Left            =   3000
      Picture         =   "f1��������.frx":2B64A2
      Top             =   1800
      Visible         =   0   'False
      Width           =   30
   End
   Begin VB.Image Image2 
      Height          =   30
      Index           =   10
      Left            =   3000
      Picture         =   "f1��������.frx":2B6762
      Top             =   1800
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Image Image2 
      Height          =   345
      Index           =   9
      Left            =   3000
      Top             =   1800
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image Image2 
      Height          =   345
      Index           =   8
      Left            =   3000
      Top             =   1800
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image Image2 
      Height          =   570
      Index           =   7
      Left            =   3000
      Picture         =   "f1��������.frx":2B6A22
      Top             =   1800
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Image Image2 
      Height          =   345
      Index           =   6
      Left            =   3000
      Picture         =   "f1��������.frx":2B8BB2
      Top             =   1800
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image Image2 
      Height          =   570
      Index           =   5
      Left            =   3000
      Picture         =   "f1��������.frx":2B90A0
      Top             =   1800
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Image Image2 
      Height          =   345
      Index           =   4
      Left            =   3000
      Top             =   1800
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image Image2 
      Height          =   570
      Index           =   3
      Left            =   3000
      Picture         =   "f1��������.frx":2BBA30
      Top             =   1800
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Image Image2 
      Height          =   345
      Index           =   2
      Left            =   3000
      Picture         =   "f1��������.frx":2BD4B0
      Top             =   1800
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image Image2 
      Height          =   570
      Index           =   1
      Left            =   3000
      Picture         =   "f1��������.frx":2BE57E
      Top             =   1800
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   495
      Index           =   0
      Left            =   1920
      Top             =   1800
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   431
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   430
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   429
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   428
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   427
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   426
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   425
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   424
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   423
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   422
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   421
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   420
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   419
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   418
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   417
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   416
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   415
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   414
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   413
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   412
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   411
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   410
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   409
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   408
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   407
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   406
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   405
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   404
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   403
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   402
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   345
      Index           =   401
      Left            =   13560
      Top             =   8880
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   400
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   399
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   398
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   397
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   396
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   395
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   394
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   393
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   392
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   391
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   390
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   389
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   388
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   387
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   386
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   385
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   384
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   383
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   382
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   381
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   380
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   379
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   378
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   377
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   376
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   375
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   374
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   373
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   372
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   371
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   370
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   369
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   368
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   367
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   366
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   365
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   364
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   363
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   362
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   361
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   360
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   359
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   358
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   357
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   356
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   355
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   354
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   353
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   352
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   351
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   350
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   349
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   348
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   347
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   346
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   345
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   344
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   343
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   342
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   341
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   340
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   339
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   338
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   337
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   336
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   335
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   334
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   333
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   332
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   331
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   330
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   329
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   328
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   327
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   326
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   325
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   324
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   323
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   322
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   321
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   320
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   319
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   318
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   317
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   316
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   315
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   314
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   313
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   312
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   311
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   310
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   309
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   308
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   307
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   306
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   305
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   304
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   303
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   302
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   301
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   300
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   299
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   298
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   297
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   296
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   295
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   294
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   293
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   292
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   291
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   290
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   289
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   288
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   287
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   286
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   285
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   284
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   283
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   282
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   281
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   280
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   279
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   278
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   277
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   276
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   275
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   274
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   273
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   272
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   271
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   270
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   269
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   268
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   267
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   266
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   265
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   264
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   263
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   262
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   261
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   260
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   259
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   258
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   257
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   256
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   255
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   254
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   253
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   252
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   251
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   250
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   249
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   248
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   247
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   246
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   245
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   244
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   243
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   242
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   241
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   240
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   239
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   238
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   237
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   236
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   235
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   234
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   233
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   232
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   231
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   230
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   229
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   228
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   227
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   226
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   225
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   224
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   223
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   222
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   221
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   220
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   219
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   218
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   217
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   216
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   215
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   214
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   213
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   212
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   211
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   210
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   209
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   208
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   207
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   206
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   205
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   204
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   203
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   202
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   201
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   199
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   198
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   197
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   196
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   195
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   194
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   193
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   192
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   191
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   190
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   189
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   188
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   187
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   186
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   185
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   184
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   183
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   182
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   181
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   180
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   179
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   178
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   177
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   176
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   175
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   174
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   173
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   172
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   171
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   170
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   169
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   168
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   167
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   166
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   165
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   164
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   163
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   162
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   161
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   160
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   159
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   158
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   157
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   156
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   155
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   154
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   153
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   152
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   151
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   150
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   149
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   148
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   147
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   146
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   145
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   144
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   143
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   142
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   141
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   140
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   139
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   138
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   137
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   136
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   135
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   134
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   133
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   132
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   131
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   130
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   129
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   128
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   127
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   126
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   125
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   124
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   123
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   122
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   121
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   120
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   119
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   118
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   117
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   116
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   115
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   114
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   113
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   112
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   111
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   110
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   109
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   108
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   107
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   106
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   105
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   104
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   103
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   102
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   101
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   100
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   99
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   98
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   97
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   96
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   95
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   94
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   93
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   92
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   91
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   90
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   89
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   88
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   87
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   86
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   85
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   84
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   83
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   82
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   81
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   80
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   79
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   78
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   77
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   76
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   75
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   74
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   73
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   72
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   71
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   70
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   69
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   68
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   67
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   66
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   65
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   64
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   63
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   62
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   61
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   60
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   59
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   58
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   57
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   56
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   55
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   54
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   53
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   52
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   51
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   50
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   49
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   48
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   47
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   46
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   45
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   44
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   43
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   42
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   41
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   40
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   39
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   38
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   37
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   36
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   35
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   34
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   33
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   32
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   31
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   30
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   29
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   28
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   27
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   26
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   25
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   24
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   23
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   22
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   21
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   20
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   19
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   18
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   17
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   16
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   15
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   14
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   13
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   12
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   11
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   10
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   9
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   8
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   7
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   6
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   5
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   4
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   3
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   2
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   1
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Index           =   15
      Left            =   3360
      TabIndex        =   15
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Index           =   14
      Left            =   3360
      TabIndex        =   14
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Index           =   13
      Left            =   3360
      TabIndex        =   13
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Index           =   12
      Left            =   3360
      TabIndex        =   12
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Index           =   11
      Left            =   3360
      TabIndex        =   11
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Index           =   10
      Left            =   3360
      TabIndex        =   10
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Index           =   9
      Left            =   3360
      TabIndex        =   9
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Index           =   8
      Left            =   3360
      TabIndex        =   8
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Index           =   7
      Left            =   3360
      TabIndex        =   7
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Index           =   6
      Left            =   3360
      TabIndex        =   6
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Index           =   5
      Left            =   3360
      TabIndex        =   5
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Index           =   4
      Left            =   3360
      TabIndex        =   4
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Index           =   3
      Left            =   3360
      TabIndex        =   3
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Index           =   2
      Left            =   3360
      TabIndex        =   2
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Index           =   1
      Left            =   3360
      TabIndex        =   1
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   $"f1��������.frx":2BF5A6
      Height          =   8055
      Index           =   0
      Left            =   15
      TabIndex        =   0
      Top             =   0
      Width           =   360
   End
   Begin VB.Image Image10 
      Height          =   10995
      Left            =   6000
      Picture         =   "f1��������.frx":2BF65D
      Top             =   0
      Width           =   915
   End
   Begin VB.Image Image4 
      Height          =   10980
      Index           =   1
      Left            =   6000
      Picture         =   "f1��������.frx":2C0D47
      Top             =   0
      Width           =   900
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'��ʾ��ʿ���Ӽ�Ҫ��3a��image1�ͱ�������ƣ���֡�
Dim diming As Long
Dim p2 As Long
Dim zuobima As Long
Dim shuoming9 As Long
Dim q As Integer
Dim beijing As Long
Dim nimabi As Long
Dim w As Integer
Dim shuoming1 As Long
Dim e As Long

Dim mmt As Long
Dim �Ͻ����� As Long
Dim jiangji As Long
Dim a(20) As Long
Dim di(20) As Long
Dim x(20) As Long
Dim ������ As Long
Dim wsz As Long
Dim ����(500) As Long
Dim ��ƣ(500) As Long
Dim ����(500) As Long
Dim c(-6 To 200, -6 To 200) As Long 'c=0�����˵�,�Ѿ�ɾȥ ��С�Ҵ�
Dim zza As Long
Dim zzd As Long
Dim ssss As Long
Dim zzs As Long
Dim mmd As Long
Dim mmn As Long
Dim ab As Long
Dim mmb As Long
Dim mmc As Long
Dim qqx(-6 To 160) As Long
Dim qqy(-6 To 160) As Long
Dim qqq As Long
Dim zzx As Long
Dim zab As Long
Dim zppn As Long
Dim ppd As Long
Dim ppc As Long
Dim ppb As Long
Dim nigan1 As Long
Dim nigan2 As Long



Private Sub Command1_Click()
Image13.Visible = True
Command1.Visible = False
Command2.Visible = False
End Sub

Private Sub Command2_Click()
Command1.Visible = False
Command2.Visible = False
Timer7.Enabled = False
p2����.Enabled = True
youhao.Enabled = False
Image13.Visible = True
p2 = 2
End Sub

Private Sub dimai_Timer()

 If ��ƣ(401) > 50 And (Image1(0).Visible = True) And Image1(0).Left > 10000 And Image1(0).Top > 6000 Then '���ҷ���ϵͳ
Label2.Caption = Label2.Caption & "         �����ɳ�ʹ�ߣ�����������Զ�����ദ"
 Dim x1 As Long
 '��������
 For x1 = 201 To 400
 If Image1(x1).Visible = True Then 'ȫ���Լ�����
 ��ƣ(x1) = ��ƣ(x1) + 4
 End If
 Next
 


'��������12��ʼ
mmc = 0
zx = 9500
For mmd = 201 To 400 'ѡ���
   If (Image1(mmd).Visible = False) And (��ƣ(mmd) < 3) And (����(mmd) = 3) And (mmc < 12) Then '������Ҫ�� ���֣�����
   mmc = mmc + 1 '��������
  ����ǰ.Enabled = True
   '����λ�ÿ�ʼ  δ
mmb = 0
ab = 0
mmn = 0


mmn = qqx(mmc) * 200 '���ȸ��õ�
ab = qqy(mmc) * 200
'����λ�ý���
  
  di(12) = di(12) - 1 '����   ������Ҫ�� ����
  Image1(mmd).Visible = True
  Image1(mmd).Left = 14000 - mmn '�������
  Image1(mmd).Top = zx - ab '�����ϱ�
  ����(mmd) = 3

����ǰ.Enabled = True

End If '�ǻ����ִ�еĶ�������
  ����ǰ.Enabled = True
   Next

'�����������
'���������12��ʼ
mmc = 0
zx = 9500
For mmd = 201 To 400 'ѡ���
   If (Image1(mmd).Visible = False) And (��ƣ(mmd) < 3) And (����(mmd) = 2) And (mmc < 12) Then '������Ҫ�� ���֣�����
   mmc = mmc + 1 '��������
  ����ǰ.Enabled = True
   '����λ�ÿ�ʼ  δ
mmb = 0
ab = 0
mmn = 0

mmn = qqx(mmc) * 200 '���ȸ��õ�
ab = qqy(mmc) * 200
'����λ�ý���
  
  di(9) = di(9) - 1 '����   ������Ҫ�� ����
  Image1(mmd).Visible = True
  Image1(mmd).Left = 14000 - mmn '�������
  Image1(mmd).Top = zx - ab '�����ϱ�
  ����(mmd) = 3

����ǰ.Enabled = True

End If '�ǻ����ִ�еĶ�������
  ����ǰ.Enabled = True
   Next

'���󹭱�����


'���ҷ�������


End If

If Image1(0).Left >= 11900 And Image1(0).Top >= 7200 And Image1(0).Visible = True Then
'�ر�



'������Χ��




mmc = 0 '������Χ��ʼ
For mmd = 201 To 400 'ѡ���
   If (Image1(mmd).Visible = False) And (��ƣ(mmd) < 3) And (����(mmd) = 1) And (mmc < 1) Then '������Ҫ�� ���֣�����
   mmc = mmc + 1 '��������
di(6) = di(6) - 1 '����   ������Ҫ�� ����
  Image1(mmd).Visible = True
  Image1(mmd).Left = Image1(0).Left '�������
  Image1(mmd).Top = Image1(0).Top - 100 '�����ϱ�
  ����(mmd) = 2
End If '�ǻ����ִ�еĶ�������
 Next
'1������

mmc = 0
For mmd = 201 To 400 'ѡ���
   If (Image1(mmd).Visible = False) And (��ƣ(mmd) < 3) And (����(mmd) = 1) And (mmc < 1) Then '������Ҫ�� ���֣�����
   mmc = mmc + 1 '��������
di(6) = di(6) - 1 '����   ������Ҫ�� ����
  Image1(mmd).Visible = True
  Image1(mmd).Left = Image1(0).Left '�������
  Image1(mmd).Top = Image1(0).Top + 100 '�����ϱ�
  ����(mmd) = 1
End If '�ǻ����ִ�еĶ�������
 Next
'1������
mmc = 0
For mmd = 201 To 400 'ѡ���
   If (Image1(mmd).Visible = False) And (��ƣ(mmd) < 3) And (����(mmd) = 1) And (mmc < 1) Then '������Ҫ�� ���֣�����
   mmc = mmc + 1 '��������
di(6) = di(6) - 1 '����   ������Ҫ�� ����
  Image1(mmd).Visible = True
  Image1(mmd).Left = Image1(0).Left - 100 '�������
  Image1(mmd).Top = Image1(0).Top  '�����ϱ�
  ����(mmd) = 4
End If '�ǻ����ִ�еĶ�������
 Next
'1������
mmc = 0
For mmd = 201 To 400 'ѡ���
   If (Image1(mmd).Visible = False) And (��ƣ(mmd) < 3) And (����(mmd) = 1) And (mmc < 1) Then '������Ҫ�� ���֣�����
   mmc = mmc + 1 '��������
di(6) = di(6) - 1 '����   ������Ҫ�� ����
  Image1(mmd).Visible = True
  Image1(mmd).Left = Image1(0).Left + 100 '�������
  Image1(mmd).Top = Image1(0).Top '�����ϱ�
  ����(mmd) = 3
End If '�ǻ����ִ�еĶ�������
 Next
'1������

'������Χ����




Else
'�ָ�timer7.enabled=true
����ǰ.Enabled = True
�о�������.Enabled = True
�г���.Enabled = True


End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKey1 Then
Label2.Caption = Label2.Caption & "    ���Ϸ�����갴ť     "
End If



If KeyCode = vbKeyF1 Then
Label2.Caption = Label2.Caption & "��Ҫ�Ұ� f1"
End If
'                        p2p2p2p2p2p2p2p2p2p2p2p2p2p2p2 p2 p2 p2 p2 p2 p2 p2 p2 p2 p2 p2 p2 p2 p2 p2 p2 p2
If p2 = 2 Then
If KeyCode = vbKeyUp Then
����(201) = 1 + (����(201) Mod 4) 'p2�Ľ���
End If
If KeyCode = vbKeyDown And Image1(201).Visible = True Then 'p2�ı���
Dim p2wan As Long
For p2wan = 202 To 400
����(p2wan) = ����(201)
Next
End If

End If


For qqq = 1 To 6 '����λ�ð�����ʼ
qqx(qqq) = 1
qqy(qqq) = qqq
Next
For qqq = 7 To 12
qqx(qqq) = 2
qqy(qqq) = qqq - 6
Next
For qqq = 13 To 18
qqx(qqq) = 3
qqy(qqq) = qqq - 12
Next
For qqq = 19 To 24
qqx(qqq) = 4
qqy(qqq) = qqq - 18
Next
For qqq = 25 To 30
qqx(qqq) = 5
qqy(qqq) = qqq - 24
Next             '����λ�ð�������





End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKey8 And a(3) - 1000 > 0 Then 'ɱţ
If x(1) > 29 Then
a(3) = a(3) - 1000 '���ݿ�ʼ
Label2.Caption = Label2.Caption & "   ʿ�������˺ܸ���   "
Label2.Caption = Label2.Caption & "   ʿ�������˾���"
   For zx = 1 To 199
 If (��ƣ(zx) < 3) And (��ƣ(zx) > 0) Then
  ��ƣ(zx) = ��ƣ(zx) - 1
  End If
  Next '���ݽ���
x(1) = 0
End If
End If 'ɱţ����


If ������ = 0 Then


'    ���׿���                                                                                    ������  ��ʼ
 If KeyCode = vbKeyO And Shift = 2 And zuobima = 1 Then '                                                                         ������  ��ʼ
           '                                                                         ������  ��ʼ
 Timer6.Enabled = True '                                                                                          ������  ��ʼ
 Label2.Caption = Label2.Caption & "  �Ѿ��ɹ�����"
 
 End If
 
 

'zzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzz30 ������ʼzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzz

' �������30


     If KeyCode = vbKeyR Then
   If x(1) < 26 Then
   Label2.Caption = Label2.Caption & "���غ����ߣ�3�����һ�غ�"
   End If
   If x(1) > 29 Then '�������ݿ�ʼ
   mmc = 0 '��������
   zx = 800 '�����������λ��
   For mmd = 1 To 199 'ѡ���
   If (Image1(mmd).Visible = False) And (��ƣ(mmd) < 3) And (����(mmd) = 3) And (mmc < 30) Then '������Ҫ�� ���֣�����
   mmc = mmc + 1 '��������
   Timer3.Enabled = True '��ǰ
   '����λ�ÿ�ʼ  δ
mmb = 0
ab = 0
mmn = 0
If c(6, 6) = 1 Then '��һƬ������ô��
zx = 1600
End If
If c(6, 15) = 1 Then
zx = 2400

End If
mmn = qqx(mmc) * 100 '���ȸ��õ�
ab = qqy(mmc) * 100
'����λ�ý���
  a(2) = a(2) - 1 '����
  a(12) = a(12) - 1 '����   ������Ҫ�� ����
  Image1(mmd).Visible = True
  Image1(mmd).Left = 800 - mmn '�������
  Image1(mmd).Top = zx - ab '�����ϱ�

 x(1) = 0
End If '�ǻ����ִ�еĶ�������
   Next
  If mmc = 0 Then
 Label2.Caption = Label2.Caption & "�ޱ�"
   End If
   '�������ݽ���
   End If
   End If '�����30����









' ��������30


     If KeyCode = vbKeyT Then
   If x(1) < 26 Then
   Label2.Caption = Label2.Caption & "���غ����ߣ�3�����һ�غ�"
   End If
   If x(1) > 29 Then '�������ݿ�ʼ
   mmc = 0 '��������
   zx = 800 '�����������λ��
   For mmd = 1 To 199 'ѡ���
   If (Image1(mmd).Visible = False) And (��ƣ(mmd) < 3) And (����(mmd) = 2) And (mmc < 30) Then '������Ҫ�� ���֣�����
   mmc = mmc + 1 '��������
   Timer3.Enabled = True '��ǰ
   '����λ�ÿ�ʼ  δ
mmb = 0
ab = 0
mmn = 0
If c(6, 6) = 1 Then '��һƬ������ô��
zx = 1600
End If
If c(6, 15) = 1 Then
zx = 2400

End If
mmn = qqx(mmc) * 100 '���ȸ��õ�
ab = qqy(mmc) * 100
'����λ�ý���
  a(2) = a(2) - 1 '����
  a(9) = a(9) - 1 '����   ������Ҫ�� ����
  Image1(mmd).Visible = True
  Image1(mmd).Left = 800 - mmn '�������
  Image1(mmd).Top = zx - ab '�����ϱ�
'������ʼ

 x(1) = 0

End If '�ǻ����ִ�еĶ�������
   Next
  If mmc = 0 Then
 Label2.Caption = Label2.Caption & "�ޱ�"
   End If
   '�������ݽ���
   End If
   End If '������30����










' ����ǹ��30


     If KeyCode = vbKeyY Then
   If x(1) < 26 Then
   Label2.Caption = Label2.Caption & "���غ����ߣ�3�����һ�غ�"
   End If
   If x(1) > 29 Then '�������ݿ�ʼ
   mmc = 0 '��������
   zx = 800 '�����������λ��
   For mmd = 1 To 199 'ѡ���
   If (Image1(mmd).Visible = False) And (��ƣ(mmd) < 3) And (����(mmd) = 1) And (mmc < 30) Then '������Ҫ�� ���֣�����
   mmc = mmc + 1 '��������
   Timer3.Enabled = True '��ǰ
   '����λ�ÿ�ʼ  δ
mmb = 0
ab = 0
mmn = 0
If c(6, 6) = 1 Then '��һƬ������ô��
zx = 1600
End If
If c(6, 15) = 1 Then
zx = 2400

End If
mmn = qqx(mmc) * 100 '���ȸ��õ�
ab = qqy(mmc) * 100
'����λ�ý���
  a(2) = a(2) - 1 '����
  a(6) = a(6) - 1 '����   ������Ҫ�� ����
  Image1(mmd).Visible = True
  Image1(mmd).Left = 800 - mmn '�������
  Image1(mmd).Top = zx - ab '�����ϱ�
 '������ʼ

 x(1) = 0

End If '�ǻ����ִ�еĶ�������
   Next
  If mmc = 0 Then
 Label2.Caption = Label2.Caption & "�ޱ�"
   End If
   '�������ݽ���
   End If
   End If '��ǹ��30����

'zzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzz30��������zzzzzzzzzzzzzzzzzzzzzzzzzz





'��������ʼ


' ����ǹ��1

If KeyCode = vbKeyU Then
If x(1) < 26 Then
Label2.Caption = Label2.Caption & "���غ����ߣ�3�����һ�غ�"
End If
If x(1) > 29 Then '�������ݿ�ʼ
mmc = 0 '��������
     zx = 1600 '�����������λ��
For mmd = 1 To 199 'ѡ���
If (Image1(mmd).Visible = False) And (��ƣ(mmd) < 3) And (����(mmd) = 1) And (mmc < 1) Then '������Ҫ�� ���֣�����
 mmc = mmc + 1 '��������
 '����λ�ÿ�ʼ  δ
mmb = 0
ab = 0
mmn = 0
'����λ�ý���
a(2) = a(2) - 1 '����
a(6) = a(6) - 1 '����   ������Ҫ�� ����
Image1(mmd).Visible = True
 Image1(mmd).Left = 1400 + mmn '�������
 Image1(mmd).Top = zx + ab '�����ϱ�

x(1) = 0
x(1) = 0
 '��������
    Timer3.Enabled = True '��ǰ����
  End If '�ǻ����ִ�еĶ�������
  Next
  If mmc = 0 Then
 Label2.Caption = Label2.Caption & "�ޱ�"
End If

 '�������ݽ���
End If
End If '��ǹ��1����

' ��������1

If KeyCode = vbKeyI Then
If x(1) < 26 Then
Label2.Caption = Label2.Caption & "���غ����ߣ�3�����һ�غ�"
End If
If x(1) > 29 Then '�������ݿ�ʼ
mmc = 0 '��������
     zx = 1600 '�����������λ��
For mmd = 1 To 199 'ѡ���
If (Image1(mmd).Visible = False) And (��ƣ(mmd) < 3) And (����(mmd) = 2) And (mmc < 1) Then '������Ҫ�� ���֣�����
 mmc = mmc + 1 '��������
 '����λ�ÿ�ʼ  δ
mmb = 0
ab = 0
mmn = 0
'����λ�ý���
a(2) = a(2) - 1 '����
a(9) = a(9) - 1 '����  ������Ҫ�� ����
Image1(mmd).Visible = True
 Image1(mmd).Left = 1400 + mmn '�������
 Image1(mmd).Top = zx + ab '�����ϱ�

x(1) = 0
 '��������
    Timer3.Enabled = True '��ǰ����
  End If '�ǻ����ִ�еĶ�������
 Next
 If mmc = 0 Then
 Label2.Caption = Label2.Caption & " �ޱ�"
End If

  '�������ݽ���
End If
End If '������1����

' �������1


If KeyCode = vbKeyO Then
If x(1) < 26 Then
Label2.Caption = Label2.Caption & "   ���غ����ߣ�3�����һ�غ�"
End If
If x(1) > 29 Then '�������ݿ�ʼ
mmc = 0 '��������
     zx = 1600 '�����������λ��
For mmd = 1 To 199 'ѡ���
If (Image1(mmd).Visible = False) And (��ƣ(mmd) < 3) And (����(mmd) = 3) And (mmc < 1) Then '������Ҫ�� ���֣�����
 mmc = mmc + 1 '��������
 '����λ�ÿ�ʼ  δ
mmb = 0
ab = 0
mmn = 0
'����λ�ý���
a(2) = a(2) - 1 '����
a(12) = a(12) - 1 '����  ������Ҫ�� ����
Image1(mmd).Visible = True
 Image1(mmd).Left = 1400 + mmn '�������
 Image1(mmd).Top = zx + ab '�����ϱ�

x(1) = 0
 '��������
    Timer3.Enabled = True '��ǰ����
  End If '�ǻ����ִ�еĶ�������
  Next
  If mmc = 0 Then
 Label2.Caption = Label2.Caption & "�ޱ�"
End If

 '�������ݽ���
End If
End If '�����1����
'����������



If KeyCode = vbKey2 And a(4) - 5000 > 0 Then '����ũ��
If x(1) > 29 Then
a(4) = a(4) - 5000
a(14) = a(14) + 1
x(1) = 0
End If
End If '����ũ�����
If KeyCode = vbKey7 And a(3) - 3000 > 0 Then '��������
If x(1) > 29 Then
a(3) = a(3) - 3000
a(1) = a(1) + 1 '����
a(2) = a(2) + 5 '����
x(1) = 0
End If
End If '�������ս���
If KeyCode = vbKey3 And a(4) - 5000 > 0 Then '������ҵ
If x(1) > 29 Then
a(4) = a(4) - 5000
a(15) = a(15) + 50
x(1) = 0
End If
End If '������ҵ����



'��������ʼ


'��1ǹ��

If KeyCode = vbKeyJ Then
If x(1) < 26 Then
Label2.Caption = Label2.Caption & "  ���غ����ߣ�3�����һ�غ�"
End If
If x(1) > 27 And a(2) > 0 Then '����һ��
x(1) = 0
tt = 0
For zx = 1 To 199 '�б�
If (��ƣ(zx) > 2) And (Image1(zx).Visible = False) And (tt < 1) And a(4) - (20000 - a(1)) > 0 Then
tt = 1
 Label2.Caption = Label2.Caption & "  �Ѿ�����һ��ʿ��"
����(zx) = 1 '                        '������Ҫ�ı���
��ƣ(zx) = 0 '��ƣ0������� �ڳ���
a(4) = a(4) - (20000 - a(1))
a(6) = a(6) + 1 '                    '������Ҫ�ı���
End If
Next
If tt = 0 Then
Label2.Caption = Label2.Caption & "   ���ڲ����б�"
End If
End If '����һ��ĩ
End If '��һǹ��ĩ


 '��1����
 
 
 If KeyCode = vbKeyK Then
If x(1) < 26 Then
Label2.Caption = Label2.Caption & " ���غ����ߣ�3�����һ�غ�"
End If
If x(1) > 27 And a(2) > 0 Then '����һ��
x(1) = 0
tt = 0
For zx = 1 To 199 '�б�
If (��ƣ(zx) > 2) And (Image1(zx).Visible = False) And (tt < 1) And a(4) - (20000 - a(1)) > 0 Then
tt = 1
 Label2.Caption = Label2.Caption & " �Ѿ�����һ��ʿ��"
����(zx) = 2 '                        '������Ҫ�ı���
��ƣ(zx) = 0 '��ƣ0������� �ڳ���
a(4) = a(4) - (20000 - a(1))
a(9) = a(9) + 1 '                    '������Ҫ�ı���
End If
Next
If tt = 0 Then
Label2.Caption = Label2.Caption & " ���ڲ����б�"
End If
End If '����һ��ĩ
End If '��һ����ĩ


'��1qi��


If KeyCode = vbKeyL Then
If x(1) < 26 Then
Label2.Caption = Label2.Caption & " ���غ����ߣ�3�����һ�غ�"
End If
If x(1) > 27 And a(2) > 0 Then '����һ��
x(1) = 0
tt = 0
For zx = 1 To 199 '�б�
If (��ƣ(zx) > 2) And (Image1(zx).Visible = False) And (tt < 1) And a(4) - (20000 - a(1)) > 0 Then
tt = 1
 Label2.Caption = Label2.Caption & "  �Ѿ�����һ��ʿ��"
����(zx) = 3
��ƣ(zx) = 0 '��ƣ0������� �ڳ���
a(4) = a(4) - (20000 - a(1))
a(12) = a(12) + 1
End If
Next
If tt = 0 Then
Label2.Caption = Label2.Caption & " ���ڲ����б�"
End If
End If '����һ��ĩ
End If '��һ���ĩ
'����������

'�������ʼ

'��5ǹ��

If KeyCode = vbKeyF Then '                        '������Ҫ�ı���
If x(1) < 26 Then
Label2.Caption = Label2.Caption & "   ���غ����ߣ�3�����һ�غ�"
End If
If x(1) > 27 And a(2) > 0 Then '����һ��
x(1) = 0
tt = 0
For zx = 1 To 199 '�б�ѭ��
If (��ƣ(zx) > 2) And (Image1(zx).Visible = False) And (tt < 5) And a(4) - (20000 - a(1)) > 0 Then
tt = 1 + tt
 Label2.Caption = Label2.Caption & "   �Ѿ�����һ��ʿ��"
����(zx) = 1 '                        '������Ҫ�ı���
��ƣ(zx) = 0 '��ƣ0������� �ڳ���
a(4) = a(4) - (20000 - a(1))
a(6) = a(6) + 1 '                    '������Ҫ�ı���
End If
Next '�б�ѭ��ĩ
If tt = 0 Then
Label2.Caption = Label2.Caption & "  ���ڲ����б�"
End If
End If '����һ��ĩ
End If '��5ǹ��ĩ


 '��5���� '
 
 If KeyCode = vbKeyG Then                        '������Ҫ�ı���
If x(1) < 26 Then
Label2.Caption = Label2.Caption & "  ���غ����ߣ�3�����һ�غ�"
End If
If x(1) > 27 And a(2) > 0 Then '����һ��
x(1) = 0
tt = 0
For zx = 1 To 199 '�б�ѭ��
If (��ƣ(zx) > 2) And (Image1(zx).Visible = False) And (tt < 5) And a(4) - (20000 - a(1)) > 0 Then
tt = 1 + tt
 Label2.Caption = Label2.Caption & "   �Ѿ�����һ��ʿ��"
����(zx) = 2 '                        '������Ҫ�ı���
��ƣ(zx) = 0 '��ƣ0������� �ڳ���
a(4) = a(4) - (20000 - a(1))
a(9) = a(9) + 1 '                    '������Ҫ�ı���
End If
Next '�б�ѭ��ĩ
If tt = 0 Then
Label2.Caption = Label2.Caption & "   ���ڲ����б�"
End If
End If '����һ��ĩ
End If '��5����ĩ

 '��5��� '
 
 If KeyCode = vbKeyH Then                        '������Ҫ�ı���
If x(1) < 26 Then
Label2.Caption = Label2.Caption & "  ���غ����ߣ�3�����һ�غ�"
End If
If x(1) > 27 And a(2) > 0 Then '����һ��
x(1) = 0
tt = 0
For zx = 1 To 199 '�б�ѭ��
If (��ƣ(zx) > 2) And (Image1(zx).Visible = False) And (tt < 5) And a(4) - (20000 - a(1)) > 0 Then
tt = 1 + tt
 Label2.Caption = Label2.Caption & "   �Ѿ�����һ��ʿ��"
����(zx) = 3 '                        '������Ҫ�ı���
��ƣ(zx) = 0 '��ƣ0������� �ڳ���
a(4) = a(4) - (20000 - a(1))
a(12) = a(12) + 1 '                    '������Ҫ�ı���
End If
Next '�б�ѭ��ĩ
If tt = 0 Then
Label2.Caption = Label2.Caption & "   ���ڲ����б�"
End If
End If '����һ��ĩ
End If '��5���ĩ
'���������








If KeyCode = vbKeyReturn Then '���Ͻ�
������ = 1
Image1(0).Visible = True
Image1(0).Left = 700
Image1(0).Top = 700
 ����(0) = 0
����(0) = 1
End If






End If '��������  ����


If ������ = 1 Then
'����                                                            ������ ��ʼ
'����                                                            ������ ��ʼ
'����                                                            ������ ��ʼ
'����                                                            ������ ��ʼ
'����                                                            ������ ��ʼ
'����                                                            ������ ��ʼ
 
 If KeyCode = vbKey0 Then '��������
 For x(4) = 1 To 199
 If Image1(x(4)).Visible = True Then
 ��ƣ(x(4)) = ��ƣ(x(4)) + 2
 End If
 Next
 Image1(0).Visible = False
 ������ = 0
 End If

If KeyCode = vbKeyReturn And Shift = 1 And Image1(0).Left < 1000 And Image1(0).Top < 1000 Then 'jin�Ͻ�

zuobima = 1
Label2.Caption = Label2.Caption & "    ����Կ�����,shifttwo Qq1Xiang24"
������ = 0
Image1(0).Visible = False


End If

Dim zz As Long
If KeyCode = vbKeyQ Then '���ҷ���

For zz = 1 To 199
����(zz) = ����(0)
Next
End If
'��ɱ��
If KeyCode = vbKeyP Then
Image1(0).Left = Image1(0).Left + 4000
End If


If KeyCode = vbKeySpace Then   '��ɱ��'��ɱ��And jiangji = 0
 jiangji = 1

      If ����(0) = 1 Then
         Image1(0).Picture = Image2(31) '����
    End If
 If ����(0) = 2 Then
 Image1(0).Picture = Image2(33)
 End If
 If ����(0) = 3 Then
 Image1(0).Picture = Image2(34)
 End If
 If ����(0) = 4 Then
 Image1(0).Picture = Image2(32)
 End If
Dim ss As Long
Dim ssa As Long
Dim ssd As Long
Dim ssf As Long
Dim ssg As Long
 
 For ss = 201 To 401
 ssa = Image1(ss).Left
 ssd = Image1(ss).Top
 ssf = Image1(0).Left
 ssg = Image1(0).Top
If ssd <= ssg + 300 And ssd >= ssg - 300 And ssa <= ssf + 300 And ssa >= ssf - 300 Then    '��Χ�˱�����
 ��ƣ(ss) = ��ƣ(ss) + 1 '������1

 End If
 Next
End If '��ɱ��ĩ
If KeyCode = vbKeyE Then 'ȫ����ǰ
Timer3.Enabled = True
End If
If KeyCode = vbKeyX Then '����
Timer3.Enabled = False
End If
If (KeyCode = vbKeyD) And (c(Image1(0).Top / 100, (Image1(0).Left + 100) / 100) = 0) Then  '����������
����(0) = 4
Image1(0).Left = Image1(0).Left + 100
End If

If (KeyCode = vbKeyA) And (c(Image1(0).Top / 100, (Image1(0).Left - 100) / 100) = 0) Then
����(0) = 3
Image1(0).Left = Image1(0).Left - 100
End If

If (KeyCode = vbKeyS) And (c((Image1(0).Top + 100) / 100, (Image1(0).Left) / 100) = 0) Then
����(0) = 2
Image1(0).Top = Image1(0).Top + 100
End If

If (KeyCode = vbKeyW) And (c((Image1(0).Top - 100) / 100, (Image1(0).Left) / 100) = 0) Then
����(0) = 1
Image1(0).Top = Image1(0).Top - 100
End If '�������ҽ���







End If '                                                    ���������� ����




End Sub

Private Sub Form_Load() 'һ��ʼ��
Timer6.Enabled = False '���׹�
Dim tuxing As Long
For tuxing = 1 To 34
Image2(80 + tuxing).Picture = Image2(tuxing).Picture
Next





 '��������
Dim aaa As Long
beijing = 0

For aaa = 1 To 431
����(aaa) = 4
��ƣ(aaa) = 4

����(aaa) = 3
Image1(aaa).Left = 10000
Image1(aaa).Top = 5000
Image1(aaa).Picture = Image2(1).Picture '�Ͻ�����tu��
Image1(aaa).Visible = False
Next
For aaa = 1 To 200
����(aaa) = 4
Next
��ƣ(0) = 5
Label2.Caption = "   �Ͻ���ܲٸմ���̣����ۣ��������лָ�                �ҷ�����Ȩ�������Ѻù�ϵ       1�����ɿ�˵��     �밴�س���"
Image1(0).Picture = Image2(1).Picture '��һ�γ���Ҳ�ɱ��򣬽�ͼ����tu��
������ = 0
For aaa = 212 To 264 '�мӱ���ʼ
����(aaa) = 1
��ƣ(aaa) = 0
Next
For aaa = 265 To 328
����(aaa) = 2
��ƣ(aaa) = 0
Next
For aaa = 329 To 392
����(aaa) = 3
��ƣ(aaa) = 0
Next
di(6) = 53
di(9) = 64
di(12) = 64
 'di jia bing����


Dim i As Long

For i = 1 To 15
a(i) = 0
Label1(i).Enabled = True
Label1(i).Height = 255
Label1(i).Left = 900
Label1(i).Top = 540 * (i - 1)
Label1(i).Width = 1050
Next
a(2) = 100
a(4) = 400000
a(15) = 900
a(14) = 30
a(3) = 300
Image1(200).Visible = True '�� ����
Image1(200).Picture = Image2(35)
����(200) = 5
��ƣ(200) = 0
����(200) = 4
Image1(0).Visible = False
Image1(200).Left = 400
Image1(200).Top = 400
Image1(401).Visible = True
Image1(401).Picture = Image2(35).Picture
Image1(401).Top = 8800
Image1(401).Left = 13500
����(401) = 6






'һ��ʼ������ʼ

Image1(201).Visible = True
��ƣ(201) = 0
����(201) = 3
����(201) = 2
Image1(201).Top = 5000
Image1(201).Left = 7500
Dim f1 As Long
For f1 = 1 To 10
Image1(201 + f1).Visible = True
��ƣ(201 + f1) = 0
����(201 + f1) = 3
����(201 + f1) = 2
Image1(201 + f1).Left = 7500
Next
Image1(202).Top = 4000
Image1(203).Top = 3000
Image1(204).Top = 2000
Image1(205).Top = 1000
Image1(206).Top = 6000
Image1(207).Top = 7000
Image1(208).Top = 8000
Image1(209).Top = 9000
Image1(210).Top = 5500
Image1(211).Top = 4500


'һ��ʼ��������
End Sub






Private Sub heliu_Timer()
Dim f1 As Long
For f1 = 0 To 400
If Image1(f1).Visible = True And Image1(f1).Left <= 6500 And Image1(f1).Left >= 6300 Then
��ƣ(f1) = ��ƣ(f1) + 1
End If
Next

End Sub

Private Sub Image13_Click()
Timer1.Enabled = True
Image13.Visible = False
heliu.Enabled = True
End Sub

Private Sub Image14_Click()
Dim asni As Long
For asni = 201 To 400
Image1(asni).BorderStyle = 1
Next
End Sub

Private Sub Image15_Click()
Dim asni As Long
For asni = 201 To 400
Image1(asni).BorderStyle = 0
Next
End Sub

Private Sub Image2_Click(Index As Integer)
Dim fff As Long
For fff = 1 To 34
Image2(fff).Picture = Image2(80 + fff).Picture
Next
End Sub

Private Sub Image5_Click()
Dim fff As Long
For fff = 1 To 26
Image2(fff).Picture = Image2(fff + 40).Picture
Next

End Sub

Private Sub Image6_Click()

shuoming1 = shuoming1 + 1
Image7.Visible = True
If shuoming1 = 2 Then

Image7.Visible = False
shuoming1 = 0
End If
End Sub



Private Sub Image8_Click()
Image8.Picture = Image11.Picture
End Sub

Private Sub Image9_Click()
nimabi = nimabi + 1
shuoming9 = 0
If nimabi Mod 2 = 0 Then
For fff = 0 To 15
Label1(fff).Visible = True
Next
Label2.Visible = True
Label3.Visible = True

Else
For fff = 0 To 15
Label1(fff).Visible = False
Next
Label3.Visible = False
Label2.Visible = False
End If
End Sub

Private Sub kuai_Timer()

Dim fff As Long
For fff = 0 To 200
If Image1(fff).Visible = True And Image1(201).Visible = True And Image1(fff).Left + 1500 > Image1(201).Left And Image1(fff).Left - 1500 < Image1(201).Left And Image1(fff).Top + 1500 > Image1(201).Top And Image1(fff).Top - 1500 < Image1(201).Top Then
Image8.Visible = True '�Ż�
Image8.Top = Image1(fff).Top
Image8.Left = Image1(fff).Left '�˴��ظ��Ż�����δ���

End If
Next

For fff = 201 To 400 'fendiwoshuangfang

Next
If Image1(201).Visible = True Then
Image4(0).Left = Image1(201).Left + 200 '���˵��
Image4(0).Top = Image1(201).Top + 200
End If
If Image1(0).Visible = True Then
Image16.Top = -Image16.Height + Image1(0).Top '����˵��
Image16.Left = Image1(0).Left
Image16.Visible = True

End If
If Image1(0).Visible = False Then '����˵����ʧ
Image16.Visible = False
End If
Dim chengchi1 As Long '�ǳر���ʼ
Dim chengchi2 As Long
If chengchi1 > chengchi2 Then
Label2.Caption = Label2.Caption & "   �ҷ��ǳ���������"
End If
chengchi2 = chengchi1
chengchi1 = ��ƣ(200)
'�ǳر������

'������


nigan2 = nigan1
nigan1 = ��ƣ(0)
'���������

a(13) = ��ƣ(401)
Dim ssssl As Long
If Image1(0).Left + 1000 > Image1(401).Left And Image1(0).Top + 1000 > Image1(201).Top Then
Image1(401).Visible = True
Else: Image1(401).Visible = False

End If
Dim liang As Long
If a(3) > 0 Then
liang = 0
End If
If a(3) < 0 And liang = 0 Then '��ʳû��
Label2.Caption = Label2.Caption & "��ʳû�ˣ�ʿ���߹���"
liang = 1
For x(7) = 1 To 199
 If Image1(x(7)).Visible = True Then '�������
 ��ƣ(x(7)) = ��ƣ(x(7)) + 3
 End If
 Next
For x(7) = 1 To 199
 If Image1(x(7)).Visible = False And ��ƣ(x(7)) < 3 Then '�����б�
 ��ƣ(x(7)) = ��ƣ(x(7)) + 3
 a(����(x(7)) * 3 + 3) = a(����(x(7)) * 3 + 3) - 1
 End If
 Next
 End If
If ��ƣ(401) > 1500 Then
Image1(0).Enabled = False
Image1(0).Visible = False
Image1(401).Visible = False
Image1(401).Visible = False
If p2 = 0 Then
Label2.Caption = "�ɹ����Ƶ��˳ǳسɹ��ɹ��ɹ��ɹ��ɹ��ɹ��ɹ��ɹ�"
End If
If p2 = 2 Then
Label2.Caption = "����ʤ��"
End If
man.Enabled = False
kuai.Enabled = False
Timer7.Enabled = False

For ssssl = 1 To 400
Image1(ssssl).Visible = False
Image1(ssssl).Enabled = False
dimai.Enabled = False
�г���.Enabled = False
Next
Image1(200).Visible = True
End If
If ��ƣ(0) > 150 Or ��ƣ(200) > 1500 Then 'ʧ��ʧ��ʧ��ʧ��ʧ��ʧ��ʧ��ʧ��ʧ��ʧ��
Image1(0).Enabled = False
Image1(0).Visible = False
Image1(200).Visible = False
Image1(200).Visible = False
If p2 = 0 Then
Label2.Caption = "ʧ��ʧ��ʧ��ʧ��ʧ��ʧ��ʧ��ʧ��ʧ��ʧ�� �����ˣ��뿴���Լ�����"
End If
If p2 = 2 Then
Label2.Caption = "���ʤ��"

End If
man.Enabled = False
kuai.Enabled = False
Timer7.Enabled = False

For ssssl = 1 To 400
Image1(ssssl).Visible = False
Image1(ssssl).Enabled = False
dimai.Enabled = False
�г���.Enabled = False
Next
End If


For qq = 0 To 431
If ��ƣ(qq) > 2 And ����(qq) < 4 And ����(qq) > 0 Then
Image1(qq).Visible = False
End If
Next '��ƣ����ʧa(3) = a(3) + a(14)
End Sub

Private Sub man_Timer()

Dim fff As Long 'huoli
For fff = 0 To 400
If Image8.Visible = True And Image1(fff).Visible = True And Image1(fff).Left >= Image8.Left - 500 And Image1(fff).Left < Image8.Left + Image8.Width And Image1(fff).Top >= Image8.Top - 500 And Image1(fff).Top < Image8.Top + Image8.Height Then
��ƣ(fff) = ��ƣ(fff) + 85
End If
Next

Label2.Caption = ""
Image4(0).Visible = False

End Sub



Private Sub p2����_Timer()

tt = 0
For zx = 201 To 400 '�б�
If (��ƣ(zx) > 2) And (Image1(zx).Visible = False) And (tt < 5) Then
tt = 1 + tt
����(zx) = 1
��ƣ(zx) = 0 '��ƣ0������� �ڳ���
di(12) = di(12) + 1
End If
Next
tt = 0
For zx = 201 To 400 '�б�
If (��ƣ(zx) > 2) And (Image1(zx).Visible = False) And (tt < 5) Then
tt = 1 + tt
����(zx) = 3
��ƣ(zx) = 0 '��ƣ0������� �ڳ���
di(6) = di(6) + 1
End If
Next

tt = 0
For zx = 201 To 400 '�б�
If (��ƣ(zx) > 2) And (Image1(zx).Visible = False) And (tt < 5) Then
tt = 1 + tt
����(zx) = 2
��ƣ(zx) = 0 '��ƣ0������� �ڳ���
di(9) = di(9) + 1
End If
Next
Label2.Caption = Label2.Caption & "�����Ұ���        �����������褾�      ����ɱ���褱���" & di(6) & " " & di(9) & " " & di(12)
End Sub

Private Sub Timer1_Timer() '�Զ���




If x(1) Mod 2 = 0 Then
Image10.Visible = False
Image4(1).Visible = True
Else
Image10.Visible = True
Image4(1).Visible = False
End If
Dim a20 As Long
a(10) = ��ƣ(200)
a(7) = ��ƣ(0)
For a20 = 1 To 15
Label1(a20).Caption = a(a20)
Next
x(1) = x(1) + 1
Label1(7).Caption = a(7)


End Sub



Private Sub Timer10_Timer()

End Sub

Private Sub Timer11_Timer()

End Sub

Private Sub Timer12_Timer()
If beijing > 10 Then

 '��������
beijing = 0
End If
End Sub

Private Sub Timer13_Timer()
beijing = beijing + 1
End Sub


Private Sub Timer2_Timer() '���ģ�����
Dim chihe As Long
Dim chihe1 As Long
chihe1 = 0
For chihe = 1 To 199
If Image1(chihe).Visible = True Then
chihe1 = chihe1 + 1
End If
Next
a(4) = a(4) + a(15)
a(3) = a(3) - a(6) - a(9) - a(12) - 1 + a(14) - chihe1
End Sub



Private Sub Timer3_Timer() ' ��ǰ��

Dim iii As Long

For x(5) = 1 To 199
 If Image1(x(5)).Visible = True And ��ƣ(x(5)) < 3 Then
 
 If ����(x(5)) > 2 Then
   If ����(x(5)) = 4 Then
   iii = 1
   Else
   iii = -1
   End If
   If (Image1(x(5)).Visible = True) And (c(Image1(x(5)).Top / 100, (Image1(x(5)).Left + iii * 100) / 100) = 0) Then '��ǰ��û����
    If ����(x(5)) = 3 Then
    Image1(x(5)).Left = Image1(x(5)).Left + iii * 200
    Else
    Image1(x(5)).Left = Image1(x(5)).Left + iii * 100
    End If
   End If
End If
If ����(x(5)) < 3 Then
   If ����(x(5)) = 2 Then
   iii = 1
   Else: iii = -1
   End If
   If (Image1(x(5)).Visible = True) And (c(Image1(x(5)).Left / 100, (Image1(x(5)).Top + iii * 100) / 100) = 0) Then
    If ����(x(5)) = 3 Then
    Image1(x(5)).Top = Image1(x(5)).Top + iii * 200
    Else: Image1(x(5)).Top = Image1(x(5)).Top + iii * 100
    End If
   End If
End If
End If
Next
End Sub



Private Sub Timer4_Timer() '�Ƿ����0�ɣ��ڼ��У��ڼ���


For w = -5 To 120
For q = -5 To 160
c(q, w) = 0 '�ڼ��У��ڼ���
Next
Next
For e = 0 To 400
If (Image1(e).Visible = True) And (��ƣ(e) < 3) Then '������ֲ�������
c(Image1(e).Left / 100, Image1(e).Top / 100) = 1
End If
Next
End Sub






Private Sub Timer5_Timer() 'ת����ͼ
Dim oo As Long
For oo = 0 To 400
If ����(oo) = 1 Then
 If ����(oo) = 1 Then
  Image1(oo).Picture = Image2(3).Picture
End If
If ����(oo) = 2 Then
  Image1(oo).Picture = Image2(7).Picture
End If
If ����(oo) = 3 Then
  Image1(oo).Picture = Image2(5).Picture
End If
If ����(oo) = 4 Then
  Image1(oo).Picture = Image2(1).Picture
End If
End If


If ����(oo) = 2 Then
 If ����(oo) = 1 Then
  Image1(oo).Picture = Image2(15).Picture
End If
If ����(oo) = 2 Then
  Image1(oo).Picture = Image2(17).Picture
End If
If ����(oo) = 3 Then
  Image1(oo).Picture = Image2(16).Picture
End If
If ����(oo) = 4 Then
  Image1(oo).Picture = Image2(18).Picture
End If
End If


If ����(oo) = 3 Then
 If ����(oo) = 1 Then
  Image1(oo).Picture = Image2(19).Picture
End If
If ����(oo) = 2 Then
  Image1(oo).Picture = Image2(23).Picture
End If
If ����(oo) = 3 Then
  Image1(oo).Picture = Image2(25).Picture
End If
If ����(oo) = 4 Then
  Image1(oo).Picture = Image2(21).Picture
End If
End If


If ����(oo) = 0 Then
 If ����(oo) = 1 Then
  Image1(oo).Picture = Image2(27).Picture
End If
If ����(oo) = 2 Then
  Image1(oo).Picture = Image2(29).Picture
End If
If ����(oo) = 3 Then
  Image1(oo).Picture = Image2(30).Picture
End If
If ����(oo) = 4 Then
  Image1(oo).Picture = Image2(28).Picture
End If
End If
Next 'ת����ͼ
End Sub

Private Sub Timer6_Timer() '����

Label2.Caption = Label2.Caption & "   ��timer6����ģʽ��0.3min"
a(3) = 500000
a(4) = 50000000
Timer1.Interval = 40


End Sub




Private Sub Timer7_Timer()
'���˼ӱ�
tt = 0
For zx = 201 To 400 '�б�
If (��ƣ(zx) > 2) And (Image1(zx).Visible = False) And (tt < 25) Then
tt = 1 + tt
����(zx) = 1
��ƣ(zx) = 0 '��ƣ0������� �ڳ���
di(12) = di(12) + 1
End If
Next
tt = 0
For zx = 201 To 400 '�б�
If (��ƣ(zx) > 2) And (Image1(zx).Visible = False) And (tt < 25) Then
tt = 1 + tt
����(zx) = 3
��ƣ(zx) = 0 '��ƣ0������� �ڳ���
di(6) = di(6) + 1
End If
Next

tt = 0
For zx = 201 To 400 '�б�
If (��ƣ(zx) > 2) And (Image1(zx).Visible = False) And (tt < 25) Then
tt = 1 + tt
����(zx) = 2
��ƣ(zx) = 0 '��ƣ0������� �ڳ���
di(9) = di(9) + 1
End If
Next


Label2.Caption = Label2.Caption & "��1 ��˵��         ��������ǵо�      ����ɱ����˱���" & di(6) & " " & di(9) & " " & di(12)
End Sub

Private Sub Timer8_Timer()

End Sub

Private Sub Timer9_Timer()

End Sub


Private Sub youhao_Timer()






�г���.Enabled = False
����ǰ.Enabled = False
Dim you1 As Long
For you1 = 0 To 200
If Image1(you1).Left > 5200 And Image1(you1).Visible = True Then
Label2.Caption = Label2.Caption & "   �Ѻù�ϵ����"


youhao.Enabled = False
End If
Next

End Sub

Private Sub ��ɱ��_Timer()
Dim xy As Long 'ɱɱɱɱɱɱɱɱɱɱɱɱ��ɱ�У����ϵı�����״̬
Dim ss As Long
Dim ssa As Long
Dim ssd As Long
Dim ssf As Long
Dim ssg As Long
Dim xx As Long
Dim xxd As Long
Dim ppo As Long
Dim nima As Long

For xy = 1 To 199
For ss = 201 To 401
ppo = 0 '����
If (Image1(xy).Visible = True) Then
If (Image1(ss).Visible = True) Then  '˫������

 xx = ����(xy)
 xxd = ����(ss)
 ssa = Image1(ss).Left
 ssd = Image1(ss).Top
 ssf = Image1(xy).Left
 ssg = Image1(xy).Top '˫��Ϊֹ
 
 
 
 If ����(xy) <> 2 Then '�ǹ�




If (ssa <= ssf + 200 And ssa >= ssf - 200) And (xx = 1) And ((ssd <= ssg) And (ssd + 400 >= ssg)) Then   'λ�ã�������Χ������Χ������Χ������Χ
         If xxd = 2 Then
        ��ƣ(ss) = ��ƣ(ss) + 1
        ppo = 1
        Else
        ��ƣ(ss) = ��ƣ(ss) + 3
        ppo = 1
        End If
End If

    If (ssa <= ssf + 200 And ssa >= ssf - 200) And (xx = 2) And ((ssd >= ssg) And (ssd - 400 <= ssg)) Then 'λ��
         If xxd = 2 Then
        ��ƣ(ss) = ��ƣ(ss) + 1
        ppo = 1
        Else
        ��ƣ(ss) = ��ƣ(ss) + 3
        ppo = 1
        End If
End If

   If ((ssa <= ssf) And (ssa + 400 >= ssf)) And (xx = 3) And (ssd <= ssg + 200 And ssg <= ssd) Then 'λ��
         If xxd = 4 Then
        ��ƣ(ss) = ��ƣ(ss) + 1
        ppo = 1
        Else
        ��ƣ(ss) = ��ƣ(ss) + 3
        ppo = 1
        End If
End If

   If ((ssa >= ssf) And (ssa - 400 <= ssf)) And (xx = 4) And (ssd <= ssg + 200 And ssg - 200 <= ssd) Then  'λ��
         If xxd = 3 Then
        ��ƣ(ss) = ��ƣ(ss) + 1
        ppo = 1
        Else
        ��ƣ(ss) = ��ƣ(ss) + 3
        ppo = 1
        End If

End If



Else

If (ssa <= ssf + 200 And ssa >= ssf - 200) And (xx = 1) And (ssg - 1000 >= ssd And ssg - 1500 <= ssd) Then '��
         If xxd = 2 Then
        ��ƣ(ss) = ��ƣ(ss) + 1
        ppo = 1


        Else
        ��ƣ(ss) = ��ƣ(ss) + 3
        ppo = 1
 
        End If
End If
    If (ssa <= ssf + 200 And ssa >= ssf - 200) And (xx = 2) And (ssg + 1000 <= ssd And ssg + 1500 >= ssd) Then 'λ��
         If xxd = 1 Then
        ��ƣ(ss) = ��ƣ(ss) + 1
        ppo = 1

        Else
        ��ƣ(ss) = ��ƣ(ss) + 3
        ppo = 1

        End If

End If
   If (ssf - 1000 > ssa And ssf - 1500 < ssa) And (xx = 3) And (ssd <= ssg + 200 And ssg - 200 <= ssd) Then 'λ��
         If xxd = 4 Then
        ��ƣ(ss) = ��ƣ(ss) + 1
        ppo = 1

        Else
        ��ƣ(ss) = ��ƣ(ss) + 3
        ppo = 1

        End If
End If

   If (ssf + 1000 < ssa And ssf + 1500 > ssa) And (xx = 4) And (ssd <= ssg + 200 And ssg - 200 <= ssd) Then 'λ��
         If xxd = 3 Then
        ��ƣ(ss) = ��ƣ(ss) + 1
        ppo = 1

        Else
        ��ƣ(ss) = ��ƣ(ss) + 3
        ppo = 1

        End If
End If
 
        End If



If ppo = 1 Then
'ɱ������

If ����(xy) = 1 Then '����
If ����(xy) = 1 Then
         Image1(xy).Picture = Image2(4)
    End If
 If ����(xy) = 2 Then
 Image1(xy).Picture = Image2(8)
 End If
 If ����(xy) = 3 Then
 Image1(xy).Picture = Image2(6)
 End If
 If ����(xy) = 4 Then
 Image1(xy).Picture = Image2(2)
 End If
End If
 If ����(xy) = 2 Then '����
If ����(xy) = 1 Then
         Image1(xy).Picture = Image2(12)
    End If
 If ����(xy) = 2 Then
 Image1(xy).Picture = Image2(13)
 End If
 If ����(xy) = 3 Then
 Image1(xy).Picture = Image2(14)
 End If
 If ����(xy) = 4 Then
 Image1(xy).Picture = Image2(9)
 End If
End If
 If ����(xy) = 3 Then '����
If ����(xy) = 1 Then
         Image1(xy).Picture = Image2(20)
    End If
 If ����(xy) = 2 Then
 Image1(xy).Picture = Image2(24)
 End If
 If ����(xy) = 3 Then
 Image1(xy).Picture = Image2(26)
 End If
 If ����(xy) = 4 Then
 Image1(xy).Picture = Image2(22)
 End If '               ɱɱɱɱɱɱɱɱɱɱɱɱ��ɱ�н���
End If
End If
End If
End If
Next
Next


 
End Sub

Private Sub ��ɱ��_Timer()
Dim xy As Long 'ɱɱɱɱɱɱɱɱɱɱɱɱ��ɱ�У����ϵı�����״̬
Dim ss As Long
Dim ssa As Long
Dim ssd As Long
Dim ssf As Long
Dim ssg As Long
Dim xx As Long
Dim xxd As Long
Dim ppo As Long
Dim nima As Long


For xy = 201 To 400
For ss = 0 To 200
ppo = 0
If (Image1(xy).Visible = True) Then
If (Image1(ss).Visible = True) Then '˫������

 xx = ����(xy)
 xxd = ����(ss)
 ssa = Image1(ss).Left
 ssd = Image1(ss).Top
 ssf = Image1(xy).Left
 ssg = Image1(xy).Top '˫��Ϊ
 
 
 
 If ����(xy) <> 2 Then '�ǹ�




If (ssa <= ssf + 200 And ssa >= ssf - 200) And (xx = 1) And ((ssd <= ssg) And (ssd + 400 >= ssg)) Then 'λ�ã�������Χ������Χ������Χ������Χ
         If xxd = 2 Then
        ��ƣ(ss) = ��ƣ(ss) + 1
        ppo = 1
        Else
        ��ƣ(ss) = ��ƣ(ss) + 3
        ppo = 1
        End If
End If

    If (ssa <= ssf + 200 And ssa >= ssf - 200) And (xx = 2) And ((ssd >= ssg) And (ssd - 400 <= ssg)) Then  'λ��
         If xxd = 2 Then
        ��ƣ(ss) = ��ƣ(ss) + 1
        ppo = 1
        Else
        ��ƣ(ss) = ��ƣ(ss) + 3
        ppo = 1
        End If
End If

   If ((ssa <= ssf) And (ssa + 400 >= ssf)) And (xx = 3) And (ssd <= ssg + 200 And ssg <= ssd) Then 'λ��
         If xxd = 4 Then
        ��ƣ(ss) = ��ƣ(ss) + 1
        ppo = 1
        Else
        ��ƣ(ss) = ��ƣ(ss) + 3
        ppo = 1
        End If
End If

   If ((ssa >= ssf) And (ssa - 400 <= ssf)) And (xx = 4) And (ssd <= ssg + 200 And ssg - 200 <= ssd) Then 'λ��
         If xxd = 3 Then
        ��ƣ(ss) = ��ƣ(ss) + 1
        ppo = 1
        Else
        ��ƣ(ss) = ��ƣ(ss) + 3
        ppo = 1
        End If

End If



Else
For nima = 1 To 3
If (ssa <= ssf + 200 And ssa >= ssf - 200) And (xx = 1) And (ssg - 1000 >= ssd And ssg - 1500 <= ssd) Then  '��
         If xxd = 2 Then
        ��ƣ(ss) = ��ƣ(ss) + 1
        ppo = 1
        Else
        ��ƣ(ss) = ��ƣ(ss) + 3
        ppo = 1

        End If
End If
    If (ssa <= ssf + 200 And ssa >= ssf - 200) And (xx = 2) And (ssg + 1000 <= ssd And ssg + 1500 >= ssd) Then 'λ��
         If xxd = 2 Then
        ��ƣ(ss) = ��ƣ(ss) + 1
        ppo = 1

        Else
        ��ƣ(ss) = ��ƣ(ss) + 3
        ppo = 1

        End If

End If
   If (ssf - 1000 > ssa And ssf - 1500 < ssa) And (xx = 3) And (ssd <= ssg + 200 And ssg - 200 <= ssd) Then 'λ��
         If xxd = 4 Then
        ��ƣ(ss) = ��ƣ(ss) + 1
        ppo = 1

        Else
        ��ƣ(ss) = ��ƣ(ss) + 3
        ppo = 1

        End If
End If

   If (ssf + 1000 < ssa And ssf + 1500 > ssa) And (xx = 4) And (ssd <= ssg + 200 And ssg - 200 <= ssd) Then  'λ��
         If xxd = 3 Then
        ��ƣ(ss) = ��ƣ(ss) + 1
        ppo = 1

        Else
        ��ƣ(ss) = ��ƣ(ss) + 3
        ppo = 1

        End If
End If
 Next
        End If



If ppo = 1 Then
'ɱ������

If ����(xy) = 1 Then '����
If ����(xy) = 1 Then
         Image1(xy).Picture = Image2(4)
    End If
 If ����(xy) = 2 Then
 Image1(xy).Picture = Image2(8)
 End If
 If ����(xy) = 3 Then
 Image1(xy).Picture = Image2(6)
 End If
 If ����(xy) = 4 Then
 Image1(xy).Picture = Image2(2)
 End If
End If
 If ����(xy) = 2 Then '����
If ����(xy) = 1 Then
         Image1(xy).Picture = Image2(12)
    End If
 If ����(xy) = 2 Then
 Image1(xy).Picture = Image2(13)
 End If
 If ����(xy) = 3 Then
 Image1(xy).Picture = Image2(14)
 End If
 If ����(xy) = 4 Then
 Image1(xy).Picture = Image2(9)
 End If
End If
 If ����(xy) = 3 Then '����
If ����(xy) = 1 Then
         Image1(xy).Picture = Image2(20)
    End If
 If ����(xy) = 2 Then
 Image1(xy).Picture = Image2(24)
 End If
 If ����(xy) = 3 Then
 Image1(xy).Picture = Image2(26)
 End If
 If ����(xy) = 4 Then
 Image1(xy).Picture = Image2(22)
 End If '               ɱɱɱɱɱɱɱɱɱɱɱɱ��ɱ�н���
End If
End If
End If
End If
Next
Next
End Sub


Private Sub �г���_Timer()

For qqq = 1 To 6 '����λ�ð�����ʼ
qqx(qqq) = 1
qqy(qqq) = qqq
Next
For qqq = 7 To 12
qqx(qqq) = 2
qqy(qqq) = qqq - 6
Next
For qqq = 13 To 18
qqx(qqq) = 3
qqy(qqq) = qqq - 12
Next
For qqq = 19 To 24
qqx(qqq) = 4
qqy(qqq) = qqq - 18
Next
For qqq = 25 To 30
qqx(qqq) = 5
qqy(qqq) = qqq - 24
Next             '����λ�ð�������
'��������6��ʼ
mmc = 0
zx = 10100
For mmd = 201 To 400 'ѡ���
   If (Image1(mmd).Visible = False) And (��ƣ(mmd) < 3) And (����(mmd) = 3) And (mmc < 6) Then '������Ҫ�� ���֣�����
   mmc = mmc + 1 '��������
  ����ǰ.Enabled = True
   '����λ�ÿ�ʼ  δ
mmb = 0
ab = 0
mmn = 0
If c(138, 98) = 1 Then '��һƬ������ô��
zx = 9300
End If
If c(138, 90) = 1 Then
zx = 8500

End If
mmn = qqx(mmc) * 300 '���ȸ��õ�
ab = qqy(mmc) * 400
'����λ�ý���
  
  di(12) = di(12) - 1 '����   ������Ҫ�� ����
  Image1(mmd).Visible = True
  Image1(mmd).Left = 14000 - mmn '�������
  Image1(mmd).Top = zx - ab '�����ϱ�
  ����(mmd) = 3

����ǰ.Enabled = True
'��������
End If '�ǻ����ִ�еĶ�������
  ����ǰ.Enabled = True
   Next

'�����������
'���ϳ����6��ʼ
mmc = 0
zx = 10100
For mmd = 201 To 400 'ѡ���
   If (Image1(mmd).Visible = False) And (��ƣ(mmd) < 3) And (����(mmd) = 3) And (mmc < 6) Then '������Ҫ�� ���֣�����
   mmc = mmc + 1 '��������
  ����ǰ.Enabled = True
   '����λ�ÿ�ʼ  δ
mmb = 0
ab = 0
mmn = 0
If c(138, 98) = 1 Then '��һƬ������ô��
zx = 9300
End If
If c(138, 90) = 1 Then
zx = 8500

End If
mmn = qqx(mmc) * 300 '���ȸ��õ�
ab = qqy(mmc) * 400
'����λ�ý���
  
  di(12) = di(12) - 1 '����   ������Ҫ�� ����
  Image1(mmd).Visible = True
  Image1(mmd).Left = 14000 - mmn '�������
  Image1(mmd).Top = zx - ab '�����ϱ�
  ����(mmd) = 1

����ǰ.Enabled = True
'��������
End If '�ǻ����ִ�еĶ�������
  ����ǰ.Enabled = True
   Next

'�����������


'����ǹ12��ʼ
mmc = 0
zx = 10100
For mmd = 201 To 400 'ѡ���
   If (Image1(mmd).Visible = False) And (��ƣ(mmd) < 3) And (����(mmd) = 1) And (mmc < 12) Then '������Ҫ�� ���֣�����
   mmc = mmc + 1 '��������
 
   '����λ�ÿ�ʼ  δ
mmb = 0
ab = 0
mmn = 0
If c(138, 98) = 1 Then '��һƬ������ô��
zx = 9300
End If
If c(138, 90) = 1 Then
zx = 8500

End If
mmn = qqx(mmc) * 500 '���ȸ��õ�
ab = qqy(mmc) * 400
'����λ�ý���
  
  di(6) = di(6) - 1 '����   ������Ҫ�� ����
  Image1(mmd).Visible = True
  Image1(mmd).Left = 14000 - mmn '�������
  Image1(mmd).Top = zx - ab '�����ϱ�
  

 ����(mmd) = 1

End If '�ǻ����ִ�еĶ�������
   Next

'���ϳ�ǹ12����
End Sub

Private Sub �о�������_Timer()
Dim jkl As Long

For jkl = 201 To 400
If ����(jkl) = 3 And Image1(jkl).Left <= 400 And Image1(jkl).Top >= 1000 Then '�������
����(jkl) = 1
End If
Next

For jkl = 201 To 400
If Image1(jkl).Left <= 400 And Image1(jkl).Top > 400 Then  '��
����(jkl) = 1
End If
Next

For jkl = 201 To 400
If ����(jkl) = 3 And Image1(jkl).Left >= 11000 And Image1(jkl).Top <= 4800 Then '�������
����(jkl) = 3
End If
Next
For jkl = 201 To 400
If ����(jkl) = 3 And Image1(jkl).Left <= 6000 And Image1(jkl).Left >= 5000 And Image1(jkl).Top <= 4800 And Image1(jkl).Top >= 3000 Then '�������
����(jkl) = 1
End If
Next
For jkl = 201 To 400
If ����(jkl) = 3 And Image1(jkl).Left <= 6000 And Image1(jkl).Left >= 5000 And Image1(jkl).Top <= 400 Then  '�������
����(jkl) = 3
End If
Next
For jkl = 201 To 400
If ����(jkl) = 1 And Image1(jkl).Left >= 135 And Image1(jkl).Top <= 400 Then 'ǹ������
����(jkl) = 3
End If
Next

For jkl = 201 To 400
If Image1(jkl).Top <= 400 And Image1(jkl).Left > 400 Then '��
����(jkl) = 3
End If
Next


End Sub



Private Sub ����ǰ_Timer()
Dim iii As Long

For x(6) = 201 To 400
 If Image1(x(5)).Visible = True And ��ƣ(x(6)) < 3 Then
 
 If ����(x(6)) > 2 Then
   If ����(x(6)) = 4 Then
   iii = 1
   Else
   iii = -1
   End If
    If (Image1(x(6)).Visible = True) And (c(Image1(x(6)).Top / 100, (Image1(x(6)).Left + iii * 100) / 100) = 0) Then '��ǰ��û����
    If ����(x(6)) = 3 Then
    Image1(x(6)).Left = Image1(x(6)).Left + iii * 200
    Else
    Image1(x(6)).Left = Image1(x(6)).Left + iii * 100
    End If
   End If
End If
If ����(x(6)) < 3 Then
   If ����(x(6)) = 2 Then
   iii = 1
   Else: iii = -1
   End If
   If (Image1(x(6)).Visible = True) And (c(Image1(x(6)).Left / 100, (Image1(x(6)).Top + iii * 100) / 100) = 0) Then
    If ����(x(6)) = 3 Then
    Image1(x(6)).Top = Image1(x(6)).Top + iii * 200
    Else: Image1(x(6)).Top = Image1(x(6)).Top + iii * 100
    End If
   End If
End If
End If
Next
End Sub



Private Sub ���_Timer()
If Image8.Visible = True Then

Image8.Picture = Image12.Picture

End If
End Sub

Private Sub ���ָ���ƣ_Timer()

 If ��ƣ(401) > 0 Then
 ��ƣ(401) = ��ƣ(401) - 1
 End If
Dim jixie As Long
jixie = 0
Do While ��ƣ(0) > 0 And jixie = 0
��ƣ(0) = ��ƣ(0) - 1
jixie = 1
Loop
End Sub

Private Sub �ؽ���_Timer()
jiangji = 0
End Sub

Private Sub ���Ʊ߽�_Timer()
Dim vvbb As Long
For vvbb = 0 To 431
If Image1(vvbb).Visible = True Then
 If Image1(vvbb).Left <= 0 Then 'way1
 ����(vvbb) = 4
 

 End If
 If Image1(vvbb).Left >= 15000 Or Image1(vvbb).Left = 14900 Then 'way2
 ����(vvbb) = 3
 
End If
 If Image1(vvbb).Top <= -100 Or Image1(vvbb).Top <= 0 Then
 ����(vvbb) = 2
End If
 If Image1(vvbb).Top >= 11000 Or Image1(vvbb).Top = 11100 Then
 ����(vvbb) = 1
 End If
End If
Next

If Image1(0).Visible = True Then '����Ͻ����翪ʼ
 If Image1(0).Left <= 0 Then
 ��ƣ(0) = ��ƣ(0) + 1
Image1(0).Left = 5000
 Image1(0).Top = 5000
 Label2.Caption = Label2.Caption & "   ����磬��������"
  a(1) = a(1) - 1
  
  Label1(7).Caption = a(7)

 

 End If
 If Image1(0).Left >= 15000 Then

Image1(0).Left = 5000 '����Ͻ����翪ʼ

 Image1(0).Top = 5000
 Label2.Caption = Label2.Caption & "  ����磬��������"
 ��ƣ(0) = ��ƣ(0) + 1
 a(1) = a(1) - 1
 
End If
 If Image1(0).Top <= 0 Then
 
Image1(0).Left = 5000 '����Ͻ����翪ʼ
 Image1(0).Top = 5000
 Label2.Caption = Label2.Caption & "   ����磬��������"
 ��ƣ(0) = ��ƣ(0) + 1
 a(1) = a(1) - 1
End If
 If Image1(0).Top >= 11000 Then
 
Image1(0).Left = 5000 '����Ͻ����翪ʼ
 Image1(0).Top = 5000
 Label2.Caption = Label2.Caption & "   ����磬��������"
��ƣ(0) = ��ƣ(0) + 1
 a(1) = a(1) - 1
 End If
End If

End Sub

Private Sub ��ʾ_Timer()
If Image1(201).Visible = True Then '˵��
Image4(0).Visible = True
End If
Dim aq As Long
If aq = 10 Then
Label2.Caption = ""
End If
If a(2) < 3 Then
Label2.Caption = Label2.Caption & "         ����ʧȥ���ģ�ʿ�����ӻ�û�ӣ��㿴�Ű죬�����3000�� ����7���� �ɰ�������"
End If
If (a(3) < 0) Or (a(4) < 0) Then
Label2.Caption = Label2.Caption & "�����Դ�Ѿ��Ǹ�ֵ���뱣��"
aq = aq + 1
End If
If ��ƣ(0) > 120 And Image1(0).Visible = True Then

Label2.Caption = Label2.Caption & "    ��ľ����Ѿ��ǵ�ֵ�������ˣ��뱣��"
aq = aq + 1
End If
End Sub






