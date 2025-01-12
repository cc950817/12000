VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "mscomm32.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "電腦硬體裝修乙級第一站 第1題"
   ClientHeight    =   6615
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12555
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   12555
   StartUpPosition =   2  '螢幕中央
   Begin VB.CommandButton Command2 
      Caption         =   "Connect Bluetooth"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      TabIndex        =   4
      Top             =   2640
      Width           =   3255
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   10680
      Top             =   5400
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   4
      DTREnable       =   -1  'True
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1200
      Top             =   5280
   End
   Begin VB.CommandButton Command1 
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Index           =   3
      Left            =   5160
      Style           =   1  '圖片外觀
      TabIndex        =   2
      Top             =   5400
      Width           =   1995
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Red LED"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Index           =   2
      Left            =   7320
      Style           =   1  '圖片外觀
      TabIndex        =   1
      Top             =   4200
      Width           =   3000
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Green LED"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Index           =   1
      Left            =   1800
      Style           =   1  '圖片外觀
      TabIndex        =   0
      Top             =   4200
      Width           =   3000
   End
   Begin VB.Label Label1 
      Alignment       =   2  '置中對齊
      BackStyle       =   0  '透明
      BorderStyle     =   1  '單線固定
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   18
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      TabIndex        =   3
      Top             =   1680
      Width           =   4935
   End
   Begin VB.Shape G 
      BorderColor     =   &H00000000&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  '實心
      Height          =   285
      Index           =   0
      Left            =   5535
      Shape           =   3  '圓形
      Top             =   3375
      Width           =   270
   End
   Begin VB.Shape G 
      BorderColor     =   &H00000000&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  '實心
      Height          =   285
      Index           =   1
      Left            =   5220
      Shape           =   3  '圓形
      Top             =   3375
      Width           =   270
   End
   Begin VB.Shape G 
      BorderColor     =   &H00000000&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  '實心
      Height          =   285
      Index           =   2
      Left            =   4920
      Shape           =   3  '圓形
      Top             =   3375
      Width           =   270
   End
   Begin VB.Shape G 
      BorderColor     =   &H00000000&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  '實心
      Height          =   285
      Index           =   3
      Left            =   4605
      Shape           =   3  '圓形
      Top             =   3375
      Width           =   270
   End
   Begin VB.Shape G 
      BorderColor     =   &H00000000&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  '實心
      Height          =   285
      Index           =   4
      Left            =   4305
      Shape           =   3  '圓形
      Top             =   3375
      Width           =   270
   End
   Begin VB.Shape G 
      BorderColor     =   &H00000000&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  '實心
      Height          =   285
      Index           =   5
      Left            =   3990
      Shape           =   3  '圓形
      Top             =   3375
      Width           =   270
   End
   Begin VB.Shape G 
      BorderColor     =   &H00000000&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  '實心
      Height          =   285
      Index           =   6
      Left            =   3690
      Shape           =   3  '圓形
      Top             =   3375
      Width           =   270
   End
   Begin VB.Shape G 
      BorderColor     =   &H00000000&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  '實心
      Height          =   285
      Index           =   7
      Left            =   3375
      Shape           =   3  '圓形
      Top             =   3375
      Width           =   270
   End
   Begin VB.Shape R 
      BorderColor     =   &H00000000&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  '實心
      Height          =   285
      Index           =   0
      Left            =   7995
      Shape           =   3  '圓形
      Top             =   3375
      Width           =   270
   End
   Begin VB.Shape R 
      BorderColor     =   &H00000000&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  '實心
      Height          =   285
      Index           =   1
      Left            =   7680
      Shape           =   3  '圓形
      Top             =   3375
      Width           =   270
   End
   Begin VB.Shape R 
      BorderColor     =   &H00000000&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  '實心
      Height          =   285
      Index           =   2
      Left            =   7380
      Shape           =   3  '圓形
      Top             =   3375
      Width           =   270
   End
   Begin VB.Shape R 
      BorderColor     =   &H00000000&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  '實心
      Height          =   285
      Index           =   3
      Left            =   7065
      Shape           =   3  '圓形
      Top             =   3375
      Width           =   270
   End
   Begin VB.Shape R 
      BorderColor     =   &H00000000&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  '實心
      Height          =   285
      Index           =   4
      Left            =   6765
      Shape           =   3  '圓形
      Top             =   3375
      Width           =   270
   End
   Begin VB.Shape R 
      BorderColor     =   &H00000000&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  '實心
      Height          =   285
      Index           =   5
      Left            =   6450
      Shape           =   3  '圓形
      Top             =   3375
      Width           =   270
   End
   Begin VB.Shape R 
      BorderColor     =   &H00000000&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  '實心
      Height          =   285
      Index           =   6
      Left            =   6150
      Shape           =   3  '圓形
      Top             =   3375
      Width           =   270
   End
   Begin VB.Shape R 
      BorderColor     =   &H00000000&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  '實心
      Height          =   285
      Index           =   7
      Left            =   5835
      Shape           =   3  '圓形
      Top             =   3375
      Width           =   270
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a, b(99), c As Integer
Dim BluetoothConnected As Boolean

Private Sub Command1_Click(Index As Integer)
    a = Index
    c = 0
End Sub

Private Sub display(no)
    For i = 0 To 7
        If no Mod 2 = 1 And a = 1 Then G(i).FillColor = vbGreen
        If no Mod 2 = 1 And a = 2 Then R(i).FillColor = vbRed
        no = no \ 2
    Next i
End Sub

Private Sub Command2_Click()
    If BluetoothConnected Then
        Debug.Print "Output: R0"
        Debug.Print "Output: G0"
        BluetoothConnected = False
        Command2.Caption = "Connect Bluetooth"
    Else
        BluetoothConnected = True
        Command2.Caption = "Disconnect Bluetooth"
        Debug.Print "Output: R0"
        Debug.Print "Output: G0"
    End If
End Sub

Private Sub Form_Load()
    BluetoothConnected = False
End Sub

Private Sub Timer1_Timer()
    b(0) = 1
    b(1) = 2
    b(2) = 4
    b(3) = 8
    b(4) = &H10
    b(5) = &H20
    b(6) = &H40
    b(7) = &H80
    Label1.Caption = "Current Time:" & Time$
    For i = 0 To 7
        G(i).FillColor = vbWhite
        R(i).FillColor = vbWhite
    Next i
    If BluetoothConnected Then
        For i = 0 To 7
            G(i).FillColor = RGB(0, 128, 0)
            R(i).FillColor = RGB(128, 0, 0)
        Next i
        If a = 1 Then
            Debug.Print "Output: G" & b(c)
            display (b(c))
        End If
        If a = 2 And c <= 8 Then
            Debug.Print "Output: R" & 2 ^ c
            display (2 ^ c)
        End If
    End If
    If a = 3 Then End
    If c > 15 Then c = 15 Else c = c + 1
End Sub

