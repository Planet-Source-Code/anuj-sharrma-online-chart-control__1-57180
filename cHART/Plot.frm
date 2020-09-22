VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "OnLine Charting"
   ClientHeight    =   5520
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8175
   Icon            =   "Plot.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5520
   ScaleWidth      =   8175
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   4680
      TabIndex        =   6
      Text            =   "Anuj"
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Resume"
      Height          =   375
      Left            =   3360
      TabIndex        =   5
      Top             =   0
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Pause"
      Height          =   375
      Left            =   2040
      TabIndex        =   4
      Top             =   0
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   960
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   0
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   0
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   6720
      TabIndex        =   1
      Top             =   0
      Width           =   1335
   End
   Begin VB.PictureBox Pic 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E8F9FB&
      FillStyle       =   0  'Solid
      Height          =   4455
      Left            =   120
      ScaleHeight     =   293
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   525
      TabIndex        =   0
      Top             =   480
      Width           =   7935
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   120
      Top             =   4560
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Created By: Anuj sharma"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   270
      Left            =   255
      TabIndex        =   7
      Top             =   5220
      Width           =   2370
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------------------------------------------------
'                                       On-line Chart
' Developed by: Anuj sharma
' e-Mail: anujsharrma@yahoo.com
'--------------------------------------------------------------------------------------------------------------------------
Dim dat(5000) As Single
Dim knt As Integer
Dim max As Single
Dim min As Single
Dim Scrw As Integer
Dim Scrh As Integer
Dim scalx As Single
Dim scaly As Single
Dim X As Single
Dim Y As Single
Dim Pause As Boolean

Private Sub Command1_Click()
    End
End Sub

Private Sub Command2_Click()
    Pause = True
End Sub

Private Sub Command3_Click()
    Pause = False
    Timer1.Enabled = True
End Sub

Private Sub Form_Load()
    dat(0) = -999
    Me.Show
    knt = 0
End Sub

Private Sub DrawAxis()
    Dim k As Integer
    
    If dat(0) = -999 Then Exit Sub
    
    Pic.Cls
    Scrw = Pic.ScaleWidth
    Scrh = Pic.ScaleHeight
    
    scalx = Scrw * 0.9 / (knt + 1)
    scaly = Scrh * 0.9 / (max - min)
    
    Text2.Text = ((Int(max / 10) + 1) * 10)
    Pic.Line (Scrw * 0.1, 0)-(Scrw * 0.1, Scrh), RGB(127, 0, 0)
    Pic.Line (0, Scrh * 0.9)-(Scrw, Scrh * 0.9), RGB(127, 0, 0)
    Pic.Line (Scrw * 0.1, Scrh * 0.45)-(Scrw, Scrh * 0.45), RGB(0, 127, 0)
    Pic.PSet (0, Scrh * 0.45)
    Pic.Print ((max + min) / 2)
    
    X = (Scrw * 0.1) + scalx
    Y = (Scrh * 0.9) - (scaly * (dat(0) - min))
        Pic.Line (X - scalx, Scrh * 0.9)-(X, Y), RGB(0, 127, 0), BF
    For k = 1 To knt - 1
        X = X + scalx
        Y = (Scrh * 0.9) - (scaly * (dat(k) - min))
            Pic.Line (X - scalx, Scrh * 0.9)-(X, Y), RGB(0, 127, 0), BF
    Next
End Sub

Private Sub Form_Resize()
    Pic.Width = Me.Width * 0.95
    Pic.Height = Me.Height * 0.8
    Call DrawAxis
End Sub

Private Sub Pic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim idx As Integer
    
    If Timer1.Enabled = False Then
        idx = ((X - (Scrw * 0.1) - (scalx / 2)) / scalx)
        If idx >= 0 And idx < knt Then
            Text3.Text = idx & ", " & dat(idx)
        End If
    End If
End Sub

Private Sub Timer1_Timer()
    If knt > 0 Then
        dat(knt) = dat(knt - 1) + IIf(Int(Rnd() * 200) Mod 2 = 0, -Rnd() * 8, Rnd() * 10)
        If dat(knt) < 0 Then dat(knt) = 10
    Else
        dat(knt) = 66.6
    End If
    If knt = 0 Then
        max = dat(knt)
        max = ((Int(max / 10) + 1) * 10)
        min = dat(knt)
        min = ((Int(min / 10) - 1) * 10)
    Else
        If max < dat(knt) Then
            max = dat(knt)
            max = ((Int(max / 10) + 1) * 10)
        End If
        If min > dat(knt) Then
            min = dat(knt)
            min = ((Int(min / 10) - 1) * 10)
        End If
    End If
    Text1.Text = dat(knt)
    knt = knt + 1
    Call DrawAxis
    If Pause = True Then Timer1.Enabled = False
End Sub
