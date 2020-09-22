VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "BitBlt Power"
   ClientHeight    =   5280
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9960
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   9960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton ClosewCB 
      Caption         =   "Close Routine"
      Height          =   435
      Left            =   525
      TabIndex        =   11
      Top             =   4560
      Width           =   1155
   End
   Begin VB.CommandButton WipeCB 
      Caption         =   "Wipe"
      Height          =   435
      Left            =   1805
      TabIndex        =   10
      Top             =   4560
      Width           =   1155
   End
   Begin VB.CommandButton CircleCB 
      Caption         =   "Circle"
      Height          =   435
      Left            =   5645
      TabIndex        =   9
      Top             =   4560
      Width           =   1155
   End
   Begin VB.CommandButton ImplodeCB 
      Caption         =   "Implode"
      Height          =   435
      Left            =   8205
      TabIndex        =   8
      Top             =   4560
      Width           =   1155
   End
   Begin VB.CommandButton TCB 
      Caption         =   "Tenda"
      Height          =   435
      Left            =   6925
      TabIndex        =   7
      Top             =   4560
      Width           =   1155
   End
   Begin VB.CommandButton HourDblCB 
      Caption         =   "Hour (Double)"
      Height          =   435
      Left            =   3085
      TabIndex        =   6
      Top             =   4560
      Width           =   1155
   End
   Begin VB.CommandButton HourCB 
      Caption         =   "Hour (inverse)"
      Height          =   435
      Left            =   4365
      TabIndex        =   5
      Top             =   4560
      Width           =   1155
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   4680
      TabIndex        =   4
      Text            =   "20"
      Top             =   4020
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2520
      TabIndex        =   3
      Text            =   "100000"
      Top             =   4020
      Width           =   1095
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Refresh"
      Height          =   435
      Left            =   540
      TabIndex        =   2
      Top             =   3900
      Width           =   1155
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      Height          =   3285
      Left            =   5040
      Picture         =   "BitBlt.frx":0000
      ScaleHeight     =   215
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   312
      TabIndex        =   1
      Top             =   180
      Width           =   4740
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   3285
      Left            =   180
      Picture         =   "BitBlt.frx":3125A
      ScaleHeight     =   215
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   315
      TabIndex        =   0
      Top             =   180
      Width           =   4785
   End
   Begin VB.Label BlockLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Block Size"
      Height          =   195
      Left            =   3780
      TabIndex        =   13
      Top             =   4080
      Width           =   750
   End
   Begin VB.Label DelayLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Delay"
      Height          =   195
      Left            =   1920
      TabIndex        =   12
      Top             =   4080
      Width           =   405
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
















Private Sub CircleCB_Click()
Const PI = 3.1415
Dim ray, angle As Double

dwRop& = &HCC0020
hDestDC& = Picture2.hDC
hSrcDC& = Picture1.hDC
DoEvents
ray = Sqr(Picture1.ScaleHeight ^ 2 + Picture1.ScaleWidth ^ 2) / 2
For i = ray To 0 Step -1.9
    For angle = 0 To 2 * PI Step 0.01
        X = i * Cos(angle) + (Picture1.ScaleWidth / 2)
        Y = i * Sin(angle) + (Picture1.ScaleHeight / 2)
        suc& = BitBlt(hDestDC&, X, Y, 3, 3, hSrcDC&, _
      X, Y, dwRop&)
    Next
Next
End Sub

Private Sub ClosewCB_Click()
Dim ImgX, ImgY As Integer
Dim NumLoop As Integer
Dim HalfHeight As Integer
Dim i As Integer
Dim suc&
Dim dwRop&

hDestDC& = Picture2.hDC
hSrcDC& = Picture1.hDC
ImgX = Picture1.ScaleWidth
ImgY = Picture1.ScaleHeight
HalfHeight = ImgY / 2
dwRop& = &HCC0020
wait = 55000
For i = 0 To HalfHeight
    Y = i
    For X = i To ImgX - i
        suc& = BitBlt(hDestDC&, X, Y, 1, 1, hSrcDC&, _
      X, Y, dwRop&)
    Next
    For a = 1 To wait: Next
    X = ImgX - i
    For Y = i To ImgY - i
        suc& = BitBlt(hDestDC&, X, Y, 1, 1, hSrcDC&, _
      X, Y, dwRop&)
    Next
    For a = 1 To wait: Next
    Y = ImgY - i
    For X = ImgX - i To i Step -1
        suc& = BitBlt(hDestDC&, X, Y, 1, 1, hSrcDC&, _
      X, Y, dwRop&)
    Next
    For a = 1 To wait: Next
    X = i
    For Y = ImgY - i To i Step -1
        suc& = BitBlt(hDestDC&, X, Y, 1, 1, hSrcDC&, _
      X, Y, dwRop&)
    Next
    For a = 1 To wait: Next
    DoEvents
Next
End Sub



Private Sub Command10_Click()

End Sub










Private Sub Command7_Click()
   
    Picture2.Cls
End Sub





Private Sub HourCB_Click()
Const PI = 3.1415
Dim ray, angle As Double
dwRop& = &HCC0020
hDestDC& = Picture2.hDC
hSrcDC& = Picture1.hDC
DoEvents


    For angle = 2 * PI To 0 Step -0.01
        a = Tan(angle)
        b = Cos(angle)
        c = Sin(angle)
        If Abs(a * (Picture1.ScaleWidth / 2)) < (Picture1.ScaleHeight / 2) Then
            For X = 0.5 * (Sgn(b) - 1) * (Picture1.ScaleWidth / 2) To 0.5 * (1 + Sgn(b)) * (Picture1.ScaleWidth / 2)
            suc& = BitBlt(hDestDC&, (Picture1.ScaleWidth / 2) + X, (Picture1.ScaleHeight / 2) + a * X, 5, 5, hSrcDC&, _
      (Picture1.ScaleWidth / 2) + X, (Picture1.ScaleHeight / 2) + a * X, dwRop&)
        Next
        Else
            For Y = 0.5 * (Sgn(c) - 1) * (Picture1.ScaleWidth / 2) To 0.5 * (1 + Sgn(c)) * (Picture1.ScaleWidth / 2)
            suc& = BitBlt(hDestDC&, (Picture1.ScaleWidth / 2) + Y / a, (Picture1.ScaleHeight / 2) + Y, 5, 5, hSrcDC&, _
      (Picture1.ScaleWidth / 2) + Y / a, (Picture1.ScaleHeight / 2) + Y, dwRop&)
            Next
        End If
    Next
End Sub

Private Sub HourDblCB_Click()
Const PI = 3.1415
Dim ray, angle As Double
dwRop& = &HCC0020
hDestDC& = Picture2.hDC
hSrcDC& = Picture1.hDC
DoEvents


    For angle = 0 To 2 * PI Step 0.01
        a = Tan(angle)
        b = Cos(angle)
        c = Sin(angle)
        If Abs(a * (Picture1.ScaleWidth / 2)) < (Picture1.ScaleHeight / 2) Then
            For X = -0.5 * (1 + Sgn(b)) * (Picture1.ScaleWidth / 2) To 0.5 * (1 + Sgn(b)) * (Picture1.ScaleWidth / 2) Step Sgn(b)
            suc& = BitBlt(hDestDC&, (Picture1.ScaleWidth / 2) + X, (Picture1.ScaleHeight / 2) + a * X, 5, 5, hSrcDC&, _
      (Picture1.ScaleWidth / 2) + X, (Picture1.ScaleHeight / 2) + a * X, dwRop&)
        Next
        Else
            For Y = -0.5 * (1 + Sgn(c)) * (Picture1.ScaleWidth / 2) To 0.5 * (1 + Sgn(c)) * (Picture1.ScaleWidth / 2) Step Sgn(c)
            suc& = BitBlt(hDestDC&, (Picture1.ScaleWidth / 2) + Y / a, (Picture1.ScaleHeight / 2) + Y, 5, 5, hSrcDC&, _
      (Picture1.ScaleWidth / 2) + Y / a, (Picture1.ScaleHeight / 2) + Y, dwRop&)
            Next
        End If
    Next
End Sub

Private Sub ImplodeCB_Click()
Const PI = 3.1415
Dim ray, angle As Double
dwRop& = &HCC0020
hDestDC& = Picture2.hDC
hSrcDC& = Picture1.hDC
DoEvents
ray = Sqr(Picture1.ScaleHeight ^ 2 + Picture1.ScaleWidth ^ 2) / 2
For i = ray To 0 Step -1.1
    For angle = 0 To 5 * PI Step 0.01
        X = i * Tan(angle) + (Picture1.ScaleWidth / 2)
        Y = i * Cos(angle) + (Picture1.ScaleHeight / 2)
        suc& = BitBlt(hDestDC&, X, Y, 16, 16, hSrcDC&, _
      X, Y, dwRop&)
    Next
Next
End Sub

Private Sub TCB_Click()
Dim ImgX, ImgY As Integer
Dim NumLoop As Integer
Dim HalfHeight As Integer
Dim i As Integer
Dim suc&
Dim dwRop&

hDestDC& = Picture2.hDC
hSrcDC& = Picture1.hDC

dwRop& = &HCC0020
For i = 0 To Picture1.ScaleHeight / 2
For a = 1 To CLng(Text1.Text): Next
        suc& = BitBlt(hDestDC&, 0, i, Picture1.ScaleWidth, 1, hSrcDC&, 0, i, dwRop&)
        suc& = BitBlt(hDestDC&, 0, Picture1.ScaleHeight - i, Picture1.ScaleWidth, 1, _
        hSrcDC&, 0, Picture1.ScaleHeight - i, dwRop&)
Next
End Sub

Private Sub WipeCB_Click()
Dim ImgX, ImgY As Integer
Dim NumLoop As Integer
Dim HalfHeight As Integer
Dim i As Integer
Dim suc&
Dim dwRop&
Dim blocco&
Dim blocco1&
Dim blocco2&
Dim tempor&
Dim Xtemp&

hDestDC& = Picture2.hDC
hSrcDC& = Picture1.hDC
ImgX = Picture1.ScaleWidth
ImgY = Picture1.ScaleHeight
HalfHeight = ImgY / 2
dwRop& = &HCC0020
wait = CLng(Text1.Text)
blocco = CLng(Text2.Text)
blocco1 = blocco
blocco2 = blocco
For i = 0 To HalfHeight Step blocco1
    Y = i
    For X = i To ImgX - i Step blocco1
        If X + blocco1 > ImgX Then
            blocco1 = ImgX - X
        End If
        suc& = BitBlt(hDestDC&, X, Y, blocco1, blocco2, hSrcDC&, _
      X, Y, dwRop&)
      tempor = X
    Next
    For a = 1 To wait: Next
    X = tempor
    For Y = i + blocco1 To ImgY - i Step blocco2
        If Y + blocco2 > ImgY Then
            blocco2 = ImgY - Y
        End If
        suc& = BitBlt(hDestDC&, X, Y, blocco1, blocco2, hSrcDC&, _
      X, Y, dwRop&)
      tempor = Y
    Next
    For a = 1 To wait: Next
    Y = tempor
    tempor = X
    For X = tempor - blocco To i Step -blocco
        suc& = BitBlt(hDestDC&, X, Y, blocco, blocco2, hSrcDC&, _
      X, Y, dwRop&)
        Xtemp = X
    Next
    For a = 1 To wait: Next
    X = Xtemp
    tempor = Y
    For Y = tempor - blocco To i - blocco Step -blocco
        suc& = BitBlt(hDestDC&, X, Y, blocco, blocco, hSrcDC&, _
      X, Y, dwRop&)
    Next
    For a = 1 To wait: Next
    DoEvents
    blocco1 = blocco
    blocco2 = blocco
Next
End Sub
