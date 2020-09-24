VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Draw Gradients using API"
   ClientHeight    =   7380
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11790
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   492
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   786
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command9 
      Caption         =   "2 color Gradient"
      Height          =   435
      Left            =   10005
      TabIndex        =   9
      Top             =   5460
      Width           =   1500
   End
   Begin VB.PictureBox picCenter 
      BackColor       =   &H0000FFFF&
      Height          =   375
      Left            =   4875
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   6900
      Width           =   375
   End
   Begin VB.CommandButton Command8 
      Caption         =   "5 color Gradient"
      Height          =   435
      Left            =   10005
      TabIndex        =   0
      Top             =   480
      Width           =   1500
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   10035
      Style           =   2  'Dropdown List
      TabIndex        =   8
      ToolTipText     =   "Direction"
      Top             =   5130
      Width           =   1455
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   10020
      Style           =   2  'Dropdown List
      TabIndex        =   6
      ToolTipText     =   "Direction to fill the colors"
      Top             =   4050
      Width           =   1470
   End
   Begin VB.CommandButton Command7 
      Caption         =   "3 color Gradient"
      Height          =   435
      Left            =   10005
      TabIndex        =   7
      Top             =   4665
      Width           =   1500
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Clear"
      Height          =   435
      Left            =   10005
      TabIndex        =   10
      Top             =   6300
      Width           =   1500
   End
   Begin VB.CommandButton Command5 
      Caption         =   "All"
      Height          =   435
      Left            =   10005
      TabIndex        =   5
      Top             =   3585
      Width           =   1500
   End
   Begin VB.CommandButton Command4 
      Caption         =   "TopRight"
      Height          =   435
      Left            =   10005
      TabIndex        =   2
      Top             =   1785
      Width           =   1500
   End
   Begin VB.CommandButton Command3 
      Caption         =   "TopLeft"
      Height          =   435
      Left            =   10005
      TabIndex        =   1
      Top             =   1275
      Width           =   1500
   End
   Begin VB.CommandButton Command2 
      Caption         =   "BottomLeft"
      Height          =   435
      Left            =   10005
      TabIndex        =   3
      Top             =   2280
      Width           =   1500
   End
   Begin MSComDlg.CommonDialog cDLG 
      Left            =   45
      Top             =   5970
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picTopLeft 
      BackColor       =   &H000000FF&
      Height          =   375
      Left            =   45
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   60
      Width           =   375
   End
   Begin VB.PictureBox picTopRight 
      BackColor       =   &H0000FF00&
      Height          =   375
      Left            =   9540
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   45
      Width           =   375
   End
   Begin VB.PictureBox picBottomRight 
      BackColor       =   &H00000000&
      Height          =   375
      Left            =   9525
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   6495
      Width           =   375
   End
   Begin VB.PictureBox picBottomLeft 
      BackColor       =   &H00FF0000&
      Height          =   375
      Left            =   45
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   6495
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "BottomRight"
      Height          =   435
      Left            =   10005
      TabIndex        =   4
      Top             =   2790
      Width           =   1500
   End
   Begin VB.PictureBox picMain 
      AutoRedraw      =   -1  'True
      Height          =   6795
      Left            =   465
      ScaleHeight     =   449
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   596
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   45
      Width           =   9000
   End
   Begin VB.Frame Frame1 
      Caption         =   "4 color gradient"
      Height          =   3510
      Left            =   9870
      TabIndex        =   17
      Top             =   990
      Width           =   1785
   End
   Begin VB.Label Label5 
      Caption         =   "5 (center-middle)"
      Height          =   270
      Left            =   5370
      TabIndex        =   23
      Top             =   7005
      Width           =   1305
   End
   Begin VB.Label Label4 
      Caption         =   "4"
      Height          =   225
      Left            =   135
      TabIndex        =   22
      Top             =   6885
      Width           =   255
   End
   Begin VB.Label Label3 
      Caption         =   "3"
      Height          =   270
      Left            =   9615
      TabIndex        =   21
      Top             =   6885
      Width           =   270
   End
   Begin VB.Label Label2 
      Caption         =   "2"
      Height          =   255
      Left            =   9645
      TabIndex        =   20
      Top             =   435
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "1"
      Height          =   270
      Left            =   150
      TabIndex        =   19
      Top             =   480
      Width           =   270
   End
   Begin VB.Label lblTime 
      AutoSize        =   -1  'True
      Height          =   240
      Left            =   495
      TabIndex        =   16
      Top             =   6930
      Width           =   2310
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Sub Command1_Click()
    Dim aTick As Long
    aTick = GetTickCount
    DoGradient45Colors picMain, picTopLeft.BackColor, picTopRight.BackColor, picBottomLeft.BackColor, _
            picBottomRight.BackColor, , , , , , BottomRight
    picMain.Refresh
    lblTime = "Time: " & ((GetTickCount - aTick) / 1000) & " seconds."
End Sub

Private Sub Command2_Click()
    Dim aTick As Long
    aTick = GetTickCount
    DoGradient45Colors picMain, picTopLeft.BackColor, picTopRight.BackColor, picBottomLeft.BackColor, _
        picBottomRight.BackColor, , , , , , BottomLeft
    picMain.Refresh
    lblTime = "Time: " & ((GetTickCount - aTick) / 1000) & " seconds."
End Sub

Private Sub Command3_Click()
    Dim aTick As Long
    aTick = GetTickCount
    DoGradient45Colors picMain, picTopLeft.BackColor, picTopRight.BackColor, picBottomLeft.BackColor, _
        picBottomRight.BackColor, , , , , , TopLeft
    picMain.Refresh
    lblTime = "Time: " & ((GetTickCount - aTick) / 1000) & " seconds."
End Sub

Private Sub Command4_Click()
    Dim aTick As Long
    aTick = GetTickCount
    DoGradient45Colors picMain, picTopLeft.BackColor, picTopRight.BackColor, picBottomLeft.BackColor, _
            picBottomRight.BackColor, , , , , , TopRight
    picMain.Refresh
    lblTime = "Time: " & ((GetTickCount - aTick) / 1000) & " seconds."
End Sub

Private Sub Command5_Click()
    Dim aTick As Long
    aTick = GetTickCount
    DoGradient45Colors picMain, picTopLeft.BackColor, picTopRight.BackColor, picBottomLeft.BackColor, _
            picBottomRight.BackColor, , , , , , , Combo1.ListIndex
    picMain.Refresh
    lblTime = "Time: " & ((GetTickCount - aTick) / 1000) & " seconds."
End Sub

Private Sub Command6_Click()
    picMain.Cls
End Sub

Private Sub Command7_Click()
    Dim aTick As Long
    aTick = GetTickCount
    If Combo2.ListIndex = 0 Then
        DoGradient3Colors picMain, picTopLeft.BackColor, picTopRight.BackColor, _
            picBottomRight.BackColor, FillHor, , , , 50
    Else
        DoGradient3Colors picMain, picTopLeft.BackColor, picTopRight.BackColor, _
        picBottomRight.BackColor, FillVer, , , 50
    End If
    picMain.Refresh
    lblTime = "Time: " & ((GetTickCount - aTick) / 1000) & " seconds."
End Sub

Private Sub Command8_Click()
    Dim aTick As Long
    aTick = GetTickCount
    DoGradient45Colors picMain, picTopLeft.BackColor, picTopRight.BackColor, picBottomLeft.BackColor, _
            picBottomRight.BackColor, picCenter.BackColor, , , , , , Combo1.ListIndex
    picMain.Refresh
    lblTime = "Time: " & ((GetTickCount - aTick) / 1000) & " seconds."
End Sub

Private Sub Command9_Click()
    Dim aTick As Long
    aTick = GetTickCount
    If Combo2.ListIndex = 0 Then
        DoGradient picMain, picTopLeft.BackColor, picTopRight.BackColor, FillHor, , , , 100
    Else
        DoGradient picMain, picTopLeft.BackColor, picTopRight.BackColor, FillVer, , , 100
    End If
    picMain.Refresh
    lblTime = "Time: " & ((GetTickCount - aTick) / 1000) & " seconds."
End Sub

Private Sub Form_Load()
    ScaleMode = vbPixels
           
    Combo1.AddItem "Slash /" ' from top-right to bottom-left
    Combo1.AddItem "BackSlash \" ' from top-left to bottom-right
    Combo1.ListIndex = 0
    
    Combo2.AddItem "Horizontal"
    Combo2.AddItem "Vertical"
    Combo2.ListIndex = 0
    
    Form1.Show
End Sub

Private Sub picBottomLeft_Click()
    cDLG.Color = picBottomLeft.BackColor
    cDLG.ShowColor
    picBottomLeft.BackColor = cDLG.Color
End Sub

Private Sub picBottomRight_Click()
    cDLG.Color = picBottomRight.BackColor
    cDLG.ShowColor
    picBottomRight.BackColor = cDLG.Color
End Sub

Private Sub picCenter_Click()
    cDLG.Color = picTopRight.BackColor
    cDLG.ShowColor
    picCenter.BackColor = cDLG.Color
    Command8_Click
End Sub

Private Sub picTopLeft_Click()
    cDLG.Color = picTopLeft.BackColor
    cDLG.ShowColor
    picTopLeft.BackColor = cDLG.Color
End Sub

Private Sub picTopRight_Click()
    cDLG.Color = picTopRight.BackColor
    cDLG.ShowColor
    picTopRight.BackColor = cDLG.Color
End Sub
