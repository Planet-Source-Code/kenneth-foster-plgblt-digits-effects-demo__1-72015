VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "PlgBlt---digit effects"
   ClientHeight    =   5025
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3780
   LinkTopic       =   "Form1"
   ScaleHeight     =   5025
   ScaleWidth      =   3780
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox p3 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   750
      Left            =   195
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   35
      TabIndex        =   38
      Top             =   360
      Width           =   555
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   315
      Left            =   360
      TabIndex        =   37
      Top             =   3840
      Width           =   900
   End
   Begin VB.CommandButton cmdColor 
      Caption         =   "Back Color"
      Height          =   420
      Left            =   90
      TabIndex        =   34
      Top             =   4515
      Width           =   1155
   End
   Begin VB.Frame Frame2 
      Caption         =   "Effects"
      Height          =   2430
      Left            =   1620
      TabIndex        =   26
      Top             =   2085
      Width           =   1650
      Begin VB.OptionButton optDir 
         Caption         =   "Effect 4"
         Height          =   195
         Index           =   7
         Left            =   90
         TabIndex        =   36
         Top             =   2130
         Width           =   945
      End
      Begin VB.OptionButton optDir 
         Caption         =   "Effect 3"
         Height          =   240
         Index           =   6
         Left            =   90
         TabIndex        =   33
         Top             =   1875
         Width           =   945
      End
      Begin VB.OptionButton optDir 
         Caption         =   "Effect 2"
         Height          =   315
         Index           =   5
         Left            =   90
         TabIndex        =   32
         Top             =   1575
         Width           =   1350
      End
      Begin VB.OptionButton optDir 
         Caption         =   "Left to Right"
         Height          =   255
         Index           =   0
         Left            =   90
         TabIndex        =   31
         Top             =   255
         Value           =   -1  'True
         Width           =   1170
      End
      Begin VB.OptionButton optDir 
         Caption         =   "Right to Left"
         Height          =   255
         Index           =   1
         Left            =   90
         TabIndex        =   30
         Top             =   510
         Width           =   1260
      End
      Begin VB.OptionButton optDir 
         Caption         =   "Top to Bottom"
         Height          =   255
         Index           =   2
         Left            =   90
         TabIndex        =   29
         Top             =   795
         Width           =   1395
      End
      Begin VB.OptionButton optDir 
         Caption         =   "Bottom to Top"
         Height          =   255
         Index           =   3
         Left            =   90
         TabIndex        =   28
         Top             =   1080
         Width           =   1365
      End
      Begin VB.OptionButton optDir 
         Caption         =   "Effect 1"
         Height          =   240
         Index           =   4
         Left            =   90
         TabIndex        =   27
         Top             =   1365
         Width           =   1110
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Keypad"
      Height          =   1680
      Left            =   75
      TabIndex        =   15
      Top             =   2055
      Width           =   1440
      Begin VB.CommandButton cmdDig 
         Caption         =   "2"
         Height          =   285
         Index           =   2
         Left            =   540
         TabIndex        =   25
         Top             =   255
         Width           =   270
      End
      Begin VB.CommandButton cmdDig 
         Caption         =   "3"
         Height          =   285
         Index           =   3
         Left            =   840
         TabIndex        =   24
         Top             =   255
         Width           =   270
      End
      Begin VB.CommandButton cmdDig 
         Caption         =   "4"
         Height          =   285
         Index           =   4
         Left            =   240
         TabIndex        =   23
         Top             =   585
         Width           =   270
      End
      Begin VB.CommandButton cmdDig 
         Caption         =   "1"
         Height          =   285
         Index           =   1
         Left            =   240
         TabIndex        =   22
         Top             =   255
         Width           =   270
      End
      Begin VB.CommandButton cmdDig 
         Caption         =   "5"
         Height          =   285
         Index           =   5
         Left            =   540
         TabIndex        =   21
         Top             =   585
         Width           =   270
      End
      Begin VB.CommandButton cmdDig 
         Caption         =   "6"
         Height          =   285
         Index           =   6
         Left            =   840
         TabIndex        =   20
         Top             =   585
         Width           =   270
      End
      Begin VB.CommandButton cmdDig 
         Caption         =   "7"
         Height          =   285
         Index           =   7
         Left            =   240
         TabIndex        =   19
         Top             =   915
         Width           =   270
      End
      Begin VB.CommandButton cmdDig 
         Caption         =   "8"
         Height          =   285
         Index           =   8
         Left            =   540
         TabIndex        =   18
         Top             =   915
         Width           =   270
      End
      Begin VB.CommandButton cmdDig 
         Caption         =   "9"
         Height          =   285
         Index           =   9
         Left            =   840
         TabIndex        =   17
         Top             =   915
         Width           =   270
      End
      Begin VB.CommandButton cmdDig 
         Caption         =   "0"
         Height          =   285
         Index           =   0
         Left            =   240
         TabIndex        =   16
         Top             =   1260
         Width           =   870
      End
   End
   Begin VB.PictureBox pDig 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   480
      Index           =   9
      Left            =   6435
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   420
      ScaleWidth      =   270
      TabIndex        =   12
      Top             =   2040
      Width           =   330
   End
   Begin VB.PictureBox pDig 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   480
      Index           =   8
      Left            =   6420
      Picture         =   "Form1.frx":0104
      ScaleHeight     =   420
      ScaleWidth      =   270
      TabIndex        =   11
      Top             =   1560
      Width           =   330
   End
   Begin VB.PictureBox pDig 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   480
      Index           =   7
      Left            =   6420
      Picture         =   "Form1.frx":0181
      ScaleHeight     =   420
      ScaleWidth      =   270
      TabIndex        =   10
      Top             =   1080
      Width           =   330
   End
   Begin VB.PictureBox pDig 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   480
      Index           =   6
      Left            =   6420
      Picture         =   "Form1.frx":0230
      ScaleHeight     =   420
      ScaleWidth      =   270
      TabIndex        =   9
      Top             =   585
      Width           =   330
   End
   Begin VB.PictureBox pDig 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   480
      Index           =   5
      Left            =   6405
      Picture         =   "Form1.frx":0337
      ScaleHeight     =   420
      ScaleWidth      =   270
      TabIndex        =   8
      Top             =   90
      Width           =   330
   End
   Begin VB.PictureBox pDig 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   480
      Index           =   4
      Left            =   6030
      Picture         =   "Form1.frx":0439
      ScaleHeight     =   420
      ScaleWidth      =   270
      TabIndex        =   7
      Top             =   2025
      Width           =   330
   End
   Begin VB.PictureBox pDig 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   480
      Index           =   3
      Left            =   6030
      Picture         =   "Form1.frx":0535
      ScaleHeight     =   420
      ScaleWidth      =   270
      TabIndex        =   6
      Top             =   1545
      Width           =   330
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   330
      Left            =   3285
      MaxLength       =   1
      TabIndex        =   5
      Top             =   1785
      Width           =   270
   End
   Begin VB.PictureBox pDig 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   480
      Index           =   2
      Left            =   6030
      Picture         =   "Form1.frx":062D
      ScaleHeight     =   420
      ScaleWidth      =   270
      TabIndex        =   4
      Top             =   1065
      Width           =   330
   End
   Begin VB.PictureBox pDig 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   480
      Index           =   0
      Left            =   6030
      Picture         =   "Form1.frx":072F
      ScaleHeight     =   420
      ScaleWidth      =   270
      TabIndex        =   3
      Top             =   90
      Width           =   330
   End
   Begin VB.PictureBox pDig 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   480
      Index           =   1
      Left            =   6030
      Picture         =   "Form1.frx":0840
      ScaleHeight     =   420
      ScaleWidth      =   270
      TabIndex        =   2
      Top             =   570
      Width           =   330
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   150
      Top             =   4035
   End
   Begin VB.PictureBox p2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   4920
      Picture         =   "Form1.frx":08E4
      ScaleHeight     =   28
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   18
      TabIndex        =   1
      Top             =   300
      Width           =   300
   End
   Begin VB.PictureBox p1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   750
      Left            =   1005
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   34
      TabIndex        =   0
      Top             =   360
      Width           =   540
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1545
      MousePointer    =   15  'Size All
      TabIndex        =   35
      Top             =   1125
      Width           =   165
   End
   Begin VB.Label Label2 
      Caption         =   "PixBox can be any size,but larger is slower."
      Height          =   240
      Left            =   75
      TabIndex        =   14
      Top             =   45
      Width           =   3075
   End
   Begin VB.Label Label1 
      Caption         =   "Textbox...numbers only"
      Height          =   195
      Left            =   1545
      TabIndex        =   13
      Top             =   1845
      Width           =   1710
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'****************************************************
'*
'*         Project Name :PlgBlt---digit effects demo
'*        Version Number: 1.0.1
'*           Author Name: Ken Foster
'*                 Date : April 21, 2009
'*        Freeware - Use anyway you want.
'*
'****************************************************

'To use T.O.P for a search , highlight just the word or words you want to search for.
'Press Ctrl + F3
' To continue search , just use F3 until all uses are found.

'***************** Table of Procedures *************
'   Private Sub Form_Load
'   Private Sub cmdAdd_Click
'   Private Sub cmdColor_Click
'   Private Sub cmdDig_Click
'   Private Sub optDir_Click
'   Private Sub Text1_KeyPress
'   Private Sub Text1_KeyUp
'   Private Sub Timer1_Timer
'***************** End of Table ********************

Option Explicit
   
Private Type POINTAPI
   X As Long
   Y As Long
End Type

Dim Ft(2) As POINTAPI
Dim ret As Long
Dim ret1 As Long
Dim cc As Integer
Dim ad As Integer
Dim ac As Integer
Dim Selected As Integer
Dim Resz As Boolean

Private Declare Function PlgBlt Lib "gdi32" (ByVal hdcDest As Long, lpPoint As POINTAPI, ByVal hdcSrc As Long, ByVal nXSrc As Long, ByVal nYSrc As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hbmMask As Long, ByVal xMask As Long, ByVal yMask As Long) As Long

Private Sub Form_Load()
If App.PrevInstance = True Then
    MsgBox "There Is Already Another Instance Of This Application Running, Please Close It And Try Again.", vbExclamation, "Error"
    End
End If
Resz = False
   'put a 0 in picbox to start things up
   Ft(0).X = 0
   Ft(0).Y = 0
   Ft(1).X = p1.ScaleWidth
   Ft(1).Y = 0
   Ft(2).X = 0
   Ft(2).Y = p1.ScaleHeight
   ret = PlgBlt(p1.hDC, Ft(0), p2.hDC, 0, 0, p2.ScaleWidth, p2.ScaleHeight, 0, 0, 0)
   ret1 = PlgBlt(p3.hDC, Ft(0), p2.hDC, 0, 0, p2.ScaleWidth, p2.ScaleHeight, 0, 0, 0)
   p1.Refresh
End Sub

Private Sub cmdAdd_Click()
   ad = ad + 1
   Resz = False
   If ad > 9 Then
      ad = 0
      ac = ac + 1
      If ac > 9 Then ac = 0
   End If
      p2.Picture = pDig(ad).Picture
   Timer1.Enabled = True
End Sub

Private Sub cmdColor_Click()
   Dim colchg As Long
   
   colchg = ShowColor
   If colchg = -1 Then Exit Sub
   p1.BackColor = colchg
   p2.BackColor = colchg
   'load last selected digit back into picbox
   Ft(0).X = 0
   Ft(0).Y = 0
   Ft(1).X = p1.ScaleWidth
   Ft(1).Y = 0
   Ft(2).X = 0
   Ft(2).Y = p1.ScaleHeight
   ret = PlgBlt(p1.hDC, Ft(0), p2.hDC, 0, 0, p2.ScaleWidth, p2.ScaleHeight, 0, 0, 0)
   p1.Refresh
End Sub

Private Sub cmdDig_Click(Index As Integer)
   p2.Picture = pDig(Index).Picture
   ad = Index
   Timer1.Enabled = True
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = 1 Then
      Label3.Left = Label3.Left + X
      Label3.Top = Label3.Top + Y
      p1.Width = p1.Width + X
      p1.Height = p1.Height + Y
      p3.Width = p1.Width + X
      p3.Height = p1.Height + Y
      Resz = True
   End If
End Sub

Private Sub Label3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   p3.Picture = LoadPicture()
   Timer1.Enabled = True
End Sub

Private Sub optDir_Click(Index As Integer)
   Selected = Index
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   'input numbers only
   If KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Or KeyAscii = 8 Then
      Text1.Text = ""
      Exit Sub
   Else
      KeyAscii = 0
   End If
   
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
   If Text1.Text = "" Then Exit Sub
   p2.Picture = pDig(Int(Text1.Text)).Picture
   Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
   p1.Picture = LoadPicture()
   Select Case Selected
      Case 0
         Ft(0).X = 0
         Ft(0).Y = 0
         
         Ft(1).X = cc
         Ft(1).Y = 0
         
         Ft(2).X = 0
         Ft(2).Y = p1.ScaleHeight
         
         ret = PlgBlt(p1.hDC, Ft(0), p2.hDC, 0, 0, p2.ScaleWidth, p2.ScaleHeight, 0, 0, 0)
         If ad = 0 Or Resz = True Then ret1 = PlgBlt(p3.hDC, Ft(0), pDig(ac).hDC, 0, 0, p2.ScaleWidth, p2.ScaleHeight, 0, 0, 0)
         If cc > p1.ScaleWidth Then
            Timer1.Enabled = False
            cc = 0
         End If
         
      Case 1
         Ft(0).X = p1.ScaleWidth - cc
         Ft(0).Y = 0
         
         Ft(1).X = p1.ScaleWidth
         Ft(1).Y = 0
         
         Ft(2).X = p1.ScaleWidth - cc
         Ft(2).Y = p1.ScaleHeight
         
         ret = PlgBlt(p1.hDC, Ft(0), p2.hDC, 0, 0, p2.ScaleWidth, p2.ScaleHeight, 0, 0, 0)
         If ad = 0 Or Resz = True Then ret1 = PlgBlt(p3.hDC, Ft(0), pDig(ac).hDC, 0, 0, p2.ScaleWidth, p2.ScaleHeight, 0, 0, 0)
         If cc > p1.ScaleWidth Then
            Timer1.Enabled = False
            cc = 0
         End If
         
      Case 2
         Ft(0).X = 0
         Ft(0).Y = 0
         
         Ft(1).X = p1.ScaleWidth
         Ft(1).Y = 0
         
         Ft(2).X = 0
         Ft(2).Y = cc
         
         ret = PlgBlt(p1.hDC, Ft(0), p2.hDC, 0, 0, p2.ScaleWidth, p2.ScaleHeight, 0, 0, 0)
         If ad = 0 Or Resz = True Then ret1 = PlgBlt(p3.hDC, Ft(0), pDig(ac).hDC, 0, 0, p2.ScaleWidth, p2.ScaleHeight, 0, 0, 0)
         If cc > p1.ScaleHeight Then
            Timer1.Enabled = False
            cc = 0
         End If
         
      Case 3
         Ft(0).X = 0
         Ft(0).Y = p1.ScaleHeight - cc
         
         Ft(1).X = p1.ScaleWidth
         Ft(1).Y = p1.ScaleHeight - cc
         
         Ft(2).X = 0
         Ft(2).Y = p1.ScaleHeight
         
         ret = PlgBlt(p1.hDC, Ft(0), p2.hDC, 0, 0, p2.ScaleWidth, p2.ScaleHeight, 0, 0, 0)
         If ad = 0 Or Resz = True Then ret1 = PlgBlt(p3.hDC, Ft(0), pDig(ac).hDC, 0, 0, p2.ScaleWidth, p2.ScaleHeight, 0, 0, 0)
         If cc > p1.ScaleHeight Then
            Timer1.Enabled = False
            cc = 0
         End If
         
      Case 4
         Ft(0).X = 0
         Ft(0).Y = p1.ScaleHeight - cc
         
         Ft(1).X = p1.ScaleWidth
         Ft(1).Y = -1
         
         Ft(2).X = 0
         Ft(2).Y = p1.ScaleHeight
         
         ret = PlgBlt(p1.hDC, Ft(0), p2.hDC, 0, 0, p2.ScaleWidth, p2.ScaleHeight, 0, 0, 0)
         If ad = 0 Or Resz = True Then ret1 = PlgBlt(p3.hDC, Ft(0), pDig(ac).hDC, 0, 0, p2.ScaleWidth, p2.ScaleHeight, 0, 0, 0)
         If cc > p1.ScaleHeight Then
            Timer1.Enabled = False
            cc = 0
         End If
         
      Case 5
         Ft(0).X = 0
         Ft(0).Y = p1.ScaleHeight - cc
         
         Ft(1).X = p1.ScaleWidth / 2 + (cc - (p1.ScaleHeight - p1.ScaleWidth / 2))
         Ft(1).Y = -1
         
         Ft(2).X = 0
         Ft(2).Y = p1.ScaleHeight
         
         ret = PlgBlt(p1.hDC, Ft(0), p2.hDC, 0, 0, p2.ScaleWidth, p2.ScaleHeight, 0, 0, 0)
         If ad = 0 Or Resz = True Then ret1 = PlgBlt(p3.hDC, Ft(0), pDig(ac).hDC, 0, 0, p2.ScaleWidth, p2.ScaleHeight, 0, 0, 0)
         If cc > p1.ScaleHeight Then
            Timer1.Enabled = False
            cc = 0
         End If
         
      Case 6
         Ft(0).X = p1.ScaleWidth / 2 - cc
         Ft(0).Y = 0
         
         Ft(1).X = p1.ScaleWidth / 2 + cc
         Ft(1).Y = 0
         
         Ft(2).X = p1.ScaleWidth / 2 - cc
         Ft(2).Y = p1.ScaleHeight
         
         ret = PlgBlt(p1.hDC, Ft(0), p2.hDC, 0, 0, p2.ScaleWidth, p2.ScaleHeight, 0, 0, 0)
         If ad = 0 Or Resz = True Then ret1 = PlgBlt(p3.hDC, Ft(0), pDig(ac).hDC, 0, 0, p2.ScaleWidth, p2.ScaleHeight, 0, 0, 0)
         If cc > p1.ScaleWidth / 2 Then
            Timer1.Enabled = False
            cc = 0
         End If
      Case 7
         Ft(0).X = p1.ScaleWidth / 2 - cc
         Ft(0).Y = p1.ScaleHeight / 2 - (cc * (p1.ScaleHeight / p1.ScaleWidth))
         
         Ft(1).X = p1.ScaleWidth / 2 + cc
         Ft(1).Y = p1.ScaleHeight / 2 - (cc * (p1.ScaleHeight / p1.ScaleWidth))
         
         Ft(2).X = p1.ScaleWidth / 2 - cc
         Ft(2).Y = p1.ScaleHeight / 2 + (cc * (p1.ScaleHeight / p1.ScaleWidth))
         
         ret = PlgBlt(p1.hDC, Ft(0), p2.hDC, 0, 0, p2.ScaleWidth, p2.ScaleHeight, 0, 0, 0)
         If ad = 0 Or Resz = True Then ret1 = PlgBlt(p3.hDC, Ft(0), pDig(ac).hDC, 0, 0, p2.ScaleWidth, p2.ScaleHeight, 0, 0, 0)
         If cc > p1.ScaleHeight / 2 - 7 Then
            Timer1.Enabled = False
            cc = 0
         End If
   End Select
   cc = cc + 1
   p1.Refresh
   p3.Refresh
End Sub
