VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   Caption         =   "Image Viewer"
   ClientHeight    =   6150
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7740
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6150
   ScaleWidth      =   7740
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog dlgOpenImg 
      Left            =   7080
      Top             =   6360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6135
      Left            =   0
      ScaleHeight     =   6135
      ScaleWidth      =   7695
      TabIndex        =   0
      Top             =   0
      Width           =   7695
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   5175
         Left            =   120
         TabIndex        =   5
         Top             =   840
         Width           =   7455
         Begin VB.Image Image1 
            Height          =   5175
            Left            =   0
            Top             =   0
            Width           =   7455
         End
      End
      Begin VB.CommandButton btnOpen 
         Caption         =   "&Open Image"
         Height          =   375
         Left            =   6360
         TabIndex        =   2
         Top             =   120
         Width           =   1215
      End
      Begin VB.PictureBox Picture2 
         Height          =   735
         Left            =   120
         ScaleHeight     =   675
         ScaleWidth      =   2235
         TabIndex        =   1
         Top             =   0
         Visible         =   0   'False
         Width           =   2295
         Begin VB.OptionButton optUnStretch 
            DownPicture     =   "frmMain.frx":08CA
            Height          =   495
            Left            =   1680
            Picture         =   "frmMain.frx":0AD8
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   120
            Width           =   495
         End
         Begin VB.OptionButton optStretch 
            DownPicture     =   "frmMain.frx":0BB6
            Height          =   495
            Left            =   1080
            Picture         =   "frmMain.frx":0DCB
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   120
            Width           =   495
         End
         Begin VB.Image imgZoomOut 
            Height          =   240
            Left            =   600
            Picture         =   "frmMain.frx":0EA6
            Top             =   240
            Width           =   240
         End
         Begin VB.Image imgZoomIn 
            Height          =   240
            Left            =   120
            Picture         =   "frmMain.frx":1026
            Top             =   240
            Width           =   240
         End
      End
      Begin VB.Label Label2 
         Caption         =   "First Open an Image"
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Image Viewer         By: Sergio del Rio"
         Height          =   495
         Left            =   2520
         TabIndex        =   6
         Top             =   120
         Width           =   1335
      End
   End
   Begin VB.Image imgUnStretch2 
      Height          =   330
      Left            =   3960
      Picture         =   "frmMain.frx":1277
      Top             =   6360
      Width           =   360
   End
   Begin VB.Image imgUnStretch1 
      Height          =   180
      Left            =   3600
      Picture         =   "frmMain.frx":1485
      Top             =   6360
      Width           =   210
   End
   Begin VB.Image imgStretch2 
      Height          =   330
      Left            =   3120
      Picture         =   "frmMain.frx":1563
      Top             =   6360
      Width           =   360
   End
   Begin VB.Image imgStretch1 
      Height          =   180
      Left            =   2760
      Picture         =   "frmMain.frx":1778
      Top             =   6360
      Width           =   210
   End
   Begin VB.Image imgZoomOut3 
      Height          =   330
      Left            =   2280
      Picture         =   "frmMain.frx":1853
      Top             =   6360
      Width           =   360
   End
   Begin VB.Image imgZoomIn3 
      Height          =   330
      Left            =   960
      Picture         =   "frmMain.frx":1B69
      Top             =   6360
      Width           =   360
   End
   Begin VB.Image imgZoomOut2 
      Height          =   330
      Left            =   1800
      Picture         =   "frmMain.frx":1E8F
      Top             =   6360
      Width           =   360
   End
   Begin VB.Image imgZoomOut1 
      Height          =   240
      Left            =   1440
      Picture         =   "frmMain.frx":219B
      Top             =   6360
      Width           =   240
   End
   Begin VB.Image imgZoomIn1 
      Height          =   240
      Left            =   120
      Picture         =   "frmMain.frx":231B
      Top             =   6360
      Width           =   240
   End
   Begin VB.Image imgZoomIn2 
      Height          =   330
      Left            =   480
      Picture         =   "frmMain.frx":256C
      Top             =   6360
      Width           =   360
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnOpen_Click()
        Image1.Height = "5175"
        Image1.Width = "7455"
        Image1.Top = "0"
        Image1.Left = "0"
        Image1.Stretch = False
        optStretch.Value = False
        optUnStretch.Value = False
    
    With dlgOpenImg
        .DialogTitle = "Open Image"
        .CancelError = False
        .Filter = "Image files (*.*)|*.*"
        .ShowOpen
        Image1.Picture = LoadPicture(dlgOpenImg.FileName)
    End With
    Picture2.Visible = True
    
    If Image1.Picture.Height > "5175" Then
        Image1.Stretch = True
        optStretch.Value = True
    Else
        Image1.Stretch = False
    End If
End Sub


Private Sub imgZoomIn_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Error
If imgZoomOut.Enabled = False Then
    imgZoomOut.Enabled = True
End If
    imgZoomIn.Picture = imgZoomIn2.Picture
    Image1.Height = Image1.Height + 600
    Image1.Width = Image1.Width + 600
    Image1.Top = Image1.Top - 300
    Image1.Left = Image1.Left - 300
Exit Sub
Error:
imgZoomIn.Enabled = False
End Sub

Private Sub imgZoomIn_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgZoomIn.Picture = imgZoomIn3.Picture
End Sub

Private Sub imgZoomIn_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgZoomIn.Picture = imgZoomIn1.Picture
End Sub


Private Sub imgZoomOut_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Error
    imgZoomOut.Picture = imgZoomOut2.Picture
    Image1.Height = Image1.Height - 600
    Image1.Width = Image1.Width - 600
    Image1.Top = Image1.Top + 300
    Image1.Left = Image1.Left + 300
Exit Sub
Error:
imgZoomOut.Enabled = False
End Sub

Private Sub imgZoomOut_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgZoomOut.Picture = imgZoomOut3.Picture
End Sub

Private Sub imgZoomOut_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgZoomIn.Picture = imgZoomIn1.Picture
End Sub

Private Sub optStretch_Click()
    Image1.Stretch = True
End Sub

Private Sub optUnStretch_Click()
    Image1.Stretch = False
End Sub
Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        imgZoomIn.Picture = imgZoomIn1.Picture
        imgZoomOut.Picture = imgZoomOut1.Picture
End Sub

Private Sub Timer1_Timer()
    If Stretch.Text = "True" Then
        Image1.Stretch = True
    End If
    If Stretch.Text = "False" Then
        Image1.Stretch = False
    End If
Timer1.Enabled = True
End Sub
