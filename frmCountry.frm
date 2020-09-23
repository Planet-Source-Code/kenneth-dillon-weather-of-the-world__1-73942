VERSION 5.00
Object = "{7888C00A-4808-4D27-9AAE-BD36EC13D16F}#1.0#0"; "LVbuttons.ocx"
Begin VB.Form frmCountry 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   1170
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   945
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1170
   ScaleWidth      =   945
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin LVbuttons.LaVolpeButton cmdAnimate 
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   2040
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Play Animation"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      BCOL            =   15790320
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   99
      MICON           =   "frmCountry.frx":0000
      ALIGN           =   1
      IMGLST          =   "(None)"
      IMGICON         =   "(None)"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   240
      Top             =   1680
   End
   Begin VB.Image picSource 
      Height          =   1095
      Left            =   960
      Top             =   120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image imgLgCountry 
      Height          =   1080
      Left            =   45
      MousePointer    =   99  'Custom
      ToolTipText     =   "Right Click To Close"
      Top             =   45
      Visible         =   0   'False
      Width           =   840
   End
End
Attribute VB_Name = "frmCountry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim PicRatio As Single
Dim BoxWidth As Integer
Dim BoxHeight As Integer
Dim zooming As Boolean
Dim MinLeft As Integer
Dim MinTop As Integer
Dim MaxLeft As Integer
Dim MaxTop As Integer

Private Sub cmdAnimate_Click()
   If cmdAnimate.Caption = "Exit" Then
      Unload Me
   Else
      Animation = True
      Timer1.Enabled = True
   End If
End Sub

Private Sub Form_Load()
  Set cmdAnimate.MouseIcon = frmWeatherMain.ImageList1.ListImages(3).Picture
  frmCountry.Caption = sFrmName ' & " Of " & scntName
  SizePic PictureName
  frmCountry.Show vbModal
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set frmCountry = Nothing
End Sub

Private Sub imgLgCountry_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
  If Button = 2 Then
    Timer1.Enabled = True
  End If
End Sub

Private Sub SizePic(PicName As String)
  'load first pic in in picSource, get ratio,
  'size imgLgCountry, send size, empty box1
  'On Error Resume Next
  picSource.Picture = LoadPicture(PicName)
  PicRatio = picSource.Width / picSource.Height
  
  If PicRatio > 1.33 Then 'pic is landscape
    BoxWidth = Screen.Width / 2.3
    BoxHeight = (Screen.Width / PicRatio) / 2.3
  End If

  If PicRatio < 1.33 Then
    BoxHeight = Screen.Height / 2.3 'pic is portrait
    BoxWidth = (Screen.Height * PicRatio) / 2.3
  End If

  If PicRatio = 1.33 Then 'pic is square
    BoxHeight = Screen.Height / 2.3
    BoxWidth = Screen.Width / 2.3
  End If
  
  Call ShowPic(BoxWidth, BoxHeight, PicName)
  imgLgCountry.Visible = True
End Sub

Public Sub ShowPic(BoxWidth As Integer, BoxHeight As Integer, PicName As String)
  'empty box2, size box2,load imgLgCountry from  box1
  imgLgCountry.Visible = False
  imgLgCountry.Height = BoxHeight
  imgLgCountry.Width = BoxWidth
  imgLgCountry.Picture = LoadPicture(PicName)
  picSource.Picture = LoadPicture()
  imgLgCountry.Top = 0
  imgLgCountry.Left = 5
  frmCountry.Height = imgLgCountry.Height + 360
   frmCountry.Width = imgLgCountry.Width + 110
   If PlayAnimation Then
      cmdAnimate.Caption = "Play Animation"
      cmdAnimate.Visible = True
   End If
   cmdAnimate.Top = frmCountry.Height - 850
   cmdAnimate.Left = (frmCountry.Width / 2) - (cmdAnimate.Width / 2)
End Sub

Public Sub ZoomPicture(BoxWidth As Integer, BoxHeight As Integer)
   Dim x, Y As Single
   On Error Resume Next
  
   x = BoxWidth
   Y = BoxHeight
   x = x / 1.0351
   Y = Y / 1.0351
   BoxWidth = x
   BoxHeight = Y
    
   Call ShowZoom(BoxWidth, BoxHeight, PictureName)
   'Center picture
   frmCountry.Left = (Screen.Width / 2) - (frmCountry.Width / 2)
   frmCountry.Top = (Screen.Height / 2) - (frmCountry.Height / 2)
   imgLgCountry.Visible = True
   If Y < 425 Or x < 200 Then
      Timer1.Enabled = False
      Unload Me
   End If
End Sub

Private Sub Timer1_Timer()
   Call ZoomPicture(imgLgCountry.Width, imgLgCountry.Height)
End Sub

Private Sub ShowZoom(BoxWidth As Integer, BoxHeight As Integer, PicName As String)
   imgLgCountry.Height = BoxHeight
   imgLgCountry.Width = BoxWidth
   frmCountry.Height = imgLgCountry.Height + 300
   frmCountry.Width = imgLgCountry.Width + 25
End Sub
