VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7245
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   6165
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   483
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   411
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdFlip 
      Caption         =   "Flip"
      Height          =   375
      Left            =   4920
      TabIndex        =   9
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CommandButton cmdCut 
      Caption         =   "Cut"
      Height          =   375
      Left            =   4920
      TabIndex        =   8
      Top             =   120
      Width           =   1095
   End
   Begin VB.CheckBox chkSelect 
      Caption         =   "Selection On"
      Height          =   255
      Left            =   4800
      TabIndex        =   6
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5280
      Top             =   2640
   End
   Begin VB.PictureBox picMerge 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   4680
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   97
      TabIndex        =   5
      Top             =   2880
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.PictureBox SelHolder 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   4680
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   97
      TabIndex        =   4
      Top             =   2400
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdPaste 
      Caption         =   "Paste"
      Height          =   375
      Left            =   4920
      TabIndex        =   3
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "Copy"
      Height          =   375
      Left            =   4920
      TabIndex        =   1
      Top             =   600
      Width           =   1095
   End
   Begin VB.PictureBox picBG 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   7005
      Left            =   120
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   467
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   300
      TabIndex        =   0
      Top             =   120
      Width           =   4500
      Begin VB.PictureBox picSelect 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         BorderStyle     =   0  'None
         Height          =   1335
         Left            =   2760
         ScaleHeight     =   89
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   105
         TabIndex        =   2
         Top             =   120
         Visible         =   0   'False
         Width           =   1575
         Begin VB.Shape PicSelectShape 
            BorderStyle     =   3  'Dot
            DrawMode        =   6  'Mask Pen Not
            Height          =   495
            Left            =   120
            Top             =   120
            Width           =   1335
         End
         Begin VB.Image SelectImage 
            Height          =   495
            Left            =   120
            Stretch         =   -1  'True
            Top             =   720
            Width           =   1335
         End
      End
      Begin VB.PictureBox picBlank 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   120
         ScaleHeight     =   615
         ScaleWidth      =   975
         TabIndex        =   7
         Top             =   1560
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Shape SelectShape 
         BorderStyle     =   3  'Dot
         DrawMode        =   6  'Mask Pen Not
         FillColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   120
         Top             =   840
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Image Image1 
         Height          =   615
         Left            =   120
         Stretch         =   -1  'True
         Top             =   120
         Width           =   975
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal Nwidth As Long, ByVal Nheight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Dim Selecting As Boolean
Dim SelX, SelY As Long
Dim SelectWidth, SelectHeight, SelectTop, SelectLeft As Double
Dim BSelectWidth, BSelectHeight, BSelectTop, BSelectLeft As Double
Dim SelectPicX, SelectPicY As Integer
Dim P As POINTAPI
Private Type POINTAPI
    X As Long
    Y As Long
End Type
Const SRCCOPY = &HCC0020

Private Sub chkSelect_Click()
    If chkSelect.Value = 0 Then
        Selecting = False
    End If
End Sub

Private Sub cmdCopy_Click()
    Clipboard.Clear
    Clipboard.SetData SelectImage.Picture
End Sub

Private Sub cmdFlip_Click()
    If picSelect.Visible = True Then
        With SelHolder
            .PaintPicture .Picture, 0, .ScaleHeight, .ScaleWidth, -1 * .ScaleHeight
            .Picture = .Image
        End With
        SelectImage.Picture = SelHolder.Image
    Else
        With picMerge
            .PaintPicture .Picture, 0, .ScaleHeight, .ScaleWidth, -1 * .ScaleHeight
            .Picture = .Image
        End With
        Image1.Picture = picMerge.Image
    End If
End Sub

Private Sub Form_Load()
    picMerge.Picture = picBG.Image
    picBG.Top = Me.ScaleTop
    picBG.Left = Me.ScaleLeft
    Image1.Top = picBG.ScaleTop
    Image1.Left = picBG.ScaleLeft
    Image1.Width = picBG.Width
    Image1.Height = picBG.Height
    Image1.Picture = picMerge.Image
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    X = X / Screen.TwipsPerPixelX
    Y = Y / Screen.TwipsPerPixelY
    If Button = 1 Then
        If picSelect.Visible = True Then
            SelectLeft = picSelect.Left
            SelectTop = picSelect.Top
            SelectWidth = picSelect.Width
            SelectHeight = picSelect.Height
            BSelectLeft = picBlank.Left
            BSelectTop = picBlank.Top
            BSelectWidth = picBlank.Width
            BSelectHeight = picBlank.Height
            picSelect.Visible = False
            SelectShape.Visible = False
            StretchBlt picMerge.hdc, BSelectLeft, BSelectTop, BSelectWidth, BSelectHeight, picBlank.hdc, 0, 0, BSelectWidth, BSelectHeight, SRCCOPY
            Image1.Picture = picMerge.Image
            StretchBlt picMerge.hdc, SelectLeft, SelectTop, SelectWidth, SelectHeight, SelHolder.hdc, 0, 0, SelectWidth, SelectHeight, SRCCOPY
            Image1.Picture = picMerge.Image
            picBlank.Visible = False
            Selecting = False
        End If
        
        If chkSelect.Value = 1 Then 'Checked
            Selecting = True
            SelX = X
            SelY = Y
            SelectShape.Left = Int(X)
            SelectShape.Top = Int(Y)
            SelectShape.Width = 0
            SelectShape.Height = 0
            SelectShape.Visible = True
        End If
    End If
    If Button = 2 Then
        SelectLeft = picSelect.Left
        SelectTop = picSelect.Top
        SelectWidth = picSelect.Width
        SelectHeight = picSelect.Height
        BSelectLeft = picBlank.Left
        BSelectTop = picBlank.Top
        BSelectWidth = picBlank.Width
        BSelectHeight = picBlank.Height
        picSelect.Visible = False
        SelectShape.Visible = False
        StretchBlt picMerge.hdc, BSelectLeft, BSelectTop, BSelectWidth, BSelectHeight, picBlank.hdc, 0, 0, BSelectWidth, BSelectHeight, SRCCOPY
        Image1.Picture = picMerge.Image
        StretchBlt picMerge.hdc, SelectLeft, SelectTop, SelectWidth, SelectHeight, SelHolder.hdc, 0, 0, SelectWidth, SelectHeight, SRCCOPY
        Image1.Picture = picMerge.Image
        picBlank.Visible = False
        Selecting = False
    End If
    Image1.Picture = picMerge.Image
    Image1.Refresh
    picMerge.Picture = picMerge.Image
    picMerge.Refresh
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    X = Int(X / Screen.TwipsPerPixelX)
    Y = Int(Y / Screen.TwipsPerPixelY)
    If Selecting Then
        Image1.MousePointer = vbCrosshair
    Else
        Image1.MousePointer = vbDefault
    End If
    If Selecting And Button = 1 Then
        If X >= SelX Then
            If Y >= SelY Then
                If X >= Image1.Width And Y >= Image1.Height Then
                    SelectShape.Top = SelY
                    SelectShape.Left = SelX
                    SelectShape.Height = Image1.Height - SelY
                    SelectShape.Width = Image1.Width - SelX
                ElseIf X >= Image1.Width Then
                    SelectShape.Top = SelY
                    SelectShape.Left = SelX
                    SelectShape.Height = Y - SelY
                    SelectShape.Width = Image1.Width - SelX
                ElseIf Y >= Image1.Height Then
                    SelectShape.Top = SelY
                    SelectShape.Left = SelX
                    SelectShape.Height = Image1.Height - SelY
                    SelectShape.Width = X - SelX
                Else
                    SelectShape.Top = SelY
                    SelectShape.Left = SelX
                    SelectShape.Height = Y - SelY
                    SelectShape.Width = X - SelX
                End If
            Else
                If X >= Image1.Width And Y <= Image1.Top Then
                    SelectShape.Top = 0
                    SelectShape.Left = SelX
                    SelectShape.Height = SelY
                    SelectShape.Width = Image1.Width - SelX
                ElseIf X >= Image1.Width Then
                    SelectShape.Top = Y
                    SelectShape.Left = SelX
                    SelectShape.Height = SelY - Y
                    SelectShape.Width = Image1.Width - SelX
                ElseIf Y <= Image1.Top Then
                    SelectShape.Top = 0
                    SelectShape.Left = SelX
                    SelectShape.Height = SelY
                    SelectShape.Width = X - SelX
                Else
                    SelectShape.Top = Y
                    SelectShape.Left = SelX
                    SelectShape.Height = SelY - Y
                    SelectShape.Width = X - SelX
                End If
            End If
        Else
            If Y >= SelY Then
                If X <= Image1.Left And Y >= Image1.Height Then
                    SelectShape.Top = SelY
                    SelectShape.Left = 0
                    SelectShape.Height = Image1.Height - SelY
                    SelectShape.Width = SelX
                ElseIf X <= Image1.Left Then
                    SelectShape.Top = SelY
                    SelectShape.Left = 0
                    SelectShape.Height = Y - SelY
                    SelectShape.Width = SelX
                ElseIf Y >= Image1.Height Then
                    SelectShape.Top = SelY
                    SelectShape.Left = X
                    SelectShape.Height = Image1.Height - SelY
                    SelectShape.Width = SelX - X
                Else
                    SelectShape.Top = SelY
                    SelectShape.Left = X
                    SelectShape.Height = Y - SelY
                    SelectShape.Width = SelX - X
                End If
            Else
                If X <= Image1.Left And Y <= Image1.Top Then
                    SelectShape.Top = 0
                    SelectShape.Left = 0
                    SelectShape.Height = SelY
                    SelectShape.Width = SelX
                ElseIf X <= Image1.Left Then
                    SelectShape.Top = Y
                    SelectShape.Left = 0
                    SelectShape.Height = SelY - Y
                    SelectShape.Width = SelX
                ElseIf Y <= Image1.Top Then
                    SelectShape.Top = 0
                    SelectShape.Left = X
                    SelectShape.Height = SelY
                    SelectShape.Width = SelX - X
                Else
                    SelectShape.Top = Y
                    SelectShape.Left = X
                    SelectShape.Height = SelY - Y
                    SelectShape.Width = SelX - X
                End If
            End If
        End If
        SelectShape.Visible = True
        Image1.Picture = picMerge.Image
        Image1.Refresh
    End If
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    X = Int(X / Screen.TwipsPerPixelX)
    Y = Int(Y / Screen.TwipsPerPixelY)
    If Button = 1 Then
        If Selecting = True Then
            If X = SelX And Y = SelY Then
                Exit Sub
            End If
            picSelect.Left = SelectShape.Left
            picSelect.Top = SelectShape.Top
            picSelect.Height = SelectShape.Height
            picSelect.Width = SelectShape.Width
            SelectLeft = SelectShape.Left
            SelectTop = SelectShape.Top
            SelectWidth = SelectShape.Width
            SelectHeight = SelectShape.Height
            myinternalcopy picMerge, SelHolder, PicSelectShape
            SelectImage.Left = 0
            SelectImage.Top = 0
            SelectImage.Width = picSelect.Width
            SelectImage.Height = picSelect.Height
            SelectImage.Picture = SelHolder.Image 'Copy image
            picBlank.Left = picSelect.Left
            picBlank.Top = picSelect.Top
            picBlank.Width = picSelect.Width
            picBlank.Height = picSelect.Height
            picBlank.Visible = True
            PicSelectShape.Left = 0
            PicSelectShape.Top = 0
            PicSelectShape.Width = picSelect.Width
            PicSelectShape.Height = picSelect.Height
            picSelect.Visible = True
            SelectShape.Visible = False
            Selecting = False
            GoTo Woops
        End If
    End If
    Image1.Picture = picMerge.Image
    Image1.Refresh
Woops:
End Sub

Public Sub myinternalcopy(Spic As PictureBox, Dpic As PictureBox, TempselShape As Shape)
    Dpic.Width = SelectWidth
    Dpic.Height = SelectHeight
    Dpic.Picture = LoadPicture()
    StretchBlt Dpic.hdc, 0, 0, SelectWidth, SelectHeight, Spic.hdc, SelectLeft, SelectTop, SelectWidth, SelectHeight, SRCCOPY
    Dpic.Picture = Dpic.Image
    Dpic.Refresh
End Sub

Private Sub picBlank_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        SelectLeft = picSelect.Left
        SelectTop = picSelect.Top
        SelectWidth = picSelect.Width
        SelectHeight = picSelect.Height
        BSelectLeft = picBlank.Left
        BSelectTop = picBlank.Top
        BSelectWidth = picBlank.Width
        BSelectHeight = picBlank.Height
        picSelect.Visible = False
        SelectShape.Visible = False
        StretchBlt picMerge.hdc, BSelectLeft, BSelectTop, BSelectWidth, BSelectHeight, picBlank.hdc, 0, 0, BSelectWidth, BSelectHeight, SRCCOPY
        Image1.Picture = picMerge.Image
        StretchBlt picMerge.hdc, SelectLeft, SelectTop, SelectWidth, SelectHeight, SelHolder.hdc, 0, 0, SelectWidth, SelectHeight, SRCCOPY
        Image1.Picture = picMerge.Image
        picBlank.Visible = False
        Selecting = False
    End If
    Image1.Picture = picMerge.Image
    Image1.Refresh
End Sub

Private Sub picSelect_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If picSelect.Visible = True Then
            picSelect.ScaleMode = 2
            GetCursorPos P
            P.X = Int(P.X)
            P.Y = Int(P.Y)
            SelectPicX = P.X - picSelect.Left
            SelectPicY = P.Y - picSelect.Top
            Timer1.Enabled = True
        End If
    End If
    If Button = 2 Then
        SelectLeft = picSelect.Left
        SelectTop = picSelect.Top
        SelectWidth = picSelect.Width
        SelectHeight = picSelect.Height
        BSelectLeft = picBlank.Left
        BSelectTop = picBlank.Top
        BSelectWidth = picBlank.Width
        BSelectHeight = picBlank.Height
        picSelect.Visible = False
        SelectShape.Visible = False
        StretchBlt picMerge.hdc, BSelectLeft, BSelectTop, BSelectWidth, BSelectHeight, picBlank.hdc, 0, 0, BSelectWidth, BSelectHeight, SRCCOPY
        Image1.Picture = picMerge.Image
        StretchBlt picMerge.hdc, SelectLeft, SelectTop, SelectWidth, SelectHeight, SelHolder.hdc, 0, 0, SelectWidth, SelectHeight, SRCCOPY
        Image1.Picture = picMerge.Image
        picBlank.Visible = False
        Selecting = False
    End If
End Sub

Private Sub picSelect_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Timer1.Enabled = False
    picSelect.ScaleMode = 3
    SelectLeft = picSelect.Left
    SelectTop = picSelect.Top
    SelectWidth = picSelect.Width
    SelectHeight = picSelect.Height
End Sub

Private Sub SelectImage_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picSelect_MouseDown Button, Shift, X, Y
End Sub

Private Sub SelectImage_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SelectImage.MousePointer = 15
End Sub

Private Sub SelectImage_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picSelect_MouseUp Button, Shift, X, Y
End Sub

Private Sub Timer1_Timer()
    GetCursorPos P
    P.X = Int(P.X)
    P.Y = Int(P.Y)
    picSelect.Left = P.X - SelectPicX
    picSelect.Top = P.Y - SelectPicY
End Sub

Private Sub cmdPaste_Click()
    picBlank.Visible = False
    Me.Enabled = False
    MousePointer = 11
    Masterpasting = True
    picSelect.Left = 0
    picSelect.Top = 0
    SelHolder.Picture = Clipboard.GetData(vbCFBitmap)
    picSelect.Width = SelHolder.Width
    picSelect.Height = SelHolder.Height
    SelectImage.Left = 0
    SelectImage.Top = 0
    SelectImage.Width = picSelect.Width
    SelectImage.Height = picSelect.Height
    SelectImage.Picture = SelHolder.Image
    PicSelectShape.Left = 0
    PicSelectShape.Top = 0
    PicSelectShape.Height = picSelect.Height
    PicSelectShape.Width = picSelect.Width
    picSelect.Visible = True
    Me.Enabled = True
    MousePointer = 0
End Sub

Private Sub cmdCut_Click()
    picSelect.Picture = picBlank.Image
    SelHolder.Picture = LoadPicture()
    SelHolder.Height = SelectHeight
    SelHolder.Width = SelectWidth
    picSelect.Height = SelectHeight
    picSelect.Width = SelectWidth
    StretchBlt SelHolder.hdc, 0, 0, SelectWidth, SelectHeight, picMerge.hdc, SelectLeft, SelectTop, SelectWidth, SelectHeight, SRCCOPY
    Clipboard.SetData SelHolder.Image
    StretchBlt picMerge.hdc, SelectLeft, SelectTop, SelectWidth, SelectHeight, picSelect.hdc, 0, 0, SelectWidth, SelectHeight, SRCCOPY
    picMerge.Refresh
    Image1.Picture = picMerge.Image
    picSelect.Visible = False
    SelectShape.Visible = False
End Sub


