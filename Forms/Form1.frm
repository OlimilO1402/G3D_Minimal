VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Example for minimal requirements on 3d graphics"
   ClientHeight    =   8400
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12135
   LinkTopic       =   "Form1"
   ScaleHeight     =   8400
   ScaleWidth      =   12135
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton BtnCameraDwn 
      Caption         =   "Camera v dwn"
      Height          =   375
      Left            =   9840
      TabIndex        =   9
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton BtnCameraUp 
      Caption         =   "Camera ^ up"
      Height          =   375
      Left            =   8640
      TabIndex        =   8
      Top             =   120
      Width           =   1215
   End
   Begin VB.ComboBox CmbObj3D 
      Height          =   315
      Left            =   120
      TabIndex        =   7
      Text            =   "Combo1"
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton BtnRotSpeedDown 
      Caption         =   "RotSpeed -"
      Height          =   375
      Left            =   6480
      TabIndex        =   5
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton BtnRotSpeedUp 
      Caption         =   "RotSpeed +"
      Height          =   375
      Left            =   5280
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Rotate <="
      Height          =   255
      Left            =   4080
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   0
      Top             =   480
   End
   Begin VB.CommandButton BtnZoomOut 
      Caption         =   "Zoom Out"
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton BtnZoomIn 
      Caption         =   "Zoom In"
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   7695
      Left            =   120
      ScaleHeight     =   7635
      ScaleWidth      =   11835
      TabIndex        =   0
      Top             =   600
      Width           =   11895
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   195
      Left            =   11160
      TabIndex        =   10
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   195
      Left            =   7920
      TabIndex        =   6
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Obj3D As Object3D
Dim c As Long
Dim rotangle As Double
Dim rotspeed As Double
Dim Projection As Matrix34
Dim Center As Point3
Dim alpha_x As Double

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Picture1_KeyDown KeyCode, Shift
End Sub

Private Sub Form_Load()
    CmbObj3D.AddItem "Würfel"
    CmbObj3D.AddItem "Ikosaeder"
    CmbObj3D.ListIndex = 0
    M3D.Pi = 4 * Atn(1)
    M3D.Pi2 = 4 * Atn(1) * 2
    alpha_x = Pi / 2
    
    rotspeed = 1
    Timer1.Interval = 10
    Timer1.Enabled = False
    BtnZoomIn.Caption = "Zoom In"
    BtnZoomOut.Caption = "Zoom Out"
    Picture1.ScaleMode = vbPixels
    Picture1.AutoRedraw = True
    UpdateLbl
    Draw
End Sub

Private Sub CmbObj3D_Click()
    Select Case CmbObj3D.ListIndex
    Case 0: Obj3D = CreateCube
    Case 1: Obj3D = CreateIcosahedron
    End Select
    Draw
End Sub

Private Sub Check1_Click()
    Timer1.Enabled = Check1.Value = vbChecked
    Draw
End Sub

Private Sub UpdateLbl()
    Label1.Caption = "rs: " & Round(rotspeed, 2)
    Label2.Caption = "alp: " & Round(alpha_x * 180 / Pi, 2)
End Sub
Private Sub BtnZoomIn_Click()
    M3D.ZoomIn
    Draw
End Sub
Private Sub BtnZoomOut_Click()
    M3D.ZoomOut
    Draw
End Sub

Private Sub BtnRotSpeedUp_Click()
    Picture1_KeyDown vbKeyRight, 0
End Sub
Private Sub BtnRotSpeedDown_Click()
    Picture1_KeyDown vbKeyLeft, 0
End Sub
Private Sub BtnCameraUp_Click()
    Picture1_KeyDown vbKeyUp, 0
End Sub
Private Sub BtnCameraDwn_Click()
    Picture1_KeyDown vbKeyDown, 0
End Sub

Private Sub Form_Resize()
    Dim brdr: brdr = 8 * Screen.TwipsPerPixelX
    Dim l: l = brdr
    Dim T: T = Picture1.Top
    Dim W: W = (Me.ScaleWidth - 2 * l)
    Dim H: H = Me.ScaleHeight - T - brdr
    If W > 0 And H > 0 Then
        Picture1.Move l, T, W, H
        InitProjection Picture1
        Draw
    End If
End Sub

Private Sub Picture1_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyUp:    alpha_x = alpha_x - 5 * Pi / 180: If alpha_x < 0 Then alpha_x = alpha_x + 2 * Pi
    Case vbKeyDown:  alpha_x = alpha_x + 5 * Pi / 180: If alpha_x > Pi Then alpha_x = alpha_x - 2 * Pi
    Case vbKeyLeft:  rotspeed = rotspeed - IIf(-0.9 <= rotspeed And rotspeed <= 1, 0.1, 1)
    Case vbKeyRight: rotspeed = rotspeed + IIf(-1 <= rotspeed And rotspeed <= 0.9, 0.1, 1)
    End Select
    InitProjection Picture1
    Draw
    UpdateLbl
End Sub

Private Sub InitProjection(aPB As PictureBox)
    Dim tx As Double: tx = aPB.ScaleWidth / 2
    Dim tz As Double: tz = aPB.ScaleHeight / 2
    Dim ax As Double
    If M3D.sc = 0 Then M3D.sc = 37
    ax = alpha_x
    Projection = New_Matrix34Camera(300, 60 / M3D.sc, 60 / M3D.sc, tx, tz, 0, 0, -5, ax, 0, 0)
End Sub

Private Sub Draw()
    Picture1.Cls
    Dim i As Long
    Dim start As Long: start = 0
    Dim color As Long: color = vbBlue 'colors(c)
    
    InitProjection Picture1
    M3D.DrawPoint3_projected Picture1, Projection, Center, color
    Dim rotobj As Object3D
    For i = start To c
        rotobj = Obj3D
        rotobj.points = M3D.Rotate_xy(rotobj.points, Center.X, Center.Y, rotangle * Pi / 180)
        'color = colors(i)
        M3D.DrawObj3D_projected Picture1, Projection, rotobj, color
    Next
End Sub

Private Sub Timer1_Timer()
    If Check1.Value = vbChecked Then
        rotangle = rotangle + rotspeed
        Draw
    End If
End Sub
