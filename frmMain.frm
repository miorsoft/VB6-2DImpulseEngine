VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Physic Engine"
   ClientHeight    =   7965
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12930
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   531
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   862
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "ADD Regular Poly"
      Height          =   615
      Left            =   11040
      TabIndex        =   7
      Top             =   4320
      Width           =   975
   End
   Begin VB.CheckBox chkJPG 
      Caption         =   "Save Jpg Frames"
      Height          =   495
      Left            =   11160
      TabIndex        =   6
      Top             =   5520
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "ADD Brick"
      Height          =   615
      Left            =   11040
      TabIndex        =   4
      Top             =   3480
      Width           =   975
   End
   Begin VB.ComboBox cmbScene 
      Height          =   315
      Left            =   11040
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1680
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ADD circle"
      Height          =   615
      Left            =   11040
      TabIndex        =   2
      Top             =   2640
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "(RE) START"
      Height          =   615
      Left            =   11040
      TabIndex        =   1
      Top             =   240
      Width           =   975
   End
   Begin VB.PictureBox PIC 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   6135
      Left            =   120
      ScaleHeight     =   6135
      ScaleWidth      =   9615
      TabIndex        =   0
      Top             =   120
      Width           =   9615
   End
   Begin VB.Label Label1 
      Caption         =   "SCENE"
      Height          =   255
      Left            =   11040
      TabIndex        =   5
      Top             =   1440
      Width           =   1095
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit





Private Sub chkJPG_Click()
    SaveFrames = (chkJPG.Value = vbChecked)
End Sub

Private Sub cmbScene_Change()
    CreateScene frmMain.cmbScene.ListIndex
End Sub

Private Sub cmbScene_Click()
    CreateScene frmMain.cmbScene.ListIndex
End Sub

Private Sub Command1_Click()


    CreateScene frmMain.cmbScene.ListIndex




End Sub

Private Sub Command2_Click()
    CREATECircle Vec2(PicW * 0.5, 0), 5 + Rnd * 20, DefDensity
End Sub

Private Sub Command3_Click()


'    CREATERandomPoly Vec2(PicW \ 2, 0), DefDensity
    CreateBox Vec2(PicW \ 2, 0), 60, 30

End Sub

Private Sub Command4_Click()
    CreateRegularPoly Vec2(PicW \ 2, 0), 7 + Rnd * 30, 7 + Rnd * 30, 3 + Int(Rnd * 10), -Int(Rnd * 2), DefDensity
End Sub

Private Sub Form_Activate()


    MAINLOOP
End Sub

Private Sub Form_Load()


    If Dir(App.Path & "\Frames", vbDirectory) = vbNullString Then MkDir App.Path & "\Frames"
    If Dir(App.Path & "\Frames\*.*") <> vbNullString Then Kill App.Path & "\Frames\*.*"

    PIC.Height = 360
    PIC.Width = Int(PIC.Height * 16 / 9)



    pHDC = PIC.hDC
    PicW = PIC.Width
    PicH = PIC.Height

    InitMATH
    InitMaterials


    cmbScene.AddItem "First"
    cmbScene.AddItem "Distance Joints"
    cmbScene.AddItem "1 Pin Joints"
    cmbScene.AddItem "2 Pins Joints"
    cmbScene.AddItem "2 Pins Joints II"
    cmbScene.AddItem "Slope"
    cmbScene.AddItem "Gum Bridge"
    


    cmbScene.ListIndex = 0

    Randomize Timer
    InitRC


    CreateScene 0

    Version = App.Major & "." & App.Minor & "." & App.Revision
    CreateIntroFrames


End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    CreateOuttroFrames

End Sub

Private Sub Form_Unload(Cancel As Integer)
    UnloadRC

    Erase Contacts
    Erase Body


    End

End Sub

