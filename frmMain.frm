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
   Begin VB.CommandButton Command3 
      Caption         =   "add poly"
      Height          =   495
      Left            =   11040
      TabIndex        =   4
      Top             =   2880
      Width           =   855
   End
   Begin VB.ComboBox cmbDrawMode 
      Height          =   315
      Left            =   11040
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "add circle"
      Height          =   495
      Left            =   11040
      TabIndex        =   2
      Top             =   1200
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
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Const Density As Double = 1

Private Sub cmbDrawMode_Click()
    RenderMode = cmbDrawMode.ListIndex
End Sub

Private Sub Command1_Click()


    Dim I   As Long

    NofBodies = 0
    NJ = 0


    For I = 1 To 20
        CREATECircle Vec2(I * 55, 50), 5 + Rnd * (20), Density
    Next

'AddJoint 2, 3, 50
'AddJoint 4, 5, 50
'AddJoint 6, 7, 50
'AddJoint 8, 9, 50
'AddJoint 10, 11, 50



    For I = 20 + 1 To 20 + 9

        CREATECircle Vec2((I - 20 - 1) * 75, PicH + 40), 65, Density

        BodySetStatic I
    Next

AddJoint 20 + 4, 5, 200

    ' AddJoint 1, 2, 50

    MAINLOOP




End Sub

Private Sub Command2_Click()
    CREATECircle Vec2(PicW * 0.5, 0), 5 + Rnd * 20, Density
End Sub

Private Sub Command3_Click()


    CREATERandomPoly Vec2(PicW \ 2, 0), Density

End Sub

Private Sub Form_Load()

    PIC.Height = 360
    PIC.Width = Int(PIC.Height * 16 / 9)



    pHDC = PIC.hDC
    PicW = PIC.Width
    PicH = PIC.Height

    InitMATH
    InitMaterials


    cmbDrawMode.AddItem "API"
    cmbDrawMode.AddItem "Antialias"
    cmbDrawMode.ListIndex = 0

    Randomize Timer
    InitRC

End Sub

Private Sub Form_Unload(Cancel As Integer)
    UnloadRC

    Erase Contacts
    Erase Body


    End

End Sub

