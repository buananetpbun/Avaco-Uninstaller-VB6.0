VERSION 5.00
Begin VB.Form FrmPilihLaporan 
   Appearance      =   0  'Flat
   BackColor       =   &H0000FF00&
   BorderStyle     =   0  'None
   ClientHeight    =   2055
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4035
   LinkTopic       =   "Form4"
   Picture         =   "FrmPilihLaporan.frx":0000
   ScaleHeight     =   2055
   ScaleWidth      =   4035
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   390
      Left            =   80
      Picture         =   "FrmPilihLaporan.frx":0B7A
      ScaleHeight     =   390
      ScaleWidth      =   2205
      TabIndex        =   6
      Top             =   0
      Width           =   2205
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H0038BCFF&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   200
      ScaleHeight     =   825
      ScaleWidth      =   3615
      TabIndex        =   1
      Top             =   480
      Width           =   3640
      Begin VB.OptionButton OptDName 
         Appearance      =   0  'Flat
         BackColor       =   &H0038BCFF&
         Caption         =   "Just View "" Display Name """
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   80
         TabIndex        =   4
         Top             =   40
         Width           =   3375
      End
      Begin VB.OptionButton OptDnameUString 
         Appearance      =   0  'Flat
         BackColor       =   &H0038BCFF&
         Caption         =   "Display Name and Uninstall String "
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   80
         TabIndex        =   3
         Top             =   280
         Width           =   3615
      End
      Begin VB.OptionButton OptDetailInformasi 
         Appearance      =   0  'Flat
         BackColor       =   &H0038BCFF&
         Caption         =   "Report Information Details"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   80
         TabIndex        =   2
         Top             =   540
         Width           =   3015
      End
   End
   Begin VB.Image Image8 
      Height          =   480
      Left            =   240
      Picture         =   "FrmPilihLaporan.frx":13E6
      Top             =   1440
      Width           =   480
   End
   Begin VB.Label LblBatal 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cancel"
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   2760
      MouseIcon       =   "FrmPilihLaporan.frx":1CB0
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   1530
      Width           =   1095
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   375
      Left            =   2760
      Shape           =   4  'Rounded Rectangle
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label LblOK 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "OK"
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   1560
      MouseIcon       =   "FrmPilihLaporan.frx":1FBA
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   1530
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   375
      Left            =   1560
      Shape           =   4  'Rounded Rectangle
      Top             =   1440
      Width           =   1095
   End
End
Attribute VB_Name = "FrmPilihLaporan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'--> avaco Uninstaller 2002
'--> version 1.00
'--> Version Language : English
'--> By Agus Ramadhani
'--> avaco software
'--> http://avaco-software.tripod.com
'--> avaco@9cy.Com
'--> 2002-2003
'--> Don't forget to Vote :)

Private Sub Form_Load()
OptDName.Value = True
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
LblOK.FontBold = False
LblBatal.FontBold = False
End Sub

Private Sub LblBatal_Click()
Unload Me
End Sub

Private Sub LblBatal_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
LblOK.FontBold = False
LblBatal.FontBold = True
End Sub

Private Sub LblOK_Click()
On Error Resume Next
If OptDName.Value = True Then
Kill App.Path & "\temp" & ".tmp"
Call ModMain.SaveFile1(FrmMain)
FrmLaporan.Show vbModal, FrmMain
End If

If OptDnameUString.Value = True Then
Kill App.Path & "\temp" & ".tmp"
Call ModMain.SaveFile2(FrmMain)
FrmLaporan.Show vbModal, FrmMain
End If

If OptDetailInformasi.Value = True Then
Kill App.Path & "\temp" & ".tmp"
Call ModMain.SaveFile3(FrmMain)
FrmLaporan.Show vbModal, FrmMain
End If
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormDrag Me
End Sub

Private Sub LblOK_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
LblOK.FontBold = True
LblBatal.FontBold = False
End Sub

Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormDrag Me
End Sub
