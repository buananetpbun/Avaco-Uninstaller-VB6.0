VERSION 5.00
Begin VB.Form frmInformasi 
   AutoRedraw      =   -1  'True
   BackColor       =   &H0000FF00&
   BorderStyle     =   0  'None
   ClientHeight    =   3915
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5925
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frminfo.frx":0000
   ScaleHeight     =   3915
   ScaleWidth      =   5925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   80
      Picture         =   "frminfo.frx":0F13
      ScaleHeight     =   375
      ScaleWidth      =   3030
      TabIndex        =   18
      Top             =   0
      Width           =   3030
   End
   Begin VB.TextBox TxtRName 
      Appearance      =   0  'Flat
      BackColor       =   &H0038BCFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   480
      Width           =   3975
   End
   Begin VB.TextBox TxtContact 
      Appearance      =   0  'Flat
      BackColor       =   &H0038BCFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   3000
      Width           =   4095
   End
   Begin VB.TextBox TxtUIAbout 
      Appearance      =   0  'Flat
      BackColor       =   &H0038BCFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   2640
      Width           =   3495
   End
   Begin VB.TextBox TxtHLink 
      Appearance      =   0  'Flat
      BackColor       =   &H0038BCFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   2280
      Width           =   3495
   End
   Begin VB.TextBox TxtDVersion 
      Appearance      =   0  'Flat
      BackColor       =   &H0038BCFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   1920
      Width           =   3975
   End
   Begin VB.TextBox TxtPublisher 
      Appearance      =   0  'Flat
      BackColor       =   &H0038BCFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   1560
      Width           =   3975
   End
   Begin VB.TextBox TxtUString 
      Appearance      =   0  'Flat
      BackColor       =   &H0038BCFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   1200
      Width           =   3975
   End
   Begin VB.TextBox TxtDName 
      Appearance      =   0  'Flat
      BackColor       =   &H0038BCFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   840
      Width           =   3975
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   5280
      MouseIcon       =   "frminfo.frx":19BB
      MousePointer    =   99  'Custom
      Picture         =   "frminfo.frx":1CC5
      Top             =   2560
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   5280
      MouseIcon       =   "frminfo.frx":258F
      MousePointer    =   99  'Custom
      Picture         =   "frminfo.frx":2899
      Top             =   2160
      Width           =   480
   End
   Begin VB.Image ImgDetailInfo 
      Height          =   480
      Left            =   240
      MouseIcon       =   "frminfo.frx":3163
      Picture         =   "frminfo.frx":346D
      ToolTipText     =   "Detail Informasi"
      Top             =   3240
      Width           =   480
   End
   Begin VB.Label LblDes 
      BackStyle       =   0  'Transparent
      Caption         =   "This Information found on registry window"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   840
      TabIndex        =   17
      Top             =   3480
      Width           =   3255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Registry Name :"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   240
      TabIndex        =   16
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Contact :"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "URL Info About :"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Help Link :"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label LblOK 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "OK"
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   4560
      MouseIcon       =   "frminfo.frx":3D37
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   3435
      Width           =   1095
   End
   Begin VB.Shape ShapeOK 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   375
      Left            =   4560
      Shape           =   4  'Rounded Rectangle
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Display Version :"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label lblprogname 
      BackStyle       =   0  'Transparent
      Caption         =   "Publisher :"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Uninstall String :"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Display Name :"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   1335
   End
End
Attribute VB_Name = "frmInformasi"
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

Private Declare Function ShellExecute Lib _
   "shell32.dll" Alias "ShellExecuteA" _
   (ByVal hWnd As Long, _
   ByVal lpOperation As String, _
   ByVal lpFile As String, _
   ByVal lpParameters As String, _
   ByVal lpDirectory As String, _
   ByVal nShowCmd As Long) As Long
    
Private Const SW_SHOWNORMAL = 1
Private Sub Form_Load()
FrmMain.GetInformasi
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormDrag Me
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.LblOK.FontBold = False
End Sub

Private Sub Image1_Click()
On Error GoTo perbaiki
If TxtHLink.Text = "?" Then
MsgBox "Information for Help Link Not found", vbInformation, "Detail Informasi"
Else
 ShellExecute Me.hWnd, _
        vbNullString, TxtHLink, _
        vbNullString, _
        "c:\", _
        SW_SHOWNORMAL
End If
Exit Sub
perbaiki:
End Sub

Private Sub Image2_Click()
On Error GoTo perbaiki
If TxtUIAbout.Text = "?" Then
MsgBox "Information for Url Info About not found", vbInformation, "Detail Informasi"
Else
ShellExecute Me.hWnd, _
        vbNullString, TxtUIAbout, _
        vbNullString, _
        "c:\", _
        SW_SHOWNORMAL
End If
Exit Sub
perbaiki:
End Sub

Private Sub LblOK_Click()
Unload Me
End Sub

Private Sub LblOK_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.LblOK.FontBold = True
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormDrag Me
End Sub
