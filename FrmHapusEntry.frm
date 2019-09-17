VERSION 5.00
Begin VB.Form FrmHapusEntry 
   BackColor       =   &H0000FF00&
   BorderStyle     =   0  'None
   ClientHeight    =   3255
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4665
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form5"
   Picture         =   "FrmHapusEntry.frx":0000
   ScaleHeight     =   3255
   ScaleWidth      =   4665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   100
      Picture         =   "FrmHapusEntry.frx":0CA4
      ScaleHeight     =   375
      ScaleWidth      =   1965
      TabIndex        =   7
      Top             =   0
      Width           =   1970
   End
   Begin VB.TextBox TxtRegname 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Text            =   "Text2"
      Top             =   3360
      Width           =   4095
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   4320
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label LblUninstall 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Uninstall"
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   240
      MouseIcon       =   "FrmHapusEntry.frx":1484
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   2720
      Width           =   1335
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   375
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label LblHapus 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Remove"
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   2160
      MouseIcon       =   "FrmHapusEntry.frx":178E
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   2720
      Width           =   1095
   End
   Begin VB.Label LblBatal 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cancel"
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   3360
      MouseIcon       =   "FrmHapusEntry.frx":1A98
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   2720
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   375
      Left            =   2160
      Shape           =   4  'Rounded Rectangle
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   375
      Left            =   3360
      Shape           =   4  'Rounded Rectangle
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label LblDname 
      BackStyle       =   0  'Transparent
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   840
      Width           =   3975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   $"FrmHapusEntry.frx":1DA2
      Height          =   855
      Left            =   240
      TabIndex        =   2
      Top             =   1680
      Width           =   4215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Are you sure to remove list entry on registry ?"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   4215
   End
End
Attribute VB_Name = "FrmHapusEntry"
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

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormDrag Me
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.LblUninstall.FontBold = False
Me.LblHapus.FontBold = False
Me.LblBatal.FontBold = False
End Sub

Private Sub LblBatal_Click()
Unload Me
End Sub

Private Sub LblBatal_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.LblUninstall.FontBold = False
Me.LblHapus.FontBold = False
Me.LblBatal.FontBold = True
End Sub

Private Sub LblHapus_Click()
Call DeleteKey(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Uninstall\" + TxtRegname.Text)
FrmMain.New_Refresh
Unload Me
End Sub

Private Sub LblHapus_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.LblUninstall.FontBold = False
Me.LblHapus.FontBold = True
Me.LblBatal.FontBold = False
End Sub

Private Sub LblUninstall_Click()
Unload Me
FrmMain.Show_FormUninstall

End Sub

Private Sub LblUninstall_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.LblUninstall.FontBold = True
Me.LblHapus.FontBold = False
Me.LblBatal.FontBold = False
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormDrag Me
End Sub
