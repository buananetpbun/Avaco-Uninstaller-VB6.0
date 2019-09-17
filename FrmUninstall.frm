VERSION 5.00
Begin VB.Form frmUninstall 
   AutoRedraw      =   -1  'True
   BackColor       =   &H0000FF00&
   BorderStyle     =   0  'None
   ClientHeight    =   2055
   ClientLeft      =   -45
   ClientTop       =   -330
   ClientWidth     =   4035
   ClipControls    =   0   'False
   Icon            =   "FrmUninstall.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmUninstall.frx":000C
   ScaleHeight     =   2055
   ScaleWidth      =   4035
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox TxtRegname 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Text            =   "Text2"
      Top             =   2280
      Width           =   3375
   End
   Begin VB.TextBox TxtDname 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1080
      Width           =   3615
   End
   Begin VB.Label LblBersihkan 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Clear"
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   240
      MouseIcon       =   "FrmUninstall.frx":0A63
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   1515
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label LblInfomasi 
      BackStyle       =   0  'Transparent
      Caption         =   $"FrmUninstall.frx":0D6D
      Height          =   855
      Left            =   240
      TabIndex        =   4
      Top             =   600
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   375
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   1440
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label LblBatal 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cancel"
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   2760
      MouseIcon       =   "FrmUninstall.frx":0DFC
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   1515
      Width           =   1095
   End
   Begin VB.Label LblUninstall 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Uninstall"
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   1560
      MouseIcon       =   "FrmUninstall.frx":1106
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   1515
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
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "FrmUninstall.frx":1410
      Top             =   520
      Width           =   480
   End
   Begin VB.Label LblInfomasi2 
      BackStyle       =   0  'Transparent
      Caption         =   "Are you sure to uninstall this program ?"
      Height          =   375
      Left            =   885
      TabIndex        =   0
      Top             =   720
      Width           =   3135
   End
End
Attribute VB_Name = "frmUninstall"
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
LblBersihkan.FontBold = False
LblUninstall.FontBold = False
LblBatal.FontBold = False
End Sub

Private Sub LblBatal_Click()
Unload Me
End Sub

Private Sub LblBatal_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
LblBersihkan.FontBold = False
LblUninstall.FontBold = False
LblBatal.FontBold = True
End Sub

Private Sub LblBersihkan_Click()
Call DeleteKey(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Uninstall\" + TxtRegname.Text)
Unload Me
FrmMain.New_Refresh
End Sub

Private Sub LblBersihkan_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
LblBersihkan.FontBold = True
LblUninstall.FontBold = False
LblBatal.FontBold = False
End Sub

Private Sub LblUninstall_Click()
LblInfomasi2.Visible = False
TxtDname.Visible = False
Image1.Visible = False
Shape1.Visible = False
LblUninstall.Visible = False
LblInfomasi.Visible = True
LblBersihkan.Visible = True
Shape3.Visible = True
FrmMain.Get_Uninstall
End Sub

Private Sub LblUninstall_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
LblBersihkan.FontBold = False
LblUninstall.FontBold = True
LblBatal.FontBold = False
End Sub
