VERSION 5.00
Begin VB.Form FrmEntryBaru 
   BackColor       =   &H0000FF00&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3255
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4665
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   Picture         =   "FrmEntryBaru.frx":0000
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
      Picture         =   "FrmEntryBaru.frx":0C2F
      ScaleHeight     =   375
      ScaleWidth      =   1935
      TabIndex        =   8
      Top             =   0
      Width           =   1935
   End
   Begin VB.TextBox TxtUninstallString 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   240
      TabIndex        =   2
      Text            =   "Command Uninstall"
      Top             =   2040
      Width           =   4215
   End
   Begin VB.TextBox TxtDiplayName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   240
      TabIndex        =   1
      Text            =   "Program Name"
      Top             =   1440
      Width           =   4215
   End
   Begin VB.TextBox TxtRegName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Text            =   "Key Registry Name"
      Top             =   840
      Width           =   4215
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   4440
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Label LblOK 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "OK"
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   2160
      MouseIcon       =   "FrmEntryBaru.frx":1380
      MousePointer    =   99  'Custom
      TabIndex        =   7
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
      MouseIcon       =   "FrmEntryBaru.frx":168A
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   2720
      Width           =   1095
   End
   Begin VB.Shape ShapeOK 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   375
      Left            =   2160
      Shape           =   4  'Rounded Rectangle
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Shape ShapeBatal 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   375
      Left            =   3360
      Shape           =   4  'Rounded Rectangle
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label LblRegName 
      BackStyle       =   0  'Transparent
      Caption         =   "Registry Name (key) :"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   600
      Width           =   3015
   End
   Begin VB.Label LblUninstallString 
      BackStyle       =   0  'Transparent
      Caption         =   "Uninstall String :"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label LblDiplayName 
      BackStyle       =   0  'Transparent
      Caption         =   "Display Name :"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      Width           =   1335
   End
End
Attribute VB_Name = "FrmEntryBaru"
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

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.LblOK.FontBold = False
Me.LblBatal.FontBold = False
End Sub

Private Sub LblBatal_Click()
Unload Me
End Sub

Private Sub LblBatal_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.LblOK.FontBold = False
Me.LblBatal.FontBold = True
End Sub

Private Sub LblOK_Click()
If TxtRegname.Text = "" Then
MsgBox "Please Don't blank to Key registry Name !", vbCritical, "Error !!"
Exit Sub
End If
If TxtDiplayName.Text = "" Then
MsgBox "Please Don't blank to Program Name !", vbCritical, "Error !!"
Exit Sub
End If
If TxtUninstallString.Text = "" Then
MsgBox "Please Don't blank to Uninstall Command ", vbCritical, "Error !!"
Exit Sub
End If

Call SaveString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Uninstall\" + TxtRegname.Text, "DisplayName", TxtDiplayName.Text)
Call SaveString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Uninstall\" + TxtRegname.Text, "UninstallString", TxtUninstallString.Text)
FrmMain.New_Refresh
Unload Me
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormDrag Me
End Sub


Private Sub LblOK_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.LblOK.FontBold = True
Me.LblBatal.FontBold = False
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormDrag Me
End Sub
