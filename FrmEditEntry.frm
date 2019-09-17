VERSION 5.00
Begin VB.Form FrmEditEntry 
   Appearance      =   0  'Flat
   BackColor       =   &H0000FF00&
   BorderStyle     =   0  'None
   Caption         =   "Form6"
   ClientHeight    =   3255
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4665
   LinkTopic       =   "Form6"
   Picture         =   "FrmEditEntry.frx":0000
   ScaleHeight     =   3255
   ScaleWidth      =   4665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   480
      Top             =   2640
   End
   Begin VB.TextBox TxtDname 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   360
      TabIndex        =   1
      Top             =   1560
      Width           =   3975
   End
   Begin VB.TextBox TxtUString 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   360
      TabIndex        =   0
      Top             =   2160
      Width           =   3975
   End
   Begin VB.Line Line1 
      X1              =   360
      X2              =   4320
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label LblRegName 
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
      Left            =   360
      TabIndex        =   7
      Top             =   720
      Width           =   3975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Display Name :"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Uninstall String :"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Edit For Registry Name (key) :"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   480
      Width           =   3015
   End
   Begin VB.Label LblBatal 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cancel"
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   3240
      MouseIcon       =   "FrmEditEntry.frx":0BC9
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   2720
      Width           =   1095
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   375
      Left            =   3240
      Shape           =   4  'Rounded Rectangle
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label LblOK 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "OK"
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   2040
      MouseIcon       =   "FrmEditEntry.frx":0ED3
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   2720
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   375
      Left            =   2040
      Shape           =   4  'Rounded Rectangle
      Top             =   2640
      Width           =   1095
   End
End
Attribute VB_Name = "FrmEditEntry"
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
Call SaveString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Uninstall\" + LblRegName.Caption, "DisplayName", TxtDname.Text)
Call SaveString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Uninstall\" + LblRegName.Caption, "UninstallString", TxtUString.Text)
 FrmMain.New_Refresh
 Unload Me
End Sub

Private Sub LblOK_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.LblOK.FontBold = True
Me.LblBatal.FontBold = False
End Sub

Private Sub Timer1_Timer()
If TxtUString.Text = "" Then
TxtUString.Text = "?"
End If
End Sub
