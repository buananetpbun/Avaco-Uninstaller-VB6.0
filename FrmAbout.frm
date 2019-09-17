VERSION 5.00
Begin VB.Form FrmAbout 
   BackColor       =   &H0000FF00&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3420
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   5280
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmAbout.frx":0000
   ScaleHeight     =   3420
   ScaleWidth      =   5280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "FREEWARE !!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Label lblWebSiteSaya 
      BackStyle       =   0  'Transparent
      Caption         =   "Http://Avaco-Software.Tripod.com"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   2760
      MouseIcon       =   "FrmAbout.frx":7955
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   3120
      Width           =   2535
   End
   Begin VB.Label lblEmailSaya 
      BackStyle       =   0  'Transparent
      Caption         =   "e-mail : Avaco_Soft22@Yahoo.com"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   180
      Left            =   2760
      MouseIcon       =   "FrmAbout.frx":7C5F
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   2925
      Width           =   2415
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   2160
      Picture         =   "FrmAbout.frx":7F69
      Top             =   2880
      Width           =   480
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "If you find bugs on my program, please give me report and send from my email. Thanks."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   2520
      Width           =   4575
   End
End
Attribute VB_Name = "FrmAbout"
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

Private Sub lblEmailSaya_Click()
   ShellExecute Me.hWnd, _
        vbNullString, _
        "mailto:Avaco_Soft22@yahoo.com", _
        vbNullString, _
        "c:\", _
        SW_SHOWNORMAL
End Sub

Private Sub lblWebSiteSaya_Click()
    ShellExecute Me.hWnd, _
        vbNullString, _
        "http://Avaco-Software.tripod.com", _
        vbNullString, _
        "c:\", _
        SW_SHOWNORMAL
End Sub
Private Sub Form_Click()
Unload Me
End Sub

Private Sub Label1_Click()
Unload Me
End Sub
