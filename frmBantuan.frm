VERSION 5.00
Begin VB.Form frmBantuan 
   Caption         =   "Help"
   ClientHeight    =   7155
   ClientLeft      =   2370
   ClientTop       =   2400
   ClientWidth     =   7425
   Icon            =   "frmBantuan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7155
   ScaleWidth      =   7425
   StartUpPosition =   1  'CenterOwner
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton CmdExit 
      Caption         =   "OK"
      Height          =   375
      Left            =   6240
      TabIndex        =   2
      Top             =   6750
      Width           =   1095
   End
   Begin VB.ListBox List2 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   4260
      ItemData        =   "frmBantuan.frx":08CA
      Left            =   0
      List            =   "frmBantuan.frx":0A2D
      TabIndex        =   1
      Top             =   2400
      Width           =   7425
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2370
      ItemData        =   "frmBantuan.frx":2461
      Left            =   0
      List            =   "frmBantuan.frx":249E
      TabIndex        =   0
      Top             =   0
      Width           =   7425
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sorry this help for indonesian Version :)"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   600
      TabIndex        =   3
      Top             =   6840
      Width           =   2775
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   375
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Image ImBantu 
      Height          =   480
      Left            =   0
      MouseIcon       =   "frmBantuan.frx":278D
      Picture         =   "frmBantuan.frx":2A97
      ToolTipText     =   "Bantuan"
      Top             =   6700
      Width           =   480
   End
End
Attribute VB_Name = "frmBantuan"
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

DefLng A-W
DefSng X-Z

Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub List1_Click()
Dim i, context$, res&
i = List1.ListIndex
context$ = List1.List(i) & Chr$(0)
If context$ <> "" Then
   res& = SendMessageLong(List2.hWnd, LB_FINDSTRINGEXACT, -1&, ByVal context$)
   List2.ListIndex = res
   If List2.ListIndex > 0 Then
      List2.TopIndex = List2.ListIndex - 1
   End If
End If

End Sub

