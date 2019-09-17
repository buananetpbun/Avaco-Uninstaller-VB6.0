VERSION 5.00
Begin VB.Form FrmOption 
   BackColor       =   &H0000FF00&
   BorderStyle     =   0  'None
   ClientHeight    =   3915
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5925
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   Picture         =   "FrmOptions.frx":0000
   ScaleHeight     =   3915
   ScaleWidth      =   5925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   390
      Left            =   80
      Picture         =   "FrmOptions.frx":0BB8
      ScaleHeight     =   390
      ScaleWidth      =   1605
      TabIndex        =   19
      Top             =   0
      Width           =   1605
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H0038BCFF&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   240
      ScaleHeight     =   705
      ScaleWidth      =   5385
      TabIndex        =   12
      Top             =   1920
      Width           =   5415
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H0038BCFF&
         Caption         =   "StartUp Windows"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   4
         Left            =   1920
         TabIndex        =   18
         Top             =   360
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H0038BCFF&
         Caption         =   "Start Menu"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   15
         Top             =   60
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H0038BCFF&
         Caption         =   "Desktop"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H0038BCFF&
         Caption         =   "Quick Launch"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   1920
         TabIndex        =   13
         Top             =   60
         Width           =   1575
      End
      Begin VB.Label LblOKShortCut 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Apply"
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   4440
         MouseIcon       =   "FrmOptions.frx":129D
         MousePointer    =   99  'Custom
         TabIndex        =   17
         Top             =   330
         Width           =   855
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H000080FF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         Height          =   375
         Left            =   4440
         Shape           =   4  'Rounded Rectangle
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H0038BCFF&
      ForeColor       =   &H80000008&
      Height          =   280
      Left            =   240
      ScaleHeight     =   255
      ScaleWidth      =   5385
      TabIndex        =   10
      Top             =   2760
      Width           =   5415
      Begin VB.CheckBox ChkUninstall 
         Appearance      =   0  'Flat
         BackColor       =   &H0038BCFF&
         Caption         =   "Show Uninstall Dialog !"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   0
         Value           =   1  'Checked
         Width           =   2400
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H0038BCFF&
      ForeColor       =   &H80000008&
      Height          =   280
      Left            =   240
      ScaleHeight     =   255
      ScaleWidth      =   5385
      TabIndex        =   6
      Top             =   1320
      Width           =   5415
      Begin VB.CheckBox chkUrut 
         Appearance      =   0  'Flat
         BackColor       =   &H0038BCFF&
         Caption         =   "Ascending \ Descending"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   0
         Value           =   1  'Checked
         Width           =   2400
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H0038BCFF&
      ForeColor       =   &H80000008&
      Height          =   280
      Left            =   240
      ScaleHeight     =   255
      ScaleWidth      =   5385
      TabIndex        =   2
      Top             =   720
      Width           =   5415
      Begin VB.OptionButton OptIconBesar 
         Appearance      =   0  'Flat
         BackColor       =   &H0038BCFF&
         Caption         =   "Large Icon"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   3840
         TabIndex        =   9
         Top             =   30
         Width           =   1575
      End
      Begin VB.OptionButton OptList 
         Appearance      =   0  'Flat
         BackColor       =   &H0038BCFF&
         Caption         =   "List"
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   0
         Left            =   1320
         TabIndex        =   5
         Top             =   30
         Width           =   735
      End
      Begin VB.OptionButton OptDetail 
         Appearance      =   0  'Flat
         BackColor       =   &H0038BCFF&
         Caption         =   "Details"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   120
         TabIndex        =   4
         Top             =   30
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton OptIconKecil 
         Appearance      =   0  'Flat
         BackColor       =   &H0038BCFF&
         Caption         =   "Small Icon"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   2400
         TabIndex        =   3
         Top             =   30
         Width           =   1095
      End
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Create Shortcut on :"
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   1680
      Width           =   2415
   End
   Begin VB.Label LblOK 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "OK"
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   4560
      MouseIcon       =   "FrmOptions.frx":15A7
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   3375
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   375
      Left            =   4560
      Shape           =   4  'Rounded Rectangle
      Top             =   3285
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "FrmOptions.frx":18B1
      Top             =   3240
      Width           =   480
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Sort List View as :"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Show List View as :"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   3255
   End
End
Attribute VB_Name = "FrmOption"
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

Private Declare Function OSfCreateShellLink Lib "vb6stkit.dll" Alias "fCreateShellLink" (ByVal lpstrFolderName As String, ByVal lpstrLinkName As String, ByVal lpstrLinkPath As String, ByVal lpstrLinkArguments As String, ByVal fPrivate As Long, ByVal sParent As String) As Long

Private Sub chkUrut_Click()
   If chkUrut.Value = vbChecked Then
      FrmMain.lstview.SortOrder = lvwAscending
    Else
      FrmMain.lstview.SortOrder = lvwDescending
    End If
    FrmMain.lstview.Refresh
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormDrag Me
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
LblOK.FontBold = False
LblOKShortCut.FontBold = False
End Sub

Private Sub LblOK_Click()
Me.Hide
End Sub

Private Sub LblOK_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
LblOK.FontBold = True
End Sub

Private Sub LblOKShortCut_Click()
Dim lReturn As Long
Select Case True
  Case Option1(1).Value
    lReturn = OSfCreateShellLink("Programs" & vbNullChar, "AVACO Uninstaller 2002", (App.Path & "\" & App.EXEName), "" & vbNullChar, True, "$(Start Menu)")
  Case Option1(2).Value
    lReturn = OSfCreateShellLink("..\..\Desktop" & vbNullChar, "AVACO Uninstaller 2002", (App.Path & "\" & App.EXEName), "" & vbNullChar, True, "$(Programs)")
  Case Option1(3).Value
    lReturn = OSfCreateShellLink("..\..\Application Data\Microsoft\Internet Explorer\Quick Launch" & vbNullChar, "AVACO Uninstaller 2002", (App.Path & "\" & App.EXEName), "" & vbNullChar, True, "$(Programs)")
    Case Option1(4).Value
    lReturn = OSfCreateShellLink("StartUp" & vbNullChar, "AVACO Uninstaller 2002", (App.Path & "\" & App.EXEName), "" & vbNullChar, True, "$(Programs)")
End Select
If lReturn = 0 Then
  MsgBox "Error for Create Shortcut!"
Else
  MsgBox "Success for Create Shortcut!"
End If
End Sub

Private Sub LblOKShortCut_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
LblOKShortCut.FontBold = True
End Sub

Private Sub OptDetail_Click()
FrmMain.lstview.View = lvwReport
End Sub

Private Sub OptIconBesar_Click()
FrmMain.lstview.View = lvwIcon
End Sub

Private Sub OptIconKecil_Click()
FrmMain.lstview.View = lvwSmallIcon
End Sub

Private Sub OptList_Click(Index As Integer)
FrmMain.lstview.View = lvwList
End Sub

Private Sub Picture3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormDrag Me
End Sub

Private Sub Picture5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
LblOKShortCut.FontBold = False
End Sub
