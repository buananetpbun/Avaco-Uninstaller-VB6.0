VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmBackupRegistry 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H0000FF00&
   BorderStyle     =   0  'None
   Caption         =   "Form3"
   ClientHeight    =   5640
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8205
   ClipControls    =   0   'False
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmRegbackup.frx":0000
   ScaleHeight     =   5640
   ScaleWidth      =   8205
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox CheckBSUS 
      Appearance      =   0  'Flat
      BackColor       =   &H0038BCFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   360
      MaskColor       =   &H00E0E0E0&
      TabIndex        =   12
      Top             =   4140
      Width           =   225
   End
   Begin VB.FileListBox File1 
      Height          =   1650
      Left            =   240
      Pattern         =   "*.Reg"
      TabIndex        =   4
      Top             =   2280
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox TxtSimpan 
      Appearance      =   0  'Flat
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
      Height          =   315
      Left            =   3360
      TabIndex        =   3
      ToolTipText     =   "Write file name for save"
      Top             =   4110
      Width           =   2745
   End
   Begin VB.TextBox TxtBackup 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   3255
      Left            =   2400
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Top             =   720
      Width           =   5655
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      ForeColor       =   &H80000008&
      Height          =   280
      Left            =   150
      ScaleHeight     =   255
      ScaleWidth      =   2205
      TabIndex        =   0
      Top             =   435
      Width           =   2235
      Begin VB.CheckBox chkBakup 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         Caption         =   "<< See list Backup"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   80
         TabIndex        =   1
         Top             =   0
         Width           =   2400
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   120
      Top             =   5520
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   1575
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRegbackup.frx":1221
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRegbackup.frx":153D
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRegbackup.frx":1859
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3375
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRegbackup.frx":2135
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRegbackup.frx":2A11
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView TreeViewBackup 
      Height          =   3255
      Left            =   150
      TabIndex        =   5
      ToolTipText     =   "Please double click for Export backup to Registry"
      Top             =   720
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   5741
      _Version        =   393217
      Indentation     =   0
      LabelEdit       =   1
      Style           =   7
      ImageList       =   "ImageList2"
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label LblBSUS 
      BackStyle       =   0  'Transparent
      Caption         =   "Backup All Uninstall (Key)"
      Enabled         =   0   'False
      Height          =   255
      Left            =   620
      TabIndex        =   13
      Top             =   4170
      Width           =   2295
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H0038BCFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   375
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   4080
      Width           =   2655
   End
   Begin VB.Label LblOK 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Save"
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   6120
      MouseIcon       =   "FrmRegbackup.frx":32ED
      MousePointer    =   99  'Custom
      TabIndex        =   11
      Top             =   4160
      Width           =   855
   End
   Begin VB.Label LblBatal 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cancel"
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   7080
      MouseIcon       =   "FrmRegbackup.frx":35F7
      MousePointer    =   99  'Custom
      TabIndex        =   10
      Top             =   4155
      Width           =   855
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Example >> ; Write new information "
      Height          =   255
      Left            =   360
      TabIndex        =   9
      Top             =   5160
      Width           =   6015
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   $"FrmRegbackup.frx":3901
      Height          =   615
      Left            =   360
      TabIndex        =   8
      Top             =   4680
      Width           =   6015
   End
   Begin VB.Shape Shape3 
      Height          =   915
      Left            =   165
      Top             =   4560
      Width           =   7860
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   7320
      Picture         =   "FrmRegbackup.frx":39AF
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   480
   End
   Begin VB.Label LblInformasi 
      BackStyle       =   0  'Transparent
      Caption         =   ">> New Backup :"
      Height          =   255
      Left            =   2400
      TabIndex        =   7
      Top             =   480
      Width           =   5655
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "File :"
      Height          =   255
      Left            =   3000
      TabIndex        =   6
      Top             =   4140
      Width           =   480
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   375
      Left            =   7095
      Shape           =   4  'Rounded Rectangle
      Top             =   4080
      Width           =   855
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   375
      Left            =   6120
      Shape           =   4  'Rounded Rectangle
      Top             =   4080
      Width           =   855
   End
End
Attribute VB_Name = "FrmBackupRegistry"
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

Const SW_SHOWNORMAL = 1
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub Nodes()
Dim s As Node
Dim i As Integer
File1.Refresh
TreeViewBackup.Nodes.Clear
Set s = TreeViewBackup.Nodes.Add(, , "r", "File Reg Backup", 3)
For i = 0 To File1.ListCount - 1
    Set s = TreeViewBackup.Nodes.Add("r", tvwChild, , File1.List(i), 2, 1)
Next i
TreeViewBackup.Nodes(1).Expanded = True
End Sub

Sub Hapus()
On Error GoTo agus
Dim result As Integer
result = MsgBox("Are you sure remove Backup Registry : " & TreeViewBackup.SelectedItem.Text & "?", vbInformation + vbYesNo, "Delete")
If result = vbYes Then
On Error GoTo agus
Kill App.Path & "\Backup Registry\" & TreeViewBackup.SelectedItem.Text
Nodes
End If
Exit Sub
agus:
MsgBox "File for remove not found !!", vbOKOnly + vbCritical, "Error !!"
End Sub

Private Sub CheckBSUS_Click()
If CheckBSUS.Value = vbChecked Then
On Error Resume Next
Dim fName As String
fName = App.Path & "\" & "temp" & ".tmp"
SaveKey "HKEY_LOCAL_MACHINE" & "\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall", fName
Call LoadText
Call LoadText
Kill App.Path & "\temp" & ".tmp"
LblInformasi.Caption = ">> Backup for : All Uninstall (Key) on registry"
TreeViewBackup.Nodes.Clear
chkBakup.Caption = "<< See list Backup"
chkBakup.Value = vbUnchecked
LblOK.Enabled = False
TxtSimpan.Text = ""
LblBSUS.Enabled = True
Else
chkBakup.Caption = "<< See list Backup"
chkBakup.Value = vbUnchecked
TreeViewBackup.Nodes.Clear
LblOK.Enabled = True
LblBSUS.Enabled = False
FrmMain.Backup_Registry
Call LoadText
Call LoadText
Kill App.Path & "\temp" & ".tmp"
End If
End Sub

Private Sub chkBakup_Click()
On Error Resume Next
If chkBakup.Value = vbChecked Then
Nodes
LblInformasi.Caption = ">> Please double click for Export backup to Registry"
TxtBackup.Text = ""
LblOK.Enabled = False
chkBakup.Caption = "Back >>"
TxtSimpan.Text = ""
CheckBSUS.Enabled = False
CheckBSUS.Value = vbUnchecked
Else
CheckBSUS.Value = vbUnchecked
CheckBSUS.Enabled = True
chkBakup.Caption = "<< See list Backup"
TreeViewBackup.Nodes.Clear
TxtBackup.Text = ""
LblOK.Enabled = True
FrmMain.Backup_Registry
Call LoadText
Call LoadText
Kill App.Path & "\temp" & ".tmp"
End If
End Sub

Private Sub Form_Load()
On Error Resume Next
MkDir App.Path + "\Backup Registry\"
File1.Path = App.Path & "\Backup Registry\"
Call LoadText
Call LoadText
Kill App.Path & "\temp" & ".tmp"
TxtSimpan.SetFocus
End Sub

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
Dim fName As String
Dim num, ad As String
Dim retval, result As String
num = LCase(TxtSimpan.Text)

fName = App.Path & "\Backup Registry\" & num & ".reg"
ad = LCase(TxtSimpan.Text) & ".reg"
If TxtBackup.Text = "" Then
MsgBox "List on text not found !!", vbCritical, "error !!"
Exit Sub
End If

retval = Dir$(App.Path & "\Backup Registry\" & num & ".reg")

If retval = ad Then
result = MsgBox("File with name [ " & retval & " ] already exist." + vbCrLf & _
"Would you like to replace the existing file ?", vbInformation + vbYesNo, "Informatin !!")
If result = vbYes Then
fName = App.Path & "\Backup Registry\" & num & ".reg"
Close #1
Open fName For Output As #1
Print #1, TxtBackup.Text
Close #1
LblOK.Enabled = False
TxtBackup.Text = ""
TxtSimpan.Text = ""
TreeViewBackup.Refresh

End If
Exit Sub
End If

Close #1
Open fName For Output As #1
Print #1, TxtBackup.Text
Close #1
LblOK.Enabled = False
TxtBackup.Text = ""
TxtSimpan.Text = ""
TreeViewBackup.Refresh

End Sub

Private Sub LblOK_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.LblOK.FontBold = True
Me.LblBatal.FontBold = False
End Sub

Private Sub Timer1_Timer()
If TxtSimpan.Text = "" Then
LblOK.Enabled = False
Else
LblOK.Enabled = True
End If
End Sub


Sub kembalikan_registry()
If Not TreeViewBackup.SelectedItem Is Nothing Then
ShellExecute 0, vbNullString, App.Path & "\Backup Registry\" & TreeViewBackup.SelectedItem.Text & ".", vbNullString, "", SW_SHOWNORMAL
Else: MsgBox "Nothing to open!", vbExclamation, App.Title
End If
End Sub


Sub LoadText()
Dim fName As String
Dim readln
Dim textload
On Error GoTo ld_err
fName = App.Path & "\temp" & ".tmp"
TxtBackup.Text = ""
Open App.Path & "\temp" & ".tmp" For Input As #1
        Do While Not EOF(1)
            Line Input #1, readln
           textload = textload + readln + Chr$(13) + Chr$(10)
If Len(textload) >= 500000 Then
      MsgBox "This File {" & App.Path & "\temp" & ".tmp" & "} too large to Open", vbCritical, "Error !!"
      GoTo ld_end
    End If
  Loop
TxtBackup = textload
ld_err:
    Resume ld_end
ld_end:
    On Error Resume Next
    Close #1
End Sub

Sub panggil_Nodes()
Nodes
End Sub

Private Sub TreeViewBackup_DblClick()
Call kembalikan_registry
End Sub

Private Sub TreeViewBackup_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then PopupMenu FrmPopup.Menu2
End Sub

Private Sub TreeViewBackup_NodeClick(ByVal Node As MSComctlLib.Node)
On Error GoTo ld_err
Dim fName As String
fName = App.Path & "\Backup Registry\" & Node.Text & "."
Dim readln
Dim wow
TxtBackup.Text = ""
 Open fName For Input As #1
        Do While Not EOF(1)
            Line Input #1, readln
            wow = wow + readln + Chr$(13) + Chr$(10)
   If Len(wow) >= 30000 Then
      MsgBox "This file {" & App.Path & "\temp" & ".tmp" & "} too large to Open", vbCritical, "Error !!"
      GoTo ld_end
    End If
  Loop
TxtBackup = wow
ld_err:
 Resume ld_end
ld_end:
On Error Resume Next
Close #1
End Sub
