VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H0038BCFF&
   BorderStyle     =   0  'None
   Caption         =   "Avaco Uninstaller 2002"
   ClientHeight    =   5790
   ClientLeft      =   90
   ClientTop       =   90
   ClientWidth     =   8445
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00FFFFFF&
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   8445
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   720
      Top             =   5040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":08CA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   120
      Top             =   5040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":11A6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   1320
      Top             =   5160
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1800
      Top             =   5160
   End
   Begin MSComctlLib.ListView lstview 
      Height          =   2985
      Left            =   120
      TabIndex        =   2
      ToolTipText     =   "Select program on the list and then double click list view for run Uninstall program"
      Top             =   1560
      Width           =   7140
      _ExtentX        =   12594
      _ExtentY        =   5265
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      PictureAlignment=   5
      _Version        =   393217
      Icons           =   "ImageList2"
      SmallIcons      =   "ImageList1"
      ForeColor       =   0
      BackColor       =   16777215
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
      MouseIcon       =   "FrmMain.frx":1A86
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Display Name"
         Object.Width           =   5291
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Uninstall String"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Registry Name"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Publisher"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Display Version"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Help Link"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "URL Info About "
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Contact"
         Object.Width           =   4304
      EndProperty
   End
   Begin VB.Label LblMinimized 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   6600
      MouseIcon       =   "FrmMain.frx":1DA0
      MousePointer    =   99  'Custom
      TabIndex        =   19
      ToolTipText     =   "Restore"
      Top             =   0
      Width           =   135
   End
   Begin VB.Label LblRestore 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   6720
      MouseIcon       =   "FrmMain.frx":20AA
      MousePointer    =   99  'Custom
      TabIndex        =   18
      ToolTipText     =   "Minimize"
      Top             =   0
      Width           =   255
   End
   Begin VB.Label LblMaximized 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   6960
      MouseIcon       =   "FrmMain.frx":23B4
      MousePointer    =   99  'Custom
      TabIndex        =   17
      ToolTipText     =   "Maximize"
      Top             =   0
      Width           =   135
   End
   Begin VB.Label LblTanggalDanJam 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   4800
      TabIndex        =   16
      Top             =   4605
      Width           =   2295
   End
   Begin VB.Label Lblbantu 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Help"
      Height          =   255
      Left            =   3180
      MouseIcon       =   "FrmMain.frx":26BE
      TabIndex        =   15
      Top             =   465
      Width           =   495
   End
   Begin VB.Label LblTools 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Tools"
      Height          =   255
      Left            =   2580
      MouseIcon       =   "FrmMain.frx":29C8
      TabIndex        =   14
      Top             =   465
      Width           =   495
   End
   Begin VB.Label LblUninstall 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Uninstall"
      Height          =   255
      Left            =   1785
      MouseIcon       =   "FrmMain.frx":2CD2
      TabIndex        =   13
      Top             =   465
      Width           =   735
   End
   Begin VB.Label LblFile 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "File"
      Height          =   255
      Left            =   240
      MouseIcon       =   "FrmMain.frx":2FDC
      TabIndex        =   12
      Top             =   465
      Width           =   375
   End
   Begin VB.Label LblTampilan 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "View"
      Height          =   255
      Left            =   1215
      MouseIcon       =   "FrmMain.frx":32E6
      TabIndex        =   11
      Top             =   465
      Width           =   480
   End
   Begin VB.Label LblEdit 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Edit"
      Height          =   255
      Left            =   720
      MouseIcon       =   "FrmMain.frx":35F0
      TabIndex        =   10
      Top             =   465
      Width           =   375
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Info Details"
      Height          =   255
      Index           =   1
      Left            =   1100
      TabIndex        =   9
      Top             =   1305
      Width           =   1095
   End
   Begin VB.Image ImgDetailInfo 
      Height          =   480
      Left            =   1365
      MouseIcon       =   "FrmMain.frx":38FA
      MousePointer    =   99  'Custom
      Picture         =   "FrmMain.frx":3C04
      ToolTipText     =   "Information Detail"
      Top             =   795
      Width           =   480
   End
   Begin VB.Image ImBantu 
      Height          =   480
      Left            =   5235
      MouseIcon       =   "FrmMain.frx":44CE
      MousePointer    =   99  'Custom
      Picture         =   "FrmMain.frx":47D8
      ToolTipText     =   "Help"
      Top             =   795
      Width           =   480
   End
   Begin VB.Image ImgUninstall 
      Height          =   480
      Left            =   390
      MouseIcon       =   "FrmMain.frx":50A2
      MousePointer    =   99  'Custom
      Picture         =   "FrmMain.frx":53AC
      ToolTipText     =   "Uninstall"
      Top             =   795
      Width           =   480
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Uninstall"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   8
      Top             =   1300
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Help"
      Height          =   255
      Index           =   5
      Left            =   5235
      TabIndex        =   7
      Top             =   1300
      Width           =   495
   End
   Begin VB.Image ImgSetings 
      Height          =   480
      Left            =   4320
      MouseIcon       =   "FrmMain.frx":5C76
      MousePointer    =   99  'Custom
      Picture         =   "FrmMain.frx":5F80
      ToolTipText     =   "Options"
      Top             =   795
      Width           =   480
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Options"
      Height          =   255
      Index           =   4
      Left            =   4200
      TabIndex        =   6
      Top             =   1300
      Width           =   735
   End
   Begin VB.Image ImgKeluar 
      Height          =   480
      Left            =   6120
      MouseIcon       =   "FrmMain.frx":684A
      MousePointer    =   99  'Custom
      Picture         =   "FrmMain.frx":6B54
      ToolTipText     =   "Exit"
      Top             =   795
      Width           =   480
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Exit"
      Height          =   255
      Index           =   6
      Left            =   6040
      TabIndex        =   5
      Top             =   1300
      Width           =   615
   End
   Begin VB.Image ImgLaporan 
      Height          =   480
      Left            =   2400
      MouseIcon       =   "FrmMain.frx":741E
      MousePointer    =   99  'Custom
      Picture         =   "FrmMain.frx":7728
      ToolTipText     =   "Report Information Registry"
      Top             =   820
      Width           =   480
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Report"
      Height          =   255
      Index           =   2
      Left            =   2280
      TabIndex        =   4
      Top             =   1300
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   3360
      MouseIcon       =   "FrmMain.frx":7FF2
      MousePointer    =   99  'Custom
      Picture         =   "FrmMain.frx":82FC
      ToolTipText     =   "Backup Registry"
      Top             =   795
      Width           =   480
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Backup"
      Height          =   255
      Index           =   3
      Left            =   3240
      TabIndex        =   3
      Top             =   1300
      Width           =   735
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   8160
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Image ImgAbout 
      Appearance      =   0  'Flat
      Height          =   885
      Left            =   8280
      MouseIcon       =   "FrmMain.frx":8BC6
      MousePointer    =   99  'Custom
      Picture         =   "FrmMain.frx":8ED0
      ToolTipText     =   "About AVACO"
      Top             =   480
      Width           =   870
   End
   Begin VB.Image RS 
      Height          =   165
      Left            =   7200
      Picture         =   "FrmMain.frx":B7A2
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   270
   End
   Begin VB.Image UR 
      Height          =   465
      Left            =   5760
      Picture         =   "FrmMain.frx":BA4C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1365
   End
   Begin VB.Image picFormResize 
      Height          =   90
      Left            =   2280
      MousePointer    =   8  'Size NW SE
      Picture         =   "FrmMain.frx":C057
      Top             =   5520
      Width           =   90
   End
   Begin VB.Image ImgForm 
      Height          =   510
      Left            =   1320
      Picture         =   "FrmMain.frx":C0CB
      Top             =   0
      Width           =   3810
   End
   Begin VB.Label LblStatus 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   820
      TabIndex        =   1
      Top             =   4605
      Width           =   2295
   End
   Begin VB.Label LblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Found :"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   4605
      Width           =   975
   End
   Begin VB.Image UL 
      Height          =   465
      Left            =   0
      Picture         =   "FrmMain.frx":CC06
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1365
   End
   Begin VB.Image LS 
      Height          =   165
      Left            =   6720
      Picture         =   "FrmMain.frx":EDB4
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   210
   End
   Begin VB.Image BL 
      Height          =   240
      Left            =   7320
      Picture         =   "FrmMain.frx":EFDA
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   300
   End
   Begin VB.Image BM 
      Height          =   150
      Left            =   6960
      Picture         =   "FrmMain.frx":F3DC
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   135
   End
   Begin VB.Image BR 
      Height          =   285
      Left            =   7080
      Picture         =   "FrmMain.frx":F536
      Stretch         =   -1  'True
      Top             =   840
      Width           =   315
   End
   Begin VB.Image UM 
      Height          =   510
      Left            =   6960
      Picture         =   "FrmMain.frx":FA38
      Stretch         =   -1  'True
      Top             =   840
      Width           =   165
   End
   Begin VB.Image bak 
      Height          =   1155
      Left            =   30
      Picture         =   "FrmMain.frx":FF42
      Stretch         =   -1  'True
      Top             =   480
      Width           =   2505
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'--> Avaco Uninstaller 2002
'--> version 1.00
'--> Version Language : English
'--> By Agus Ramadhani
'--> avaco software
'--> http://avaco-software.tripod.com
'--> avaco@9cy.Com
'--> 2002-2003
'--> Don't forget to Vote :)

Public FormFlag As Boolean, FX As Long, FY As Long
Public FormFirst As Boolean, AX As Long, AY As Long
Private Declare Function WinExec Lib "kernel32" (ByVal lpCmdLine As String, ByVal nCmdShow As Long) As Long
Dim GetlocString, RegName, UString, Dname, Publisher, DVersion, HelpLink, UIAbout, Contact As String
Dim iKetetapan As Integer
Dim fTimer

Sub GetInformasi()

RegName = lstview.SelectedItem.Key
Dname = GetString(HKEY_LOCAL_MACHINE, GetlocString & lstview.SelectedItem.Key, "DisplayName")
UString = GetString(HKEY_LOCAL_MACHINE, GetlocString & lstview.SelectedItem.Key, "UninstallString")
Publisher = Trim(GetString(HKEY_LOCAL_MACHINE, GetlocString & lstview.SelectedItem.Key, "Publisher"))
DVersion = Trim(GetString(HKEY_LOCAL_MACHINE, GetlocString & lstview.SelectedItem.Key, "DisplayVersion"))
HelpLink = Trim(GetString(HKEY_LOCAL_MACHINE, GetlocString & lstview.SelectedItem.Key, "HelpLink"))
UIAbout = Trim(GetString(HKEY_LOCAL_MACHINE, GetlocString & lstview.SelectedItem.Key, "URLInfoAbout"))
Contact = Trim(GetString(HKEY_LOCAL_MACHINE, GetlocString & lstview.SelectedItem.Key, "Contact"))

If Len(RegName) = 0 Then
frmInformasi.TxtRName = "?"
Else
frmInformasi.TxtRName = RegName
End If

If Len(Dname) = 0 Then
frmInformasi.TxtDname.Text = "?"
Else
frmInformasi.TxtDname.Text = Dname
End If

If Len(UString) = 0 Then
frmInformasi.TxtUString.Text = "?"
Else
frmInformasi.TxtUString.Text = UString
End If

If Len(Publisher) = 0 Then
frmInformasi.TxtPublisher.Text = "?"
Else
frmInformasi.TxtPublisher.Text = Publisher
End If

If Len(DVersion) = 0 Then
frmInformasi.TxtDVersion.Text = "?"
Else
frmInformasi.TxtDVersion.Text = DVersion
End If

If Len(HelpLink) = 0 Then
frmInformasi.TxtHLink.Text = "?"
Else
frmInformasi.TxtHLink.Text = HelpLink
End If

If Len(UIAbout) = 0 Then
frmInformasi.TxtUIAbout.Text = "?"
Else
frmInformasi.TxtUIAbout.Text = UIAbout
End If

If Len(Contact) = 0 Then
frmInformasi.TxtContact.Text = "?"
Else
frmInformasi.TxtContact.Text = Contact
End If
    
End Sub

Private Sub GetKetReg()
GetlocString = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\"
ModRegistry.GetKeyNames HKEY_LOCAL_MACHINE, GetlocString
End Sub

Private Sub ShowUninstallList()
On Error Resume Next
Dim LokasiItem As ListItem
Call GetKetReg
For iKetetapan = 1 To sKeys.Count - 0
    Dname = GetString(HKEY_LOCAL_MACHINE, GetlocString & sKeys(iKetetapan), "DisplayName")
    UString = GetString(HKEY_LOCAL_MACHINE, GetlocString & sKeys(iKetetapan), "UninstallString")
    Publisher = GetString(HKEY_LOCAL_MACHINE, GetlocString & sKeys(iKetetapan), "Publisher")
    DVersion = GetString(HKEY_LOCAL_MACHINE, GetlocString & sKeys(iKetetapan), "DisplayVersion")
    HelpLink = GetString(HKEY_LOCAL_MACHINE, GetlocString & sKeys(iKetetapan), "HelpLink")
    UIAbout = GetString(HKEY_LOCAL_MACHINE, GetlocString & sKeys(iKetetapan), "URLInfoAbout")
    Contact = GetString(HKEY_LOCAL_MACHINE, GetlocString & sKeys(iKetetapan), "Contact")

    If Len(Dname) > 0 Then
        Set LokasiItem = lstview.ListItems.Add(, sKeys(iKetetapan), Dname, 1, 1)
        
        If Len(UString) = 0 Then
           LokasiItem.SubItems(1) = "?"
        Else
           LokasiItem.SubItems(1) = UString
        End If
            
        If Len(sKeys(iKetetapan)) = 0 Then
           LokasiItem.SubItems(2) = "?"
        Else
           LokasiItem.SubItems(2) = sKeys(iKetetapan)
        End If
           
        If Len(Publisher) = 0 Then
           LokasiItem.SubItems(3) = "?"
        Else
           LokasiItem.SubItems(3) = Publisher
        End If
           
        If Len(DVersion) = 0 Then
           LokasiItem.SubItems(4) = "?"
        Else
           LokasiItem.SubItems(4) = DVersion
        End If
        
        If Len(HelpLink) = 0 Then
           LokasiItem.SubItems(5) = "?"
        Else
           LokasiItem.SubItems(5) = HelpLink
        End If
        
        If Len(UIAbout) = 0 Then
           LokasiItem.SubItems(6) = "?"
        Else
           LokasiItem.SubItems(6) = UIAbout
        End If
        
        If Len(Contact) = 0 Then
           LokasiItem.SubItems(7) = "?"
        Else
           LokasiItem.SubItems(7) = Contact
        End If
End If
    LblStatus.Caption = lstview.ListItems.Count & " Installed Programs"
Next iKetetapan
    Set sKeys = Nothing
   
End Sub

Sub Show_FormUninstall()
If FrmOption.ChkUninstall.Value = vbChecked Then
Dname = GetString(HKEY_LOCAL_MACHINE, GetlocString & lstview.SelectedItem.Key, "DisplayName")
RegName = lstview.SelectedItem.Key
frmUninstall.TxtDname.Text = Dname
frmUninstall.TxtRegname.Text = RegName
frmUninstall.Show vbModal, FrmMain
Else
Get_Uninstall
End If
End Sub

Sub Get_Uninstall()
Dim strRemove As String
strRemove = GetString(HKEY_LOCAL_MACHINE, GetlocString & lstview.SelectedItem.Key, "UninstallString")
WinExec strRemove, 1
End Sub

Private Sub Bak_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Tulisan_Menu_tipis
End Sub

Private Sub Form_Load()
On Error Resume Next
    Set sKeys = New Collection
    lstview.Refresh
    ShowUninstallList
    LblTanggalDanJam.Caption = Now
    Me.Icon = ImgUninstall.Picture
    Form_Resize
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Tulisan_Menu_tipis
End Sub

Sub Tulisan_Menu_tipis()
LblFile.FontBold = False
LblEdit.FontBold = False
LblTampilan.FontBold = False
LblUninstall.FontBold = False
LblTools.FontBold = False
Lblbantu.FontBold = False
Label2(0).FontBold = False
Label2(1).FontBold = False
Label2(2).FontBold = False
Label2(3).FontBold = False
Label2(4).FontBold = False
Label2(5).FontBold = False
Label2(6).FontBold = False
End Sub
Private Sub Form_Resize()
On Error Resume Next
Call SkinForm
lstview.Width = Me.Width - 280 + 45
lstview.Height = Me.Height - 1960
ImgAbout.Left = Me.Width - 1050
LblInfo.Top = Me.Height - 355
LblStatus.Top = Me.Height - 355
LblTanggalDanJam.Top = Me.Height - 355
LblTanggalDanJam.Left = Me.Width - 2540
LblMinimized.Left = Me.Width - 540
LblRestore.Left = Me.Width - 430
LblMaximized.Left = Me.Width - 200
picFormResize.Top = Me.Height - 100
picFormResize.Left = Me.Width - 100
lstview.ColumnHeaders(1).Width = 5500
Line1.X2 = Me.Width - 850

End Sub

Private Sub Form_Unload(Cancel As Integer)
    GetlocString = ""
    Dim Form As Form
    For Each Form In Forms
        If Form.Name <> Me.Name Then
            Unload Form
            Set Form = Nothing
        End If
    Next Form
End Sub

Sub SkinForm()
UL.Top = 0
UL.Left = 0
BL.Top = Me.Height - BL.Height - 0
BL.Left = 0
LS.Top = UL.Height
LS.Height = Me.Height - UL.Height - BL.Height
LS.Left = 0
UM.Width = Me.Width - UL.Width - UR.Width
UM.Left = UL.Width
UM.Top = 0
UR.Top = 0
UR.Left = Me.Width - UR.Width - 0
BR.Left = Me.Width - BR.Width - 0
BR.Top = Me.Height - BR.Height - 0
RS.Left = Me.Width - RS.Width - 0
RS.Height = Me.Height - UR.Height - BR.Height
RS.Top = UR.Height
BM.Top = Me.Height - BM.Height - 0
BM.Left = BL.Width
BM.Width = Me.Width - BL.Width - BR.Width
End Sub

Private Sub Image1_Click()
Call Backup_Registry
FrmBackupRegistry.Show vbModal, FrmMain
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2(0).FontBold = False
Label2(1).FontBold = False
Label2(2).FontBold = False
Label2(3).FontBold = True
Label2(4).FontBold = False
Label2(5).FontBold = False
Label2(6).FontBold = False
End Sub

Private Sub ImBantu_Click()
frmBantuan.Show
End Sub

Private Sub ImBantu_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2(0).FontBold = False
Label2(1).FontBold = False
Label2(2).FontBold = False
Label2(3).FontBold = False
Label2(4).FontBold = False
Label2(5).FontBold = True
Label2(6).FontBold = False
End Sub

Private Sub ImgAbout_Click()
FrmAbout.Show vbModal, FrmMain
End Sub

Private Sub ImgDetailInfo_Click()
frmInformasi.Show vbModal, FrmMain
End Sub

Private Sub ImgDetailInfo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2(0).FontBold = False
Label2(1).FontBold = True
Label2(2).FontBold = False
Label2(3).FontBold = False
Label2(4).FontBold = False
Label2(5).FontBold = False
Label2(6).FontBold = False
End Sub

Private Sub ImgForm_DblClick()
UM_DblClick
End Sub

Private Sub ImgForm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Tulisan_Menu_tipis
End Sub

Private Sub ImgKeluar_Click()
Unload Me
End Sub

Private Sub ImgKeluar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2(0).FontBold = False
Label2(1).FontBold = False
Label2(2).FontBold = False
Label2(3).FontBold = False
Label2(4).FontBold = False
Label2(5).FontBold = False
Label2(6).FontBold = True
End Sub

Private Sub ImgLaporan_Click()
FrmPilihLaporan.Show vbModal, FrmMain
End Sub

Private Sub ImgLaporan_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2(0).FontBold = False
Label2(1).FontBold = False
Label2(2).FontBold = True
Label2(3).FontBold = False
Label2(4).FontBold = False
Label2(5).FontBold = False
Label2(6).FontBold = False

End Sub

Private Sub ImgSetings_Click()
FrmOption.Show vbModal, FrmMain
End Sub

Private Sub ImgSetings_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2(0).FontBold = False
Label2(1).FontBold = False
Label2(2).FontBold = False
Label2(3).FontBold = False
Label2(4).FontBold = True
Label2(5).FontBold = False
Label2(6).FontBold = False
End Sub

Private Sub ImgUninstall_Click()
Call Show_FormUninstall
End Sub

Sub Backup_Registry()
On Error Resume Next
Dim fName As String
fName = App.Path & "\" & "temp" & ".tmp"
SaveKey "HKEY_LOCAL_MACHINE" & "\" & GetlocString & lstview.SelectedItem.Key, fName
Dname = GetString(HKEY_LOCAL_MACHINE, GetlocString & lstview.SelectedItem.Key, "DisplayName")
FrmBackupRegistry.LblInformasi = ">> Backup For : " & Dname
End Sub

Private Sub ImgUninstall_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2(0).FontBold = True
Label2(1).FontBold = False
Label2(2).FontBold = False
Label2(3).FontBold = False
Label2(4).FontBold = False
Label2(5).FontBold = False
Label2(6).FontBold = False

End Sub

Private Sub Lblbantu_Click()
PopupMenu FrmPopup.Menu8
End Sub

Private Sub Lblbantu_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
LblFile.FontBold = False
LblEdit.FontBold = False
LblTampilan.FontBold = False
LblUninstall.FontBold = False
LblTools.FontBold = False
Me.Lblbantu.FontBold = True
End Sub

Private Sub LblEdit_Click()
PopupMenu FrmPopup.Menu6
End Sub

Private Sub LblEdit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
LblFile.FontBold = False
LblEdit.FontBold = True
Me.Lblbantu.FontBold = False
LblTampilan.FontBold = False
LblUninstall.FontBold = False
LblTools.FontBold = False
End Sub

Private Sub LblFile_Click()
PopupMenu FrmPopup.Menu3
End Sub

Private Sub LblFile_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
LblFile.FontBold = True
LblEdit.FontBold = False
Me.Lblbantu.FontBold = False
LblTampilan.FontBold = False
LblUninstall.FontBold = False
LblTools.FontBold = False
End Sub

Private Sub LblMaximized_Click()
Me.WindowState = vbMaximized
End Sub

Private Sub LblMinimized_Click()
Me.WindowState = vbNormal
End Sub

Private Sub LblRestore_Click()

Me.WindowState = vbMinimized
End Sub

Sub MnuHapusEntry_Click()
Dname = GetString(HKEY_LOCAL_MACHINE, GetlocString & lstview.SelectedItem.Key, "DisplayName")
RegName = lstview.SelectedItem.Key
FrmHapusEntry.LblDname.Caption = Dname
FrmHapusEntry.TxtRegname.Text = RegName
FrmHapusEntry.Show vbModal, FrmMain
End Sub

Sub MnuEditEntry_Click()
Dname = GetString(HKEY_LOCAL_MACHINE, GetlocString & lstview.SelectedItem.Key, "DisplayName")
UString = GetString(HKEY_LOCAL_MACHINE, GetlocString & lstview.SelectedItem.Key, "UninstallString")
RegName = lstview.SelectedItem.Key
FrmEditEntry.LblRegName.Caption = RegName
FrmEditEntry.TxtDname.Text = Dname
FrmEditEntry.TxtUString.Text = UString
FrmEditEntry.Show vbModal, FrmMain
End Sub

Private Sub LblTampilan_Click()
PopupMenu FrmPopup.Menu7
End Sub

Private Sub LblTampilan_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
LblFile.FontBold = False
LblEdit.FontBold = False
LblTampilan.FontBold = True
LblUninstall.FontBold = False
LblTools.FontBold = False
Me.Lblbantu.FontBold = False
End Sub

Private Sub LblTools_Click()
PopupMenu FrmPopup.Menu4
End Sub

Private Sub LblTools_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
LblFile.FontBold = False
LblEdit.FontBold = False
LblTampilan.FontBold = False
LblUninstall.FontBold = False
LblTools.FontBold = True
Me.Lblbantu.FontBold = False
End Sub

Private Sub LblUninstall_Click()
PopupMenu FrmPopup.Menu5
End Sub

Private Sub LblUninstall_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
LblFile.FontBold = False
LblEdit.FontBold = False
LblTampilan.FontBold = False
LblUninstall.FontBold = True
LblTools.FontBold = False
Me.Lblbantu.FontBold = False
End Sub

Private Sub lstview_DblClick()
Call Show_FormUninstall
End Sub

Private Sub lstview_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then PopupMenu FrmPopup.Menu
End Sub

Private Sub lstview_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Tulisan_Menu_tipis
End Sub

Private Sub Timer1_Timer()
LblTanggalDanJam.Caption = Now
End Sub

Private Sub MnuInformation_Click()
frmInformasi.Show vbModal, FrmMain
End Sub

Sub New_Refresh()
lstview.ListItems.Clear
Set sKeys = New Collection
ShowUninstallList
lstview.Refresh
lstview.ColumnHeaders(1).Width = 5500
End Sub

Private Sub UL_DblClick()
UM_DblClick
End Sub

Private Sub UM_DblClick()
If Me.WindowState = vbMinimized Then Exit Sub
If Me.WindowState = vbMaximized Then
Me.WindowState = vbNormal
Else
Me.WindowState = vbMaximized
End If
End Sub

Private Sub UM_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormDrag Me
End Sub

Private Sub UR_DblClick()
UM_DblClick
End Sub

Private Sub UR_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormDrag Me
End Sub

Private Sub ImgForm_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormDrag Me
End Sub

Private Sub picFormResize_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
       If FormFlag = False Then
        FX = picFormResize.Left + (picFormResize.Width / 2)
        FY = picFormResize.Top + (picFormResize.Height / 2)
        AX = X
        AY = Y
        FormFlag = True
        FormFirst = True
        fTimer = Timer + 0.1
    End If
End Sub

Private Sub PicFormResize_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   On Error Resume Next
    DoEvents
    If FormFlag = True Then
        If Timer < fTimer Then Exit Sub:
        fTimer = Timer + 0.2
        Dim posW As Long, posH As Long
        posW = FX + X - IIf(FormFirst = True, X, AX)
        posH = FY + Y - IIf(FormFirst = True, Y, AY)
        If posW < 3870 Then posW = 3870:
        If posH < 1485 Then posH = 1485:
        FX = Int(IIf(FormFirst = True, FX, posW) / 15) * 15
        FY = Int(IIf(FormFirst = True, FY, posH) / 15) * 15
        If FormFirst = False Then
            Dim formW, formH
            picFormResize.Left = FX
            picFormResize.Top = FY
            formW = FX + 100
            formH = FY + 100
             Me.Move Me.Left, Me.Top, formW, formH
        End If
        FormFirst = False
    End If
End Sub

Private Sub picFormResize_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    FormFlag = False
    Me.AutoRedraw = False
    SaveSetting App.EXEName, "Visual", "StartSize", Str(Me.Width) & "," & Str(Me.Height)
End Sub

