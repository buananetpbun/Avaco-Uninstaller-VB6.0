VERSION 5.00
Begin VB.Form FrmPopup 
   BackColor       =   &H000080FF&
   BorderStyle     =   0  'None
   ClientHeight    =   1260
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   6450
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1260
   ScaleWidth      =   6450
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu Menu 
      Caption         =   "Menu"
      Begin VB.Menu MnuUninstall 
         Caption         =   "Uninstall"
      End
      Begin VB.Menu MnuInformasi 
         Caption         =   "Information Details"
      End
      Begin VB.Menu MnuLaporan 
         Caption         =   "Report Information"
      End
      Begin VB.Menu spr1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRegEntry 
         Caption         =   "Registry Entry"
         Begin VB.Menu MnuEditEntry 
            Caption         =   "Edit Entry"
         End
         Begin VB.Menu MnuHEntry 
            Caption         =   "Delete Entry"
         End
         Begin VB.Menu spr2 
            Caption         =   "-"
         End
         Begin VB.Menu MnuEBaru 
            Caption         =   "New Entry"
         End
      End
      Begin VB.Menu MnuPReg 
         Caption         =   "Edit From Regedit"
      End
      Begin VB.Menu spr23 
         Caption         =   "-"
      End
      Begin VB.Menu MnuBakup 
         Caption         =   "Backup Uninstall Reg ..."
      End
      Begin VB.Menu spr4 
         Caption         =   "-"
      End
      Begin VB.Menu MnuPG 
         Caption         =   "Programs Group"
      End
      Begin VB.Menu spr12 
         Caption         =   "-"
      End
      Begin VB.Menu MnuRefresh 
         Caption         =   "Refresh"
      End
   End
   Begin VB.Menu Menu2 
      Caption         =   "Menu2"
      Begin VB.Menu MnuEBR 
         Caption         =   "Export Backup Registry"
      End
      Begin VB.Menu MnuHapusFile 
         Caption         =   "Delete File Backup"
      End
      Begin VB.Menu spr5 
         Caption         =   "-"
      End
      Begin VB.Menu MnuRefresh2 
         Caption         =   "Refresh"
      End
   End
   Begin VB.Menu Menu3 
      Caption         =   "Menu3"
      Begin VB.Menu mnuBUR 
         Caption         =   "Backup Uninstall Reg ..."
      End
      Begin VB.Menu MnuLaporan2 
         Caption         =   "Report Information"
      End
      Begin VB.Menu spr6 
         Caption         =   "-"
      End
      Begin VB.Menu Keluar 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu Menu4 
      Caption         =   "Menu4"
      Begin VB.Menu mnuPilihan 
         Caption         =   "Options"
      End
   End
   Begin VB.Menu Menu5 
      Caption         =   "Menu5"
      Begin VB.Menu MnuUninstall2 
         Caption         =   "Uninstall"
      End
      Begin VB.Menu spr7 
         Caption         =   "-"
      End
      Begin VB.Menu MnuDinformasi 
         Caption         =   "Information Detail"
      End
   End
   Begin VB.Menu Menu6 
      Caption         =   "Menu6"
      Begin VB.Menu MnuEEntry2 
         Caption         =   "Edit Entry"
      End
      Begin VB.Menu MnuHEntry2 
         Caption         =   "Delete Entry"
      End
      Begin VB.Menu spr9 
         Caption         =   "-"
      End
      Begin VB.Menu MnuEBaru2 
         Caption         =   "New Entry"
      End
      Begin VB.Menu spr13 
         Caption         =   "-"
      End
      Begin VB.Menu MnuPReg2 
         Caption         =   "Edit From Regedit"
      End
      Begin VB.Menu MnuPG2 
         Caption         =   "Programs Group"
      End
   End
   Begin VB.Menu Menu7 
      Caption         =   "Menu7"
      Begin VB.Menu MnuDetail 
         Caption         =   "Details"
      End
      Begin VB.Menu MnuList 
         Caption         =   "List"
      End
      Begin VB.Menu mnuIconBesar 
         Caption         =   "Small Icon"
      End
      Begin VB.Menu MnuIconKecil 
         Caption         =   "Large Icon "
      End
      Begin VB.Menu spr10 
         Caption         =   "-"
      End
      Begin VB.Menu MnuMinimalWindow 
         Caption         =   "Minimize"
      End
      Begin VB.Menu MnuKWindow 
         Caption         =   "Restore"
      End
      Begin VB.Menu MnuMaksWindow 
         Caption         =   "Maximize"
      End
   End
   Begin VB.Menu Menu8 
      Caption         =   "Menu8"
      Begin VB.Menu MnuBantuan 
         Caption         =   "Help Topics"
      End
      Begin VB.Menu spr11 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About Avaco"
      End
   End
End
Attribute VB_Name = "FrmPopup"
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

Private Sub Keluar_Click()
Unload FrmMain
End Sub

Private Sub mnuAbout_Click()
FrmAbout.Show vbModal, FrmMain
End Sub

Private Sub MnuBakup_Click()
FrmMain.Backup_Registry
FrmBackupRegistry.Show vbModal, FrmMain
End Sub

Private Sub MnuBantuan_Click()
frmBantuan.Show
End Sub

Private Sub mnuBUR_Click()
FrmMain.Backup_Registry
FrmBackupRegistry.Show vbModal, FrmMain
End Sub

Private Sub MnuDetail_Click()
FrmMain.lstview.View = lvwReport
End Sub


Private Sub MnuDinformasi_Click()
frmInformasi.Show vbModal, FrmMain
End Sub

Private Sub MnuEBaru_Click()
FrmEntryBaru.Show vbModal, FrmMain
End Sub

Private Sub MnuEBaru2_Click()
FrmEntryBaru.Show vbModal, FrmMain
End Sub

Private Sub MnuEBR_Click()
FrmBackupRegistry.kembalikan_registry
End Sub

Private Sub MnuEditEntry_Click()
FrmMain.MnuEditEntry_Click
End Sub

Private Sub MnuEEntry2_Click()
FrmMain.MnuEditEntry_Click
End Sub


Private Sub MnuHapusFile_Click()
FrmBackupRegistry.Hapus
End Sub

Private Sub MnuHEntry_Click()
FrmMain.MnuHapusEntry_Click
End Sub

Private Sub MnuHEntry2_Click()
FrmMain.MnuHapusEntry_Click
End Sub

Private Sub mnuIconBesar_Click()
FrmMain.lstview.View = lvwIcon
End Sub

Private Sub MnuIconKecil_Click()
FrmMain.lstview.View = lvwSmallIcon
End Sub

Private Sub MnuInformasi_Click()
frmInformasi.Show vbModal, FrmMain
End Sub

Private Sub MnuKWindow_Click()
FrmMain.WindowState = vbNormal
End Sub

Private Sub MnuLaporan_Click()
FrmPilihLaporan.Show vbModal, FrmMain
End Sub

Private Sub MnuLaporan2_Click()
FrmPilihLaporan.Show vbModal, FrmMain
End Sub

Private Sub MnuList_Click()
FrmMain.lstview.View = lvwList
End Sub

Private Sub MnuMaksWindow_Click()
FrmMain.WindowState = vbMaximized
End Sub

Private Sub MnuMinimalWindow_Click()
FrmMain.WindowState = vbMinimized
End Sub

Private Sub MnuPG_Click()
Shell ("explorer C:\WINDOWS\Start Menu\Programs"), vbNormalFocus
End Sub

Private Sub MnuPG2_Click()
Shell ("explorer C:\WINDOWS\Start Menu\Programs"), vbNormalFocus
End Sub

Private Sub mnuPilihan_Click()
FrmOption.Show vbModal, FrmMain
End Sub

Private Sub MnuPReg_Click()
Dim result
result = MsgBox("You want to modify from regedit ? for information location is :" + vbCrLf & _
"[HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{Nama Program}]", vbYesNo, "Call Regedit.exe")
If result = vbYes Then
Shell ("c:\windows\regedit.exe"), vbNormalFocus
End If
End Sub

Private Sub MnuPReg2_Click()
Call MnuPReg_Click
End Sub

Private Sub MnuRefresh_Click()
FrmMain.New_Refresh
End Sub

Private Sub MnuRefresh2_Click()
FrmBackupRegistry.TreeViewBackup.Refresh
FrmBackupRegistry.panggil_Nodes
End Sub

Private Sub MnuUninstall_Click()
FrmMain.Show_FormUninstall
End Sub

Private Sub MnuUninstall2_Click()
FrmMain.Show_FormUninstall
End Sub
