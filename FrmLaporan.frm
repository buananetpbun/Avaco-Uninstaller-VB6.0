VERSION 5.00
Begin VB.Form FrmLaporan 
   BackColor       =   &H0000FF00&
   BorderStyle     =   0  'None
   Caption         =   "Form3"
   ClientHeight    =   5640
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8205
   LinkTopic       =   "Form3"
   Picture         =   "FrmLaporan.frx":0000
   ScaleHeight     =   5640
   ScaleWidth      =   8205
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   390
      Left            =   80
      Picture         =   "FrmLaporan.frx":0EBB
      ScaleHeight     =   390
      ScaleWidth      =   1605
      TabIndex        =   7
      Top             =   0
      Width           =   1605
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4610
      Left            =   120
      ScaleHeight     =   4605
      ScaleWidth      =   7965
      TabIndex        =   0
      Top             =   920
      Width           =   7970
      Begin VB.PictureBox picRuler 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   0
         Picture         =   "FrmLaporan.frx":1510
         ScaleHeight     =   270
         ScaleWidth      =   7950
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   0
         Width           =   7950
      End
      Begin VB.TextBox TxtLaporan 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   4360
         Left            =   0
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   1
         Top             =   240
         Width           =   7935
      End
   End
   Begin VB.Label LblTutup 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Close"
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   3840
      MouseIcon       =   "FrmLaporan.frx":B70A
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   570
      Width           =   1095
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   375
      Left            =   3840
      Shape           =   4  'Rounded Rectangle
      Top             =   480
      Width           =   1095
   End
   Begin VB.Image Image8 
      Height          =   480
      Left            =   7320
      Picture         =   "FrmLaporan.frx":BA14
      Top             =   360
      Width           =   480
   End
   Begin VB.Label LblBuka 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Open"
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   2640
      MouseIcon       =   "FrmLaporan.frx":C2DE
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   570
      Width           =   1095
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   375
      Left            =   2640
      Shape           =   4  'Rounded Rectangle
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label LblSimpan 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Save"
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   1440
      MouseIcon       =   "FrmLaporan.frx":C5E8
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   570
      Width           =   1095
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   375
      Left            =   1440
      Shape           =   4  'Rounded Rectangle
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label LblPrint 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Print"
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   240
      MouseIcon       =   "FrmLaporan.frx":C8F2
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   570
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   375
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   480
      Width           =   1095
   End
End
Attribute VB_Name = "FrmLaporan"
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

Sub LoadText()
Dim fName As String
Dim readln
Dim textload
On Error GoTo ld_err
fName = App.Path & "\temp" & ".tmp"
TxtLaporan.Text = ""
Open App.Path & "\temp" & ".tmp" For Input As #1
        Do While Not EOF(1)
            Line Input #1, readln
           textload = textload + readln + Chr$(13) + Chr$(10)
If Len(textload) >= 500000 Then
      MsgBox "This file {" & App.Path & "\temp" & ".tmp" & "} too large for open", vbCritical
      GoTo ld_end
    End If
  Loop
 TxtLaporan = textload
ld_err:
    Resume ld_end
ld_end:
    On Error Resume Next
    Close #1
End Sub

Private Sub Form_Load()
Call LoadText
Call LoadText
Kill App.Path & "\temp" & ".tmp"
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormDrag Me
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Panggil_Fontbold_tipis
End Sub

Private Sub LblBuka_Click()
  Dim filename As String
   filename = OpenDialog(Me, "Text Files (*.txt)|*.txt|All files (*.*)|*.*", _
                   "Open", "")
If Len(filename) Then
 On Error GoTo ld_err
Dim readln
Dim textload
TxtLaporan.Text = ""
 Open filename For Input As #1
        Do While Not EOF(1)
            Line Input #1, readln
            textload = textload + readln + Chr$(13) + Chr$(10)

   If Len(textload) >= 30000 Then
      MsgBox "This file {" & App.Path & "\temp" & ".tmp" & "} too large for open", vbCritical, "Error !!"
      GoTo ld_end
    End If
  Loop
 TxtLaporan = textload
ld_err:
 Resume ld_end
ld_end:
On Error Resume Next
  Close #1
  End If
End Sub

Private Sub LblBuka_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.LblPrint.FontBold = False
Me.LblBuka.FontBold = True
Me.LblSimpan.FontBold = False
Me.LblTutup.FontBold = False

End Sub

Private Sub LblPrint_Click()
On Error GoTo perbaiki:
Printer.Print ""
Printer.FontName = "Arial"
Printer.FontSize = 8
Printer.FontBold = False
Printer.Print Now
Printer.Print ""
Printer.Print Me.TxtLaporan.Text
Printer.EndDoc
Exit Sub
perbaiki:
MsgBox "Error : " & Err.Description & ", Please Check your printer ?", vbCritical, "Print Error !!"
End Sub

Private Sub LblPrint_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.LblPrint.FontBold = True
Me.LblBuka.FontBold = False
Me.LblSimpan.FontBold = False
Me.LblTutup.FontBold = False
End Sub

Private Sub LblSimpan_Click()
  Dim filename As String
  On Local Error Resume Next
  filename = SaveDialog(Me, "Text Files (*.txt)|*.txt", _
                       "Save", "", "")
If Len(filename) Then
     Close #1
Open filename For Output As #1
Print #1, TxtLaporan.Text
Close #1
  End If
End Sub

Private Sub LblSimpan_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.LblPrint.FontBold = False
Me.LblBuka.FontBold = False
Me.LblSimpan.FontBold = True
Me.LblTutup.FontBold = False

End Sub

Private Sub LblTutup_Click()
Unload Me
End Sub

Private Sub LblTutup_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.LblPrint.FontBold = False
Me.LblBuka.FontBold = False
Me.LblSimpan.FontBold = False
Me.LblTutup.FontBold = True

End Sub
Sub Panggil_Fontbold_tipis()
Me.LblPrint.FontBold = False
Me.LblBuka.FontBold = False
Me.LblSimpan.FontBold = False
Me.LblTutup.FontBold = False
End Sub

Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormDrag Me
End Sub

Private Sub TxtLaporan_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Panggil_Fontbold_tipis
End Sub
