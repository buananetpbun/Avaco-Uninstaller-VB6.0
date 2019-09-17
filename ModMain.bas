Attribute VB_Name = "ModMain"
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

Sub Main()
  
    Load FrmMain
    FrmMain.Show
    Load FrmSplash
    FrmSplash.Show
 
End Sub

Public Sub SaveFile1(ff As Form)
On Error Resume Next
    Dim fName As String

    Dim i As Long
    fName = App.Path & "\" & "temp" & ".tmp"
    Close #1
       
            Open fName For Output As #1
            Print #1, "--> Create Report On : " & Now
            Print #1, "--> AVACO UNINSTALLER 2002 V1.00 - Copyright © 2002 AVACO, Pp"
            Print #1, "--> By Agus Ramadhani - AVACO_Soft22@Yahoo.com"
            Print #1, "--> Http://Avaco-Software.tripod.com"
            Print #1, "--> Find : " & ff.lstview.ListItems.Count & " Installed Programs"
            Print #1, ""
            Print #1, "INSTALLED PROGRAMS :"
            Print #1, ""
            For i = 1 To ff.lstview.ListItems.Count
            Print #1, ff.lstview.ListItems(i).Text
        Next i
    
End Sub

Public Sub SaveFile2(ff As Form)
On Error Resume Next
    Dim fName As String

    Dim i As Long
fName = App.Path & "\" & "temp" & ".tmp"
        Close #1
       
            Open fName For Output As #1
            Print #1, "--> Create Report On : " & Now
            Print #1, "--> AVACO UNINSTALLER 2002 V1.00 - Copyright © 2002 AVACO, Pp"
            Print #1, "--> By Agus Ramadhani - AVACO_Soft22@Yahoo.com"
            Print #1, "--> Http://Avaco-Software.tripod.com"
            Print #1, "--> Find : " & ff.lstview.ListItems.Count & " Installed Programs"
            Print #1, ""
            Print #1, "INSTALLED PROGRAMS :"
            Print #1, ""
            For i = 1 To ff.lstview.ListItems.Count
            Print #1, ff.lstview.ListItems(i).Text
            Print #1, ff.lstview.ListItems(i).ListSubItems(1).Text
            Print #1, ""
        Next i
    
End Sub

Public Sub SaveFile3(ff As Form)
On Error Resume Next
    Dim fName As String

    Dim i As Long
fName = App.Path & "\" & "temp" & ".tmp"
        Close #1
       
            Open fName For Output As #1
            Print #1, "--> Create Report On : " & Now
            Print #1, "--> AVACO UNINSTALLER 2002 V1.00 - Copyright © 2002 AVACO, Pp"
            Print #1, "--> By Agus Ramadhani - AVACO_Soft22@Yahoo.com"
            Print #1, "--> Http://Avaco-Software.tripod.com"
            Print #1, "--> Find : " & ff.lstview.ListItems.Count & " Installed Programs"
            Print #1, ""
            Print #1, "INSTALLED PROGRAMS :"
            Print #1, ""
            For i = 1 To ff.lstview.ListItems.Count
            Print #1, "Display Name --> " & ff.lstview.ListItems(i).Text
            Print #1, "Uninstall String --> " & ff.lstview.ListItems(i).ListSubItems(1).Text
            Print #1, "Registry Name --> " & ff.lstview.ListItems(i).ListSubItems(2).Text
            Print #1, "Publisher --> " & ff.lstview.ListItems(i).ListSubItems(3).Text
            Print #1, "Display Version --> " & ff.v.ListItems(i).ListSubItems(4).Text
            Print #1, "HelpLink --> " & ff.v.ListItems(i).ListSubItems(5).Text
            Print #1, "URL Info About --> " & ff.lstview.ListItems(i).ListSubItems(6).Text
            Print #1, "Contact --> " & ff.lstview.ListItems(i).ListSubItems(7).Text
            Print #1, ""
        Next i
    
End Sub

