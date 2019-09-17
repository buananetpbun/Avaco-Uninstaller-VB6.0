Attribute VB_Name = "ModHelp"
Option Explicit
Option Base 1

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

Public Const LB_FINDSTRINGEXACT = &H1A2

Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" _
(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long



