VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form frmWebBrowse 
   Caption         =   "Verizon Messenger"
   ClientHeight    =   5580
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7515
   Icon            =   "frmWebBrowse.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5580
   ScaleWidth      =   7515
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin SHDocVwCtl.WebBrowser webBrowser 
      Height          =   5655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7575
      ExtentX         =   13361
      ExtentY         =   9975
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   1
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "res://C:\WINDOWS\SYSTEM\SHDOCLC.DLL/dnserror.htm#http:///"
   End
End
Attribute VB_Name = "frmWebBrowse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Form_Resize
End Sub

Private Sub Form_Resize()
webBrowser.Width = Me.Width - 500 - Me.Left
webBrowser.Height = Me.Height - 1000
End Sub

Private Sub webBrowser_DocumentComplete(ByVal pDisp As Object, URL As Variant)
On Error Resume Next
Kill "C:\tempVZW.html"
End Sub

