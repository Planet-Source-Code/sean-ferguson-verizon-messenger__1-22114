VERSION 5.00
Begin VB.Form frmSendTo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Send To"
   ClientHeight    =   2610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4320
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2610
   ScaleWidth      =   4320
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   2160
      Width           =   2895
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove"
      Height          =   315
      Left            =   3120
      TabIndex        =   3
      Top             =   1132
      Width           =   1095
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   315
      Left            =   3120
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox txtAddNumber 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2895
   End
   Begin VB.ListBox lstRecipients 
      Height          =   1620
      ItemData        =   "frmSendTo.frx":0000
      Left            =   120
      List            =   "frmSendTo.frx":0002
      TabIndex        =   2
      Top             =   480
      Width           =   2895
   End
End
Attribute VB_Name = "frmSendTo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()
If txtAddNumber.Text = "" Then Exit Sub
lstRecipients.AddItem txtAddNumber.Text
txtAddNumber.Text = ""
If lstRecipients.ListCount >= 10 Then
    cmdAdd.Enabled = False
    txtAddNumber.Enabled = False
End If
End Sub

Private Sub cmdOK_Click()
txtAddNumber.Text = ""
If lstRecipients.ListCount = 0 Then
    frmMessenger.recipients.Caption = "(Please click the button to the right to select recipients)"
    frmMessenger.cmdSend.Enabled = False
    Me.Hide
    Exit Sub
End If
frmMessenger.recipients.Caption = lstRecipients.List(0)
If lstRecipients.ListCount > 1 Then
    For X = 1 To lstRecipients.ListCount - 1
        frmMessenger.recipients.Caption = frmMessenger.recipients.Caption & ", " & lstRecipients.List(X)
    Next X
End If
Me.Hide
If frmMessenger.txtFrom.Text <> "" And frmMessenger.txtMessage.Text <> "" Then frmMessenger.cmdSend.Enabled = True
End Sub

Private Sub cmdRemove_Click()
If lstRecipients.SelCount > 0 Then lstRecipients.RemoveItem lstRecipients.ListIndex
If lstRecipients.ListCount < 10 Then
    cmdAdd.Enabled = True
    txtAddNumber.Enabled = True
End If
End Sub
