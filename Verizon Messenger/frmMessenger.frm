VERSION 5.00
Begin VB.Form frmMessenger 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Verizon Messenger"
   ClientHeight    =   5640
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6555
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmMessenger.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MouseIcon       =   "frmMessenger.frx":0442
   ScaleHeight     =   5640
   ScaleWidth      =   6555
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear Form"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4560
      Width           =   2175
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send Message"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4560
      Width           =   2175
   End
   Begin VB.OptionButton priHigh 
      BackColor       =   &H00000000&
      Caption         =   "High"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2160
      TabIndex        =   7
      Top             =   4080
      Width           =   735
   End
   Begin VB.OptionButton priNormal 
      BackColor       =   &H00000000&
      Caption         =   "Normal"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1200
      TabIndex        =   6
      Top             =   4080
      Value           =   -1  'True
      Width           =   855
   End
   Begin VB.TextBox txtNumLeft 
      Height          =   285
      Left            =   6000
      Locked          =   -1  'True
      MaxLength       =   120
      TabIndex        =   17
      TabStop         =   0   'False
      Text            =   "120"
      Top             =   3480
      Width           =   375
   End
   Begin VB.TextBox txtSubject 
      Height          =   285
      Left            =   1200
      MaxLength       =   120
      TabIndex        =   4
      Top             =   2085
      Width           =   5175
   End
   Begin VB.TextBox txtCallbackNum 
      Height          =   285
      Left            =   1200
      MaxLength       =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   5175
   End
   Begin VB.TextBox txtMessage 
      Height          =   855
      Left            =   1200
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   2520
      Width           =   5175
   End
   Begin VB.TextBox txtFrom 
      Height          =   285
      Left            =   1200
      MaxLength       =   120
      TabIndex        =   2
      Text            =   "(Enter your e-mail address)"
      Top             =   1275
      Width           =   5175
   End
   Begin VB.CommandButton cmdDeliverTo 
      Caption         =   "..."
      Height          =   255
      Left            =   6120
      TabIndex        =   1
      Top             =   960
      Width           =   255
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Field names colored in white are required, all others are optional."
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   5280
      Width           =   6375
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Priority:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   19
      Top             =   4080
      Width           =   855
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00808080&
      X1              =   1200
      X2              =   6360
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00808080&
      X1              =   4200
      X2              =   4200
      Y1              =   3960
      Y2              =   3480
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "E-mail address, subject, and message all affect the 120 character maximum."
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   1200
      TabIndex        =   18
      Top             =   3525
      Width           =   2895
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Characters Remaining:"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   4320
      TabIndex        =   16
      Top             =   3525
      Width           =   1695
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Subject:"
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   2130
      Width           =   855
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Callback #:"
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   1725
      Width           =   855
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Message:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "From:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label recipients 
      BackStyle       =   0  'Transparent
      Caption         =   "(Please click the button to the right to select recipients)"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1200
      TabIndex        =   11
      Top             =   960
      Width           =   4815
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   120
      X2              =   6480
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Deliver To:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Please fill in the fields below to send a message to a Verizon Wireless Celluar Phone."
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   2400
      TabIndex        =   0
      Top             =   360
      Width           =   4095
   End
   Begin VB.Image imgVZW 
      Height          =   885
      Left            =   -240
      Picture         =   "frmMessenger.frx":074C
      Stretch         =   -1  'True
      ToolTipText     =   "Go to myVZW.com"
      Top             =   -120
      Width           =   2625
   End
End
Attribute VB_Name = "frmMessenger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClear_Click()
clearYN = MsgBox("Do you want to clear the form?", vbYesNo, "Clear")
If clearYN = vbYes Then
    txtFrom.Text = "(Enter your e-mail address)"
    txtMessage.Text = ""
    txtCallbackNum.Text = ""
    txtSubject.Text = ""
    txtNumLeft.Text = 120
    priNormal.Value = True
    recipients.Caption = "(Please click the button to the right to select recipients)"
    frmSendTo.lstRecipients.Clear
    cmdSend.Enabled = False
End If
End Sub

Private Sub cmdDeliverTo_Click()
frmSendTo.Show , frmMessenger
End Sub

Private Sub cmdSend_Click()
vzwDisclaimer = "________________________________________________________________________" & vbCrLf & _
                "                               DISCLAIMER                               " & vbCrLf & _
                "________________________________________________________________________" & vbCrLf & _
                vbCrLf & _
                "Verizon wireless is not responsible for any messages that are lost or" & vbCrLf & _
                "significantly delayed due to transmission via the Internet." & vbCrLf & _
                "Information sent via text messages, including your wireless phone" & vbCrLf & _
                "number, may be intercepted by third parties during transmission over" & vbCrLf & _
                "the Internet without your or Verizon Wireless' permission. Verizon" & vbCrLf & _
                "Wireless is not responsible for the number or content of messages sent" & vbCrLf & _
                "to customers using the Wireless Text Messaging Service. Customers are" & vbCrLf & _
                "responsible for the cost of messages recieved in excess of the text" & vbCrLf & _
                "messaging plan chosen."
sendMsg = MsgBox(vzwDisclaimer, vbOKCancel, "Disclaimer")
If sendMsg = vbCancel Then Exit Sub
Dim msgPriority As String
If priNormal.Value = True Then msgPriority = "normal" Else msgPriority = "urgent"
Open "C:\tempVZW.html" For Output As #1
Print #1, "<HTML>"
Print #1, "<BODY ONLOAD='document.form.submit();'>"
Print #1, "Please wait... Redirecting to Verizon Wireless..."
Print #1, "<DIV STYLE='visibility: hidden;'>"
Print #1, "<FORM ACTION='http://www.msg.myvzw.com/servlet/SmsServlet' METHOD=POST NAME=form>"
Print #1, "<INPUT TYPE=HIDDEN NAME=bgcolor VALUE=FFFFFF>"
Print #1, "<INPUT TYPE=HIDDEN NAME=msg_type VALUE=messaging>"
Print #1, "<TEXTAREA NAME=min>" & recipients.Caption & "</TEXTAREA>"
Print #1, "<TEXTAREA NAME=senderName>" & txtFrom.Text & "</TEXTAREA>"
Print #1, "<TEXTAREA NAME=subject>" & txtSubject.Text & "</TEXTAREA>"
Print #1, "<TEXTAREA NAME=from>" & txtCallbackNum.Text & "</TEXTAREA>"
Print #1, "<INPUT TYPE=HIDDEN NAME=priority VALUE=" & msgPriority & ">"
Print #1, "<TEXTAREA NAME=message>" & txtMessage.Text & "</TEXTAREA>"
Print #1, "</FORM>"
Print #1, "</DIV>"
Print #1, "</BODY>"
Print #1, "</HTML>"
Close #1
If MsgBox("Do you wish to view the results of your send?", vbYesNo + vbQuestion, "View Results") = vbYes Then frmWebBrowse.Show Else Call MsgBox("Your message has been sent.", vbOKOnly, "Sent")
frmWebBrowse.webBrowser.Navigate "C:\tempVZW.html"
    txtFrom.Text = "(Enter your e-mail address)"
    txtMessage.Text = ""
    txtCallbackNum.Text = ""
    txtSubject.Text = ""
    txtNumLeft.Text = 120
    priNormal.Value = True
    recipients.Caption = "(Please click the button to the right to select recipients)"
    frmSendTo.lstRecipients.Clear
    cmdSend.Enabled = False
End Sub

Private Sub txtFrom_GotFocus()
If txtFrom.Text = "(Enter your e-mail address)" Then txtFrom.Text = ""
txtFrom.MaxLength = 120 - Len(txtMessage.Text) - Len(txtSubject.Text) - Len(txtCallbackNum.Text)
End Sub

Private Sub txtFrom_Change()
txtNumLeft.Text = 120 - Len(txtMessage.Text) - Len(txtSubject.Text) - Len(txtCallbackNum.Text) - Len(txtFrom.Text)
If txtFrom.Text <> "" And txtMessage.Text <> "" And frmSendTo.lstRecipients.ListCount > 0 Then cmdSend.Enabled = True Else cmdSend.Enabled = False
End Sub

Private Sub txtMessage_GotFocus()
txtMessage.MaxLength = 120 - Len(txtFrom.Text) - Len(txtSubject.Text) - Len(txtCallbackNum.Text)
End Sub

Private Sub txtMessage_Change()
txtNumLeft.Text = 120 - Len(txtFrom.Text) - Len(txtSubject.Text) - Len(txtCallbackNum.Text) - Len(txtMessage.Text)
If txtFrom.Text <> "" And txtMessage.Text <> "" And frmSendTo.lstRecipients.ListCount > 0 Then cmdSend.Enabled = True Else cmdSend.Enabled = False
End Sub

Private Sub txtSubject_GotFocus()
txtSubject.MaxLength = 120 - Len(txtFrom.Text) - Len(txtMessage.Text) - Len(txtCallbackNum.Text)
End Sub

Private Sub txtSubject_Change()
txtNumLeft.Text = 120 - Len(txtFrom.Text) - Len(txtMessage.Text) - Len(txtCallbackNum.Text) - Len(txtSubject.Text)
End Sub

Private Sub txtCallbackNum_GotFocus()
txtCallbackNum.MaxLength = 120 - Len(txtFrom.Text) - Len(txtMessage.Text) - Len(txtSubject.Text)
End Sub

Private Sub txtCallbackNum_Change()
txtNumLeft.Text = 120 - Len(txtFrom.Text) - Len(txtMessage.Text) - Len(txtSubject.Text) - Len(txtCallbackNum.Text)
End Sub

