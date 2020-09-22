VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   Caption         =   "MSN Messenger Contact Transfer Utility"
   ClientHeight    =   7965
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9390
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7965
   ScaleWidth      =   9390
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton SelectBut 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Refresh Contacts"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   3
      Left            =   165
      MaskColor       =   &H00FFFFFF&
      MouseIcon       =   "Form1.frx":617A
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   1695
      Width           =   4185
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   -90
      TabIndex        =   24
      Top             =   6120
      Width           =   9540
   End
   Begin VB.CommandButton SelectBut 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Select &REVERSE"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   2
      Left            =   7545
      MaskColor       =   &H00FFFFFF&
      MouseIcon       =   "Form1.frx":6A44
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   1695
      Width           =   1665
   End
   Begin VB.CommandButton SelectBut 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Select &NONE"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   1
      Left            =   5955
      MaskColor       =   &H00FFFFFF&
      MouseIcon       =   "Form1.frx":730E
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   1695
      Width           =   1575
   End
   Begin VB.CommandButton SelectBut 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Select &ALL"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   4365
      MaskColor       =   &H00FFFFFF&
      MouseIcon       =   "Form1.frx":7BD8
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   1695
      Width           =   1575
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   4725
      Top             =   5715
   End
   Begin VB.CommandButton Command3 
      Caption         =   "EXPORT EMAILS TO TEXT FILE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Left            =   1605
      MouseIcon       =   "Form1.frx":84A2
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":8D6C
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6300
      Width           =   1365
   End
   Begin VB.CommandButton Command2 
      Caption         =   "TRANSFER TO OTHER MSN"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Left            =   7860
      MouseIcon       =   "Form1.frx":EEE6
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":F7B0
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6285
      Width           =   1365
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5265
      Top             =   5535
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
            Picture         =   "Form1.frx":1592A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2C2EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":42CAE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView Lw 
      Height          =   3615
      Left            =   165
      TabIndex        =   4
      Top             =   2040
      Width           =   9045
      _ExtentX        =   15954
      _ExtentY        =   6376
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      PictureAlignment=   5
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
      Picture         =   "Form1.frx":59670
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SEND MESSAGE SELECTED"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1515
      Left            =   165
      MouseIcon       =   "Form1.frx":5A272
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":5AB3C
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6300
      Width           =   1365
   End
   Begin VB.Label Status 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blocked users"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Index           =   2
      Left            =   3270
      TabIndex        =   23
      Top             =   5730
      Width           =   1200
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   2
      Left            =   2985
      Stretch         =   -1  'True
      Top             =   5715
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   1
      Left            =   1560
      Stretch         =   -1  'True
      Top             =   5715
      Width           =   240
   End
   Begin VB.Label Status 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Online users"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Index           =   1
      Left            =   1845
      TabIndex        =   22
      Top             =   5730
      Width           =   1065
   End
   Begin VB.Label Status 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Offline users"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Index           =   0
      Left            =   420
      TabIndex        =   21
      Top             =   5730
      Width           =   1080
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   0
      Left            =   150
      Stretch         =   -1  'True
      Top             =   5700
      Width           =   240
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Free VB Project (2007)"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   4
      Left            =   7155
      TabIndex        =   17
      Top             =   1335
      Width           =   1965
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "www.cemhaner.com"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   3
      Left            =   7365
      TabIndex        =   16
      Top             =   1155
      Width           =   1740
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Coded by B.Cem HANER"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   2
      Left            =   7005
      TabIndex        =   15
      Top             =   990
      Width           =   2115
   End
   Begin VB.Label dedecttext 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DEDECT NEW MSN"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   0
      Left            =   6945
      TabIndex        =   14
      Top             =   255
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.Shape Dedectpan 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      Height          =   300
      Left            =   6570
      Shape           =   4  'Rounded Rectangle
      Top             =   225
      Visible         =   0   'False
      Width           =   2610
   End
   Begin VB.Label Label4 
      Caption         =   "TRANSFER BUTTON WILL ENABLING AUTOMATICALLY AFTER LOGIN..."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Index           =   2
      Left            =   3435
      TabIndex        =   13
      Top             =   7350
      Width           =   4230
   End
   Begin VB.Label Label4 
      Caption         =   "Please sign off current MSN account and log'in other account. Next; click a TRANSFER Button !"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Index           =   1
      Left            =   3435
      TabIndex        =   12
      Top             =   6585
      Width           =   4020
   End
   Begin VB.Label Label4 
      Caption         =   "OK ! ALL CONTACTS ARE RECEIVED !"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   3435
      TabIndex        =   11
      Top             =   6315
      Width           =   3885
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Logon email"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   315
      TabIndex        =   10
      Top             =   1200
      Width           =   1320
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Index           =   2
      Left            =   1800
      TabIndex        =   9
      Top             =   1200
      Width           =   60
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Active MSN Profile"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   1
      Left            =   255
      TabIndex        =   8
      Top             =   255
      Width           =   1965
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Index           =   1
      Left            =   1800
      TabIndex        =   7
      Top             =   945
      Width           =   60
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Your status"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   315
      TabIndex        =   6
      Top             =   945
      Width           =   1260
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "MSN Name"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   330
      TabIndex        =   5
      Top             =   675
      Width           =   1155
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Index           =   0
      Left            =   1800
      TabIndex        =   0
      Top             =   690
      Width           =   1755
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E4C2AB&
      BackStyle       =   1  'Opaque
      Height          =   360
      Index           =   1
      Left            =   165
      Top             =   195
      Width           =   9045
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      Height          =   1065
      Index           =   2
      Left            =   165
      Top             =   540
      Width           =   9045
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents msn As MessengerAPI.Messenger
Attribute msn.VB_VarHelpID = -1


Private Sub Command1_Click()
Dim MessageID As String
MessageID = Trim(Lw.SelectedItem.Text)
txtmsg = InputBox("Please write a message", MessageID & " -> will send") & ""
If Trim(txtmsg) = "" Then
    Exit Sub
End If


If Lw.ListItems.Count > 0 Then
    Set MsnWindow = msn.InstantMessage(MessageID)
    SendKeys txtmsg 'this is the message to be sent
    Pause 0.1
    SendKeys "{ENTER}"
    SendKeys "{ENTER}"
    SendKeys "{ESC}"
Else
    MsgBox "Select a contact", vbExclamation, "Error"
End If



End Sub

Private Sub Command2_Click()
On Error GoTo Hatam
Timer1.Enabled = False

Dim ContactEmail As String, t As Integer, i As Integer
If Not MsgBox(Me.Tag & "'s selected contacts are transferring from " & Me.Tag & " to " & Label2(1).Caption & vbLf & "Are you sure ?", vbInformation + vbYesNo) = vbYes Then
    Exit Sub
End If
    Command2.Enabled = False
    Me.SetFocus
    
    t = 0
    For i = 1 To Lw.ListItems.Count
        
        If Lw.ListItems.Item(i).Selected = True Then
            t = t + 1
            ContactEmail = Lw.ListItems.Item(i).Text
            msn.AddContact 0, Trim(ContactEmail)
            Pause 1
            SendKeys "{ENTER}"
            SendKeys "{ESC}"
        End If
    
    Next
    
    
    
    MsgBox t & " contacts transferred successfully!"
    Timer1.Enabled = True
    Command2.Enabled = True
Exit Sub

Hatam:
    MsgBox "Transfer error !", vbCritical + vbOKOnly, " Error"
    Timer1.Enabled = True
    Command2.Enabled = True
End Sub

Private Sub Command3_Click()
On Error GoTo HataVar
Dim OutFileName As String, i As Integer
OutFileName = App.Path & "\" & msn.MySigninName & ".TXT"

Close #1: Open OutFileName For Output As #1

For i = 1 To Lw.ListItems.Count
    Print #1, Trim(Lw.ListItems.Item(i).Text) & Chr(13)
Next i
Close #1

    MsgBox "All contacts are exported to " & vbLf & OutFileName, vbInformation + vbOKOnly, " Export successfully"

Exit Sub

HataVar:
    MsgBox "Contacts are not exported !", vbCritical + vbOKOnly, " Error"
    On Error GoTo 0
End Sub





Private Sub Form_Load()
Set msn = New MessengerAPI.Messenger
    
Call MSNStatus
If Label2(1).Caption = "UNKNOWN" Then
    MsgBox "Please logon  will transfer Messanger account and re-run program"
    End
Else
    Me.Tag = msn.MySigninName
End If
    
Aktarma = False

Lw.ColumnHeaders.Add 1, "A", "Contact address", 3000
Lw.ColumnHeaders.Add 2, "B", "User Identify", Lw.Width - 3250

Call UserAL
Image1(0).Picture = ImageList1.ListImages(2).Picture
Image1(1).Picture = ImageList1.ListImages(1).Picture
Image1(2).Picture = ImageList1.ListImages(3).Picture
End Sub


Private Sub MSNStatus()
On Error Resume Next
Select Case msn.MyStatus
    Case 14: Label2(1).Caption = "BE RIGHT BACK"
    Case 10: Label2(1).Caption = "BUSY"
    Case 2: Label2(1).Caption = "ONLINE"
    Case 3: Label2(1).Caption = "AWAY"
    Case 50: Label2(1).Caption = "IN A CALL"
    Case 66: Label2(1).Caption = "OUT TO LUNCH"
    Case 6: Label2(1).Caption = "APPEAR OFFLINE"
    Case Else: Label2(1).Caption = "UNKNOWN"
End Select
Label2(0).Caption = msn.MyFriendlyName
Label2(2).Caption = msn.MySigninName

End Sub

Private Sub Lw_DblClick()
'Call Command1_Click

    Dim msncontact As IMessengerContact
    Dim msncontacts As IMessengerContacts
    Set msncontacts = msn.MyContacts

End Sub

Private Sub SelectBut_Click(Index As Integer)
Select Case Index
    Case 0
        For i = 1 To Lw.ListItems.Count
            Lw.ListItems.Item(i).Selected = True
        Next i
    Case 1
        For i = 1 To Lw.ListItems.Count
            Lw.ListItems.Item(i).Selected = False
        Next i
    Case 2
        For i = 1 To Lw.ListItems.Count
            If Lw.ListItems.Item(i).Selected = True Then
                Lw.ListItems.Item(i).Selected = False
            Else
                Lw.ListItems.Item(i).Selected = True
            End If
        Next i
    Case 3
        UserAL
        Me.Tag = msn.MySigninName
End Select
Lw.SetFocus
End Sub

Private Sub Timer1_Timer()
On Error GoTo Hatali
Call MSNStatus
If Lw.ListItems.Count > 0 And msn.MySigninName <> Me.Tag Then
    dedecttext(0).Visible = True: Dedectpan.Visible = True: Command2.Enabled = True
End If

Exit Sub

Hatali:
On Error GoTo 0

End Sub

Private Sub UserAL()
Lw.ListItems.Clear
dedecttext(0).Visible = False: Dedectpan.Visible = False: Command2.Enabled = False
    
    Dim msncontact As IMessengerContact
    Dim msncontacts As IMessengerContacts
    Set msncontacts = msn.MyContacts
    


For Each msncontact In msncontacts
   Set itmx = Lw.ListItems.Add(, , " " & msncontact.SigninName, 2)
      If msncontact.Status = MISTATUS_OFFLINE Then
        itmx.SmallIcon = 2
      Else
        itmx.SmallIcon = 1
      End If
            
      If msncontact.Blocked = True Then
        itmx.SmallIcon = 3
      End If
      
      itmx.SubItems(1) = msncontact.FriendlyName
    

    'List1.AddItem (msncontact.SigninName)
Next

'If List1.ListCount < 1 Then
'    MsgBox "Please run to MSN Messenger..."
'    End
'End If
End Sub

Public Sub Pause(interval)
Current = Timer
Do While Timer - Current < Val(interval)
DoEvents
Loop
End Sub
 
