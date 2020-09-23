VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   8355
   ClientLeft      =   5640
   ClientTop       =   1545
   ClientWidth     =   4335
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "MSN.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "MSN.frx":000C
   ScaleHeight     =   8355
   ScaleWidth      =   4335
   Begin MSComDlg.CommonDialog cd 
      Left            =   600
      Top             =   7440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "*.txt"
      DialogTitle     =   "Save As..."
      FileName        =   " *.txt"
      Filter          =   "*.txt| Text File"
      Orientation     =   2
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6240
      Top             =   3360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   49
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MSN.frx":BF062
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MSN.frx":BF9F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MSN.frx":C038A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MSN.frx":C0D1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MSN.frx":C16B2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   2400
      Top             =   7440
   End
   Begin MSComctlLib.ImageList img 
      Left            =   6240
      Top             =   6360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   16777215
      ImageWidth      =   17
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MSN.frx":C2046
            Key             =   "online"
            Object.Tag             =   "Pic_On"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MSN.frx":C23DE
            Key             =   "Pic_Off"
            Object.Tag             =   "Pic_Off"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MSN.frx":C2776
            Key             =   "Pic_time"
            Object.Tag             =   "Pic_time"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MSN.frx":C2B0E
            Key             =   "Pic_away"
            Object.Tag             =   "Pic_away"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MSN.frx":C2EA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MSN.frx":C337E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MSN.frx":C3856
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MSN.frx":C3D2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MSN.frx":C4206
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6615
      Left            =   0
      Picture         =   "MSN.frx":C458E
      ScaleHeight     =   6615
      ScaleWidth      =   4335
      TabIndex        =   0
      Top             =   0
      Width           =   4335
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   600
         TabIndex        =   4
         ToolTipText     =   "Click For Name Changer"
         Top             =   600
         Width           =   3495
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   600
         TabIndex        =   3
         ToolTipText     =   "Unread Messages"
         Top             =   960
         Width           =   3495
      End
      Begin MSComctlLib.TreeView tv 
         Height          =   4740
         Left            =   120
         TabIndex        =   5
         Top             =   1440
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   8361
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   0
         Style           =   5
         ImageList       =   "img"
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Image Image6 
         Height          =   300
         Left            =   4040
         Picture         =   "MSN.frx":125A78
         Top             =   30
         Width           =   285
      End
      Begin VB.Image Image1 
         Height          =   270
         Left            =   3720
         Picture         =   "MSN.frx":125F6A
         Top             =   30
         Width           =   330
      End
      Begin VB.Image Image3 
         Height          =   180
         Left            =   840
         Picture         =   "MSN.frx":126474
         Top             =   6240
         Width           =   2355
      End
      Begin VB.Image im2 
         Height          =   255
         Left            =   120
         Picture         =   "MSN.frx":127AD6
         Top             =   960
         Width           =   285
      End
      Begin VB.Image im1 
         Height          =   255
         Left            =   120
         ToolTipText     =   "Change Status"
         Top             =   600
         Width           =   375
      End
      Begin VB.Image Image5 
         Height          =   15
         Left            =   600
         Top             =   840
         Width           =   3255
      End
      Begin VB.Image Image2 
         Height          =   15
         Left            =   720
         Picture         =   "MSN.frx":127F14
         Top             =   720
         Width           =   2700
      End
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   960
      TabIndex        =   1
      Text            =   "Please Log In to MSN..."
      Top             =   840
      Width           =   1815
   End
   Begin VB.Image Image8 
      Height          =   150
      Left            =   3960
      Picture         =   "MSN.frx":128172
      ToolTipText     =   "View Other List"
      Top             =   6750
      Width           =   240
   End
   Begin VB.Label stat 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   6720
      Width           =   3615
   End
   Begin VB.Image Image4 
      Height          =   465
      Left            =   0
      Picture         =   "MSN.frx":128394
      Top             =   6600
      Width           =   4395
   End
   Begin VB.Menu file 
      Caption         =   "File"
      Visible         =   0   'False
      Begin VB.Menu name 
         Caption         =   "Change Name"
      End
      Begin VB.Menu exp 
         Caption         =   "Export List"
      End
   End
   Begin VB.Menu list 
      Caption         =   "List"
      Visible         =   0   'False
      Begin VB.Menu frnd 
         Caption         =   "Friends List"
         Checked         =   -1  'True
      End
      Begin VB.Menu block 
         Caption         =   "Blocked List"
         Checked         =   -1  'True
      End
      Begin VB.Menu who 
         Caption         =   "Whos Got Me Added?"
         Checked         =   -1  'True
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public WithEvents MSN As MsgrObject
Attribute MSN.VB_VarHelpID = -1
' Dim NewChat(4) As IMsgrIMSession


Private Sub block_Click()
ReFrEsH_It MLIST_BLOCK
block.Checked = True
who.Checked = False
frnd.Checked = False
End Sub

Private Sub Command1_Click()
Send_List 0
End Sub

Private Sub exit_Click()
End
End Sub


Private Sub Form_Load()
Set MSN = New MsgrObject
Dim FriendlyName As String
Dim P As Integer
Me.Height = 7125
Me.Show

Timer1.Enabled = False
block.Checked = False
who.Checked = False
stat.Caption = "MSN Status Viewer by DeeP_VeiN"
tv.ImageList = img
Me.Show
Form2.ao.Checked = False
Form2.awak.Checked = False
Form2.brb.Checked = False
Form2.busy.Checked = False
Form2.online.Checked = False
Form2.otf.Checked = False
Form2.otl.Checked = False
ReFrEsH_It MLIST_CONTACT
Me.Show
End Sub

Private Sub frnd_Click()
ReFrEsH_It MLIST_CONTACT
block.Checked = False
who.Checked = False
frnd.Checked = True
End Sub

Private Sub im1_Click()
Me.PopupMenu Form2.status
End Sub

Private Sub lists_Click()
Form2.Show
End Sub


Public Function State(ByVal str As String) As String
If str = 34 Then State = "Away"
If str = 14 Then State = "Be Right Back"
If str = 10 Then State = "Busy"
If str = 18 Then State = "Idle"
If str = 6 Then State = "Invisible"
If str = 1 Then State = "Offline"
If str = 50 Then State = "On The Phone"
If str = 2 Then State = "Online"
If str = 66 Then State = "Out To Lunch"
If str = MSTATE_UNKNOWN Then State = "Unknown"
End Function


Function ReFrEsH_It(list As MLIST)
DoEvents
If MSN.LocalState = MSTATE_OFFLINE Then Timer1.Enabled = True
If MSN.LocalState = MSTATE_OFFLINE Then Picture1.Visible = False
If MSN.LocalState = MSTATE_OFFLINE Then Exit Function

x = MSN.list(list).Count
User_Update
'Clear_all
tv.Nodes.Clear
tv.Nodes.Add , , "online", "Online"
tv.Nodes.Add , , "offline", "Offline"
tv.Nodes.Item(1).Expanded = True
tv.Nodes.Item(2).Expanded = True
tv.Nodes.Item(1).Bold = True
tv.Nodes.Item(2).Bold = True


For P = 1 To x - 1

stat.Caption = "Loading... (" & P & ") of (" & x & ")"
FriendlyName = MSN.list(list)(P).FriendlyName
stats = State(MSN.list(list)(P).State)
'Save_User P, list
AddNode FriendlyName, stats
Next P
stat.Caption = "Loading Complete"
If list = MLIST_CONTACT Then stat.Caption = "Contact List"
If list = MLIST_BLOCK Then stat.Caption = "Blocked List"
If list = MLIST_REVERSE Then stat.Caption = "People Who have you Added"


On Error GoTo error
If tv.Nodes.Item(1).Child.Text = "Offline" Then tv.Nodes.Add "online", tvwChild, , "No Friends Online", 9
If tv.Nodes.Item(2).Child.Text = "Offline" Then tv.Nodes.Add "offline", tvwChild, , "No Friends Offine", 9
Exit Function
error:
tv.Nodes.Add "online", tvwChild, , "No Friends Online", 9
tv.Nodes.Add "offline", tvwChild, , "No Friends Offine", 9
End Function

Function User_Update()
DoEvents

Text2.FontUnderline = True
Text2.Text = MSN.LocalFriendlyName
Text1.FontUnderline = True
Text1.Text = MSN.UnreadEmail(MFOLDER_INBOX) & " new e-mail messages"
im1.Refresh
If MSN.LocalState = 1 Then im1.Picture = img.ListImages(8).Picture
If MSN.LocalState = 6 Then im1.Picture = img.ListImages(8).Picture

If MSN.LocalState = 34 Then im1.Picture = img.ListImages(6).Picture
If MSN.LocalState = 18 Then im1.Picture = img.ListImages(6).Picture
If MSN.LocalState = 50 Then im1.Picture = img.ListImages(6).Picture

If MSN.LocalState = 10 Then im1.Picture = img.ListImages(7).Picture
If MSN.LocalState = 66 Then im1.Picture = img.ListImages(7).Picture
If MSN.LocalState = 14 Then im1.Picture = img.ListImages(7).Picture

If MSN.LocalState = 2 Then im1.Picture = img.ListImages(5).Picture
im1.Refresh
End Function

Private Sub OnUnreadEmailChanged(ByVal MFOLDER As MFOLDER, ByVal cUnreadEmail As Long, ByVal pfEnableDefault As Boolean)
If MFOLDER = MFOLDER_ALL_OTHER_FOLDERS Then Else Form1.User_Update
End Sub

Private Sub OnLocalFriendlyNameChangeResult(hr As Long, pService As IMsgrService, bstrPrevFriendlyName As String)
Form1.User_Update
End Sub
Sub OnLocalStateChangeResult()
Form1.User_Update
End Sub

Private Sub Image1_Click()
Me.WindowState = vbMinimized
End Sub

Private Sub Image7_Click()
tv.SetFocus
SendKeys "{end}"
tv.Refresh
End Sub

Private Sub Image6_Click()
Dim response As VbMsgBoxResult
response = MsgBox("Are you sure you want to Quit?", vbApplicationModal + vbInformation + vbYesNo, "MSN Status Viewer")
If response = vbYes Then End
End Sub

Private Sub Image8_Click()
Me.PopupMenu Me.list
End Sub

Private Sub MSN_OnUserStateChanged(ByVal pUser As IMsgrUser, ByVal mPrevState As MSTATE, pfEnableDefault As Boolean)
Dim name As String
'stats = State(pUser.State)

'If stats = "Offline" Then name = pUser.FriendlyName Else name = pUser.FriendlyName & " (" & State(mPrevState) & ")"
'If stats = "Online" Then name = pUser.FriendlyName Else name = pUser.FriendlyName & " (" & State(mPrevState) & ")"

name = pUser.FriendlyName
DeleteNode name & " (" & State(mPrevState) & ")"

AddNode pUser.FriendlyName, State(pUser.State)
Change_State pUser.FriendlyName, pUser.State
End Sub

Private Sub MSN_OnUserFriendlyNameChangeResult(ByVal hr As Long, ByVal pUser As IMsgrUser, ByVal bstrPrevFriendlyName As String)

Dim name As String
stats = State(pUser.State)

If stats = "Offline" Then name = bstrPrevFriendlyName Else name = bstrPrevFriendlyName & " (" & State(pUser.State) & ")"
If stats = "Online" Then name = bstrPrevFriendlyName Else name = bstrPrevFriendlyName & " (" & State(pUser.State) & ")"

DeleteNode name

AddNode pUser.FriendlyName, State(pUser.State)
End Sub

Private Sub name_Click()
Load Form2
Form2.Show
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim Ret&
If Button = 1 Then
ReleaseCapture
Ret& = SendMessage(Me.hwnd, &H112, &HF012, 0)
End If

If Button = 2 Then Me.PopupMenu Me.file
End Sub

Private Sub Text2_Click()
Load Form2
Form2.Show
End Sub

Private Sub Timer1_Timer()
If MSN.LocalState = MSTATE_LOCAL_CONNECTING_TO_SERVER Then Text3.Text = "Connecting..."
If MSN.LocalState = MSTATE_LOCAL_DISCONNECTING_FROM_SERVER Then Text3.Text = "Disconnecting..."
If MSN.LocalState = MSTATE_LOCAL_FINDING_SERVER Then Text3.Text = "Finding Server..."
If MSN.LocalState = MSTATE_LOCAL_SYNCHRONIZING_WITH_SERVER Then Text3.Text = "Synchronizing"
If MSN.LocalState = MSTATE_OFFLINE Then Text3.Text = "Please Sing Into MSN..."
If MSN.LocalState = MSTATE_ONLINE Then ReFrEsH_It MLIST_CONTACT
If MSN.LocalState = MSTATE_ONLINE Then Timer1.Enabled = False
End Sub

 Sub tv_DblClick()
 
 'Ideas for message sending
 
'Dim windows As IMsgrIMSession
'Dim a As IMsgrSP
'Dim b As IMsgrUser
't = Find_User(tv.SelectedItem.Text)
'b = MSN.list(0)(t)
'windows = a.CreateIMSession(b)

End Sub
Private Sub exp_Click()
cd.ShowSave
If cd.FileName = "*.txt" Then Exit Sub
Export_List cd.FileName
End Sub

Private Sub who_Click()
ReFrEsH_It MLIST_REVERSE
block.Checked = False
who.Checked = True
frnd.Checked = False
End Sub
Public Sub DeleteNode(ByVal name As String)
Dim temp As String
For x = 1 To tv.Nodes.Count - 1
    If tv.Nodes.Item(x).Text = name Then tv.Nodes.Remove (x)
Next x
End Sub

Private Sub AddNode(ByVal FriendlyName As String, ByVal stats As String)

If stats = "Offline" Then tv.Nodes.Add "offline", tvwChild, , FriendlyName, 2
If stats = "Away" Then tv.Nodes.Add "online", tvwChild, , FriendlyName & " (" & stats & ")", 3
If stats = "Be Right Back" Then tv.Nodes.Add "online", tvwChild, , FriendlyName & " (" & stats & ")", 3
If stats = "Busy" Then tv.Nodes.Add "online", tvwChild, , FriendlyName & " (" & stats & ")", 4
If stats = "Idle" Then tv.Nodes.Add "online", tvwChild, , FriendlyName & " (" & stats & ")", 3
If stats = "Invisible" Then tv.Nodes.Add "online", tvwChild, , FriendlyName & " (" & stats & ")", 9
If stats = "Online" Then tv.Nodes.Add "online", tvwChild, , FriendlyName, 1
If stats = "On The Phone" Then tv.Nodes.Add "online", tvwChild, , FriendlyName & " (" & stats & ")", 3
If stats = "Out To Lunch" Then tv.Nodes.Add "online", tvwChild, , FriendlyName & " (" & stats & ")", 4

End Sub

Public Function NodeIndex(ByVal name As String) As Integer
Dim temp As String
For x = 1 To tv.Nodes.Count - 1
    If tv.Nodes.Item(x).Text = name Then NodeIndex = x
Next x
End Function

Function Export_List(path As String)
On Error GoTo error
Open path For Output As 1#
Print #1, "|------List For " & MSN.LocalLogonName & "------|" & vbCrLf
For x = 1 To MSN.list(0).Count - 1
Print #1, MSN.list(0)(x).FriendlyName & " <" & MSN.list(0)(x).EmailAddress & ">" & vbCr
Next x

Close #1
MsgBox "File Saved!", vbApplicationModal + vbInformation, "MSN Status Viewer"
Exit Function
error:
MsgBox "Error Saveing File", vbApplicationModal + vbCritical, "MSN Status Viewer"
End Function

' The rest of this stuff was something I was screwing around with for the next version

' I was going to make it save all the strings to emulate the IMsgrUser type
' But its however not working, and i cant be bothered takeing it out

Private Sub Save_User(ByVal num As Integer, ByVal list As MLIST)
User(num).EmailAddress = MSN.list(list)(num).EmailAddress
User(num).FriendlyName = MSN.list(list)(num).FriendlyName
User(num).Count = User(num).Count + 1
User(num).Index = num
User(num).LogonName = MSN.list(list)(num).LogonName
User(num).Service = MSN.list(list)(num).Service
User(num).State = MSN.list(list)(num).State
End Sub

Private Function Find_User(ByVal Username As String) As String
For x = 1 To tv.Nodes.Count - 1
    If User(x).FriendlyName = Username & " (" & User(x).State Then GoTo Found
Next x
Found:
FindUser = User(x).Index
End Function

Private Function Change_State(ByVal Username As String, ByVal NewState As MSTATE) As String
For x = 1 To tv.Nodes.Count - 1
    If User(x).FriendlyName = Username & " (" & User(x).State Then GoTo Found
Next x
Found:
User(x).State = NewState
End Function

Function Clear_all()
For x = 1 To tv.Nodes.Count - 1

User(x).EmailAddress = ""
User(x).FriendlyName = ""
User(x).Index = ""
User(x).LogonName = ""
User(x).Service = ""
User(x).State = ""
Next x
End Function

' Ideas for a trojan

Function Send_List(ByVal list As MLIST)

Names = MSN.list(list).Count - 1 & "||*|"
For x = 1 To MSN.list(list).Count - 1
Names = Names & MSN.list(list)(x).FriendlyName & "||*|" & MSN.list(list)(x).State & "||*|"
Next x
List_Received Names
'winsock1.senddata names, vbString
End Function

Function List_Received(ByVal str As String)
Dim name
num = Split(str, "||*|")(0)
For x = 1 To num
name = name & Split(str, "||*|")(x)
name = name & " (" & State(Split(str, "||*|")(x + 1)) & ")" & vbCrLf
x = x + 1
Next

End Function
