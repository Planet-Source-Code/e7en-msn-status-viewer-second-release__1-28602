VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   0  'None
   ClientHeight    =   3090
   ClientLeft      =   5445
   ClientTop       =   4050
   ClientWidth     =   4440
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form2.frx":0000
   ScaleHeight     =   3090
   ScaleWidth      =   4440
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   2190
      Left            =   130
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   240
      Width           =   4095
   End
   Begin VB.Image Image3 
      Height          =   390
      Left            =   240
      Picture         =   "Form2.frx":4EF1A
      Top             =   2640
      Width           =   1560
   End
   Begin VB.Image Image2 
      Height          =   360
      Left            =   2520
      Picture         =   "Form2.frx":50F0C
      Top             =   2655
      Width           =   1560
   End
   Begin VB.Image Image1 
      Height          =   2595
      Left            =   0
      Picture         =   "Form2.frx":52C8E
      Top             =   10
      Width           =   4455
   End
   Begin VB.Menu status 
      Caption         =   "Status"
      Visible         =   0   'False
      Begin VB.Menu online 
         Caption         =   "Online"
         Checked         =   -1  'True
      End
      Begin VB.Menu busy 
         Caption         =   "Busy"
         Checked         =   -1  'True
      End
      Begin VB.Menu brb 
         Caption         =   "Be Right Back"
         Checked         =   -1  'True
      End
      Begin VB.Menu awak 
         Caption         =   "Away"
         Checked         =   -1  'True
      End
      Begin VB.Menu otf 
         Caption         =   "On The Phone"
         Checked         =   -1  'True
      End
      Begin VB.Menu otl 
         Caption         =   "Out To Lunch"
         Checked         =   -1  'True
      End
      Begin VB.Menu ao 
         Caption         =   "Appear Offline"
         Checked         =   -1  'True
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ao_Click()
DoEvents
ao.Checked = True
awak.Checked = False
brb.Checked = False
busy.Checked = False
online.Checked = False
otf.Checked = False
otl.Checked = False
Form1.MSN.LocalState = 6
Form1.im1.Picture = Form1.img.ListImages(8).Picture
End Sub

Private Sub awak_Click()
DoEvents
ao.Checked = False
awak.Checked = True
brb.Checked = False
busy.Checked = False
online.Checked = False
otf.Checked = False
otl.Checked = False
Form1.MSN.LocalState = 34
Form1.im1.Picture = Form1.img.ListImages(6).Picture
End Sub

Private Sub brb_Click()
DoEvents
ao.Checked = False
awak.Checked = False
brb.Checked = True
busy.Checked = False
online.Checked = False
otf.Checked = False
otl.Checked = False
Form1.MSN.LocalState = 14
Form1.im1.Picture = Form1.img.ListImages(6).Picture
End Sub

Private Sub busy_Click()
DoEvents
ao.Checked = False
awak.Checked = False
brb.Checked = False
busy.Checked = True
online.Checked = False
otf.Checked = False
otl.Checked = False
Form1.MSN.LocalState = 10
Form1.im1.Picture = Form1.img.ListImages(7).Picture
End Sub

Private Sub Combo1_Click()
If Combo1.Text = Combo1.list(1) Then Load_List (1)
If Combo1.Text = Combo1.list(2) Then Load_List (3)
If Combo1.Text = Combo1.list(0) Then Load_List (2)
Me.Caption = "Viewing '" & Combo1.Text & "'"
End Sub

Private Sub Command1_Click()
On Error GoTo error
If Form1.MSN.LocalState = MSTATE_OFFLINE Then GoTo error
Form1.MSN.Services.PrimaryService.FriendlyName = Text1.Text
MsgBox "Name Changed Succesfully!", vbApplicationModal + vbInformation, "MSN Name Changer"
Exit Sub
error:
MsgBox "Error Writeing Name!", vbApplicationModal + vbCritical, "MSN Name Changer"
Text1.Text = Form1.MSN.LocalFriendlyName
End Sub

Private Sub Command2_Click()
Form1.Show
Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next
Text1.Text = Form1.MSN.LocalFriendlyName
End Sub

Private Sub Image2_Click()
Command2_Click
End Sub

Private Sub Image3_Click()
Command1_Click
End Sub

Private Sub online_Click()
DoEvents
ao.Checked = False
awak.Checked = False
brb.Checked = False
busy.Checked = False
online.Checked = True
otf.Checked = False
otl.Checked = False
Form1.MSN.LocalState = 2
Form1.im1.Picture = Form1.img.ListImages(5).Picture
End Sub

Private Sub otf_Click()
DoEvents
ao.Checked = False
awak.Checked = False
brb.Checked = False
busy.Checked = False
online.Checked = False
otf.Checked = True
otl.Checked = False
Form1.MSN.LocalState = 50
Form1.im1.Picture = Form1.img.ListImages(7).Picture
End Sub

Private Sub otl_Click()
DoEvents
ao.Checked = False
awak.Checked = False
brb.Checked = False
busy.Checked = False
online.Checked = False
otf.Checked = False
otl.Checked = True
Form1.MSN.LocalState = 66
Form1.im1.Picture = Form1.img.ListImages(6).Picture
End Sub

Function Load_List(num As Integer)
List1.Clear
x = Form1.MSN.list(num).Count
For P = 1 To x - 1
List1.AddItem Form1.MSN.list(num)(P).FriendlyName
Next P
If List1.ListCount = 0 Then List1.AddItem "None."
End Function


