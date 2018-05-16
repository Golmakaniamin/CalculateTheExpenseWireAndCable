VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form12 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Å—œ«“‘"
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10185
   BeginProperty Font 
      Name            =   "B Titr"
      Size            =   9.75
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form12.frx":0000
   LinkTopic       =   "Form12"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   3120
   ScaleWidth      =   10185
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   120
      Top             =   2040
   End
   Begin VB.CommandButton Command1 
      Caption         =   "‘—Ê⁄"
      BeginProperty Font 
         Name            =   "B Titr"
         Size            =   26.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   9975
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   1508
      _Version        =   393216
      Appearance      =   1
      Max             =   1000
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "0"
      Height          =   495
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "B Titr"
         Size            =   26.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1080
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   2040
      Width           =   9015
   End
End
Attribute VB_Name = "Form12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim sd As String
Dim db1 As New ADODB.Connection
Dim rs1(3) As New ADODB.Recordset

Label2.Caption = 0
Command1.Enabled = False
Timer1.Enabled = True
ProgressBar1.Max = 28
ProgressBar1.Min = 0
ProgressBar1.Value = 0

Label1.Caption = "«Ê“«‰"
DoEvents
Form4.Adodc1.RecordSource = "select * from infomavad"
Form4.Adodc1.Refresh
If Form4.Adodc1.Recordset.RecordCount > 0 Then
  Form4.Adodc1.Recordset.MoveFirst
  Do
    Form15.Adodc2.RecordSource = "select * from ozanunder Where idmade= '" + Trim(Str(Form4.Adodc1.Recordset.Fields!idmavad)) + "'"
    Form15.Adodc2.Refresh
    If Form15.Adodc2.Recordset.RecordCount > 0 Then
      Form15.Adodc2.Recordset.MoveFirst
      Do
        sd = Val(Form15.Adodc2.Recordset.Fields!meghdar) * Val(Form4.Adodc1.Recordset.Fields!zarib)
        Form15.Adodc2.Recordset.Fields!meghdar2 = sd
        Form15.Adodc2.Recordset.Update
        Form15.Adodc2.Recordset.MoveNext
      Loop Until Form15.Adodc2.Recordset.EOF = True
    End If
    Form4.Adodc1.Recordset.MoveNext
  Loop Until Form4.Adodc1.Recordset.EOF = True
End If

ProgressBar1.Value = 1
Label1.Caption = "ê—œ‘ „Ê«œ «Ê·ÌÂ"
DoEvents
Call Form7.Command3_Click

ProgressBar1.Value = 2
Label1.Caption = "«” Â·«ò"
DoEvents
Call Form24.Command1_Click

ProgressBar1.Value = 3
Label1.Caption = "”—»«—"
DoEvents
Call Form26.Command1_Click

ProgressBar1.Value = 4
Label1.Caption = "„’—› «” «‰œ«—œ"
DoEvents
Call Form6.Command1_Click

ProgressBar1.Value = 5
Label1.Caption = "—«œ"
DoEvents
Call Form9.Command1_Click
Call Form9.Command1_Click
Call Form9.Command1_Click

ProgressBar1.Value = 6
Label1.Caption = "À«‰ÊÌÂ"
DoEvents
Call Form10.Command1_Click
Call Form10.Command1_Click
Call Form10.Command1_Click

ProgressBar1.Value = 7
Label1.Caption = "‰Â«ÌÌ"
DoEvents
Call Form11.Command1_Click
Call Form11.Command1_Click
Call Form11.Command1_Click

ProgressBar1.Value = 8
Label1.Caption = "òÊ—Â"
DoEvents
Call Form13.Command1_Click
Call Form13.Command1_Click
Call Form13.Command1_Click

ProgressBar1.Value = 9
Label1.Caption = "‰Â«ÌÌ"
DoEvents
Call Form11.Command1_Click
Call Form11.Command1_Click
Call Form11.Command1_Click

ProgressBar1.Value = 10
Label1.Caption = "À«‰ÊÌÂ"
DoEvents
Call Form10.Command1_Click
Call Form10.Command1_Click
Call Form10.Command1_Click

ProgressBar1.Value = 11
Label1.Caption = "—«œ"
DoEvents
Call Form9.Command1_Click
Call Form9.Command1_Click
Call Form9.Command1_Click

ProgressBar1.Value = 12
Label1.Caption = "À«‰ÊÌÂ"
DoEvents
Call Form10.Command1_Click
Call Form10.Command1_Click
Call Form10.Command1_Click

ProgressBar1.Value = 13
Label1.Caption = "‰Â«ÌÌ"
DoEvents
Call Form11.Command1_Click
Call Form11.Command1_Click
Call Form11.Command1_Click

ProgressBar1.Value = 14
Label1.Caption = "òÊ—Â"
DoEvents
Call Form13.Command1_Click
Call Form13.Command1_Click
Call Form13.Command1_Click

ProgressBar1.Value = 15
Label1.Caption = " «»"
DoEvents
Call Form1.Command1_Click

ProgressBar1.Value = 16
Label1.Caption = "»«‰ç—"
DoEvents
Call Form28.Command1_Click

ProgressBar1.Value = 17
Label1.Caption = "«” —‰œ— 6 +1"
DoEvents
Call Form14.Command1_Click

ProgressBar1.Value = 18
Label1.Caption = "«” —‰œ— 36 + 1"
DoEvents
Call Form16.Command1_Click

ProgressBar1.Value = 19
Label1.Caption = "«” —‰œ— 4 + 1"
DoEvents
Call Form17.Command1_Click

ProgressBar1.Value = 20
Label1.Caption = "œ—«„  ÊÌ” —"
DoEvents
Call Form18.Command1_Click

ProgressBar1.Value = 21
Label1.Caption = "„Œ«»—« Ì"
DoEvents
Call Form19.Command1_Click

ProgressBar1.Value = 22
Label1.Caption = "«ò” —Êœ—"
DoEvents
Call Form20.Command1_Click

ProgressBar1.Value = 23
Label1.Caption = "„’—› «” «‰œ«—œ"
DoEvents
Call Form6.Command1_Click

ProgressBar1.Value = 24
Label1.Caption = "«ò” —Êœ—"
DoEvents
Call Form20.Command1_Click

ProgressBar1.Value = 25
Label1.Caption = "»” Â »‰œÌ"
DoEvents
Call Form21.Command1_Click

ProgressBar1.Value = 26
Label1.Caption = "«‰»«— „Õ’Ê·"
DoEvents
Call Form22.Command1_Click

ProgressBar1.Value = 27
Label1.Caption = "”—»«—"
DoEvents
Call Form26.Command1_Click

ProgressBar1.Value = 28
Label1.Caption = "Å—œ«“‘ „Ê›ﬁÌ  Å–Ì— »Êœ"
DoEvents
Command1.Enabled = True
Timer1.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
Form3.Show
Me.Hide
End Sub

Private Sub Timer1_Timer()
Label2 = Val(Label2) + 1
End Sub
