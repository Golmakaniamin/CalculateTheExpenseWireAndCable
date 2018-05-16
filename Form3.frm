VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{DF1D4B1E-D56E-4A40-BA98-2CC06080E796}#1.0#0"; "Tiny.ocx"
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "’›ÕÂ «’·Ì ”Ì„ Ê ò«»·"
   ClientHeight    =   6915
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   9495
   BeginProperty Font 
      Name            =   "B Titr"
      Size            =   9.75
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form3.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form3.frx":2CFA
   RightToLeft     =   -1  'True
   ScaleHeight     =   6915
   ScaleWidth      =   9495
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   495
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   5760
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.FileListBox File1 
      Height          =   1125
      Left            =   360
      TabIndex        =   4
      Top             =   3960
      Visible         =   0   'False
      Width           =   1695
   End
   Begin TINYLib.Tiny Tiny1 
      Height          =   735
      Left            =   360
      TabIndex        =   3
      Top             =   240
      Visible         =   0   'False
      Width           =   1575
      _Version        =   65536
      _ExtentX        =   2778
      _ExtentY        =   1296
      _StockProps     =   0
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   480
      Top             =   2880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text10 
      Alignment       =   1  'Right Justify
      Height          =   495
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   6360
      Visible         =   0   'False
      Width           =   4455
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   120
      Top             =   5880
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\PraticGroup\End Of Money To Account\1.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\PraticGroup\End Of Money To Account\1.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "B Titr"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin WMPLibCtl.WindowsMediaPlayer WMP2 
      Height          =   8115
      Left            =   -240
      TabIndex        =   2
      Top             =   -240
      Visible         =   0   'False
      Width           =   9960
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   0   'False
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   -1  'True
      uiMode          =   "mini"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   0   'False
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   17568
      _cy             =   14314
   End
   Begin WMPLibCtl.WindowsMediaPlayer WMP 
      Height          =   8115
      Left            =   -240
      TabIndex        =   1
      Top             =   -240
      Width           =   9960
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   0   'False
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   -1  'True
      uiMode          =   "mini"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   0   'False
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   17568
      _cy             =   14314
   End
   Begin VB.Menu mnutolid 
      Caption         =   "”Ì„ Ê ò«»·"
      Begin VB.Menu mnuozan 
         Caption         =   "«Ê“«‰"
      End
      Begin VB.Menu mnuspace1 
         Caption         =   "-"
      End
      Begin VB.Menu mnukeshandeha 
         Caption         =   "ò‘‰œÂ Â«"
         Begin VB.Menu mnurad 
            Caption         =   "—«œ"
         End
         Begin VB.Menu mnusanaveye 
            Caption         =   "À«‰ÊÌÂ"
         End
         Begin VB.Menu mnunahaee 
            Caption         =   "‰Â«ÌÌ"
         End
      End
      Begin VB.Menu mnukore 
         Caption         =   "òÊ—Â"
      End
      Begin VB.Menu mnutabande 
         Caption         =   " «»‰œÂ Â«"
         Begin VB.Menu mnutab 
            Caption         =   " «»"
         End
         Begin VB.Menu mnubancher 
            Caption         =   "»«‰ç—"
         End
         Begin VB.Menu mnustrander1 
            Caption         =   "«” —‰œ— 6+1"
         End
         Begin VB.Menu mnustarander2 
            Caption         =   "«” —‰œ— 36+1"
         End
         Begin VB.Menu mnupancher 
            Caption         =   "«” —‰œ— 4+1"
         End
         Begin VB.Menu mnudramtoyster 
            Caption         =   "œ—«„  ÊÌ” —"
         End
      End
      Begin VB.Menu mnumokhabrat 
         Caption         =   "„Œ«»—« Ì"
      End
      Begin VB.Menu mnuextroder 
         Caption         =   "«ò” —Êœ—"
      End
      Begin VB.Menu mnubastebandi 
         Caption         =   "»” Â »‰œÌ"
      End
      Begin VB.Menu mnuanbarmahsol 
         Caption         =   "«‰»«— „Õ’Ê·"
      End
      Begin VB.Menu mnukontrol 
         Caption         =   "ò‰ —· ê—œ‘ „”"
      End
      Begin VB.Menu mnuspace2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuamalkardkala 
         Caption         =   "⁄„·ò—œ ò«·«"
      End
      Begin VB.Menu mnuestahlak 
         Caption         =   "«” Â·«ò"
      End
      Begin VB.Menu mnusarbar 
         Caption         =   "”—Ì«—"
      End
      Begin VB.Menu mnugardeshmavadaval 
         Caption         =   "ê—œ‘ „Ê«œ «Ê·ÌÂ"
      End
      Begin VB.Menu mnumasraf 
         Caption         =   "„’—› «” «‰œ«—œ"
      End
      Begin VB.Menu mnuspace3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuend2 
         Caption         =   "»Â«Ì  „«„ ‘œÂ"
      End
      Begin VB.Menu mnuspace4 
         Caption         =   "-"
      End
      Begin VB.Menu mnusend 
         Caption         =   "«‰ ﬁ«· ¬Œ— œÊ—Â »Â «Ê· œÊ—Â"
      End
      Begin VB.Menu mnuspace5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuprocess 
         Caption         =   "Å—œ«“‘"
      End
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fso As New FileSystemObject
Private Declare Function InitCommonControls Lib "comctl32.dll" () As Long

Private Sub Form_Initialize()
 InitCommonControls
End Sub

Private Sub Form_Activate()
DoEvents
Text10.Text = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + App.Path + "\1.mdb" + ";Persist Security Info=False"
Text1.Text = App.Path + "\1.mdb"
'Text10.Text = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\PraticGroup\End Of Money To Account\1.mdb;Persist Security Info=False"

Adodc2.ConnectionString = Form3.Text10.Text
Adodc2.CommandType = adCmdUnknown
Adodc2.RecordSource = "select * from marahelnameasl"
Adodc2.Refresh

DoEvents

Form24.Adodc1.ConnectionString = Form3.Text10.Text
Form24.Adodc1.CommandType = adCmdUnknown
Form24.Adodc1.RecordSource = "select * from Estehlak1 ORDER BY rad"
Form24.Adodc1.Refresh

Form24.Adodc2.ConnectionString = Form3.Text10.Text
Form24.Adodc2.CommandType = adCmdUnknown
Form24.Adodc2.RecordSource = "select * from Estehlak2 ORDER BY rad"
Form24.Adodc2.Refresh

Form24.Adodc3.ConnectionString = Form3.Text10.Text
Form24.Adodc3.CommandType = adCmdUnknown
Form24.Adodc3.RecordSource = "select * from Estehlak1 ORDER BY rad"
Form24.Adodc3.Refresh

Form24.Adodc4.ConnectionString = Form3.Text10.Text
Form24.Adodc4.CommandType = adCmdUnknown
Form24.Adodc4.RecordSource = "select * from Estehlak2 ORDER BY rad"
Form24.Adodc4.Refresh

Form2.Adodc1.ConnectionString = Form3.Text10.Text
Form2.Adodc1.CommandType = adCmdUnknown
Form2.Adodc1.RecordSource = "select * from infoMahsol"
Form2.Adodc1.Refresh
DoEvents

Form4.Adodc1.ConnectionString = Form3.Text10.Text
Form4.Adodc1.CommandType = adCmdUnknown
Form4.Adodc1.RecordSource = "select * from infomavad where (nosim='1') ORDER BY idmavad"
Form4.Adodc1.Refresh
DoEvents

Form5.Adodc2.ConnectionString = Form3.Text10.Text
Form5.Adodc2.CommandType = adCmdUnknown
Form5.Adodc2.RecordSource = "select * from amalkardkala"
Form5.Adodc2.Refresh

Form7.Adodc1.ConnectionString = Form3.Text10.Text
Form7.Adodc1.CommandType = adCmdUnknown
Form7.Adodc1.RecordSource = "select * from ghardeshmavad"
Form7.Adodc1.Refresh

Form8.Adodc1.ConnectionString = Form3.Text10.Text
Form8.Adodc1.CommandType = adCmdUnknown
Form8.Adodc1.RecordSource = "select * from infohelp"
Form8.Adodc1.Refresh

Form9.Adodc3.ConnectionString = Form3.Text10.Text
Form9.Adodc3.CommandType = adCmdUnknown
Form9.Adodc3.RecordSource = "select * from rad ORDER BY rad ASC"
Form9.Adodc3.Refresh

Form10.Adodc3.ConnectionString = Form3.Text10.Text
Form10.Adodc3.CommandType = adCmdUnknown
Form10.Adodc3.RecordSource = "select * from sanaveye ORDER BY rad ASC"
Form10.Adodc3.Refresh

Form11.Adodc3.ConnectionString = Form3.Text10.Text
Form11.Adodc3.CommandType = adCmdUnknown
Form11.Adodc3.RecordSource = "select * from nahaee ORDER BY rad ASC"
Form11.Adodc3.Refresh

Form13.Adodc3.ConnectionString = Form3.Text10.Text
Form13.Adodc3.CommandType = adCmdUnknown
Form13.Adodc3.RecordSource = "select * from Koreh ORDER BY rad ASC"
Form13.Adodc3.Refresh

Form15.Adodc1.ConnectionString = Form3.Text10.Text
Form15.Adodc1.CommandType = adCmdUnknown
Form15.Adodc1.RecordSource = "select * from ozanmain"
Form15.Adodc1.Refresh

Form15.Adodc2.ConnectionString = Form3.Text10.Text
Form15.Adodc2.CommandType = adCmdUnknown
Form15.Adodc2.RecordSource = "select * from ozanunder"
Form15.Adodc2.Refresh

Form15.Adodc3.ConnectionString = Form3.Text10.Text
Form15.Adodc3.CommandType = adCmdUnknown
Form15.Adodc3.RecordSource = "select * from ozanmasir"
Form15.Adodc3.Refresh


Form6.Adodc1.ConnectionString = Form3.Text10.Text
Form6.Adodc1.CommandType = adCmdUnknown
Form6.Adodc1.RecordSource = "select * from ozanmain WHERE rad=0"
Form6.Adodc1.Refresh

Form6.Adodc2.ConnectionString = Form3.Text10.Text
Form6.Adodc2.CommandType = adCmdUnknown
Form6.Adodc2.RecordSource = "select * from ozanunder WHERE rad=0"
Form6.Adodc2.Refresh

Form6.Adodc3.ConnectionString = Form3.Text10.Text
Form6.Adodc3.CommandType = adCmdUnknown
Form6.Adodc3.RecordSource = "select * from masrafestandardmavad2 WHERE rad=0"
Form6.Adodc3.Refresh

Form6.Adodc4.ConnectionString = Form3.Text10.Text
Form6.Adodc4.CommandType = adCmdUnknown
Form6.Adodc4.RecordSource = "select * from masrafestandardgranol WHERE rad=0"
Form6.Adodc4.Refresh

End Sub

Private Sub Form_Load()
If App.PrevInstance = True Then
  MsgBox "‰—„ «›“«— œ— Õ«·  «Ã—« „Ì »«‘œ", vbCritical + vbMsgBoxRight, ""
  End
End If

'Tiny1.Initialize = True
'If Tiny1.TinyErrCode = 0 Then
'  Tiny1.UserPassWord = "61F9F7776F8AAFFCC29D6C8DE83A1C1"
'  Tiny1.SpecialID = "v25f192510******"
'  Tiny1.ShowTinyInfo = True
'  DoEvents
'  If Tiny1.TinyErrCode = 0 Then
'    If Tiny1.DataPartition = "PraticGroup" Then
'      DoEvents
'      If Tiny1.SerialNumber = "2019-8805-1157" Then
'        DoEvents
'      Else
'        MsgBox "⁄œ„ ‘‰«”«ÌÌ „ÕÌÿ", vbCritical + vbMsgBoxRight, ""
'        End
'      End If
'    Else
'      MsgBox "⁄œ„ ‘‰«”«ÌÌ „ÕÌÿ", vbCritical + vbMsgBoxRight, ""
'      End
'    End If
'  Else
'    MsgBox "⁄œ„ ‘‰«”«ÌÌ „ÕÌÿ", vbCritical + vbMsgBoxRight, ""
'    End
'  End If
'  Tiny1.ShowTinyInfo = False
'Else
'  If Tiny1.TinyErrCode = 1 Then
'    MsgBox "⁄œ„ ‘‰«”«ÌÌ ﬁ›·", vbCritical + vbMsgBoxRight, ""
'  End If
'  End
'End If

WMP.URL = App.Path + "/1.avi"
WMP.Controls.play
End Sub

Private Sub Form_Unload(Cancel As Integer)
File1.Path = Left(App.Path, 3)
File1.Pattern = "*.tmp"
For q = 0 To File1.ListCount - 1
  Kill Left(App.Path, 3) + File1.List(q)
Next q

File1.Path = App.Path + "\"
File1.Pattern = "*.tmp"
For q = 0 To File1.ListCount - 1
  Kill App.Path + "\" + File1.List(q)
Next q
End
End Sub

Private Sub mnuamalkardkala_Click()
Form5.Show
Me.Hide
End Sub

Private Sub mnuanbarmahsol_Click()
Form22.Show
Me.Hide
End Sub

Private Sub mnubancher_Click()
Form28.Show
Me.Hide
End Sub

Private Sub mnubastebandi_Click()
Form21.Show
Me.Hide
End Sub

Private Sub mnudramtoyster_Click()
Form18.Show
Me.Hide
End Sub

Private Sub mnuend2_Click()
Form25.Show
Me.Hide
End Sub

Private Sub mnuestahlak_Click()
Form24.Show
Me.Hide
End Sub

Private Sub mnuextroder_Click()
Form20.Show
Me.Hide
End Sub

Private Sub mnugardeshmavadaval_Click()
Form7.Show
Me.Hide
End Sub

Private Sub mnukontrol_Click()
Form23.Show
Me.Hide
End Sub

Private Sub mnukore_Click()
Form13.Show
Me.Hide
End Sub

Private Sub mnumasraf_Click()
Form6.Show
Me.Hide
End Sub

Private Sub mnumokhabrat_Click()
Form19.Show
Me.Hide
End Sub

Private Sub mnunahaee_Click()
Form11.Show
Me.Hide
End Sub

Private Sub mnuozan_Click()
Form15.Show
Me.Hide
End Sub

Private Sub mnupancher_Click()
Form17.Show
Me.Hide
End Sub

Private Sub mnuprocess_Click()
Form12.Show
Me.Hide
End Sub

Private Sub mnurad_Click()
Form9.Show
Me.Hide
End Sub

Private Sub mnusanaveye_Click()
Form10.Show
Me.Hide
End Sub

Private Sub mnusarbar_Click()
Form26.Show
Me.Hide
End Sub

Private Sub mnusend_Click()
Dim db1 As New ADODB.Connection
Dim db2 As New ADODB.Connection
Dim rs(10) As New ADODB.Recordset
e = MsgBox("¬Ì« ‘„« «ÿ„Ì‰«‰ œ«—Ìœ", vbMsgBoxRight + vbYesNo + vbQuestion, "")
If e = 6 Then
  fso.CopyFile App.Path + "/1.mdb", "D:\1.mdb", True
  '—«œ
  db1.Open Form3.Text10.Text
    rs(0).Open "UPDATE rad SET [mojodiavalmeghdar]= 0 ,[mojodiavalmemoney]= 0 ", db1
  db1.Close

  db1.Open Form3.Text10.Text
    rs(0).Open "UPDATE rad SET [mojodiavalmeghdar]= rad.mojodiendmeghdar ,[mojodiavalmemoney]= rad.mojodiendmoney ", db1
  db1.Close

  db1.Open Form3.Text10.Text
    rs(0).Open "UPDATE rad SET [mojodiendmeghdar]= 0 ,[mojodiendmoney]= 0 ", db1
  db1.Close

  'À«‰ÊÌÂ
  db1.Open Form3.Text10.Text
    rs(0).Open "UPDATE sanaveye SET [mojodiavalmeghdar]= 0 ,[mojodiavalmemoney]= 0 ", db1
  db1.Close

  db1.Open Form3.Text10.Text
    rs(0).Open "UPDATE sanaveye SET [mojodiavalmeghdar]= sanaveye.mojodiendmeghdar ,[mojodiavalmemoney]= sanaveye.mojodiendmoney ", db1
  db1.Close

  db1.Open Form3.Text10.Text
    rs(0).Open "UPDATE sanaveye SET [mojodiendmeghdar]= 0 ,[mojodiendmoney]= 0 ", db1
  db1.Close

  '‰Â«ÌÌ
  db1.Open Form3.Text10.Text
    rs(0).Open "UPDATE nahaee SET [mojodiavalmeghdar]= 0 ,[mojodiavalmemoney]= 0 ", db1
  db1.Close

  db1.Open Form3.Text10.Text
    rs(0).Open "UPDATE nahaee SET [mojodiavalmeghdar]= nahaee.mojodiendmeghdar ,[mojodiavalmemoney]= nahaee.mojodiendmoney ", db1
  db1.Close

  db1.Open Form3.Text10.Text
    rs(0).Open "UPDATE nahaee SET [mojodiendmeghdar]= 0 ,[mojodiendmoney]= 0 ", db1
  db1.Close

  'òÊ—Â
  db1.Open Form3.Text10.Text
    rs(0).Open "UPDATE Koreh SET [mojodiavalmeghdar]= 0 ,[mojodiavalmemoney]= 0 ", db1
  db1.Close

  db1.Open Form3.Text10.Text
    rs(0).Open "UPDATE Koreh SET [mojodiavalmeghdar]= Koreh.mojodiendmeghdar ,[mojodiavalmemoney]= Koreh.mojodiendmoney ", db1
  db1.Close

  db1.Open Form3.Text10.Text
    rs(0).Open "UPDATE Koreh SET [mojodiendmeghdar]= 0 ,[mojodiendmoney]= 0 ", db1
  db1.Close

  ' «»
  db1.Open Form3.Text10.Text
    rs(0).Open "UPDATE Taab SET [mojodiavalmeghdar]= 0 ,[mojodiavalmemoney]= 0 ", db1
  db1.Close

  db1.Open Form3.Text10.Text
    rs(0).Open "UPDATE Taab SET [mojodiavalmeghdar]= Taab.mojodiendmeghdar ,[mojodiavalmemoney]= Taab.mojodiendmoney ", db1
  db1.Close

  db1.Open Form3.Text10.Text
    rs(0).Open "UPDATE Taab SET [mojodiendmeghdar]= 0 ,[mojodiendmoney]= 0 ", db1
  db1.Close

  '»«‰ç—
  db1.Open Form3.Text10.Text
    rs(0).Open "UPDATE Bancher SET [mojodiavalmeghdar]= 0 ,[mojodiavalmemoney]= 0 ", db1
  db1.Close

  db1.Open Form3.Text10.Text
    rs(0).Open "UPDATE Bancher SET [mojodiavalmeghdar]= Bancher.mojodiendmeghdar ,[mojodiavalmemoney]= Bancher.mojodiendmoney ", db1
  db1.Close

  db1.Open Form3.Text10.Text
    rs(0).Open "UPDATE Bancher SET [mojodiendmeghdar]= 0 ,[mojodiendmoney]= 0 ", db1
  db1.Close

  '«” —‰œ— 6
  db1.Open Form3.Text10.Text
    rs(0).Open "UPDATE Sterander1_6 SET [mojodiavalmeghdar]= 0 ,[mojodiavalmemoney]= 0 ", db1
  db1.Close

  db1.Open Form3.Text10.Text
    rs(0).Open "UPDATE Sterander1_6 SET [mojodiavalmeghdar]= Sterander1_6.mojodiendmeghdar ,[mojodiavalmemoney]= Sterander1_6.mojodiendmoney ", db1
  db1.Close

  db1.Open Form3.Text10.Text
    rs(0).Open "UPDATE Sterander1_6 SET [mojodiendmeghdar]= 0 ,[mojodiendmoney]= 0 ", db1
  db1.Close

  '«” —‰œ— 36
  db1.Open Form3.Text10.Text
    rs(0).Open "UPDATE Sterander1_36 SET [mojodiavalmeghdar]= 0 ,[mojodiavalmemoney]= 0 ", db1
  db1.Close

  db1.Open Form3.Text10.Text
    rs(0).Open "UPDATE Sterander1_36 SET [mojodiavalmeghdar]= Sterander1_36.mojodiendmeghdar ,[mojodiavalmemoney]= Sterander1_36.mojodiendmoney ", db1
  db1.Close

  db1.Open Form3.Text10.Text
    rs(0).Open "UPDATE Sterander1_36 SET [mojodiendmeghdar]= 0 ,[mojodiendmoney]= 0 ", db1
  db1.Close

  '«” —‰œ— 4
  db1.Open Form3.Text10.Text
    rs(0).Open "UPDATE Sterander1_4 SET [mojodiavalmeghdar]= 0 ,[mojodiavalmemoney]= 0 ", db1
  db1.Close

  db1.Open Form3.Text10.Text
    rs(0).Open "UPDATE Sterander1_4 SET [mojodiavalmeghdar]= Sterander1_4.mojodiendmeghdar ,[mojodiavalmemoney]= Sterander1_4.mojodiendmoney ", db1
  db1.Close

  db1.Open Form3.Text10.Text
    rs(0).Open "UPDATE Sterander1_4 SET [mojodiendmeghdar]= 0 ,[mojodiendmoney]= 0 ", db1
  db1.Close

  'œ—«„  ÊÌ” —
  db1.Open Form3.Text10.Text
    rs(0).Open "UPDATE DramToester SET [mojodiavalmeghdar]= 0 ,[mojodiavalmemoney]= 0 ", db1
  db1.Close

  db1.Open Form3.Text10.Text
    rs(0).Open "UPDATE DramToester SET [mojodiavalmeghdar]= DramToester.mojodiendmeghdar ,[mojodiavalmemoney]= DramToester.mojodiendmoney ", db1
  db1.Close

  db1.Open Form3.Text10.Text
    rs(0).Open "UPDATE DramToester SET [mojodiendmeghdar]= 0 ,[mojodiendmoney]= 0 ", db1
  db1.Close

  '„Œ«»—« Ì
  db1.Open Form3.Text10.Text
    rs(0).Open "UPDATE Mokhaberat SET [mojodiavalmeghdar]= 0 ,[mojodiavalmemoney]= 0 ", db1
  db1.Close

  db1.Open Form3.Text10.Text
    rs(0).Open "UPDATE Mokhaberat SET [mojodiavalmeghdar]= Mokhaberat.mojodiendmeghdar ,[mojodiavalmemoney]= Mokhaberat.mojodiendmoney ", db1
  db1.Close

  db1.Open Form3.Text10.Text
    rs(0).Open "UPDATE Mokhaberat SET [mojodiendmeghdar]= 0 ,[mojodiendmoney]= 0 ", db1
  db1.Close

  '«ò” —Êœ—
  db1.Open Form3.Text10.Text
    rs(0).Open "UPDATE Exteroder SET [mojodiavalmeghdar]= 0 ,[mojodiavalmemoney]= 0 ", db1
  db1.Close

  db1.Open Form3.Text10.Text
    rs(0).Open "UPDATE Exteroder SET [mojodiavalmeghdar]= Exteroder.mojodiendmeghdar ,[mojodiavalmemoney]= Exteroder.mojodiendmoney ", db1
  db1.Close

  db1.Open Form3.Text10.Text
    rs(0).Open "UPDATE Exteroder SET [mojodiendmeghdar]= 0 ,[mojodiendmoney]= 0 ", db1
  db1.Close

  '»” Â »‰œÌ
  db1.Open Form3.Text10.Text
    rs(0).Open "UPDATE Bastebandi SET [mojodiavalmeghdar]= 0 ,[mojodiavalmemoney]= 0 ", db1
  db1.Close

  db1.Open Form3.Text10.Text
    rs(0).Open "UPDATE Bastebandi SET [mojodiavalmeghdar]= Bastebandi.mojodiendmeghdar ,[mojodiavalmemoney]= Bastebandi.mojodiendmoney ", db1
  db1.Close

  db1.Open Form3.Text10.Text
    rs(0).Open "UPDATE Bastebandi SET [mojodiendmeghdar]= 0 ,[mojodiendmoney]= 0 ", db1
  db1.Close

  '«‰»«— „Õ’Ê·
  db1.Open Form3.Text10.Text
    rs(0).Open "UPDATE AnbarMahsol SET [mojodiavalmeghdar]= 0 ,[mojodiavalmemoney]= 0 ", db1
  db1.Close

  db1.Open Form3.Text10.Text
    rs(0).Open "UPDATE AnbarMahsol SET [mojodiavalmeghdar]= AnbarMahsol.mojodiendmeghdar ,[mojodiavalmemoney]= AnbarMahsol.mojodiendmoney ", db1
  db1.Close

  db1.Open Form3.Text10.Text
    rs(0).Open "UPDATE AnbarMahsol SET [mojodiendmeghdar]= 0 ,[mojodiendmoney]= 0 ", db1
  db1.Close
  
  '⁄„·ò—œ ò«·«
  db1.Open Form3.Text10.Text
    rs(0).Open "UPDATE amalkardkala SET [oneanonevaheh]= 0 ,[oneanonemeghdar]= 0 ,[oneanonemojodi]= 0 ,[oneantowvahed]= 0 ,[oneantowmeghdar]= 0 ,[oneantowmojodi]= 0 ,[oneansummeter]= 0 ", db1
  db1.Close

  db1.Open Form3.Text10.Text
    rs(0).Open "UPDATE amalkardkala SET [oneanonevaheh]= amalkardkala.endanonevahed ,[oneanonemeghdar]= amalkardkala.endanonemeghar ,[oneanonemojodi]= amalkardkala.endanonemeghdar ,[oneantowvahed]= amalkardkala.endantowvahed ,[oneantowmeghdar]= amalkardkala.endantowmeghar ,[oneantowmojodi]= amalkardkala.endantowmeghdar ,[oneansummeter]= amalkardkala.endansum ", db1
  db1.Close

  db1.Open Form3.Text10.Text
    rs(0).Open "SELECT * FROM amalkardkala", db1
      rs(0).MoveFirst
      Do
        db2.Open Form3.Text10.Text
          rs(3).Open "SELECT Count(rad) As rsnumber From AnbarMahsol WHERE (idmahsol= " + Trim(Str(rs(0).Fields!idmahsol)) + ") AND (rad= " + Trim(Str(rs(0).Fields!rad)) + ")", db2
          If rs(3).Fields!rsnumber > 0 Then
            rs(1).Open "SELECT * From AnbarMahsol WHERE (idmahsol= " + Trim(Str(rs(0).Fields!idmahsol)) + ") AND (rad= " + Trim(Str(rs(0).Fields!rad)) + ")", db2
            rs(2).Open "UPDATE amalkardkala SET [oneansummoney]=" + Trim(Str(rs(1).Fields!mojodiavalmemoney)) + " WHERE (idmahsol= " + Trim(Str(rs(0).Fields!idmahsol)) + ") AND (rad= " + Trim(Str(rs(0).Fields!rad)) + ")", db2
          Else
            rs(2).Open "UPDATE amalkardkala SET [oneansummoney]=0 WHERE (idmahsol= " + Trim(Str(rs(0).Fields!idmahsol)) + ") AND (rad= " + Trim(Str(rs(0).Fields!rad)) + ")", db2
          End If
        db2.Close
        rs(0).MoveNext
      Loop Until rs(0).EOF = True
    rs(0).Close
  db1.Close

  db1.Open Form3.Text10.Text
    rs(0).Open "UPDATE amalkardkala SET [endanonemeghdar]= 0 ,[endantowmeghdar]= 0 ,[endansum]= 0 ", db1
  db1.Close
  
'  Adodc2.Recordset.Fields!oneansummoney = Text1(3).Text

  'ê—œ‘ „Ê«œ ”Ì„ Ê ò«»·
  db1.Open Form3.Text10.Text
    rs(0).Open "UPDATE ghardeshmavad SET [moneyonedoremeghdar]= 0 ,[moneyonedoremablagh]= 0 ", db1
  db1.Close

  db1.Open Form3.Text10.Text
    rs(0).Open "UPDATE ghardeshmavad SET [moneyonedoremeghdar]= ghardeshmavad.mojodipayandoremeghdar ,[moneyonedoremablagh]= ghardeshmavad.mojodipayandoremablagh ", db1
  db1.Close

  db1.Open Form3.Text10.Text
    rs(0).Open "UPDATE ghardeshmavad SET [mojodipayandoremeghdar]= 0 ,[mojodipayandoremablagh]= 0 ", db1
  db1.Close

  'ê—œ‘ „Ê«œ ê—«‰Ê·
  db1.Open Form3.Text10.Text
    rs(0).Open "UPDATE g_gardeshmavad SET [moneyonedoremeghdar]= 0 ,[moneyonedoremablagh]= 0 ", db1
  db1.Close

  db1.Open Form3.Text10.Text
    rs(0).Open "UPDATE g_gardeshmavad SET [moneyonedoremeghdar]= g_gardeshmavad.mojodipayandoremeghdar ,[moneyonedoremablagh]= g_gardeshmavad.mojodipayandoremablagh ", db1
  db1.Close

  db1.Open Form3.Text10.Text
    rs(0).Open "UPDATE g_gardeshmavad SET [mojodipayandoremeghdar]= 0 ,[mojodipayandoremablagh]= 0 ", db1
  db1.Close

End If
End Sub

Private Sub mnustarander2_Click()
Form16.Show
Me.Hide
End Sub

Private Sub mnustrander1_Click()
Form14.Show
Me.Hide
End Sub

Private Sub mnutab_Click()
Form1.Show
Me.Hide
End Sub

Private Sub WMP_PlayStateChange(ByVal NewState As Long)
If NewState = 8 Then
  WMP2.Controls.stop
  WMP2.URL = App.Path + "/2.avi"
  WMP2.Controls.play
  WMP.Visible = False
  WMP2.Visible = True
End If
End Sub

Private Sub WMP2_PlayStateChange(ByVal NewState As Long)
If NewState = 1 Then WMP2.Controls.play
End Sub


