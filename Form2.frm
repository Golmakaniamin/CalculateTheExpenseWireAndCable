VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ÊÚÑíÝ ãÍÕæá"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4560
   BeginProperty Font 
      Name            =   "B Titr"
      Size            =   9.75
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   5175
   ScaleWidth      =   4560
   StartUpPosition =   2  'CenterScreen
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form2.frx":2CFA
      Height          =   3255
      Left            =   120
      TabIndex        =   8
      Top             =   1800
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   5741
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   26
      FormatLocked    =   -1  'True
      AllowDelete     =   -1  'True
      RightToLeft     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "B Titr"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "B Titr"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "idmahsol"
         Caption         =   "ßÏ ãÍÕæá"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1080
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "mahsol"
         Caption         =   "äÇã ãÍÕæá"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1080
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   120
      Top             =   120
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=1.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=1.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "infoMahsol"
      Caption         =   "Adodc1"
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
   Begin VB.CommandButton Command3 
      Caption         =   "ËÈÊ"
      Height          =   375
      Left            =   3120
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "æíÑÇíÔ"
      Height          =   375
      Left            =   3120
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ÌÏíÏ"
      Height          =   375
      Left            =   3120
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   240
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   495
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Label4"
      Height          =   255
      Left            =   1440
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   1440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "äÇã ãÍÕæá"
      Height          =   495
      Left            =   1560
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "˜Ï ãÍÕæá :"
      Height          =   495
      Left            =   1560
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim q As Integer

Private Sub Command1_Click()
Label4.Caption = 1
Text1.Text = ""
q = 1
If Adodc1.Recordset.RecordCount > 0 Then
  Adodc1.Recordset.Sort = "idmahsol"
  Adodc1.Recordset.MoveFirst
  Do
    If Adodc1.Recordset.Fields!idmahsol <> q Then Exit Do
    q = q + 1
    Adodc1.Recordset.MoveNext
  Loop Until Adodc1.Recordset.EOF = True
End If
Label2.Caption = q
Text1.SetFocus
End Sub

Private Sub Command2_Click()
If Adodc1.Recordset.RecordCount > 0 Then
  DataGrid1.Col = 0
  Label2.Caption = DataGrid1.Text
  DataGrid1.Col = 1
  Text1.Text = DataGrid1.Text
  Label4.Caption = 2
End If
End Sub

Private Sub Command3_Click()
If (Label2.Caption = "") Or (Text1.Text = "") Then
  MsgBox "áØÝÇ ÊãÇãí ÝíáÏ åÇ ÑÇ Ê˜ãíá äãÇííÏ", vbCritical + vbMsgBoxRight, ""
  Exit Sub
End If
If Label4.Caption = 1 Then
  Adodc1.Refresh
  Adodc1.Recordset.AddNew
  Adodc1.Recordset.Fields!idmahsol = Label2.Caption
  Adodc1.Recordset.Fields!mahsol = Text1.Text
  Adodc1.Recordset.Update
  MsgBox "ÇØáÇÚÇÊ ÈÇ ãæÝÞíÊ ËÈÊ ÔÏ", vbInformation + vbMsgBoxRight, ""
End If
If Label4.Caption = 2 Then
  Adodc1.Refresh
  Adodc1.Recordset.Find "idmahsol=" + Label2.Caption, , adSearchForward, 1
  Adodc1.Recordset.Fields!mahsol = Text1.Text
  Adodc1.Recordset.Update
  MsgBox "ÇØáÇÚÇÊ ÈÇ ãæÝÞíÊ ÊÛííÑ íÏÇ ˜ÑÏ", vbInformation + vbMsgBoxRight, ""
End If
DataGrid1.Refresh
Command1.SetFocus
End Sub

Private Sub Form_Activate()
If Adodc1.Recordset.RecordCount > 0 Then
  Adodc1.Recordset.MoveFirst
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Form15.Show
End Sub

Private Sub Text1_GotFocus()
Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Command3.SetFocus
End Sub
