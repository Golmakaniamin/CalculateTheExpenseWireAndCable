VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Begin VB.Form Form25 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "»Â«Ì  „«„ ‘œÂ"
   ClientHeight    =   8535
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10725
   BeginProperty Font 
      Name            =   "B Titr"
      Size            =   9.75
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form25.frx":0000
   LinkTopic       =   "Form25"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   8535
   ScaleWidth      =   10725
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "ç«Å"
      Height          =   495
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   7920
      Width           =   10455
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
      RecordSource    =   "baha2"
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
   Begin MSDataGridLib.DataGrid DataGrid4 
      Bindings        =   "Form25.frx":2CFA
      Height          =   7695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   13573
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   29
      FormatLocked    =   -1  'True
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
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "name"
         Caption         =   ""
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
         DataField       =   "simocable"
         Caption         =   "”Ì„ Ê ﬂ«»·"
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
      BeginProperty Column02 
         DataField       =   "granol"
         Caption         =   "ê—«‰Ê· ”«“Ì"
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
      BeginProperty Column03 
         DataField       =   "sum"
         Caption         =   "Ã„⁄"
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
            ColumnWidth     =   2729.764
         EndProperty
         BeginProperty Column01 
         EndProperty
         BeginProperty Column02 
         EndProperty
         BeginProperty Column03 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Form25"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim q(22) As String, q1(22) As String
Dim db1 As New ADODB.Connection
Dim rs1 As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset

Private Sub Command1_Click()
Form47.Show
End Sub

Private Sub Form_Activate()
Adodc1.Recordset.Find "rad=1", , adSearchForward, 1
Form7.Adodc1.RecordSource = "Select sum(moneyonedoremablagh) as amin1 From ghardeshmavad Where (nomade=1)"
Form7.Adodc1.Refresh
Adodc1.Recordset.Fields!simocable = Form7.Adodc1.Recordset.Fields!amin1
q(1) = Form7.Adodc1.Recordset.Fields!amin1
Form7.Adodc1.RecordSource = "Select sum(moneyonedoremablagh) as amin1 From g_gardeshmavad Where (nomade=1)"
Form7.Adodc1.Refresh
Adodc1.Recordset.Fields!granol = Form7.Adodc1.Recordset.Fields!amin1
q1(1) = Form7.Adodc1.Recordset.Fields!amin1
Adodc1.Recordset.Fields!Sum = Val(Adodc1.Recordset.Fields!granol) + Val(Adodc1.Recordset.Fields!simocable)
Adodc1.Recordset.Update


Adodc1.Recordset.Find "rad=2", , adSearchForward, 1
Form7.Adodc1.RecordSource = "Select sum(kharidteydoremablagh) as amin1 From ghardeshmavad Where (nomade=1)"
Form7.Adodc1.Refresh
Adodc1.Recordset.Fields!simocable = Form7.Adodc1.Recordset.Fields!amin1
q(2) = Form7.Adodc1.Recordset.Fields!amin1
Form7.Adodc1.RecordSource = "Select sum(kharidteydoremablagh) as amin1 From g_gardeshmavad Where (nomade=1)"
Form7.Adodc1.Refresh
Adodc1.Recordset.Fields!granol = Form7.Adodc1.Recordset.Fields!amin1
q1(2) = Form7.Adodc1.Recordset.Fields!amin1
Adodc1.Recordset.Fields!Sum = Val(Adodc1.Recordset.Fields!granol) + Val(Adodc1.Recordset.Fields!simocable)
Adodc1.Recordset.Update


Adodc1.Recordset.Find "rad=3", , adSearchForward, 1
Form7.Adodc1.RecordSource = "Select sum(naghlazgeranolmablagh) as amin1 From ghardeshmavad Where (nomade=1)"
Form7.Adodc1.Refresh
Adodc1.Recordset.Fields!simocable = Form7.Adodc1.Recordset.Fields!amin1
q(3) = Form7.Adodc1.Recordset.Fields!amin1
Adodc1.Recordset.Fields!granol = 0
q1(3) = 0
Adodc1.Recordset.Fields!Sum = Val(Adodc1.Recordset.Fields!granol) + Val(Adodc1.Recordset.Fields!simocable)
Adodc1.Recordset.Update


Adodc1.Recordset.Find "rad=4", , adSearchForward, 1
Adodc1.Recordset.Fields!simocable = 0
q(4) = 0
db1.Open Adodc1.ConnectionString
  rs1.Open "Select Sum(foroshteydoremablagh) As rssum From g_gardeshmavad Where (nomade=1)", db1
    Adodc1.Recordset.Fields!granol = rs1.Fields!rssum
  rs1.Close
db1.Close
q1(4) = Adodc1.Recordset.Fields!granol
Adodc1.Recordset.Fields!Sum = Val(Adodc1.Recordset.Fields!granol) + Val(Adodc1.Recordset.Fields!simocable)
Adodc1.Recordset.Update


Adodc1.Recordset.Find "rad=5", , adSearchForward, 1
Form26.Adodc4.Recordset.Find "rad=1", , adSearchForward, 1
'Adodc1.Recordset.Fields!simocable = Form26.Adodc4.Recordset.Fields!zayeat
Adodc1.Recordset.Fields!simocable = 0
'q(5) = Form26.Adodc4.Recordset.Fields!zayeat
q(5) = 0
Form7.Adodc1.RecordSource = "Select sum(zayeatmablagh) as amin1 From g_gardeshmavad Where (nomade=1)"
Form7.Adodc1.Refresh
Adodc1.Recordset.Fields!granol = Form7.Adodc1.Recordset.Fields!amin1
q1(5) = Form7.Adodc1.Recordset.Fields!amin1
Adodc1.Recordset.Fields!Sum = Val(Adodc1.Recordset.Fields!granol) + Val(Adodc1.Recordset.Fields!simocable)
Adodc1.Recordset.Update


Adodc1.Recordset.Find "rad=6", , adSearchForward, 1
Form7.Adodc1.RecordSource = "Select sum(mojodipayandoremablagh) as amin1 From ghardeshmavad Where (nomade=1)"
Form7.Adodc1.Refresh
Adodc1.Recordset.Fields!simocable = Form7.Adodc1.Recordset.Fields!amin1
q(6) = Form7.Adodc1.Recordset.Fields!amin1
Form7.Adodc1.RecordSource = "Select sum(mojodipayandoremablagh) as amin1 From g_gardeshmavad Where (nomade=1)"
Form7.Adodc1.Refresh
Adodc1.Recordset.Fields!granol = Form7.Adodc1.Recordset.Fields!amin1
q1(6) = Form7.Adodc1.Recordset.Fields!amin1
Adodc1.Recordset.Fields!Sum = Val(Adodc1.Recordset.Fields!granol) + Val(Adodc1.Recordset.Fields!simocable)
Adodc1.Recordset.Update

Adodc1.Recordset.Find "rad=7", , adSearchForward, 1
db1.Open Adodc1.ConnectionString
rs1.Open "Select * From ghardeshmavad Where (nomade=1) and (idmade=14)", db1
Adodc1.Recordset.Fields!simocable = (Val(q(1)) + Val(q(2)) + Val(q(3)) + Val(q(4)) + Val(q(5)) - Val(q(6))) - rs1.Fields!masrafteydoremablagh
q(7) = Adodc1.Recordset.Fields!simocable
rs1.Close
db1.Close
Adodc1.Recordset.Fields!granol = (Val(q1(1)) + Val(q1(2)) + Val(q1(3))) - (Val(q1(4)) + Val(q1(5)) + Val(q1(6)))
q1(7) = Adodc1.Recordset.Fields!granol
Adodc1.Recordset.Fields!Sum = Val(Adodc1.Recordset.Fields!granol) + Val(Adodc1.Recordset.Fields!simocable)
Adodc1.Recordset.Update

Adodc1.Recordset.Find "rad=8", , adSearchForward, 1
Form26.Adodc4.Recordset.Find "rad=997", , adSearchForward, 1
Adodc1.Recordset.Fields!simocable = Form26.Adodc4.Recordset.Fields!dastmozd
q(8) = Form26.Adodc4.Recordset.Fields!dastmozd
Form26.Adodc4.Recordset.Find "rad=998", , adSearchForward, 1
Adodc1.Recordset.Fields!granol = Form26.Adodc4.Recordset.Fields!dastmozd
q1(8) = Form26.Adodc4.Recordset.Fields!dastmozd
Adodc1.Recordset.Fields!Sum = Val(Adodc1.Recordset.Fields!granol) + Val(Adodc1.Recordset.Fields!simocable)
Adodc1.Recordset.Update


Adodc1.Recordset.Find "rad=9", , adSearchForward, 1
Form26.Adodc4.Recordset.Find "rad=997", , adSearchForward, 1
Adodc1.Recordset.Fields!simocable = Val(Form26.Adodc4.Recordset.Fields!sarbarvahed) + Val(Form26.Adodc4.Recordset.Fields!sarbarjazb)
q(9) = Val(Form26.Adodc4.Recordset.Fields!sarbarvahed) + Val(Form26.Adodc4.Recordset.Fields!sarbarjazb)
Form26.Adodc4.Recordset.Find "rad=998", , adSearchForward, 1
Adodc1.Recordset.Fields!granol = Form26.Adodc4.Recordset.Fields!sarbarvahed
q1(9) = Form26.Adodc4.Recordset.Fields!sarbarvahed
Adodc1.Recordset.Fields!Sum = Val(Adodc1.Recordset.Fields!granol) + Val(Adodc1.Recordset.Fields!simocable)
Adodc1.Recordset.Update


Adodc1.Recordset.Find "rad=10", , adSearchForward, 1
Form26.Adodc4.Recordset.Find "rad=997", , adSearchForward, 1
Adodc1.Recordset.Fields!simocable = Form26.Adodc4.Recordset.Fields!estehlak
q(10) = Form26.Adodc4.Recordset.Fields!estehlak
Form26.Adodc4.Recordset.Find "rad=998", , adSearchForward, 1
Adodc1.Recordset.Fields!granol = Form26.Adodc4.Recordset.Fields!estehlak
q1(10) = Form26.Adodc4.Recordset.Fields!estehlak
Adodc1.Recordset.Fields!Sum = Val(Adodc1.Recordset.Fields!granol) + Val(Adodc1.Recordset.Fields!simocable)
Adodc1.Recordset.Update

Adodc1.Recordset.Find "rad=11", , adSearchForward, 1
Form21.Adodc1.Recordset.Find "rad=99999", , adSearchForward, 1
sd = Form21.Adodc1.Recordset.Fields!baste
Adodc1.Recordset.Fields!simocable = sd
q(21) = Adodc1.Recordset.Fields!simocable
Adodc1.Recordset.Fields!granol = 0
Adodc1.Recordset.Fields!Sum = Val(Adodc1.Recordset.Fields!granol) + Val(Adodc1.Recordset.Fields!simocable)
Adodc1.Recordset.Update

Adodc1.Recordset.Find "rad=17", , adSearchForward, 1
Form22.Adodc1.Recordset.Find "rad=99999", , adSearchForward, 1
Adodc1.Recordset.Fields!simocable = Form22.Adodc1.Recordset.Fields!mojodiavalmemoney
q(14) = Form22.Adodc1.Recordset.Fields!mojodiavalmemoney
Form7.Adodc1.RecordSource = "Select sum(moneyonedoremablagh) as amin1 From g_gardeshmavad Where (nomade=2)"
Form7.Adodc1.Refresh
Adodc1.Recordset.Fields!granol = Form7.Adodc1.Recordset.Fields!amin1
q1(14) = Form7.Adodc1.Recordset.Fields!amin1
Adodc1.Recordset.Fields!Sum = Val(Adodc1.Recordset.Fields!granol) + Val(Adodc1.Recordset.Fields!simocable)
Adodc1.Recordset.Update

Adodc1.Recordset.Find "rad=18", , adSearchForward, 1
Adodc1.Recordset.Fields!simocable = 0
q(15) = 0
Form7.Adodc1.RecordSource = "Select sum(masrafteydoremablagh) as amin1 From g_gardeshmavad Where (nomade=2)"
Form7.Adodc1.Refresh
Adodc1.Recordset.Fields!granol = Form7.Adodc1.Recordset.Fields!amin1
q1(15) = Form7.Adodc1.Recordset.Fields!amin1
Adodc1.Recordset.Fields!Sum = Val(Adodc1.Recordset.Fields!granol) + Val(Adodc1.Recordset.Fields!simocable)
Adodc1.Recordset.Update

Adodc1.Recordset.Find "rad=19", , adSearchForward, 1
Adodc1.Recordset.Fields!simocable = 0
q(16) = 0
Form7.Adodc1.RecordSource = "Select sum(zayeatmablagh) as amin1 From g_gardeshmavad Where (nomade=2)"
Form7.Adodc1.Refresh
Adodc1.Recordset.Fields!granol = Form7.Adodc1.Recordset.Fields!amin1
q1(16) = Form7.Adodc1.Recordset.Fields!amin1
Adodc1.Recordset.Fields!Sum = Val(Adodc1.Recordset.Fields!granol) + Val(Adodc1.Recordset.Fields!simocable)
Adodc1.Recordset.Update

Adodc1.Recordset.Find "rad=20", , adSearchForward, 1
Form22.Adodc1.Recordset.Find "rad=99999", , adSearchForward, 1
Adodc1.Recordset.Fields!simocable = Form22.Adodc1.Recordset.Fields!mojodiendmoney
q(17) = Form22.Adodc1.Recordset.Fields!mojodiendmoney
Form7.Adodc1.RecordSource = "Select sum(mojodipayandoremablagh) as amin1 From g_gardeshmavad Where (nomade=2)"
Form7.Adodc1.Refresh
Adodc1.Recordset.Fields!granol = Form7.Adodc1.Recordset.Fields!amin1
q1(17) = Form7.Adodc1.Recordset.Fields!amin1
Adodc1.Recordset.Fields!Sum = Val(Adodc1.Recordset.Fields!granol) + Val(Adodc1.Recordset.Fields!simocable)
Adodc1.Recordset.Update

Adodc1.Recordset.Find "rad=12", , adSearchForward, 1
Form26.Adodc4.Recordset.Find "rad=997", , adSearchForward, 1
Adodc1.Recordset.Fields!simocable = Val(Form26.Adodc4.Recordset.Fields!kaladarjaryanavaldore) - Val(q(14))
q(11) = Val(Form26.Adodc4.Recordset.Fields!kaladarjaryanavaldore) - Val(q(14))
Adodc1.Recordset.Fields!granol = 0
q1(11) = 0
Adodc1.Recordset.Fields!Sum = Val(Adodc1.Recordset.Fields!granol) + Val(Adodc1.Recordset.Fields!simocable)
Adodc1.Recordset.Update


Adodc1.Recordset.Find "rad=13", , adSearchForward, 1
Form26.Adodc4.Recordset.Find "rad=997", , adSearchForward, 1
Adodc1.Recordset.Fields!simocable = Val(Form26.Adodc4.Recordset.Fields!hazkalapayandore) - Val(q(17))
q(12) = Val(Form26.Adodc4.Recordset.Fields!hazkalapayandore) - Val(q(17))
Adodc1.Recordset.Fields!granol = 0
q1(12) = 0
Adodc1.Recordset.Fields!Sum = Val(Adodc1.Recordset.Fields!granol) + Val(Adodc1.Recordset.Fields!simocable)
Adodc1.Recordset.Update


Adodc1.Recordset.Find "rad=14", , adSearchForward, 1
Adodc1.Recordset.Fields!simocable = (Val(q(7)) + Val(q(8)) + Val(q(9)) + Val(q(10)) + Val(q(11)) + Val(q(21))) - Val(q(12))
q(13) = Adodc1.Recordset.Fields!simocable
Adodc1.Recordset.Fields!granol = (Val(q1(7)) + Val(q1(8)) + Val(q1(9)) + Val(q1(10)) + Val(q1(11))) - Val(q1(12))
q1(13) = (Val(q1(7)) + Val(q1(8)) + Val(q1(9)) + Val(q1(10)) + Val(q1(11))) - Val(q1(12))
Adodc1.Recordset.Fields!Sum = Val(Adodc1.Recordset.Fields!granol) + Val(Adodc1.Recordset.Fields!simocable)
Adodc1.Recordset.Update

Adodc1.Recordset.Find "rad=15", , adSearchForward, 1
Form9.Adodc1.Recordset.Find "rad=99998", , adSearchForward, 1
Adodc1.Recordset.Fields!simocable = Form9.Adodc1.Recordset.Fields!gheymattamam
q(18) = Adodc1.Recordset.Fields!simocable
Adodc1.Recordset.Fields!granol = 0
Adodc1.Recordset.Fields!Sum = Val(Adodc1.Recordset.Fields!granol) + Val(Adodc1.Recordset.Fields!simocable)
Adodc1.Recordset.Update

Adodc1.Recordset.Find "rad=16", , adSearchForward, 1
Adodc1.Recordset.Fields!simocable = (Val(q(13)) - Val(q(18)))
q(19) = Adodc1.Recordset.Fields!simocable
Adodc1.Recordset.Fields!granol = 0
Adodc1.Recordset.Fields!Sum = Val(Adodc1.Recordset.Fields!granol) + Val(Adodc1.Recordset.Fields!simocable)
Adodc1.Recordset.Update

Adodc1.Recordset.Find "rad=21", , adSearchForward, 1
Adodc1.Recordset.Fields!simocable = Val(q(14)) + Val(q(19)) - Val(q(17))

q(22) = Adodc1.Recordset.Fields!simocable
Adodc1.Recordset.Fields!granol = (Val(q1(14)) + Val(q1(13))) - (Val(q1(15)) + Val(q1(16)) + Val(q1(17)))

q1(22) = Adodc1.Recordset.Fields!granol
Adodc1.Recordset.Fields!Sum = Val(Adodc1.Recordset.Fields!granol) + Val(Adodc1.Recordset.Fields!simocable)
Adodc1.Recordset.Update


Adodc1.Recordset.Find "rad=22", , adSearchForward, 1
db1.Open Adodc1.ConnectionString
  rs1.Open "Select * From AnbarMahsol Where (rad=99999)", db1
    Adodc1.Recordset.Fields!simocable = Val(rs1.Fields!naghlbebadmoney) - Val(q(22))
  rs1.Close
db1.Close

db1.Open Adodc1.ConnectionString
  rs1.Open "Select Sum(foroshteydoremablagh) As rssum From g_gardeshmavad Where (nomade=2)", db1
    Adodc1.Recordset.Fields!granol = Val(rs1.Fields!rssum) - Val(q1(22))
  rs1.Close
db1.Close

Adodc1.Recordset.Fields!Sum = Val(Adodc1.Recordset.Fields!granol) + Val(Adodc1.Recordset.Fields!simocable)
Adodc1.Recordset.Update

Adodc1.Recordset.Sort = "rad"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Form3.Show
Me.Hide
End Sub


