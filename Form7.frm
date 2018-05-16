VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form7 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ÑÏÔ ãæÇÏ Çæáíå"
   ClientHeight    =   8205
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   14550
   BeginProperty Font 
      Name            =   "B Titr"
      Size            =   9.75
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form7.frx":0000
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   8205
   ScaleWidth      =   14550
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "ÇÝÒæÏä ãæÇÏ ÌÏíÏ"
      Height          =   495
      Left            =   5400
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   120
      Width           =   2295
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form7.frx":2CFA
      Height          =   7335
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   14295
      _ExtentX        =   25215
      _ExtentY        =   12938
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   1
      RowHeight       =   23
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
      ColumnCount     =   19
      BeginProperty Column00 
         DataField       =   "idmade"
         Caption         =   "˜Ï ãÍÕæá"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "amin_1"
         Caption         =   "äÇã ãÍÕæá"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "moneyonedoremeghdar"
         Caption         =   "ãæÌæÏí Çæá ÏæÑå ãÞÏÇÑ"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "moneyonedoremablagh"
         Caption         =   "ãæÌæÏí Çæá ÏæÑå ãÈáÛ"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "kharidteydoremeghdar"
         Caption         =   "ÎÑíÏ Øí ÏæÑå ãÞÏÇÑ"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "kharidteydoremablagh"
         Caption         =   "ÎÑíÏ Øí ÏæÑå ãÈáÛ"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "naghlazgeranolmeghdar"
         Caption         =   "äÞá ÇÒ ˜ÑÇäæá ÓÇÒí ãÞÏÇÑ"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "naghlazgeranolmablagh"
         Caption         =   "äÞá ÇÒ ˜ÑÇäæá ÓÇÒí ãÈáÛ"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "mojodiamademasrafmeghdar"
         Caption         =   "ÂãÇÏå ÈÑÇí ãÕÑÝ ãÞÏÇÑ"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column09 
         DataField       =   "mojodiamademasraffi"
         Caption         =   "ÂãÇÏå ÈÑÇí ãÕÑÝ Ýí"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column10 
         DataField       =   "mojodiamademasrafmablagh"
         Caption         =   "ÂãÇÏå ÈÑÇí ãÕÑÝ ãÈáÛ"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column11 
         DataField       =   "masrafteydoremeghdar"
         Caption         =   "ãÕÑÝ Øí ÏæÑå ãÞÏÇÑ"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column12 
         DataField       =   "masrafteydoremablagh"
         Caption         =   "ãÕÑÝ Øí ÏæÑå ãÈáÛ"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column13 
         DataField       =   "foroshteydoremeghdar"
         Caption         =   "ÝÑæÔ Øí ÏæÑå ãÞÏÇÑ"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column14 
         DataField       =   "foroshteydoremablagh"
         Caption         =   "ÝÑæÔ Øí ÏæÑå ãÈáÛ"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column15 
         DataField       =   "zayeatmeghdar"
         Caption         =   "ÖÇíÚÇÊ ãÞÏÇÑ"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column16 
         DataField       =   "zayeatmablagh"
         Caption         =   "ÖÇíÚÇÊ ãÈáÛ"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column17 
         DataField       =   "mojodipayandoremeghdar"
         Caption         =   "ÇíÇä ÏæÑå ãÞÏÇÑ"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column18 
         DataField       =   "mojodipayandoremablagh"
         Caption         =   "ÇíÇä ÏæÑå ãÈáÛ"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2775.118
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2835.213
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   2805.166
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   2775.118
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   2775.118
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   2910.047
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   2880
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   3479.811
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   2775.118
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   3449.764
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   2789.858
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   2775.118
         EndProperty
         BeginProperty Column13 
            ColumnWidth     =   2775.118
         EndProperty
         BeginProperty Column14 
            ColumnWidth     =   2775.118
         EndProperty
         BeginProperty Column15 
            ColumnWidth     =   2775.118
         EndProperty
         BeginProperty Column16 
            ColumnWidth     =   2775.118
         EndProperty
         BeginProperty Column17 
            ColumnWidth     =   3105.071
         EndProperty
         BeginProperty Column18 
            ColumnWidth     =   3075.024
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Èå ÑæÒ ÑÓÇäí"
      Height          =   495
      Left            =   2760
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   120
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Ç"
      Height          =   495
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   120
      Width           =   2535
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   8160
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\PraticGroup\End Of Money To Account\1.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\PraticGroup\End Of Money To Account\1.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "ghardeshmavad"
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
   Begin VB.OptionButton Option2 
      Alignment       =   1  'Right Justify
      Caption         =   "ãæÇÏ Çæáíå ˜ã˜í"
      Height          =   495
      Left            =   10800
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
   Begin VB.OptionButton Option1 
      Alignment       =   1  'Right Justify
      Caption         =   "ãæÇÏ Çæáíå ãÕÑÝí"
      Height          =   495
      Left            =   12720
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Value           =   -1  'True
      Width           =   1695
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   9480
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\PraticGroup\End Of Money To Account\1.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\PraticGroup\End Of Money To Account\1.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "p_gardeshmavad"
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
   Begin VB.Menu mnutarif 
      Caption         =   "ÊÚÇÑíÝ"
      Begin VB.Menu mnumavadkomaki 
         Caption         =   "ãæÇÏ Çæáíå ˜ã˜í"
      End
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db1 As New ADODB.Connection
Dim rs1(10) As New ADODB.Recordset
Dim rs(10) As New ADODB.Recordset

Private Sub Command1_Click()
Dim db1 As New ADODB.Connection
Dim db2 As New ADODB.Connection
Dim rs(10) As New ADODB.Recordset
db1.Open Form3.Text10.Text
  rs(0).Open "SELECT Count(nosim) As rsumber From infomavad WHERE (nosim='1')", db1
    If rs(0).Fields!rsumber > 0 Then
      rs(1).Open "SELECT * From infomavad WHERE (nosim = '1')", db1
        rs(1).MoveFirst
        Do
          rs(2).Open "SELECT Count(nomade) As rsnumber FROM ghardeshmavad WHERE (nomade=1) AND (idmade=" + Trim(Str(rs(1).Fields!idmavad)) + ")", db1
            If rs(2).Fields!rsnumber = 0 Then
              db2.Open Form3.Text10.Text
                rs(3).Open "INSERT INTO ghardeshmavad (idmade,nomade,amin_1,nogra,moneyonedoremeghdar,moneyonedoremablagh,kharidteydoremeghdar,kharidteydoremablagh,naghlazgeranolmeghdar,naghlazgeranolmablagh,mojodiamademasrafmeghdar,mojodiamademasraffi,mojodiamademasrafmablagh,masrafteydoremeghdar,masrafteydoremablagh,foroshteydoremeghdar,foroshteydoremablagh,zayeatmeghdar,zayeatmablagh,mojodipayandoremeghdar,mojodipayandoremablagh) VALUES (" + Trim(Str(rs(1).Fields!idmavad)) + ",1,'" + Trim(rs(1).Fields!mavad) + "'," + Trim(Str(rs(1).Fields!nogra)) + ",'0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0')", db2
              db2.Close
            End If
          rs(2).Close
          rs(1).MoveNext
        Loop Until rs(1).EOF = True
      rs(1).Close
    End If
  rs(0).Close
db1.Close
End Sub

Private Sub Command2_Click()
Dim db1 As New ADODB.Connection
Dim rs1 As New ADODB.Recordset
db1.Open Form3.Text10.Text
  rs1.Open "DELETE FROM p_gardeshmavad", db1
db1.Close
Adodc1.RecordSource = "SELECT * FROM ghardeshmavad"
Adodc1.Refresh
Adodc1.Recordset.MoveFirst
Do
  Adodc3.Refresh
  Adodc3.Recordset.AddNew
  Adodc3.Recordset.Fields!idmade = Adodc1.Recordset.Fields!idmade
  Adodc3.Recordset.Fields!nomade = Adodc1.Recordset.Fields!nomade
  Adodc3.Recordset.Fields!moneyonedoremeghdar = Adodc1.Recordset.Fields!moneyonedoremeghdar
  Adodc3.Recordset.Fields!moneyonedoremablagh = Adodc1.Recordset.Fields!moneyonedoremablagh
  Adodc3.Recordset.Fields!kharidteydoremeghdar = Adodc1.Recordset.Fields!kharidteydoremeghdar
  Adodc3.Recordset.Fields!kharidteydoremablagh = Adodc1.Recordset.Fields!kharidteydoremablagh
  Adodc3.Recordset.Fields!naghlazgeranolmeghdar = Adodc1.Recordset.Fields!naghlazgeranolmeghdar
  Adodc3.Recordset.Fields!naghlazgeranolmablagh = Adodc1.Recordset.Fields!naghlazgeranolmablagh
  Adodc3.Recordset.Fields!mojodiamademasrafmeghdar = Adodc1.Recordset.Fields!mojodiamademasrafmeghdar
  Adodc3.Recordset.Fields!mojodiamademasraffi = Adodc1.Recordset.Fields!mojodiamademasraffi
  Adodc3.Recordset.Fields!mojodiamademasrafmablagh = Adodc1.Recordset.Fields!mojodiamademasrafmablagh
  Adodc3.Recordset.Fields!masrafteydoremeghdar = Adodc1.Recordset.Fields!masrafteydoremeghdar
  Adodc3.Recordset.Fields!masrafteydoremablagh = Adodc1.Recordset.Fields!masrafteydoremablagh
  Adodc3.Recordset.Fields!foroshteydoremeghdar = Adodc1.Recordset.Fields!foroshteydoremeghdar
  Adodc3.Recordset.Fields!foroshteydoremablagh = Adodc1.Recordset.Fields!foroshteydoremablagh
  Adodc3.Recordset.Fields!zayeatmeghdar = Adodc1.Recordset.Fields!zayeatmeghdar
  Adodc3.Recordset.Fields!zayeatmablagh = Adodc1.Recordset.Fields!zayeatmablagh
  Adodc3.Recordset.Fields!mojodipayandoremeghdar = Adodc1.Recordset.Fields!mojodipayandoremeghdar
  Adodc3.Recordset.Fields!mojodipayandoremablagh = Adodc1.Recordset.Fields!mojodipayandoremablagh
  If Adodc1.Recordset.Fields!nomade = 1 Then
    Form4.Adodc1.Recordset.Find "idmavad=" + Trim(Str(Adodc1.Recordset.Fields!idmade)), , adSearchForward, 1
    Adodc3.Recordset.Fields!Name = Form4.Adodc1.Recordset.Fields!mavad
  End If
  
  If Adodc1.Recordset.Fields!nomade = 2 Then
    Form8.Adodc1.Recordset.Find "idmavad=" + Trim(Str(Adodc1.Recordset.Fields!idmade)), , adSearchForward, 1
    Adodc3.Recordset.Fields!Name = Form8.Adodc1.Recordset.Fields!mavad
  End If
  
  Adodc3.Recordset.Update
  Adodc1.Recordset.MoveNext
Loop Until Adodc1.Recordset.EOF = True

Form44.Show
End Sub

Public Sub Command3_Click()
Adodc1.ConnectionString = Form3.Text10.Text
Adodc1.CommandType = adCmdUnknown
Adodc1.RecordSource = "SELECT * FROM ghardeshmavad WHERE nomade=1 Order by idmade"
Adodc1.Refresh

Adodc1.Recordset.MoveFirst
If Adodc1.Recordset.RecordCount > 0 Then
  Adodc1.Recordset.MoveFirst
  Do
    If Adodc1.Recordset.Fields!nogra = 1 Then
      'ãÓÊÑÈ
      If Adodc1.Recordset.Fields!idmade = 14 Then
        Adodc1.Recordset.Fields!naghlazgeranolmeghdar = 0
        Adodc1.Recordset.Fields!naghlazgeranolmablagh = 0
        db1.Open Form3.Text10.Text
          rs(2).Open "SELECT * FROM infomavad WHERE (nogra='1') AND (mastebach='1')", db1
            rs(2).MoveFirst
            Do
              rs(3).Open "SELECT Count(nomade) As rsnumber FROM g_gardeshmavad WHERE (nomade=2) AND (idmade=" + Trim(Str(rs(2).Fields!idmavad)) + ")", db1
                If rs(3).Fields!rsnumber > 0 Then
                  rs(1).Open "SELECT * FROM g_gardeshmavad WHERE (nomade=2) AND (idmade=" + Trim(Str(rs(2).Fields!idmavad)) + ")", db1
                    Adodc1.Recordset.Fields!naghlazgeranolmeghdar = Val(Adodc1.Recordset.Fields!naghlazgeranolmeghdar) + Val(rs(1).Fields!masrafteydoremeghdar)
                    Adodc1.Recordset.Fields!naghlazgeranolmablagh = Val(Adodc1.Recordset.Fields!naghlazgeranolmablagh) + Val(rs(1).Fields!masrafteydoremablagh)
                  rs(1).Close
                End If
              rs(3).Close
              rs(2).MoveNext
            Loop Until rs(2).EOF = True
          rs(2).Close
        db1.Close

        Adodc1.Recordset.Fields!mojodiamademasrafmeghdar = Round(Val(Adodc1.Recordset.Fields!moneyonedoremeghdar) + Val(Adodc1.Recordset.Fields!kharidteydoremeghdar) + Val(Adodc1.Recordset.Fields!naghlazgeranolmeghdar))
        Adodc1.Recordset.Fields!mojodiamademasrafmablagh = Round(Val(Adodc1.Recordset.Fields!moneyonedoremablagh) + Val(Adodc1.Recordset.Fields!kharidteydoremablagh) + Val(Adodc1.Recordset.Fields!naghlazgeranolmablagh))

        If Adodc1.Recordset.Fields!mojodiamademasrafmeghdar = 0 Then
          Adodc1.Recordset.Fields!mojodiamademasraffi = 0
        Else
          Adodc1.Recordset.Fields!mojodiamademasraffi = Round(Val(Adodc1.Recordset.Fields!mojodiamademasrafmablagh) / Val(Adodc1.Recordset.Fields!mojodiamademasrafmeghdar))
        End If

        Adodc1.Recordset.Fields!masrafteydoremeghdar = Round(Val(Adodc1.Recordset.Fields!mojodiamademasrafmeghdar) - (Val(Adodc1.Recordset.Fields!foroshteydoremeghdar) + Val(Adodc1.Recordset.Fields!zayeatmeghdar) + Val(Adodc1.Recordset.Fields!mojodipayandoremeghdar)))

        If Val(Adodc1.Recordset.Fields!mojodiamademasrafmeghdar) <> 0 Then
          Adodc1.Recordset.Fields!masrafteydoremablagh = Round(Val(Adodc1.Recordset.Fields!mojodiamademasrafmablagh) / Val(Adodc1.Recordset.Fields!mojodiamademasrafmeghdar) * Val(Adodc1.Recordset.Fields!masrafteydoremeghdar))
          Adodc1.Recordset.Fields!foroshteydoremablagh = Round(Val(Adodc1.Recordset.Fields!mojodiamademasrafmablagh) / Val(Adodc1.Recordset.Fields!mojodiamademasrafmeghdar) * Val(Adodc1.Recordset.Fields!foroshteydoremeghdar))
          Adodc1.Recordset.Fields!zayeatmablagh = Round(Val(Adodc1.Recordset.Fields!mojodiamademasrafmablagh) / Val(Adodc1.Recordset.Fields!mojodiamademasrafmeghdar) * Val(Adodc1.Recordset.Fields!zayeatmeghdar))
          Adodc1.Recordset.Fields!mojodipayandoremablagh = Round(Val(Adodc1.Recordset.Fields!mojodiamademasrafmablagh) / Val(Adodc1.Recordset.Fields!mojodiamademasrafmeghdar) * Val(Adodc1.Recordset.Fields!mojodipayandoremeghdar))
        Else
          Adodc1.Recordset.Fields!masrafteydoremablagh = 0
          Adodc1.Recordset.Fields!foroshteydoremablagh = 0
          Adodc1.Recordset.Fields!zayeatmablagh = 0
          Adodc1.Recordset.Fields!mojodipayandoremablagh = 0
        End If
        Adodc1.Recordset.Update
        
      Else
      
        db1.Open Form3.Text10.Text
          rs(0).Open "SELECT SUM(meghdar) as rssum FROM masrafestandardmavad2 WHERE (idmade='" + Trim(Str(Adodc1.Recordset.Fields!idmade)) + "')", db1
            If IsNull(rs(0).Fields!rssum) = False Then
              Adodc1.Recordset.Fields!masrafteydoremeghdar = Round(rs(0).Fields!rssum)
            Else
              Adodc1.Recordset.Fields!masrafteydoremeghdar = 0
            End If
          rs(0).Close
        db1.Close
    
        Adodc1.Recordset.Fields!mojodiamademasrafmeghdar = Round(Val(Adodc1.Recordset.Fields!masrafteydoremeghdar) + Val(Adodc1.Recordset.Fields!foroshteydoremeghdar) + Val(Adodc1.Recordset.Fields!zayeatmeghdar) + Val(Adodc1.Recordset.Fields!mojodipayandoremeghdar))
        Adodc1.Recordset.Fields!naghlazgeranolmeghdar = ((Val(Adodc1.Recordset.Fields!masrafteydoremeghdar) + Val(Adodc1.Recordset.Fields!foroshteydoremeghdar) + Val(Adodc1.Recordset.Fields!zayeatmeghdar) + Val(Adodc1.Recordset.Fields!mojodipayandoremeghdar))) - (Val(Adodc1.Recordset.Fields!moneyonedoremeghdar) + Val(Adodc1.Recordset.Fields!kharidteydoremeghdar))
        db1.Open Form3.Text10.Text
          rs(0).Open "SELECT Count(nomade) As rsnumber FROM g_gardeshmavad WHERE (nomade=2) AND (idmade=" + Trim(Str(Adodc1.Recordset.Fields!idmade)) + ")", db1
            If rs(0).Fields!rsnumber > 0 Then
              rs(1).Open "SELECT * FROM g_gardeshmavad WHERE (nomade=2) AND (idmade=" + Trim(Str(Adodc1.Recordset.Fields!idmade)) + ")", db1
                Adodc1.Recordset.Fields!naghlazgeranolmablagh = rs(1).Fields!masrafteydoremablagh
              rs(1).Close
            Else
              Adodc1.Recordset.Fields!naghlazgeranolmablagh = 0
            End If
          rs(0).Close
        db1.Close
    
        Adodc1.Recordset.Fields!mojodiamademasrafmablagh = Round(Val(Adodc1.Recordset.Fields!moneyonedoremablagh) + Val(Adodc1.Recordset.Fields!kharidteydoremablagh) + Val(Adodc1.Recordset.Fields!naghlazgeranolmablagh))
        If Adodc1.Recordset.Fields!mojodiamademasrafmeghdar = 0 Then
          Adodc1.Recordset.Fields!mojodiamademasraffi = 0
        Else
          Adodc1.Recordset.Fields!mojodiamademasraffi = Round(Val(Adodc1.Recordset.Fields!mojodiamademasrafmablagh) / Val(Adodc1.Recordset.Fields!mojodiamademasrafmeghdar))
        End If
    
        If Val(Adodc1.Recordset.Fields!mojodiamademasrafmeghdar) <> 0 Then
          Adodc1.Recordset.Fields!masrafteydoremablagh = Round(Val(Adodc1.Recordset.Fields!mojodiamademasrafmablagh) / Val(Adodc1.Recordset.Fields!mojodiamademasrafmeghdar) * Val(Adodc1.Recordset.Fields!masrafteydoremeghdar))
          Adodc1.Recordset.Fields!foroshteydoremablagh = Round(Val(Adodc1.Recordset.Fields!mojodiamademasrafmablagh) / Val(Adodc1.Recordset.Fields!mojodiamademasrafmeghdar) * Val(Adodc1.Recordset.Fields!foroshteydoremeghdar))
          Adodc1.Recordset.Fields!zayeatmablagh = Round(Val(Adodc1.Recordset.Fields!mojodiamademasrafmablagh) / Val(Adodc1.Recordset.Fields!mojodiamademasrafmeghdar) * Val(Adodc1.Recordset.Fields!zayeatmeghdar))
          Adodc1.Recordset.Fields!mojodipayandoremablagh = Round(Val(Adodc1.Recordset.Fields!mojodiamademasrafmablagh) / Val(Adodc1.Recordset.Fields!mojodiamademasrafmeghdar) * Val(Adodc1.Recordset.Fields!mojodipayandoremeghdar))
        Else
          Adodc1.Recordset.Fields!masrafteydoremablagh = 0
          Adodc1.Recordset.Fields!foroshteydoremablagh = 0
          Adodc1.Recordset.Fields!zayeatmablagh = 0
          Adodc1.Recordset.Fields!mojodipayandoremablagh = 0
        End If
        Adodc1.Recordset.Update
      End If
    Else
      If Adodc1.Recordset.Fields!idmade = 30 Then
        db1.Open Form3.Text10.Text
          rs(2).Open "SELECT Sum(mojodiavalmeghdar) As mojodiavalmeghdar1 ,Sum(mojodiavalmemoney) As mojodiavalmemoney1 ,Sum(tolidteydoremeghdar) As tolidteydoremeghdar1 ,Sum(tolidteydoremoney) As tolidteydoremoney1 ,Sum(naghlbebadmoney) As naghlbebadmoney1 ,Sum(naghlbebadmeghdar) As naghlbebadmeghdar1 ,Sum(mojodiendmeghdar) As mojodiendmeghdar1 ,Sum(mojodiendmoney) As mojodiendmoney1 FROM Newmes", db1
            If IsNull(rs(2).Fields!naghlbebadmeghdar1) = False Then
              Adodc1.Recordset.Fields!moneyonedoremeghdar = rs(2).Fields!mojodiavalmeghdar1
              Adodc1.Recordset.Fields!moneyonedoremablagh = rs(2).Fields!mojodiavalmemoney1
              Adodc1.Recordset.Fields!kharidteydoremeghdar = rs(2).Fields!tolidteydoremeghdar1
              Adodc1.Recordset.Fields!kharidteydoremablagh = rs(2).Fields!tolidteydoremoney1
              Adodc1.Recordset.Fields!naghlazgeranolmeghdar = 0
              Adodc1.Recordset.Fields!naghlazgeranolmablagh = 0
            
              Adodc1.Recordset.Fields!mojodiamademasrafmeghdar = Val(Adodc1.Recordset.Fields!moneyonedoremeghdar) + Val(Adodc1.Recordset.Fields!kharidteydoremeghdar) + Val(Adodc1.Recordset.Fields!naghlazgeranolmeghdar)
              Adodc1.Recordset.Fields!mojodiamademasrafmablagh = Val(Adodc1.Recordset.Fields!moneyonedoremablagh) + Val(Adodc1.Recordset.Fields!kharidteydoremablagh) + Val(Adodc1.Recordset.Fields!naghlazgeranolmablagh)
              If Adodc1.Recordset.Fields!mojodiamademasrafmeghdar <> 0 Then
                Adodc1.Recordset.Fields!mojodiamademasraffi = Round(Val(Adodc1.Recordset.Fields!mojodiamademasrafmablagh) / Val(Adodc1.Recordset.Fields!mojodiamademasrafmeghdar))
              Else
                Adodc1.Recordset.Fields!mojodiamademasraffi = 0
              End If
              Adodc1.Recordset.Fields!masrafteydoremeghdar = rs(2).Fields!naghlbebadmeghdar1
              Adodc1.Recordset.Fields!masrafteydoremablagh = rs(2).Fields!naghlbebadmoney1
              Adodc1.Recordset.Fields!foroshteydoremeghdar = 0
              Adodc1.Recordset.Fields!foroshteydoremablagh = 0
              Adodc1.Recordset.Fields!zayeatmeghdar = 0
              Adodc1.Recordset.Fields!zayeatmablagh = 0
              Adodc1.Recordset.Fields!mojodipayandoremeghdar = rs(2).Fields!mojodiendmeghdar1
              Adodc1.Recordset.Fields!mojodipayandoremablagh = rs(2).Fields!mojodiendmoney1
            Else
              Adodc1.Recordset.Fields!moneyonedoremeghdar = 0
              Adodc1.Recordset.Fields!moneyonedoremablagh = 0
              Adodc1.Recordset.Fields!kharidteydoremeghdar = 0
              Adodc1.Recordset.Fields!kharidteydoremablagh = 0
              Adodc1.Recordset.Fields!naghlazgeranolmeghdar = 0
              Adodc1.Recordset.Fields!naghlazgeranolmablagh = 0
            
              Adodc1.Recordset.Fields!mojodiamademasrafmeghdar = 0
              Adodc1.Recordset.Fields!mojodiamademasrafmablagh = 0
              Adodc1.Recordset.Fields!mojodiamademasraffi = 0
            
              Adodc1.Recordset.Fields!masrafteydoremeghdar = 0
              Adodc1.Recordset.Fields!masrafteydoremablagh = 0
              Adodc1.Recordset.Fields!foroshteydoremeghdar = 0
              Adodc1.Recordset.Fields!foroshteydoremablagh = 0
              Adodc1.Recordset.Fields!zayeatmeghdar = 0
              Adodc1.Recordset.Fields!zayeatmablagh = 0
              Adodc1.Recordset.Fields!mojodipayandoremeghdar = 0
              Adodc1.Recordset.Fields!mojodipayandoremablagh = 0
            End If
            Adodc1.Recordset.Update
          rs(2).Close
        db1.Close
      Else
        Adodc1.Recordset.Fields!naghlazgeranolmeghdar = 0
        Adodc1.Recordset.Fields!naghlazgeranolmablagh = 0
      
        Adodc1.Recordset.Fields!mojodiamademasrafmeghdar = Round(Val(Adodc1.Recordset.Fields!moneyonedoremeghdar) + Val(Adodc1.Recordset.Fields!kharidteydoremeghdar) + Val(Adodc1.Recordset.Fields!naghlazgeranolmeghdar))
        Adodc1.Recordset.Fields!mojodiamademasrafmablagh = Round(Val(Adodc1.Recordset.Fields!moneyonedoremablagh) + Val(Adodc1.Recordset.Fields!kharidteydoremablagh) + Val(Adodc1.Recordset.Fields!naghlazgeranolmablagh))
      
        Adodc1.Recordset.Fields!masrafteydoremeghdar = Val(Adodc1.Recordset.Fields!mojodiamademasrafmeghdar) - (Val(Adodc1.Recordset.Fields!foroshteydoremeghdar) + Val(Adodc1.Recordset.Fields!zayeatmeghdar) + Val(Adodc1.Recordset.Fields!mojodipayandoremeghdar))
        'MsgBox Adodc1.Recordset.Fields!masrafteydoremeghdar
        
        If Adodc1.Recordset.Fields!mojodiamademasrafmeghdar = 0 Then
          Adodc1.Recordset.Fields!mojodiamademasraffi = 0
        Else
          Adodc1.Recordset.Fields!mojodiamademasraffi = Round(Val(Adodc1.Recordset.Fields!mojodiamademasrafmablagh) / Val(Adodc1.Recordset.Fields!mojodiamademasrafmeghdar))
        End If
    
        If Val(Adodc1.Recordset.Fields!mojodiamademasrafmeghdar) <> 0 Then
          Adodc1.Recordset.Fields!masrafteydoremablagh = Round(Val(Adodc1.Recordset.Fields!mojodiamademasrafmablagh) / Val(Adodc1.Recordset.Fields!mojodiamademasrafmeghdar) * Val(Adodc1.Recordset.Fields!masrafteydoremeghdar))
          Adodc1.Recordset.Fields!foroshteydoremablagh = Round(Val(Adodc1.Recordset.Fields!mojodiamademasrafmablagh) / Val(Adodc1.Recordset.Fields!mojodiamademasrafmeghdar) * Val(Adodc1.Recordset.Fields!foroshteydoremeghdar))
          Adodc1.Recordset.Fields!zayeatmablagh = Round(Val(Adodc1.Recordset.Fields!mojodiamademasrafmablagh) / Val(Adodc1.Recordset.Fields!mojodiamademasrafmeghdar) * Val(Adodc1.Recordset.Fields!zayeatmeghdar))
          Adodc1.Recordset.Fields!mojodipayandoremablagh = Round(Val(Adodc1.Recordset.Fields!mojodiamademasrafmablagh) / Val(Adodc1.Recordset.Fields!mojodiamademasrafmeghdar) * Val(Adodc1.Recordset.Fields!mojodipayandoremeghdar))
        Else
          Adodc1.Recordset.Fields!masrafteydoremablagh = 0
          Adodc1.Recordset.Fields!foroshteydoremablagh = 0
          Adodc1.Recordset.Fields!zayeatmablagh = 0
          Adodc1.Recordset.Fields!mojodipayandoremablagh = 0
        End If
        Adodc1.Recordset.Update
      End If
    End If
    Adodc1.Recordset.MoveNext
  Loop Until Adodc1.Recordset.EOF = True
End If
End Sub

Private Sub Form_Activate()
Call Option1_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
Form3.Show
End Sub

Private Sub mnumavadkomaki_Click()
Form8.Show
End Sub

Private Sub Option1_Click()
db1.Open Form3.Text10.Text
  Adodc1.ConnectionString = Form3.Text10.Text
  Adodc1.CommandType = adCmdUnknown
  Adodc1.RecordSource = "SELECT * FROM ghardeshmavad WHERE nomade=1 Order by idmade"
  Adodc1.Refresh

  Adodc1.Recordset.MoveFirst
  Do
    rs(0).Open "SELECT * FROM infomavad WHERE (idmavad=" + Trim(Str(Adodc1.Recordset.Fields!idmade)) + ")", db1
      Adodc1.Recordset.Fields!amin_1 = rs(0).Fields!mavad
      Adodc1.Recordset.Fields!nogra = rs(0).Fields!nogra
      Adodc1.Recordset.Update
    rs(0).Close
    Adodc1.Recordset.MoveNext
  Loop Until Adodc1.Recordset.EOF = True
db1.Close
End Sub

Private Sub Option2_Click()
db1.Open Form3.Text10.Text
  Adodc1.ConnectionString = Form3.Text10.Text
  Adodc1.CommandType = adCmdUnknown
  Adodc1.RecordSource = "SELECT * FROM ghardeshmavad WHERE nomade=2 Order by idmade"
  Adodc1.Refresh

  Adodc1.Recordset.MoveFirst
  Do
    rs(0).Open "SELECT * FROM infohelp WHERE (idmavad=" + Trim(Str(Adodc1.Recordset.Fields!idmade)) + ")", db1
      Adodc1.Recordset.Fields!amin_1 = rs(0).Fields!mavad
      Adodc1.Recordset.Update
    rs(0).Close
    Adodc1.Recordset.MoveNext
  Loop Until Adodc1.Recordset.EOF = True
db1.Close
End Sub
