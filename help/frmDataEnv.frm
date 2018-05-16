VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmDataEnv 
   Caption         =   "Form1"
   ClientHeight    =   4515
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7515
   LinkTopic       =   "Form1"
   ScaleHeight     =   4515
   ScaleWidth      =   7515
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   5160
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   3120
      TabIndex        =   1
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   3000
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   2040
      Width           =   2055
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   735
      Left            =   720
      Top             =   960
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1296
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
End
Attribute VB_Name = "frmDataEnv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'Adodc1.CommandType = adCmdUnknown
'Adodc1.RecordSource = "select sum(naghlbebadmeghdar) as naghlbebadmeghdar1 from Taab where (gothr='" + Text2.Text + "') "
'Adodc1.Refresh
'Text1.Text = Adodc1.Recordset.Fields!naghlbebadmeghdar1

'Adodc1.CommandType = adCmdUnknown
'Adodc1.RecordSource = "select sum(tolidteydoremeghdar) as tolidteydoremeghdar1 from Exteroder where (gothr='" + Text2.Text + "') "
'Adodc1.Refresh
'Text1.Text = Adodc1.Recordset.Fields!tolidteydoremeghdar1

'Adodc1.CommandType = adCmdUnknown
'Adodc1.RecordSource = "select sum(naghlbebadmeghdar) as naghlbebadmeghdar1 from Taab where (gothr='" + Text2.Text + "') "
'Adodc1.Refresh
'Text1.Text = Adodc1.Recordset.Fields!naghlbebadmeghdar1

'Adodc1.CommandType = adCmdUnknown
'Adodc1.RecordSource = "select sum(tolidteydoremeghdar) as tolidteydoremeghdar1 from Exteroder where (gothr='" + Text2.Text + "') "
'Adodc1.Refresh
'Text1.Text = Adodc1.Recordset.Fields!tolidteydoremeghdar1

'Adodc1.CommandType = adCmdUnknown
'Adodc1.RecordSource = "Select sum(bahamavad) as qqq From infomavad "
'Adodc1.Refresh
'Text1.Text = Adodc1.Recordset.Fields!qqq

'Adodc1.CommandType = adCmdUnknown
'Adodc1.RecordSource = "select sum(meghdar) as meghdar1 from masrafestandardgranol where (idmade='" + Text2.Text + "') "
'Adodc1.Refresh
'Text1.Text = Adodc1.Recordset.Fields!meghdar1


'Adodc1.CommandType = adCmdUnknown
'Adodc1.RecordSource = "select sum(naghlbebadmoney) as naghlbebadmeghdar1 from Koreh where (name='" + Text2.Text + "') "
'Adodc1.Refresh
'Text1.Text = Adodc1.Recordset.Fields!naghlbebadmeghdar1

Adodc1.CommandType = adCmdUnknown
Adodc1.RecordSource = "select sum(naghlbebadmoney) as naghlbebadmeghdar1 from taab where (idmahsol=10)and(rad>=11) and (rad<=27) "
Adodc1.Refresh
Text1.Text = Adodc1.Recordset.Fields!naghlbebadmeghdar1

End Sub


