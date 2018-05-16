VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form15 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«Ê“«‰"
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   14670
   BeginProperty Font 
      Name            =   "B Titr"
      Size            =   9.75
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form15.frx":0000
   LinkTopic       =   "Form15"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   6855
   ScaleWidth      =   14670
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   " ⁄—Ì› „Õ’Ê·« "
      Height          =   6615
      Left            =   4920
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   120
      Width           =   9615
      Begin VB.CommandButton Command4 
         Caption         =   "ê—œ‘ „Õ’Ê·"
         Height          =   375
         Left            =   4920
         RightToLeft     =   -1  'True
         TabIndex        =   60
         Top             =   2280
         Width           =   1455
      End
      Begin VB.ComboBox Combo7 
         Height          =   465
         Left            =   120
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   38
         Top             =   1560
         Width           =   1935
      End
      Begin VB.ComboBox Combo6 
         Height          =   465
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Text            =   "Combo3"
         Top             =   1560
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton Command1 
         Caption         =   "ç«Å"
         Height          =   375
         Left            =   3360
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   2280
         Width           =   1455
      End
      Begin VB.CommandButton Command8 
         Caption         =   "ÊÌ—«Ì‘"
         Height          =   375
         Left            =   6480
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   2280
         Width           =   1455
      End
      Begin VB.ComboBox Combo3 
         Height          =   465
         Left            =   6600
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Text            =   "Combo3"
         Top             =   360
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton Command5 
         Caption         =   "ÃœÌœ"
         Height          =   375
         Left            =   8040
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   2280
         Width           =   1455
      End
      Begin VB.CommandButton Command7 
         Caption         =   "À» "
         Height          =   375
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   2280
         Width           =   1575
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Õ–›"
         Height          =   375
         Left            =   1800
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   2280
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   495
         Index           =   4
         Left            =   6600
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   1560
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   495
         Index           =   3
         Left            =   6600
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   495
         Index           =   2
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   360
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   495
         Index           =   1
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   495
         Index           =   0
         Left            =   3480
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   960
         Width           =   1935
      End
      Begin VB.ComboBox Combo1 
         Height          =   465
         Left            =   6600
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   360
         Width           =   1935
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "Form15.frx":2CFA
         Height          =   3735
         Left            =   120
         TabIndex        =   24
         Top             =   2760
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   6588
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
         ColumnCount     =   8
         BeginProperty Column00 
            DataField       =   "rad"
            Caption         =   "—œÌ›"
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
            DataField       =   "kodemahsol"
            Caption         =   "òœ „Õ’Ê·"
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
            DataField       =   "nomahsol"
            Caption         =   "‰Ê⁄ „Õ’Ê·"
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
            DataField       =   "size"
            Caption         =   "”«Ì“"
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
         BeginProperty Column04 
            DataField       =   "gothr"
            Caption         =   "ﬁÿ—"
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
         BeginProperty Column05 
            DataField       =   "propertikhas"
            Caption         =   "„‘Œ’Â Œ«’"
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
         BeginProperty Column06 
            DataField       =   "ger"
            Caption         =   "»” Â »‰œÌ"
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
            DataField       =   "nomes"
            Caption         =   "‰Ê⁄ „”"
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
               ColumnWidth     =   840.189
            EndProperty
            BeginProperty Column01 
            EndProperty
            BeginProperty Column02 
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1170.142
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1094.74
            EndProperty
            BeginProperty Column05 
            EndProperty
            BeginProperty Column06 
            EndProperty
            BeginProperty Column07 
            EndProperty
         EndProperty
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "‰Ê⁄ »” Â »‰œÌ"
         Height          =   495
         Left            =   2160
         RightToLeft     =   -1  'True
         TabIndex        =   39
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Height          =   495
         Index           =   9
         Left            =   3480
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "—œÌ› :"
         Height          =   495
         Index           =   8
         Left            =   5400
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "„‘Œ’Â Œ«’"
         Height          =   495
         Index           =   4
         Left            =   5400
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "ﬁÿ—"
         Height          =   495
         Index           =   3
         Left            =   8640
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "‰Ê⁄ „Õ’Ê·"
         Height          =   495
         Index           =   2
         Left            =   8640
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "òœ „Õ’Ê·"
         Height          =   495
         Index           =   1
         Left            =   2160
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "”«Ì“"
         Height          =   495
         Index           =   0
         Left            =   2160
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "‰«„ „Õ’Ê·"
         Height          =   495
         Left            =   8640
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   360
         Width           =   855
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5895
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   4725
      _ExtentX        =   8334
      _ExtentY        =   10398
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "„Ê«œ «Ê·ÌÂ „’—›Ì"
      TabPicture(0)   =   "Form15.frx":2D0F
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(1)=   "DataGrid2"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   " ⁄ÌÌ‰ „—«Õ·"
      TabPicture(1)   =   "Form15.frx":2D2B
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DataGrid3"
      Tab(1).Control(1)=   "List4"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "„”  «»ÌœÂ ‘œÂ"
      TabPicture(2)   =   "Form15.frx":2D47
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label2(20)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label2(22)"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label2(24)"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Label2(26)"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Label2(21)"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Label2(23)"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Label2(25)"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "Label2(27)"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "Text1(6)"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "Text1(7)"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "Text1(8)"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "Text1(9)"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).Control(12)=   "Command2"
      Tab(2).Control(12).Enabled=   0   'False
      Tab(2).Control(13)=   "Text1(10)"
      Tab(2).Control(13).Enabled=   0   'False
      Tab(2).Control(14)=   "Text1(11)"
      Tab(2).Control(14).Enabled=   0   'False
      Tab(2).Control(15)=   "Text1(12)"
      Tab(2).Control(15).Enabled=   0   'False
      Tab(2).Control(16)=   "Text1(13)"
      Tab(2).Control(16).Enabled=   0   'False
      Tab(2).Control(17)=   "Option1"
      Tab(2).Control(17).Enabled=   0   'False
      Tab(2).Control(18)=   "Option2"
      Tab(2).Control(18).Enabled=   0   'False
      Tab(2).Control(19)=   "Option3"
      Tab(2).Control(19).Enabled=   0   'False
      Tab(2).ControlCount=   20
      Begin VB.OptionButton Option3 
         Alignment       =   1  'Right Justify
         Caption         =   "¬·Ê„Ì‰ÌÊ„"
         Height          =   495
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   61
         Top             =   480
         Width           =   1455
      End
      Begin VB.OptionButton Option2 
         Alignment       =   1  'Right Justify
         Caption         =   "„”  Ê·Ìœ ‘œÂ"
         Height          =   495
         Left            =   3120
         RightToLeft     =   -1  'True
         TabIndex        =   59
         Top             =   480
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Caption         =   "„”  «»ÌœÂ ‘œÂ"
         Height          =   495
         Left            =   1560
         RightToLeft     =   -1  'True
         TabIndex        =   58
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   495
         Index           =   13
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   52
         Top             =   4440
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   495
         Index           =   12
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   51
         Top             =   3960
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   495
         Index           =   11
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   50
         Top             =   3480
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   495
         Index           =   10
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   49
         Top             =   3000
         Width           =   2175
      End
      Begin VB.CommandButton Command2 
         Caption         =   "À» "
         Height          =   735
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   48
         Top             =   5040
         Width           =   3015
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   495
         Index           =   9
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   46
         Top             =   2520
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   495
         Index           =   8
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   44
         Top             =   2040
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   495
         Index           =   7
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   42
         Top             =   1560
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   495
         Index           =   6
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   40
         Top             =   1080
         Width           =   2175
      End
      Begin VB.ListBox List4 
         Height          =   3510
         ItemData        =   "Form15.frx":2D63
         Left            =   -71760
         List            =   "Form15.frx":2D85
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   600
         Width           =   1335
      End
      Begin VB.Frame Frame2 
         Caption         =   "«›“Êœ‰ „Ê«œ"
         Height          =   1695
         Left            =   -74880
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   360
         Width           =   4455
         Begin VB.ComboBox Combo5 
            Height          =   465
            Left            =   2880
            RightToLeft     =   -1  'True
            TabIndex        =   7
            Text            =   "Combo5"
            Top             =   480
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.ComboBox Combo4 
            Height          =   465
            Left            =   2280
            RightToLeft     =   -1  'True
            TabIndex        =   6
            Text            =   "Combo4"
            Top             =   480
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.CommandButton Command3 
            Caption         =   "À» "
            Height          =   465
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   5
            Top             =   1080
            Width           =   4215
         End
         Begin VB.TextBox Text1 
            Height          =   495
            Index           =   5
            Left            =   120
            TabIndex        =   4
            Top             =   480
            Width           =   1455
         End
         Begin VB.ComboBox Combo2 
            Height          =   465
            Left            =   2280
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   480
            Width           =   2055
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "„ﬁœ«—"
            Height          =   495
            Index           =   7
            Left            =   1560
            RightToLeft     =   -1  'True
            TabIndex        =   8
            Top             =   480
            Width           =   615
         End
      End
      Begin MSDataGridLib.DataGrid DataGrid2 
         Bindings        =   "Form15.frx":2DFB
         Height          =   3615
         Left            =   -74880
         TabIndex        =   9
         Top             =   2160
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   6376
         _Version        =   393216
         AllowUpdate     =   -1  'True
         HeadLines       =   1
         RowHeight       =   29
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
         ColumnCount     =   4
         BeginProperty Column00 
            DataField       =   "qq"
            Caption         =   "‰«„ „«œÂ"
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
            DataField       =   "meghdar2"
            Caption         =   "„ﬁœ«— ‰Â«ÌÌ"
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
            DataField       =   "meghdar"
            Caption         =   "„ﬁœ«— «Ê“«‰"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1080
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "mainacc"
            Caption         =   "„Õ«”»Â „ —«é"
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
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column01 
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column02 
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column03 
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid DataGrid3 
         Bindings        =   "Form15.frx":2E10
         Height          =   5175
         Left            =   -74880
         TabIndex        =   11
         Top             =   600
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   9128
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   24
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
            DataField       =   "rad1"
            Caption         =   "«Ê·ÊÌ "
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
            DataField       =   "name"
            Caption         =   "„—Õ·Â"
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
            Size            =   143
            BeginProperty Column00 
               ColumnWidth     =   915.024
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "„ÊÃÊœÌ Å«Ì«‰ œÊ—Â („»·€)"
         Enabled         =   0   'False
         Height          =   495
         Index           =   27
         Left            =   2400
         RightToLeft     =   -1  'True
         TabIndex        =   56
         Top             =   4440
         Width           =   2175
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "„’—› ÿÌ œÊ—Â („»·€)"
         Enabled         =   0   'False
         Height          =   495
         Index           =   25
         Left            =   2400
         RightToLeft     =   -1  'True
         TabIndex        =   55
         Top             =   3480
         Width           =   2175
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Ê«—œÂ ÿÌ œÊ—Â („»·€)"
         Enabled         =   0   'False
         Height          =   495
         Index           =   23
         Left            =   2400
         RightToLeft     =   -1  'True
         TabIndex        =   54
         Top             =   2520
         Width           =   2175
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "„ÊÃÊœÌ «Ê· œÊ—Â („»·€)"
         Enabled         =   0   'False
         Height          =   495
         Index           =   21
         Left            =   2400
         RightToLeft     =   -1  'True
         TabIndex        =   53
         Top             =   1560
         Width           =   2175
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "„ÊÃÊœÌ Å«Ì«‰ œÊ—Â („ﬁœ«—)"
         Enabled         =   0   'False
         Height          =   495
         Index           =   26
         Left            =   2400
         RightToLeft     =   -1  'True
         TabIndex        =   47
         Top             =   3960
         Width           =   2175
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "„’—› ÿÌ œÊ—Â („ﬁœ«—)"
         Enabled         =   0   'False
         Height          =   495
         Index           =   24
         Left            =   2400
         RightToLeft     =   -1  'True
         TabIndex        =   45
         Top             =   3000
         Width           =   2175
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Ê«—œÂ ÿÌ œÊ—Â („ﬁœ«—)"
         Enabled         =   0   'False
         Height          =   495
         Index           =   22
         Left            =   2400
         RightToLeft     =   -1  'True
         TabIndex        =   43
         Top             =   2040
         Width           =   2175
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "„ÊÃÊœÌ «Ê· œÊ—Â („ﬁœ«—)"
         Enabled         =   0   'False
         Height          =   495
         Index           =   20
         Left            =   2400
         RightToLeft     =   -1  'True
         TabIndex        =   41
         Top             =   1080
         Width           =   2175
      End
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   2760
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
      RecordSource    =   "ozanunder"
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   1440
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
      RecordSource    =   "ozanmain"
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
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   4080
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
      RecordSource    =   "ozanmasir"
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
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   330
      Left            =   5160
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
      RecordSource    =   "P_ozan"
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
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   495
      Left            =   6720
      RightToLeft     =   -1  'True
      TabIndex        =   57
      Top             =   3240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   495
      Index           =   10
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   35
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Height          =   495
      Index           =   13
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   34
      Top             =   6240
      Width           =   3855
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "„Ã„Ê⁄ :"
      Height          =   495
      Index           =   12
      Left            =   4080
      RightToLeft     =   -1  'True
      TabIndex        =   33
      Top             =   6240
      Width           =   735
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   375
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Menu mnumain1 
      Caption         =   " ⁄—Ì›"
      Begin VB.Menu mnumahsol 
         Caption         =   "„Õ’Ê·"
      End
      Begin VB.Menu mnumavadmasraf 
         Caption         =   "„Ê«œ «Ê·ÌÂ „’—›Ì"
      End
      Begin VB.Menu mnubasteband 
         Caption         =   "»” Â »‰œÌ"
      End
   End
End
Attribute VB_Name = "Form15"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db1 As New ADODB.Connection
Dim rs1 As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim rs3 As New ADODB.Recordset
Dim arrayall(100, 3) As String
Dim q As Integer, commove As Integer

Private Sub Combo1_Click()
Call Combo1_LostFocus
End Sub

Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
  Text1(0).SetFocus
End If
End Sub

Private Sub Combo1_LostFocus()
If Combo1.ListIndex <> -1 Then
  Adodc1.ConnectionString = Form3.Text10.Text
  Adodc1.CommandType = adCmdUnknown
  Adodc1.RecordSource = "select * from ozanmain where idmahsol=" + Combo3.List(Combo1.ListIndex)
  Adodc1.Refresh
End If
End Sub

Private Sub Combo2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Text1(5).SetFocus
End Sub

Private Sub Combo7_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Command7.SetFocus
End Sub

Private Sub Command1_Click()
Dim wenumber As String
db1.Open Form3.Text10.Text
  rs1.Open "DELETE FROM P_ozan", db1
db1.Close

db1.Open Form3.Text10.Text
rs1.Open "select * from ozanmain", db1
rs1.MoveFirst
Do
  For q = 0 To 100
    arrayall(q, 0) = ""
    arrayall(q, 1) = ""
    arrayall(q, 2) = ""
    arrayall(q, 3) = ""
  Next q
  rs2.Open "select count(rad) as wenumber11 from ozanunder WHERE (idmahsol=" + Trim(Str(rs1.Fields!idmahsol)) + ") and (rad=" + Trim(Str(rs1.Fields!rad)) + ")", db1
    wenumber = rs2.Fields!wenumber11
  rs2.Close
  rs2.Open "select * from ozanunder WHERE (idmahsol=" + Trim(Str(rs1.Fields!idmahsol)) + ") and (rad=" + Trim(Str(rs1.Fields!rad)) + ") ORDER BY idmade", db1
  If wenumber > 0 Then
    q = 0
    rs2.MoveFirst
    Do
      arrayall(q, 0) = rs2.Fields!meghdar2
      arrayall(q, 1) = rs2.Fields!qq
      q = q + 1
      rs2.MoveNext
    Loop Until rs2.EOF = True
  End If
  
  rs3.Open "select count(rad) as wenumber11 from ozanmasir WHERE (idmahsol=" + Trim(Str(rs1.Fields!idmahsol)) + ") and (rad=" + Trim(Str(rs1.Fields!rad)) + ")", db1
    wenumber = rs3.Fields!wenumber11
  rs3.Close
  
  rs3.Open "select * from ozanmasir WHERE (idmahsol=" + Trim(Str(rs1.Fields!idmahsol)) + ") and (rad=" + Trim(Str(rs1.Fields!rad)) + ") ORDER BY rad1", db1
  q = 0
  If wenumber > 0 Then
    rs3.MoveFirst
    Do
      arrayall(q, 2) = rs3.Fields!rad1
      arrayall(q, 3) = rs3.Fields!Name
      q = q + 1
      rs3.MoveNext
    Loop Until rs3.EOF = True
  End If
  
  For q = 0 To 100
    If (arrayall(q, 0) <> "") Or (arrayall(q, 2) <> "") Then
      Adodc4.Refresh
      Adodc4.Recordset.AddNew
      Adodc4.Recordset.Fields!idmahsol = rs1.Fields!idmahsol
      Adodc4.Recordset.Fields!rad = rs1.Fields!rad
      Adodc4.Recordset.Fields!ID = Trim(Str(Adodc4.Recordset.Fields!idmahsol)) + "." + Trim(Str(Adodc4.Recordset.Fields!rad))
      Adodc4.Recordset.Fields!propertikhas = rs1.Fields!propertikhas
      Adodc4.Recordset.Fields!Size = rs1.Fields!Size
      Adodc4.Recordset.Fields!kodemahsol = rs1.Fields!kodemahsol
      Adodc4.Recordset.Fields!nomahsol = rs1.Fields!nomahsol
      Adodc4.Recordset.Fields!gothr = rs1.Fields!gothr
      Adodc4.Recordset.Fields!sheet1number = rs1.Fields!sheet1number
      If (arrayall(q, 2) = "") And (q > 0) Then arrayall(q, 2) = Val(arrayall(q - 1, 2)) + 1
      Adodc4.Recordset.Fields!masrad = Val(arrayall(q, 2))
      Adodc4.Recordset.Fields!masname = arrayall(q, 3)
      Adodc4.Recordset.Fields!meghdar = Val(arrayall(q, 0))
      Adodc4.Recordset.Fields!qq = arrayall(q, 1)
      Adodc4.Recordset.Update
    End If
  Next q
  rs3.Close
  rs2.Close
  rs1.MoveNext
Loop Until rs1.EOF = True

rs1.Close

db1.Close
Form46.Label1.Caption = 1
Form46.Show
End Sub

Private Sub Command2_Click()
If Adodc1.Recordset.RecordCount > 0 Then
  If Option1.Value = True Then
    For q = 6 To 13
      If Val(Text1(q).Text) <= 0 Then Text1(q).Text = 0
    Next q
    If Label5.Caption = 1 Then
      db1.Open Form3.Text10.Text
        rs1.Open "INSERT INTO NewMes (idmahsol,rad,mojodiavalmeghdar,mojodiavalmemoney,tolidteydoremeghdar,tolidteydoremoney,naghlbebadmeghdar,naghlbebadmoney,mojodiendmeghdar,mojodiendmoney) VALUES (" + Trim(Str(Adodc1.Recordset.Fields!idmahsol)) + "," + Trim(Str(Adodc1.Recordset.Fields!rad)) + "," + Text1(6).Text + "," + Text1(7).Text + "," + Text1(8).Text + "," + Text1(9).Text + "," + Text1(10).Text + "," + Text1(11).Text + "," + Text1(12).Text + "," + Text1(13).Text + ")", db1
      db1.Close
      MsgBox "«ÿ·«⁄«  »« „Ê›ﬁÌ  À»  ‘œ", vbMsgBoxRight + vbInformation, ""
    Else
      db1.Open Form3.Text10.Text
        rs1.Open "UPDATE NewMes SET [mojodiavalmeghdar]='" + Text1(6).Text + "',[mojodiavalmemoney]='" + Text1(7).Text + "',[tolidteydoremeghdar]='" + Text1(8).Text + "',[tolidteydoremoney]='" + Text1(9).Text + "',[naghlbebadmeghdar]='" + Text1(10).Text + "',[naghlbebadmoney]='" + Text1(11).Text + "',[mojodiendmeghdar]='" + Text1(12).Text + "',[mojodiendmoney]='" + Text1(13).Text + "' WHERE (idmahsol=" + Trim(Str(Adodc1.Recordset.Fields!idmahsol)) + ") and (rad = " + Trim(Str(Adodc1.Recordset.Fields!rad)) + ")", db1
      db1.Close
      MsgBox "«ÿ·«⁄«  »« „Ê›ﬁÌ   €ÌÌ— ÅÌœ« ò—œ", vbMsgBoxRight + vbInformation, ""
    End If
    Adodc1.Recordset.Fields!nomes = "„”  «»ÌœÂ ‘œÂ"
    Adodc1.Recordset.Update
  End If
  
  If Option2.Value = True Then
    Adodc1.Recordset.Fields!nomes = "„”  Ê·Ìœ ‘œÂ"
    Adodc1.Recordset.Update
  End If

  If Option3.Value = True Then
    Adodc1.Recordset.Fields!nomes = "¬·Ê„Ì‰ÌÊ„"
    Adodc1.Recordset.Update
  End If
End If
End Sub

Private Sub Command3_Click()
Dim sd As String
If (Combo2.ListIndex = -1) Or (Text1(5).Text = "") Then
  MsgBox "·ÿ›«  „«„Ì ›Ì·œ Â« —«  ò„Ì· ‰„«ÌÌœ", vbCritical + vbMsgBoxRight, ""
  Exit Sub
End If

If Adodc2.Recordset.RecordCount > 0 Then
  Adodc2.Recordset.MoveFirst
  Do
    If Adodc2.Recordset.Fields!idmade = Combo4.List(Combo2.ListIndex) Then
      Exit Sub
    End If
    Adodc2.Recordset.MoveNext
  Loop Until Adodc2.Recordset.EOF = True
End If

Adodc2.Refresh
Adodc2.Recordset.AddNew
Adodc2.Recordset.Fields!idmahsol = Adodc1.Recordset.Fields!idmahsol
Adodc2.Recordset.Fields!rad = Adodc1.Recordset.Fields!rad
Adodc2.Recordset.Fields!idmade = Combo4.List(Combo2.ListIndex)
sd = Val(Text1(5).Text) * Val(Combo5.List(Combo2.ListIndex))
Adodc2.Recordset.Fields!meghdar2 = sd
Adodc2.Recordset.Fields!meghdar = Text1(5).Text
Adodc2.Recordset.Fields!qq = Combo2.Text
Adodc2.Recordset.Update
Label2(13).Caption = Val(Label2(13).Caption) + Val(Text1(5).Text * Combo5.List(Combo2.ListIndex))
MsgBox "«ÿ·«⁄«  »« „Ê›ﬁÌ  À»  ‘œ", vbInformation + vbMsgBoxRight, ""
End Sub

Private Sub Command4_Click()
Dim db1 As New ADODB.Connection
Dim db2 As New ADODB.Connection
Dim rs(10) As New ADODB.Recordset

If Adodc1.Recordset.RecordCount > 0 Then
  db1.Open Form3.Text10.Text
    rs(0).Open "DELETE FROM P_AllMah", db1
  db1.Close
  
  db1.Open Form3.Text10.Text
    rs(0).Open "DELETE FROM P_AllMah_under", db1
  db1.Close
  
  db1.Open Form3.Text10.Text
    '«Ê“«‰
    rs(0).Open "SELECT Count(idmahsol) As rsnumber FROM ozanunder WHERE (idmahsol= " + Trim(Str(Adodc1.Recordset.Fields!idmahsol)) + ") AND (rad =" + Trim(Str(Adodc1.Recordset.Fields!rad)) + ")", db1
      If rs(0).Fields!rsnumber > 0 Then
        rs(1).Open "SELECT * FROM ozanunder WHERE (idmahsol= " + Trim(Str(Adodc1.Recordset.Fields!idmahsol)) + ") AND (rad =" + Trim(Str(Adodc1.Recordset.Fields!rad)) + ")", db1
          rs(1).MoveFirst
          Do
            '„’—› «” «‰œ«—œ  Ê·Ìœ
            '„’—› «” «‰œ«—œ «ò” —Êœ—
            rs(2).Open "SELECT Count(idmahsol) As rsnumber FROM masrafestandardmavad2 WHERE (idmahsol= " + Trim(Str(Adodc1.Recordset.Fields!idmahsol)) + ") AND (rad =" + Trim(Str(Adodc1.Recordset.Fields!rad)) + ") AND (idmade='" + rs(1).Fields!idmade + "')", db1
              If rs(2).Fields!rsnumber > 0 Then
                rs(3).Open "SELECT * FROM masrafestandardmavad2 WHERE (idmahsol= " + Trim(Str(Adodc1.Recordset.Fields!idmahsol)) + ") AND (rad =" + Trim(Str(Adodc1.Recordset.Fields!rad)) + ") AND (idmade='" + rs(1).Fields!idmade + "')", db1
                  tmp1 = rs(3).Fields!meghdar
                  tmp2 = rs(3).Fields!meghdar1
                rs(3).Close
              Else
                tmp1 = 0
                tmp2 = 0
              End If
            rs(2).Close
            
            '„’—› «” «‰œ«—œ ê—«‰Ê·
            rs(2).Open "SELECT Count(idmahsol) As rsnumber FROM masrafestandardgranol WHERE (idmahsol= " + Trim(Str(Adodc1.Recordset.Fields!idmahsol)) + ") AND (rad =" + Trim(Str(Adodc1.Recordset.Fields!rad)) + ") AND (idmade='" + rs(1).Fields!idmade + "')", db1
              If rs(2).Fields!rsnumber > 0 Then
                rs(3).Open "SELECT * FROM masrafestandardgranol WHERE (idmahsol= " + Trim(Str(Adodc1.Recordset.Fields!idmahsol)) + ") AND (rad =" + Trim(Str(Adodc1.Recordset.Fields!rad)) + ") AND (idmade='" + rs(1).Fields!idmade + "')", db1
                  tmp3 = rs(3).Fields!meghdar
                rs(3).Close
              Else
                tmp3 = 0
              End If
            rs(2).Close
            
            db2.Open Form3.Text10.Text
             rs(4).Open "INSERT INTO P_AllMah_under (rad,name,q1,q2,q3,q4) VALUES (" + rs(1).Fields!idmade + ",'" + rs(1).Fields!qq + "','" + rs(1).Fields!meghdar2 + "','" + Trim(Str(tmp1)) + "','" + tmp2 + "','" + Trim(Str(tmp3)) + "')", db2
            db2.Close
            rs(1).MoveNext
          Loop Until rs(1).EOF = True
        rs(1).Close
      End If
    rs(0).Close

    
    ' «»
    rs(0).Open "SELECT Count(idmahsol) As rsnumber FROM Taab WHERE (idmahsol= " + Trim(Str(Adodc1.Recordset.Fields!idmahsol)) + ") AND (rad =" + Trim(Str(Adodc1.Recordset.Fields!rad)) + ")", db1
      If rs(0).Fields!rsnumber > 0 Then
        rs(1).Open "SELECT * FROM Taab WHERE (idmahsol= " + Trim(Str(Adodc1.Recordset.Fields!idmahsol)) + ") AND (rad =" + Trim(Str(Adodc1.Recordset.Fields!rad)) + ")", db1
          db2.Open Form3.Text10.Text
            rs(2).Open "INSERT INTO P_AllMah (rad,name,vaznmes,bahamavad,dastmozd,sarbar,estelak,gheymattamam,mojodiavalmeghdar,mojodiavalmemoney,tolidteydoremeghdar,tolidteydoremoney,naghlbebadmeghdar,naghlbebadmoney,mojodiendmeghdar,mojodiendmoney,bahagranol,bastebandim) VALUES (1,' «»'," + Trim(Str(rs(1).Fields!vaznmes)) + "," + Trim(Str(rs(1).Fields!bahamavad)) + "," + Trim(Str(rs(1).Fields!dastmozd)) + "," + Trim(Str(rs(1).Fields!sarbar)) + "," + Trim(Str(rs(1).Fields!estelak)) + "," + Trim(Str(rs(1).Fields!gheymattamam)) + "," + Trim(Str(rs(1).Fields!mojodiavalmeghdar)) + "," + Trim(Str(rs(1).Fields!mojodiavalmemoney)) + "," + Trim(Str(rs(1).Fields!tolidteydoremeghdar)) + "," + Trim(Str(rs(1).Fields!tolidteydoremoney)) + "," + Trim(Str(rs(1).Fields!naghlbebadmeghdar)) + "," + Trim(Str(rs(1).Fields!naghlbebadmoney)) + "," + Trim(Str(rs(1).Fields!mojodiendmeghdar)) + "," + Trim(Str(rs(1).Fields!mojodiendmoney)) + ",0,0)", db2
          db2.Close
        rs(1).Close
      End If
    rs(0).Close
  
    '»«‰ç—
    rs(0).Open "SELECT Count(idmahsol) As rsnumber FROM Bancher WHERE (idmahsol= " + Trim(Str(Adodc1.Recordset.Fields!idmahsol)) + ") AND (rad =" + Trim(Str(Adodc1.Recordset.Fields!rad)) + ")", db1
      If rs(0).Fields!rsnumber > 0 Then
        rs(1).Open "SELECT * FROM Bancher WHERE (idmahsol= " + Trim(Str(Adodc1.Recordset.Fields!idmahsol)) + ") AND (rad =" + Trim(Str(Adodc1.Recordset.Fields!rad)) + ")", db1
          db2.Open Form3.Text10.Text
            rs(2).Open "INSERT INTO P_AllMah (rad,name,vaznmes,bahamavad,dastmozd,sarbar,estelak,gheymattamam,mojodiavalmeghdar,mojodiavalmemoney,tolidteydoremeghdar,tolidteydoremoney,naghlbebadmeghdar,naghlbebadmoney,mojodiendmeghdar,mojodiendmoney,bahagranol,bastebandim) VALUES (2,'»«‰ç—'," + Trim(Str(rs(1).Fields!vaznmes)) + "," + Trim(Str(rs(1).Fields!bahamavad)) + "," + Trim(Str(rs(1).Fields!dastmozd)) + "," + Trim(Str(rs(1).Fields!sarbar)) + "," + Trim(Str(rs(1).Fields!estelak)) + "," + Trim(Str(rs(1).Fields!gheymattamam)) + "," + Trim(Str(rs(1).Fields!mojodiavalmeghdar)) + "," + Trim(Str(rs(1).Fields!mojodiavalmemoney)) + "," + Trim(Str(rs(1).Fields!tolidteydoremeghdar)) + "," + Trim(Str(rs(1).Fields!tolidteydoremoney)) + "," + Trim(Str(rs(1).Fields!naghlbebadmeghdar)) + "," + Trim(Str(rs(1).Fields!naghlbebadmoney)) + "," + Trim(Str(rs(1).Fields!mojodiendmeghdar)) + "," + Trim(Str(rs(1).Fields!mojodiendmoney)) + ",0,0)", db2
          db2.Close
        rs(1).Close
      End If
    rs(0).Close
  
    '«” —‰œ— 6
    rs(0).Open "SELECT Count(idmahsol) As rsnumber FROM Sterander1_6 WHERE (idmahsol= " + Trim(Str(Adodc1.Recordset.Fields!idmahsol)) + ") AND (rad =" + Trim(Str(Adodc1.Recordset.Fields!rad)) + ")", db1
      If rs(0).Fields!rsnumber > 0 Then
        rs(1).Open "SELECT * FROM Sterander1_6 WHERE (idmahsol= " + Trim(Str(Adodc1.Recordset.Fields!idmahsol)) + ") AND (rad =" + Trim(Str(Adodc1.Recordset.Fields!rad)) + ")", db1
          db2.Open Form3.Text10.Text
            rs(2).Open "INSERT INTO P_AllMah (rad,name,vaznmes,bahamavad,dastmozd,sarbar,estelak,gheymattamam,mojodiavalmeghdar,mojodiavalmemoney,tolidteydoremeghdar,tolidteydoremoney,naghlbebadmeghdar,naghlbebadmoney,mojodiendmeghdar,mojodiendmoney,bahagranol,bastebandim) VALUES (3,'«” —‰œ— 1+6'," + Trim(Str(rs(1).Fields!vaznmes)) + "," + Trim(Str(rs(1).Fields!bahamavad)) + "," + Trim(Str(rs(1).Fields!dastmozd)) + "," + Trim(Str(rs(1).Fields!sarbar)) + "," + Trim(Str(rs(1).Fields!estelak)) + "," + Trim(Str(rs(1).Fields!gheymattamam)) + "," + Trim(Str(rs(1).Fields!mojodiavalmeghdar)) + "," + Trim(Str(rs(1).Fields!mojodiavalmemoney)) + "," + Trim(Str(rs(1).Fields!tolidteydoremeghdar)) + "," + Trim(Str(rs(1).Fields!tolidteydoremoney)) + "," + Trim(Str(rs(1).Fields!naghlbebadmeghdar)) + "," + Trim(Str(rs(1).Fields!naghlbebadmoney)) + "," + Trim(Str(rs(1).Fields!mojodiendmeghdar)) + "," + Trim(Str(rs(1).Fields!mojodiendmoney)) + ",0,0)", db2
          db2.Close
        rs(1).Close
      End If
    rs(0).Close
  
    '«” —‰œ— 36
    rs(0).Open "SELECT Count(idmahsol) As rsnumber FROM Sterander1_36 WHERE (idmahsol= " + Trim(Str(Adodc1.Recordset.Fields!idmahsol)) + ") AND (rad =" + Trim(Str(Adodc1.Recordset.Fields!rad)) + ")", db1
      If rs(0).Fields!rsnumber > 0 Then
        rs(1).Open "SELECT * FROM Sterander1_36 WHERE (idmahsol= " + Trim(Str(Adodc1.Recordset.Fields!idmahsol)) + ") AND (rad =" + Trim(Str(Adodc1.Recordset.Fields!rad)) + ")", db1
          db2.Open Form3.Text10.Text
            rs(2).Open "INSERT INTO P_AllMah (rad,name,vaznmes,bahamavad,dastmozd,sarbar,estelak,gheymattamam,mojodiavalmeghdar,mojodiavalmemoney,tolidteydoremeghdar,tolidteydoremoney,naghlbebadmeghdar,naghlbebadmoney,mojodiendmeghdar,mojodiendmoney,bahagranol,bastebandim) VALUES (4,'«” —‰œ— 1+36'," + Trim(Str(rs(1).Fields!vaznmes)) + "," + Trim(Str(rs(1).Fields!bahamavad)) + "," + Trim(Str(rs(1).Fields!dastmozd)) + "," + Trim(Str(rs(1).Fields!sarbar)) + "," + Trim(Str(rs(1).Fields!estelak)) + "," + Trim(Str(rs(1).Fields!gheymattamam)) + "," + Trim(Str(rs(1).Fields!mojodiavalmeghdar)) + "," + Trim(Str(rs(1).Fields!mojodiavalmemoney)) + "," + Trim(Str(rs(1).Fields!tolidteydoremeghdar)) + "," + Trim(Str(rs(1).Fields!tolidteydoremoney)) + "," + Trim(Str(rs(1).Fields!naghlbebadmeghdar)) + "," + Trim(Str(rs(1).Fields!naghlbebadmoney)) + "," + Trim(Str(rs(1).Fields!mojodiendmeghdar)) + "," + Trim(Str(rs(1).Fields!mojodiendmoney)) + ",0,0)", db2
          db2.Close
        rs(1).Close
      End If
    rs(0).Close
  
    '«” —‰œ— 4
    rs(0).Open "SELECT Count(idmahsol) As rsnumber FROM Sterander1_4 WHERE (idmahsol= " + Trim(Str(Adodc1.Recordset.Fields!idmahsol)) + ") AND (rad =" + Trim(Str(Adodc1.Recordset.Fields!rad)) + ")", db1
      If rs(0).Fields!rsnumber > 0 Then
        rs(1).Open "SELECT * FROM Sterander1_4 WHERE (idmahsol= " + Trim(Str(Adodc1.Recordset.Fields!idmahsol)) + ") AND (rad =" + Trim(Str(Adodc1.Recordset.Fields!rad)) + ")", db1
          db2.Open Form3.Text10.Text
            rs(2).Open "INSERT INTO P_AllMah (rad,name,vaznmes,bahamavad,dastmozd,sarbar,estelak,gheymattamam,mojodiavalmeghdar,mojodiavalmemoney,tolidteydoremeghdar,tolidteydoremoney,naghlbebadmeghdar,naghlbebadmoney,mojodiendmeghdar,mojodiendmoney,bahagranol,bastebandim) VALUES (5,'«” —‰œ— 1+4'," + Trim(Str(rs(1).Fields!vaznmes)) + "," + Trim(Str(rs(1).Fields!bahamavad)) + "," + Trim(Str(rs(1).Fields!dastmozd)) + "," + Trim(Str(rs(1).Fields!sarbar)) + "," + Trim(Str(rs(1).Fields!estelak)) + "," + Trim(Str(rs(1).Fields!gheymattamam)) + "," + Trim(Str(rs(1).Fields!mojodiavalmeghdar)) + "," + Trim(Str(rs(1).Fields!mojodiavalmemoney)) + "," + Trim(Str(rs(1).Fields!tolidteydoremeghdar)) + "," + Trim(Str(rs(1).Fields!tolidteydoremoney)) + "," + Trim(Str(rs(1).Fields!naghlbebadmeghdar)) + "," + Trim(Str(rs(1).Fields!naghlbebadmoney)) + "," + Trim(Str(rs(1).Fields!mojodiendmeghdar)) + "," + Trim(Str(rs(1).Fields!mojodiendmoney)) + ",0,0)", db2
          db2.Close
        rs(1).Close
      End If
    rs(0).Close
  
    'œ—«„
    rs(0).Open "SELECT Count(idmahsol) As rsnumber FROM DramToester WHERE (idmahsol= " + Trim(Str(Adodc1.Recordset.Fields!idmahsol)) + ") AND (rad =" + Trim(Str(Adodc1.Recordset.Fields!rad)) + ")", db1
      If rs(0).Fields!rsnumber > 0 Then
        rs(1).Open "SELECT * FROM DramToester WHERE (idmahsol= " + Trim(Str(Adodc1.Recordset.Fields!idmahsol)) + ") AND (rad =" + Trim(Str(Adodc1.Recordset.Fields!rad)) + ")", db1
          db2.Open Form3.Text10.Text
            rs(2).Open "INSERT INTO P_AllMah (rad,name,vaznmes,bahamavad,dastmozd,sarbar,estelak,gheymattamam,mojodiavalmeghdar,mojodiavalmemoney,tolidteydoremeghdar,tolidteydoremoney,naghlbebadmeghdar,naghlbebadmoney,mojodiendmeghdar,mojodiendmoney,bahagranol,bastebandim) VALUES (6,'œ—«„  ÊÌ” —'," + Trim(Str(rs(1).Fields!vaznmes)) + "," + Trim(Str(rs(1).Fields!bahamavad)) + "," + Trim(Str(rs(1).Fields!dastmozd)) + "," + Trim(Str(rs(1).Fields!sarbar)) + "," + Trim(Str(rs(1).Fields!estelak)) + "," + Trim(Str(rs(1).Fields!gheymattamam)) + "," + Trim(Str(rs(1).Fields!mojodiavalmeghdar)) + "," + Trim(Str(rs(1).Fields!mojodiavalmemoney)) + "," + Trim(Str(rs(1).Fields!tolidteydoremeghdar)) + "," + Trim(Str(rs(1).Fields!tolidteydoremoney)) + "," + Trim(Str(rs(1).Fields!naghlbebadmeghdar)) + "," + Trim(Str(rs(1).Fields!naghlbebadmoney)) + "," + Trim(Str(rs(1).Fields!mojodiendmeghdar)) + "," + Trim(Str(rs(1).Fields!mojodiendmoney)) + ",0,0)", db2
          db2.Close
        rs(1).Close
      End If
    rs(0).Close
  
    '„Œ«»—« Ì
    rs(0).Open "SELECT Count(idmahsol) As rsnumber FROM Mokhaberat WHERE (idmahsol= " + Trim(Str(Adodc1.Recordset.Fields!idmahsol)) + ") AND (rad =" + Trim(Str(Adodc1.Recordset.Fields!rad)) + ")", db1
      If rs(0).Fields!rsnumber > 0 Then
        rs(1).Open "SELECT * FROM Mokhaberat WHERE (idmahsol= " + Trim(Str(Adodc1.Recordset.Fields!idmahsol)) + ") AND (rad =" + Trim(Str(Adodc1.Recordset.Fields!rad)) + ")", db1
          db2.Open Form3.Text10.Text
            rs(2).Open "INSERT INTO P_AllMah (rad,name,vaznmes,bahamavad,dastmozd,sarbar,estelak,gheymattamam,mojodiavalmeghdar,mojodiavalmemoney,tolidteydoremeghdar,tolidteydoremoney,naghlbebadmeghdar,naghlbebadmoney,mojodiendmeghdar,mojodiendmoney,bahagranol,bastebandim) VALUES (7,'„Œ«»—« Ì'," + Trim(Str(rs(1).Fields!vaznmes)) + "," + Trim(Str(rs(1).Fields!bahamavad)) + "," + Trim(Str(rs(1).Fields!dastmozd)) + "," + Trim(Str(rs(1).Fields!sarbar)) + "," + Trim(Str(rs(1).Fields!estelak)) + "," + Trim(Str(rs(1).Fields!gheymattamam)) + "," + Trim(Str(rs(1).Fields!mojodiavalmeghdar)) + "," + Trim(Str(rs(1).Fields!mojodiavalmemoney)) + "," + Trim(Str(rs(1).Fields!tolidteydoremeghdar)) + "," + Trim(Str(rs(1).Fields!tolidteydoremoney)) + "," + Trim(Str(rs(1).Fields!naghlbebadmeghdar)) + "," + Trim(Str(rs(1).Fields!naghlbebadmoney)) + "," + Trim(Str(rs(1).Fields!mojodiendmeghdar)) + "," + Trim(Str(rs(1).Fields!mojodiendmoney)) + ",0,0)", db2
          db2.Close
        rs(1).Close
      End If
    rs(0).Close
  
    '«ò” —Êœ—
    rs(0).Open "SELECT Count(idmahsol) As rsnumber FROM Exteroder WHERE (idmahsol= " + Trim(Str(Adodc1.Recordset.Fields!idmahsol)) + ") AND (rad =" + Trim(Str(Adodc1.Recordset.Fields!rad)) + ")", db1
      If rs(0).Fields!rsnumber > 0 Then
        rs(1).Open "SELECT * FROM Exteroder WHERE (idmahsol= " + Trim(Str(Adodc1.Recordset.Fields!idmahsol)) + ") AND (rad =" + Trim(Str(Adodc1.Recordset.Fields!rad)) + ")", db1
          db2.Open Form3.Text10.Text
            rs(2).Open "INSERT INTO P_AllMah (rad,name,vaznmes,bahamavad,dastmozd,sarbar,estelak,gheymattamam,mojodiavalmeghdar,mojodiavalmemoney,tolidteydoremeghdar,tolidteydoremoney,naghlbebadmeghdar,naghlbebadmoney,mojodiendmeghdar,mojodiendmoney,bahagranol,bastebandim) VALUES (8,'«ò” —Êœ—'," + Trim(Str(rs(1).Fields!vaznmes)) + "," + Trim(Str(rs(1).Fields!bahaymavad)) + "," + Trim(Str(rs(1).Fields!dastmozd)) + "," + Trim(Str(rs(1).Fields!sarbar)) + "," + Trim(Str(rs(1).Fields!estelak)) + "," + Trim(Str(rs(1).Fields!gheymattamam)) + "," + Trim(Str(rs(1).Fields!mojodiavalmeghdar)) + "," + Trim(Str(rs(1).Fields!mojodiavalmemoney)) + "," + Trim(Str(rs(1).Fields!tolidteydoremeghdar)) + "," + Trim(Str(rs(1).Fields!tolidteydoremoney)) + "," + Trim(Str(rs(1).Fields!naghlbebadmeghdar)) + "," + Trim(Str(rs(1).Fields!naghlbebadmoney)) + "," + Trim(Str(rs(1).Fields!mojodiendmeghdar)) + "," + Trim(Str(rs(1).Fields!mojodiendmoney)) + "," + Trim(Str(rs(1).Fields!granol)) + ",0)", db2
          db2.Close
        rs(1).Close
      End If
    rs(0).Close
  
    '»” Â »‰œÌ
    rs(0).Open "SELECT Count(idmahsol) As rsnumber FROM Bastebandi WHERE (idmahsol= " + Trim(Str(Adodc1.Recordset.Fields!idmahsol)) + ") AND (rad =" + Trim(Str(Adodc1.Recordset.Fields!rad)) + ")", db1
      If rs(0).Fields!rsnumber > 0 Then
        rs(1).Open "SELECT * FROM Bastebandi WHERE (idmahsol= " + Trim(Str(Adodc1.Recordset.Fields!idmahsol)) + ") AND (rad =" + Trim(Str(Adodc1.Recordset.Fields!rad)) + ")", db1
          db2.Open Form3.Text10.Text
            rs(2).Open "INSERT INTO P_AllMah (rad,name,vaznmes,bahamavad,dastmozd,sarbar,estelak,gheymattamam,mojodiavalmeghdar,mojodiavalmemoney,tolidteydoremeghdar,tolidteydoremoney,naghlbebadmeghdar,naghlbebadmoney,mojodiendmeghdar,mojodiendmoney,bahagranol,bastebandim) VALUES (9,'»” Â »‰œÌ'," + Trim(Str(rs(1).Fields!vaznmes)) + "," + Trim(Str(rs(1).Fields!bahamavad)) + "," + Trim(Str(rs(1).Fields!dastmozd)) + "," + Trim(Str(rs(1).Fields!sarbar)) + "," + Trim(Str(rs(1).Fields!estelak)) + "," + Trim(Str(rs(1).Fields!gheymattamam)) + "," + Trim(Str(rs(1).Fields!mojodiavalmeghdar)) + "," + Trim(Str(rs(1).Fields!mojodiavalmemoney)) + "," + Trim(Str(rs(1).Fields!tolidteydoremeghdar)) + "," + Trim(Str(rs(1).Fields!tolidteydoremoney)) + "," + Trim(Str(rs(1).Fields!naghlbebadmeghdar)) + "," + Trim(Str(rs(1).Fields!naghlbebadmoney)) + "," + Trim(Str(rs(1).Fields!mojodiendmeghdar)) + "," + Trim(Str(rs(1).Fields!mojodiendmoney)) + ",0," + Trim(Str(rs(1).Fields!baste)) + ")", db2
          db2.Close
        rs(1).Close
      End If
    rs(0).Close
  
    '«‰»«— „Õ’Ê·
    rs(0).Open "SELECT Count(idmahsol) As rsnumber FROM AnbarMahsol WHERE (idmahsol= " + Trim(Str(Adodc1.Recordset.Fields!idmahsol)) + ") AND (rad =" + Trim(Str(Adodc1.Recordset.Fields!rad)) + ")", db1
      If rs(0).Fields!rsnumber > 0 Then
        rs(1).Open "SELECT * FROM AnbarMahsol WHERE (idmahsol= " + Trim(Str(Adodc1.Recordset.Fields!idmahsol)) + ") AND (rad =" + Trim(Str(Adodc1.Recordset.Fields!rad)) + ")", db1
          db2.Open Form3.Text10.Text
            rs(2).Open "INSERT INTO P_AllMah (rad,name,vaznmes,bahamavad,dastmozd,sarbar,estelak,gheymattamam,mojodiavalmeghdar,mojodiavalmemoney,tolidteydoremeghdar,tolidteydoremoney,naghlbebadmeghdar,naghlbebadmoney,mojodiendmeghdar,mojodiendmoney,bahagranol,bastebandim) VALUES (10,'«‰»«— „Õ’Ê·'," + Trim(Str(rs(1).Fields!vaznmes)) + "," + Trim(Str(rs(1).Fields!bahamavad)) + "," + Trim(Str(rs(1).Fields!dastmozd)) + "," + Trim(Str(rs(1).Fields!sarbar)) + "," + Trim(Str(rs(1).Fields!estelak)) + "," + Trim(Str(rs(1).Fields!gheymattamam)) + "," + Trim(Str(rs(1).Fields!mojodiavalmeghdar)) + "," + Trim(Str(rs(1).Fields!mojodiavalmemoney)) + "," + Trim(Str(rs(1).Fields!tolidteydoremeghdar)) + "," + Trim(Str(rs(1).Fields!tolidteydoremoney)) + "," + Trim(Str(rs(1).Fields!naghlbebadmeghdar)) + "," + Trim(Str(rs(1).Fields!naghlbebadmoney)) + "," + Trim(Str(rs(1).Fields!mojodiendmeghdar)) + "," + Trim(Str(rs(1).Fields!mojodiendmoney)) + ",0,0)", db2
          db2.Close
        rs(1).Close
      End If
    rs(0).Close
  
  Form46.Label1.Caption = 2
  Form46.Show
  db1.Close
End If
End Sub

Private Sub Command5_Click()
Label3.Caption = 1
For q = 0 To 4
  Text1(q).Text = ""
Next q
q = 1
If Adodc1.Recordset.RecordCount > 0 Then
  Adodc1.Recordset.Sort = "rad"
  Adodc1.Recordset.MoveFirst
  Do
    If Adodc1.Recordset.Fields!rad <> q Then Exit Do
    q = q + 1
    Adodc1.Recordset.MoveNext
  Loop Until Adodc1.Recordset.EOF = True
End If
Label2(9).Caption = q
Label2(10).Caption = 0
Text1(3).SetFocus
End Sub

Private Sub Command7_Click()
If Combo1.ListIndex <> -1 Then
  If (Label3.Caption = 1) And (Label2(9).Caption <> "") Then
    Adodc1.Recordset.AddNew
    Adodc1.Recordset.Fields!idmahsol = Combo3.List(Combo1.ListIndex)
    Adodc1.Recordset.Fields!rad = Label2(9).Caption
    Adodc1.Recordset.Fields!propertikhas = Text1(0).Text
    Adodc1.Recordset.Fields!Size = Text1(1).Text
    Adodc1.Recordset.Fields!kodemahsol = Text1(2).Text
    Adodc1.Recordset.Fields!nomahsol = Text1(3).Text
    Adodc1.Recordset.Fields!gothr = Text1(4).Text
    Adodc1.Recordset.Fields!ger = Combo7.List(Combo7.ListIndex)
    Adodc1.Recordset.Fields!gercode = Combo6.List(Combo7.ListIndex)
    Adodc1.Recordset.Fields!nomes = "„”  Ê·Ìœ ‘œÂ"
    Adodc1.Recordset.Update
    MsgBox "«ÿ·«⁄«  »« „Ê›ﬁÌ  À»  ‘œ", vbMsgBoxRight + vbInformation, ""
    Call Command5_Click
    Call Combo1_LostFocus
  End If

  If (Label3.Caption = 2) And (Label2(9).Caption <> "") Then
    Adodc1.Recordset.Fields!idmahsol = Combo3.List(Combo1.ListIndex)
    Adodc1.Recordset.Fields!rad = Label2(9).Caption
    Adodc1.Recordset.Fields!propertikhas = Text1(0).Text
    Adodc1.Recordset.Fields!Size = Text1(1).Text
    Adodc1.Recordset.Fields!kodemahsol = Text1(2).Text
    Adodc1.Recordset.Fields!nomahsol = Text1(3).Text
    Adodc1.Recordset.Fields!gothr = Text1(4).Text
    Adodc1.Recordset.Fields!ger = Combo7.List(Combo7.ListIndex)
    Adodc1.Recordset.Fields!gercode = Combo6.List(Combo7.ListIndex)
    Adodc1.Recordset.Update
    MsgBox "«ÿ·«⁄«  »« „Ê›ﬁÌ   €ÌÌ— ÅÌœ« ò—œ", vbMsgBoxRight + vbInformation, ""
  End If
End If
End Sub

Private Sub Command8_Click()
If Adodc1.Recordset.RecordCount > 0 Then
  DataGrid1.Col = 0
  Label2(9).Caption = DataGrid1.Text
  
  DataGrid1.Col = 1
  Text1(2).Text = DataGrid1.Text
  
  DataGrid1.Col = 2
  Text1(3).Text = DataGrid1.Text
  
  DataGrid1.Col = 3
  Text1(1).Text = DataGrid1.Text
  
  DataGrid1.Col = 4
  Text1(4).Text = DataGrid1.Text
  
  DataGrid1.Col = 5
  Text1(0).Text = DataGrid1.Text
  
  DataGrid1.Col = 6
  For q = 0 To Combo7.ListCount - 1
    If Combo7.List(q) = DataGrid1.Text Then Combo7.ListIndex = q
  Next q
  
  Label3.Caption = 2
End If
End Sub

Private Sub DataGrid1_Click()
If Adodc1.Recordset.RecordCount > 0 Then
  List4.Clear
  For q = 0 To Form13.List1.ListCount - 1
    List4.AddItem Form13.List1.List(q)
  Next q
  commove = 1
  q = Adodc1.Recordset.Fields!rad
  
  Adodc2.ConnectionString = Form3.Text10.Text
  Adodc2.CommandType = adCmdUnknown
  Adodc2.RecordSource = "select sum(meghdar2) as amin12 from ozanunder where (idmahsol=" + Trim(Str(Combo3.List(Combo1.ListIndex))) + ") and (rad = " + Trim(Str(q)) + ")"
  Adodc2.Refresh
  If IsNull(Adodc2.Recordset.Fields!amin12) Then
    Label2(13).Caption = 0
  Else
    Label2(13).Caption = Adodc2.Recordset.Fields!amin12
  End If
  
  Adodc2.ConnectionString = Form3.Text10.Text
  Adodc2.CommandType = adCmdUnknown
  Adodc2.RecordSource = "select * from ozanunder where (idmahsol=" + Trim(Str(Combo3.List(Combo1.ListIndex))) + ") and (rad = " + Trim(Str(q)) + ")"
  Adodc2.Refresh
  If Adodc2.Recordset.RecordCount > 0 Then
    Adodc2.Recordset.MoveFirst
    Do
      Form4.Adodc1.Recordset.Find "idmavad=" + Adodc2.Recordset.Fields!idmade, , adSearchForward, 1
      DataGrid2.Col = 0
      DataGrid2.Text = Form4.Adodc1.Recordset.Fields!mavad
      Adodc2.Recordset.MoveNext
    Loop Until Adodc2.Recordset.EOF = True
  End If
  DataGrid2.Refresh
  
  Adodc3.ConnectionString = Form3.Text10.Text
  Adodc3.CommandType = adCmdUnknown
  Adodc3.RecordSource = "select * from ozanmasir where (idmahsol=" + Trim(Str(Combo3.List(Combo1.ListIndex))) + ") and (rad = " + Trim(Str(q)) + ")"
  
  Adodc3.Refresh
  DataGrid3.Refresh
  
  Adodc3.Refresh
  DataGrid3.Refresh
  
  If Adodc3.Recordset.RecordCount > 0 Then
    Adodc3.Recordset.MoveFirst
    Do
      w = -1
      For q = 0 To List4.ListCount - 1
        If List4.List(q) = Adodc3.Recordset.Fields!Name Then
          w = q
        End If
      Next q
      If w <> -1 Then List4.RemoveItem w
      Adodc3.Recordset.MoveNext
    Loop Until Adodc3.Recordset.EOF = True
  End If
  DataGrid3.Refresh
  
  If IsNull(Adodc1.Recordset.Fields!nomes) = False Then
    If Adodc1.Recordset.Fields!nomes = "„”  «»ÌœÂ ‘œÂ" Then
      Option1.Value = True
      db1.Open Form3.Text10.Text
        rs1.Open "SELECT Count(mojodiavalmeghdar) As rsnumber From NewMes WHERE (idmahsol=" + Trim(Str(Adodc1.Recordset.Fields!idmahsol)) + ") And (rad=" + Trim(Str(Adodc1.Recordset.Fields!rad)) + ")", db1
          If rs1.Fields!rsnumber > 0 Then
            rs2.Open "SELECT * From NewMes WHERE (idmahsol=" + Trim(Str(Adodc1.Recordset.Fields!idmahsol)) + ") And (rad=" + Trim(Str(Adodc1.Recordset.Fields!rad)) + ")", db1
              Text1(6).Text = rs2.Fields!mojodiavalmeghdar
              Text1(7).Text = rs2.Fields!mojodiavalmemoney
              Text1(8).Text = rs2.Fields!tolidteydoremeghdar
              Text1(9).Text = rs2.Fields!tolidteydoremoney
              Text1(10).Text = rs2.Fields!naghlbebadmeghdar
              Text1(11).Text = rs2.Fields!naghlbebadmoney
              Text1(12).Text = rs2.Fields!mojodiendmeghdar
              Text1(13).Text = rs2.Fields!mojodiendmoney
            rs2.Close
            Label5.Caption = 2
          Else
            For q = 6 To 13
               Text1(q).Text = ""
            Next q
            Label5.Caption = 1
          End If
        rs1.Close
      db1.Close
    End If
    If Adodc1.Recordset.Fields!nomes = "„”  Ê·Ìœ ‘œÂ" Then
      Option2.Value = True
    End If
    
    If Adodc1.Recordset.Fields!nomes = "¬·Ê„Ì‰ÌÊ„" Then
      Option3.Value = True
    End If
  End If
End If
End Sub

Private Sub DataGrid2_BeforeDelete(Cancel As Integer)
Label2(13).Caption = Val(Label2(13).Caption) - Val(Adodc2.Recordset.Fields!meghdar)
End Sub

Private Sub DataGrid3_BeforeDelete(Cancel As Integer)
Dim db1 As New ADODB.Connection
Dim rs(10) As New ADODB.Recordset

If Adodc3.Recordset.RecordCount > 0 Then
    List4.AddItem Adodc3.Recordset.Fields!Name
    Select Case Adodc3.Recordset.Fields!Name

      Case " «»"
        db1.Open Form3.Text10.Text
          rs(0).Open "DELETE From Taab WHERE (idmahsol=" + Trim(Str(Adodc1.Recordset.Fields!idmahsol)) + ") and (rad=" + Trim(Str(Adodc1.Recordset.Fields!rad)) + ")", db1
        db1.Close

      Case "«” —‰œ— 4 + 1"
        db1.Open Form3.Text10.Text
          rs(0).Open "DELETE From Sterander1_4 WHERE (idmahsol=" + Trim(Str(Adodc1.Recordset.Fields!idmahsol)) + ") and (rad=" + Trim(Str(Adodc1.Recordset.Fields!rad)) + ")", db1
        db1.Close
        
      Case "«” —‰œ— 6 +1"
        db1.Open Form3.Text10.Text
          rs(0).Open "DELETE From Sterander1_6 WHERE (idmahsol=" + Trim(Str(Adodc1.Recordset.Fields!idmahsol)) + ") and (rad=" + Trim(Str(Adodc1.Recordset.Fields!rad)) + ")", db1
        db1.Close
      
      Case "«” —‰œ— 36 + 1"
        db1.Open Form3.Text10.Text
          rs(0).Open "DELETE From Sterander1_36 WHERE (idmahsol=" + Trim(Str(Adodc1.Recordset.Fields!idmahsol)) + ") and (rad=" + Trim(Str(Adodc1.Recordset.Fields!rad)) + ")", db1
        db1.Close
        
      Case "œ—«„  ÊÌ” —"
        db1.Open Form3.Text10.Text
          rs(0).Open "DELETE From DramToester WHERE (idmahsol=" + Trim(Str(Adodc1.Recordset.Fields!idmahsol)) + ") and (rad=" + Trim(Str(Adodc1.Recordset.Fields!rad)) + ")", db1
        db1.Close

      Case "„Œ«»—« Ì"
        db1.Open Form3.Text10.Text
          rs(0).Open "DELETE From Mokhaberat WHERE (idmahsol=" + Trim(Str(Adodc1.Recordset.Fields!idmahsol)) + ") and (rad=" + Trim(Str(Adodc1.Recordset.Fields!rad)) + ")", db1
        db1.Close
        
      Case "«ò” —Êœ—"
        db1.Open Form3.Text10.Text
          rs(0).Open "DELETE From Exteroder WHERE (idmahsol=" + Trim(Str(Adodc1.Recordset.Fields!idmahsol)) + ") and (rad=" + Trim(Str(Adodc1.Recordset.Fields!rad)) + ")", db1
        db1.Close
        
      Case "»” Â »‰œÌ"
        db1.Open Form3.Text10.Text
          rs(0).Open "DELETE  From Bastebandi WHERE (idmahsol=" + Trim(Str(Adodc1.Recordset.Fields!idmahsol)) + ") and (rad=" + Trim(Str(Adodc1.Recordset.Fields!rad)) + ")", db1
        db1.Close
        
      Case "«‰»«— „Õ’Ê·"
        db1.Open Form3.Text10.Text
          rs(0).Open "DELETE  From AnbarMahsol WHERE (idmahsol=" + Trim(Str(Adodc1.Recordset.Fields!idmahsol)) + ") and (rad=" + Trim(Str(Adodc1.Recordset.Fields!rad)) + ")", db1
        db1.Close
        
    End Select
End If
End Sub

Private Sub Form_Activate()
commove = 0

Form2.Adodc1.ConnectionString = Form3.Text10.Text
Form2.Adodc1.CommandType = adCmdUnknown
Form2.Adodc1.RecordSource = "SELECT * FROM infoMahsol"
Form2.Adodc1.Refresh

If Form2.Adodc1.Recordset.RecordCount > 0 Then
  Combo1.Clear
  Combo3.Clear
  Form2.Adodc1.Recordset.Sort = "idmahsol"
  Form2.Adodc1.Recordset.MoveFirst
  Do
    Combo1.AddItem Form2.Adodc1.Recordset.Fields!mahsol
    Combo3.AddItem Form2.Adodc1.Recordset.Fields!idmahsol
    Form2.Adodc1.Recordset.MoveNext
  Loop Until Form2.Adodc1.Recordset.EOF = True
End If

Form29.Adodc1.ConnectionString = Form3.Text10.Text
Form29.Adodc1.CommandType = adCmdUnknown
Form29.Adodc1.RecordSource = "SELECT * FROM infogher"
Form29.Adodc1.Refresh

If Form29.Adodc1.Recordset.RecordCount > 0 Then
  Combo7.Clear
  Combo6.Clear
  Form29.Adodc1.Recordset.Sort = "id"
  Form29.Adodc1.Recordset.MoveFirst
  Do
    Combo7.AddItem Form29.Adodc1.Recordset.Fields!Name
    Combo6.AddItem Form29.Adodc1.Recordset.Fields!ID
    Form29.Adodc1.Recordset.MoveNext
  Loop Until Form29.Adodc1.Recordset.EOF = True
End If

Form4.Adodc1.ConnectionString = Form3.Text10.Text
Form4.Adodc1.CommandType = adCmdUnknown
Form4.Adodc1.RecordSource = "select * from infomavad where (nosim='1') ORDER BY idmavad"
Form4.Adodc1.Refresh
If Form4.Adodc1.Recordset.RecordCount > 0 Then
  Combo2.Clear
  Combo4.Clear
  Combo5.Clear
  Form4.Adodc1.Recordset.Sort = "idmavad"
  Form4.Adodc1.Recordset.MoveFirst
  Do
    Combo2.AddItem Form4.Adodc1.Recordset.Fields!mavad
    Combo4.AddItem Form4.Adodc1.Recordset.Fields!idmavad
    Combo5.AddItem Form4.Adodc1.Recordset.Fields!zarib
    Form4.Adodc1.Recordset.MoveNext
  Loop Until Form4.Adodc1.Recordset.EOF = True
End If

Adodc1.ConnectionString = Form3.Text10.Text
Adodc1.CommandType = adCmdUnknown
Adodc1.RecordSource = "select * from ozanmain WHERE rad=0"
Adodc1.Refresh

Adodc2.ConnectionString = Form3.Text10.Text
Adodc2.CommandType = adCmdUnknown
Adodc2.RecordSource = "select * from ozanunder WHERE rad=0"
Adodc2.Refresh

Adodc3.ConnectionString = Form3.Text10.Text
Adodc3.CommandType = adCmdUnknown
Adodc3.RecordSource = "select * from ozanmasir WHERE rad=0"
Adodc3.Refresh

Form4.Adodc1.ConnectionString = Form3.Text10.Text
Form4.Adodc1.CommandType = adCmdUnknown
Form4.Adodc1.RecordSource = "select * from infomavad "
Form4.Adodc1.Refresh

Label2(13).Caption = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
Form3.Show
Me.Hide
End Sub

Private Sub List4_DblClick()
If List4.ListIndex <> -1 Then
  q = Adodc3.Recordset.RecordCount
  Adodc3.Refresh
  Adodc3.Recordset.AddNew
  Adodc3.Recordset.Fields!idmahsol = Adodc1.Recordset.Fields!idmahsol
  Adodc3.Recordset.Fields!rad = Adodc1.Recordset.Fields!rad
  Adodc3.Recordset.Fields!rad1 = q + 1
  Adodc3.Recordset.Fields!Name = List4.List(List4.ListIndex)
  List4.RemoveItem (List4.ListIndex)
  Adodc3.Recordset.Update
End If
End Sub

Private Sub mnubasteband_Click()
Form29.Show
End Sub

Private Sub mnumahsol_Click()
Form2.Show
End Sub

Private Sub mnumavadmasraf_Click()
Form4.Show
End Sub

Private Sub Option1_Click()
  For q = 6 To 13
    Text1(q).Enabled = True
  Next q
  For q = 20 To 27
    Label2(q).Enabled = True
  Next q
End Sub

Private Sub Option2_Click()
  For q = 6 To 13
    Text1(q).Enabled = False
  Next q
  For q = 20 To 27
    Label2(q).Enabled = False
  Next q
End Sub

Private Sub Text1_Change(Index As Integer)
If (Index >= 6) And (Index <= 13) Then
  Text1(10).Text = (Val(Text1(8).Text) + Val(Text1(6).Text)) - Val(Text1(12).Text)
  
  If (Val(Text1(6).Text) + Val(Text1(8).Text)) <> 0 Then
    q1 = (Val(Text1(7).Text) + Val(Text1(9).Text)) / (Val(Text1(6).Text) + Val(Text1(8).Text))
    Text1(11).Text = Round(Val(Text1(10).Text) * q1)
    Text1(13).Text = Round(Val(Text1(12).Text) * q1)
  End If

End If
End Sub

Private Sub Text1_GotFocus(Index As Integer)
Text1(Index).SelStart = 0
Text1(Index).SelLength = Len(Text1(Index).Text)
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
  Select Case Index
    Case 3
      Text1(4).SetFocus
      
    Case 4
      Text1(0).SetFocus
      
    Case 0
      Text1(2).SetFocus
      
    Case 2
      Text1(1).SetFocus
      
    Case 1
      Combo7.SetFocus
      
    Case 5
      Command3.SetFocus
      
    Case 6
      Text1(7).SetFocus
      
    Case 7
      Text1(8).SetFocus
      
    Case 8
      Text1(9).SetFocus
      
    Case 9
      Text1(10).SetFocus
      
    Case 10
      Text1(11).SetFocus
      
    Case 11
      Text1(12).SetFocus
      
    Case 12
      Text1(13).SetFocus
      
    Case 13
      Command2.SetFocus
      
  End Select
End If
End Sub
