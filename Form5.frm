VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form5 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "⁄„·ò—œ ò«·«"
   ClientHeight    =   10290
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11160
   BeginProperty Font 
      Name            =   "B Titr"
      Size            =   9.75
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form5.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   10290
   ScaleWidth      =   11160
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "ç«Å"
      Height          =   495
      Left            =   360
      RightToLeft     =   -1  'True
      TabIndex        =   65
      Top             =   9600
      Width           =   3495
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   1320
      Top             =   0
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
      Caption         =   "Adodc3"
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
   Begin VB.ComboBox Combo8 
      Height          =   465
      Left            =   7920
      RightToLeft     =   -1  'True
      TabIndex        =   62
      Text            =   "Combo3"
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.ComboBox Combo7 
      Height          =   465
      Left            =   8040
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   61
      Top             =   120
      Width           =   1935
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   3960
      Top             =   0
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
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
      Connect         =   ""
      OLEDBString     =   ""
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
   Begin VB.Frame Frame5 
      Caption         =   "„ÊÃÊœÌ ¬Œ— œÊ—Â"
      Height          =   4695
      Left            =   360
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   4200
      Width           =   3495
      Begin VB.Frame Frame9 
         Caption         =   "«‰»«— œÊ„"
         Height          =   1575
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   54
         Top             =   2160
         Width           =   3015
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   495
            Index           =   12
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   57
            Top             =   960
            Width           =   2055
         End
         Begin VB.ComboBox Combo6 
            Height          =   465
            ItemData        =   "Form5.frx":2CFA
            Left            =   1440
            List            =   "Form5.frx":2D0A
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   56
            Top             =   360
            Width           =   735
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   495
            Index           =   10
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   55
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "„ÊÃÊœÌ"
            Height          =   495
            Index           =   22
            Left            =   2160
            RightToLeft     =   -1  'True
            TabIndex        =   60
            Top             =   960
            Width           =   735
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Ê«Õœ"
            Height          =   495
            Index           =   21
            Left            =   2160
            RightToLeft     =   -1  'True
            TabIndex        =   59
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "„ﬁœ«—"
            Height          =   495
            Index           =   20
            Left            =   840
            RightToLeft     =   -1  'True
            TabIndex        =   58
            Top             =   360
            Width           =   495
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "«‰»«— «Ê·"
         Height          =   1575
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   47
         Top             =   480
         Width           =   3015
         Begin VB.ComboBox Combo5 
            Height          =   465
            ItemData        =   "Form5.frx":2D26
            Left            =   1440
            List            =   "Form5.frx":2D36
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   50
            Top             =   360
            Width           =   735
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   495
            Index           =   9
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   49
            Top             =   360
            Width           =   735
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   495
            Index           =   8
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   48
            Top             =   960
            Width           =   2055
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Ê«Õœ"
            Height          =   495
            Index           =   17
            Left            =   2160
            RightToLeft     =   -1  'True
            TabIndex        =   53
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "„ﬁœ«—"
            Height          =   495
            Index           =   16
            Left            =   840
            RightToLeft     =   -1  'True
            TabIndex        =   52
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "„ÊÃÊœÌ"
            Height          =   495
            Index           =   9
            Left            =   2160
            RightToLeft     =   -1  'True
            TabIndex        =   51
            Top             =   960
            Width           =   735
         End
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Ã„⁄ „ —«é :"
         Height          =   495
         Index           =   18
         Left            =   2040
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   3960
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Height          =   495
         Index           =   54
         Left            =   480
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   3960
         Width           =   1575
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Õ–›"
      Height          =   495
      Left            =   2160
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   9000
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "À» "
      Height          =   495
      Left            =   360
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   9000
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "„ÊÃÊœÌ «Ê· œÊ—Â"
      Height          =   5295
      Left            =   7320
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   4200
      Width           =   3495
      Begin VB.Frame Frame7 
         Caption         =   "«‰»«— œÊ„"
         Height          =   1575
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   2160
         Width           =   3015
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   495
            Index           =   11
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   36
            Top             =   360
            Width           =   735
         End
         Begin VB.ComboBox Combo2 
            Height          =   465
            ItemData        =   "Form5.frx":2D52
            Left            =   1440
            List            =   "Form5.frx":2D62
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   35
            Top             =   360
            Width           =   735
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   495
            Index           =   2
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   33
            Top             =   960
            Width           =   2055
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "„ﬁœ«—"
            Height          =   495
            Index           =   11
            Left            =   840
            RightToLeft     =   -1  'True
            TabIndex        =   38
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Ê«Õœ"
            Height          =   495
            Index           =   10
            Left            =   2160
            RightToLeft     =   -1  'True
            TabIndex        =   37
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "„ÊÃÊœÌ"
            Height          =   495
            Index           =   1
            Left            =   2160
            RightToLeft     =   -1  'True
            TabIndex        =   34
            Top             =   960
            Width           =   735
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "«‰»«— «Ê·"
         Height          =   1575
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   480
         Width           =   3015
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   495
            Index           =   1
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   28
            Top             =   960
            Width           =   2055
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   495
            Index           =   0
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   27
            Top             =   360
            Width           =   735
         End
         Begin VB.ComboBox Combo1 
            Height          =   465
            ItemData        =   "Form5.frx":2D7E
            Left            =   1440
            List            =   "Form5.frx":2D8E
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   26
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "„ÊÃÊœÌ"
            Height          =   495
            Index           =   4
            Left            =   2160
            RightToLeft     =   -1  'True
            TabIndex        =   31
            Top             =   960
            Width           =   735
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "„ﬁœ«—"
            Height          =   495
            Index           =   8
            Left            =   840
            RightToLeft     =   -1  'True
            TabIndex        =   30
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Ê«Õœ"
            Height          =   495
            Index           =   7
            Left            =   2160
            RightToLeft     =   -1  'True
            TabIndex        =   29
            Top             =   360
            Width           =   735
         End
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   495
         Index           =   3
         Left            =   480
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   3960
         Width           =   1575
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Height          =   495
         Index           =   50
         Left            =   480
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   4560
         Width           =   1575
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Ã„⁄ „»·€"
         Height          =   495
         Index           =   6
         Left            =   2040
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   3960
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Ã„⁄ „ —«é :"
         Height          =   495
         Index           =   3
         Left            =   2040
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   4560
         Width           =   975
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "›—Ê‘"
      Height          =   5295
      Left            =   3960
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   4200
      Width           =   3255
      Begin VB.Frame Frame3 
         Caption         =   "›—Ê‘ œ— œ”  «ﬁœ«„"
         Height          =   2175
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   2520
         Width           =   3015
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   495
            Index           =   6
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   44
            Top             =   360
            Width           =   735
         End
         Begin VB.ComboBox Combo4 
            Height          =   465
            ItemData        =   "Form5.frx":2DAA
            Left            =   1440
            List            =   "Form5.frx":2DBA
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   43
            Top             =   360
            Width           =   735
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   495
            Index           =   7
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   14
            Top             =   960
            Width           =   2175
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "„ﬁœ«—"
            Height          =   495
            Index           =   14
            Left            =   840
            RightToLeft     =   -1  'True
            TabIndex        =   46
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Ê«Õœ"
            Height          =   495
            Index           =   0
            Left            =   2160
            RightToLeft     =   -1  'True
            TabIndex        =   45
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "„ —«é"
            Height          =   495
            Index           =   19
            Left            =   2280
            RightToLeft     =   -1  'True
            TabIndex        =   17
            Top             =   960
            Width           =   615
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Ã„⁄ ›—Ê‘ :"
            Height          =   495
            Index           =   15
            Left            =   1920
            RightToLeft     =   -1  'True
            TabIndex        =   16
            Top             =   1560
            Width           =   975
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Height          =   495
            Index           =   52
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   15
            Top             =   1560
            Width           =   1815
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "›—Ê‘ ÿÌ œÊ—Â"
         Height          =   2175
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   360
         Width           =   3015
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   495
            Index           =   4
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   40
            Top             =   360
            Width           =   735
         End
         Begin VB.ComboBox Combo3 
            Height          =   465
            ItemData        =   "Form5.frx":2DD6
            Left            =   1440
            List            =   "Form5.frx":2DE6
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   39
            Top             =   360
            Width           =   735
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   495
            Index           =   5
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   6
            Top             =   960
            Width           =   2175
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "„ﬁœ«—"
            Height          =   495
            Index           =   13
            Left            =   840
            RightToLeft     =   -1  'True
            TabIndex        =   42
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Ê«Õœ"
            Height          =   495
            Index           =   12
            Left            =   2160
            RightToLeft     =   -1  'True
            TabIndex        =   41
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Height          =   495
            Index           =   51
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   12
            Top             =   1560
            Width           =   1815
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Ã„⁄ ›—Ê‘ :"
            Height          =   495
            Index           =   2
            Left            =   1920
            RightToLeft     =   -1  'True
            TabIndex        =   8
            Top             =   1560
            Width           =   975
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "„ —«é"
            Height          =   495
            Index           =   5
            Left            =   2280
            RightToLeft     =   -1  'True
            TabIndex        =   7
            Top             =   960
            Width           =   615
         End
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Ã„⁄ —Ì«·Ì :"
         Height          =   495
         Left            =   2160
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   4680
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Height          =   495
         Index           =   53
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   4680
         Width           =   2055
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   2640
      Top             =   0
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form5.frx":2E02
      Height          =   3375
      Left            =   360
      TabIndex        =   64
      Top             =   720
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   5953
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
      Caption         =   "·Ì”  ò«·« Ê „Ê«œ «Ê·ÌÂ „’—›Ì «” «‰œ«—œ"
      ColumnCount     =   6
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
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
         BeginProperty Column02 
         EndProperty
         BeginProperty Column03 
         EndProperty
         BeginProperty Column04 
         EndProperty
         BeginProperty Column05 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   330
      Left            =   0
      Top             =   0
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
      RecordSource    =   "p_amalkardkala"
      Caption         =   "Adodc3"
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
      Caption         =   "‰«„ „Õ’Ê·"
      Height          =   495
      Left            =   9960
      RightToLeft     =   -1  'True
      TabIndex        =   63
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Height          =   495
      Index           =   55
      Left            =   6840
      RightToLeft     =   -1  'True
      TabIndex        =   24
      Top             =   9600
      Width           =   2175
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Ã„⁄  Ê·Ìœ ÿÌ œÊ—Â :"
      Height          =   495
      Left            =   9120
      RightToLeft     =   -1  'True
      TabIndex        =   23
      Top             =   9600
      Width           =   1695
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim blnexist As Boolean
Dim blnexist1 As Boolean

Private Sub Combo7_Click()
Call Combo7_LostFocus
End Sub

Private Sub Combo7_LostFocus()
If Combo7.ListIndex <> -1 Then
  Adodc1.ConnectionString = Form3.Text10.Text
  Adodc1.CommandType = adCmdUnknown
  Adodc1.RecordSource = "select * from ozanmain where idmahsol=" + Combo8.List(Combo7.ListIndex)
  Adodc1.Refresh
End If
End Sub

Private Sub Command1_Click()
Adodc2.RecordSource = "SELECT * FROM amalkardkala WHERE (idmahsol=" + Trim(Str(Adodc1.Recordset.Fields!idmahsol)) + ") and (rad=" + Trim(Str(Adodc1.Recordset.Fields!rad)) + ")"
Adodc2.Refresh

If Adodc2.Recordset.RecordCount = 0 Then
  Adodc2.Recordset.AddNew
  Adodc2.Recordset.Fields!oneanonevaheh = Combo1.List(Combo1.ListIndex)
  Adodc2.Recordset.Fields!oneanonemeghdar = Text1(0).Text
  Adodc2.Recordset.Fields!oneanonemojodi = Text1(1).Text
  Adodc2.Recordset.Fields!oneantowvahed = Combo2.List(Combo2.ListIndex)
  Adodc2.Recordset.Fields!oneantowmeghdar = Text1(11).Text
  Adodc2.Recordset.Fields!oneantowmojodi = Text1(2).Text
  Adodc2.Recordset.Fields!oneansummoney = Text1(3).Text
  Adodc2.Recordset.Fields!oneansummeter = Label2(50).Caption
  Adodc2.Recordset.Fields!seldorevahed = Combo3.List(Combo3.ListIndex)
  Adodc2.Recordset.Fields!seldoremeghdar = Text1(4).Text
  Adodc2.Recordset.Fields!seldoremeter = Text1(5).Text
  Adodc2.Recordset.Fields!selcodesum = Label2(51).Caption
  Adodc2.Recordset.Fields!seleghdamvahed = Combo4.List(Combo4.ListIndex)
  Adodc2.Recordset.Fields!seleghdammeghdar = Text1(6).Text
  Adodc2.Recordset.Fields!seleghdammeter = Text1(7).Text
  Adodc2.Recordset.Fields!seleghdamsum = Label2(52).Caption
  Adodc2.Recordset.Fields!selsum = Label2(53).Caption
  Adodc2.Recordset.Fields!endanonevahed = Combo5.List(Combo5.ListIndex)
  Adodc2.Recordset.Fields!endanonemeghar = Text1(9).Text
  Adodc2.Recordset.Fields!endanonemeghdar = Text1(8).Text
  Adodc2.Recordset.Fields!endantowvahed = Combo6.List(Combo6.ListIndex)
  Adodc2.Recordset.Fields!endantowmeghar = Text1(10).Text
  Adodc2.Recordset.Fields!endantowmeghdar = Text1(12).Text
  Adodc2.Recordset.Fields!endansum = Label2(54).Caption
  Adodc2.Recordset.Fields!sumtolid = Label2(55).Caption
  Adodc2.Recordset.Fields!idmahsol = Adodc1.Recordset.Fields!idmahsol
  Adodc2.Recordset.Fields!rad = Adodc1.Recordset.Fields!rad
  Adodc2.Recordset.Update
  MsgBox "«ÿ·«⁄«  »« „Ê›ﬁÌ  À»  ‘œ", vbMsgBoxRight + vbInformation, ""
Else
  Adodc2.Recordset.Fields!oneanonevaheh = Combo1.List(Combo1.ListIndex)
  Adodc2.Recordset.Fields!oneanonemeghdar = Text1(0).Text
  Adodc2.Recordset.Fields!oneanonemojodi = Text1(1).Text
  Adodc2.Recordset.Fields!oneantowvahed = Combo2.List(Combo2.ListIndex)
  Adodc2.Recordset.Fields!oneantowmeghdar = Text1(11).Text
  Adodc2.Recordset.Fields!oneantowmojodi = Text1(2).Text
  Adodc2.Recordset.Fields!oneansummoney = Text1(3).Text
  Adodc2.Recordset.Fields!oneansummeter = Label2(50).Caption
  Adodc2.Recordset.Fields!seldorevahed = Combo3.List(Combo3.ListIndex)
  Adodc2.Recordset.Fields!seldoremeghdar = Text1(4).Text
  Adodc2.Recordset.Fields!seldoremeter = Text1(5).Text
  Adodc2.Recordset.Fields!selcodesum = Label2(51).Caption
  Adodc2.Recordset.Fields!seleghdamvahed = Combo4.List(Combo4.ListIndex)
  Adodc2.Recordset.Fields!seleghdammeghdar = Text1(6).Text
  Adodc2.Recordset.Fields!seleghdammeter = Text1(7).Text
  Adodc2.Recordset.Fields!seleghdamsum = Label2(52).Caption
  Adodc2.Recordset.Fields!selsum = Label2(53).Caption
  Adodc2.Recordset.Fields!endanonevahed = Combo5.List(Combo5.ListIndex)
  Adodc2.Recordset.Fields!endanonemeghar = Text1(9).Text
  Adodc2.Recordset.Fields!endanonemeghdar = Text1(8).Text
  Adodc2.Recordset.Fields!endantowvahed = Combo6.List(Combo6.ListIndex)
  Adodc2.Recordset.Fields!endantowmeghar = Text1(10).Text
  Adodc2.Recordset.Fields!endantowmeghdar = Text1(12).Text
  Adodc2.Recordset.Fields!endansum = Label2(54).Caption
  Adodc2.Recordset.Fields!sumtolid = Label2(55).Caption
  Adodc2.Recordset.Update
  MsgBox "«ÿ·«⁄«  »« „Ê›ﬁÌ   €ÌÌ— ÅÌœ« ò—œ", vbMsgBoxRight + vbInformation, ""
End If

End Sub

Private Sub Command3_Click()
Dim db1 As New ADODB.Connection
Dim rs1 As New ADODB.Recordset
db1.Open Form3.Text10.Text
rs1.Open "DELETE FROM p_amalkardkala", db1
db1.Close


Adodc2.RecordSource = "SELECT * FROM amalkardkala "
Adodc2.Refresh

Adodc2.Recordset.MoveFirst
Do
  Adodc4.Refresh
  Adodc4.Recordset.AddNew
  Adodc4.Recordset.Fields!oneanonevaheh = Adodc2.Recordset.Fields!oneanonevaheh
  Adodc4.Recordset.Fields!oneanonemeghdar = Adodc2.Recordset.Fields!oneanonemeghdar
  Adodc4.Recordset.Fields!oneanonemojodi = Adodc2.Recordset.Fields!oneanonemojodi
  Adodc4.Recordset.Fields!oneantowvahed = Adodc2.Recordset.Fields!oneantowvahed
  Adodc4.Recordset.Fields!oneantowmeghdar = Adodc2.Recordset.Fields!oneantowmeghdar
  Adodc4.Recordset.Fields!oneantowmojodi = Adodc2.Recordset.Fields!oneantowmojodi
  Adodc4.Recordset.Fields!oneansummoney = Adodc2.Recordset.Fields!oneansummoney
  Adodc4.Recordset.Fields!oneansummeter = Adodc2.Recordset.Fields!oneansummeter
  Adodc4.Recordset.Fields!seldorevahed = Adodc2.Recordset.Fields!seldorevahed
  Adodc4.Recordset.Fields!seldoremeghdar = Adodc2.Recordset.Fields!seldoremeghdar
  Adodc4.Recordset.Fields!seldoremeter = Adodc2.Recordset.Fields!seldoremeter
  Adodc4.Recordset.Fields!selcodesum = Adodc2.Recordset.Fields!selcodesum
  Adodc4.Recordset.Fields!seleghdamvahed = Adodc2.Recordset.Fields!seleghdamvahed
  Adodc4.Recordset.Fields!seleghdammeghdar = Adodc2.Recordset.Fields!seleghdammeghdar
  Adodc4.Recordset.Fields!seleghdammeter = Adodc2.Recordset.Fields!seleghdammeter
  Adodc4.Recordset.Fields!seleghdamsum = Adodc2.Recordset.Fields!seleghdamsum
  Adodc4.Recordset.Fields!selsum = Adodc2.Recordset.Fields!selsum
  Adodc4.Recordset.Fields!endanonevahed = Adodc2.Recordset.Fields!endanonevahed
  Adodc4.Recordset.Fields!endanonemeghar = Adodc2.Recordset.Fields!endanonemeghar
  Adodc4.Recordset.Fields!endanonemeghdar = Adodc2.Recordset.Fields!endanonemeghdar
  Adodc4.Recordset.Fields!endantowvahed = Adodc2.Recordset.Fields!endantowvahed
  Adodc4.Recordset.Fields!endantowmeghar = Adodc2.Recordset.Fields!endantowmeghar
  Adodc4.Recordset.Fields!endantowmeghdar = Adodc2.Recordset.Fields!endantowmeghdar
  Adodc4.Recordset.Fields!endansum = Adodc2.Recordset.Fields!endansum
  Adodc4.Recordset.Fields!sumtolid = Adodc2.Recordset.Fields!sumtolid
  Adodc4.Recordset.Fields!idmahsol = Adodc2.Recordset.Fields!idmahsol
  Adodc4.Recordset.Fields!rad = Adodc2.Recordset.Fields!rad
  Adodc3.RecordSource = "select * from ozanmain where (idmahsol=" + Trim(Str(Adodc2.Recordset.Fields!idmahsol)) + ") and (rad=" + Trim(Str(Adodc2.Recordset.Fields!rad)) + ")"
  Adodc3.Refresh
  If Adodc3.Recordset.RecordCount > 0 Then
    Adodc4.Recordset.Fields!propertikhas = Adodc3.Recordset.Fields!propertikhas
    Adodc4.Recordset.Fields!Size = Adodc3.Recordset.Fields!Size
    Adodc4.Recordset.Fields!kodemahsol = Adodc3.Recordset.Fields!kodemahsol
    Adodc4.Recordset.Fields!nomahsol = Adodc3.Recordset.Fields!nomahsol
    Adodc4.Recordset.Fields!gothr = Adodc3.Recordset.Fields!gothr
    Form2.Adodc1.Recordset.Find "idmahsol=" + Trim(Str(Adodc2.Recordset.Fields!idmahsol)), , adSearchForward, 1
    Adodc4.Recordset.Fields!namemahsol = Form2.Adodc1.Recordset.Fields!mahsol
  Else
    Adodc4.Recordset.Fields!propertikhas = ""
    Adodc4.Recordset.Fields!Size = ""
    Adodc4.Recordset.Fields!kodemahsol = ""
    Adodc4.Recordset.Fields!nomahsol = ""
    Adodc4.Recordset.Fields!gothr = ""
    Adodc4.Recordset.Fields!namemahsol = ""
  End If
  Adodc4.Recordset.Update
  Adodc2.Recordset.MoveNext
Loop Until Adodc2.Recordset.EOF = True
Form49.Show
End Sub

Private Sub DataGrid1_Click()
If Adodc1.Recordset.RecordCount > 0 Then
For q = 0 To 12
  Text1(q).Text = 0
Next q
For q = 50 To 55
  Label2(q).Caption = 0
Next q

Adodc2.RecordSource = "SELECT * FROM amalkardkala WHERE (idmahsol=" + Trim(Str(Adodc1.Recordset.Fields!idmahsol)) + ") and (rad=" + Trim(Str(Adodc1.Recordset.Fields!rad)) + ")"
Adodc2.Refresh

If Adodc2.Recordset.RecordCount > 0 Then
  For q = 0 To Combo1.ListCount - 1
     If Combo1.List(q) = Adodc2.Recordset.Fields!oneanonevaheh Then Combo1.ListIndex = q: Exit For
  Next q
  Text1(0).Text = Adodc2.Recordset.Fields!oneanonemeghdar
  Text1(1).Text = Adodc2.Recordset.Fields!oneanonemojodi
  For q = 0 To Combo2.ListCount - 1
     If Combo2.List(q) = Adodc2.Recordset.Fields!oneantowvahed Then Combo2.ListIndex = q: Exit For
  Next q
  Text1(11).Text = Adodc2.Recordset.Fields!oneantowmeghdar
  Text1(2).Text = Adodc2.Recordset.Fields!oneantowmojodi
  Text1(3).Text = Adodc2.Recordset.Fields!oneansummoney
  Label2(50).Caption = Adodc2.Recordset.Fields!oneansummeter
  For q = 0 To Combo3.ListCount - 1
     If Combo3.List(q) = Adodc2.Recordset.Fields!seldorevahed Then Combo3.ListIndex = q: Exit For
  Next q
  Text1(4).Text = Adodc2.Recordset.Fields!seldoremeghdar
  Text1(5).Text = Adodc2.Recordset.Fields!seldoremeter
  Label2(51).Caption = Adodc2.Recordset.Fields!selcodesum
  For q = 0 To Combo4.ListCount - 1
     If Combo4.List(q) = Adodc2.Recordset.Fields!seleghdamvahed Then Combo4.ListIndex = q: Exit For
  Next q
  Text1(6).Text = Adodc2.Recordset.Fields!seleghdammeghdar
  Text1(7).Text = Adodc2.Recordset.Fields!seleghdammeter
  Label2(52).Caption = Adodc2.Recordset.Fields!seleghdamsum
  Label2(53).Caption = Adodc2.Recordset.Fields!selsum
  For q = 0 To Combo5.ListCount - 1
     If Combo5.List(q) = Adodc2.Recordset.Fields!endanonevahed Then Combo5.ListIndex = q: Exit For
  Next q
  Text1(9).Text = Adodc2.Recordset.Fields!endanonemeghar
  Text1(8).Text = Adodc2.Recordset.Fields!endanonemeghdar
  For q = 0 To Combo6.ListCount - 1
     If Combo6.List(q) = Adodc2.Recordset.Fields!endantowvahed Then Combo6.ListIndex = q: Exit For
  Next q
  Text1(10).Text = Adodc2.Recordset.Fields!endantowmeghar
  Text1(12).Text = Adodc2.Recordset.Fields!endantowmeghdar
  Label2(54).Caption = Adodc2.Recordset.Fields!endansum
  Label2(55).Caption = Adodc2.Recordset.Fields!sumtolid
End If
End If
End Sub

Private Sub Form_Activate()
Form2.Adodc1.ConnectionString = Form3.Text10.Text
Form2.Adodc1.CommandType = adCmdUnknown
Form2.Adodc1.RecordSource = "select * from infoMahsol"
Form2.Adodc1.Refresh

If Form2.Adodc1.Recordset.RecordCount > 0 Then
  Combo7.Clear
  Combo8.Clear
  Form2.Adodc1.Recordset.Sort = "idmahsol"
  Form2.Adodc1.Recordset.MoveFirst
  Do
    Combo7.AddItem Form2.Adodc1.Recordset.Fields!mahsol
    Combo8.AddItem Form2.Adodc1.Recordset.Fields!idmahsol
    Form2.Adodc1.Recordset.MoveNext
  Loop Until Form2.Adodc1.Recordset.EOF = True
End If
For q = 0 To 12
  Text1(q).Text = 0
Next q
For q = 50 To 55
  Label2(q).Caption = 0
Next q
Adodc1.ConnectionString = Form3.Text10.Text
Adodc1.CommandType = adCmdUnknown
Adodc1.RecordSource = "select * from ozanmain WHERE rad=0"
Adodc1.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
Form3.Show
End Sub

Private Sub Text1_Change(Index As Integer)
Label2(50).Caption = (Val(Text1(1).Text) * Val(Text1(0).Text)) + (Val(Text1(2).Text) * Val(Text1(11).Text))
Label2(51).Caption = (Val(Text1(5).Text) * Val(Text1(4).Text))
Label2(52).Caption = (Val(Text1(6).Text) * Val(Text1(7).Text))
Label2(53).Caption = Val(Label2(52).Caption) + Val(Label2(51).Caption)
Label2(54).Caption = (Val(Text1(8).Text) * Val(Text1(9).Text)) + (Val(Text1(10).Text) * Val(Text1(12).Text))
Label2(55).Caption = (Val(Label2(51).Caption) + Val(Label2(52).Caption) + Val(Label2(54).Caption)) - Val(Label2(50).Caption)
End Sub

Private Sub Text1_GotFocus(Index As Integer)
Text1(Index).SelStart = 0
Text1(Index).SelLength = Len(Text1(Index).Text)
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
If ((KeyAscii >= 48) And (KeyAscii <= 57)) Or KeyAscii = 8 Then
Else
  KeyAscii = 0
End If
End Sub
