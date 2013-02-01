VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.4#0"; "CODEJO~3.OCX"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{A7960112-5DC4-4575-BFA3-DAD80FEE0438}#33.0#0"; "BasKomponen.ocx"
Begin VB.Form FrmUser 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5010
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8700
   LinkTopic       =   "Form1"
   ScaleHeight     =   5010
   ScaleWidth      =   8700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc AdoCabang 
      Height          =   375
      Left            =   4080
      Top             =   4320
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
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin BasKomponen.BasForm BasForm1 
      Height          =   4080
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7035
      _ExtentX        =   12409
      _ExtentY        =   7197
      ButtonMax       =   0   'False
      ButtonMin       =   0   'False
      Caption         =   ":: User Management ::"
      Object.ToolTipText     =   ":: User Management ::"
      Begin TrueOleDBGrid70.TDBGrid Grid 
         Height          =   2415
         Left            =   6120
         TabIndex        =   1
         ToolTipText     =   "Double klik untuk edit data"
         Top             =   480
         Width           =   6585
         _ExtentX        =   11615
         _ExtentY        =   4260
         _LayoutType     =   0
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).DataField=   ""
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).DataField=   ""
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   2
         Splits(0)._UserFlags=   0
         Splits(0).PartialRightColumn=   0   'False
         Splits(0).AnchorRightColumn=   -1  'True
         Splits(0).MarqueeStyle=   2
         Splits(0).AllowRowSizing=   0   'False
         Splits(0).RecordSelectorWidth=   529
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).AllowColSelect=   0   'False
         Splits(0).FetchRowStyle=   -1  'True
         Splits(0).AlternatingRowStyle=   -1  'True
         Splits(0).DividerColor=   8421376
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=2"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=3281"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=3175"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=516"
         Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(7)=   "Column(1).Width=3281"
         Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=3175"
         Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=516"
         Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         AllowUpdate     =   0   'False
         Appearance      =   2
         ColumnFooters   =   -1  'True
         DefColWidth     =   0
         HeadLines       =   2
         FootLines       =   1
         TabAction       =   2
         WrapCellPointer =   -1  'True
         MultipleLines   =   0
         CellTipsWidth   =   0
         GroupByCaption  =   "Keterangan"
         DeadAreaBackColor=   14215660
         RowDividerColor =   8454143
         RowSubDividerColor=   14215660
         DirectionAfterEnter=   1
         MaxRows         =   250000
         ViewColumnCaptionWidth=   0
         ViewColumnWidth =   0
         _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&H7DBBFF&,.bold=-1,.fontsize=825"
         _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=-1,.fontsize=825,.italic=0"
         _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
         _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=-1,.fontsize=825,.italic=0"
         _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
         _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36,.borderColor=&H80000013&"
         _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38,.bgcolor=&HFF8080&"
         _StyleDefs(20)  =   ":id=8,.fgcolor=&H80000012&"
         _StyleDefs(21)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(22)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(23)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(24)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(25)  =   "Splits(0).Style:id=13,.parent=1,.bgcolor=&H80000005&,.fgcolor=&H0&,.bold=0"
         _StyleDefs(26)  =   ":id=13,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(27)  =   ":id=13,.fontname=Verdana"
         _StyleDefs(28)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(29)  =   "Splits(0).HeadingStyle:id=14,.parent=2,.namedParent=37,.bgcolor=&H555555&"
         _StyleDefs(30)  =   ":id=14,.fgcolor=&H37D7FF&,.bold=-1,.fontsize=600,.italic=0,.underline=0"
         _StyleDefs(31)  =   ":id=14,.strikethrough=0,.charset=255"
         _StyleDefs(32)  =   ":id=14,.fontname=Terminal"
         _StyleDefs(33)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(34)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(35)  =   "Splits(0).SelectedStyle:id=18,.parent=6,.fgcolor=&HFFFF&,.borderColor=&H80FF80&"
         _StyleDefs(36)  =   "Splits(0).EditorStyle:id=17,.parent=7,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(37)  =   ":id=17,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(38)  =   ":id=17,.fontname=MS Sans Serif"
         _StyleDefs(39)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8,.namedParent=38,.fgcolor=&HFFFF&"
         _StyleDefs(40)  =   "Splits(0).EvenRowStyle:id=20,.parent=9,.bgcolor=&HFFFF&"
         _StyleDefs(41)  =   "Splits(0).OddRowStyle:id=21,.parent=10,.namedParent=37,.bgcolor=&H80FFFF&"
         _StyleDefs(42)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(43)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(44)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
         _StyleDefs(45)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
         _StyleDefs(49)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(52)  =   "Named:id=33:Normal"
         _StyleDefs(53)  =   ":id=33,.parent=0,.bgcolor=&HFF80&,.fgcolor=&HFFFFFF&,.borderColor=&H800040&"
         _StyleDefs(54)  =   "Named:id=34:Heading"
         _StyleDefs(55)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(56)  =   ":id=34,.wraptext=-1"
         _StyleDefs(57)  =   "Named:id=35:Footing"
         _StyleDefs(58)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(59)  =   ":id=35,.wraptext=0,.locked=0"
         _StyleDefs(60)  =   "Named:id=36:Selected"
         _StyleDefs(61)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(62)  =   ":id=36,.borderColor=&H80000013&"
         _StyleDefs(63)  =   "Named:id=37:Caption"
         _StyleDefs(64)  =   ":id=37,.parent=34,.alignment=2,.bgcolor=&H80000009&"
         _StyleDefs(65)  =   "Named:id=38:HighlightRow"
         _StyleDefs(66)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&HA00000&,.borderColor=&H800040&"
         _StyleDefs(67)  =   "Named:id=39:EvenRow"
         _StyleDefs(68)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(69)  =   "Named:id=40:OddRow"
         _StyleDefs(70)  =   ":id=40,.parent=33,.bgcolor=&H4000&"
         _StyleDefs(71)  =   "Named:id=41:RecordSelector"
         _StyleDefs(72)  =   ":id=41,.parent=34"
         _StyleDefs(73)  =   "Named:id=42:FilterBar"
         _StyleDefs(74)  =   ":id=42,.parent=33,.bgcolor=&HFF0000&"
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo CmbCabang 
         Height          =   345
         Left            =   2160
         TabIndex        =   16
         Tag             =   "Kode"
         Top             =   2280
         Width           =   2535
         BevelType       =   0
         _Version        =   196616
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelColorHighlight=   -2147483634
         BevelColorFace  =   -2147483627
         CheckBox3D      =   0   'False
         ForeColorEven   =   0
         BackColorEven   =   8454143
         BackColorOdd    =   65535
         RowHeight       =   423
         Columns(0).Width=   3200
         Columns(0).DataType=   8
         Columns(0).FieldLen=   4096
         _ExtentX        =   4471
         _ExtentY        =   609
         _StockProps     =   93
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo CmbType 
         Height          =   345
         Left            =   2160
         TabIndex        =   14
         Tag             =   "Kode"
         Top             =   1920
         Width           =   2535
         DataFieldList   =   "Column 0"
         AllowInput      =   0   'False
         BevelType       =   0
         _Version        =   196616
         DataMode        =   2
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Row.Count       =   2
         Row(0)          =   "Admin"
         Row(1)          =   "Kasir"
         BevelColorHighlight=   -2147483634
         BevelColorFace  =   -2147483627
         CheckBox3D      =   0   'False
         ForeColorEven   =   0
         BackColorEven   =   8454143
         BackColorOdd    =   65535
         RowHeight       =   423
         Columns(0).Width=   4419
         Columns(0).Caption=   "Satuan"
         Columns(0).Name =   "Satuan"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         _ExtentX        =   4471
         _ExtentY        =   609
         _StockProps     =   93
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox TxtNama 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2160
         TabIndex        =   12
         Top             =   1560
         Width           =   2595
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00808080&
         Height          =   855
         Left            =   360
         ScaleHeight     =   795
         ScaleWidth      =   6435
         TabIndex        =   4
         Top             =   3000
         Width           =   6495
         Begin XtremeSuiteControls.PushButton CmdAdd 
            Height          =   615
            Left            =   120
            TabIndex        =   5
            ToolTipText     =   "Tambah data"
            Top             =   80
            Width           =   615
            _Version        =   851972
            _ExtentX        =   1085
            _ExtentY        =   1085
            _StockProps     =   79
            BackColor       =   8454143
            Appearance      =   6
            Picture         =   "FrmUser.frx":0000
            BorderGap       =   0
         End
         Begin XtremeSuiteControls.PushButton CmdDelete 
            Height          =   615
            Left            =   720
            TabIndex        =   6
            ToolTipText     =   "Hapus data"
            Top             =   75
            Width           =   615
            _Version        =   851972
            _ExtentX        =   1085
            _ExtentY        =   1085
            _StockProps     =   79
            BackColor       =   8454143
            Appearance      =   6
            Picture         =   "FrmUser.frx":59AA
            BorderGap       =   0
         End
         Begin XtremeSuiteControls.PushButton CmdSave 
            Height          =   615
            Left            =   4560
            TabIndex        =   7
            ToolTipText     =   "Simpan data"
            Top             =   80
            Width           =   615
            _Version        =   851972
            _ExtentX        =   1085
            _ExtentY        =   1085
            _StockProps     =   79
            BackColor       =   8454143
            Appearance      =   6
            Picture         =   "FrmUser.frx":B354
            BorderGap       =   0
         End
         Begin XtremeSuiteControls.PushButton CmdCancel 
            Height          =   615
            Left            =   5160
            TabIndex        =   8
            ToolTipText     =   "Cancel"
            Top             =   80
            Width           =   615
            _Version        =   851972
            _ExtentX        =   1085
            _ExtentY        =   1085
            _StockProps     =   79
            BackColor       =   8454143
            Appearance      =   6
            Picture         =   "FrmUser.frx":10CFE
            BorderGap       =   0
         End
         Begin XtremeSuiteControls.PushButton CmdQuit 
            Height          =   615
            Left            =   5760
            TabIndex        =   9
            ToolTipText     =   "Exit"
            Top             =   80
            Width           =   615
            _Version        =   851972
            _ExtentX        =   1085
            _ExtentY        =   1085
            _StockProps     =   79
            BackColor       =   8454143
            Appearance      =   6
            Picture         =   "FrmUser.frx":166A8
            BorderGap       =   0
         End
      End
      Begin VB.TextBox TxtID 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   840
         Width           =   2595
      End
      Begin VB.TextBox TxtPassword 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2160
         TabIndex        =   2
         Top             =   1200
         Width           =   2595
      End
      Begin VB.Label LB_Reload 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cabang"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000CCFF&
         Height          =   210
         Index           =   4
         Left            =   480
         MousePointer    =   99  'Custom
         TabIndex        =   17
         Top             =   2280
         Width           =   615
      End
      Begin VB.Label LB_Reload 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Type User"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000CCFF&
         Height          =   210
         Index           =   3
         Left            =   480
         MousePointer    =   99  'Custom
         TabIndex        =   15
         Top             =   1920
         Width           =   840
      End
      Begin VB.Label LB_Reload 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Pengguna"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000CCFF&
         Height          =   210
         Index           =   0
         Left            =   480
         MousePointer    =   99  'Custom
         TabIndex        =   13
         Top             =   1590
         Width           =   1320
      End
      Begin VB.Label LB_Reload 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "UserID"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000CCFF&
         Height          =   210
         Index           =   2
         Left            =   480
         MousePointer    =   99  'Custom
         TabIndex        =   11
         Top             =   870
         Width           =   540
      End
      Begin VB.Label LB_Reload 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000CCFF&
         Height          =   210
         Index           =   1
         Left            =   480
         MousePointer    =   99  'Custom
         TabIndex        =   10
         Top             =   1230
         Width           =   840
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H004C4C4C&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00808080&
         Height          =   2250
         Index           =   0
         Left            =   360
         Top             =   600
         Width           =   6225
      End
   End
End
Attribute VB_Name = "FrmUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit As Boolean

Private Sub CmbCabang_DropDown()
AdoCabang.RecordSource = ""
SQL = "Select kodecabang,NamaCabang as [Nama Cabang] from cabang "
Set RsFind = DbCon.Execute(SQL)
CmbCabang.Reset
If RsFind.BOF Then Exit Sub
AdoCabang.RecordSource = SQL
AdoCabang.Refresh
With CmbCabang
        .DataSourceList = AdoCabang
        .DataFieldList = "nama cabang"
        .Columns(1).Width = 3000
        .Columns(0).Visible = False
End With
End Sub

Private Sub CmdAdd_Click()
tombol False
Edit = False
TxtID.SetFocus
End Sub

Sub RefreshData()
Set Grid.DataSource = Nothing
SQL = "SELECT UserPengguna.UserID AS [User ID], UserPengguna.Nama, UserPengguna.Kelas " & _
    " AS Type, Cabang.NamaCabang AS Cabang, UserPengguna.Pembuat, UserPengguna.TglBuat " & _
    " AS [Tanggal Buat] FROM UserPengguna INNER JOIN Cabang ON UserPengguna.Cabang = Cabang.KodeCabang"
Set Grid.DataSource = DbCon.Execute(SQL)
Grid.Refresh

Grid.Columns(0).Alignment = dbgLeft
Grid.Columns(0).Width = 3000
Grid.Columns(1).Alignment = dbgLeft
Grid.Columns(1).Width = 3000
Grid.Columns(2).Alignment = dbgLeft
Grid.Columns(2).Width = 3000
Grid.Columns(3).Alignment = dbgLeft
Grid.Columns(3).Width = 3000
Grid.Columns(4).Alignment = dbgLeft
Grid.Columns(4).Width = 3000
Grid.Columns(5).Alignment = dbgLeft
Grid.Columns(5).Width = 3000

End Sub

Private Sub CmdCancel_Click()
tombol True
RefreshData
bersih
TxtKode.Locked = False
End Sub

Private Sub CmdDelete_Click()
If MsgBox("Yakin akan menghapus data ini?" & vbCrLf & "" _
            & "USERID : " & Trim(Grid.Columns(0).Text) + vbCrLf & "" _
            & "NAMA : " & Trim(Grid.Columns(1).Text) + vbCrLf & "", _
         vbYesNo + vbQuestion) = vbYes Then
    SQL = "delete from jenis where kodejenis ='" & Trim(Grid.Columns(0).Text) & "'"
    DbCon.Execute SQL
    MsgBox "Data terhapus"
    RefreshData
End If
End Sub

Private Sub CmdEdit_Click()
Grid_DblClick
End Sub

Private Sub CmdQuit_Click()
Unload Me
End Sub

Private Sub CmdSave_Click()
Dim id, pass As String
If Trim(TxtKode) = "" Then
    MsgBox "Kode grup masih kosong."
    TxtKode.SetFocus
    Exit Sub
ElseIf Not Len(Trim(TxtKode)) = 3 Then
    MsgBox "Kode grup harus 3 huruf."
    TxtKode.SetFocus
    Exit Sub
ElseIf Trim(TxtNama) = "" Then
    MsgBox "Nama Grup masih kosong."
    TxtNama.SetFocus
    Exit Sub
End If

id = Trans.encryp_pass(25, TxtID)
pass = Trans.encryp_pass(25, TxtPassword)

If InStr(1, TxtID, "'") Then
    id = Trans.encryp_pass(25, TxtID)
ElseIf InStr(1, TxtID, ";") Then
    id = Trans.encryp_pass(25, TxtID)
ElseIf InStr(1, TxtPassword, "'") Then
    pass = Trans.encryp_pass(25, TxtID)
ElseIf InStr(1, TxtPassword, ";") Then
    pass = Trans.encryp_pass(25, TxtID)
End If


If Not Edit Then
    SQL = "insert into userpengguna (userID,Password,nama,kelas,cabang,pembuat,TglBuat) values " & _
        " ('" & Trim(TxtKode) & _
        "','" & Trim(TxtNama) & "')"
Else
    SQL = "update jenis set namaJenis ='" & Trim(TxtNama) & "' where kodejenis ='" & _
        Trim(TxtKode) & "'"
    TxtKode.Locked = False
End If
DbCon.Execute SQL
MsgBox "Data tersimpan."
RefreshData
tombol True
bersih
End Sub

Private Sub Form_Load()
Me.Height = BasForm1.Height
Me.Width = BasForm1.Width
Grid.Left = 240

CmbCabang.ZOrder vbSendToBack
CmbType.ZOrder vbSendToBack
AdoCabang.ConnectionString = ConDb
tombol True
Edit = False
RefreshData
bersih
End Sub

Function KodeAuto()
SQL = "Select kodejenis from jenis order by kodejenis Desc"
Set RsFind = DbCon.Execute(SQL)
If RsFind.BOF Then
   KodeAuto = "0001"
Else
   KodeAuto = Format(CInt(Left(RsFind!kodejenis, 4)) + 1, "0000")
End If
End Function

Sub tombol(Status As Boolean)
CmdAdd.Visible = Status
CmdDelete.Visible = Status

CmdSave.Visible = Not Status
CmdCancel.Visible = Not Status

Grid.Visible = Status
End Sub

Sub bersih()
TxtKode = ""
TxtNama = ""
End Sub

Private Sub TxtKode_KeyDown(KeyCode As Integer, Shift As Integer)
Enter KeyCode
End Sub

Private Sub TxtKode_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase((Chr(KeyAscii))))
If KeyAscii = vbKeyReturn Then
    SendKeys "{tab}"
    KeyAscii = 0
End If
End Sub

Private Sub TxtNama_KeyDown(KeyCode As Integer, Shift As Integer)
Enter KeyCode
End Sub

Private Sub TxtNama_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase((Chr(KeyAscii))))
If KeyAscii = vbKeyReturn Then
    SendKeys "{tab}"
    KeyAscii = 0
End If
End Sub


