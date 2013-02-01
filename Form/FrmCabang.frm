VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.4#0"; "CODEJO~3.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{A7960112-5DC4-4575-BFA3-DAD80FEE0438}#33.0#0"; "BasKomponen.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form FrmCabang 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4785
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8535
   LinkTopic       =   "Form1"
   ScaleHeight     =   4785
   ScaleWidth      =   8535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtKode 
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
      Left            =   6000
      TabIndex        =   0
      Top             =   4080
      Width           =   915
   End
   Begin BasKomponen.BasForm BasForm1 
      Height          =   3960
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   7755
      _ExtentX        =   13679
      _ExtentY        =   6985
      ButtonMax       =   0   'False
      ButtonMin       =   0   'False
      Caption         =   ":: Cabang ::"
      Object.ToolTipText     =   ":: Cabang ::"
      Begin TrueOleDBGrid70.TDBGrid Grid 
         Height          =   2295
         Left            =   240
         TabIndex        =   12
         ToolTipText     =   "Double klik untuk edit data"
         Top             =   480
         Width           =   7305
         _ExtentX        =   12885
         _ExtentY        =   4048
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
         TabIndex        =   1
         Top             =   720
         Width           =   2595
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00808080&
         Height          =   855
         Left            =   240
         ScaleHeight     =   795
         ScaleWidth      =   7155
         TabIndex        =   13
         Top             =   2880
         Width           =   7215
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
            Picture         =   "FrmCabang.frx":0000
            BorderGap       =   0
         End
         Begin XtremeSuiteControls.PushButton CmdEdit 
            Height          =   615
            Left            =   720
            TabIndex        =   6
            ToolTipText     =   "Edit data"
            Top             =   80
            Width           =   615
            _Version        =   851972
            _ExtentX        =   1085
            _ExtentY        =   1085
            _StockProps     =   79
            BackColor       =   8454143
            Appearance      =   6
            Picture         =   "FrmCabang.frx":59AA
            BorderGap       =   0
         End
         Begin XtremeSuiteControls.PushButton CmdDelete 
            Height          =   615
            Left            =   1320
            TabIndex        =   7
            ToolTipText     =   "Hapus data"
            Top             =   80
            Width           =   615
            _Version        =   851972
            _ExtentX        =   1085
            _ExtentY        =   1085
            _StockProps     =   79
            BackColor       =   8454143
            Appearance      =   6
            Picture         =   "FrmCabang.frx":B354
            BorderGap       =   0
         End
         Begin XtremeSuiteControls.PushButton CmdSave 
            Height          =   615
            Left            =   5280
            TabIndex        =   8
            ToolTipText     =   "Simpan data"
            Top             =   75
            Width           =   615
            _Version        =   851972
            _ExtentX        =   1085
            _ExtentY        =   1085
            _StockProps     =   79
            BackColor       =   8454143
            Appearance      =   6
            Picture         =   "FrmCabang.frx":10CFE
            BorderGap       =   0
         End
         Begin XtremeSuiteControls.PushButton CmdCancel 
            Height          =   615
            Left            =   5880
            TabIndex        =   9
            ToolTipText     =   "Cancel"
            Top             =   75
            Width           =   615
            _Version        =   851972
            _ExtentX        =   1085
            _ExtentY        =   1085
            _StockProps     =   79
            BackColor       =   8454143
            Appearance      =   6
            Picture         =   "FrmCabang.frx":166A8
            BorderGap       =   0
         End
         Begin XtremeSuiteControls.PushButton CmdQuit 
            Height          =   615
            Left            =   6480
            TabIndex        =   10
            ToolTipText     =   "Exit"
            Top             =   75
            Width           =   615
            _Version        =   851972
            _ExtentX        =   1085
            _ExtentY        =   1085
            _StockProps     =   79
            BackColor       =   8454143
            Appearance      =   6
            Picture         =   "FrmCabang.frx":1C052
            BorderGap       =   0
         End
      End
      Begin VB.TextBox TxtAlamat 
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
         Height          =   690
         Left            =   2160
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   1080
         Width           =   2595
      End
      Begin MSMask.MaskEdBox TxtTlp 
         Height          =   330
         Left            =   2160
         TabIndex        =   4
         Top             =   2160
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         MaxLength       =   15
         Mask            =   "[&&&&]-&&&&&&&&"
         PromptChar      =   "_"
      End
      Begin VB.TextBox TxtKota 
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
         TabIndex        =   3
         Top             =   1800
         Width           =   2595
      End
      Begin VB.Label LB_Reload 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Cabang"
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
         TabIndex        =   17
         Top             =   720
         Width           =   1110
      End
      Begin VB.Label LB_Reload 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alamat"
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
         TabIndex        =   16
         Top             =   1080
         Width           =   570
      End
      Begin VB.Label LB_Reload 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kota"
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
         TabIndex        =   15
         Top             =   1800
         Width           =   360
      End
      Begin VB.Label LB_Reload 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Telepon"
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
         TabIndex        =   14
         Top             =   2160
         Width           =   675
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H004C4C4C&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00808080&
         Height          =   2010
         Index           =   0
         Left            =   360
         Top             =   600
         Width           =   7065
      End
   End
End
Attribute VB_Name = "FrmCabang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit As Boolean

Private Sub CmdAdd_Click()
tombol False
Edit = False
TxtNama.SetFocus
End Sub

Sub RefreshData()
Set Grid.DataSource = Nothing
SQL = "SELECT KodeCabang AS [Kode Cabang], NamaCabang AS [Nama Cabang], Alamat, " & _
    " Kota, NoTlp AS Telepon FROM Cabang order by KodeCabang"
Set Grid.DataSource = DbCon.Execute(SQL)
Grid.Refresh

Grid.Columns(0).Visible = False

Grid.Columns(1).Alignment = dbgLeft
Grid.Columns(1).Width = 3000
Grid.Columns(2).Alignment = dbgLeft
Grid.Columns(2).Width = 3000
Grid.Columns(3).Alignment = dbgLeft
Grid.Columns(3).Width = 2000
Grid.Columns(4).Alignment = dbgLeft
Grid.Columns(4).Width = 2000
End Sub

Private Sub CmdCancel_Click()
tombol True
RefreshData
bersih
TxtKode.Locked = False
End Sub

Private Sub CmdDelete_Click()
If MsgBox("Yakin akan menghapus data ini?" & vbCrLf & "" _
            & "KODE CABANG : " & Trim(Grid.Columns(0).Text) + vbCrLf & "" _
            & "NAMA CABANG : " & Trim(Grid.Columns(1).Text) + vbCrLf & "", _
         vbYesNo + vbQuestion) = vbYes Then
    SQL = "delete from cabang where kodeCabang ='" & Trim(Grid.Columns(0).Text) & "'"
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
If Trim(TxtNama) = "" Then
    MsgBox "Nama Grup masih kosong."
    TxtNama.SetFocus
    Exit Sub
ElseIf Trim(TxtAlamat) = "" Then
    MsgBox "Alamat Supplier masih kosong."
    TxtAlamat.SetFocus
    Exit Sub
ElseIf Trim(TxtKota) = "" Then
    MsgBox "Kota Supplier masih kosong."
    TxtKota.SetFocus
    Exit Sub
End If

If Not Edit Then
    SQL = "insert into cabang (kodeCabang,NamaCabang,alamat,kota,notlp) " & _
        " values ('" & Trim(KodeAuto) & "','" & Trim(TxtNama) & "','" & Trim(TxtAlamat) & _
        "','" & Trim(TxtKota) & "','" & Trim(TxtTlp) & "')"
Else
    SQL = "update supplier set namaSupplier ='" & Trim(TxtNama) & "',alamat='" & Trim(TxtAlamat) & _
        "',kota='" & Trim(TxtKota) & "', noTlp='" & Trim(TxtTlp) & _
        "' where kodeSupplier ='" & Trim(TxtKode) & "'"
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

tombol True
Edit = False
RefreshData
bersih
End Sub

Function KodeAuto()
SQL = "Select kodeCabang from Cabang order by kodeCabang Desc"
Set RsFind = DbCon.Execute(SQL)
If RsFind.BOF Then
   KodeAuto = "CAB-00001"
Else
   KodeAuto = "CAB-" & Format(CInt(Right(RsFind!kodeCabang, 5)) + 1, "00000")
End If
End Function

Sub tombol(Status As Boolean)
CmdAdd.Visible = Status
CmdEdit.Visible = Status
CmdDelete.Visible = Status

CmdSave.Visible = Not Status
CmdCancel.Visible = Not Status

Grid.Visible = Status
End Sub

Private Sub Grid_DblClick()
TxtKode = Trim(Grid.Columns(0).Text)
TxtNama = Trim(Grid.Columns(1).Text)
TxtAlamat = Trim(Grid.Columns(2).Text)
TxtKota = Trim(Grid.Columns(3).Text)
TxtTlp = Trim(Grid.Columns(4).Text)

TxtKode.Locked = True
Edit = True
tombol False
End Sub

Sub bersih()
TxtKode = ""
TxtNama = ""
TxtAlamat = ""
TxtKota = ""
TxtTlp.Text = "[____]-________"
End Sub

Private Sub TxtAlamat_KeyDown(KeyCode As Integer, Shift As Integer)
Enter KeyCode
End Sub

Private Sub TxtKode_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase((Chr(KeyAscii))))
If KeyAscii = vbKeyReturn Then
    SendKeys "{tab}"
    KeyAscii = 0
End If
End Sub

Private Sub TxtKota_KeyDown(KeyCode As Integer, Shift As Integer)
Enter KeyCode
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


Private Sub TxtTlp_KeyDown(KeyCode As Integer, Shift As Integer)
Enter KeyCode
End Sub


