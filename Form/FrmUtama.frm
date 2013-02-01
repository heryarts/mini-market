VERSION 5.00
Begin VB.Form FrmUtama 
   Caption         =   "Mini Market 1.0"
   ClientHeight    =   6630
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   8280
   ControlBox      =   0   'False
   Icon            =   "FrmUtama.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "FrmUtama.frx":068A
   ScaleHeight     =   6630
   ScaleWidth      =   8280
   WindowState     =   2  'Maximized
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Minimarket Anugerah"
      BeginProperty Font 
         Name            =   "Daredevil"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   4320
      TabIndex        =   0
      Top             =   360
      Width           =   3375
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   1575
      Left            =   4080
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   3855
   End
   Begin VB.Menu MFile 
      Caption         =   "&File"
      Begin VB.Menu MLogin 
         Caption         =   "Login"
         Shortcut        =   ^L
      End
      Begin VB.Menu MLogout 
         Caption         =   "Logout"
         Shortcut        =   ^O
      End
   End
   Begin VB.Menu MData 
      Caption         =   "&Data"
      Begin VB.Menu MSupplier 
         Caption         =   "Supplier"
      End
      Begin VB.Menu MPembeli 
         Caption         =   "Pembeli"
      End
      Begin VB.Menu MBaris1 
         Caption         =   "-"
      End
      Begin VB.Menu MGrup 
         Caption         =   "Grup"
      End
      Begin VB.Menu MProduk 
         Caption         =   "Produk"
      End
      Begin VB.Menu MBarang 
         Caption         =   "Barang"
      End
   End
   Begin VB.Menu MTransaksi 
      Caption         =   "&Transaksi"
   End
   Begin VB.Menu MLaporan 
      Caption         =   "&Laporan"
   End
   Begin VB.Menu MSetting 
      Caption         =   "&Setting"
      Begin VB.Menu MUser 
         Caption         =   "User"
      End
      Begin VB.Menu MCabang 
         Caption         =   "Cabang"
      End
   End
   Begin VB.Menu MExit 
      Caption         =   "&Exit"
   End
End
Attribute VB_Name = "FrmUtama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
Me.Shape1.Left = Me.Left + Me.Width - 4200
Me.Label1.Left = Me.Left + Me.Width - 3900
End Sub

Private Sub Form_Load()
bukaDatabase

End Sub

Private Sub Form_Unload(Cancel As Integer)
If MsgBox("Akan keluar aplikasi?", vbYesNo + vbInformation, "Keluar") = vbYes Then
    Cancel = 0
Else
    Cancel = 1
End If
End Sub

Private Sub MBarang_Click()
FrmBarang.Show 1
End Sub

Private Sub MCabang_Click()
FrmCabang.Show 1
End Sub

Private Sub MExit_Click()
Unload Me
End Sub

Private Sub MGrup_Click()
FrmGrup.Show 1
End Sub

Private Sub MProduk_Click()
FrmProduk.Show 1
End Sub

Private Sub MSupplier_Click()
FrmSupplier.Show 1
End Sub

Private Sub MUser_Click()
FrmUser.Show 1
End Sub
