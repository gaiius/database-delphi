VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmMenuUtama 
   BackColor       =   &H8000000C&
   Caption         =   "Menu Utama Aplikasi"
   ClientHeight    =   7095
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   9420
   Icon            =   "frmMenuUtama.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Left            =   1560
      Top             =   0
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   6600
      Width           =   9420
      _ExtentX        =   16616
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   1411
            MinWidth        =   1411
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   11360
            MinWidth        =   11360
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   1
            TextSave        =   "17:23"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   1
            TextSave        =   "10-07-2009"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuMaster 
      Caption         =   "&MASTER"
      Begin VB.Menu mnuRef 
         Caption         =   "Referensi"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuMDivisi 
         Caption         =   "Divisi"
      End
      Begin VB.Menu mnuJab 
         Caption         =   "Jabatan"
      End
      Begin VB.Menu mnuStKawin 
         Caption         =   "Status Kawin"
      End
      Begin VB.Menu mnuMKary 
         Caption         =   "Karyawan"
      End
      Begin VB.Menu mnuMKel 
         Caption         =   "Keluarga"
         Visible         =   0   'False
      End
      Begin VB.Menu grs0 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuMUser 
         Caption         =   "USER && PASSWORD"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuTranskasi 
      Caption         =   "&TRANSAKSI"
      Begin VB.Menu mnuTSPJPK 
         Caption         =   "SPJPK"
      End
      Begin VB.Menu mnuTRincian 
         Caption         =   "Rincian Biaya"
      End
   End
   Begin VB.Menu mnuLap 
      Caption         =   "&LAPORAN"
      Begin VB.Menu mnuLKary 
         Caption         =   "Karyawan"
      End
      Begin VB.Menu mnuLJaminan 
         Caption         =   "Jaminan"
      End
      Begin VB.Menu mnuLSPJPK 
         Caption         =   "SPJPK"
      End
      Begin VB.Menu mnuLRincian 
         Caption         =   "Rincian Biaya"
      End
   End
   Begin VB.Menu mnuKeluar 
      Caption         =   "KELUAR"
   End
End
Attribute VB_Name = "frmMenuUtama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim counter As Integer
Public Dibuat As String

Private Sub MDIForm_Load()
On Error Resume Next
cnKoneksi

Me.Picture = LoadPicture(App.Path & "\bg.jpg")
Dibuat = "******* SISTEM INFORMASI PEMBUATAN SURAT PENGANTAR JAMINAN PEMELIHARAAN KESEHATAN (SPJPK) DI PT.KIEC *******"
counter = 0
Timer1.Interval = 100
End Sub

Private Sub mnuJab_Click()
frmJab.Show
End Sub

Private Sub mnuKeluar_Click()
If MsgBox("Anda yakin akan keluar dari Aplikasi ? ", vbYesNo, "Keluar") = vbYes Then
    End
End If
End Sub

Private Sub mnuLJaminan_Click()
frmLapJaminan.Show
End Sub

Private Sub mnuLKary_Click()
frmLapKary.Show
End Sub

Private Sub mnuLRincian_Click()
frmLapRincian.Show
End Sub

Private Sub mnuLSPJPK_Click()
frmLapSPJPK.Show
End Sub

Private Sub mnuMDivisi_Click()
frmDivisi.Show
End Sub

Private Sub mnuMKary_Click()
frmInKary.Show
End Sub

Private Sub mnuStKawin_Click()
frmKawin.Show
End Sub

Private Sub mnuTRincian_Click()
frmRinciBiaya.Show
End Sub

Private Sub mnuTSPJPK_Click()
frmInSPJPK.Show
End Sub

Public Function TulisJalan(Hitung As Integer, _
strKalimat As String, Panjang As Integer)

  If Hitung = Len(strKalimat) + Panjang Then
     Hitung = 0
  ElseIf Hitung > Len(strKalimat) Then
     TulisJalan = strKalimat & Space(Hitung - _
                  Len(strKalimat))
  Else
     TulisJalan = Mid(strKalimat, 1, Hitung)
  End If
End Function

Private Sub Timer1_Timer()
Dim Kalimat As String
Dim pnlX1 As Panel
Set pnlX1 = StatusBar1.Panels(2)
Kalimat = Dibuat
counter = counter + 1
DoEvents
pnlX1.Text = TulisJalan(counter, Kalimat, 150)
End Sub

