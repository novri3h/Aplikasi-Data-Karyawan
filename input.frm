VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form input_data 
   BackColor       =   &H00E0E0E0&
   Caption         =   "INPUT DATA"
   ClientHeight    =   8085
   ClientLeft      =   4185
   ClientTop       =   1800
   ClientWidth     =   7710
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "input.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8085
   ScaleWidth      =   7710
   Begin VB.CommandButton cetak 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Cetak"
      Height          =   495
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   7440
      Width           =   975
   End
   Begin VB.TextBox tgl5 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   5760
      MaxLength       =   10
      TabIndex        =   32
      Top             =   4680
      Width           =   1575
   End
   Begin VB.TextBox tgl4 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   5760
      MaxLength       =   10
      TabIndex        =   31
      Top             =   4320
      Width           =   1575
   End
   Begin VB.TextBox tgl3 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   5760
      MaxLength       =   10
      TabIndex        =   30
      Top             =   3960
      Width           =   1575
   End
   Begin VB.TextBox tgl2 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   5760
      MaxLength       =   10
      TabIndex        =   29
      Top             =   3600
      Width           =   1575
   End
   Begin VB.TextBox tgl1 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   5760
      MaxLength       =   10
      TabIndex        =   28
      Top             =   3240
      Width           =   1575
   End
   Begin VB.TextBox tgl_is 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2280
      MaxLength       =   10
      TabIndex        =   27
      Top             =   2280
      Width           =   1575
   End
   Begin VB.CommandButton keluar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Keluar"
      Height          =   495
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   7440
      Width           =   975
   End
   Begin VB.CommandButton hapus 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Hapus"
      Height          =   495
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   7440
      Width           =   975
   End
   Begin VB.CommandButton edit 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Edit"
      Height          =   495
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   7440
      Width           =   975
   End
   Begin VB.CommandButton batal 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ba&tal"
      Height          =   495
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   7440
      Width           =   975
   End
   Begin VB.CommandButton simpan 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Simpan"
      Height          =   495
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   7440
      Width           =   975
   End
   Begin VB.CommandButton baru 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Baru"
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   7440
      Width           =   975
   End
   Begin MSComctlLib.ListView LvKaryawan 
      Height          =   1935
      Left            =   120
      TabIndex        =   19
      Top             =   5400
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   3413
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   14
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "NIP"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nama Karyawan"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Nama Istri/Suami"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Tgl lahir"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Anak ke 1"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Tgl Lahir"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Anak Ke2"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Tgl Lahir"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Anak ke 3"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Tgl Lahir"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Anak ke 4"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Tgl Lahir"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "Anak ke 5"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "Tgl lahir"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.TextBox Text_a5 
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   1440
      TabIndex        =   18
      Top             =   4680
      Width           =   4335
   End
   Begin VB.TextBox Text_a4 
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   1440
      TabIndex        =   17
      Top             =   4320
      Width           =   4335
   End
   Begin VB.TextBox Text_a3 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   1440
      TabIndex        =   16
      Top             =   3960
      Width           =   4335
   End
   Begin VB.TextBox Text_a2 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   1440
      TabIndex        =   15
      Top             =   3600
      Width           =   4335
   End
   Begin VB.TextBox Text_a1 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   1440
      TabIndex        =   14
      Top             =   3240
      Width           =   4335
   End
   Begin VB.TextBox Text_namais 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   2280
      TabIndex        =   13
      Top             =   1920
      Width           =   3375
   End
   Begin VB.TextBox Text_nama 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   2280
      TabIndex        =   12
      Top             =   1560
      Width           =   3375
   End
   Begin VB.TextBox Text_nip 
      Appearance      =   0  'Flat
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "dd-MM-yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1057
         SubFormatType   =   0
      EndProperty
      Height          =   330
      Left            =   2280
      TabIndex        =   11
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NIP"
      Height          =   240
      Left            =   360
      TabIndex        =   34
      Top             =   1200
      Width           =   330
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "INPUT DATA KARYAWAN"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   26
      Top             =   360
      Width           =   7215
   End
   Begin VB.Line Line6 
      X1              =   1440
      X2              =   360
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Shape Shape5 
      Height          =   375
      Left            =   360
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Line Line5 
      X1              =   1440
      X2              =   360
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Line Line4 
      X1              =   1440
      X2              =   360
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Shape Shape4 
      Height          =   2175
      Left            =   1440
      Top             =   2880
      Width           =   4335
   End
   Begin VB.Shape Shape3 
      Height          =   2175
      Left            =   360
      Top             =   2880
      Width           =   6975
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      Height          =   615
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   7455
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   7560
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   7560
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "5"
      Height          =   360
      Left            =   360
      TabIndex        =   10
      Top             =   4680
      Width           =   1155
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "4"
      Height          =   360
      Left            =   360
      TabIndex        =   9
      Top             =   4320
      Width           =   1155
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "3"
      Height          =   360
      Left            =   360
      TabIndex        =   8
      Top             =   3960
      Width           =   1155
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "2"
      Height          =   345
      Left            =   360
      TabIndex        =   7
      Top             =   3600
      Width           =   1155
   End
   Begin VB.Line Line1 
      X1              =   360
      X2              =   7320
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Tanggal Lahir"
      Height          =   360
      Left            =   5760
      TabIndex        =   6
      Top             =   2880
      Width           =   1560
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "1"
      Height          =   360
      Left            =   360
      TabIndex        =   5
      Top             =   3240
      Width           =   1155
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Nama Anak"
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   1440
      TabIndex        =   4
      Top             =   2880
      Width           =   4335
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Anak Ke"
      Height          =   360
      Left            =   360
      TabIndex        =   3
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal Lahir"
      Height          =   240
      Left            =   360
      TabIndex        =   2
      Top             =   2280
      Width           =   1560
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nama Istri/Suami"
      Height          =   240
      Left            =   360
      TabIndex        =   1
      Top             =   1920
      Width           =   1680
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nama Karyawan"
      Height          =   240
      Left            =   360
      TabIndex        =   0
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      Height          =   4215
      Left            =   120
      Top             =   1080
      Width           =   7455
   End
End
Attribute VB_Name = "input_data"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub baru_Click()
bersih
aktiv
Text_nip.SetFocus
baru.Enabled = False
simpan.Enabled = True
batal.Enabled = True
End Sub

Private Sub batal_Click()
bersih
Text_nip.SetFocus
batal.Enabled = False
baru.Enabled = True
End Sub

Private Sub Command2_Click()

End Sub

Private Sub cetak_Click()
sambung
report
lap1.DataControl1.Source = "select * from karyawan where nip = '" & Text_nip.Text & "'"
lap1.Show
lap1.WindowState = maximized
End Sub

Private Sub edit_Click()
aktiv
Text_nip.Enabled = False
edit.Enabled = False
simpan.Enabled = True
Text_nama.SetFocus
End Sub

Private Sub Form_Load()
Me.Height = 8595
Me.Left = 4000
Me.Top = 200
Me.Width = 7830

pasif
simpan.Enabled = False
batal.Enabled = False
tampil ("select * from karyawan")
End Sub

Private Sub hapus_Click()
If MsgBox("Yaakinn....mau dihapus...???", vbYesNo, "Warning..") = vbYes Then
sambung
sql = "delete from karyawan where nip = '" & Text_nip.Text & "'"
con.Execute (sql)
bersih

hapus.Enabled = False
tampil ("select * from karyawan")
End If
End Sub

Private Sub keluar_Click()
Unload Me
End Sub

Private Sub simpan_Click()
If Text_nip.Enabled = True Then
sambung
sql = "insert into karyawan values('" & Text_nip.Text & "', '" & Text_nama.Text & "', '" & Text_namais.Text & "','" & tgl_is.Text & "','" & Text_a1.Text & "','" & tgl1.Text & "','" & Text_a2.Text & "','" & tgl2.Text & "','" & Text_a3.Text & "','" & tgl3.Text & "','" & Text_a4.Text & "','" & tgl4.Text & "','" & Text_a5.Text & "','" & tgl5.Text & "') "
con.Execute (sql)
Else
sql = "update karyawan set nama = '" & Text_nama.Text & "', nama_is = '" & Text_namais.Text & "', tgl_is = '" & tgl_is.Text & "', nama1 = '" & Text_a1.Text & "', tgl1 = '" & tgl1.Text & "', nama2 = '" & Text_a2.Text & "', tgl2 = '" & tgl2.Text & "', nama3 = '" & Text_a3.Text & "', tgl3 = '" & tgl3.Text & "', nama4 = '" & Text_a4.Text & "', tgl4 = '" & tgl4.Text & "', nama5 = '" & Text_a5.Text & "', tgl5 = '" & tgl5.Text & "' where nip = '" & Text_nip.Text & "'"
con.Execute (sql)
End If
pasif
tampil ("select * from karyawan")
simpan.Enabled = False
baru.Enabled = True
End Sub

Private Sub Text_a1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then tgl1.SetFocus
End Sub

Private Sub Text_a2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then tgl2.SetFocus
End Sub

Private Sub Text_a3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then tgl3.SetFocus
End Sub

Private Sub Text_a4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then tgl4.SetFocus
End Sub

Private Sub Text_a5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then tgl5.SetFocus
End Sub

Private Sub Text_nama_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Text_namais.SetFocus
End Sub

Private Sub Text_namais_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then tgl_is.SetFocus
End Sub

Private Sub Text_nip_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Text_nama.SetFocus
End Sub

Private Sub tgl_is_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Text_a1.SetFocus
End Sub

Private Sub tgl1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Text_a2.SetFocus
End Sub

Private Sub tgl2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Text_a3.SetFocus
End Sub

Private Sub tgl3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Text_a4.SetFocus
End Sub

Private Sub tgl4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Text_a5.SetFocus
End Sub

Private Sub tgl5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then simpan.SetFocus
End Sub

Sub bersih()
Text_nip = ""
Text_nama = ""
Text_namais = ""
tgl_is = ""
Text_a1 = ""
Text_a2 = ""
Text_a3 = ""
Text_a4 = ""
Text_a5 = ""
tgl1 = ""
tgl2 = ""
tgl3 = ""
tgl4 = ""
tgl5 = ""
End Sub

Sub aktiv()
Text_nip.Enabled = True
Text_nama.Enabled = True
Text_namais.Enabled = True
tgl_is.Enabled = True
Text_a1.Enabled = True
Text_a2.Enabled = True
Text_a3.Enabled = True
Text_a4.Enabled = True
Text_a5.Enabled = True
tgl1.Enabled = True
tgl2.Enabled = True
tgl3.Enabled = True
tgl4.Enabled = True
tgl5.Enabled = True
End Sub

Sub pasif()
Text_nip.Enabled = False
Text_nama.Enabled = False
Text_namais.Enabled = False
tgl_is.Enabled = False
Text_a1.Enabled = False
Text_a2.Enabled = False
Text_a3.Enabled = False
Text_a4.Enabled = False
Text_a5.Enabled = False
tgl1.Enabled = False
tgl2.Enabled = False
tgl3.Enabled = False
tgl4.Enabled = False
tgl5.Enabled = False
End Sub

Function tampil(strsql As String)
sambung
LvKaryawan.ListItems.Clear
Dim data As ListItem
If rs.State = 1 Then rs.Close
rs.Open strsql, con, adOpenDynamic, adLockOptimistic
    While Not rs.EOF
        Set data = LvKaryawan.ListItems.Add(, , rs.Fields(0))
            data.SubItems(1) = rs.Fields(1)
            data.SubItems(2) = rs.Fields(2)
            data.SubItems(3) = rs.Fields(3)
            data.SubItems(4) = rs.Fields(4)
            data.SubItems(5) = rs.Fields(5)
            data.SubItems(6) = rs.Fields(6)
            data.SubItems(7) = rs.Fields(7)
            data.SubItems(8) = rs.Fields(8)
            data.SubItems(9) = rs.Fields(9)
            data.SubItems(10) = rs.Fields(10)
            data.SubItems(11) = rs.Fields(11)
            data.SubItems(12) = rs.Fields(12)
            data.SubItems(13) = rs.Fields(13)
        rs.MoveNext
    Wend
End Function

Private Sub LvKaryawan_Click()
    If rs.State = 1 Then rs.Close
        rs.Open "select * from karyawan where [nip] = '" & LvKaryawan.SelectedItem & "'", con
        Text_nip = rs.Fields(0)
        Text_nama = rs.Fields(1)
        Text_namais = rs.Fields(2)
        tgl_is = rs.Fields(3)
        Text_a1 = rs.Fields(4)
        tgl1 = rs.Fields(5)
        Text_a2 = rs.Fields(6)
        tgl2 = rs.Fields(7)
        Text_a3 = rs.Fields(8)
        tgl3 = rs.Fields(9)
        Text_a4 = rs.Fields(10)
        tgl4 = rs.Fields(11)
        Text_a5 = rs.Fields(12)
        tgl5 = rs.Fields(13)
End Sub

Sub report()
lap1.DataControl1.CursorLocation = ddADOUseClient
lap1.DataControl1.CursorType = ddADOOpenDynamic
lap1.DataControl1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & App.Path & "\data.mdb;Persist Security Info=False"
End Sub

