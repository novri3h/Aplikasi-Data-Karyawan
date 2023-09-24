VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form karyawan 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Form1"
   ClientHeight    =   6120
   ClientLeft      =   60
   ClientTop       =   465
   ClientWidth     =   10320
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "karyawan.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6120
   ScaleWidth      =   10320
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Keluar"
      Height          =   615
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Cetak"
      Height          =   615
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1200
      Width           =   1215
   End
   Begin MSComctlLib.ListView LvKaryawan 
      Height          =   6855
      Left            =   240
      TabIndex        =   3
      Top             =   2160
      Width           =   14655
      _ExtentX        =   25850
      _ExtentY        =   12091
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
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nama Karyawan"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Nama Istri/Suami"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Tgl Lahir"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Anak ke 1"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "Tgl Lahir"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Anak ke 2"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   7
         Text            =   "Tgl Lahir"
         Object.Width           =   2470
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Anak Ke 3"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   9
         Text            =   "Tgl Lahir"
         Object.Width           =   2470
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Anak Ke 4"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   11
         Text            =   "Tgl Lahir"
         Object.Width           =   2470
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "Anak Ke 5"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   13
         Text            =   "Tgl Lahir"
         Object.Width           =   2470
      EndProperty
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Berdasarkan Nama Karyawan"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1560
      Width           =   3375
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Berdasarkan NIP"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   240
      X2              =   5040
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DATA SELURUH KARYAWAN"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   4860
   End
End
Attribute VB_Name = "karyawan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Private Sub Command1_Click()
If Option1.Value = True Then
report
lap2.DataControl1.Source = "select * from karyawan order by nip"
lap2.Show
lap2.WindowState = maximized
Else
report
lap2.DataControl1.Source = "select * from karyawan order by nama"
lap2.Show
lap2.WindowState = maximized
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
tampil ("select * from karyawan")
End Sub

Sub report()
lap2.DataControl1.CursorLocation = ddADOUseClient
lap2.DataControl1.CursorType = ddADOOpenDynamic
lap2.DataControl1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & App.Path & "\data.mdb;Persist Security Info=False"
End Sub

Private Sub Option1_Click()
tampil ("select * from karyawan order by nip")
End Sub

Private Sub Option2_Click()
tampil ("select * from karyawan order by nama")
End Sub
