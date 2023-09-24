VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.OCX"
Begin VB.Form cari 
   BackColor       =   &H00E0E0E0&
   Caption         =   "CARI DATA"
   ClientHeight    =   6045
   ClientLeft      =   4800
   ClientTop       =   2625
   ClientWidth     =   6255
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "cari.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6045
   ScaleWidth      =   6255
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "KELUAR"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3000
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "BERSIH"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3000
      Width           =   1455
   End
   Begin MSComctlLib.ListView LvKaryawan 
      Height          =   1935
      Left            =   120
      TabIndex        =   5
      Top             =   3960
      Width           =   6015
      _ExtentX        =   10610
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
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "NIP"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nama Karyawan"
         Object.Width           =   6174
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Nama Istri/Suami"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "TAMPIL"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3000
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   3
      Top             =   2160
      Width           =   4815
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Nama Karyawan"
      Height          =   375
      Left            =   3720
      TabIndex        =   1
      Top             =   1560
      Width           =   2055
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "NIP Karyawan"
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Shape Shape5 
      Height          =   735
      Left            =   4200
      Shape           =   4  'Rounded Rectangle
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Shape Shape4 
      Height          =   735
      Left            =   360
      Shape           =   4  'Rounded Rectangle
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "CARI DATA KARYAWAN"
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
      TabIndex        =   6
      Top             =   360
      Width           =   5775
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      Height          =   735
      Left            =   2280
      Shape           =   4  'Rounded Rectangle
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Line Line6 
      X1              =   4560
      X2              =   4560
      Y1              =   1320
      Y2              =   1440
   End
   Begin VB.Line Line5 
      X1              =   4200
      X2              =   4560
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line Line4 
      X1              =   1680
      X2              =   1680
      Y1              =   1440
      Y2              =   1320
   End
   Begin VB.Line Line3 
      X1              =   2040
      X2              =   1680
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CARI BERDASARKAN"
      Height          =   240
      Left            =   2160
      TabIndex        =   2
      Top             =   1200
      Width           =   1965
   End
   Begin VB.Shape Shape2 
      Height          =   2775
      Left            =   120
      Top             =   1080
      Width           =   6015
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   6120
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      Height          =   615
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   6015
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   6120
      Y1              =   120
      Y2              =   120
   End
End
Attribute VB_Name = "cari"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
On Error GoTo salah
sambung
If Option1.Value = True Then
input_data.Text_nip.Text = Text1.Text
sql = "select * from karyawan where nip = '" & input_data.Text_nip.Text & "'"
Set rs = con.Execute(sql)
Else
input_data.Text_nama.Text = Text1.Text
sql = "select * from karyawan where nama = '" & input_data.Text_nama.Text & "'"
Set rs = con.Execute(sql)
End If
input_data.Text_nip.Text = rs.Fields(0)
input_data.Text_nama.Text = rs.Fields(1)
input_data.Text_namais.Text = rs.Fields(2)
input_data.tgl_is.Text = rs.Fields(3)
input_data.Text_a1.Text = rs.Fields(4)
input_data.tgl1.Text = rs.Fields(5)
input_data.Text_a2.Text = rs.Fields(6)
input_data.tgl2.Text = rs.Fields(7)
input_data.Text_a3.Text = rs.Fields(8)
input_data.tgl3.Text = rs.Fields(9)
input_data.Text_a4.Text = rs.Fields(10)
input_data.tgl4.Text = rs.Fields(11)
input_data.Text_a5.Text = rs.Fields(12)
input_data.tgl5.Text = rs.Fields(13)
input_data.Show
Exit Sub
salah:
MsgBox ("Data tidak ada coy...."), vbInformation, "O...Oo..."
Text1 = ""
Text1.SetFocus
input_data.keluar = True
End Sub

Private Sub Command2_Click()
Text1.Text = ""
Text1.SetFocus
tampil ("select * from karyawan")
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Form_Load()
Me.Height = 6555
Me.Left = 4740
Me.Top = 1000
Me.Width = 6375
tampil ("select * from karyawan")
Option1.Value = False
Option2.Value = False
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
        rs.MoveNext
    Wend
End Function

Private Sub Option1_Click()
tampil ("select * from karyawan order by nip")
Text1.SetFocus
End Sub

Private Sub Option2_Click()
tampil ("select * from karyawan order by nama")
Text1.SetFocus
End Sub

Private Sub Text1_Change()
If Option1.Value = True Then
tampil ("select * from karyawan where nip like '" & Text1.Text & "%'")
Else
If Option2.Value = True Then
tampil ("select * from karyawan where nama like '" & Text1.Text & "%'")
Else
MsgBox ("pilih kriteria dulu ya...."), vbInformation, "Ooopz...."
End If
End If
End Sub

Private Sub LvKaryawan_Click()
If Option1.Value = True Then
    If rs.State = 1 Then rs.Close
        rs.Open "select * from karyawan where [nip] = '" & LvKaryawan.SelectedItem & "'", con
        Text1.Text = rs.Fields(0)
Else
    If rs.State = 1 Then rs.Close
        rs.Open "select * from karyawan where [nip] = '" & LvKaryawan.SelectedItem & "'", con
        Text1.Text = rs.Fields(1)
End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Command1.SetFocus
End Sub
