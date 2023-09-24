VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.OCX"
Begin VB.Form user 
   BackColor       =   &H00E0E0E0&
   Caption         =   "INPUT USER"
   ClientHeight    =   4695
   ClientLeft      =   5790
   ClientTop       =   2625
   ClientWidth     =   4455
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4695
   ScaleWidth      =   4455
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Simpan"
      Height          =   495
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2280
      Width           =   855
   End
   Begin MSComctlLib.ListView LvUser 
      Height          =   1455
      Left            =   120
      TabIndex        =   7
      Top             =   3120
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   2566
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
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "User"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Password"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Keluar"
      Height          =   495
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2280
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Hapus"
      Height          =   495
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2280
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Baru"
      Height          =   495
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2280
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1680
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1680
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   1200
      Width           =   2415
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "INPUT DATA USER"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   360
      Width           =   3975
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      Height          =   240
      Left            =   360
      TabIndex        =   1
      Top             =   1680
      Width           =   930
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nama"
      Height          =   240
      Left            =   360
      TabIndex        =   0
      Top             =   1200
      Width           =   540
   End
   Begin VB.Shape Shape2 
      Height          =   2055
      Left            =   120
      Top             =   960
      Width           =   4215
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   4320
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   4215
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   4320
      Y1              =   120
      Y2              =   120
   End
End
Attribute VB_Name = "user"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
aktif
Text1 = ""
Text2 = ""
Text1.SetFocus
Command1.Enabled = False
Command4.Enabled = True
End Sub

Private Sub Command2_Click()
If MsgBox("Yaakinn....mau dihapus...???", vbYesNo, "Warning..") = vbYes Then
sambung
sql = "delete from login where nama = '" & Text1.Text & "'"
con.Execute (sql)
tampil ("select * from login")
Text1.Text = ""
Text2.Text = ""
Command2.Enabled = False
End If
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Command4_Click()
sambung
sql = "insert into login values('" & Text1.Text & "','" & Text2.Text & "')"
con.Execute (sql)
pasif
tampil ("select * from login")
Command1.Enabled = True
Command4.Enabled = False
End Sub

Private Sub Form_Load()
Me.Height = 5220
Me.Left = 5730
Me.Top = 2160
Me.Width = 4575
pasif
Command1.Enabled = True
Command4.Enabled = False
tampil ("select * from login")
End Sub

Function tampil(strsql As String)
sambung
LvUser.ListItems.Clear
Dim data As ListItem
If rs.State = 1 Then rs.Close
rs.Open strsql, con, adOpenDynamic, adLockOptimistic
    While Not rs.EOF
        Set data = LvUser.ListItems.Add(, , rs.Fields(0))
            data.SubItems(1) = rs.Fields(1)
        rs.MoveNext
    Wend
End Function

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Text2.SetFocus
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Command4.SetFocus
End Sub

Sub pasif()
Text1.Enabled = False
Text2.Enabled = False
End Sub

Sub aktif()
Text1.Enabled = True
Text2.Enabled = True
End Sub

Private Sub LvUser_Click()
    If rs.State = 1 Then rs.Close
        rs.Open "select * from login where [nama] = '" & LvUser.SelectedItem & "'", con
        Text1.Text = rs.Fields(0)
        Text2.Text = rs.Fields(1)
Command2.Enabled = True
End Sub
