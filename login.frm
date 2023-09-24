VERSION 5.00
Begin VB.Form login 
   BackColor       =   &H00E0E0E0&
   Caption         =   "LOG IN"
   ClientHeight    =   2070
   ClientLeft      =   5790
   ClientTop       =   4875
   ClientWidth     =   3840
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
   ScaleHeight     =   2070
   ScaleWidth      =   3840
   Begin VB.CommandButton Command2 
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
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1320
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "LOG IN"
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
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1320
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1560
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   720
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   240
      Width           =   2055
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   3600
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      Height          =   240
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   930
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User"
      Height          =   240
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   435
   End
   Begin VB.Shape Shape1 
      Height          =   1815
      Left            =   120
      Top             =   120
      Width           =   3615
   End
End
Attribute VB_Name = "login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
sambung
sql = "select * from login where nama = '" & Text1.Text & "' and pass = '" & Text2.Text & "'"
Set rs = con.Execute(sql)
If Not rs.EOF Then
With menu
.Show
.inp.Enabled = True
.car.Enabled = True
.cetak.Enabled = True
.us.Enabled = True
.keluar.Enabled = True
.Command1.Enabled = True
.Command2.Enabled = True
.Command3.Enabled = True
End With
Unload Me
Else
MsgBox ("Periksa user dan password anda, kayaknya salah deh..."), vbInformation, "Oopzz......"
Text1.Text = ""
Text2.Text = ""
Text1.SetFocus
End If
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Form_Load()
Me.Height = 2505
Me.Width = 3960
Me.Left = 5730
Me.Top = 3000
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
sambung
rs.Open "select * from login where nama = '" & Text1.Text & "'", con, adOpenDynamic, adLockOptimistic
Text2.SetFocus
    If rs.EOF Then
    MsgBox ("User salah tuh...!!"), vbInformation, "Upz.."
    Text1.Text = ""
    Text1.SetFocus
    End If
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
sambung
rs.Open "select * from login where pass = '" & Text2.Text & "'", con, adOpenDynamic, adLockOptimistic
Command1.SetFocus
    If rs.EOF Then
    MsgBox ("Password salah tuh...!!"), vbInformation, "Upz.."
    Text2.Text = ""
    Text2.SetFocus
    End If
End If
End Sub
